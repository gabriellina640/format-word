"""Apply explicit profile properties to a preserved DOCX package."""
from __future__ import annotations

import os
import math
import errno
import shutil
import tempfile
from copy import deepcopy
from dataclasses import dataclass, replace
from io import BytesIO
from pathlib import Path
from zipfile import BadZipFile, ZipFile

from docx import Document
from docx.document import Document as DocumentType
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from PIL import Image, ImageOps
from lxml import etree

from app.config import (CATEGORY_LABELS, FormatSettings, SettingsError, paper_dimensions,
                        settings_for_category, validate_settings)
from app.document_model import (DocumentOverrides, ImageEdit, classify_paragraph, fingerprint,
                                image_elements, is_protected, nearest_paragraph, own_runs, xml_query)

SUPPORTED_INPUT_EXTENSIONS = {'.docx'}
MAX_EXPANDED_BYTES = 512 * 1024 * 1024
ALIGNMENTS = {'left': WD_ALIGN_PARAGRAPH.LEFT, 'center': WD_ALIGN_PARAGRAPH.CENTER,
              'right': WD_ALIGN_PARAGRAPH.RIGHT, 'justify': WD_ALIGN_PARAGRAPH.JUSTIFY}


class FormatterError(Exception):
    """A document could not be formatted with the requested fidelity."""


@dataclass(frozen=True, slots=True)
class FormatResult:
    output_path: Path
    paragraphs: int
    warnings: tuple[str, ...] = ()


def validate_input(input_path: Path, max_input_mb: int = 50) -> None:
    if input_path.suffix.lower() not in SUPPORTED_INPUT_EXTENSIONS:
        raise FormatterError('Selecione um arquivo Word .docx. PDF e .doc não são suportados.')
    if not input_path.is_file():
        raise FormatterError('Arquivo não encontrado.')
    if input_path.stat().st_size > max_input_mb * 1024 * 1024:
        raise FormatterError(f'O arquivo deve ter no máximo {max_input_mb} MB.')
    try:
        with ZipFile(input_path) as package:
            parts = package.infolist()
            if len(parts) > 10000 or sum(p.file_size for p in parts) > MAX_EXPANDED_BYTES:
                raise FormatterError('O conteúdo descompactado do Word excede o limite de segurança.')
            if any(p.flag_bits & 1 for p in parts):
                raise FormatterError('Word protegido por senha não é suportado.')
            names = set(package.namelist())
            if not {'[Content_Types].xml', 'word/document.xml'}.issubset(names):
                raise FormatterError('O arquivo não contém um documento Word válido.')
            if any('vbaProject' in name or name.startswith('_xmlsignatures/') for name in names):
                raise FormatterError('Documentos com macros ou assinatura digital não são suportados.')
    except (BadZipFile, OSError) as exc:
        raise FormatterError('Não foi possível ler o Word. O arquivo está inválido, protegido ou inacessível.') from exc


def format_document(input_path: Path, output_dir: Path, settings: FormatSettings,
                    overrides: DocumentOverrides | None = None) -> FormatResult:
    try:
        settings = validate_settings(settings, check_assets=True)
        input_path = Path(input_path).expanduser().resolve()
        output_dir = Path(output_dir).expanduser().resolve()
        validate_input(input_path, settings.max_input_mb)
        before = fingerprint(input_path)
        overrides = deepcopy(overrides)
        if overrides is not None and (not isinstance(overrides, DocumentOverrides) or overrides.fingerprint != before):
            raise FormatterError('O documento mudou desde a revisão. Revise o arquivo novamente antes de aplicar.')
        document = Document(str(input_path))
        warnings = _inspect_content(document, settings)
        paragraphs = list(document._element.body.xpath('.//w:p'))
        _validate_overrides(overrides, paragraphs)
        protected = {el: etree.tostring(el) for el in paragraphs if is_protected(el)}
        # Section properties embedded in protected text must remain untouched too.
        protected_sections = {i for i, section in enumerate(document.sections)
            if is_protected(section._sectPr) or nearest_paragraph(section._sectPr) in protected}
        if protected_sections and (settings.header_mode != 'preserve' or settings.footer_mode != 'preserve'):
            raise FormatterError('Há seções com alterações controladas. Preserve cabeçalho e rodapé ou resolva essas revisões no Word.')
        if protected_sections:
            warnings.append('Seções com alterações controladas preservaram também papel, margens e distâncias.')
        image_alignments = {el: _effective_alignment(nearest_paragraph(el), document) for el in image_elements(document)}
        _configure_sections(document, settings, protected_sections)
        _configure_header_footer(document, settings)
        rules = {category: settings_for_category(settings, category) for category in CATEGORY_LABELS}
        expected = {}
        for index, element in enumerate(paragraphs):
            category = (overrides.paragraph_categories.get(index) if overrides else None) or classify_paragraph(element, document, settings)
            rule = None if element in protected else rules[category]
            expected[element] = rule
            if rule is None:
                continue
            paragraph = Paragraph(element, document._body)
            _format_paragraph(paragraph, rule)
            for run_element in own_runs(element):
                _format_run(Run(run_element, paragraph), rule)
        _apply_body_images(document, settings, overrides, paragraphs, expected, warnings, image_alignments)
        for element, original in protected.items():
            if etree.tostring(element) != original:
                raise FormatterError('Um trecho com revisão seria alterado. Nenhum arquivo foi publicado.')
        if fingerprint(input_path) != before:
            raise FormatterError('O documento mudou durante o processamento. Aplique novamente.')
        final_paragraphs = document._element.body.xpath('.//w:p')
        expectations = [expected.get(el) for el in final_paragraphs]
        _prepare_output_dir(output_dir)
        output_path = _save_exclusively(document, output_dir, input_path.stem, settings,
            expectations, etree.tostring(document._element.body), protected_sections)
        return FormatResult(output_path, len(paragraphs), tuple(warnings))
    except FormatterError:
        raise
    except SettingsError as exc:
        raise FormatterError(str(exc)) from exc
    except Exception as exc:
        raise FormatterError(f'Não foi possível formatar "{Path(input_path).name}": {exc}') from exc


def _validate_overrides(overrides, paragraphs):
    if overrides is None:
        return
    if not isinstance(overrides.paragraph_categories, dict) or not isinstance(overrides.image_edits, dict):
        raise FormatterError('Revisão de documento inválida.')
    for index, category in overrides.paragraph_categories.items():
        if type(index) is not int or not 0 <= index < len(paragraphs) or category not in CATEGORY_LABELS:
            raise FormatterError('Classificação de parágrafo inválida; revise o documento novamente.')
        if is_protected(paragraphs[index]):
            raise FormatterError('Não é possível reclassificar um trecho com alterações controladas.')


def _effective_alignment(element, document):
    if element is None:
        return WD_ALIGN_PARAGRAPH.LEFT
    paragraph = Paragraph(element, document._body)
    if paragraph.alignment is not None:
        return paragraph.alignment
    style = paragraph.style
    visited = set()
    while style is not None and style.style_id not in visited:
        visited.add(style.style_id)
        if style.paragraph_format.alignment is not None:
            return style.paragraph_format.alignment
        style = style.base_style
    return WD_ALIGN_PARAGRAPH.LEFT


def _apply_body_images(document, settings, overrides, paragraphs, expected, warnings, image_alignments):
    images = image_elements(document)
    edits = overrides.image_edits if overrides else {}
    if any(type(i) is not int or not 0 <= i < len(images) for i in edits):
        raise FormatterError('Imagem não encontrada; revise o documento novamente.')
    page_width, page_height = paper_dimensions(settings)
    body_width = page_width - settings.margin_left_cm - settings.margin_right_cm
    body_height = page_height - settings.margin_top_cm - settings.margin_bottom_cm
    insertion_tails = {}
    for index, element in enumerate(images):
        explicit = index in edits
        edit = edits.get(index, ImageEdit(settings.body_image_width_cm, settings.body_image_alignment))
        if not isinstance(edit, ImageEdit):
            raise FormatterError('Ajuste de imagem inválido.')
        if edit.alignment not in ('preserve', 'left', 'center', 'right') or edit.placement not in ('before', 'after'):
            raise FormatterError('Alinhamento ou posição da imagem inválidos.')
        if edit.width_cm is not None and (isinstance(edit.width_cm, bool) or not isinstance(edit.width_cm, (int, float))
                or not math.isfinite(edit.width_cm) or not .01 <= edit.width_cm <= body_width):
            raise FormatterError('A largura da imagem deve ser positiva e caber na área útil da página.')
        if edit.target_paragraph is not None and (type(edit.target_paragraph) is not int
                or not 0 <= edit.target_paragraph < len(paragraphs)):
            raise FormatterError('Escolha um parágrafo válido para posicionar a imagem.')
        if edit.width_cm is None and edit.alignment == 'preserve' and edit.target_paragraph is None:
            continue
        source = nearest_paragraph(element)
        if source is None or is_protected(source) or element.tag == qn('wp:anchor'):
            if explicit:
                raise FormatterError('Esta imagem é flutuante ou pertence a trecho protegido. Ajuste-a no Word.')
            warnings.append('Uma imagem flutuante ou protegida manteve tamanho e posição originais.')
            continue
        # Avoid touching whole shapes, linked pictures and unknown drawing geometry.
        extent = element.find(qn('wp:extent'))
        if extent is None or int(extent.get('cx', '0')) <= 0 or int(extent.get('cy', '0')) <= 0:
            raise FormatterError('A imagem não possui dimensões válidas para ajuste proporcional.')
        width, height = int(extent.get('cx')), int(extent.get('cy'))
        target = paragraphs[edit.target_paragraph] if edit.target_paragraph is not None else source
        if is_protected(target):
            raise FormatterError('Não é possível posicionar imagem junto de trecho com alterações controladas.')
        if edit.target_paragraph is not None and target.getparent() is not document._element.body:
            raise FormatterError('Para mover a imagem, escolha um parágrafo do corpo fora de tabelas e caixas de texto.')
        if edit.width_cm is not None:
            new_width = int(Cm(edit.width_cm))
            new_height = round(height * new_width / width)
            available_width = body_width
            for ancestor in target.iterancestors():
                if ancestor.tag == qn('w:tc'):
                    widths = ancestor.xpath('./w:tcPr/w:tcW')
                    if widths and widths[0].get(qn('w:type')) == 'dxa':
                        available_width = min(available_width, int(widths[0].get(qn('w:w'))) / 567 - .38)
                if ancestor.tag == qn('w:txbxContent'):
                    raise FormatterError('Redimensionamento de imagem dentro de caixa de texto exige ajuste no Word.')
            if edit.width_cm > available_width or new_height > Cm(body_height):
                raise FormatterError('A imagem proporcional não cabe na área disponível; reduza sua largura.')
            extent.set('cx', str(new_width)); extent.set('cy', str(new_height))
            for transform in xml_query(element, './a:graphic/a:graphicData/pic:pic/pic:spPr/a:xfrm/a:ext'):
                transform.set('cx', str(new_width)); transform.set('cy', str(new_height))
        # A dedicated paragraph gives a real alignment without shifting adjacent text.
        if edit.width_cm is not None or edit.target_paragraph is not None or edit.alignment != 'preserve':
            new_p = OxmlElement('w:p')
            if edit.placement == 'before' and edit.target_paragraph is not None:
                target.addprevious(new_p)
            else:
                insertion_tails.get(target, target).addnext(new_p)
                insertion_tails[target] = new_p
            paragraph = Paragraph(new_p, document._body)
            paragraph.paragraph_format.left_indent = Cm(0)
            paragraph.paragraph_format.right_indent = Cm(0)
            paragraph.paragraph_format.first_line_indent = Cm(0)
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(0)
            paragraph.paragraph_format.line_spacing = 1.0
            paragraph.alignment = ALIGNMENTS[edit.alignment] if edit.alignment != 'preserve' else image_alignments[element]
            drawing = element.getparent()
            drawing.remove(element)
            if not len(drawing):
                drawing.getparent().remove(drawing)
            new_drawing = OxmlElement('w:drawing')
            new_drawing.append(element)
            paragraph.add_run()._r.append(new_drawing)
            expected[new_p] = None


def _inspect_content(document: DocumentType, settings: FormatSettings) -> list[str]:
    body = document._element.body
    unsupported = {
        'w:ins': 'alterações controladas', 'w:del': 'alterações controladas',
        'w:moveFrom': 'alterações controladas', 'w:moveTo': 'alterações controladas',
        'w:pPrChange': 'alterações controladas', 'w:rPrChange': 'alterações controladas',
        'w:sectPrChange': 'alterações controladas', 'w:tblPrChange': 'alterações controladas',
        'w:trPrChange': 'alterações controladas', 'w:tcPrChange': 'alterações controladas',
        'w:cellIns': 'alterações controladas', 'w:cellDel': 'alterações controladas', 'w:cellMerge': 'alterações controladas',
        'w:txbxContent': 'caixas de texto', 'w:altChunk': 'conteúdo externo incorporado',
        'w:object': 'objetos incorporados',
        'm:oMath': 'equações',
    }
    warnings = []
    seen = set()
    for tag, label in unsupported.items():
        if body.xpath(f'.//{tag}'):
            if settings.complex_content_mode == 'reject':
                raise FormatterError(f'Este documento contém {label}. O perfil está configurado para bloquear conteúdo complexo.')
            if label not in seen:
                action = ('Recebem somente as regras de texto escolhidas; tamanho e posição das caixas são preservados.'
                          if tag == 'w:txbxContent' else 'Conteúdo protegido preservado; confira esses trechos no Word.')
                warnings.append(f'Documento com {label}. {action}')
                seen.add(label)
    if body.xpath('.//w:fldChar|.//w:fldSimple'):
        warnings.append('Campos e sumários foram preservados. Atualizá-los no Word pode recalcular seu texto e formatação.')
    if body.xpath('.//w:footnoteReference|.//w:endnoteReference|.//w:commentReference'):
        warnings.append('Notas e comentários foram preservados; o perfil não altera o texto desses elementos.')
    if body.xpath('.//wp:anchor'):
        warnings.append('Imagens flutuantes mantêm suas âncoras. Confira a posição no Word após alterar margens ou papel.')
    if body.xpath('.//w:tbl'):
        warnings.append('Tabelas mantêm larguras e estrutura originais. Confira o encaixe se reduziu a área útil da página.')
    if body.xpath('.//w:numPr') or any(p.style and p.style.name.startswith('List') for p in document.paragraphs):
        warnings.append('Listas mantêm a numeração original; os recuos de seus parágrafos seguem o perfil.')
    return warnings


def _set_onoff(parent, name: str, value: bool) -> None:
    element = parent.find(qn('w:' + name))
    if element is None:
        element = OxmlElement('w:' + name)
        if name == 'snapToGrid':
            parent.insert_element_before(element, 'w:spacing', 'w:ind', 'w:contextualSpacing',
                'w:mirrorIndents', 'w:suppressOverlap', 'w:jc', 'w:textDirection',
                'w:textAlignment', 'w:textboxTightWrap', 'w:outlineLvl', 'w:divId',
                'w:cnfStyle', 'w:rPr', 'w:sectPr', 'w:pPrChange')
        elif name == 'contextualSpacing':
            parent.insert_element_before(element, 'w:mirrorIndents', 'w:suppressOverlap',
                'w:jc', 'w:textDirection', 'w:textAlignment', 'w:textboxTightWrap',
                'w:outlineLvl', 'w:divId', 'w:cnfStyle', 'w:rPr', 'w:sectPr', 'w:pPrChange')
        else:
            element = getattr(parent, 'get_or_add_' + name)()
    element.set(qn('w:val'), '1' if value else '0')


def _format_run(run: Run, settings: FormatSettings) -> None:
    run.font.name = settings.font_name
    run.font.size = Pt(settings.font_size)
    rpr = run._r.get_or_add_rPr()
    fonts = rpr.get_or_add_rFonts()
    for key in list(fonts.attrib):
        if 'theme' in key.lower():
            del fonts.attrib[key]
    for slot in ('ascii', 'hAnsi', 'eastAsia', 'cs'):
        fonts.set(qn('w:' + slot), settings.font_name)
    size_cs = rpr.find(qn('w:szCs'))
    if size_cs is None:
        size_cs = OxmlElement('w:szCs')
        rpr.insert_element_before(size_cs, 'w:highlight', 'w:u', 'w:effect', 'w:bdr',
            'w:shd', 'w:fitText', 'w:vertAlign', 'w:rtl', 'w:cs', 'w:em', 'w:lang',
            'w:eastAsianLayout', 'w:specVanish', 'w:oMath', 'w:rPrChange')
    size_cs.set(qn('w:val'), str(int(settings.font_size * 2)))
    for key, complex_key in (('bold', 'bCs'), ('italic', 'iCs')):
        value = getattr(settings, key)
        if value is not None:
            setattr(run.font, key, value)
            _set_onoff(rpr, complex_key, value)
    if settings.underline is not None:
        run.font.underline = settings.underline
    if settings.font_color:
        run.font.color.rgb = RGBColor.from_string(settings.font_color)
        color = rpr.find(qn('w:color'))
        for key in list(color.attrib):
            if key != qn('w:val'):
                del color.attrib[key]


def _format_paragraph(paragraph: Paragraph, settings: FormatSettings) -> None:
    fmt = paragraph.paragraph_format
    fmt.alignment = ALIGNMENTS[settings.alignment]
    fmt.space_before = Pt(settings.paragraph_spacing_before)
    fmt.space_after = Pt(settings.paragraph_spacing_after)
    if settings.line_spacing_mode == 'multiple':
        fmt.line_spacing = settings.line_spacing
    else:
        fmt.line_spacing = Pt(settings.line_spacing)
        fmt.line_spacing_rule = WD_LINE_SPACING.EXACTLY if settings.line_spacing_mode == 'exact' else WD_LINE_SPACING.AT_LEAST
    fmt.left_indent = Cm(settings.left_indent_cm)
    fmt.right_indent = Cm(settings.right_indent_cm)
    fmt.first_line_indent = Cm(settings.first_line_indent_cm)
    for key in ('keep_with_next', 'keep_together', 'widow_control'):
        value = getattr(settings, key)
        if value is not None:
            setattr(fmt, key, value)
    ppr = paragraph._p.get_or_add_pPr()
    for key in ('beforeLines', 'afterLines'):
        ppr.spacing.attrib.pop(qn('w:' + key), None)
    for key in ('beforeAutospacing', 'afterAutospacing'):
        ppr.spacing.set(qn('w:' + key), '0')
    for key in ('leftChars', 'rightChars', 'firstLineChars', 'hangingChars', 'start', 'end', 'startChars', 'endChars'):
        ppr.ind.attrib.pop(qn('w:' + key), None)
    _set_onoff(ppr, 'contextualSpacing', False)
    _set_onoff(ppr, 'snapToGrid', False)
    # The paragraph mark determines empty-line height and newly entered text.
    old_mark = ppr.find(qn('w:rPr'))
    mark_run = OxmlElement('w:r')
    if old_mark is not None:
        mark_run.append(deepcopy(old_mark))
        ppr.remove(old_mark)
    _format_run(Run(mark_run, paragraph), settings)
    ppr.insert_element_before(mark_run.rPr, 'w:sectPr', 'w:pPrChange')


def _configure_sections(document: DocumentType, settings: FormatSettings, protected_sections=()) -> None:
    width, height = paper_dimensions(settings)
    for index, section in enumerate(document.sections):
        if index in protected_sections:
            continue
        section.orientation = WD_ORIENT.LANDSCAPE if settings.orientation == 'landscape' else WD_ORIENT.PORTRAIT
        section.page_width, section.page_height = Cm(width), Cm(height)
        section.top_margin, section.bottom_margin = Cm(settings.margin_top_cm), Cm(settings.margin_bottom_cm)
        section.left_margin, section.right_margin = Cm(settings.margin_left_cm), Cm(settings.margin_right_cm)
        section.header_distance, section.footer_distance = Cm(settings.header_distance_cm), Cm(settings.footer_distance_cm)


def _configure_header_footer(document: DocumentType, settings: FormatSettings) -> None:
    for slot in ('header', 'footer'):
        mode = getattr(settings, slot + '_mode')
        if mode == 'preserve':
            continue
        old_relationships = set()
        # Removing references from EVERY section also prevents inheritance.
        for section in document.sections:
            for ref in list(section._sectPr.findall(qn('w:' + slot + 'Reference'))):
                old_relationships.add(ref.get(qn('r:id')))
                section._sectPr.remove(ref)
        for rel_id in old_relationships:
            document.part.drop_rel(rel_id)
        if mode == 'remove':
            continue
        # Discard non-square DPI and apply EXIF orientation before computing geometry.
        with Image.open(Path(getattr(settings, slot + '_image_path')).expanduser()) as source:
            normalized = ImageOps.exif_transpose(source).convert('RGBA')
        image_data = BytesIO()
        normalized.save(image_data, format='PNG')
        normalized.close()
        for section in document.sections:
            for variant in (slot, 'first_page_' + slot, 'even_page_' + slot):
                container = getattr(section, variant)
                container.is_linked_to_previous = False
                paragraph = container.paragraphs[0]
                fmt = paragraph.paragraph_format
                fmt.left_indent = fmt.right_indent = fmt.first_line_indent = Cm(0)
                fmt.space_before = fmt.space_after = Pt(0)
                fmt.line_spacing = 1.0
                fmt.keep_with_next = fmt.keep_together = False
                paragraph.alignment = ALIGNMENTS[getattr(settings, slot + '_alignment')]
                # Never inherit typography/indents from the input's Header/Footer style.
                paragraph.style = None
                run = paragraph.add_run()
                run.font.size = Pt(1)
                image_data.seek(0)
                run.add_picture(image_data,
                                width=Cm(getattr(settings, slot + '_image_width_cm')))
                ppr = paragraph._p.get_or_add_pPr()
                _set_onoff(ppr, 'snapToGrid', False)
                _set_onoff(ppr, 'contextualSpacing', False)
                for key in ('beforeAutospacing', 'afterAutospacing'):
                    ppr.spacing.set(qn('w:' + key), '0')
                mark = OxmlElement('w:r')
                _format_run(Run(mark, paragraph), replace(settings, font_size=1, bold=False, italic=False, underline=False))
                ppr.append(mark.rPr)


def _prepare_output_dir(output_dir: Path) -> None:
    try:
        output_dir.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        raise FormatterError(f'Não foi possível acessar a pasta de saída: {exc}') from exc


def _verify_output(path: Path, settings: FormatSettings, expectations=None, expected_body=None, protected_sections=()) -> None:
    """Reopen the actual bytes, checking the profile before publishing success."""
    out = Document(str(path))
    width, height = paper_dimensions(settings)
    def check(condition: bool) -> None:
        if not condition:
            raise FormatterError('A verificação do Word gravado falhou; nenhum arquivo final foi publicado.')
    def length_matches(actual, expected) -> bool:
        return actual is not None and actual.twips == expected.twips
    if expected_body is not None:
        check(etree.tostring(out._element.body) == expected_body)
    for section_index, section in enumerate(out.sections):
        if section_index in protected_sections:
            continue
        for name, value in (('page_width', width), ('page_height', height),
                ('top_margin', settings.margin_top_cm), ('bottom_margin', settings.margin_bottom_cm),
                ('left_margin', settings.margin_left_cm), ('right_margin', settings.margin_right_cm),
                ('header_distance', settings.header_distance_cm), ('footer_distance', settings.footer_distance_cm)):
            check(length_matches(getattr(section, name), Cm(value)))
        for slot in ('header', 'footer'):
            mode = getattr(settings, slot + '_mode')
            if mode == 'remove':
                check(not section._sectPr.findall(qn('w:' + slot + 'Reference')))
            elif mode == 'image':
                for variant in (slot, 'first_page_' + slot, 'even_page_' + slot):
                    paragraph = getattr(section, variant).paragraphs[0]
                    extents = paragraph._p.xpath('.//wp:extent')
                    check(len(extents) == 1)
                    check(abs(int(extents[0].get('cx')) - Cm(getattr(settings, slot + '_image_width_cm'))) <= 1)
                    check(paragraph.alignment == ALIGNMENTS[getattr(settings, slot + '_alignment')])
    elements = out._element.body.xpath('.//w:p')
    if expectations is not None:
        check(len(elements) == len(expectations))
    for index, element in enumerate(elements):
        rule = expectations[index] if expectations is not None else settings
        if rule is None:
            continue
        settings = rule
        p = Paragraph(element, out._body)
        fmt = p.paragraph_format
        check(fmt.alignment == ALIGNMENTS[settings.alignment])
        for prop, expected in (('space_before', Pt(settings.paragraph_spacing_before)),
                ('space_after', Pt(settings.paragraph_spacing_after)), ('left_indent', Cm(settings.left_indent_cm)),
                ('right_indent', Cm(settings.right_indent_cm)), ('first_line_indent', Cm(settings.first_line_indent_cm))):
            check(length_matches(getattr(fmt, prop), expected))
        if settings.line_spacing_mode == 'multiple':
            check(fmt.line_spacing is not None and abs(fmt.line_spacing - settings.line_spacing) <= 1 / 480 + 1e-9)
        else:
            expected = WD_LINE_SPACING.EXACTLY if settings.line_spacing_mode == 'exact' else WD_LINE_SPACING.AT_LEAST
            check(fmt.line_spacing_rule == expected)
            check(length_matches(fmt.line_spacing, Pt(settings.line_spacing)))
        for prop in ('keep_with_next', 'keep_together', 'widow_control'):
            if getattr(settings, prop) is not None:
                check(getattr(fmt, prop) == getattr(settings, prop))
        runs = own_runs(element)
        properties = element.xpath('./w:pPr/w:rPr') + [run.rPr for run in runs if run.rPr is not None]
        check(len(properties) == len(runs) + 1)
        for rpr in properties:
            check(rpr.sz is not None and rpr.sz.val.pt == settings.font_size)
            check(rpr.rFonts is not None and all(rpr.rFonts.get(qn('w:' + slot)) == settings.font_name
                                               for slot in ('ascii', 'hAnsi', 'eastAsia', 'cs')))
            for key, tag in (('bold', 'b'), ('italic', 'i'), ('underline', 'u')):
                value = getattr(settings, key)
                if value is not None:
                    element_value = rpr.find(qn('w:' + tag))
                    check(element_value is not None)
                    if tag == 'u':
                        check(element_value.get(qn('w:val')) == ('single' if value else 'none'))
                    else:
                        check(element_value.val == value)
            if settings.font_color:
                check(rpr.color is not None and str(rpr.color.val) == settings.font_color)


def _save_exclusively(document: DocumentType, output_dir: Path, stem: str, settings: FormatSettings,
                      expectations=None, expected_body=None, protected_sections=()) -> Path:
    """Save fully, then publish without replacing any concurrent output."""
    # Limit filename component so generated suffixes fit common filesystems.
    stem = stem.encode('utf-8')[:180].decode('utf-8', errors='ignore')
    temp_path = None
    try:
        with tempfile.NamedTemporaryFile(prefix='.formatword-', suffix='.tmp', dir=output_dir, delete=False) as temp:
            temp_path = Path(temp.name)
            document.save(temp)
            temp.flush()
            os.fsync(temp.fileno())
        _verify_output(temp_path, settings, expectations, expected_body, protected_sections)
        counter = 1
        while True:
            extra = '' if counter == 1 else f'_{counter}'
            target = output_dir / f'{stem}{settings.output_suffix}{extra}.docx'
            try:
                # Same directory/volume; atomic and fails if the name already exists.
                os.link(temp_path, target)
                return target
            except FileExistsError:
                counter += 1
            except OSError as exc:
                if exc.errno not in {errno.ENOTSUP, errno.EPERM, errno.EXDEV, errno.ENOSYS}:
                    raise
                # FAT/exFAT and some network shares do not support hard links.
                # Reserve exclusively, never replace, and clean up a failed copy.
                try:
                    output = target.open('xb')
                except FileExistsError:
                    counter += 1
                    continue
                try:
                    with output, temp_path.open('rb') as source:
                        shutil.copyfileobj(source, output)
                        output.flush()
                        os.fsync(output.fileno())
                except BaseException:
                    target.unlink(missing_ok=True)
                    raise
                return target
    finally:
        if temp_path is not None:
            temp_path.unlink(missing_ok=True)
