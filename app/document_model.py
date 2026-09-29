"""Read-only document inspection shared by review UI and formatting engine."""
from __future__ import annotations

import hashlib
import re
from dataclasses import dataclass, field
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn, nsmap
from lxml import etree
from docx.text.paragraph import Paragraph

from app.config import FormatSettings

REVISION_TAGS = tuple(qn('w:' + name) for name in (
    'ins', 'del', 'moveFrom', 'moveTo', 'pPrChange', 'rPrChange', 'sectPrChange', 'tblPrChange',
    'trPrChange', 'tcPrChange', 'cellIns', 'cellDel', 'cellMerge'))


@dataclass(frozen=True, slots=True)
class ParagraphInfo:
    index: int
    text: str
    style_id: str
    style_name: str
    category: str
    protected: bool
    in_body: bool = True


@dataclass(frozen=True, slots=True)
class ImageInfo:
    index: int
    label: str
    width_cm: float
    height_cm: float
    paragraph_index: int
    editable: bool
    floating: bool
    preview: bytes | None = None


@dataclass(frozen=True, slots=True)
class DocumentInspection:
    fingerprint: str
    paragraphs: tuple[ParagraphInfo, ...]
    images: tuple[ImageInfo, ...]
    notices: tuple[str, ...] = ()


@dataclass(frozen=True, slots=True)
class ImageEdit:
    width_cm: float | None = None
    alignment: str = 'preserve'
    target_paragraph: int | None = None
    placement: str = 'after'


@dataclass(frozen=True, slots=True)
class DocumentOverrides:
    fingerprint: str
    paragraph_categories: dict[int, str] = field(default_factory=dict)
    image_edits: dict[int, ImageEdit] = field(default_factory=dict)


def fingerprint(path: Path) -> str:
    with Path(path).open('rb') as source:
        return hashlib.file_digest(source, 'sha256').hexdigest()


def nearest_paragraph(element):
    return next((parent for parent in element.iterancestors() if parent.tag == qn('w:p')), None)


def own_runs(element):
    """Textbox paragraphs own their runs; never format them through an outer run."""
    return [run for run in element.xpath('.//w:r') if nearest_paragraph(run) is element]


def is_protected(element) -> bool:
    property_tags = {qn('w:' + tag) for tag in ('tblPr', 'trPr', 'tcPr')}
    return (any(node.tag in REVISION_TAGS for node in element.iter())
            or any(parent.tag in REVISION_TAGS or any(
                node.tag in REVISION_TAGS for prop in parent if prop.tag in property_tags for node in prop.iter())
                for parent in element.iterancestors()))


def xml_query(element, query):
    # wp:anchor and VML/textbox nodes are generic lxml elements, without docx's xpath override.
    return etree.XPath(query, namespaces=nsmap)(element)


def classify_paragraph(element, document, settings: FormatSettings) -> str:
    paragraph = Paragraph(element, document._body)
    style = paragraph.style
    style_map = getattr(settings, 'style_categories', {})
    if style is not None:
        for key in (style.style_id, style.name):
            if key in style_map:
                return style_map[key]
    ancestors = {parent.tag for parent in element.iterancestors()}
    if qn('w:txbxContent') in ancestors:
        return 'textbox'
    if qn('w:tc') in ancestors:
        return 'table'
    # Based on named styles, never on legal text content or visual guesses.
    visited = set()
    while style is not None and style.style_id not in visited:
        visited.add(style.style_id)
        name = re.sub(r'[\s_-]+', '', style.name.casefold())
        if name in {'title', 'título', 'titulo'} or re.fullmatch(r'(heading|título|titulo)\d+', name):
            return 'title'
        if name in {'subtitle', 'subtítulo', 'subtitulo'}:
            return 'subtitle'
        if name in {'quote', 'intensequote', 'citação', 'citacao', 'citaçãointensa'}:
            return 'quote'
        if name in {'signature', 'assinatura'}:
            return 'signature'
        style = style.base_style
    return 'body'


def image_elements(document):
    # Picture-only drawings: do not resize charts, diagrams or textbox shapes.
    return [el for el in document._element.body.xpath('.//wp:inline|.//wp:anchor')
            if xml_query(el, './a:graphic/a:graphicData/pic:pic')]


def inspect_document(path: Path, settings: FormatSettings) -> DocumentInspection:
    # Lazy import keeps the parsing helpers usable by the engine without a cycle.
    from app.formatter import FormatterError, validate_input
    path = Path(path).expanduser().resolve()
    validate_input(path, settings.max_input_mb)
    before = fingerprint(path)
    try:
        document = Document(str(path))
        elements = document._element.body.xpath('.//w:p')
        paragraphs = []
        for index, element in enumerate(elements):
            paragraph = Paragraph(element, document._body)
            text = ''.join(t.text or '' for run in own_runs(element) for t in run.xpath('./w:t'))
            style = paragraph.style
            paragraphs.append(ParagraphInfo(index, text, style.style_id if style else '',
                style.name if style else '', classify_paragraph(element, document, settings), is_protected(element),
                element.getparent() is document._element.body))
        images = []
        for index, element in enumerate(image_elements(document)):
            paragraph = nearest_paragraph(element)
            extent = element.find(qn('wp:extent'))
            width = int(extent.get('cx', '0')) / 360000 if extent is not None else 0
            height = int(extent.get('cy', '0')) / 360000 if extent is not None else 0
            properties = element.find(qn('wp:docPr'))
            label = properties.get('descr') or properties.get('name') if properties is not None else ''
            floating = element.tag == qn('wp:anchor')
            preview = None
            blips = xml_query(element, './a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip')
            if blips:
                rid = blips[0].get(qn('r:embed'))
                if rid and rid in document.part.related_parts:
                    preview = document.part.related_parts[rid].blob
            images.append(ImageInfo(index, label or f'Imagem {index + 1}', width, height,
                elements.index(paragraph) if paragraph in elements else -1,
                paragraph is not None and not is_protected(paragraph) and not floating and width > 0 and height > 0
                and not any(parent.tag == qn("w:txbxContent") for parent in element.iterancestors()),
                floating, preview))
        if before != fingerprint(path):
            raise FormatterError('O documento mudou durante a leitura. Abra a revisão novamente.')
        notices = []
        if any(p.protected for p in paragraphs):
            notices.append('Trechos com alterações controladas serão preservados, sem aceitar ou rejeitar revisões.')
        if any(i.floating for i in images):
            notices.append('Imagens flutuantes são preservadas; ajustes individuais exigem convertê-las para Em linha no Word.')
        return DocumentInspection(before, tuple(paragraphs), tuple(images), tuple(notices))
    except FormatterError:
        raise
    except Exception as exc:
        raise FormatterError(f'Não foi possível inspecionar o documento: {exc}') from exc
