"""Read a DOCX reference into reviewable profile alternatives without changing it."""
from __future__ import annotations

import colorsys
from dataclasses import dataclass, replace
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from lxml import etree

from app.config import CATEGORY_LABELS, TEXT_FIELDS, FormatSettings, SettingsError, validate_settings
from app.document_model import classify_paragraph, fingerprint, is_protected, own_runs
from app.formatter import FormatterError, validate_input


@dataclass(frozen=True, slots=True)
class ImportOption:
    label: str
    values: dict
    count: int = 1
    sample: str = ''


@dataclass(frozen=True, slots=True)
class ProfileImportReport:
    page_options: tuple[ImportOption, ...]
    category_options: dict[str, tuple[ImportOption, ...]]
    notices: tuple[str, ...]
    source: Path | None = None

    def build_settings(self, page_index=0, selections: dict[str, int] | None = None) -> FormatSettings:
        """Build a validated independent snapshot from the user's choices."""
        selections = {} if selections is None else selections
        if not isinstance(selections, dict) or set(selections) - set(self.category_options):
            raise SettingsError('category_rules', 'Seleção de categoria inválida.')
        def select(options, index, field):
            if type(index) is not int or not 0 <= index < len(options):
                raise SettingsError(field, 'Selecione uma alternativa válida para ' + field + '.')
            return dict(options[index].values)
        page = select(self.page_options, page_index, 'paper_size')
        chosen = {category: select(options, selections.get(category, 0), CATEGORY_LABELS[category])
                  for category, options in self.category_options.items()}
        settings = FormatSettings(formatting_mode='by_category', **page)
        if 'body' in chosen:
            settings = replace(settings, **chosen['body'])
        settings.category_rules = {
            category: ({'mode': 'inherit'} if category == 'body' else
                       {'mode': 'custom', 'values': chosen[category]}) if category in chosen
            else {'mode': 'preserve'} for category in CATEGORY_LABELS}
        try:
            return validate_settings(settings)
        except SettingsError as exc:
            raise SettingsError(exc.field, 'A alternativa importada não cabe no perfil: ' + str(exc)) from exc


def _val(node, name='val'):
    return node.get(qn('w:' + name)) if node is not None else None


def _on(node):
    return _val(node) not in ('0', 'false', 'off')


def _chain(style):
    chain, seen = [], set()
    while style is not None and style.style_id not in seen:
        seen.add(style.style_id)
        chain.append(style)
        style = style.base_style
    return reversed(chain)


class _Reader:
    def __init__(self, document):
        self.document = document
        self.notices = []
        self.defaults = document.styles.element.find(qn('w:docDefaults'))
        self.theme = None
        for rel in document.part.rels.values():
            if rel.reltype.endswith('/theme') and not rel.is_external:
                self.theme = etree.fromstring(rel.target_part.blob,
                    parser=etree.XMLParser(resolve_entities=False, no_network=True))
                break

    def notice(self, message):
        if message not in self.notices:
            self.notices.append(message)

    def default(self, tag):
        if self.defaults is None:
            return None
        container = self.defaults.find(qn('w:' + tag + 'Default'))
        return container.find(qn('w:' + tag)) if container is not None else None

    def theme_font(self, value):
        if self.theme is None or value not in ('majorAscii', 'majorHAnsi', 'minorAscii', 'minorHAnsi'):
            return None
        group = 'majorFont' if value.startswith('major') else 'minorFont'
        # The profile stores one font. Latin is the supported script here.
        node = self.theme.find('.//' + qn('a:fontScheme') + '/' + qn('a:' + group) + '/' + qn('a:latin'))
        return node.get('typeface') if node is not None else None

    def color(self, node):
        theme = _val(node, 'themeColor')
        value = _val(node)
        resolved_theme = False
        if theme:
            mapping_keys = {'background1': 'bg1', 'background2': 'bg2',
                            'text1': 't1', 'text2': 't2'}
            default_mapping = {'background1': 'light1', 'background2': 'light2',
                               'text1': 'dark1', 'text2': 'dark2'}
            mapping = self.document.settings.element.find(qn('w:clrSchemeMapping'))
            mapped = _val(mapping, mapping_keys.get(theme, theme)) or default_mapping.get(theme, theme)
            aliases = {'dark1': 'dk1', 'light1': 'lt1', 'dark2': 'dk2', 'light2': 'lt2',
                       'hyperlink': 'hlink', 'followedHyperlink': 'folHlink'}
            # Mapping values name actual theme slots, not another mapping key.
            tag = aliases.get(mapped, mapped)
            scheme = self.theme.find('.//' + qn('a:clrScheme')) if self.theme is not None else None
            target = scheme.find(qn('a:' + tag)) if scheme is not None else None
            if target is not None and len(target):
                base_color = target[0].get('lastClr') or target[0].get('val')
                if (target[0].tag in (qn('a:srgbClr'), qn('a:sysClr'))
                        and base_color and len(base_color) == 6
                        and all(c in '0123456789abcdefABCDEF' for c in base_color)):
                    value = base_color
                    resolved_theme = True
            if not resolved_theme:
                self.notice('Cor de tema não resolvida; a cor gravada no arquivo será usada quando disponível.')
        if value and value != 'auto':
            try:
                channels = [int(value[i:i + 2], 16) for i in (0, 2, 4)]
                if len(value) != 6:
                    raise ValueError()
                shade, tint = _val(node, 'themeShade'), _val(node, 'themeTint')
                if resolved_theme and (shade or tint):
                    # Word's MS-OI29500 17.3.2.6 uses HSL luminance and gives
                    # tint precedence over shade, rather than scaling RGB.
                    hue, lightness, saturation = colorsys.rgb_to_hls(*(c / 255 for c in channels))
                    lightness = (lightness * int(tint, 16) / 255 + (1 - int(tint, 16) / 255)
                                 if tint else lightness * int(shade, 16) / 255)
                    channels = [round(c * 255) for c in colorsys.hls_to_rgb(hue, lightness, saturation)]
                    self.notice('Tons e sombras de tema são convertidos em cor fixa; pode haver pequena diferença de arredondamento em relação ao Word.')
                return ''.join(f'{c:02X}' for c in channels)
            except ValueError:
                self.notice('Cor não representável; o perfil preservará a cor do documento de destino.')
        else:
            self.notice('Cor automática ou ausente: o perfil preservará a cor do documento de destino.')
        return ''

    def run_values(self, paragraph, run):
        values = {'font_name': None, 'font_size': None, 'bold': False, 'italic': False,
                  'underline': False, 'font_color': ''}
        layers = [(self.default('rPr'), False)]
        layers.extend((style.element.find(qn('w:rPr')), True) for style in _chain(paragraph.style))
        if run is not None:
            rpr = run.find(qn('w:rPr'))
            sid = _val(rpr.find(qn('w:rStyle'))) if rpr is not None else None
            if sid:
                try:
                    style = self.document.styles.get_by_id(sid, 2)
                    layers.extend((s.element.find(qn('w:rPr')), True) for s in _chain(style))
                except (KeyError, ValueError):
                    self.notice('Um estilo de caractere não pôde ser resolvido; revise a fonte e a ênfase.')
            layers.append((rpr, False))
        for rpr, toggle in layers:
            if rpr is None:
                continue
            fonts = rpr.find(qn('w:rFonts'))
            if fonts is not None:
                theme = _val(fonts, 'asciiTheme') or _val(fonts, 'hAnsiTheme')
                name = self.theme_font(theme) if theme else None
                if theme and not name:
                    self.notice('Fonte de tema não resolvida; revise o nome da fonte sugerida.')
                    values['font_name'] = None
                name = name or _val(fonts, 'ascii') or _val(fonts, 'hAnsi')
                if name:
                    values['font_name'] = name
            size = rpr.find(qn('w:sz'))
            if size is not None:
                values['font_size'] = float(_val(size)) / 2
            for name, tag in (('bold', 'b'), ('italic', 'i')):
                node = rpr.find(qn('w:' + tag))
                if node is not None:
                    if toggle:
                        if _on(node):
                            values[name] = not values[name]
                    else:
                        values[name] = _on(node)
            under = rpr.find(qn('w:u'))
            if under is not None:
                kind = _val(under) or 'single'
                values['underline'] = kind != 'none'
                if kind not in ('none', 'single'):
                    values['underline'] = None
                    self.notice('Sublinhado decorativo não cabe no perfil; será preservado no documento de destino.')
            color = rpr.find(qn('w:color'))
            if color is not None:
                values['font_color'] = self.color(color)
        for name, fallback in (('font_name', 'Arial'), ('font_size', 12)):
            if values[name] is None:
                values[name] = fallback
                label = 'Nome da fonte' if name == 'font_name' else 'Tamanho da fonte'
                self.notice(f'{label} não resolvido no documento; sugestão padrão {fallback}, revise antes de importar.')
        return values

    def paragraph_values(self, paragraph):
        values = dict(alignment='left', line_spacing_mode='multiple', line_spacing=1.0,
            paragraph_spacing_before=0.0, paragraph_spacing_after=0.0,
            first_line_indent_cm=0.0, left_indent_cm=0.0, right_indent_cm=0.0,
            keep_with_next=False, keep_together=False, widow_control=True)
        layers = [self.default('pPr')]
        layers.extend(style.element.find(qn('w:pPr')) for style in _chain(paragraph.style))
        layers.append(paragraph._p.find(qn('w:pPr')))
        line_value, line_rule = 240, 'auto'
        for ppr in layers:
            if ppr is None:
                continue
            alignment = ppr.find(qn('w:jc'))
            if alignment is not None:
                value = _val(alignment)
                if value in ('left', 'center', 'right', 'both'):
                    values['alignment'] = 'justify' if value == 'both' else value
                else:
                    self.notice(f'Alinhamento {value!r} não representável; sugerido alinhamento à esquerda.')
                    values['alignment'] = 'left'
            for name, tag in (('keep_with_next', 'keepNext'), ('keep_together', 'keepLines'), ('widow_control', 'widowControl')):
                node = ppr.find(qn('w:' + tag))
                if node is not None:
                    values[name] = _on(node)
            spacing = ppr.find(qn('w:spacing'))
            if spacing is not None:
                for attr, key in (('before', 'paragraph_spacing_before'), ('after', 'paragraph_spacing_after')):
                    value = _val(spacing, attr)
                    if value is not None:
                        values[key] = float(value) / 20
                line = _val(spacing, 'line')
                if line is not None:
                    line_value = float(line)
                line_rule = _val(spacing, 'lineRule') or line_rule
                values['line_spacing_mode'] = {'auto': 'multiple', 'exact': 'exact', 'atLeast': 'at_least'}.get(line_rule, line_rule)
                values['line_spacing'] = line_value / (240 if line_rule == 'auto' else 20)
                if any(_val(spacing, attr) for attr in ('beforeLines', 'afterLines', 'beforeAutospacing', 'afterAutospacing')):
                    self.notice('Espaçamento automático ou em linhas não cabe no perfil; revise os valores em pontos sugeridos.')
            indent = ppr.find(qn('w:ind'))
            if indent is not None:
                for attr, key in (('left', 'left_indent_cm'), ('right', 'right_indent_cm')):
                    value = _val(indent, attr)
                    if value is not None:
                        values[key] = float(value) * 2.54 / 1440
                first, hanging = _val(indent, 'firstLine'), _val(indent, 'hanging')
                if first is not None or hanging is not None:
                    values['first_line_indent_cm'] = (-float(hanging) if hanging is not None else float(first)) * 2.54 / 1440
                if any(_val(indent, attr) for attr in ('leftChars', 'rightChars', 'firstLineChars', 'hangingChars', 'start', 'end')):
                    self.notice('Recuos em caracteres ou direcionais não são reproduzidos; revise os recuos sugeridos em centímetros.')
            if ppr.find(qn('w:numPr')) is not None:
                self.notice('Numeração e recuos derivados de listas não são copiados; revise os recuos sugeridos.')
        return values


def _page_options(document, reader):
    options = []
    mapping = {'page_width': 'page_width_cm', 'page_height': 'page_height_cm',
        'top_margin': 'margin_top_cm', 'bottom_margin': 'margin_bottom_cm',
        'left_margin': 'margin_left_cm', 'right_margin': 'margin_right_cm',
        'header_distance': 'header_distance_cm', 'footer_distance': 'footer_distance_cm'}
    defaults = FormatSettings()
    labels = dict(zip(mapping.values(), ('largura da página', 'altura da página',
        'margem superior', 'margem inferior', 'margem esquerda', 'margem direita',
        'distância do cabeçalho', 'distância do rodapé')))
    for index, section in enumerate(document.sections):
        values = {}
        for prop, field in mapping.items():
            value = getattr(section, prop)
            values[field] = value.cm if value is not None else getattr(defaults, field)
            if value is None:
                reader.notice(f'Seção {index + 1}: {labels[field]} ausente; usado padrão do aplicativo, revise a página.')
        width, height = values['page_width_cm'], values['page_height_cm']
        values['orientation'] = 'landscape' if width > height else 'portrait'
        values['page_width_cm'], values['page_height_cm'] = sorted((width, height))
        values['paper_size'] = 'custom'
        for paper, dimensions in (('A4', (21, 29.7)), ('Letter', (21.59, 27.94))):
            if all(abs(a - b) < .003 for a, b in zip(sorted((width, height)), dimensions)):
                values['paper_size'] = paper
        label = f'Seção {index + 1}: {width:.2f} × {height:.2f} cm'
        options.append(ImportOption(label, values))
        if section._sectPr.xpath('./w:cols[@w:num>1]|./w:pgBorders|./w:docGrid|./w:pgMar[@w:gutter>0]'):
            reader.notice('Colunas, bordas, grade e medianiz não são copiadas pelo perfil.')
    if len(options) > 1:
        reader.notice('Escolha uma seção de referência; a página escolhida será aplicada a todas as seções do destino.')
    return tuple(options)


def inspect_profile(path, max_input_mb=50) -> ProfileImportReport:
    """Inspect supported properties, exposing ambiguity instead of changing the source."""
    path = Path(path).expanduser().resolve()
    validate_settings(FormatSettings(max_input_mb=max_input_mb))
    validate_input(path, max_input_mb)
    before = fingerprint(path)
    try:
        document = Document(str(path))
        reader = _Reader(document)
        reader.notice('Cabeçalho e rodapé NÃO são copiados da referência; Preservar original mantém os do documento de destino.')
        reader.notice('Categorias são identificadas por estilos e posição, sem inferir o significado do texto.')
        reader.notice('Propriedades de parágrafo ausentes usam os padrões do Word: esquerda, recuos/espaços zero e entrelinha simples.')
        reader.notice('Cor ausente ou automática será preservada no destino. O perfil representa uma única fonte por categoria; fontes por sistema de escrita não são reproduzidas separadamente.')
        pages = _page_options(document, reader)
        if not pages:
            raise FormatterError('O Word não contém uma seção de página válida para importar.')
        groups = {}
        for element in document._element.body.xpath('.//w:p'):
            if is_protected(element):
                reader.notice('Trechos com alterações controladas não participam da extração do perfil.')
                continue
            paragraph = Paragraph(element, document._body)
            runs = [(r, ''.join(t.text or '' for t in r.xpath('./w:t'))) for r in own_runs(element)]
            runs = [(r, text) for r, text in runs if text.strip()]
            if not runs:
                continue
            category = classify_paragraph(element, document, FormatSettings())
            props = reader.paragraph_values(paragraph)
            candidates = [(reader.run_values(paragraph, r), len(text), text) for r, text in runs]
            for emphasis in ('bold', 'italic', 'underline'):
                if len({v[emphasis] for v, _, _ in candidates}) > 1:
                    for values, _, _ in candidates:
                        values[emphasis] = None
                    reader.notice(f'{CATEGORY_LABELS[category]}: ênfase mista no mesmo parágrafo; negrito/itálico/sublinhado divergentes serão preservados no destino.')
            local = {}
            for run_props, weight, text in candidates:
                values = {**run_props, **props}
                key = tuple(values[name] for name in TEXT_FIELDS)
                entry = local.setdefault(key, [values, 0, text[:100]])
                entry[1] += weight
            for key, (values, weight, sample) in local.items():
                group = groups.setdefault(category, {})
                entry = group.setdefault(key, [values, 0, 0, sample])
                entry[1] += 1
                entry[2] += weight
        categories = {}
        for category, alternatives in groups.items():
            ranked = sorted(alternatives.values(), key=lambda x: (x[1], x[2]), reverse=True)
            categories[category] = tuple(ImportOption(
                f'{CATEGORY_LABELS[category]} {i + 1}: {v["font_name"]}, {v["font_size"]:g} pt — {count} parágrafo(s)',
                v, count, sample) for i, (v, count, weight, sample) in enumerate(ranked))
            if len(ranked) > 1:
                reader.notice(f'{CATEGORY_LABELS[category]}: há {len(ranked)} alternativas completas; escolha uma. A primeira tem maior frequência de parágrafos, com desempate por quantidade de texto.')
        if 'body' not in categories:
            reader.notice('Nenhum corpo de texto reconhecido; a categoria Corpo será preservada no documento de destino.')
        if document._element.body.xpath('.//w:tbl|.//w:drawing|.//w:pict|.//w:numPr|.//w:fldChar|.//w:fldSimple|.//m:oMath'):
            reader.notice('Imagens, objetos, campos, listas e estrutura de tabelas não são copiados; apenas as propriedades de texto suportadas são sugeridas.')
        if document._element.body.xpath('.//w:tbl'):
            reader.notice('Formatação de estilos condicionais de tabela não é resolvida; revise fonte, ênfase, cor e recuos sugeridos para Tabela.')
        if fingerprint(path) != before:
            raise FormatterError('O documento mudou durante a leitura. Importe novamente.')
        return ProfileImportReport(pages, categories, tuple(reader.notices), path)
    except FormatterError:
        raise
    except Exception as exc:
        raise FormatterError(f'Não foi possível importar o perfil do Word: {exc}') from exc
