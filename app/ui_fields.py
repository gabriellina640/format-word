"""Shared form schema, parsing and summaries; no dependency on Tk."""
from __future__ import annotations

from dataclasses import dataclass, replace
from typing import Any

from app.config import CATEGORY_LABELS, FormatSettings, SettingsError, paper_dimensions, validate_settings

TRISTATE = {'Preservar': None, 'Ativar': True, 'Desativar': False}
ALIGNMENT = {'Esquerda': 'left', 'Centro': 'center', 'Direita': 'right', 'Justificado': 'justify'}
MODES = {'Preservar original': 'preserve', 'Remover': 'remove', 'Imagem do perfil': 'image'}


@dataclass(frozen=True)
class FormField:
    name: str
    label: str
    group: str
    kind: str = 'number'
    choices: dict[str, Any] | None = None
    help: str = ''


FIELDS = (
    FormField('formatting_mode', 'Modo de formatação', 'Texto', 'choice',
              {'Uniforme': 'uniform', 'Por categoria': 'by_category'},
              help='Por categoria permite preservar ou personalizar títulos e outros tipos de texto.'),
    FormField('font_name', 'Fonte', 'Texto', 'font', help='O Word precisa ter a fonte instalada para exibi-la.'),
    FormField('font_size', 'Tamanho da fonte (pt)', 'Texto', help='De 1 a 400 pt; aceita meio ponto, como 12,5.'),
    FormField('bold', 'Negrito', 'Texto', 'choice', TRISTATE),
    FormField('italic', 'Itálico', 'Texto', 'choice', TRISTATE),
    FormField('underline', 'Sublinhado', 'Texto', 'choice', TRISTATE),
    FormField('font_color', 'Cor do texto (hexadecimal)', 'Texto', 'text', help='Vazio preserva a cor. Exemplo: #1F2937.'),
    FormField('alignment', 'Alinhamento', 'Parágrafo', 'choice', ALIGNMENT),
    FormField('line_spacing_mode', 'Tipo de entrelinhas', 'Parágrafo', 'choice',
              {'Múltiplo': 'multiple', 'Exato (pt)': 'exact', 'Mínimo (pt)': 'at_least'}),
    FormField('line_spacing', 'Entrelinhas (múltiplo ou pt)', 'Parágrafo', help='Múltiplo: 1,5 = uma linha e meia. Exato/mínimo: pontos.'),
    FormField('paragraph_spacing_before', 'Espaço antes do parágrafo (pt)', 'Parágrafo'),
    FormField('paragraph_spacing_after', 'Espaço após o parágrafo (pt)', 'Parágrafo'),
    FormField('first_line_indent_cm', 'Primeira linha / deslocado (cm)', 'Parágrafo', help='Positivo recua a primeira linha; negativo cria deslocamento.'),
    FormField('left_indent_cm', 'Recuo esquerdo (cm)', 'Parágrafo'),
    FormField('right_indent_cm', 'Recuo direito (cm)', 'Parágrafo'),
    FormField('keep_with_next', 'Manter com o próximo parágrafo', 'Parágrafo', 'choice', TRISTATE),
    FormField('keep_together', 'Manter linhas juntas', 'Parágrafo', 'choice', TRISTATE),
    FormField('widow_control', 'Controle de viúvas e órfãs', 'Parágrafo', 'choice', TRISTATE),
    FormField('paper_size', 'Papel', 'Página', 'choice', {'A4': 'A4', 'Carta': 'Letter', 'Personalizado': 'custom'}),
    FormField('orientation', 'Orientação', 'Página', 'choice', {'Retrato': 'portrait', 'Paisagem': 'landscape'}),
    FormField('page_width_cm', 'Lado menor do papel personalizado (cm)', 'Página', help='Usado apenas quando o papel Personalizado está selecionado.'),
    FormField('page_height_cm', 'Lado maior do papel personalizado (cm)', 'Página', help='Usado apenas quando o papel Personalizado está selecionado.'),
    FormField('margin_top_cm', 'Margem superior (cm)', 'Página'),
    FormField('margin_bottom_cm', 'Margem inferior (cm)', 'Página'),
    FormField('margin_left_cm', 'Margem esquerda (cm)', 'Página'),
    FormField('margin_right_cm', 'Margem direita (cm)', 'Página'),
    FormField('header_mode', 'Cabeçalho', 'Cabeçalho', 'choice', MODES),
    FormField('header_image_path', 'Imagem do cabeçalho', 'Cabeçalho', 'image'),
    FormField('header_image_width_cm', 'Largura da imagem (cm)', 'Cabeçalho', help='Só no modo Imagem do perfil. Altura proporcional + distância + 0,1 cm devem caber na margem.'),
    FormField('header_alignment', 'Alinhamento da imagem', 'Cabeçalho', 'choice', {k:v for k,v in ALIGNMENT.items() if v != 'justify'}),
    FormField('header_distance_cm', 'Distância da borda superior (cm)', 'Cabeçalho'),
    FormField('footer_mode', 'Rodapé', 'Rodapé', 'choice', MODES),
    FormField('footer_image_path', 'Imagem do rodapé', 'Rodapé', 'image'),
    FormField('footer_image_width_cm', 'Largura da imagem (cm)', 'Rodapé', help='Só no modo Imagem do perfil. Altura proporcional + distância + 0,1 cm devem caber na margem.'),
    FormField('footer_alignment', 'Alinhamento da imagem', 'Rodapé', 'choice', {k:v for k,v in ALIGNMENT.items() if v != 'justify'}),
    FormField('footer_distance_cm', 'Distância da borda inferior (cm)', 'Rodapé'),
    FormField('output_suffix', 'Sufixo do arquivo de saída', 'Saída', 'text', help='Exemplo: _formatado. Originais nunca são sobrescritos.'),
    FormField('max_input_mb', 'Limite por arquivo (MB)', 'Saída', 'integer'),
    FormField('body_image_width_cm', 'Largura das imagens do documento (cm)', 'Imagens',
              'optional_number', help='Vazio preserva a largura original. A altura mantém a proporção.'),
    FormField('body_image_alignment', 'Alinhamento das imagens do documento', 'Imagens', 'choice',
              {'Preservar': 'preserve', 'Esquerda': 'left', 'Centro': 'center', 'Direita': 'right'}),
    FormField('complex_content_mode', 'Conteúdo complexo', 'Saída', 'choice',
              {'Preservar trechos protegidos': 'preserve', 'Recusar documento': 'reject'}),
)
FIELD_MAP = {field.name: field for field in FIELDS}


def number_text(value: int | float) -> str:
    text = str(int(value)) if float(value).is_integer() else str(value)
    return text.replace('.', ',')


def form_values(settings: FormatSettings) -> dict[str, str]:
    result = {}
    for field in FIELDS:
        value = getattr(settings, field.name)
        if field.choices:
            result[field.name] = next(label for label, option in field.choices.items() if type(option) is type(value) and option == value)
        elif field.kind in {'number', 'integer', 'optional_number'}:
            result[field.name] = '' if value is None else number_text(value)
        else:
            result[field.name] = value
    return result


def settings_from_values(values: dict[str, str], base: FormatSettings | None = None,
                         check_assets: bool = False) -> FormatSettings:
    settings = replace(base) if base is not None else FormatSettings()
    for field in FIELDS:
        raw = values.get(field.name, '')
        try:
            if field.choices:
                value = field.choices[raw]
            elif field.kind in {'number', 'integer', 'optional_number'}:
                value = (None if field.kind == 'optional_number' and not raw.strip()
                         else float(raw.strip().replace(',', '.')))
                if field.kind == 'integer':
                    if not value.is_integer():
                        raise ValueError('Informe um número inteiro.')
                    value = int(value)
            else:
                value = raw.strip()
        except (ValueError, KeyError, OverflowError) as exc:
            raise SettingsError(field.name, f'{field.label}: informe um valor válido.') from exc
        setattr(settings, field.name, value)
    return validate_settings(settings, check_assets=check_assets)


def settings_summary(settings: FormatSettings) -> str:
    width, height = paper_dimensions(settings)
    values = form_values(settings)
    line = number_text(settings.line_spacing)
    line += ' ×' if settings.line_spacing_mode == 'multiple' else ' pt'
    mode_labels = {'inherit': 'Seguir corpo', 'preserve': 'Preservar', 'custom': 'Personalizar'}
    categories = ''
    if settings.formatting_mode == 'by_category':
        lines = []
        for category, label in CATEGORY_LABELS.items():
            rule = settings.category_rules.get(category, {})
            mode = rule.get('mode', 'inherit' if category == 'body' else 'preserve')
            line_label = f'{label}: {mode_labels[mode]}'
            if mode == 'custom':
                effective = replace(settings, **rule.get('values', {}))
                displayed = form_values(effective)
                changes = [f'{FIELD_MAP[name].label}: {displayed[name] or "Preservar"}'
                           for name in rule.get('values', {})]
                line_label += ' (' + '; '.join(changes) + ')'
            lines.append(line_label)
        categories = '\n'.join(lines) + '\n'
    return (
        f'Modo: {values["formatting_mode"]}\n' + categories +
        f'{settings.font_name} · {number_text(settings.font_size)} pt · {values["alignment"]}\n'
        f'Negrito: {values["bold"]} · Itálico: {values["italic"]} · Sublinhado: {values["underline"]} · '
        f'Cor: {"#" + settings.font_color if settings.font_color else "Preservar"}\n'
        f'Entrelinhas: {values["line_spacing_mode"]} {line} · Antes/depois: '
        f'{number_text(settings.paragraph_spacing_before)}/{number_text(settings.paragraph_spacing_after)} pt\n'
        f'Recuos E/D/primeira linha: {number_text(settings.left_indent_cm)}/'
        f'{number_text(settings.right_indent_cm)}/{number_text(settings.first_line_indent_cm)} cm\n'
        f'{values["paper_size"]} · {values["orientation"]} · {number_text(width)} × {number_text(height)} cm\n'
        f'Margens superior/inferior/esquerda/direita: '
        f'{number_text(settings.margin_top_cm)}/{number_text(settings.margin_bottom_cm)}/'
        f'{number_text(settings.margin_left_cm)}/{number_text(settings.margin_right_cm)} cm\n'
        f'Cabeçalho: {values["header_mode"]} · Rodapé: {values["footer_mode"]}\n'
        f'Distância cabeçalho/rodapé: {number_text(settings.header_distance_cm)}/{number_text(settings.footer_distance_cm)} cm\n'
        + (f'Imagem cabeçalho: {number_text(settings.header_image_width_cm)} cm · {values["header_alignment"]}\n'
           if settings.header_mode == 'image' else '')
        + (f'Imagem rodapé: {number_text(settings.footer_image_width_cm)} cm · {values["footer_alignment"]}\n'
           if settings.footer_mode == 'image' else '')
        + f'Manter com próximo: {values["keep_with_next"]} · Manter linhas: {values["keep_together"]} · '
        f'Viúvas/órfãs: {values["widow_control"]}\n'
        f'Sufixo: {settings.output_suffix} · Limite por arquivo: {settings.max_input_mb} MB\n'
        f'Imagens do documento: {values["body_image_width_cm"] + " cm" if settings.body_image_width_cm is not None else "Preservar largura"}'
        f' · {values["body_image_alignment"]}\n'
        f'Conteúdo complexo: {values["complex_content_mode"]}\n'
        + ('Aplicação: texto uniforme em todas as categorias e seções.'
           if settings.formatting_mode == 'uniform' else 'Aplicação: conforme as regras por categoria, em todas as seções.')
    )


def unique_profile_name(base: str, names: set[str]) -> str:
    name, index = base, 2
    while name in names:
        name = f'{base} {index}'
        index += 1
    return name
