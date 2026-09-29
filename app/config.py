"""Profile storage and validation.

Storage validates structure while retaining migration blockers and unavailable
asset paths, so one damaged resource cannot prevent recovery of other profiles.
UI profile-save actions and the formatting engine must separately call
``validate_settings(settings, check_assets=True)`` before accepting execution.
"""

from __future__ import annotations

import json
import math
import os
import re
import shutil
import tempfile
import uuid
from dataclasses import asdict, dataclass, field, replace
from pathlib import Path
from typing import Any

from PIL import Image, ImageOps, UnidentifiedImageError

APP_NAME = 'FormatWord'
DEFAULT_PROFILE_NAME = 'Configuração atual'
MAX_IMAGE_BYTES = 8 * 1024 * 1024
# A JPEG imported within 8 MiB can expand when normalized losslessly to RGBA PNG.
# The 40 MP pixel limit bounds decoded data; 160 MiB accommodates its stored form.
MAX_NORMALIZED_IMAGE_BYTES = 160 * 1024 * 1024
MAX_IMAGE_PIXELS = 40_000_000
SUPPORTED_IMAGE_EXTENSIONS = {'.png', '.jpg', '.jpeg'}
SUPPORTED_IMAGE_TYPES = {'header', 'footer'}
FONT_OPTIONS = ('Arial', 'Times New Roman', 'Calibri', 'Cambria', 'Georgia',
                'Verdana', 'Tahoma', 'Courier New')
CATEGORY_LABELS = {'body': 'Corpo', 'title': 'Título', 'subtitle': 'Subtítulo',
                   'quote': 'Citação', 'signature': 'Assinatura', 'table': 'Tabela',
                   'textbox': 'Caixa de texto'}
TEXT_FIELDS = ('font_name', 'font_size', 'bold', 'italic', 'underline', 'font_color',
               'alignment', 'line_spacing_mode', 'line_spacing',
               'paragraph_spacing_before', 'paragraph_spacing_after',
               'first_line_indent_cm', 'left_indent_cm', 'right_indent_cm',
               'keep_with_next', 'keep_together', 'widow_control')


@dataclass(slots=True)
class FormatSettings:
    font_name: str = 'Arial'
    font_size: float = 12
    alignment: str = 'justify'
    line_spacing_mode: str = 'multiple'
    line_spacing: float = 1.5
    paragraph_spacing_before: float = 0
    paragraph_spacing_after: float = 6
    first_line_indent_cm: float = 1.25
    left_indent_cm: float = 0
    right_indent_cm: float = 0
    margin_top_cm: float = 3
    margin_bottom_cm: float = 2
    margin_left_cm: float = 3
    margin_right_cm: float = 2
    paper_size: str = 'A4'
    page_width_cm: float = 21
    page_height_cm: float = 29.7
    orientation: str = 'portrait'
    bold: bool | None = None
    italic: bool | None = None
    underline: bool | None = None
    font_color: str = ''
    keep_with_next: bool | None = None
    keep_together: bool | None = None
    widow_control: bool | None = None
    header_mode: str = 'preserve'
    footer_mode: str = 'preserve'
    header_image_path: str = ''
    footer_image_path: str = ''
    header_image_width_cm: float = 15
    footer_image_width_cm: float = 15
    header_distance_cm: float = .8
    footer_distance_cm: float = .7
    header_alignment: str = 'center'
    footer_alignment: str = 'center'
    template_path: str = ''
    output_suffix: str = '_formatado'
    max_input_mb: int = 50
    migration_warnings: tuple[str, ...] = ()
    formatting_mode: str = 'uniform'
    category_rules: dict[str, dict] = field(default_factory=dict)
    style_categories: dict[str, str] = field(default_factory=dict)
    body_image_width_cm: float | None = None
    body_image_alignment: str = 'preserve'
    complex_content_mode: str = 'preserve'


@dataclass(slots=True)
class AppConfig:
    settings: FormatSettings = field(default_factory=FormatSettings)
    stacks: dict[str, FormatSettings] = field(default_factory=dict)
    active_stack: str = ''


class SettingsError(ValueError):
    def __init__(self, field: str, message: str):
        self.field = field
        super().__init__(message)


def paper_dimensions(settings: FormatSettings) -> tuple[float, float]:
    """Return page dimensions in cm, with orientation applied to the short/long edges."""
    sizes = {'A4': (21.0, 29.7), 'Letter': (21.59, 27.94)}
    width, height = sizes.get(settings.paper_size, (settings.page_width_cm, settings.page_height_cm))
    short, long = sorted((width, height))
    return (long, short) if settings.orientation == 'landscape' else (short, long)


def _number(settings: FormatSettings, name: str, low: float, high: float) -> None:
    value = getattr(settings, name)
    if isinstance(value, bool) or not isinstance(value, (float, int)):
        raise SettingsError(name, f'{name}: informe um número entre {low:g} e {high:g}.')
    try:
        valid = math.isfinite(value) and low <= value <= high
    except OverflowError:
        valid = False
    if not valid:
        raise SettingsError(name, f'{name}: informe um número finito entre {low:g} e {high:g}.')


def _load_image(path: Path, max_bytes: int | None = None) -> Image.Image:
    """Decode a bounded PNG/JPEG and normalize EXIF without cropping or padding."""
    if max_bytes is None:
        max_bytes = MAX_IMAGE_BYTES
    try:
        if not path.is_file():
            raise ValueError('Imagem não encontrada.')
        if path.stat().st_size > max_bytes:
            raise ValueError(f'A imagem deve ter no máximo {max_bytes / (1024 * 1024):g} MB.')
        with Image.open(path) as source:
            if source.format not in {'PNG', 'JPEG'}:
                raise ValueError('O conteúdo da imagem deve ser PNG ou JPEG.')
            if source.width * source.height > MAX_IMAGE_PIXELS:
                raise ValueError('A imagem deve ter no máximo 40 megapixels.')
            source.load()
            return ImageOps.exif_transpose(source).convert('RGBA')
    except (OSError, UnidentifiedImageError, Image.DecompressionBombError) as exc:
        raise ValueError('Não foi possível decodificar o conteúdo da imagem PNG ou JPEG.') from exc


def validate_settings(settings: FormatSettings, check_assets: bool = False) -> FormatSettings:
    """Validate a profile and return an independent, normalized snapshot."""
    return _validate_settings(settings, check_assets, allow_legacy=False)


def validate_persistable_settings(settings: FormatSettings) -> FormatSettings:
    """Validate storage structure, preserving blockers and unavailable assets."""
    return _validate_settings(settings, False, allow_legacy=True)


def _validate_settings(settings: FormatSettings, check_assets: bool, *, allow_legacy: bool,
                       validate_categories: bool = True) -> FormatSettings:
    if not isinstance(settings, FormatSettings):
        raise SettingsError('settings', 'Configuração inválida.')
    result = replace(settings)
    for name in ('font_name', 'font_color', 'output_suffix', 'template_path',
                 'header_image_path', 'footer_image_path'):
        value = getattr(result, name)
        if not isinstance(value, str) or any(ord(char) < 32 or ord(char) == 127 for char in value):
            raise SettingsError(name, f'{name}: texto inválido.')
    result.font_name = result.font_name.strip()
    if not result.font_name:
        raise SettingsError('font_name', 'Informe o nome da fonte.')
    if not re.fullmatch(r'[A-Za-z0-9_-]{1,50}', result.output_suffix):
        raise SettingsError('output_suffix', 'Sufixo: use 1 a 50 letras ASCII, números, _ ou -.')
    if result.font_color:
        if not re.fullmatch(r'#?[0-9a-fA-F]{6}', result.font_color):
            raise SettingsError('font_color', 'Cor: use RRGGBB ou #RRGGBB.')
        result.font_color = result.font_color.lstrip('#').upper()
    enums = {'alignment': ('left', 'center', 'right', 'justify'),
             'line_spacing_mode': ('multiple', 'exact', 'at_least'),
             'paper_size': ('A4', 'Letter', 'custom'),
             'orientation': ('portrait', 'landscape'),
             'header_mode': ('preserve', 'remove', 'image'),
             'footer_mode': ('preserve', 'remove', 'image'),
             'header_alignment': ('left', 'center', 'right'),
             'footer_alignment': ('left', 'center', 'right'),
             'formatting_mode': ('uniform', 'by_category'),
             'body_image_alignment': ('preserve', 'left', 'center', 'right'),
             'complex_content_mode': ('preserve', 'reject')}
    for name, options in enums.items():
        value = getattr(result, name)
        if not isinstance(value, str) or value not in options:
            raise SettingsError(name, f'{name}: escolha uma opção válida.')
    for name in ('bold', 'italic', 'underline', 'keep_with_next', 'keep_together', 'widow_control'):
        value = getattr(result, name)
        if value is not None and type(value) is not bool:
            raise SettingsError(name, f'{name}: escolha preservar, ativar ou desativar.')
    _number(result, 'font_size', 1, 400)
    if result.font_size * 2 != round(result.font_size * 2):
        raise SettingsError('font_size', 'Tamanho da fonte: use incrementos de 0,5 ponto.')
    limits = (.5, 10) if result.line_spacing_mode == 'multiple' else (1, 400)
    _number(result, 'line_spacing', *limits)
    for name in ('paragraph_spacing_before', 'paragraph_spacing_after'):
        _number(result, name, 0, 1584)
    for name in ('page_width_cm', 'page_height_cm'):
        _number(result, name, 1, 55.87)
    for name in ('margin_top_cm', 'margin_bottom_cm', 'margin_left_cm', 'margin_right_cm',
                 'left_indent_cm', 'right_indent_cm', 'header_distance_cm', 'footer_distance_cm'):
        _number(result, name, 0, 55.87)
    _number(result, 'first_line_indent_cm', -55.87, 55.87)
    for name in ('header_image_width_cm', 'footer_image_width_cm'):
        _number(result, name, .01, 55.87)
    if result.body_image_width_cm is not None:
        _number(result, 'body_image_width_cm', .01, 55.87)
    if type(result.max_input_mb) is not int or not 1 <= result.max_input_mb <= 2048:
        raise SettingsError('max_input_mb', 'Limite de entrada: use um inteiro entre 1 e 2048 MB.')
    width, height = paper_dimensions(result)
    body_width = width - result.margin_left_cm - result.margin_right_cm
    if body_width <= 0:
        raise SettingsError('margin_left_cm', 'As margens devem deixar largura útil positiva.')
    if result.body_image_width_cm is not None and result.body_image_width_cm > body_width:
        raise SettingsError('body_image_width_cm', 'A largura da imagem excede a largura útil da página.')
    if height - result.margin_top_cm - result.margin_bottom_cm <= 0:
        raise SettingsError('margin_top_cm', 'As margens devem deixar altura útil positiva.')
    paragraph_width = body_width - result.left_indent_cm - result.right_indent_cm
    if paragraph_width <= 0:
        raise SettingsError('left_indent_cm', 'Os recuos devem deixar largura de texto positiva.')
    if result.first_line_indent_cm >= paragraph_width:
        raise SettingsError('first_line_indent_cm', 'O recuo da primeira linha excede a largura de texto.')
    if result.first_line_indent_cm + result.left_indent_cm < -result.margin_left_cm:
        raise SettingsError('first_line_indent_cm', 'O recuo deslocado ultrapassa a borda da página.')
    for slot, margin in (('header', result.margin_top_cm), ('footer', result.margin_bottom_cm)):
        label = 'Cabeçalho' if slot == 'header' else 'Rodapé'
        distance = getattr(result, f'{slot}_distance_cm')
        if distance > margin:
            raise SettingsError(f'{slot}_distance_cm', f'{label}: a distância deve caber na margem correspondente.')
        if getattr(result, f'{slot}_mode') == 'image':
            image_width = getattr(result, f'{slot}_image_width_cm')
            if image_width > body_width:
                raise SettingsError(f'{slot}_image_width_cm', f'{label}: a largura da imagem excede a largura útil da página.')
            if check_assets:
                try:
                    with _load_image(Path(getattr(result, f'{slot}_image_path')).expanduser(),
                                     max_bytes=MAX_NORMALIZED_IMAGE_BYTES) as image:
                        image_height = image_width * image.height / image.width
                except ValueError as exc:
                    raise SettingsError(f'{slot}_image_path', f'{label}: {exc}') from exc
                # Reserve 1 mm for the inline image paragraph's baseline/line box.
                if image_height + distance + .1 > margin:
                    raise SettingsError(f'{slot}_image_width_cm', f'{label}: a altura proporcional da imagem, distância e folga de 0,1 cm excedem a margem.')
    if not isinstance(result.migration_warnings, (tuple, list)) or any(
            not isinstance(message, str) for message in result.migration_warnings):
        raise SettingsError('migration_warnings', 'Avisos de migração inválidos.')
    result.migration_warnings = tuple(result.migration_warnings)
    if result.template_path and not allow_legacy:
        raise SettingsError('template_path', 'Template legado: escolha um modo suportado e remova o template do perfil.')
    if result.migration_warnings and not allow_legacy:
        raise SettingsError('migration_warnings', 'Revise as opções legadas do perfil: ' + '; '.join(result.migration_warnings))
    if validate_categories:
        _validate_category_rules(result, allow_legacy=allow_legacy)
    return result


def _validate_category_rules(settings: FormatSettings, *, allow_legacy: bool) -> None:
    """Normalize bounded overrides against the profile's actual page geometry."""
    if not isinstance(settings.category_rules, dict):
        raise SettingsError('category_rules', 'As regras por categoria devem ser um objeto.')
    rules = {}
    for category, rule in settings.category_rules.items():
        if not isinstance(category, str) or category not in CATEGORY_LABELS:
            raise SettingsError('category_rules', 'Categoria desconhecida nas regras do perfil.')
        if not isinstance(rule, dict) or set(rule) - {'mode', 'values'}:
            raise SettingsError('category_rules', f'{CATEGORY_LABELS[category]}: regra inválida.')
        mode, values = rule.get('mode'), rule.get('values', {})
        if not isinstance(mode, str) or mode not in ('inherit', 'preserve', 'custom'):
            raise SettingsError('category_rules', f'{CATEGORY_LABELS[category]}: escolha um modo válido.')
        if not isinstance(values, dict) or set(values) - set(TEXT_FIELDS):
            raise SettingsError('category_rules', f'{CATEGORY_LABELS[category]}: propriedades de texto inválidas.')
        # Empty category maps and a private guard keep validation finite, even
        # when profiles hold multiple rules. Assets were checked once above.
        candidate = replace(settings, category_rules={}, style_categories={}, **values)
        try:
            normalized = _validate_settings(candidate, False, allow_legacy=allow_legacy,
                                            validate_categories=False)
        except SettingsError as exc:
            raise SettingsError('category_rules', f'{CATEGORY_LABELS[category]}: {exc}') from exc
        rules[category] = {'mode': mode, 'values': {
            name: getattr(normalized, name) for name in values}}
    if not isinstance(settings.style_categories, dict):
        raise SettingsError('style_categories', 'O mapa de estilos deve ser um objeto.')
    styles = {}
    for style, category in settings.style_categories.items():
        if (not isinstance(style, str) or not style.strip()
                or any(ord(char) < 32 or ord(char) == 127 for char in style)
                or not isinstance(category, str) or category not in CATEGORY_LABELS):
            raise SettingsError('style_categories', 'Informe um estilo e uma categoria válidos.')
        styles[style] = category
    settings.category_rules = rules
    settings.style_categories = styles


def settings_for_category(base: FormatSettings, category: str) -> FormatSettings | None:
    """Return an independent effective profile, or None to preserve the paragraph."""
    if not isinstance(category, str) or category not in CATEGORY_LABELS:
        raise SettingsError('category_rules', 'Categoria desconhecida.')
    result = validate_settings(base)
    if result.formatting_mode == 'uniform':
        return result
    rule = result.category_rules.get(category, {})
    mode = rule.get('mode', 'inherit' if category == 'body' else 'preserve')
    if mode == 'inherit' and category != 'body':
        rule = result.category_rules.get('body', {})
        mode = rule.get('mode', 'inherit')
    if mode == 'preserve':
        return None
    if mode == 'custom':
        for name, value in rule['values'].items():
            setattr(result, name, value)
    return result


def get_config_dir() -> Path:
    if os.name == 'nt':
        base = Path(os.environ.get('APPDATA', Path.home() / 'AppData' / 'Roaming'))
    elif os.sys.platform == 'darwin':
        base = Path.home() / 'Library' / 'Application Support'
    else:
        base = Path(os.environ.get('XDG_CONFIG_HOME', Path.home() / '.config'))
    return base / APP_NAME


class ConfigStore:
    def __init__(self, config_dir: Path | None = None) -> None:
        self.config_dir = Path(config_dir) if config_dir is not None else get_config_dir()
        self.assets_dir = self.config_dir / 'assets'
        self.templates_dir = self.config_dir / 'templates'
        self.config_path = self.config_dir / 'settings.json'
        self.warnings: list[str] = []
        self._backup_required = False
        self.config_dir.mkdir(parents=True, exist_ok=True)
        self.assets_dir.mkdir(parents=True, exist_ok=True)
        self.templates_dir.mkdir(parents=True, exist_ok=True)

    def _decode_settings(self, raw: Any, label: str) -> FormatSettings:
        if not isinstance(raw, dict):
            raise ValueError('O perfil deve ser um objeto JSON.')
        values = {key: value for key, value in raw.items() if key in FormatSettings.__dataclass_fields__}
        review = values.get('migration_warnings', ())
        if not isinstance(review, (tuple, list)) or any(not isinstance(item, str) for item in review):
            raise ValueError('Avisos de migração inválidos.')
        review = list(review)
        for old, new in (('justify_text', 'alignment'), ('include_header', 'header_mode'),
                         ('include_footer', 'footer_mode')):
            if old in raw:
                if type(raw[old]) is not bool:
                    raise ValueError(f'{old}: valor legado inválido.')
                if new not in values:
                    values[new] = ('justify' if raw[old] else 'left') if old == 'justify_text' else (
                        'image' if raw[old] else 'preserve')
        for key in ('header_offset_x_cm', 'header_offset_y_cm', 'footer_offset_x_cm', 'footer_offset_y_cm'):
            if key in raw and raw[key] != 0:
                review.append(f'{key}={raw[key]!r}: defina distância, alinhamento e largura da imagem.')
        known = set(FormatSettings.__dataclass_fields__) | {
            'justify_text', 'include_header', 'include_footer', 'header_offset_x_cm',
            'header_offset_y_cm', 'footer_offset_x_cm', 'footer_offset_y_cm'}
        unknown = sorted(set(raw) - known)
        if unknown:
            review.append('Opções desconhecidas exigem revisão: ' + ', '.join(unknown))
        values['migration_warnings'] = tuple(review)
        settings = FormatSettings(**values)
        normalized = validate_persistable_settings(settings)
        if normalized.template_path:
            self.warnings.append(f'{label}: template legado requer revisão; arquivo preservado: {normalized.template_path}')
        self.warnings.extend(f'{label}: {message}' for message in review)
        return normalized

    def _backup(self) -> None:
        if self._backup_required and self.config_path.exists():
            target = self.config_dir / f'settings.{uuid.uuid4().hex}.bak'
            shutil.copy2(self.config_path, target)
            self._backup_required = False

    def load(self) -> AppConfig:
        self.warnings = []
        if not self.config_path.exists():
            return AppConfig()
        self._backup_required = False
        try:
            data = json.loads(self.config_path.read_text(encoding='utf-8'))
            if not isinstance(data, dict):
                raise ValueError('A raiz deve ser um objeto JSON.')
        except (OSError, UnicodeError, ValueError) as exc:
            self.warnings.append(f'Configuração recuperada com padrões: {exc}')
            self._backup_required = True
            self._try_backup()
            return AppConfig()
        try:
            settings = self._decode_settings(data.get('settings', {}), 'Configuração atual')
        except ValueError as exc:
            settings = FormatSettings()
            self.warnings.append(f'Configuração atual inválida; padrões recuperados: {exc}')
        stacks = {}
        renamed = {}
        raw_stacks = data.get('stacks', {})
        if not isinstance(raw_stacks, dict):
            self.warnings.append('Lista de perfis inválida; perfis originais preservados no backup.')
        else:
            for name, raw in raw_stacks.items():
                try:
                    if not isinstance(name, str) or not name.strip():
                        raise ValueError('Nome de perfil vazio.')
                    target_name = name
                    if name == DEFAULT_PROFILE_NAME:
                        base = target_name = f'{name} (perfil)'
                        counter = 2
                        while target_name in raw_stacks or target_name in stacks:
                            target_name = f'{base} {counter}'
                            counter += 1
                        renamed[name] = target_name
                        self.warnings.append(f'Perfil "{name}" renomeado para "{target_name}" para evitar conflito com a configuração atual.')
                    stacks[target_name] = self._decode_settings(raw, name)
                except ValueError as exc:
                    self.warnings.append(f'Perfil {name!r} inválido, preservado no backup: {exc}')
        active = data.get('active_stack', '')
        if isinstance(active, str):
            active = renamed.get(active, active)
        if not isinstance(active, str) or active not in stacks:
            active = ''
        if self.warnings:
            self._backup_required = True
            self._try_backup()
        return AppConfig(settings, stacks, active)

    def _try_backup(self) -> None:
        try:
            self._backup()
        except OSError as exc:
            self.warnings.append(f'Não foi possível criar o backup; gravação bloqueada até backup: {exc}')

    def save(self, config: AppConfig) -> None:
        """Persist all profiles without requiring unrelated assets or migrations."""
        settings = validate_persistable_settings(config.settings)
        if not isinstance(config.stacks, dict):
            raise ValueError('Lista de perfis inválida.')
        stacks = {}
        for name, profile in config.stacks.items():
            if not isinstance(name, str) or not name.strip():
                raise ValueError('Informe um nome de perfil não vazio.')
            if name == DEFAULT_PROFILE_NAME:
                raise ValueError(f'O nome "{DEFAULT_PROFILE_NAME}" é reservado.')
            stacks[name] = asdict(validate_persistable_settings(profile))
        payload = {'settings': asdict(settings), 'stacks': stacks,
                   'active_stack': config.active_stack if isinstance(config.active_stack, str)
                   and config.active_stack in stacks else ''}
        self._backup()
        descriptor, name = tempfile.mkstemp(prefix='.settings-', suffix='.tmp', dir=self.config_dir)
        temporary = Path(name)
        try:
            with os.fdopen(descriptor, 'w', encoding='utf-8') as output:
                json.dump(payload, output, indent=2, ensure_ascii=False, allow_nan=False)
                output.flush()
                os.fsync(output.fileno())
            temporary.replace(self.config_path)
        finally:
            temporary.unlink(missing_ok=True)

    def store_image(self, image_path: Path, image_type: str) -> str:
        if image_type not in SUPPORTED_IMAGE_TYPES:
            raise ValueError('Tipo de imagem inválido.')
        image_path = Path(image_path).expanduser().resolve()
        if image_path.suffix.lower() not in SUPPORTED_IMAGE_EXTENSIONS:
            raise ValueError('Use uma imagem PNG, JPG ou JPEG.')
        target = self.assets_dir / f'{image_type}_{uuid.uuid4().hex}.png'
        try:
            with _load_image(image_path) as image:
                image.save(target, format='PNG', optimize=True)
        except Exception:
            target.unlink(missing_ok=True)
            raise
        return str(target)
