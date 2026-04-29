from __future__ import annotations

import json
import os
import shutil
import uuid
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any

from PIL import Image, ImageOps


APP_NAME = "FormatWord"
MAX_IMAGE_BYTES = 8 * 1024 * 1024
SUPPORTED_IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg"}
SUPPORTED_IMAGE_TYPES = {"header", "footer"}
MAX_TEMPLATE_BYTES = 25 * 1024 * 1024
HEADER_IMAGE_PIXELS = (1600, 220)
FOOTER_IMAGE_PIXELS = (1600, 180)
FONT_OPTIONS = (
    "Arial",
    "Times New Roman",
    "Calibri",
    "Cambria",
    "Georgia",
    "Verdana",
    "Tahoma",
    "Courier New",
)


@dataclass(slots=True)
class FormatSettings:
    font_name: str = "Arial"
    font_size: int = 12
    line_spacing: float = 1.5
    paragraph_spacing_after: int = 6
    first_line_indent_cm: float = 1.25
    margin_top_cm: float = 3.0
    margin_bottom_cm: float = 2.0
    margin_left_cm: float = 3.0
    margin_right_cm: float = 2.0
    justify_text: bool = True
    include_header: bool = False
    include_footer: bool = False
    header_image_path: str = ""
    footer_image_path: str = ""
    header_offset_x_cm: float = 0.0
    header_offset_y_cm: float = 0.0
    footer_offset_x_cm: float = 0.0
    footer_offset_y_cm: float = 0.0
    template_path: str = ""
    output_suffix: str = "_formatado"
    max_input_mb: int = 50


@dataclass(slots=True)
class AppConfig:
    settings: FormatSettings = field(default_factory=FormatSettings)
    stacks: dict[str, FormatSettings] = field(default_factory=dict)
    active_stack: str = ""


def get_config_dir() -> Path:
    if os.name == "nt":
        base = Path(os.environ.get("APPDATA", Path.home() / "AppData" / "Roaming"))
    elif sys_platform_is_macos():
        base = Path.home() / "Library" / "Application Support"
    else:
        base = Path(os.environ.get("XDG_CONFIG_HOME", Path.home() / ".config"))
    return base / APP_NAME


def sys_platform_is_macos() -> bool:
    return os.sys.platform == "darwin"


class ConfigStore:
    def __init__(self, config_dir: Path | None = None) -> None:
        self.config_dir = config_dir or get_config_dir()
        self.assets_dir = self.config_dir / "assets"
        self.templates_dir = self.config_dir / "templates"
        self.config_path = self.config_dir / "settings.json"
        self.config_dir.mkdir(parents=True, exist_ok=True)
        self.assets_dir.mkdir(parents=True, exist_ok=True)
        self.templates_dir.mkdir(parents=True, exist_ok=True)

    def load(self) -> AppConfig:
        if not self.config_path.exists():
            return AppConfig()

        try:
            data = json.loads(self.config_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return AppConfig()

        settings_data = data.get("settings", {})
        if not isinstance(settings_data, dict):
            return AppConfig()

        defaults = asdict(FormatSettings())
        merged = {**defaults, **self._filter_settings(settings_data)}
        stacks: dict[str, FormatSettings] = {}
        raw_stacks = data.get("stacks", {})
        if isinstance(raw_stacks, dict):
            for name, stack_data in raw_stacks.items():
                if isinstance(name, str) and isinstance(stack_data, dict):
                    stack_merged = {**defaults, **self._filter_settings(stack_data)}
                    stacks[name] = FormatSettings(**stack_merged)

        active_stack = data.get("active_stack", "")
        if not isinstance(active_stack, str) or active_stack not in stacks:
            active_stack = ""

        return AppConfig(settings=FormatSettings(**merged), stacks=stacks, active_stack=active_stack)

    def save(self, config: AppConfig) -> None:
        payload = {
            "settings": asdict(config.settings),
            "stacks": {name: asdict(settings) for name, settings in config.stacks.items()},
            "active_stack": config.active_stack if config.active_stack in config.stacks else "",
        }
        tmp_path = self.config_path.with_suffix(".tmp")
        tmp_path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
        tmp_path.replace(self.config_path)

    def store_image(self, image_path: Path, image_type: str) -> str:
        if image_type not in SUPPORTED_IMAGE_TYPES:
            raise ValueError("Tipo de imagem inválido.")

        image_path = image_path.expanduser().resolve()
        if image_path.suffix.lower() not in SUPPORTED_IMAGE_EXTENSIONS:
            raise ValueError("Use uma imagem PNG, JPG ou JPEG.")
        if not image_path.is_file():
            raise ValueError("Imagem não encontrada.")
        if image_path.stat().st_size > MAX_IMAGE_BYTES:
            raise ValueError("A imagem deve ter no máximo 8 MB.")
        if not self._has_valid_image_signature(image_path):
            raise ValueError("O conteúdo da imagem não parece ser PNG ou JPEG válido.")

        target = self.assets_dir / f"{image_type}_{uuid.uuid4().hex}.png"
        self._normalize_image(image_path, target, image_type)
        return str(target)

    def store_template(self, template_path: Path) -> str:
        template_path = template_path.expanduser().resolve()
        if template_path.suffix.lower() != ".docx":
            raise ValueError("Use um template Word no formato .docx.")
        if not template_path.is_file():
            raise ValueError("Template não encontrado.")
        if template_path.stat().st_size > MAX_TEMPLATE_BYTES:
            raise ValueError("O template deve ter no máximo 25 MB.")

        target = self.templates_dir / f"template_{uuid.uuid4().hex}.docx"
        shutil.copy2(template_path, target)
        return str(target)

    @staticmethod
    def _has_valid_image_signature(image_path: Path) -> bool:
        try:
            signature = image_path.read_bytes()[:12]
        except OSError:
            return False
        is_png = signature.startswith(b"\x89PNG\r\n\x1a\n")
        is_jpeg = signature.startswith(b"\xff\xd8\xff")
        return is_png or is_jpeg

    @staticmethod
    def _normalize_image(source: Path, target: Path, image_type: str) -> None:
        box = HEADER_IMAGE_PIXELS if image_type == "header" else FOOTER_IMAGE_PIXELS
        try:
            with Image.open(source) as image:
                image = ImageOps.exif_transpose(image).convert("RGBA")
                image.thumbnail(box, Image.Resampling.LANCZOS)
                canvas = Image.new("RGBA", box, (255, 255, 255, 0))
                x = (box[0] - image.width) // 2
                y = (box[1] - image.height) // 2
                canvas.alpha_composite(image, (x, y))
                canvas.save(target, format="PNG", optimize=True)
        except OSError as exc:
            raise ValueError("Não foi possível processar a imagem selecionada.") from exc

    @staticmethod
    def _filter_settings(raw: dict[str, Any]) -> dict[str, Any]:
        allowed = set(FormatSettings.__dataclass_fields__)
        return {key: value for key, value in raw.items() if key in allowed}
