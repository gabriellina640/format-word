from __future__ import annotations

import json
import os
import shutil
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any


APP_NAME = "FormatWord"
MAX_IMAGE_BYTES = 8 * 1024 * 1024
SUPPORTED_IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg"}
SUPPORTED_IMAGE_TYPES = {"header", "footer"}


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
    output_suffix: str = "_formatado"
    max_input_mb: int = 50


@dataclass(slots=True)
class AppConfig:
    settings: FormatSettings = field(default_factory=FormatSettings)


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
        self.config_path = self.config_dir / "settings.json"
        self.config_dir.mkdir(parents=True, exist_ok=True)
        self.assets_dir.mkdir(parents=True, exist_ok=True)

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
        return AppConfig(settings=FormatSettings(**merged))

    def save(self, config: AppConfig) -> None:
        payload = {"settings": asdict(config.settings)}
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

        target = self.assets_dir / f"{image_type}{image_path.suffix.lower()}"
        shutil.copy2(image_path, target)
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
    def _filter_settings(raw: dict[str, Any]) -> dict[str, Any]:
        allowed = set(FormatSettings.__dataclass_fields__)
        return {key: value for key, value in raw.items() if key in allowed}
