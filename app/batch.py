"""Sequential batch jobs independent from the graphical event loop."""
from __future__ import annotations

from dataclasses import dataclass
from copy import deepcopy
from pathlib import Path
from threading import Event
from typing import Callable, Iterable

from app.config import FormatSettings, validate_settings
from app.document_model import DocumentOverrides
from app.formatter import FormatResult, FormatterError, format_document


@dataclass(frozen=True, slots=True)
class BatchItem:
    input_path: Path
    result: FormatResult | None = None
    error: str | None = None
    cancelled: bool = False


def format_documents_batch(
    paths: Iterable[Path],
    output_dir: Path,
    settings: FormatSettings,
    cancel_event: Event | None = None,
    on_result: Callable[[BatchItem], None] | None = None,
    overrides_by_path: dict[Path, DocumentOverrides] | None = None,
) -> list[BatchItem]:
    inputs = tuple(Path(path) for path in paths)
    if not inputs:
        raise ValueError('Selecione ao menos um arquivo Word.')
    snapshot = validate_settings(settings, check_assets=True)
    reviews = {Path(path).expanduser().resolve(): deepcopy(review)
               for path, review in (overrides_by_path or {}).items()}
    cancel = cancel_event if cancel_event is not None else Event()
    results = []
    for source in inputs:
        if cancel.is_set():
            item = BatchItem(source, cancelled=True)
        else:
            try:
                item = BatchItem(source, result=format_document(source, output_dir, snapshot,
                    overrides=reviews.get(source.expanduser().resolve())))
            except FormatterError as exc:
                item = BatchItem(source, error=str(exc))
        results.append(item)
        if on_result is not None:
            on_result(item)
    return results
