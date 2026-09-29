"""Small executable smoke test using only temporary documents and profiles."""
from __future__ import annotations

import json
import tempfile
import traceback
from pathlib import Path


def run_self_test(report_path: Path, *, gui: bool = True) -> int:
    report = {'ok': False, 'checks': []}
    try:
        from docx import Document
        from PIL import Image
        from app.batch import format_documents_batch
        from app.config import FormatSettings
        with tempfile.TemporaryDirectory(prefix='formatword-selftest-') as directory:
            root = Path(directory)
            picture = root / 'picture.png'
            Image.new('RGB', (200, 100), 'navy').save(picture)
            source = root / 'sample.docx'
            document = Document()
            document.add_heading('Título de teste', 1)
            document.add_paragraph('Corpo do documento')
            document.add_picture(str(picture))
            document.save(source)
            original = source.read_bytes()
            settings = FormatSettings(formatting_mode='by_category', body_image_width_cm=4,
                category_rules={'title': {'mode':'custom','values':{'font_size':18,'alignment':'center'}}})
            item = format_documents_batch([source],root/'output',settings)[0]
            if item.error or item.result is None:
                raise RuntimeError(item.error or 'Saída ausente')
            output = Document(item.result.output_path)
            if (output.paragraphs[0].runs[0].font.size.pt != 18
                    or output.paragraphs[1].runs[0].font.size.pt != 12
                    or abs(output.inline_shapes[0].width.cm - 4) > .001
                    or abs(output.inline_shapes[0].height.cm - 2) > .001
                    or source.read_bytes() != original):
                raise RuntimeError('Saída diferente das propriedades solicitadas')
            report['checks'].append('DOCX, categorias, imagem proporcional, lote e original intacto')
            if gui:
                from app.ui import FormatWordApp
                app = FormatWordApp(config_dir=root/'profiles')
                try:
                    app.withdraw()
                    app.update()
                    app.add_files([source])
                    app.profile_name.set('Perfil de teste')
                    if not app.save_profile():
                        raise RuntimeError('Falha ao salvar perfil pela interface')
                    app.update()
                finally:
                    app.destroy()
                report['checks'].append('Interface Tk, recursos empacotados e persistência isolada')
        report['ok'] = True
    except Exception:
        report['error'] = traceback.format_exc()
    report_path = Path(report_path)
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    return 0 if report['ok'] else 1
