"""Validate the generated documentation artifact and render it for review."""
from pathlib import Path
import hashlib
import json

import pymupdf as fitz
from PIL import Image, ImageDraw

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
PDF = HERE / 'Manual-Visual-Format-Word.pdf'
OUT = Path('/private/tmp/formatword-manual-review')
OUT.mkdir(exist_ok=True)


def main():
    metadata = json.loads((HERE / 'capturas.json').read_text())
    assert metadata['validation']['statuses'] == ['Concluído', 'Concluído com ressalvas', 'Erro']
    assert metadata['validation']['fictional_data']
    screenshots = [key for key in metadata if key != 'validation']
    for name in screenshots:
        with Image.open(HERE/'telas'/(name+'.png')) as image:
            image.verify()
    pdf = fitz.open(PDF)
    assert len(pdf) == 19
    assert not pdf.is_encrypted
    assert len(pdf.get_toc()) == 19
    assert len(pdf[1].get_links()) == 7
    full_text = '\n'.join(page.get_text() for page in pdf)
    for term in ['Importar de Word', 'Salvar perfil', 'Aplicar aos arquivos', 'Confirmar revisão',
                 'Concluído com ressalvas', 'Recuo especial', 'Cabeçalho e rodapé']:
        assert term in full_text, term
    for forbidden in ['/Users/', 'gabrielhenrique', '\ufffd']:
        assert forbidden not in full_text, forbidden
    assert all(page.get_images() for page in list(pdf)[:17])
    links = []
    for page in pdf:
        assert len(page.get_text().strip()) > 200
        for link in page.get_links():
            assert link['kind'] == fitz.LINK_GOTO
            assert 0 <= link['page'] < len(pdf)
            links.append(link['page']+1)
        used_fonts = {span['font'] for block in page.get_text('dict')['blocks']
                      if 'lines' in block for line in block['lines'] for span in line['spans']}
        for xref, _, _, font_name, *_ in page.get_fonts():
            if any(font_name.endswith(used) for used in used_fonts):
                assert pdf.extract_font(xref)[3], 'Used font must be embedded'
    # Detect text-text collisions in the source layout. Background boxes and
    # images are deliberately not treated as text; visual review covers them.
    layout = json.loads((HERE/'layout.json').read_text())
    boxes = layout['text_boxes']
    for i, (page, x, y, w, h) in enumerate(boxes):
        assert x >= 0 and y >= 0 and x+w <= pdf[page-1].rect.width
        for other in boxes[i+1:]:
            p2,x2,y2,w2,h2 = other
            if p2 != page:
                continue
            overlap_x = min(x+w,x2+w2)-max(x,x2)
            overlap_y = min(y+h,y2+h2)-max(y,y2)
            assert overlap_x < 1 or overlap_y < 1, ('Overlapping text', page, (x,y,w,h), other)
    digest = hashlib.sha256(PDF.read_bytes()).hexdigest()
    for filename in ['Manual-Visual-Format-Word.pdf', 'manual-de-uso-formatador.pdf']:
        assert hashlib.sha256((ROOT/'dist'/filename).read_bytes()).hexdigest() == digest
    for i,page in enumerate(pdf):
        page.get_pixmap(matrix=fitz.Matrix(1.6,1.6)).save(OUT/f'pagina-{i+1:02}.png')
    for first in range(0,len(pdf),6):
        sheet = Image.new('RGB',(1266,960),'#dae2ed')
        draw = ImageDraw.Draw(sheet)
        for j in range(first,min(first+6,len(pdf))):
            with Image.open(OUT/f'pagina-{j+1:02}.png') as img:
                img.thumbnail((610,290))
                x,y = (j-first)%2*633+10, (j-first)//2*320+22
                sheet.paste(img,(x,y))
                draw.text((x,y-17),f'Página {j+1}',fill='black')
        sheet.save(OUT/f'contato-{first+1}.png')
    report = dict(pages=len(pdf), screenshots=len(screenshots), bookmarks=len(pdf.get_toc()),
                  linked_pages=links, text_boxes=len(boxes), sha256=digest,
                  bytes=PDF.stat().st_size, distribution_copies_identical=True)
    (HERE/'validacao.json').write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n')
    print(json.dumps(report,ensure_ascii=False,indent=2))


if __name__ == '__main__':
    main()
