"""Build the shareable visual guide from genuine app screenshots (no GUI needed)."""
from pathlib import Path
import json
import shutil
from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor, white
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import Paragraph

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
META = json.loads((HERE / 'capturas.json').read_text())
PDF = HERE / 'Manual-Visual-Format-Word.pdf'
W, H = landscape(A4)
NAVY, BLUE, MUTED = '#14243D', '#245FE8', '#50627A'
for name, filename in [('Manual', 'Arial.ttf'), ('ManualBold', 'Arial Bold.ttf')]:
    pdfmetrics.registerFont(TTFont(name, '/System/Library/Fonts/Supplemental/' + filename))
pdfmetrics.registerFontFamily('Manual', normal='Manual', bold='ManualBold', italic='Manual', boldItalic='ManualBold')
c = canvas.Canvas(str(PDF), pagesize=(W, H), pageCompression=1)
c.setTitle('Format Word — Manual visual de uso')
c.setAuthor('Format Word')
c.setSubject('Passo a passo com telas reais: perfis, importação de Word e formatação de documentos')
c.setViewerPreference('DisplayDocTitle', 'true')
checks = []


def text(value, x, top, width, size=11.5, color=NAVY, bold=False, leading=None):
    p = Paragraph(value, ParagraphStyle('body', fontName='ManualBold' if bold else 'Manual',
                  fontSize=size, leading=leading or size*1.35, textColor=HexColor(color)))
    _, height = p.wrap(width, H)
    assert top + height < H-10, (c.getPageNumber(), value[:80], top+height)
    p.drawOn(c, x, H-top-height)
    checks.append([c.getPageNumber(), x, top, width, height])
    return height


def rect(x, top, width, height, fill, radius=10):
    c.setFillColor(HexColor(fill))
    c.roundRect(x, H-top-height, width, height, radius, stroke=0, fill=1)


def circle(number, x, top, radius=10):
    c.setFillColor(HexColor(BLUE)); c.setStrokeColor(white); c.setLineWidth(1.6)
    c.circle(x, H-top, radius, stroke=1, fill=1)
    c.setFillColor(white); c.setFont('ManualBold', 10)
    c.drawCentredString(x, H-top-3.4, str(number))


def start(kicker, title, subtitle):
    c.setFillColor(white); c.rect(0, 0, W, H, fill=1, stroke=0)
    c.bookmarkPage('p'+str(c.getPageNumber()))
    c.addOutlineEntry(title, 'p'+str(c.getPageNumber()), level=0)
    text(kicker.upper(), 32, 24, W-64, 9, BLUE, True)
    text(title, 32, 44, W-64, 25, bold=True)
    text(subtitle, 32, 83, W-64, 11, MUTED)
    c.setStrokeColor(HexColor('#DEE6F1')); c.setLineWidth(.7)
    c.line(32, 39, W-32, 39)
    text('FORMAT WORD  •  MANUAL VISUAL  •  OUTUBRO 2026', 32, H-29, 600, 7.5, MUTED)
    c.setFont('ManualBold', 9); c.setFillColor(HexColor(MUTED))
    c.drawRightString(W-32, 19, str(c.getPageNumber()).zfill(2))


def screen(name, x=32, top=125, width=536, maxheight=405, crop=None, markers=()):
    m = META[name]
    crop = crop or (0, 0, m['width'], m['height'])
    left, upper, right, lower = crop
    scale = min(width/(right-left), maxheight/(lower-upper))
    dw, dh = (right-left)*scale, (lower-upper)*scale
    x += (width-dw)/2
    rect(x-3, top-3, dw+6, dh+6, '#DEE6F1', 5)
    c.saveState()
    path = c.beginPath(); path.rect(x, H-top-dh, dw, dh); c.clipPath(path, stroke=0)
    c.drawImage(str(HERE/'telas'/(name+'.png')), x-left*scale,
                H-top-(m['height']-upper)*scale,
                width=m['width']*scale, height=m['height']*scale, mask='auto')
    c.restoreState()
    for number, mx, my in markers:
        assert left <= mx <= right and upper <= my <= lower, (name, mx, my, crop)
        circle(number, x+(mx-left)*scale, top+(my-upper)*scale)
    return top+dh


def steps(items, x=599, top=132, width=209):
    for n, title, body in items:
        circle(n, x-14, top+7)
        height = text(title, x+3, top-1, width-3, 12, bold=True)
        height += text(body, x+3, top+height+6, width-3, 11)
        top += height+28
    return top


def note(title, body, top=437, x=596, width=214, tint='#ECF3FF'):
    style = ParagraphStyle('note', fontName='Manual', fontSize=10, leading=13.3)
    p = Paragraph('<b>'+title+'</b><br/>'+body, style)
    _, height = p.wrap(width-24, H)
    assert top+height+24 < 548, (c.getPageNumber(), title, top+height+24)
    rect(x, top, width, height+24, tint)
    text('<b>'+title+'</b><br/>'+body, x+12, top+12, width-24, 10, leading=13.3)


def ribbon(body, top=536):
    text(body, 32, top, W-64, 10, MUTED)


def end():
    c.showPage()


# 01 — cover and two paths
start('Guia para quem vai usar', 'Format Word', 'Manual visual • Prepare o padrão uma vez e reutilize nos próximos documentos.')
text('Do Word de referência<br/>ao documento formatado.', 32, 135, 338, 29, bold=True)
text('Siga as telas e os números. Você pode importar um arquivo já formatado, conferir o perfil e aplicar esse padrão a um ou vários documentos.', 32, 270, 332, 13)
screen('17-resultados', x=395, top=129, width=414, maxheight=319)
rect(32, 372, 340, 154, '#ECF3FF')
text('PRIMEIRO USO', 48, 389, 300, 10, BLUE, True)
text('Importar → conferir → salvar', 48, 413, 300, 17, bold=True)
text('Leia as páginas 2 a 5. Depois, gere seus documentos pela página 6.', 48, 448, 299, 12)
rect(395, 472, 414, 54, '#F2F6FA')
text('<b>Já tem um perfil salvo?</b> Vá direto à página 6.<br/>Ajustes detalhados: páginas 10 a 17.', 409, 483, 386, 11)
ribbon('Telas reais do aplicativo no macOS. Nomes, documentos e valores dos exemplos são fictícios.')
end()

# 02 — overview with navigable contents
start('Antes de começar', 'Entenda as duas áreas do aplicativo', 'Abra o Format Word recebido com a automação. Tenha seus documentos no formato .docx.')
screen('01-inicio', width=502, maxheight=382, markers=[(1, 654, 131), (2, 347, 131), (3, 725, 43)])
steps([(1, 'Perfis e formatação', 'Prepare fonte, parágrafo, página e outras regras. Um perfil é um conjunto de configurações reutilizáveis.'),
       (2, 'Arquivos e resultados', 'Adicione os documentos, escolha onde salvar e gere as cópias formatadas.'),
       (3, 'Perfil no topo', 'Escolha aqui o padrão salvo que deseja usar.')], x=573, width=234)
# Clickable contents: visible label and clickable bounds share coordinates.
x = 32
for label, page in [('Importar · 3',3), ('Salvar · 5',5), ('Aplicar · 6',6), ('Revisar · 7–8',7), ('Resultados · 9',9), ('Ajustes · 10',10), ('Dúvidas · 18',18)]:
    width = pdfmetrics.stringWidth(label, 'Manual', 10) + 22
    rect(x, 517, width, 24, '#ECF3FF', 5)
    text(label, x+11, 522, width-15, 10, BLUE)
    c.linkRect('', 'p'+str(page), (x,H-541,x+width,H-517), relative=0, thickness=0)
    x += width+7
end()

# 03 — import
start('Primeiro uso • 1 de 3', 'Comece com um Word de referência', 'Use um arquivo .docx que represente o padrão que sua equipe quer repetir.')
screen('02-perfis', markers=[(1, 654, 131), (2, 18, 252), (3, 17, 208)])
steps([(1, 'Abra a área de perfis', 'Clique em <b>Perfis e formatação</b>.'),
       (2, 'Importe o padrão', 'Clique em <b>Importar de Word</b>. Na janela de arquivos, selecione o .docx de referência e confirme a abertura.'),
       (3, 'Sem um modelo pronto?', 'Use <b>Novo perfil</b>, preencha os ajustes das páginas 10 a 17 e salve.')])
note('O que será aproveitado?', 'Configurações de texto e página suportadas. O conteúdo do modelo não é colocado nos documentos.', top=430)
end()

# 04 — review imported data
start('Primeiro uso • 2 de 3', 'Confira o que foi encontrado no Word', 'A leitura oferece um ponto de partida. Revise as alternativas antes de criar o perfil.')
screen('03-importar', maxheight=417, markers=[(1, 165, 135), (2, 165, 170), (3, 510, 262), (4, 714, 620)])
steps([(1, 'Configuração de página', 'Se houver várias seções, escolha a que representa o padrão. Papel e margens serão aplicados às seções do documento de destino.'),
       (2, 'Tipo de texto', 'Passe pelos tipos disponíveis. Confira a <b>Formatação encontrada</b> e o trecho de exemplo; escolha outra alternativa se necessário.'),
       (3, 'Avisos e limites', 'Leia esta aba e o resumo em <b>Configurações do perfil</b>. A importação não copia o documento inteiro.'),
       (4, 'Use as configurações', 'Clique em <b>Usar estas configurações</b>. O perfil ainda precisa ser salvo.')])
end()

# 05 — save
start('Primeiro uso • 3 de 3', 'Dê um nome ao perfil e salve', 'Exemplo: “Escritório — documentos gerais”. Use outro perfil para outro padrão.')
screen('05-salvar-perfil', markers=[(1, 112, 168), (2, 18, 337), (3, 118, 208)])
steps([(1, 'Nome do perfil', 'Digite um nome que a equipe reconheça. O nome ajuda a escolher o padrão certo nas próximas vezes.'),
       (2, 'Confira as seções', 'Abra <b>Fonte</b>, <b>Parágrafo</b> e as demais seções. Use <b>Anterior</b>/<b>Próximo</b> ou clique direto no nome da seção.'),
       (3, 'Salvar perfil', 'Clique em <b>Salvar</b> ou <b>Salvar perfil</b>: os dois salvam as mesmas configurações. Confira a indicação <b>Perfil salvo</b>.')])
note('Salvar ≠ gerar documentos', 'Salvar guarda o padrão para reutilizar. Para criar os arquivos formatados, siga a página 6.', top=435)
end()

# 06 — main daily flow
start('No dia a dia', 'Selecione os documentos e aplique', 'Este é o caminho que você repetirá depois de preparar o perfil.')
screen('13-arquivos', markers=[(1, 725, 43), (2, 18, 172), (3, 981, 410), (4, 18, 590)])
steps([(1, 'Escolha o perfil', 'No seletor superior, escolha o padrão salvo. Abra <b>Arquivos e resultados</b>.'),
       (2, 'Adicionar DOCX', 'Selecione um ou vários arquivos .docx. Os documentos aparecem na lista.'),
       (3, 'Pasta de destino', 'Clique em <b>Escolher</b> e selecione onde guardar as cópias. Confira <b>Configuração que será aplicada</b>; role o resumo para ler tudo.'),
       (4, 'Aplicar aos arquivos', 'Clique e aguarde o lote terminar. Veja os resultados na página 9. Para ajustes individuais, faça antes a revisão das páginas 7 e 8.')])
ribbon('A aplicação usa os campos atuais, inclusive alterações ainda não salvas. Salve o perfil antes para reutilizar o mesmo padrão.')
end()

# 07 — review text
start('Revisão opcional • antes de aplicar', 'Indique quais trechos são títulos ou citações', 'Na lista de arquivos, selecione um único documento e clique em Revisar documento.')
screen('14-revisar-textos', markers=[(1, 37, 134), (2, 42, 500), (3, 582, 532), (4, 824, 652)])
steps([(1, 'Selecione o trecho', 'Na aba <b>Tipos de texto</b>, marque uma linha ou várias. Um título digitado com estilo Normal pode aparecer como Corpo.'),
       (2, 'Escolha o tipo', 'Selecione <b>Título</b>, <b>Citação</b> ou outro tipo no campo abaixo da lista.'),
       (3, 'Classifique os selecionados', 'Antes de clicar, marque a caixa de mesmo estilo <b>somente</b> se quiser afetar outros documentos. Clique em <b>Classificar selecionados</b>. Se vinculou o estilo, salve o perfil depois.'),
       (4, 'Confirme a revisão', 'Clique em <b>Confirmar revisão</b> e depois em <b>Aplicar aos arquivos</b>. “Revisado” ainda não significa que o arquivo foi gerado.')])
ribbon('As regras por tipo funcionam no modo Por tipo de texto. A revisão individual fica somente na sessão atual.')
end()

# 08 — review image
start('Revisão opcional • antes de aplicar', 'Ajuste uma imagem específica do documento', 'Em Revisar documento, abra a aba Imagens e selecione uma imagem da lista.')
screen('15-revisar-imagens', markers=[(1, 37, 134), (2, 247, 355), (3, 247, 430), (4, 824, 652)])
steps([(1, 'Selecione a imagem', 'A edição está disponível para imagens <b>Em linha</b>. Imagens flutuantes devem ser ajustadas no Word.'),
       (2, 'Tamanho e alinhamento', 'Informe <b>Largura (cm)</b> e <b>Alinhamento</b>. A altura acompanha a proporção. Largura vazia preserva o tamanho original daquela imagem.'),
       (3, 'Mova, se precisar', 'Escolha um <b>Parágrafo de referência (corpo)</b> e <b>Antes</b> ou <b>Depois</b>. Para não mover, mantenha <b>Manter posição</b>.'),
       (4, 'Confirme e aplique', 'Clique em <b>Confirmar revisão</b>. De volta à tela principal, use <b>Aplicar aos arquivos</b>.')])
ribbon('Ao redimensionar, alinhar ou mover, a automação usa um parágrafo próprio para a imagem.')
end()

# 09 — results
start('Depois de aplicar', 'Leia o resultado e abra a cópia gerada', 'O lote pode conter sucessos e erros ao mesmo tempo. Confira cada documento que vai encaminhar.')
screen('17-resultados', markers=[(1, 983, 253), (2, 18, 645), (3, 18, 687)])
steps([(1, 'Veja a coluna Resultado', '<b>Concluído:</b> documento gerado.<br/><b>Concluído com ressalvas:</b> gerado, mas há pontos para conferir.<br/><b>Erro:</b> veja o motivo.<br/><b>Cancelado:</b> arquivo não processado.'),
       (2, 'Leia os detalhes', 'Selecione a linha. O painel inferior mostra o arquivo de saída, avisos ou o motivo do erro. Role esse painel quando necessário.'),
       (3, 'Abra o resultado', 'Use <b>Abrir DOCX selecionado</b> ou <b>Abrir pasta de saída</b>. Confira o documento no Word antes de encaminhar.')])
note('Se cancelar o lote', 'O arquivo em andamento termina; os próximos são cancelados. Um erro também não impede os demais arquivos.', top=444)
end()

# 10 — font
start('Ajustes do perfil • Fonte', 'Escolha como o texto deve aparecer', 'Perfis e formatação → Fonte. Role dentro do formulário para encontrar os campos abaixo.')
screen('06-fonte', top=128, width=536, maxheight=200, crop=(25,390,965,705), markers=[(1,33,456),(2,500,456)])
screen('06-fonte-detalhe', top=340, width=536, maxheight=188, crop=(25,390,965,705), markers=[(3,33,528)])
steps([(1, 'Modo de formatação', '<b>Todo o texto:</b> usa a mesma formatação nos tipos de texto.<br/><b>Por tipo de texto:</b> permite diferenciar títulos, corpo e citações (página 14).'),
       (2, 'Fonte e tamanho', 'Escolha a fonte e informe o tamanho em <b>pt</b>. A fonte precisa estar instalada no computador que abrirá o resultado.'),
       (3, 'Estilo e cor', 'Em negrito, itálico e sublinhado, escolha Preservar, Ativar ou Desativar. Use <b>Escolher cor…</b> ou <b>Manter original</b>.')])
note('Preservar', 'Mantém a característica que já existe no documento. Manter a cor original não é o mesmo que aplicar preto.', top=441)
end()

# 11 — paragraph
start('Ajustes do perfil • Parágrafo', 'Defina o espaçamento do texto', 'Perfis e formatação → Parágrafo. Os nomes seguem as opções familiares do Word.')
screen('07-paragrafo', top=133, width=536, maxheight=315, crop=(25,390,965,705), markers=[(1,33,456),(2,500,456),(3,33,607)])
steps([(1, 'Alinhamento', 'Escolha <b>Esquerda</b>, <b>Centralizado</b>, <b>Direita</b> ou <b>Justificado</b>.'),
       (2, 'Espaçamento entre linhas', 'Use os atalhos <b>Simples</b>, <b>1,5 linhas</b> ou <b>Duplo</b> para preencher de uma vez.'),
       (3, 'Outro valor em Em', '<b>Múltiplo:</b> número de linhas, como 1,5.<br/><b>Exatamente:</b> altura fixa em pontos.<br/><b>Pelo menos:</b> altura mínima em pontos.')])
rect(32, 354, 536, 120, '#F2F6FA')
text('ESPAÇO ENTRE PARÁGRAFOS', 48, 370, 500, 10, BLUE, True)
text('<b>Espaçamento antes</b> e <b>Espaçamento depois</b> controlam a distância antes e depois de cada parágrafo. São medidos em pontos (pt). Role para baixo para encontrar o campo “depois”.', 48, 397, 500, 12)
note('Vírgula ou ponto', 'Você pode digitar 1,5 ou 1.5. Confira a unidade ao lado do campo: pt e cm medem coisas diferentes.', top=438)
end()

# 12 — indent
start('Ajustes do perfil • Parágrafo', 'Ajuste os recuos', 'Na mesma seção Parágrafo, role para baixo até Recuo especial e os recuos laterais.')
screen('07-paragrafo-detalhe', top=135, width=536, maxheight=305, crop=(25,440,965,705), markers=[(1,500,484),(2,752,518),(3,33,651)])
steps([(1, 'Recuo especial', '<b>Nenhum:</b> sem recuo especial.<br/><b>Primeira linha:</b> desloca só a primeira linha.<br/><b>Deslocado:</b> desloca as linhas seguintes.'),
       (2, 'Por (cm)', 'Digite uma medida positiva. Exemplo: Primeira linha, Por 1,25 cm. Não precisa digitar sinal negativo para Deslocado.'),
       (3, 'Recuos laterais', '<b>Recuo à esquerda</b> e <b>Recuo à direita</b> afastam o parágrafo das margens, em centímetros.')])
rect(32, 354, 536, 124, '#ECF3FF')
text('Exemplo de leitura', 48, 370, 500, 13, bold=True)
text('“Primeira linha: 1,25 cm” → só o começo do parágrafo entra.<br/>“Recuo à esquerda: 2 cm” → o parágrafo inteiro entra.<br/><br/>Os números são ilustrativos. Use as medidas pedidas pela sua equipe.', 48, 399, 500, 11.5)
note('Ao terminar', 'Clique em Salvar perfil. Navegar entre as seções mantém os valores preenchidos.', top=437)
end()

# 13 — page
start('Ajustes do perfil • Página', 'Confira papel, orientação e margens', 'Perfis e formatação → Página. Estas configurações valem para as seções do documento de destino.')
screen('08-pagina', top=129, width=536, maxheight=191, crop=(25,390,965,705), markers=[(1,33,456),(2,500,456)])
screen('08-pagina-detalhe', top=333, width=536, maxheight=191, crop=(25,505,965,705), markers=[(3,33,544)])
steps([(1, 'Tamanho do papel', 'Escolha <b>A4</b>, <b>Carta</b> ou <b>Personalizado</b>. Os lados menor e maior só são usados no papel Personalizado.'),
       (2, 'Orientação', '<b>Retrato:</b> página vertical.<br/><b>Paisagem:</b> página horizontal.'),
       (3, 'Margens', 'Role e confira as margens <b>superior</b>, <b>inferior</b>, <b>esquerda</b> e <b>direita</b>, sempre em centímetros.')])
note('Confira no Word', 'Mudar margens pode alterar a paginação e o encaixe de tabelas, assinaturas e imagens.', top=432)
end()

# 14 — category rules
start('Ajustes do perfil • Tipos de texto', 'Dê tratamentos diferentes a títulos e citações', 'Em Perfis e formatação, clique em Tipos de texto… para abrir esta janela.')
screen('12-tipos', maxheight=412, markers=[(1,12,79),(2,12,115),(3,643,641)])
steps([(1, 'Escolha o tipo', 'Selecione Corpo, Título, Subtítulo, Citação, Assinatura, Tabela ou Caixa de texto.'),
       (2, 'Escolha a regra', '<b>Seguir corpo:</b> herda a regra de Corpo, inclusive seus ajustes personalizados.<br/><b>Preservar original:</b> mantém o texto do documento de destino.<br/><b>Personalizar:</b> permite preencher os campos abaixo.'),
       (3, 'Confirme e salve', 'Clique em <b>Confirmar formatação</b> e depois salve o perfil. O modo passa a <b>Por tipo de texto</b>.')])
note('O tipo precisa estar correto', 'Se um título aparecer como Corpo, use Revisar documento (página 7). Estas regras não são o editor completo de estilos do Word.', top=438)
end()

# 15 — header footer
start('Ajustes do perfil • Cabeçalho e rodapé', 'Decida o que fazer com o timbre', 'Perfis e formatação → Cabeçalho e rodapé. A mesma lógica vale para os dois elementos.')
screen('09-cabecalho', top=129, width=536, maxheight=191, crop=(25,390,965,705))
screen('09-cabecalho-detalhe', top=334, width=536, maxheight=191, crop=(25,390,965,705))
steps([(1, 'Preservar original', 'Mantém o cabeçalho ou rodapé do <b>documento que será formatado</b>. É a opção para um timbre que já está nele.'),
       (2, 'Remover', 'Remove o cabeçalho ou rodapé correspondente das cópias geradas.'),
       (3, 'Imagem do perfil', 'Escolha uma imagem PNG/JPEG, ajuste largura, alinhamento e distância da borda. Imagem e distância precisam caber na margem.')])
note('Referência ≠ timbre copiado', 'Importar de Word não copia cabeçalho, rodapé ou imagens do modelo. Configure a imagem separadamente, se precisar.', top=432)
end()

# 16 — image global
start('Ajustes do perfil • Imagens', 'Defina um padrão para as imagens', 'Perfis e formatação → Imagens. Para uma imagem específica, use a revisão da página 8.')
screen('10-imagens', top=139, width=536, maxheight=295, crop=(25,390,965,665), markers=[(1,33,456),(2,500,456)])
steps([(1, 'Largura das imagens', 'Informe a largura em centímetros para as imagens editáveis do documento. A altura mantém a proporção.'),
       (2, 'Alinhamento', 'Escolha Esquerda, Centralizado ou Direita. Use <b>Preservar</b> para manter o alinhamento original.'),
       (3, 'Sem mudança de tamanho', 'Deixe a largura vazia para preservar o tamanho original. Salve o perfil ao terminar.')])
rect(32, 336, 536, 129, '#F2F6FA')
text('O QUE CONFERIR NA CÓPIA', 48, 352, 500, 10, BLUE, True)
text('Confira se a imagem cabe entre as margens e se ficou perto do trecho correto. Imagens flutuantes ou em caixas de texto podem exigir ajuste no próprio Word.', 48, 380, 500, 12)
note('Cabeçalho e rodapé', 'As imagens desses elementos têm controles próprios na seção Cabeçalho e rodapé.', top=437)
end()

# 17 — advanced
start('Ajustes do perfil • Opções avançadas', 'Ajuste o que for necessário para sua rotina', 'Perfis e formatação → Opções avançadas. Você pode manter os valores atuais quando já atendem ao padrão.')
screen('11-avancadas', top=129, width=536, maxheight=191, crop=(25,390,965,705))
screen('11-avancadas-detalhe', top=334, width=536, maxheight=191, crop=(25,390,965,705))
steps([(1, 'Quebras de linha e de página', '<b>Manter com o próximo:</b> mantém o parágrafo junto ao seguinte.<br/><b>Manter linhas juntas:</b> evita dividir o parágrafo.<br/><b>Controle de viúvas/órfãs:</b> evita linhas isoladas no início ou fim de página.'),
       (2, 'Nome e limite', '<b>Sufixo do arquivo de saída</b> acrescenta um texto ao nome, como _formatado. <b>Limite por arquivo</b> define o tamanho máximo em MB.'),
       (3, 'Conteúdo complexo', 'Escolha preservar trechos protegidos ou recusar esses documentos. Leia os avisos que aparecerem no resultado.')])
end()

# 18 — help
start('Consulta rápida', 'Se algo não sair como você esperava', 'Confira o caso abaixo e volte ao passo indicado. As cópias geradas ficam na pasta de destino.')
rows = [
 ('O título ficou igual ao corpo.', 'Confira o modo Por tipo de texto e as regras em Tipos de texto… (página 14). Se o título usa o estilo Normal, classifique-o em Revisar documento (página 7).'),
 ('Apareceu “Concluído com ressalvas”.', 'Selecione a linha e leia o painel de detalhes. O arquivo foi gerado. Confira no Word os trechos citados, como tabelas e imagens (página 9).'),
 ('Apareceu “Erro” ou o arquivo não abre.', 'Leia o motivo no painel de detalhes. Use um .docx válido, sem senha. Arquivos .doc e PDF não são entradas. Confira o limite por arquivo e se a pasta de destino permite salvar.'),
 ('Um campo ficou inválido ao salvar.', 'O aplicativo abre a seção correspondente. Corrija o valor indicado, confira cm ou pt e tente salvar novamente. Vírgula e ponto decimal são aceitos.'),
 ('O resultado ficou diferente no Word.', 'Confira a fonte instalada, a paginação, as margens e os tipos de texto. A importação não clona o modelo inteiro nem reproduz toda a estrutura de tabelas e listas.'),
 ('Minha revisão individual desapareceu.', 'Ela vale para a sessão atual. Remover o arquivo, limpar a lista ou encerrar o aplicativo descarta esses ajustes. Se editar o original fora do aplicativo, revise-o novamente.'),
]
y=124
for index, (title, body) in enumerate(rows):
    rect(32,y,W-64,61,'#F2F6FA' if index%2==0 else '#FFFFFF',6)
    text(title,44,y+10,228,11.3,bold=True)
    text(body,288,y+9,505,10.4,leading=13)
    y+=67
end()

# 19 — handoff checklist
start('Para consultar sempre', 'Rotina pronta para repetir', 'Não é necessário recriar o perfil a cada documento.')
left = [(1,'Escolha o perfil certo','Confira o nome no topo e o resumo das configurações.'),
        (2,'Adicione os DOCX e escolha a pasta','Faça a revisão individual apenas quando o documento precisar.'),
        (3,'Aplique e leia os resultados','Abra cada cópia que será encaminhada. Resolva os avisos.'),
        (4,'Confira no Word','Revise títulos, citações, paginação, assinaturas, timbre, tabelas e imagens.')]
steps(left,x=59,top=135,width=331)
rect(426,127,383,193,'#ECF3FF')
text('Três coisas diferentes',442,144,345,16,bold=True)
text('<b>Word de referência:</b> exemplo de formatação usado para preencher o perfil.<br/><br/><b>Perfil salvo:</b> conjunto de regras para reutilizar.<br/><br/><b>Documento de entrada:</b> arquivo que receberá a formatação em uma nova cópia.',442,179,345,12)
rect(426,335,383,177,'#F2F6FA')
text('Cuidados com a organização',442,351,345,14,bold=True)
text('<b>Duplicar</b> ajuda a criar outro padrão. <b>Restaurar</b> recupera os valores salvos do perfil; <b>Excluir</b> apaga o perfil.<br/><br/><b>Remover selecionados</b> e <b>Limpar lista</b> só retiram itens da lista. Os originais não são apagados nem sobrescritos pela formatação.',442,382,345,11)
ribbon('As telas podem variar de aparência conforme o sistema. Use os nomes dos controles para localizar as mesmas ações.')
end()

c.save()
# Pages expose PDF bookmarks and the overview has internal navigation links.
(ROOT/'dist').mkdir(exist_ok=True)
shutil.copy2(PDF, ROOT/'dist'/PDF.name)
shutil.copy2(PDF, ROOT/'dist'/'manual-de-uso-formatador.pdf')
(HERE/'layout.json').write_text(json.dumps(dict(pages=c.getPageNumber()-1, text_boxes=checks),indent=2)+'\n')
print(f'{PDF}: {c.getPageNumber()-1} páginas, {PDF.stat().st_size:,} bytes')
