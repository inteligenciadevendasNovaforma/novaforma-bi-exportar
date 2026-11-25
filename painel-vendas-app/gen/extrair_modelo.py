# Extrai da aba '(Data Aprovação) EVOLUÇÃO VENDA' do exemplo (sheet8; a aba (Data APP) tem o mesmo layout) os estilos,
# larguras e alturas e gera src/modeloPainelVendas.ts
# uso: python3 gen/extrair_modelo.py <pasta do xlsx descompactado> src/modeloPainelVendas.ts
import re, json, sys
SRC, OUT = sys.argv[1], sys.argv[2]
sheet = open(SRC + '/xl/worksheets/sheet8.xml', encoding='utf-8').read()
styles = open(SRC + '/xl/styles.xml', encoding='utf-8').read()
def blocos(tag):
    inner = re.search(r'<%s[^>]*>(.*?)</%s>' % (tag, tag), styles, re.S).group(1)
    sub = {'fonts': 'font', 'fills': 'fill', 'borders': 'border', 'cellXfs': 'xf', 'numFmts': 'numFmt'}[tag]
    return re.findall(r'<%s\b(?:[^>]*/>|.*?</%s>)' % (sub, sub), inner, re.S)
fonts, fills, borders, xfs = blocos('fonts'), blocos('fills'), blocos('borders'), blocos('cellXfs')
numfmts = {int(re.search(r'numFmtId="(\d+)"', n).group(1)): n for n in blocos('numFmts')}
# papéis das células na aba (conferidos com openpyxl):
#  titulo B2:M2 | faixaB B4/D4 (branco) | ufVend C4 | vend E4:G4 | fat H4:J4 | cart K4 | cli L4:N4 | mix O4:S4
#  cab B5:N5 | mFB..mCT O5:S5 | rep B | uf C | sup D | num E:K | qtd L:N
#  mix com valor: mixO (O, borda esq.), mixV (P:R), mixS (S, borda dir.); vazio (rosa): mixOV, mixVV, mixSV
PAPEIS = {'base': 0, 'linha': 1, 'titulo': 23, 'faixaB': 84, 'ufVend': 85, 'vend': 2, 'fat': 86, 'cart': 87, 'cli': 88, 'mix': 3,
          'cab': 4, 'mFB': 6, 'mFC': 7, 'mPT': 9, 'mAP': 10, 'mCT': 89, 'rep': 11, 'uf': 12, 'sup': 29, 'num': 90, 'qtd': 91,
          'mixO': 92, 'mixV': 93, 'mixVV': 94, 'mixS': 95, 'mixOV': 96, 'mixSV': 97}
usados = sorted(set(PAPEIS.values()))
fmap, filmap, bmap, nmap = {}, {0: 0, 1: 1}, {0: 0}, {}
nf, nfi, nb, nn = [], [fills[0], fills[1]], [borders[0]], []
idx = lambda a, xf: int(re.search(a + r'="(\d+)"', xf).group(1)) if re.search(a + r'="(\d+)"', xf) else 0
novo = []
for x in usados:
    xf = xfs[x]
    f, fi, b, n = idx('fontId', xf), idx('fillId', xf), idx('borderId', xf), idx('numFmtId', xf)
    if f not in fmap: fmap[f] = len(nf); nf.append(fonts[f])
    if fi not in filmap: filmap[fi] = len(nfi); nfi.append(fills[fi])
    if b not in bmap: bmap[b] = len(nb); nb.append(borders[b])
    if n >= 164 and n not in nmap: nmap[n] = 164 + len(nn); nn.append(re.sub(r'numFmtId="\d+"', 'numFmtId="%d"' % (164 + len(nn)), numfmts[n]))
    x2 = re.sub(r'fontId="\d+"', 'fontId="%d"' % fmap[f], xf)
    x2 = re.sub(r'fillId="\d+"', 'fillId="%d"' % filmap[fi], x2)
    x2 = re.sub(r'borderId="\d+"', 'borderId="%d"' % bmap[b], x2)
    if n >= 164: x2 = re.sub(r'numFmtId="\d+"', 'numFmtId="%d"' % nmap[n], x2)
    if n and 'applyNumberFormat' not in x2: x2 = x2.replace('<xf ', '<xf applyNumberFormat="1" ', 1)
    novo.append(re.sub(r'\s*xfId="\d+"', '', x2).replace('<xf ', '<xf xfId="0" ', 1))
xmap = {x: i for i, x in enumerate(usados)}
papeis = {k: xmap[v] for k, v in PAPEIS.items()}
# estilos extras (não existem na aba): nota do rodapé (cinza, itálico) e texto/valor/percentual da aba Dados (sem borda)
nf.append('<font><i/><sz val="9"/><color rgb="FF666666"/><name val="Arial"/></font>')
papeis['nota'] = len(novo); novo.append('<xf xfId="0" numFmtId="0" fontId="%d" fillId="0" borderId="0" applyFont="1"/>' % (len(nf) - 1))
nf.append('<font><sz val="10"/><color rgb="FF333333"/><name val="Arial"/></font>')
fD = len(nf) - 1
papeis['dTxt'] = len(novo); novo.append('<xf xfId="0" numFmtId="49" fontId="%d" fillId="0" borderId="0" applyNumberFormat="1" applyFont="1"/>' % fD)
papeis['dValor'] = len(novo); novo.append('<xf xfId="0" numFmtId="4" fontId="%d" fillId="0" borderId="0" applyNumberFormat="1" applyFont="1"/>' % fD)
papeis['dQtd'] = len(novo); novo.append('<xf xfId="0" numFmtId="1" fontId="%d" fillId="0" borderId="0" applyNumberFormat="1" applyFont="1"/>' % fD)
papeis['dPct'] = len(novo); novo.append('<xf xfId="0" numFmtId="10" fontId="%d" fillId="0" borderId="0" applyNumberFormat="1" applyFont="1"/>' % fD)
argb = lambda s: re.sub(r'rgb="00([0-9A-Fa-f]{6})"', r'rgb="FF\1"', s)
xml = ('<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    + ('<numFmts count="%d">%s</numFmts>' % (len(nn), ''.join(nn)) if nn else '')
    + '<fonts count="%d">%s</fonts>' % (len(nf), ''.join(nf))
    + '<fills count="%d">%s</fills>' % (len(nfi), ''.join(nfi))
    + '<borders count="%d">%s</borders>' % (len(nb), ''.join(nb))
    + '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    + '<cellXfs count="%d">%s</cellXfs>' % (len(novo), ''.join(novo))
    + '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>')
xml = argb(re.sub(r'>\s+<', '><', xml))
cols = re.sub(r'>\s+<', '><', re.search(r'<cols>(.*?)</cols>', sheet, re.S).group(1))
ht = lambda r: float(re.search(r'<row r="%d"[^>]*? ht="([\d.]+)"' % r, sheet).group(1))
alturas = {'r1': ht(1), 'titulo': ht(2), 'faixa': ht(3), 'grupo': ht(4), 'cab': ht(5), 'linha': ht(6)}
ts = ('// Gerado por gen/extrair_modelo.py a partir da aba (Data Aprovação) EVOLUÇÃO VENDA do exemplo (não editar à mão).\n'
      '// Estilos (cellXfs) remapeados; ESTILO = papel de cada célula; ALTURA = alturas das linhas (pt).\n'
      'export const COLS_XML = ' + json.dumps(cols) + ';\n'
      'export const STYLES_XML = ' + json.dumps(xml, ensure_ascii=False) + ';\n'
      'export const ESTILO = ' + json.dumps(papeis) + ';\n'
      'export const ALTURA = ' + json.dumps(alturas) + ';\n')
open(OUT, 'w', encoding='utf-8').write(ts)
print(papeis, alturas); print(xml[:400])
