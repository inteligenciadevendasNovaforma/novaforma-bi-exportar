# Extrai da aba '(Data do APP) PAINEL METAS | Re' do exemplo o molde do Excel (estilos, colunas, linhas, mesclagens) e gera src/modeloMetasRep.ts
# Uso: python3 gen/extrair_modelo.py <pasta com o xlsx descompactado> src/modeloMetasRep.ts
import re, json, sys, xml.etree.ElementTree as ET
SRC = sys.argv[1]; OUT = sys.argv[2]
NS = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
sheet = open(SRC + '/xl/worksheets/sheet6.xml', encoding='utf-8').read()
styles = open(SRC + '/xl/styles.xml', encoding='utf-8').read()
sst = ET.parse(SRC + '/xl/sharedStrings.xml').getroot()
strings = [''.join(t.text or '' for t in si.iter('{%s}t' % NS['m'])) for si in sst.findall('m:si', NS)]

def blocos(tag):
    inner = re.search(r'<%s[^>]*>(.*?)</%s>' % (tag, tag), styles, re.S).group(1)
    sub = {'fonts': 'font', 'fills': 'fill', 'borders': 'border', 'cellXfs': 'xf', 'numFmts': 'numFmt'}[tag]
    return re.findall(r'<%s\b(?:[^>]*/>|.*?</%s>)' % (sub, sub), inner, re.S)
fonts, fills, borders, xfs = blocos('fonts'), blocos('fills'), blocos('borders'), blocos('cellXfs')
numfmts = {int(re.search(r'numFmtId="(\d+)"', n).group(1)): n for n in blocos('numFmts')}

root = ET.fromstring(sheet)
rows = {}
for r in root.find('m:sheetData', NS).findall('m:row', NS):
    cells = []
    for c in r.findall('m:c', NS):
        col = re.match(r'[A-Z]+', c.get('r')).group(0)
        ci = 0
        for ch in col: ci = ci * 26 + ord(ch) - 64
        v = c.find('m:v', NS)
        val = None if v is None else (strings[int(v.text)] if c.get('t') == 's' else float(v.text))
        cells.append([ci, int(c.get('s', 0)), val])
    rows[int(r.get('r'))] = {'ht': float(r.get('ht')) if r.get('ht') else None, 's': int(r.get('s')) if r.get('s') else None, 'c': cells}
merges = [m.get('ref') for m in root.find('m:mergeCells', NS).findall('m:mergeCell', NS)]
cols = re.search(r'<cols>(.*?)</cols>', sheet, re.S).group(1)

# percentuais: 57 = vazio (texto rosa sobre rosa), 37 = valor (vermelho escuro sobre rosa), 73 = 100% ou mais (verde)
PCT = {'vazio': 57, 'valor': 37, 'meta': 73}
usados = sorted({c[1] for rn in range(1, 16) for c in rows[rn]['c']} | {rows[rn]['s'] for rn in range(1, 16) if rows[rn]['s'] is not None} | {0} | set(PCT.values()))
fmap, filmap, bmap, nmap = {}, {0: 0, 1: 1}, {0: 0}, {}
nf, nfi, nb, nn = [], [fills[0], fills[1]], [borders[0]], []
def idx(attr, xf):
    m = re.search(attr + r'="(\d+)"', xf); return int(m.group(1)) if m else 0
novoXf = []
for x in usados:
    xf = xfs[x]
    f, fi, b, n = idx('fontId', xf), idx('fillId', xf), idx('borderId', xf), idx('numFmtId', xf)
    if f not in fmap: fmap[f] = len(nf); nf.append(fonts[f])
    if fi not in filmap: filmap[fi] = len(nfi); nfi.append(fills[fi])
    if b not in bmap: bmap[b] = len(nb); nb.append(borders[b])
    if n >= 164 and n not in nmap: nmap[n] = 164 + len(nn); nn.append(re.sub(r'numFmtId="\d+"', 'numFmtId="%d"' % (164 + len(nn)), numfmts[n]))
    xf2 = re.sub(r'fontId="\d+"', 'fontId="%d"' % fmap[f], xf)
    xf2 = re.sub(r'fillId="\d+"', 'fillId="%d"' % filmap[fi], xf2)
    xf2 = re.sub(r'borderId="\d+"', 'borderId="%d"' % bmap[b], xf2)
    if n >= 164: xf2 = re.sub(r'numFmtId="\d+"', 'numFmtId="%d"' % nmap[n], xf2)
    xf2 = re.sub(r'\s*xfId="\d+"', '', xf2).replace('<xf ', '<xf xfId="0" ', 1)
    novoXf.append(xf2)
xmap = {x: i for i, x in enumerate(usados)}
argb = lambda s: re.sub(r'rgb="00([0-9A-Fa-f]{6})"', r'rgb="FF\1"', s)
stylesXml = ('<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    + ('<numFmts count="%d">%s</numFmts>' % (len(nn), ''.join(nn)) if nn else '')
    + '<fonts count="%d">%s</fonts>' % (len(nf), ''.join(nf))
    + '<fills count="%d">%s</fills>' % (len(nfi), ''.join(nfi))
    + '<borders count="%d">%s</borders>' % (len(nb), ''.join(nb))
    + '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    + '<cellXfs count="%d">%s</cellXfs>' % (len(novoXf), ''.join(novoXf))
    + '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>')
stylesXml = argb(re.sub(r'>\s+<', '><', stylesXml))

KPI = {'Objetivo Clientes': 'obj', 'Clientes(V)': 'cliV', 'Clientes (F)': 'cliF', '%': 'pct'}
TAB = {'Objetivo': 'obj', 'Vendido mês': 'vend', 'Vendido Ontem': 'ontem', 'Tendência': 'tend', 'Faturado': 'fat', 'Obj x Fat': 'objFat',
       'Carteira': 'cart', 'Faturado + Carteira': 'fatCart', 'Cliente Mês': 'cliMes', 'Clientes Ontem': 'cliOntem',
       'Clientes Faturados': 'cliFat', 'Clientes Carteira': 'cliCart'}
FAMS = ['FORRO E ACESS. BR', 'FORRO E ACESS. COLOR', 'PORTA PVC', 'ACESS. PARA PISO', 'CANTINHO']

def segmento(r0, r1):
    seg = []
    kpiHdr = tabHdr = None
    modo = None; fi = 0
    for rn in range(r0, r1 + 1):
        r = rows.get(rn, {'ht': None, 's': None, 'c': []})
        out = []
        byc = {c[0]: c for c in r['c']}
        b = byc.get(2, [2, 0, None])[2]
        txt = b.strip() if isinstance(b, str) else ''
        if txt.startswith('Objetivo Clientes'): kpiHdr = {c[0]: KPI.get(str(c[2]).strip()) for c in r['c'] if c[2] is not None}; modo = 'kpiH'
        elif txt == 'Família': tabHdr = {c[0]: TAB.get(str(c[2]).strip()) for c in r['c'] if c[2] is not None}; modo = 'tabH'
        elif modo == 'kpiH': modo = 'kpiV'
        elif modo == 'tabH': modo = 'tab'; fi = 0
        elif modo == 'kpiV': modo = None
        for ci, s, val in r['c']:
            k = None
            if modo == 'kpiV' and ci in kpiHdr and kpiHdr[ci]: k = 'k:' + kpiHdr[ci]
            elif modo == 'tab':
                linha = 'f%d' % fi if txt in FAMS else ('t' if txt == '' else None)
                if linha:
                    if ci == 2: k = linha + ':nome' if linha != 't' else None
                    elif ci in tabHdr and tabHdr[ci]: k = linha + ':' + tabHdr[ci]
            elif isinstance(val, str) and val.startswith('PAINEL DE VENDAS - MÊS |'): k = 'titulo'
            if k: out.append([ci, xmap[s], '=' + k])
            elif isinstance(val, str): out.append([ci, xmap[s], val])
            else: out.append([ci, xmap[s]])
        if modo == 'tab' and txt in FAMS: fi += 1
        if modo == 'tab' and txt == '' and any(c[2] is not None for c in r['c']): modo = None
        seg.append({'h': r['ht'], 's': xmap[r['s']] if r['s'] is not None else None, 'c': out})
    ms = []
    for m in merges:
        a, b2 = m.split(':')
        ra, rb = int(re.search(r'\d+', a).group(0)), int(re.search(r'\d+', b2).group(0))
        if r0 <= ra <= r1:
            ms.append([re.match(r'[A-Z]+', a).group(0), ra - r0, re.match(r'[A-Z]+', b2).group(0), rb - r0])
    return {'linhas': seg, 'merges': ms}

modelo = {'topo': segmento(1, 2), 'bloco': segmento(3, 15)}
pct = {k: xmap[v] for k, v in PCT.items()}
ts = ('// Gerado por gen/extrair_modelo.py a partir da aba "(Data do APP) PAINEL METAS | Re" do exemplo (não editar à mão).\n'
      '// Linhas: h = altura, s = estilo da linha, c = [coluna (1 = A), estilo, texto fixo ou "=chave"]\n'
      'export const COLS_XML = ' + json.dumps(re.sub(r'>\s+<', '><', cols)) + ';\n'
      'export const STYLES_XML = ' + json.dumps(stylesXml, ensure_ascii=False) + ';\n'
      '// estilos dos percentuais (% e Obj x Fat): vazio (rosa), valor (vermelho escuro sobre rosa), 100% ou mais (verde)\n'
      'export const ESTILO_PCT = ' + json.dumps(pct) + ';\n'
      'export interface LinhaModelo { h: number | null; s: number | null; c: (string | number)[][]; }\n'
      'export interface Segmento { linhas: LinhaModelo[]; merges: (string | number)[][]; }\n'
      'export const MODELO: { topo: Segmento; bloco: Segmento } = ' + json.dumps(modelo, ensure_ascii=False) + ';\n')
open(OUT, 'w', encoding='utf-8').write(ts)
print('ok', len(ts), 'estilos', len(novoXf), 'pct', pct)
for k, sg in modelo.items():
    keys = [c[2] for l in sg['linhas'] for c in l['c'] if len(c) > 2 and str(c[2]).startswith('=')]
    print(k, len(sg['linhas']), len(keys), keys[:8], sg['merges'])
