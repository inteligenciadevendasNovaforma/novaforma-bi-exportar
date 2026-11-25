// Gerador mínimo de .xlsx (sem bibliotecas): zip sem compressão + XML das planilhas.
export interface LinhaDado { mes: string; unidade: string; valor: number; }
export interface ImagemXlsx { png: Uint8Array; w: number; h: number; }
export interface EntradaXlsx {
    titulo: string;
    geradoEm: string;
    kpis: { rotulo: string; valor: string }[];
    imagens: ImagemXlsx[];
    meses: string[];
    unidades: string[];
    matriz: number[][]; // [unidade][mes]
    linhas: LinhaDado[];
    opcoes: OpcoesExport;
}
export interface OpcoesExport { cards: boolean; graficos: boolean; matriz: boolean; dados: boolean; }

const enc = new TextEncoder();
const CRC: number[] = (() => {
    const t: number[] = [];
    for (let n = 0; n < 256; n++) {
        let c = n;
        for (let k = 0; k < 8; k++) { c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; }
        t[n] = c >>> 0;
    }
    return t;
})();
function crc32(b: Uint8Array): number {
    let c = 0xFFFFFFFF;
    for (let i = 0; i < b.length; i++) { c = CRC[(c ^ b[i]) & 255] ^ (c >>> 8); }
    return (c ^ 0xFFFFFFFF) >>> 0;
}
function esc(s: string): string {
    return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
function letra(i: number): string {
    let s = "";
    let n = i + 1;
    while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
    return s;
}
function sTxt(ref: string, texto: string, estilo = 0): string {
    return `<c r="${ref}" t="inlineStr" s="${estilo}"><is><t xml:space="preserve">${esc(texto)}</t></is></c>`;
}
function sNum(ref: string, n: number, estilo = 0): string {
    return `<c r="${ref}" s="${estilo}"><v>${n}</v></c>`;
}
function sFor(ref: string, formula: string, valor: number, estilo = 0): string {
    return `<c r="${ref}" s="${estilo}"><f>${formula}</f><v>${valor}</v></c>`;
}

interface Arquivo { nome: string; dados: Uint8Array; }

function zip(arquivos: Arquivo[]): Uint8Array {
    const partes: Uint8Array[] = [];
    const central: Uint8Array[] = [];
    let offset = 0;
    const u16 = (v: number) => [v & 255, (v >>> 8) & 255];
    const u32 = (v: number) => [v & 255, (v >>> 8) & 255, (v >>> 16) & 255, (v >>> 24) & 255];
    for (const a of arquivos) {
        const nome = enc.encode(a.nome);
        const crc = crc32(a.dados);
        const local = new Uint8Array([
            0x50, 0x4B, 0x03, 0x04, ...u16(20), ...u16(0x0800), ...u16(0), ...u16(0), ...u16(0x21),
            ...u32(crc), ...u32(a.dados.length), ...u32(a.dados.length), ...u16(nome.length), ...u16(0)
        ]);
        partes.push(local, nome, a.dados);
        central.push(new Uint8Array([
            0x50, 0x4B, 0x01, 0x02, ...u16(20), ...u16(20), ...u16(0x0800), ...u16(0), ...u16(0), ...u16(0x21),
            ...u32(crc), ...u32(a.dados.length), ...u32(a.dados.length), ...u16(nome.length), ...u16(0), ...u16(0),
            ...u16(0), ...u16(0), ...u32(0), ...u32(offset)
        ]), nome);
        offset += local.length + nome.length + a.dados.length;
    }
    const tamCentral = central.reduce((s, p) => s + p.length, 0);
    const fim = new Uint8Array([
        0x50, 0x4B, 0x05, 0x06, ...u16(0), ...u16(0), ...u16(arquivos.length), ...u16(arquivos.length),
        ...u32(tamCentral), ...u32(offset), ...u16(0)
    ]);
    const todos = [...partes, ...central, fim];
    const total = todos.reduce((s, p) => s + p.length, 0);
    const saida = new Uint8Array(total);
    let pos = 0;
    for (const p of todos) { saida.set(p, pos); pos += p.length; }
    return saida;
}

const NS = `xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`;
const XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`;
const EMU = 9525;
const LARG_IMG = 400; // largura das imagens na planilha (px)

export function gerarXlsx(e: EntradaXlsx): Uint8Array {
    const nm = e.meses.length;
    const nu = e.unidades.length;
    const o = e.opcoes;
    const imagens = o.graficos ? e.imagens : [];
    const alturas = imagens.map(i => Math.round(LARG_IMG * i.h / i.w));
    const alturaMax = Math.max(0, ...alturas);
    const colTot = 2 + nm;

    // ---- posições (1-based) ----
    let linha = 4;
    const linhaCards = linha;
    if (o.cards) { linha += 3; }
    const linhaImg = linha;
    if (imagens.length) { linha += Math.ceil(alturaMax / 20) + 2; }
    const linhaTitMatriz = linha;
    const linhaCab = linhaTitMatriz + 1;
    const primeira = linhaCab + 1;
    const linhaTotal = primeira + nu;
    const ultima = primeira + nu - 1;
    const totalGeral = e.matriz.reduce((s, r) => s + r.reduce((a, b) => a + b, 0), 0);

    // ---- aba Painel ----
    let faixa = sTxt("B2", e.titulo.toUpperCase(), 11);
    for (let c = 2; c < colTot; c++) { faixa += `<c r="${letra(c)}2" s="11"/>`; }
    faixa += sTxt(letra(colTot) + "2", "Gerado em " + e.geradoEm, 12);
    let rows = `<row r="2" ht="26" customHeight="1">${faixa}</row>`;

    if (o.cards) {
        const colsKpi = [1, 3, 5];
        let kl = "", kv = "";
        e.kpis.slice(0, 3).forEach((k, i) => {
            kl += sTxt(letra(colsKpi[i]) + linhaCards, k.rotulo, 13) + `<c r="${letra(colsKpi[i] + 1)}${linhaCards}" s="13"/>`;
            const ref = letra(colsKpi[i]) + (linhaCards + 1);
            kv += (i === 0
                ? (o.matriz ? sFor(ref, `${letra(colTot)}${linhaTotal}`, totalGeral, 14) : sNum(ref, totalGeral, 14))
                : sTxt(ref, k.valor, 15)) + `<c r="${letra(colsKpi[i] + 1)}${linhaCards + 1}" s="15"/>`;
        });
        rows += `<row r="${linhaCards}">${kl}</row><row r="${linhaCards + 1}" ht="26" customHeight="1">${kv}</row>`;
    }

    if (o.matriz) {
        rows += `<row r="${linhaTitMatriz}">${sTxt("B" + linhaTitMatriz, "Matriz por mês e unidade", 5)}</row>`;
        let cab = sTxt("B" + linhaCab, "Unidade", 1);
        e.meses.forEach((m, j) => { cab += sTxt(letra(2 + j) + linhaCab, m, 7); });
        cab += sTxt(letra(colTot) + linhaCab, "Total", 7);
        rows += `<row r="${linhaCab}">${cab}</row>`;
        e.unidades.forEach((u, i) => {
            const r = primeira + i;
            let c = sTxt("B" + r, u);
            let somaLinha = 0;
            e.meses.forEach((_, j) => { c += sNum(letra(2 + j) + r, e.matriz[i][j], 8); somaLinha += e.matriz[i][j]; });
            c += sFor(letra(colTot) + r, `SUM(C${r}:${letra(1 + nm)}${r})`, somaLinha, 9);
            rows += `<row r="${r}">${c}</row>`;
        });
        let ct = sTxt("B" + linhaTotal, "Total", 16);
        e.meses.forEach((_, j) => {
            const col = letra(2 + j);
            const s = e.matriz.reduce((a, r) => a + r[j], 0);
            ct += sFor(col + linhaTotal, `SUM(${col}${primeira}:${col}${ultima})`, s, 17);
        });
        ct += sFor(letra(colTot) + linhaTotal, `SUM(${letra(colTot)}${primeira}:${letra(colTot)}${ultima})`, totalGeral, 17);
        rows += `<row r="${linhaTotal}">${ct}</row>`;
    }

    const temDesenho = imagens.length > 0;
    const sheet1 = XML + `<worksheet ${NS}><sheetViews><sheetView showGridLines="0" workbookViewId="0"/></sheetViews>` +
        `<cols><col min="1" max="1" width="3" customWidth="1"/><col min="2" max="${Math.max(14, colTot + 1)}" width="14" customWidth="1"/></cols>` +
        `<sheetData>${rows}</sheetData>${temDesenho ? '<drawing r:id="rId1"/>' : ""}</worksheet>`;
    const rels1 = XML + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>`;

    const colsImg = [1, 5];
    let pics = "";
    imagens.forEach((img, i) => {
        const cx = LARG_IMG * EMU, cy = alturas[i] * EMU;
        pics += `<xdr:oneCellAnchor><xdr:from><xdr:col>${colsImg[i] ?? 1}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${linhaImg - 1}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
            `<xdr:ext cx="${cx}" cy="${cy}"/>` +
            `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${i + 2}" name="Grafico ${i + 1}"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>` +
            `<xdr:blipFill><a:blip r:embed="rId${i + 1}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>` +
            `<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>`;
    });
    const drawing = XML + `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${pics}</xdr:wsDr>`;
    const drawingRels = XML + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        imagens.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image${i + 1}.png"/>`).join("") +
        `</Relationships>`;

    // ---- aba Dados (opcional) ----
    let sheet2 = "";
    if (o.dados) {
        let linhasDados = `<row r="1">${sTxt("A1", "Mês", 1)}${sTxt("B1", "Unidade", 1)}${sTxt("C1", "Devoluções", 1)}</row>`;
        e.linhas.forEach((l, i) => {
            const r = i + 2;
            linhasDados += `<row r="${r}">${sTxt("A" + r, l.mes)}${sTxt("B" + r, l.unidade)}${sNum("C" + r, l.valor)}</row>`;
        });
        sheet2 = XML + `<worksheet ${NS}><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>` +
            `<cols><col min="1" max="3" width="16" customWidth="1"/></cols><sheetData>${linhasDados}</sheetData></worksheet>`;
    }

    const styles = XML + `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
        `<fonts count="8"><font><sz val="11"/><name val="Calibri"/></font>` +
        `<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>` +
        `<font><b/><sz val="18"/><color rgb="FF0B2A4A"/><name val="Calibri"/></font>` +
        `<font><sz val="11"/><color rgb="FF6B7280"/><name val="Calibri"/></font>` +
        `<font><b/><sz val="11"/><name val="Calibri"/></font>` +
        `<font><b/><sz val="16"/><color rgb="FF0B2A4A"/><name val="Calibri"/></font>` +
        `<font><b/><sz val="14"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>` +
        `<font><sz val="10"/><color rgb="FFCBD5E1"/><name val="Calibri"/></font></fonts>` +
        `<fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>` +
        `<fill><patternFill patternType="solid"><fgColor rgb="FF0B2A4A"/><bgColor indexed="64"/></patternFill></fill>` +
        `<fill><patternFill patternType="solid"><fgColor rgb="FFF3F4F6"/><bgColor indexed="64"/></patternFill></fill></fills>` +
        `<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>` +
        `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
        `<cellXfs count="18"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>` +
        `<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>` +
        `<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>` +
        `<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>` +
        `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>` +
        `<xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0" applyFont="1"/>` +
        `<xf numFmtId="0" fontId="5" fillId="0" borderId="0" xfId="0" applyFont="1"/>` +
        `<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="right"/></xf>` +
        `<xf numFmtId="3" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>` +
        `<xf numFmtId="3" fontId="4" fillId="0" borderId="0" xfId="0" applyFont="1" applyNumberFormat="1"/>` +
        `<xf numFmtId="3" fontId="5" fillId="0" borderId="0" xfId="0" applyFont="1" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="left"/></xf>` +
        `<xf numFmtId="0" fontId="6" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>` +
        `<xf numFmtId="0" fontId="7" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="right"/></xf>` +
        `<xf numFmtId="0" fontId="3" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1"/>` +
        `<xf numFmtId="3" fontId="5" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="left"/></xf>` +
        `<xf numFmtId="0" fontId="5" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1"/>` +
        `<xf numFmtId="0" fontId="4" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1"/>` +
        `<xf numFmtId="3" fontId="4" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1" applyNumberFormat="1"/></cellXfs>` +
        `<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;

    const idEstilo = o.dados ? 3 : 2;
    const workbook = XML + `<workbook ${NS}><sheets><sheet name="Painel" sheetId="1" r:id="rId1"/>${o.dados ? '<sheet name="Dados" sheetId="2" r:id="rId2"/>' : ""}</sheets></workbook>`;
    const wbRels = XML + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>` +
        (o.dados ? `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>` : "") +
        `<Relationship Id="rId${idEstilo}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`;
    const rootRels = XML + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`;
    const tipos = XML + `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>` +
        `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
        `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>` +
        (o.dados ? `<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>` : "") +
        `<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>` +
        (temDesenho ? `<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>` : "") +
        `</Types>`;

    const t = (nome: string, txt: string): Arquivo => ({ nome, dados: enc.encode(txt) });
    const arquivos: Arquivo[] = [
        t("[Content_Types].xml", tipos),
        t("_rels/.rels", rootRels),
        t("xl/workbook.xml", workbook),
        t("xl/_rels/workbook.xml.rels", wbRels),
        t("xl/styles.xml", styles),
        t("xl/worksheets/sheet1.xml", sheet1)
    ];
    if (temDesenho) {
        arquivos.push(
            t("xl/worksheets/_rels/sheet1.xml.rels", rels1),
            t("xl/drawings/drawing1.xml", drawing),
            t("xl/drawings/_rels/drawing1.xml.rels", drawingRels),
            ...imagens.map((img, i): Arquivo => ({ nome: `xl/media/image${i + 1}.png`, dados: img.png }))
        );
    }
    if (o.dados) { arquivos.push(t("xl/worksheets/sheet2.xml", sheet2)); }
    return zip(arquivos);
}
export function paraBase64(b: Uint8Array): string {
    let s = "";
    for (let i = 0; i < b.length; i += 0x8000) { s += String.fromCharCode.apply(null, Array.from(b.subarray(i, i + 0x8000))); }
    return btoa(s);
}
