// Gerador de .xlsx do Pedido Diário v2 (sem bibliotecas): zip sem compressão + XML das planilhas.
// Só células (sem imagens): aba com a tabela (título, cabeçalho, matriz por dia, notas) e, opcional, aba Dados.
export interface LinhaPedido { nome: string; uf: string; sup: string; mes: number | null; vend: number; q: (number | null)[]; }
// cards/medidores: só existiam na v1 (aceitos e ignorados, para links antigos continuarem gerando a tabela)
export interface OpcoesPedido { tabela: boolean; dados: boolean; cards?: boolean; medidores?: boolean; }
export interface EntradaPedido {
    titulo: string; geradoEm: string; periodo?: string; filtros?: string; dias: number[]; linhas: LinhaPedido[]; opcoes: OpcoesPedido;
}
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
function sTxt(ref: string, texto: string, estilo = 0): string {
    return `<c r="${ref}" t="inlineStr" s="${estilo}"><is><t xml:space="preserve">${esc(texto)}</t></is></c>`;
}
function sNum(ref: string, n: number, estilo = 0): string {
    return `<c r="${ref}" s="${estilo}"><v>${n}</v></c>`;
}
function vazio(ref: string, estilo: number): string { return `<c r="${ref}" s="${estilo}"/>`; }

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
function letra(i: number): string {
    let s = "";
    let n = i + 1;
    while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
    return s;
}

// estilos (índices de cellXfs)
const S = { titulo: 1, cab: 2, nome: 3, uf: 4, sup: 5, dia: 6, vazio: 7, nota: 8, dCab: 9, dTxt: 10, dNum: 11, dVal: 12, semDados: 13 };

export function gerarXlsxPedido(e: EntradaPedido): Uint8Array {
    const o = e.opcoes;
    const dias = e.dias || [];
    const nd = dias.length;
    const abas: { nome: string; xml: string }[] = [];

    if (o.tabela || !o.dados) {
        const ult = 3 + Math.max(nd, 1); // última coluna da tabela (0 = A)
        const L = letra(ult);
        const cheias = (r: number, estilo: number, de = 2): string => { let s = ""; for (let c = de; c <= ult; c++) { s += vazio(letra(c) + r, estilo); } return s; };
        let rows = "";
        const merges: string[] = [];
        let r = 2;
        rows += `<row r="${r}" ht="30" customHeight="1">${sTxt("B" + r, (e.titulo || "").toUpperCase(), S.titulo)}${cheias(r, S.titulo)}</row>`;
        merges.push(`B${r}:${L}${r}`); r++;
        let cab = sTxt("B" + r, "REPRESENTANTES", S.cab) + sTxt("C" + r, "UF", S.cab) + sTxt("D" + r, "SUP", S.cab);
        if (!nd) { cab += vazio("E" + r, S.cab); }
        dias.forEach((d, j) => { cab += sNum(letra(4 + j) + r, d, S.cab); });
        rows += `<row r="${r}" ht="21" customHeight="1">${cab}</row>`;
        r++;
        if (!e.linhas.length) {
            rows += `<row r="${r}" ht="20" customHeight="1">${sTxt("B" + r, "Sem representantes para os filtros escolhidos.", S.semDados)}${cheias(r, S.semDados, 2)}</row>`;
            merges.push(`B${r}:${L}${r}`); r++;
        }
        e.linhas.forEach(l => {
            let c = sTxt("B" + r, l.nome, S.nome) + sTxt("C" + r, l.uf, S.uf) + sTxt("D" + r, l.sup, S.sup);
            if (!nd) { c += vazio("E" + r, S.vazio); }
            dias.forEach((_, j) => {
                const v = l.q ? l.q[j] : null;
                c += v === null || v === undefined ? vazio(letra(4 + j) + r, S.vazio) : sNum(letra(4 + j) + r, v, S.dia);
            });
            rows += `<row r="${r}" ht="18" customHeight="1">${c}</row>`;
            r++;
        });
        r++;
        const notas = ["Gerado em " + e.geradoEm, e.periodo || "", e.filtros || ""].filter(x => !!x);
        notas.forEach(n => {
            rows += `<row r="${r}">${sTxt("B" + r, n, S.nota)}${cheias(r, S.nota)}</row>`;
            merges.push(`B${r}:${L}${r}`); r++;
        });
        let cols = `<col min="1" max="1" width="1.5" customWidth="1"/><col min="2" max="2" width="40" customWidth="1"/><col min="3" max="3" width="6" customWidth="1"/><col min="4" max="4" width="13" customWidth="1"/>`;
        cols += `<col min="5" max="${4 + Math.max(nd, 1)}" width="4.6" customWidth="1"/>`;
        abas.push({
            nome: "Pedido Diário",
            xml: XML + `<worksheet ${NS}><sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><sheetViews><sheetView showGridLines="0" workbookViewId="0"><pane ySplit="3" topLeftCell="A4" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>` +
                `<cols>${cols}</cols><sheetData>${rows}</sheetData>` +
                `<mergeCells count="${merges.length}">${merges.map(m => `<mergeCell ref="${m}"/>`).join("")}</mergeCells>` +
                `<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/><pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="0"/></worksheet>`
        });
    }

    const ultD = letra(4 + nd);
    const nLin = Math.max(e.linhas.length + 1, 1);
    if (o.dados) {
        const cabs = ["Representante", "UF", "SUP", "CLI MÊS", ...dias.map(d => "Dia " + d), "Vendido no mês (R$)"];
        let d = `<row r="1">${cabs.map((h, i) => sTxt(letra(i) + "1", h, S.dCab)).join("")}</row>`;
        e.linhas.forEach((l, i) => {
            const r = i + 2;
            let c = sTxt("A" + r, l.nome, S.dTxt) + sTxt("B" + r, l.uf, S.dTxt) + sTxt("C" + r, l.sup, S.dTxt);
            if (l.mes !== null && l.mes !== undefined) { c += sNum("D" + r, l.mes, S.dNum); }
            dias.forEach((_, j) => { const v = l.q ? l.q[j] : null; if (v !== null && v !== undefined) { c += sNum(letra(4 + j) + r, v, S.dNum); } });
            c += sNum(letra(4 + nd) + r, Number(l.vend) || 0, S.dVal);
            d += `<row r="${r}">${c}</row>`;
        });
        abas.push({
            nome: "Dados",
            xml: XML + `<worksheet ${NS}><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>` +
                `<cols><col min="1" max="1" width="40" customWidth="1"/><col min="2" max="2" width="7" customWidth="1"/><col min="3" max="3" width="14" customWidth="1"/><col min="4" max="${4 + nd}" width="9" customWidth="1"/><col min="${5 + nd}" max="${5 + nd}" width="20" customWidth="1"/></cols>` +
                `<sheetData>${d}</sheetData><autoFilter ref="A1:${ultD}${nLin}"/></worksheet>`
        });
    }

    const styles = XML + `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
        `<numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0.00"/></numFmts>` +
        `<fonts count="8"><font><sz val="11"/><name val="Calibri"/></font>` +
        `<font><b/><sz val="16"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>` +
        `<font><b/><sz val="11"/><color rgb="FF003463"/><name val="Calibri"/></font>` +
        `<font><b/><sz val="10"/><color rgb="FF333333"/><name val="Arial"/></font>` +
        `<font><sz val="10"/><color rgb="FF333333"/><name val="Arial"/></font>` +
        `<font><sz val="9"/><color rgb="FF6B7280"/><name val="Calibri"/></font>` +
        `<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>` +
        `<font><i/><sz val="10"/><color rgb="FF6B7280"/><name val="Arial"/></font></fonts>` +
        `<fills count="5"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>` +
        `<fill><patternFill patternType="solid"><fgColor rgb="FF003463"/><bgColor indexed="64"/></patternFill></fill>` +
        `<fill><patternFill patternType="solid"><fgColor rgb="FFD6DCE4"/><bgColor indexed="64"/></patternFill></fill>` +
        `<fill><patternFill patternType="solid"><fgColor rgb="FFFFD5DA"/><bgColor indexed="64"/></patternFill></fill></fills>` +
        `<borders count="3"><border><left/><right/><top/><bottom/><diagonal/></border>` +
        `<border><left/><right/><top/><bottom style="thin"><color rgb="FFD9D9D9"/></bottom><diagonal/></border>` +
        `<border><left style="thin"><color rgb="FFFFFFFF"/></left><right style="thin"><color rgb="FFFFFFFF"/></right><top/><bottom style="thin"><color rgb="FFD9D9D9"/></bottom><diagonal/></border></borders>` +
        `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
        `<cellXfs count="14"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>` +
        `<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>` +
        `<xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>` +
        `<xf numFmtId="0" fontId="3" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>` +
        `<xf numFmtId="0" fontId="3" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>` +
        `<xf numFmtId="0" fontId="4" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>` +
        `<xf numFmtId="1" fontId="4" fillId="0" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>` +
        `<xf numFmtId="0" fontId="4" fillId="4" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>` +
        `<xf numFmtId="0" fontId="5" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="left" wrapText="1"/></xf>` +
        `<xf numFmtId="0" fontId="6" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>` +
        `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>` +
        `<xf numFmtId="1" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>` +
        `<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>` +
        `<xf numFmtId="0" fontId="7" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf></cellXfs>` +
        `<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;

    const iDados = abas.findIndex(a => a.nome === "Dados");
    const workbook = XML + `<workbook ${NS}><sheets>${abas.map((a, i) => `<sheet name="${esc(a.nome)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join("")}</sheets>` +
        (iDados >= 0 ? `<definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="${iDados}" hidden="1">Dados!$A$1:$${ultD}$${nLin}</definedName></definedNames>` : "") +
        `</workbook>`;
    const wbRels = XML + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        abas.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join("") +
        `<Relationship Id="rId${abas.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`;
    const rootRels = XML + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`;
    const tipos = XML + `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>` +
        `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
        abas.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("") +
        `<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`;

    const t = (nome: string, txt: string): Arquivo => ({ nome, dados: enc.encode(txt) });
    return zip([
        t("[Content_Types].xml", tipos), t("_rels/.rels", rootRels), t("xl/workbook.xml", workbook),
        t("xl/_rels/workbook.xml.rels", wbRels), t("xl/styles.xml", styles),
        ...abas.map((a, i) => t(`xl/worksheets/sheet${i + 1}.xml`, a.xml))
    ]);
}
export function paraBase64(b: Uint8Array): string {
    let s = "";
    for (let i = 0; i < b.length; i += 0x8000) { s += String.fromCharCode.apply(null, Array.from(b.subarray(i, i + 0x8000))); }
    return btoa(s);
}
