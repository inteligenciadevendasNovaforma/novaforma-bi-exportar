// Gerador de .xlsx do Painel Metas (sem bibliotecas): zip sem compressão + XML.
// A aba "Painel Metas" repete a aba PAINEL METAS do exemplo: mesmas colunas, alturas, mesclagens, cores e formatos
// (molde em modeloMetas.ts, extraído da planilha). Bloco total no topo; depois um bloco por UF (1º no molde ES, demais no molde MG).
import { MODELO, COLS_XML, STYLES_XML, Segmento } from "./modeloMetas";

export type N = number | null;
// v: objetivo, vendido mês, vendido ontem, tendência, faturado, obj x fat, carteira, faturado + carteira, cliente mês, clientes ontem, clientes faturados, clientes carteira
export interface LinhaMetas { nome: string; v: N[]; d: N[]; }
// k: objetivo clientes, clientes (V), clientes (F), %, clientes carteira, tendência (F+C), prazo (V), nota mix (V); cli: clientes por dia (índice 0 = dia 1)
export interface BlocoMetas { uf: string; k: N[]; total: LinhaMetas; fams: LinhaMetas[]; cli: N[]; }
export interface OpcoesMetas { dados: boolean; }
export interface EntradaMetas {
    titulo: string; atualizacao: string; geradoEm: string; periodo?: string; filtros?: string; blocos: BlocoMetas[]; opcoes: OpcoesMetas;
}
export const COLS_V = ["obj", "vend", "ontem", "tend", "fat", "objFat", "cart", "fatCart", "cliMes", "cliOntem", "cliFat", "cliCart"];
export const COLS_K = ["obj", "cliV", "cliF", "pct", "cliCart", "tend", "prazo", "mix"];

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
function letra(i: number): string { // 1 = A
    let s = "";
    let n = i;
    while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
    return s;
}
function cTxt(ref: string, texto: string, estilo: number): string {
    return `<c r="${ref}" s="${estilo}" t="inlineStr"><is><t xml:space="preserve">${esc(texto)}</t></is></c>`;
}
function cNum(ref: string, n: number, estilo: number): string { return `<c r="${ref}" s="${estilo}"><v>${n}</v></c>`; }
function cVazio(ref: string, estilo: number): string { return `<c r="${ref}" s="${estilo}"/>`; }

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
    const saida = new Uint8Array(todos.reduce((s, p) => s + p.length, 0));
    let pos = 0;
    for (const p of todos) { saida.set(p, pos); pos += p.length; }
    return saida;
}

const NS = `xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`;
const XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`;

function num(v: N | undefined): number | null { return v === null || v === undefined || isNaN(Number(v)) ? null : Number(v); }

// valor de uma chave do molde ("k:obj", "f0:vend", "t:d12", "c:d3", "f1:nome", "titulo", "atualizacao", "uf:titulo")
function valor(chave: string, b: BlocoMetas, e: EntradaMetas): string | number | null {
    if (chave === "titulo") { return e.titulo; }
    if (chave === "atualizacao") { return e.atualizacao; }
    if (chave === "uf:titulo") { return "PAINEL METAS | " + b.uf; }
    const [lin, campo] = chave.split(":");
    if (lin === "k") { const i = COLS_K.indexOf(campo); return i < 0 ? null : num(b.k[i]); }
    if (lin === "c") { const d = Number(campo.slice(1)); return num(b.cli[d - 1]) ?? 0; }
    const linha = lin === "t" ? b.total : b.fams[Number(lin.slice(1))];
    if (!linha) { return null; }
    if (campo === "nome") { return linha.nome; }
    if (campo.charAt(0) === "d" && /^d\d+$/.test(campo)) { return num(linha.d[Number(campo.slice(1)) - 1]); }
    const i = COLS_V.indexOf(campo);
    return i < 0 ? null : num(linha.v[i]);
}

function escreverSegmento(seg: Segmento, r0: number, b: BlocoMetas, e: EntradaMetas, merges: string[]): string {
    let xml = "";
    seg.linhas.forEach((l, i) => {
        const r = r0 + i;
        let attrs = `r="${r}"`;
        if (l.s !== null) { attrs += ` s="${l.s}" customFormat="1"`; }
        if (l.h !== null) { attrs += ` ht="${l.h}" customHeight="1"`; }
        let cs = "";
        l.c.forEach(c => {
            const ref = letra(Number(c[0])) + r;
            const estilo = Number(c[1]);
            if (c.length < 3) { cs += cVazio(ref, estilo); return; }
            const t = String(c[2]);
            if (t.charAt(0) !== "=") { cs += cTxt(ref, t, estilo); return; }
            const v = valor(t.slice(1), b, e);
            if (v === null || v === "") { cs += cVazio(ref, estilo); } else if (typeof v === "number") { cs += cNum(ref, v, estilo); } else { cs += cTxt(ref, v, estilo); }
        });
        xml += `<row ${attrs}>${cs}</row>`;
    });
    seg.merges.forEach(m => merges.push(`${m[0]}${r0 + Number(m[1])}:${m[2]}${r0 + Number(m[3])}`));
    return xml;
}

function abaDados(e: EntradaMetas): string {
    const cabs = ["UF", "Família", "Objetivo", "Vendido mês", "Vendido Ontem", "Tendência", "Faturado", "Obj x Fat", "Carteira", "Faturado + Carteira",
        "Cliente Mês", "Clientes Ontem", "Clientes Faturados", "Clientes Carteira"];
    for (let d = 1; d <= 31; d++) { cabs.push("Dia " + (d < 10 ? "0" : "") + d); }
    let rows = `<row r="1">${cabs.map((h, i) => cTxt(letra(i + 1) + "1", h, 0)).join("")}</row>`;
    let r = 2;
    e.blocos.forEach(b => {
        [...b.fams, b.total].forEach(l => {
            let c = cTxt("A" + r, b.uf || "Total", 0) + cTxt("B" + r, l.nome, 0);
            l.v.forEach((v, i) => { const n = num(v); if (n !== null) { c += cNum(letra(i + 3) + r, n, 0); } });
            for (let d = 0; d < 31; d++) { const n = num(l.d[d]); if (n !== null) { c += cNum(letra(15 + d) + r, n, 0); } }
            rows += `<row r="${r}">${c}</row>`; r++;
        });
        let c = cTxt("A" + r, b.uf || "Total", 0) + cTxt("B" + r, "CLIENTES (por dia)", 0);
        for (let d = 0; d < 31; d++) { const n = num(b.cli[d]); if (n !== null) { c += cNum(letra(15 + d) + r, n, 0); } }
        rows += `<row r="${r}">${c}</row>`; r++;
    });
    r++;
    ["Gerado em " + e.geradoEm, e.atualizacao, e.periodo || "", e.filtros || ""].filter(x => !!x).forEach(n => {
        rows += `<row r="${r}">${cTxt("A" + r, n, 0)}</row>`; r++;
    });
    return XML + `<worksheet ${NS}><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>` +
        `<cols><col min="1" max="1" width="8" customWidth="1"/><col min="2" max="2" width="24" customWidth="1"/><col min="3" max="14" width="13" customWidth="1"/><col min="15" max="45" width="9" customWidth="1"/></cols>` +
        `<sheetData>${rows}</sheetData></worksheet>`;
}

export function gerarXlsxMetas(e: EntradaMetas): Uint8Array {
    const merges: string[] = [];
    let rows = "";
    let r = 1;
    const blocos = e.blocos.length ? e.blocos : [{ uf: "", k: [], total: { nome: "", v: [], d: [] }, fams: [], cli: [] }];
    blocos.forEach((b, i) => {
        const seg = i === 0 ? MODELO.topo : i === 1 ? MODELO.uf1 : MODELO.ufN;
        rows += escreverSegmento(seg, r, b, e, merges);
        r += seg.linhas.length;
    });
    const abas: { nome: string; xml: string }[] = [{
        nome: "Painel Metas",
        xml: XML + `<worksheet ${NS}><sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><dimension ref="A1:AK${r - 1}"/><sheetViews><sheetView workbookViewId="0"/></sheetViews>` +
            `<sheetFormatPr defaultRowHeight="12.8"/><cols>${COLS_XML}</cols><sheetData>${rows}</sheetData>` +
            `<mergeCells count="${merges.length}">${merges.map(m => `<mergeCell ref="${m}"/>`).join("")}</mergeCells>` +
            `<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/><pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="0"/></worksheet>`
    }];
    if (e.opcoes && e.opcoes.dados) { abas.push({ nome: "Dados", xml: abaDados(e) }); }

    const workbook = XML + `<workbook ${NS}><sheets>${abas.map((a, i) => `<sheet name="${esc(a.nome)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join("")}</sheets></workbook>`;
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
        t("xl/_rels/workbook.xml.rels", wbRels), t("xl/styles.xml", XML + STYLES_XML),
        ...abas.map((a, i) => t(`xl/worksheets/sheet${i + 1}.xml`, a.xml))
    ]);
}
export function paraBase64(b: Uint8Array): string {
    let s = "";
    for (let i = 0; i < b.length; i += 0x8000) { s += String.fromCharCode.apply(null, Array.from(b.subarray(i, i + 0x8000))); }
    return btoa(s);
}
