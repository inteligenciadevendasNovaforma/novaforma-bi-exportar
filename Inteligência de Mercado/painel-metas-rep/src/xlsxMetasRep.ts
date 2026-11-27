// Gerador de .xlsx do Painel Metas | Representante (sem bibliotecas): zip sem compressão + XML.
// A aba "(Data do APP) PAINEL METAS | Re" repete a aba do exemplo: mesmas colunas, alturas, mesclagens, cores e formatos
// (molde em modeloMetasRep.ts, extraído da planilha). Um bloco de 13 linhas por representante, na ordem da tela.
import { MODELO, COLS_XML, STYLES_XML, ESTILO_PCT, Segmento } from "./modeloMetasRep";

export type N = number | null;
// v: objetivo, vendido mês, vendido ontem, tendência, faturado, obj x fat, carteira, faturado + carteira, cliente mês, clientes ontem, clientes faturados, clientes carteira
export interface LinhaRep { nome: string; v: N[]; }
// k: objetivo clientes, clientes (V), clientes (F), % clientes (V)
export interface BlocoRep { cod: string; nome: string; uf: string; k: N[]; total: LinhaRep; fams: LinhaRep[]; }
export interface OpcoesRep { tabela: boolean; dados: boolean; }
export interface EntradaRep {
    titulo: string; aba: string; atualizacao: string; geradoEm: string; periodo?: string; filtros?: string; blocos: BlocoRep[]; opcoes: OpcoesRep;
}
export const ABA = "(Data do APP) PAINEL METAS | Re";
export const COLS_V = ["obj", "vend", "ontem", "tend", "fat", "objFat", "cart", "fatCart", "cliMes", "cliOntem", "cliFat", "cliCart"];
export const COLS_K = ["obj", "cliV", "cliF", "pct"];
export const ROTULOS_V = ["Objetivo", "Vendido mês", "Vendido Ontem", "Tendência", "Faturado", "Obj x Fat", "Carteira", "Faturado + Carteira",
    "Cliente Mês", "Clientes Ontem", "Clientes Faturados", "Clientes Carteira"];
// Cliente Mês, Clientes Ontem e Clientes Faturados aparecem como 0 quando vazios (como na aba); Clientes Carteira fica vazio
export const ZERO_SE_VAZIO = [8, 9, 10];
export const IDX_PCT = 5;
// título do bloco, igual à aba: "PAINEL DE VENDAS - MÊS | NOME - UF"
export function tituloBloco(b: BlocoRep): string { return "PAINEL DE VENDAS - MÊS | " + b.nome + (b.uf ? " - " + b.uf : ""); }

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

function valor(chave: string, b: BlocoRep): number | string | null {
    if (chave === "titulo") { return tituloBloco(b); }
    const [lin, campo] = chave.split(":");
    if (lin === "k") { const i = COLS_K.indexOf(campo); return i < 0 ? null : num(b.k[i]); }
    const linha = lin === "t" ? b.total : b.fams[Number(lin.slice(1))];
    if (!linha) { return null; }
    if (campo === "nome") { return linha.nome; }
    const i = COLS_V.indexOf(campo);
    if (i < 0) { return null; }
    const v = num(linha.v[i]);
    return v === null && ZERO_SE_VAZIO.indexOf(i) >= 0 ? 0 : v;
}
function estiloPct(v: number | string | null, chave: string, estilo: number): number {
    const pct = chave === "k:pct" || /:objFat$/.test(chave);
    if (!pct) { return estilo; }
    if (v === null || v === "") { return ESTILO_PCT.vazio; }
    return chave !== "k:pct" && typeof v === "number" && v >= 1 ? ESTILO_PCT.meta : ESTILO_PCT.valor;
}

function escreverSegmento(seg: Segmento, r0: number, b: BlocoRep | null, merges: string[]): string {
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
            if (t.charAt(0) !== "=" || !b) { cs += t.charAt(0) === "=" ? cVazio(ref, estilo) : cTxt(ref, t, estilo); return; }
            const chave = t.slice(1);
            const v = valor(chave, b);
            const st = estiloPct(v, chave, estilo);
            if (v === null || v === "") { cs += cVazio(ref, st); } else if (typeof v === "number") { cs += cNum(ref, v, st); } else { cs += cTxt(ref, v, st); }
        });
        xml += `<row ${attrs}>${cs}</row>`;
    });
    seg.merges.forEach(m => merges.push(`${m[0]}${r0 + Number(m[1])}:${m[2]}${r0 + Number(m[3])}`));
    return xml;
}

function abaDados(e: EntradaRep): string {
    const cabs = ["Código", "Representante", "UF", "Objetivo Clientes", "Clientes(V)", "Clientes (F)", "%", "Família", ...ROTULOS_V];
    let rows = `<row r="1">${cabs.map((h, i) => cTxt(letra(i + 1) + "1", h, 0)).join("")}</row>`;
    let r = 2;
    e.blocos.forEach(b => {
        [...b.fams, b.total].forEach(l => {
            let c = cTxt("A" + r, b.cod, 0) + cTxt("B" + r, b.nome, 0) + cTxt("C" + r, b.uf, 0);
            b.k.forEach((v, i) => { const n = num(v); if (n !== null) { c += cNum(letra(4 + i) + r, n, 0); } });
            c += cTxt("H" + r, l.nome, 0);
            l.v.forEach((v, i) => { const n = num(v); if (n !== null) { c += cNum(letra(9 + i) + r, n, 0); } });
            rows += `<row r="${r}">${c}</row>`; r++;
        });
    });
    r++;
    ["Gerado em " + e.geradoEm, e.atualizacao, e.periodo || "", e.filtros || ""].filter(x => !!x).forEach(n => {
        rows += `<row r="${r}">${cTxt("A" + r, n, 0)}</row>`; r++;
    });
    return XML + `<worksheet ${NS}><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>` +
        `<cols><col min="1" max="1" width="10" customWidth="1"/><col min="2" max="2" width="38" customWidth="1"/><col min="3" max="7" width="11" customWidth="1"/><col min="8" max="8" width="24" customWidth="1"/><col min="9" max="20" width="13" customWidth="1"/></cols>` +
        `<sheetData>${rows}</sheetData></worksheet>`;
}

export function gerarXlsxMetasRep(e: EntradaRep): Uint8Array {
    const abas: { nome: string; xml: string }[] = [];
    if (!e.opcoes || e.opcoes.tabela !== false) {
        const merges: string[] = [];
        let rows = escreverSegmento(MODELO.topo, 1, null, merges);
        let r = 1 + MODELO.topo.linhas.length;
        e.blocos.forEach(b => { rows += escreverSegmento(MODELO.bloco, r, b, merges); r += MODELO.bloco.linhas.length; });
        abas.push({
            nome: (e.aba || ABA).slice(0, 31),
            xml: XML + `<worksheet ${NS}><sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><dimension ref="A1:N${r - 1}"/><sheetViews><sheetView workbookViewId="0" showGridLines="0"/></sheetViews>` +
                `<sheetFormatPr defaultRowHeight="12.8"/><cols>${COLS_XML}</cols><sheetData>${rows}</sheetData>` +
                (merges.length ? `<mergeCells count="${merges.length}">${merges.map(m => `<mergeCell ref="${m}"/>`).join("")}</mergeCells>` : "") +
                `<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/><pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="0"/></worksheet>`
        });
    }
    if (!abas.length || (e.opcoes && e.opcoes.dados)) { abas.push({ nome: "Dados", xml: abaDados(e) }); }

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

// ---------- carga compacta (link curto ou área de transferência) ----------
// R$ em centavos inteiros, percentuais x 1.000.000 inteiros, contagens inteiras; famílias numa tabela (F) e índice por linha.
//   B: [nome, uf, cod, [objCli, cliV, cliF, pct], total[12], [[fi, v[12]], ...]]  (vazio = null; nulls do fim cortados)
export interface CargaRep {
    tipo: "painelMetasRep"; versao: number; titulo: string; aba: string; atualizacao: string; geradoEm: string; periodo?: string; filtros?: string;
    opcoes: OpcoesRep; F: string[]; B: unknown[][];
}
const PCT_K = 3;
const ehPct = (i: number): boolean => i === IDX_PCT;
const ehCont = (i: number): boolean => i >= 8;
function corta(a: N[]): N[] { const b = a.slice(); while (b.length && (b[b.length - 1] === null || b[b.length - 1] === undefined)) { b.pop(); } return b; }
function comp(v: N, i: number): N {
    if (v === null || v === undefined || isNaN(v)) { return null; }
    return ehPct(i) ? Math.round(v * 1e6) : ehCont(i) ? Math.round(v) : Math.round(v * 100);
}
function desc(v: N | undefined, i: number): N {
    if (v === null || v === undefined) { return null; }
    return ehPct(i) ? v / 1e6 : ehCont(i) ? v : v / 100;
}
export function compactar(e: EntradaRep): CargaRep {
    const F: string[] = [];
    const fi = (n: string): number => { let i = F.indexOf(n); if (i < 0) { i = F.length; F.push(n); } return i; };
    const B = e.blocos.map(b => [
        b.nome, b.uf, b.cod,
        corta(b.k.map((v, i) => (v === null || v === undefined ? null : i === PCT_K ? Math.round(v * 1e6) : Math.round(v)))),
        corta(b.total.v.map(comp)),
        b.fams.map(l => [fi(l.nome), corta(l.v.map(comp))])
    ]);
    return { tipo: "painelMetasRep", versao: 1, titulo: e.titulo, aba: e.aba, atualizacao: e.atualizacao, geradoEm: e.geradoEm, periodo: e.periodo, filtros: e.filtros, opcoes: e.opcoes, F, B };
}
export function expandir(c: CargaRep): EntradaRep {
    const v12 = (a: unknown): N[] => { const x = (Array.isArray(a) ? a : []) as N[]; const o: N[] = []; for (let i = 0; i < 12; i++) { o.push(desc(x[i], i)); } return o; };
    const blocos: BlocoRep[] = (c.B || []).map(b => {
        const k = (Array.isArray(b[3]) ? b[3] : []) as N[];
        const kk: N[] = []; for (let i = 0; i < 4; i++) { const v = k[i]; kk.push(v === null || v === undefined ? null : i === PCT_K ? v / 1e6 : v); }
        return {
            nome: String(b[0] || ""), uf: String(b[1] || ""), cod: String(b[2] || ""), k: kk,
            total: { nome: "TOTAL", v: v12(b[4]) },
            fams: ((b[5] || []) as unknown[][]).map(l => ({ nome: (c.F || [])[Number(l[0])] || "", v: v12(l[1]) }))
        };
    });
    return {
        titulo: c.titulo || "PAINEL METAS | REPRESENTANTE", aba: c.aba || ABA, atualizacao: c.atualizacao || "", geradoEm: c.geradoEm || "", periodo: c.periodo || "", filtros: c.filtros || "",
        blocos, opcoes: { tabela: !c.opcoes || c.opcoes.tabela !== false, dados: !!(c.opcoes && c.opcoes.dados) }
    };
}
