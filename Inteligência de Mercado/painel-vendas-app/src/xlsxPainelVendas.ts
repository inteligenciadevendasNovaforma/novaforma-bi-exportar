// Gerador de .xlsx do Painel de Vendas (Data Aprovação / Data do APP), sem bibliotecas: zip sem compressão + XML.
// A aba repete a aba (Data Aprovação|Data APP) EVOLUÇÃO VENDA do exemplo: título B2:M2 "PAINEL DE VENDAS - DIA",
// faixas na linha 4 (VENDIDO E4:G4, FATURADO H4:J4, CARTEIRA K4, CLIENTES (VENDIDO) L4:N4, MIX (VENDIDO %) O4:S4),
// cabeçalho na linha 5 (REPRESENTANTES | UF | SUP | 3 meses x 3 | Em Aberto | FB FC PT AP CT) e uma linha por representante.
// Estilos, larguras e alturas vêm de modeloPainelVendas.ts (extraídos da planilha). Mix vazio = célula rosa, como na aba.
// Também ficam aqui o formato compacto (link / área de transferência) e o texto TSV do botão Copiar, usados pelo visual e pela página.
// O mesmo arquivo atende os dois visuais (Aprovação e APP) e o tipo "painelVendas" da página de download.
import { COLS_XML, STYLES_XML, ESTILO, ALTURA } from "./modeloPainelVendas";

export type N = number | null;
// v: 0-2 vendido (3 meses), 3-5 faturado, 6 carteira, 7-9 clientes, 10-14 mix FB FC PT AP CT (fração)
export interface LinhaPV { cod: string; nome: string; uf: string; sup: string; v: N[]; }
export interface OpcoesPV { tabela: boolean; dados: boolean; }
export interface EntradaPV {
    titulo: string; aba: string; regra: string; meses: string[]; geradoEm: string; periodo?: string; filtros?: string;
    linhas: LinhaPV[]; opcoes: OpcoesPV;
}

export const FAMILIAS = ["FB", "FC", "PT", "AP", "CT"];
export const GRUPOS = ["VENDIDO", "FATURADO", "CARTEIRA", "CLIENTES (VENDIDO)", "MIX (VENDIDO %)"];
export const NCOLS = 15;
// rótulos das 15 colunas de valores (linha 5 da aba); m = títulos dos 3 meses
export function rotulos(m: string[]): string[] {
    const t = [0, 1, 2].map(i => m[i] || ["Mês −2", "Mês −1", "Mês atual"][i]);
    return [...t, ...t, "Em Aberto", ...t, ...FAMILIAS];
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
function cTxt(ref: string, texto: string, estilo: number): string {
    return `<c r="${ref}" s="${estilo}" t="inlineStr"><is><t xml:space="preserve">${esc(texto)}</t></is></c>`;
}
function cNum(ref: string, n: number, estilo: number): string { return `<c r="${ref}" s="${estilo}"><v>${n}</v></c>`; }
function cVazio(ref: string, estilo: number): string { return `<c r="${ref}" s="${estilo}"/>`; }
const col = (i: number): string => String.fromCharCode(65 + i); // 0 = A (até S)

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
const S = ESTILO;
const ok = (v: N | undefined): v is number => v !== null && v !== undefined && !isNaN(v);

function abaPainel(e: EntradaPV): string {
    const linha = (r: number, ht: number, cels: string): string => `<row r="${r}" s="${S.linha}" customFormat="1" ht="${ht}" customHeight="1">${cels}</row>`;
    let rows = linha(1, ALTURA.r1, "");
    let t = cTxt("B2", e.titulo || "PAINEL DE VENDAS - DIA", S.titulo);
    for (let c = 2; c <= 12; c++) { t += cVazio(col(c) + "2", S.titulo); }
    rows += linha(2, ALTURA.titulo, t);
    rows += linha(3, ALTURA.faixa, "");
    // linha 4: faixas (mesmos estilos e mesclagens da aba)
    let g = cVazio("B4", S.faixaB) + cTxt("C4", "UF do vendedor", S.ufVend) + cVazio("D4", S.faixaB);
    g += cTxt("E4", "VENDIDO", S.vend) + cVazio("F4", S.vend) + cVazio("G4", S.vend);
    g += cTxt("H4", "FATURADO", S.fat) + cVazio("I4", S.fat) + cVazio("J4", S.fat);
    g += cTxt("K4", "CARTEIRA", S.cart);
    g += cTxt("L4", "CLIENTES (VENDIDO)", S.cli) + cVazio("M4", S.cli) + cVazio("N4", S.cli);
    g += cTxt("O4", "MIX (VENDIDO %)", S.mix) + ["P", "Q", "R", "S"].map(c => cVazio(c + "4", S.mix)).join("");
    rows += linha(4, ALTURA.grupo, g);
    // linha 5: cabeçalho
    const rot = rotulos(e.meses);
    const estFam = [S.mFB, S.mFC, S.mPT, S.mAP, S.mCT];
    let h = cTxt("B5", "REPRESENTANTES", S.cab) + cTxt("C5", "UF", S.cab) + cTxt("D5", "SUP", S.cab);
    rot.forEach((x, i) => { h += cTxt(col(4 + i) + "5", x, i < 10 ? S.cab : estFam[i - 10]); });
    rows += linha(5, ALTURA.cab, h);
    let r = 6;
    e.linhas.forEach(l => {
        let c = cTxt("B" + r, l.nome, S.rep) + cTxt("C" + r, l.uf, S.uf) + cTxt("D" + r, l.sup, S.sup);
        for (let i = 0; i < NCOLS; i++) {
            const ref = col(4 + i) + r, v = l.v[i];
            if (i < 10) {
                const est = i >= 7 ? S.qtd : S.num;
                c += ok(v) ? cNum(ref, i >= 7 ? Math.round(v) : v, est) : cVazio(ref, est);
            } else {
                const k = i - 10;
                const est = ok(v) ? (k === 0 ? S.mixO : k === 4 ? S.mixS : S.mixV) : (k === 0 ? S.mixOV : k === 4 ? S.mixSV : S.mixVV);
                c += ok(v) ? cNum(ref, v, est) : cVazio(ref, est);
            }
        }
        rows += linha(r, ALTURA.linha, c);
        r++;
    });
    if (!e.linhas.length) {
        rows += linha(r, ALTURA.linha, cTxt("B" + r, "Sem representantes para os filtros escolhidos.", S.rep));
        r++;
    }
    r++;
    ["Gerado em " + e.geradoEm + (e.regra ? " · Regra: " + e.regra : ""), e.periodo || "", e.filtros || ""].filter(x => !!x)
        .forEach(n => { rows += `<row r="${r}">${cTxt("B" + r, n, S.nota)}</row>`; r++; });
    return XML + `<worksheet ${NS}><dimension ref="A1:S${r - 1}"/><sheetViews><sheetView workbookViewId="0"/></sheetViews>` +
        `<sheetFormatPr defaultRowHeight="12.8"/><cols>${COLS_XML}</cols><sheetData>${rows}</sheetData>` +
        `<mergeCells count="5"><mergeCell ref="B2:M2"/><mergeCell ref="E4:G4"/><mergeCell ref="H4:J4"/><mergeCell ref="L4:N4"/><mergeCell ref="O4:S4"/></mergeCells>` +
        `<pageSetup paperSize="9" scale="100" orientation="portrait"/></worksheet>`;
}

// cabeçalho da aba Dados: grupo + coluna ("Vendido setembro 2026", "Mix FB"...)
export function cabecalhoDados(m: string[]): string[] {
    const rot = rotulos(m);
    const pref = ["Vendido", "Vendido", "Vendido", "Faturado", "Faturado", "Faturado", "Carteira", "Clientes (V)", "Clientes (V)", "Clientes (V)", "Mix %", "Mix %", "Mix %", "Mix %", "Mix %"];
    return ["Representante", "UF", "SUP", ...rot.map((x, i) => pref[i] + " " + x)];
}

function abaDados(e: EntradaPV): { xml: string; n: number } {
    const cab = cabecalhoDados(e.meses);
    let rows = `<row r="1">${cab.map((h, i) => cTxt(col(i) + "1", h, S.cab)).join("")}</row>`;
    let r = 2;
    e.linhas.forEach(l => {
        let c = cTxt("A" + r, l.nome, S.dTxt) + cTxt("B" + r, l.uf, S.dTxt) + cTxt("C" + r, l.sup, S.dTxt);
        l.v.forEach((v, i) => { if (ok(v)) { c += cNum(col(3 + i) + r, v, i < 7 ? S.dValor : i < 10 ? S.dQtd : S.dPct); } });
        rows += `<row r="${r}">${c}</row>`;
        r++;
    });
    const n = r - 1;
    const ult = col(cab.length - 1);
    return {
        n, xml: XML + `<worksheet ${NS}><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>` +
            `<cols><col min="1" max="1" width="40" customWidth="1"/><col min="2" max="2" width="6" customWidth="1"/><col min="3" max="3" width="13" customWidth="1"/><col min="4" max="${cab.length}" width="16" customWidth="1"/></cols>` +
            `<sheetData>${rows}</sheetData><autoFilter ref="A1:${ult}${Math.max(n, 1)}"/></worksheet>`
    };
}

export function gerarXlsxPainelVendas(e: EntradaPV): Uint8Array {
    const abas: { nome: string; xml: string }[] = [];
    let nDados = 0;
    const nomeAba = (e.aba || "PAINEL DE VENDAS").replace(/[\\/?*[\]:]/g, " ").slice(0, 31);
    if (e.opcoes.tabela || !e.opcoes.dados) { abas.push({ nome: nomeAba, xml: abaPainel(e) }); }
    if (e.opcoes.dados) { const d = abaDados(e); nDados = d.n; abas.push({ nome: "Dados", xml: d.xml }); }
    const iDados = abas.findIndex(a => a.nome === "Dados");
    const ultD = col(cabecalhoDados(e.meses).length - 1);
    const workbook = XML + `<workbook ${NS}><sheets>${abas.map((a, i) => `<sheet name="${esc(a.nome)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join("")}</sheets>` +
        (iDados >= 0 ? `<definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="${iDados}" hidden="1">Dados!$A$1:$${ultD}$${Math.max(nDados, 1)}</definedName></definedNames>` : "") + `</workbook>`;
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

// ---------- formato compacto (link da página de download e área de transferência) ----------
// Supervisões numa tabela (S, índice); cada linha: [nome, uf, iSup, v0..v14] com valores em R$ arredondados a inteiro
// (a aba mostra #,##0), clientes inteiros e mix em décimos de milésimo (4 casas); null = célula vazia.
// { tipo: "painelVendas", versao, titulo, aba, regra, meses, geradoEm, periodo, filtros, opcoes, S: [sups], L: [...] }
export interface CargaPV {
    tipo: "painelVendas"; versao: number; titulo: string; aba: string; regra: string; meses: string[];
    geradoEm: string; periodo?: string; filtros?: string; opcoes: OpcoesPV; S: string[]; L: (string | number | null)[][];
}
export function compactar(e: EntradaPV): CargaPV {
    const S: string[] = [];
    const iSup = (s: string): number => { let i = S.indexOf(s); if (i < 0) { i = S.length; S.push(s); } return i; };
    return {
        tipo: "painelVendas", versao: 1, titulo: e.titulo, aba: e.aba, regra: e.regra, meses: e.meses, geradoEm: e.geradoEm,
        periodo: e.periodo, filtros: e.filtros, opcoes: e.opcoes, S,
        L: e.linhas.map(l => {
            const v = l.v.map((x, i) => (!ok(x) ? null : i < 10 ? Math.round(x) : Math.round(x * 10000)));
            while (v.length && v[v.length - 1] === null) { v.pop(); }
            return [l.nome, l.uf, iSup(l.sup), ...v];
        })
    };
}
export function expandir(c: CargaPV): EntradaPV {
    const S = c.S || [];
    const linhas: LinhaPV[] = (c.L || []).map(a => {
        const v: N[] = [];
        for (let i = 0; i < NCOLS; i++) {
            const x = a[3 + i];
            v.push(x === null || x === undefined || x === "" ? null : (i < 10 ? Number(x) : Number(x) / 10000));
        }
        return { cod: "", nome: String(a[0] || ""), uf: String(a[1] || ""), sup: S[Number(a[2])] || "", v };
    });
    const o = c.opcoes || { tabela: true, dados: false };
    return {
        titulo: c.titulo || "PAINEL DE VENDAS - DIA", aba: c.aba || "PAINEL DE VENDAS", regra: c.regra || "", meses: c.meses || [],
        geradoEm: c.geradoEm || "", periodo: c.periodo || "", filtros: c.filtros || "", linhas, opcoes: { tabela: !!o.tabela, dados: !!o.dados }
    };
}

// ---------- texto para colar no Excel (botão Copiar) ----------
// Tabulações e quebras de linha; mesmo layout da aba (faixas, cabeçalho e uma linha por representante).
// Valores sem separador de milhar e com vírgula decimal (Excel em português lê como número); mix em % com vírgula.
export function tsv(e: EntradaPV): string {
    const limpa = (s: string): string => String(s).replace(/[\t\r\n]+/g, " ");
    const num = (v: N, i: number): string => {
        if (!ok(v)) { return ""; }
        if (i >= 10) { return (Math.round(v * 1000) / 10).toFixed(1).replace(".", ",") + "%"; }
        if (i >= 7) { return String(Math.round(v)); }
        return String(Math.round(v));
    };
    const linhas = [
        ["", "", "", "VENDIDO", "", "", "FATURADO", "", "", "CARTEIRA", "CLIENTES (VENDIDO)", "", "", "MIX (VENDIDO %)", "", "", "", ""].join("\t"),
        ["REPRESENTANTES", "UF", "SUP", ...rotulos(e.meses)].join("\t")
    ];
    e.linhas.forEach(l => linhas.push([limpa(l.nome), limpa(l.uf), limpa(l.sup), ...l.v.map((v, i) => num(v, i))].join("\t")));
    return linhas.join("\r\n") + "\r\n";
}
