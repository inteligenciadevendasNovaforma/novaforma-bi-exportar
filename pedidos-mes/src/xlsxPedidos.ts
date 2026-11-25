// Gerador de .xlsx do Pedidos Mês (sem bibliotecas): zip sem compressão + XML.
// A aba "PEDIDOS MÊS" repete a aba do exemplo: título B2:F2, faixa vazia, cabeçalho C4:G4 (Representante | Cód. Cliente |
// Razão Social | Família Comercial | Vendido), representante e cliente só na 1ª linha do grupo, faixa cinza depois de cada
// representante, valores "R$ "#,##0.00. Estilos, larguras e alturas vêm de modeloPedidos.ts (extraídos da planilha).
// Também ficam aqui o formato compacto (link / área de transferência) e o texto TSV do botão Copiar, usados pelo visual e pela página.
import { COLS_XML, STYLES_XML, ESTILO, ALTURA } from "./modeloPedidos";

export interface FamPM { ord: number; nome: string; v: number; }
export interface CliPM { cod: string; razao: string; fams: FamPM[]; }
export interface RepPM { cod: string; nome: string; total: number; clis: CliPM[]; }
export interface OpcoesPedidos { tabela: boolean; dados: boolean; }
export interface EntradaPedidos { titulo: string; geradoEm: string; periodo?: string; filtros?: string; reps: RepPM[]; opcoes: OpcoesPedidos; }

export const CABECALHO = ["Representante", "Cód. Cliente", "Razão Social", "Família Comercial", "Vendido"];

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
const v2 = (v: number): number => Math.round((Number(v) || 0) * 100) / 100;

function abaPedidos(e: EntradaPedidos): string {
    const linha = (r: number, ht: number, cels: string): string => `<row r="${r}" s="${S.linha}" customFormat="1" ht="${ht}" customHeight="1">${cels}</row>`;
    let rows = linha(1, ALTURA.r1, "");
    rows += linha(2, ALTURA.titulo, cTxt("B2", e.titulo || "PEDIDOS MÊS", S.titulo) + ["C", "D", "E", "F"].map(c => cVazio(c + "2", S.titulo)).join(""));
    rows += linha(3, ALTURA.faixa, "");
    rows += linha(4, ALTURA.cab, CABECALHO.map((h, i) => cTxt(String.fromCharCode(67 + i) + "4", h, S.cab)).join(""));
    let r = 5;
    e.reps.forEach(rep => {
        let primeiraRep = true;
        rep.clis.forEach(cli => {
            cli.fams.forEach((f, i) => {
                let c = primeiraRep ? cTxt("C" + r, rep.nome, S.rep) : cVazio("C" + r, S.repVazio);
                c += i === 0 ? cTxt("D" + r, cli.cod, S.txt) + cTxt("E" + r, cli.razao, S.txt) : cVazio("D" + r, S.txtVazio) + cVazio("E" + r, S.txtVazio);
                c += cTxt("F" + r, f.nome, S.fam) + cNum("G" + r, v2(f.v), S.valor);
                rows += linha(r, ALTURA.linha, c);
                primeiraRep = false; r++;
            });
        });
        rows += linha(r, ALTURA.sep, cVazio("C" + r, S.sep) + cVazio("D" + r, S.sep) + cVazio("E" + r, S.sepE) + cVazio("F" + r, S.sep) + cVazio("G" + r, S.sep));
        r++;
    });
    if (!e.reps.length) {
        rows += linha(r, ALTURA.linha, cTxt("C" + r, "Sem pedidos para os filtros escolhidos.", S.txt) + ["D", "E", "F", "G"].map(c => cVazio(c + r, S.txtVazio)).join(""));
        r++;
    }
    r++;
    ["Gerado em " + e.geradoEm, e.periodo || "", e.filtros || ""].filter(x => !!x).forEach(n => { rows += `<row r="${r}">${cTxt("C" + r, n, S.nota)}</row>`; r++; });
    return XML + `<worksheet ${NS}><dimension ref="A1:G${r - 1}"/><sheetViews><sheetView workbookViewId="0"/></sheetViews>` +
        `<sheetFormatPr defaultRowHeight="12.8"/><cols>${COLS_XML}</cols><sheetData>${rows}</sheetData>` +
        `<mergeCells count="1"><mergeCell ref="B2:F2"/></mergeCells>` +
        `<pageSetup paperSize="9" scale="100" orientation="portrait"/></worksheet>`;
}

function abaDados(e: EntradaPedidos): { xml: string; n: number } {
    let rows = `<row r="1">${CABECALHO.map((h, i) => cTxt(String.fromCharCode(65 + i) + "1", h, S.cab)).join("")}</row>`;
    let r = 2;
    e.reps.forEach(rep => rep.clis.forEach(cli => cli.fams.forEach(f => {
        rows += `<row r="${r}">${cTxt("A" + r, rep.nome, S.dTxt)}${cTxt("B" + r, cli.cod, S.dTxt)}${cTxt("C" + r, cli.razao, S.dTxt)}${cTxt("D" + r, f.nome, S.dTxt)}${cNum("E" + r, v2(f.v), S.dValor)}</row>`;
        r++;
    })));
    const n = r - 1;
    return {
        n, xml: XML + `<worksheet ${NS}><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>` +
            `<cols><col min="1" max="1" width="42.6" customWidth="1"/><col min="2" max="2" width="13" customWidth="1"/><col min="3" max="3" width="42.6" customWidth="1"/><col min="4" max="4" width="23.2" customWidth="1"/><col min="5" max="5" width="15" customWidth="1"/></cols>` +
            `<sheetData>${rows}</sheetData><autoFilter ref="A1:E${Math.max(n, 1)}"/></worksheet>`
    };
}

export function gerarXlsxPedidos(e: EntradaPedidos): Uint8Array {
    const abas: { nome: string; xml: string }[] = [];
    let nDados = 0;
    if (e.opcoes.tabela || !e.opcoes.dados) { abas.push({ nome: "PEDIDOS MÊS", xml: abaPedidos(e) }); }
    if (e.opcoes.dados) { const d = abaDados(e); nDados = d.n; abas.push({ nome: "Dados", xml: d.xml }); }
    const iDados = abas.findIndex(a => a.nome === "Dados");
    const workbook = XML + `<workbook ${NS}><sheets>${abas.map((a, i) => `<sheet name="${esc(a.nome)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join("")}</sheets>` +
        (iDados >= 0 ? `<definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="${iDados}" hidden="1">Dados!$A$1:$E$${Math.max(nDados, 1)}</definedName></definedNames>` : "") + `</workbook>`;
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
// Nomes de família numa tabela (F, índice = posição) e a hierarquia representante > cliente > [família, centavos, ...]:
//   { tipo, versao, titulo, geradoEm, periodo, filtros, opcoes, F: [nomes], R: [[nomeRep, [[codCli, razao, [iFam, centavos, ...]], ...]], ...] }
export interface CargaPedidos {
    tipo: "pedidosMes"; versao: number; titulo: string; geradoEm: string; periodo?: string; filtros?: string; opcoes: OpcoesPedidos;
    F: string[]; R: [string, [string, string, number[]][]][];
}
export function compactar(e: EntradaPedidos): CargaPedidos {
    const F: string[] = [];
    const iFam = (n: string): number => { let i = F.indexOf(n); if (i < 0) { i = F.length; F.push(n); } return i; };
    return {
        tipo: "pedidosMes", versao: 1, titulo: e.titulo, geradoEm: e.geradoEm, periodo: e.periodo, filtros: e.filtros, opcoes: e.opcoes, F,
        R: e.reps.map(r => [r.nome, r.clis.map(c => {
            const fv: number[] = [];
            c.fams.forEach(f => { fv.push(iFam(f.nome), Math.round((Number(f.v) || 0) * 100)); });
            return [c.cod, c.razao, fv] as [string, string, number[]];
        })] as [string, [string, string, number[]][]])
    };
}
export function expandir(c: CargaPedidos): EntradaPedidos {
    const F = c.F || [];
    const reps: RepPM[] = (c.R || []).map(([nome, clis]) => {
        let total = 0;
        const cs: CliPM[] = (clis || []).map(([cod, razao, fv]) => {
            const fams: FamPM[] = [];
            for (let i = 0; i + 1 < (fv || []).length; i += 2) { const v = fv[i + 1] / 100; total += v; fams.push({ ord: fv[i] + 1, nome: F[fv[i]] || "", v }); }
            return { cod: String(cod), razao: String(razao), fams };
        });
        return { cod: "", nome: String(nome), total, clis: cs };
    });
    const o = c.opcoes || { tabela: true, dados: false };
    return { titulo: c.titulo || "PEDIDOS MÊS", geradoEm: c.geradoEm || "", periodo: c.periodo || "", filtros: c.filtros || "", reps, opcoes: { tabela: !!o.tabela, dados: !!o.dados } };
}

// ---------- texto para colar no Excel (botão Copiar) ----------
// Tabulações e quebras de linha; mesmo layout da aba (representante e cliente só na 1ª linha do grupo).
// Valor sem separador de milhar e com vírgula decimal (Excel em português lê como número).
export function tsv(e: EntradaPedidos): string {
    const limpa = (s: string): string => String(s).replace(/[\t\r\n]+/g, " ");
    const linhas = [CABECALHO.join("\t")];
    e.reps.forEach(rep => {
        let primeira = true;
        rep.clis.forEach(cli => cli.fams.forEach((f, i) => {
            linhas.push([primeira ? limpa(rep.nome) : "", i === 0 ? limpa(cli.cod) : "", i === 0 ? limpa(cli.razao) : "", limpa(f.nome), v2(f.v).toFixed(2).replace(".", ",")].join("\t"));
            primeira = false;
        }));
    });
    return linhas.join("\r\n") + "\r\n";
}
