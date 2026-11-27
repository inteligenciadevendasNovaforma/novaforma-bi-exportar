"use strict";
// Oportunidade de Mix: carga compacta (visual -> página de download) e gerador do .xlsx (sem bibliotecas).
// Usado pelo visual (compactar) e pela página de download (expandir + gerarXlsxMix).
// Abas equivalentes às 3 do dashboard original: Resumo (por família), Sem venda por família (listas) e Representantes (mix e última venda).

export interface Familia { s: string; nome: string; modelo: string; cor: string; texto: string; }
// cores NovaForma (texto da faixa com contraste >= 4,5:1)
export const FAMILIAS: Familia[] = [
    { s: "FB", nome: "Forro Branco", modelo: "FORRO E ACESS. BR", cor: "#1F7A70", texto: "#FFFFFF" },
    { s: "FC", nome: "Forro Colorido", modelo: "FORRO E ACESS. COLOR", cor: "#A35F12", texto: "#FFFFFF" },
    { s: "PT", nome: "Porta PVC", modelo: "PORTA PVC", cor: "#0B2A4A", texto: "#FFFFFF" },
    { s: "AP", nome: "Acessórios para piso", modelo: "ACESS. PARA PISO", cor: "#1465B3", texto: "#FFFFFF" },
    { s: "CT", nome: "Cantinho", modelo: "CANTINHO", cor: "#56606B", texto: "#FFFFFF" }
];

export interface FamRep { n: number; v: number; u: string; }   // pedidos no período, vendido (R$), última venda (yyyy-mm-dd ou "")
export interface RepMix { cod: string; nome: string; sup: string; uf: string; ped: number; vend: number; v30: number; f: FamRep[]; }
export interface EntradaMix {
    titulo: string; geradoEm: string; periodo: string; ini: string; fim: string; filtros: string; ordem: string;
    base: number; cobertura: string; reps: RepMix[];
}
// carga compacta: R = [cod, nome, sup, uf, pedidos, vendido (centavos), vendido 30 dias (centavos), (n, centavos, yyyymmdd ou 0) x 5]
export interface CargaMix {
    tipo: "oportunidadeMix"; versao: number; t: string; g: string; p: string; i: string; f: string; fl: string; o: string; b: number; c: string;
    R: (string | number)[][];
}
export const PREFIXO_COLA = "NFOM1.";

const ymd = (s: string): number => (s ? Number(s.replace(/-/g, "")) : 0);
const iso = (n: number): string => { const s = String(n || ""); return s.length === 8 ? `${s.slice(0, 4)}-${s.slice(4, 6)}-${s.slice(6)}` : ""; };

export function compactar(e: EntradaMix): CargaMix {
    return {
        tipo: "oportunidadeMix", versao: 1, t: e.titulo, g: e.geradoEm, p: e.periodo, i: e.ini, f: e.fim, fl: e.filtros, o: e.ordem, b: e.base, c: e.cobertura,
        R: e.reps.map(r => {
            const l: (string | number)[] = [r.cod, r.nome, r.sup, r.uf, r.ped, Math.round(r.vend * 100), Math.round(r.v30 * 100)];
            r.f.forEach(x => l.push(x.n, Math.round(x.v * 100), ymd(x.u)));
            return l;
        })
    };
}
export function expandir(c: CargaMix): EntradaMix {
    const num = (x: unknown): number => { const n = Number(x); return isFinite(n) ? n : 0; };
    return {
        titulo: c.t || "Oportunidade de Mix", geradoEm: c.g || "", periodo: c.p || "", ini: c.i || "", fim: c.f || "", filtros: c.fl || "", ordem: c.o || "",
        base: num(c.b), cobertura: c.c || "",
        reps: (c.R || []).map(l => ({
            cod: String(l[0]), nome: String(l[1]), sup: String(l[2]), uf: String(l[3]), ped: num(l[4]), vend: num(l[5]) / 100, v30: num(l[6]) / 100,
            f: FAMILIAS.map((_, k) => ({ n: num(l[7 + k * 3]), v: num(l[8 + k * 3]) / 100, u: iso(num(l[9 + k * 3])) }))
        }))
    };
}

// dias corridos entre a última venda e o fim do período (null = sem registro)
export function diasSem(u: string, fim: string): number | null {
    if (!u || !fim) { return null; }
    return Math.round((Date.parse(fim + "T12:00:00Z") - Date.parse(u + "T12:00:00Z")) / 86400000);
}
export const mixDe = (r: RepMix): number => r.f.filter(x => x.n > 0).length;

// ---------------- xlsx ----------------
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
    return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/[\x00-\x08\x0b\x0c\x0e-\x1f]/g, "");
}
function col(n: number): string { let s = ""; for (n++; n; n = Math.floor((n - 1) / 26)) { s = String.fromCharCode(65 + (n - 1) % 26) + s; } return s; }
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
        const local = new Uint8Array([0x50, 0x4B, 0x03, 0x04, ...u16(20), ...u16(0x0800), ...u16(0), ...u16(0), ...u16(0x21),
            ...u32(crc), ...u32(a.dados.length), ...u32(a.dados.length), ...u16(nome.length), ...u16(0)]);
        partes.push(local, nome, a.dados);
        central.push(new Uint8Array([0x50, 0x4B, 0x01, 0x02, ...u16(20), ...u16(20), ...u16(0x0800), ...u16(0), ...u16(0), ...u16(0x21),
            ...u32(crc), ...u32(a.dados.length), ...u32(a.dados.length), ...u16(nome.length), ...u16(0), ...u16(0),
            ...u16(0), ...u16(0), ...u32(0), ...u32(offset)]), nome);
        offset += local.length + nome.length + a.dados.length;
    }
    const tamCentral = central.reduce((s, p) => s + p.length, 0);
    const fim = new Uint8Array([0x50, 0x4B, 0x05, 0x06, ...u16(0), ...u16(0), ...u16(arquivos.length), ...u16(arquivos.length), ...u32(tamCentral), ...u32(offset), ...u16(0)]);
    const todos = [...partes, ...central, fim];
    const saida = new Uint8Array(todos.reduce((s, p) => s + p.length, 0));
    let pos = 0;
    for (const p of todos) { saida.set(p, pos); pos += p.length; }
    return saida;
}

const NS = `xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`;
const XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`;
// índices de cellXfs (ver styles)
const S = {
    normal: 0, titulo: 1, sub: 2, nota: 3, cab: 4, txt: 5, int: 6, brl: 7, pct: 8, data: 9, totTxt: 10, totInt: 11, totBrl: 12, totPct: 13,
    famBase: 14,   // 14..18 = faixa de cada família (FB..CT)
    cont: 19, lista: 20, alerta: 21, semVenda: 22, txtC: 23
};
type Cel = { t: "s"; v: string; s: number } | { t: "n"; v: number; s: number } | { t: "e"; s: number };
const cs = (v: string, s = S.txt): Cel => ({ t: "s", v, s });
const cn = (v: number, s = S.int): Cel => ({ t: "n", v, s });
const ce = (s = S.txt): Cel => ({ t: "e", s });
function serial(isoData: string): number | null {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(isoData || "");
    return m ? Math.round((Date.UTC(+m[1], +m[2] - 1, +m[3]) - Date.UTC(1899, 11, 30)) / 86400000) : null;
}
function linha(r: number, cels: (Cel | null)[], ht?: number): string {
    return `<row r="${r}"${ht ? ` ht="${ht}" customHeight="1"` : ""}>` + cels.map((c, i) => {
        if (!c) { return ""; }
        const ref = col(i) + r;
        if (c.t === "s") { return `<c r="${ref}" t="inlineStr" s="${c.s}"><is><t xml:space="preserve">${esc(c.v)}</t></is></c>`; }
        if (c.t === "n") { return `<c r="${ref}" s="${c.s}"><v>${isFinite(c.v) ? c.v : 0}</v></c>`; }
        return `<c r="${ref}" s="${c.s}"/>`;
    }).join("") + `</row>`;
}
function aba(cols: number[], linhas: string, extra: { merges?: string[]; congelar?: string; filtro?: string; paisagem?: boolean } = {}): string {
    const m = extra.merges || [];
    const pane = extra.congelar ? `<pane ySplit="${Number(/\d+/.exec(extra.congelar)?.[0] || 1) - 1}" topLeftCell="${extra.congelar}" activePane="bottomLeft" state="frozen"/>` : "";
    return XML + `<worksheet ${NS}><sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><sheetViews><sheetView showGridLines="0" workbookViewId="0">${pane}</sheetView></sheetViews>` +
        `<cols>${cols.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join("")}</cols>` +
        `<sheetData>${linhas}</sheetData>` + (extra.filtro ? `<autoFilter ref="${extra.filtro}"/>` : "") +
        (m.length ? `<mergeCells count="${m.length}">${m.map(x => `<mergeCell ref="${x}"/>`).join("")}</mergeCells>` : "") +
        `<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/><pageSetup paperSize="9" orientation="${extra.paisagem ? "landscape" : "portrait"}" fitToWidth="1" fitToHeight="0"/></worksheet>`;
}

export function gerarXlsxMix(e: EntradaMix): Uint8Array {
    const reps = e.reps;
    const listas = FAMILIAS.map((_, k) => reps.filter(r => r.f[k].n === 0));
    const abas: { nome: string; xml: string; filtro?: string }[] = [];

    // ---- Resumo ----
    {
        let x = ""; const mg: string[] = []; let r = 2;
        const L = 8; // colunas B..H
        const larga = (texto: string, s: number, ht?: number): void => {
            const cels: (Cel | null)[] = [null, cs(texto, s)]; for (let i = 2; i < L; i++) { cels.push(ce(s)); }
            x += linha(r, cels, ht); mg.push(`B${r}:${col(L - 1)}${r}`); r++;
        };
        larga(e.titulo, S.titulo, 28);
        larga("Representantes sem venda por família de produto", S.sub, 20);
        if (e.periodo) { larga(e.periodo, S.nota); }
        larga(e.filtros || "Filtros: nenhum (todos os representantes da base)", S.nota);
        larga(`Representantes no filtro: ${reps.length.toLocaleString("pt-BR")} de ${e.base.toLocaleString("pt-BR")} na base`, S.nota);
        if (e.cobertura) { larga(e.cobertura, S.alerta, 30); }
        r++;
        x += linha(r, [null, cs("Família", S.cab), cs("Sigla", S.cab), cs("Sem venda", S.cab), cs("Com venda", S.cab), cs("% sem venda", S.cab), cs("Pedidos no período", S.cab), cs("Vendido no período (R$)", S.cab)], 30);
        const ini = r + 1; r++;
        FAMILIAS.forEach((f, k) => {
            const sem = listas[k].length, com = reps.length - sem;
            const ped = reps.reduce((a, p) => a + p.f[k].n, 0), vend = reps.reduce((a, p) => a + p.f[k].v, 0);
            x += linha(r, [null, cs(f.nome), cs(f.s, S.txtC), cn(sem), cn(com), cn(reps.length ? sem / reps.length : 0, S.pct), cn(ped), cn(Math.round(vend * 100) / 100, S.brl)], 20);
            r++;
        });
        const tPed = reps.reduce((a, p) => a + p.ped, 0), tVend = reps.reduce((a, p) => a + p.vend, 0);
        x += linha(r, [null, cs("Total das 5 famílias (pedidos distintos)", S.totTxt), cs("", S.totTxt), ce(S.totInt), ce(S.totInt), ce(S.totPct), cn(tPed, S.totInt), cn(Math.round(tVend * 100) / 100, S.totBrl)], 22);
        r += 2;
        const completo = reps.filter(p => mixDe(p) === 5).length, nenhum = reps.filter(p => mixDe(p) === 0).length;
        x += linha(r, [null, cs("Mix completo (vendeu nas 5 famílias)", S.txt), ce(), cn(completo)], 20); r++;
        x += linha(r, [null, cs("Sem venda em nenhuma das 5 famílias", S.txt), ce(), cn(nenhum)], 20); r++;
        x += linha(r, [null, cs("Ticket médio (vendido ÷ pedidos distintos)", S.txt), ce(), cn(tPed ? Math.round(tVend / tPed * 100) / 100 : 0, S.brl)], 20); r += 2;
        void ini;
        [
            "Sem venda = nenhum pedido válido da família no período, pela data de entrada do pedido. Pedido válido: faturamento e carteira (sem recusa) com saldo positivo por documento; o mesmo pedido nas duas bases conta uma vez.",
            "Base: representantes ativos e não bloqueados no SAP, sem Venda Direta. A carteira é a posição atual carregada.",
            "Gerado em " + e.geradoEm + (e.ordem ? " · Ordenação das listas: " + e.ordem : "")
        ].forEach(n => larga(n, S.nota, 30));
        abas.push({ nome: "Resumo", xml: aba([1.5, 34, 8, 12, 12, 13, 17, 22], x, { merges: mg }) });
    }

    // ---- Sem venda por família (listas lado a lado, como no painel) ----
    {
        let x = ""; const mg: string[] = []; let r = 2;
        const t: (Cel | null)[] = [null, cs(e.titulo + " · Representantes sem venda por família", S.titulo)]; for (let i = 2; i <= 5; i++) { t.push(ce(S.titulo)); }
        x += linha(r, t, 26); mg.push(`B${r}:F${r}`); r++;
        if (e.periodo) { const p: (Cel | null)[] = [null, cs(e.periodo, S.nota)]; for (let i = 2; i <= 5; i++) { p.push(ce(S.nota)); } x += linha(r, p); mg.push(`B${r}:F${r}`); r++; }
        r++;
        const ini = r;
        x += linha(r, [null, ...FAMILIAS.map((f, k) => cs(f.s + " · " + f.nome, S.famBase + k))], 26); r++;
        x += linha(r, [null, ...listas.map(l => cs(l.length.toLocaleString("pt-BR") + " sem venda", S.cont))], 20); r++;
        const max = Math.max(0, ...listas.map(l => l.length));
        if (!max) { x += linha(r, [null, ...FAMILIAS.map(() => cs("Nenhum representante sem venda", S.lista))]); r++; }
        for (let i = 0; i < max; i++) {
            x += linha(r, [null, ...listas.map(l => l[i] ? cs(l[i].nome + (l[i].uf ? " (" + l[i].uf + ")" : ""), S.lista) : ce(S.lista))]);
            r++;
        }
        abas.push({ nome: "Sem venda por família", xml: aba([1.5, 40, 40, 40, 40, 40], x, { merges: mg, congelar: "A" + (ini + 2), paisagem: true }) });
    }

    // ---- Representantes (dados por representante: mix e última venda) ----
    {
        const cab = ["Código", "Representante", "Supervisão", "UF", "Vendido no período (R$)", "Pedidos válidos", "Ticket médio (R$)", "Mix vendido (de 5)", "Vendido 30 dias (R$)"];
        FAMILIAS.forEach(f => cab.push(f.s + " pedidos", f.s + " última venda", f.s + " dias sem venda"));
        cab.push("Faltam no período");
        let x = linha(1, cab.map(h => cs(h, S.cab)), 32);
        reps.forEach((p, i) => {
            const r = i + 2;
            const cels: Cel[] = [cs(p.cod), cs(p.nome), cs(p.sup), cs(p.uf, S.txtC), cn(p.vend, S.brl), cn(p.ped), p.ped ? cn(Math.round(p.vend / p.ped * 100) / 100, S.brl) : ce(), cn(mixDe(p)), cn(p.v30, S.brl)];
            p.f.forEach(f => {
                const d = serial(f.u), dias = diasSem(f.u, e.fim);
                cels.push(f.n ? cn(f.n) : cs("Sem venda", S.semVenda));
                cels.push(d === null ? cs("Sem registro", S.txtC) : cn(d, S.data));
                cels.push(dias === null ? ce() : cn(dias));
            });
            cels.push(cs(FAMILIAS.filter((_, k) => p.f[k].n === 0).map(f => f.s).join(", ")));
            x += linha(r, cels);
        });
        const ult = col(cab.length - 1) + Math.max(reps.length + 1, 1);
        abas.push({
            nome: "Representantes", filtro: "A1:" + ult,
            xml: aba([10, 40, 16, 6, 16, 10, 14, 10, 16, ...FAMILIAS.flatMap(() => [10, 13, 10]), 18], x, { congelar: "C2", filtro: "A1:" + ult, paisagem: true })
        });
    }

    const fams = FAMILIAS.map(f => `<fill><patternFill patternType="solid"><fgColor rgb="FF${f.cor.slice(1).toUpperCase()}"/><bgColor indexed="64"/></patternFill></fill>`).join("");
    const xf = (num: number, font: number, fill: number, border: number, al = "", extra = ""): string =>
        `<xf numFmtId="${num}" fontId="${font}" fillId="${fill}" borderId="${border}" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"${al ? ` applyAlignment="1"><alignment ${al}/></xf>` : `${extra}/>`}`;
    const styles = XML + `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
        `<numFmts count="3"><numFmt numFmtId="164" formatCode="#,##0.00"/><numFmt numFmtId="165" formatCode="dd/mm/yyyy"/><numFmt numFmtId="166" formatCode="0.0%"/></numFmts>` +
        `<fonts count="8"><font><sz val="11"/><name val="Segoe UI"/></font>` +
        `<font><b/><sz val="16"/><color rgb="FF0B2A4A"/><name val="Segoe UI"/></font>` +
        `<font><sz val="12"/><color rgb="FF3A4756"/><name val="Segoe UI"/></font>` +
        `<font><sz val="10"/><color rgb="FF4B5563"/><name val="Segoe UI"/></font>` +
        `<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Segoe UI"/></font>` +
        `<font><b/><sz val="11"/><color rgb="FF0B2A4A"/><name val="Segoe UI"/></font>` +
        `<font><b/><sz val="11"/><color rgb="FF7A2E0B"/><name val="Segoe UI"/></font>` +
        `<font><b/><sz val="11"/><color rgb="FFB42318"/><name val="Segoe UI"/></font></fonts>` +
        `<fills count="${5 + FAMILIAS.length}"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>` +
        `<fill><patternFill patternType="solid"><fgColor rgb="FF0B2A4A"/><bgColor indexed="64"/></patternFill></fill>` +
        `<fill><patternFill patternType="solid"><fgColor rgb="FFE8EDF3"/><bgColor indexed="64"/></patternFill></fill>` +
        `<fill><patternFill patternType="solid"><fgColor rgb="FFFFF4E5"/><bgColor indexed="64"/></patternFill></fill>` + fams + `</fills>` +
        `<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>` +
        `<border><left/><right/><top/><bottom style="thin"><color rgb="FFD5DCE4"/></bottom><diagonal/></border></borders>` +
        `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
        `<cellXfs count="24">` +
        xf(0, 0, 0, 0) +                                                    // 0 normal
        xf(0, 1, 0, 0, `vertical="center"`) +                                // 1 título
        xf(0, 2, 0, 0, `vertical="center"`) +                                // 2 subtítulo
        xf(0, 3, 0, 0, `vertical="top" wrapText="1"`) +                      // 3 nota
        xf(0, 4, 2, 0, `horizontal="center" vertical="center" wrapText="1"`) + // 4 cabeçalho
        xf(0, 0, 0, 1, `vertical="center"`) +                                // 5 texto
        xf(3, 0, 0, 1, `vertical="center"`) +                                // 6 inteiro
        xf(164, 0, 0, 1, `vertical="center"`) +                              // 7 R$
        xf(166, 0, 0, 1, `vertical="center"`) +                              // 8 %
        xf(165, 0, 0, 1, `horizontal="center" vertical="center"`) +          // 9 data
        xf(0, 5, 3, 1, `vertical="center"`) +                                // 10 total texto
        xf(3, 5, 3, 1, `vertical="center"`) +                                // 11 total inteiro
        xf(164, 5, 3, 1, `vertical="center"`) +                              // 12 total R$
        xf(166, 5, 3, 1, `vertical="center"`) +                              // 13 total %
        FAMILIAS.map((_, k) => xf(0, 4, 5 + k, 0, `horizontal="center" vertical="center"`)).join("") + // 14..18 família
        xf(0, 5, 3, 1, `horizontal="center" vertical="center"`) +            // 19 contagem
        xf(0, 0, 0, 1, `vertical="center"`) +                                // 20 nome na lista
        xf(0, 6, 4, 0, `vertical="center" wrapText="1"`) +                   // 21 alerta de cobertura
        xf(0, 7, 0, 1, `horizontal="center" vertical="center"`) +            // 22 "Sem venda"
        xf(0, 0, 0, 1, `horizontal="center" vertical="center"`) +            // 23 texto centralizado
        `</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;

    const nomes = abas.map((a, i) => a.filtro ? `<definedName name="_xlnm._FilterDatabase" localSheetId="${i}" hidden="1">'${a.nome}'!$${a.filtro.replace(":", ":$").replace(/([A-Z]+)(\d+)/g, "$1$$$2")}</definedName>` : "").join("");
    const workbook = XML + `<workbook ${NS}><sheets>${abas.map((a, i) => `<sheet name="${esc(a.nome)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join("")}</sheets>` +
        (nomes ? `<definedNames>${nomes}</definedNames>` : "") + `</workbook>`;
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
