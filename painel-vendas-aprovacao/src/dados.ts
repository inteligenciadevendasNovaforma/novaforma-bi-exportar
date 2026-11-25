// Leitura da medida de carga do Painel de Vendas: uma linha de texto por representante (linha do visual):
//   T|codigo|representante|uf|sup|V1|V2|V3|F1|F2|F3|carteira|Q1|Q2|Q3|FB|FC|PT|AP|CT     (T = A Aprovação, P APP; vazio = célula vazia)
// No 1º representante vem o cabeçalho depois de "~":  titulo1|titulo2|titulo3|dataRef|anos|meses|dias|supervisoes|representantes|ufs|familias|materiais
import { LinhaPV, N, NCOLS } from "./xlsxPainelVendas";

export type Ordem = "nome" | "vendido" | "codigo";
export interface Carga { cabecalho: string[]; linhas: LinhaPV[]; }
// resultado da conferência do formato: "ok", "outra" (medida do outro visual), "errada" (outra medida qualquer)
export type Formato = "ok" | "outra" | "errada";

const numCod = (s: string): number => (/^\d+$/.test(s) ? Number(s) : Number.MAX_SAFE_INTEGER);
function n(s: string | undefined): N {
    if (s === undefined || s === "") { return null; }
    const v = Number(s);
    return isNaN(v) ? null : v;
}

// confere a 1ª célula que tem linha de dados (a do 1º representante pode trazer só o cabeçalho, que começa com "~")
export function conferir(celulas: string[], tipo: string, outro: string): Formato {
    const amostra = celulas.map(x => x.split("~")[0]).find(x => x.trim().length > 0);
    if (!amostra) { return celulas.some(x => x.indexOf("~") >= 0) ? "ok" : "errada"; }
    const c = amostra.split(/\r?\n/)[0].split("|");
    const numeros = c.slice(5, 20).every(x => x === "" || /^-?\d+(\.\d+)?$/.test(x));
    if (c.length !== 20 || !numeros) { return "errada"; }
    if (c[0] === tipo) { return "ok"; }
    return c[0] === outro ? "outra" : "errada";
}

export function lerCarga(celulas: string[], ordem: Ordem): Carga {
    let cabecalho: string[] = [];
    const porCod = new Map<string, LinhaPV>();
    celulas.forEach(cel => {
        (cel || "").split(/\r?\n/).forEach(l => {
            if (!l) { return; }
            const til = l.indexOf("~");
            const corpo = til >= 0 ? l.slice(0, til) : l;
            if (til >= 0) { cabecalho = l.slice(til + 1).split("|"); }
            const c = corpo.split("|");
            if (c.length < 20) { return; }
            const v: N[] = [];
            for (let i = 0; i < NCOLS; i++) { v.push(n(c[5 + i])); }
            porCod.set(c[1], { cod: c[1], nome: c[2], uf: c[3], sup: c[4], v });
        });
    });
    const linhas = Array.from(porCod.values());
    const porNome = (a: LinhaPV, b: LinhaPV): number => a.nome.localeCompare(b.nome, "pt-BR") || numCod(a.cod) - numCod(b.cod);
    if (ordem === "vendido") { linhas.sort((a, b) => (b.v[2] || 0) - (a.v[2] || 0) || porNome(a, b)); }
    else if (ordem === "codigo") { linhas.sort((a, b) => numCod(a.cod) - numCod(b.cod) || porNome(a, b)); }
    else { linhas.sort(porNome); }
    return { cabecalho, linhas };
}
