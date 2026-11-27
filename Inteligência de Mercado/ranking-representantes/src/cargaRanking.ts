"use strict";
// Carga do PDF do Ranking (visual -> página de download). Usado pelo visual (compactar) e pela página (expandir).
// Uma página por Supervisão × UF; os textos já vêm prontos do visual (o desenho na página fica igual ao PDF antigo).

export interface RepPdf { r: string; v: number; q: number; pos: number; zebra: number; }
export interface PaginaPdf { reps: RepPdf[]; isPos: boolean; eyebrow: string; title: string; sub: string; foot: string; stats: { l: string; v: string }[]; }

// tipo "ranking" (v1): pos = 1 para Positivação; t = título; p = páginas
// cada página: e = linha de cima (supervisão · UF), s = subtítulo, f = rodapé, k = totais [rótulo, valor],
// r = linhas [posição, representante, vendido em centavos, clientes] na ordem do ranking
export interface CargaRanking {
    tipo: "ranking";
    versao: number;
    pos: number;
    t: string;
    p: { e: string; s: string; f: string; k: [string, string][]; r: [number, string, number, number][] }[];
}

export const LIMITE_PAGINAS = 60;
export const PREFIXO_COLA = "NFRK1.";

export function compactar(pages: PaginaPdf[]): CargaRanking {
    return {
        tipo: "ranking", versao: 1,
        pos: pages.length && pages[0].isPos ? 1 : 0,
        t: pages.length ? pages[0].title : "",
        p: pages.map(pg => ({
            e: pg.eyebrow, s: pg.sub, f: pg.foot,
            k: pg.stats.map(s => [s.l, s.v] as [string, string]),
            r: pg.reps.map(x => [x.pos, x.r, Math.round(x.v * 100), x.q] as [number, string, number, number])
        }))
    };
}

export function expandir(c: CargaRanking): PaginaPdf[] {
    if (!c || c.tipo !== "ranking" || !Array.isArray(c.p)) { throw new Error("Os dados do ranking estão incompletos."); }
    if (c.p.length > LIMITE_PAGINAS) { throw new Error("O PDF passaria de " + LIMITE_PAGINAS + " páginas. Filtre Supervisão ou UF e exporte de novo."); }
    const isPos = c.pos === 1;
    return c.p.map(p => ({
        isPos, title: c.t || ("Ranking de " + (isPos ? "Positivação" : "Vendas")),
        eyebrow: p.e || "", sub: p.s || "", foot: p.f || "",
        stats: (p.k || []).map(k => ({ l: String(k[0]), v: String(k[1]) })),
        reps: (p.r || []).map((x, i) => ({ pos: Number(x[0]), r: String(x[1]), v: Number(x[2]) / 100, q: Number(x[3]), zebra: i }))
    })).filter(p => p.reps.length > 0);
}
