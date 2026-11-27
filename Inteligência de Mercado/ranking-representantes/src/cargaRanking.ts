"use strict";
// Carga do PDF do Ranking (visual -> página de download). Usado pelo visual (compactar) e pela página (expandir).
// v1.0.7+: UMA lista contínua, igual à tela (sem separar por Supervisão ou UF). A página de download quebra em várias
// páginas só quando a lista não cabe. Os textos já vêm prontos do visual. Cargas antigas (v1, uma página por Supervisão × UF) ainda são aceitas.

export interface RepPdf { r: string; v: number; q: number; pos: number; zebra: number; }
// showV: mostra a coluna "Vendido (R$)" (false = "Ocultar valores" na tela)
export interface PaginaPdf { reps: RepPdf[]; isPos: boolean; showV?: boolean; eyebrow: string; title: string; sub: string; foot: string; stats: { l: string; v: string }[]; }

// tipo "ranking": pos = 1 para Positivação; sv = 0 quando os valores estão ocultos; t = título
// versão 2 (lista única): e = linha de cima (filtros da tela), s = subtítulo, f = rodapé, k = totais [rótulo, valor],
//   r = linhas [posição, representante, vendido em centavos, clientes] na ordem do ranking
// versão 1 (antiga): p = uma entrada {e, s, f, k, r} por página
export interface CargaRanking {
    tipo: "ranking";
    versao: number;
    pos: number;
    sv?: number;
    saida?: "jpg";   // v1.0.8+: "jpg" = uma imagem única (JPEG) em vez do PDF
    t: string;
    e?: string; s?: string; f?: string; k?: [string, string][]; r?: [number, string, number, number][];
    p?: { e: string; s: string; f: string; k: [string, string][]; r: [number, string, number, number][] }[];
}

export const LIMITE_PAGINAS = 60;
export const PREFIXO_COLA = "NFRK1.";

// ---- layout das páginas (canvas de 880 px de largura; página A4 = 880 x 1245) ----
export const PAG_H = 1245;
export const TOPO1 = 343;      // primeira linha da 1ª página (abaixo do cabeçalho grande)
export const TOPON = 150;      // primeira linha das páginas seguintes (faixa fina + cabeçalho da tabela)
export const PASSO = 47.4;
export const LINHA_H = 41;
export const MARGEM_PAG = 64;  // espaço no pé das páginas que não são a última (número da página)
export const rodapeH = (linhas: number) => 108 + (Math.max(1, linhas) - 1) * 18;
const cabem = (topo: number, limite: number) => Math.max(1, Math.floor((limite - topo - LINHA_H) / PASSO) + 1);

// Quantas linhas vão em cada página (n linhas; "linhas" = nº de linhas do texto do rodapé). Uma página só, se couber.
export function particionar(n: number, linhas: number): number[] {
    if (n <= 0) return [];
    const ultima = (topo: number) => cabem(topo, PAG_H - rodapeH(linhas));
    const cheia = (topo: number) => cabem(topo, PAG_H - MARGEM_PAG);
    if (n <= ultima(TOPO1)) return [n];
    const out: number[] = [];
    let resto = n;
    const c1 = cheia(TOPO1), cn = cheia(TOPON), cu = ultima(TOPON);
    out.push(Math.min(c1, resto)); resto -= out[0];
    while (resto > cu) { const k = Math.min(cn, resto); out.push(k); resto -= k; }
    if (resto > 0) out.push(resto);
    return out;
}

// ---- imagem única (JPEG): lista contínua, sem páginas ----
export const IMG_W = 880;                    // largura lógica (a mesma do PDF); em pixels = IMG_W x escala
export const IMG_MAX_ALTURA = 16000;         // altura máxima segura de um canvas, em pixels
export const IMG_MAX_AREA = 268000000;       // área máxima (pixels) dos navegadores de computador
export const IMG_MAX_AREA_IOS = 16700000;    // Safari do iPhone/iPad aceita bem menos
export const MSG_IMAGEM_GRANDE = "A lista é grande demais para uma imagem única; use o PDF ou filtre por supervisão.";
export const alturaImagem = (n: number, linhas: number) => Math.round(Math.max(946, TOPO1 + (Math.max(1, n) - 1) * PASSO + LINHA_H + rodapeH(linhas)));
// Maior escala (2 -> 1,5 -> 1) em que a lista inteira cabe num canvas; 0 = não cabe nem na escala 1
export function escalaImagem(n: number, linhas: number, ios = false): number {
    const H = alturaImagem(n, linhas), area = ios ? IMG_MAX_AREA_IOS : IMG_MAX_AREA;
    for (const s of [2, 1.5, 1]) { if (H * s <= IMG_MAX_ALTURA && IMG_W * s * H * s <= area) return s; }
    return 0;
}
export const ehIos = () => typeof navigator !== "undefined" && (/iP(hone|ad|od)/.test(navigator.userAgent) || (navigator.platform === "MacIntel" && (navigator as any).maxTouchPoints > 1));

export function compactar(pg: PaginaPdf, saida?: "jpg"): CargaRanking {
    return {
        tipo: "ranking", versao: 2,
        ...(saida === "jpg" ? { saida } : {}),
        pos: pg.isPos ? 1 : 0, sv: pg.showV === false ? 0 : 1,
        t: pg.title, e: pg.eyebrow, s: pg.sub, f: pg.foot,
        k: pg.stats.map(s => [s.l, s.v] as [string, string]),
        r: pg.reps.map(x => [x.pos, x.r, Math.round(x.v * 100), x.q] as [number, string, number, number])
    };
}

const reps = (r: [number, string, number, number][] | undefined) =>
    (r || []).map((x, i) => ({ pos: Number(x[0]), r: String(x[1]), v: Number(x[2]) / 100, q: Number(x[3]), zebra: i }));

// Carga -> documentos a desenhar (v2: um; v1: um por Supervisão × UF)
export function expandir(c: CargaRanking): PaginaPdf[] {
    if (!c || c.tipo !== "ranking" || (!Array.isArray(c.p) && !Array.isArray(c.r))) { throw new Error("Os dados do ranking estão incompletos."); }
    const isPos = c.pos === 1, showV = c.sv !== 0;
    const titulo = c.t || ("Ranking de " + (isPos ? "Positivação" : "Vendas"));
    if (Array.isArray(c.p)) {
        if (c.p.length > LIMITE_PAGINAS) { throw new Error("O PDF passaria de " + LIMITE_PAGINAS + " páginas. Filtre Supervisão ou UF e exporte de novo."); }
        return c.p.map(p => ({
            isPos, showV, title: titulo, eyebrow: p.e || "", sub: p.s || "", foot: p.f || "",
            stats: (p.k || []).map(k => ({ l: String(k[0]), v: String(k[1]) })), reps: reps(p.r)
        })).filter(p => p.reps.length > 0);
    }
    const d: PaginaPdf = {
        isPos, showV, title: titulo, eyebrow: c.e || "", sub: c.s || "", foot: c.f || "",
        stats: (c.k || []).map(k => ({ l: String(k[0]), v: String(k[1]) })), reps: reps(c.r)
    };
    return d.reps.length ? [d] : [];
}
