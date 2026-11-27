"use strict";
// Desenho do PDF do Ranking de Representantes. Roda na PÁGINA de download (não no visual).
// Mesmo desenho do PDF que o visual gerava até a v1.0.4: canvas 880 px de largura (escala 2) -> JPEG -> PDF montado à mão,
// v1.0.7: lista contínua como na tela (várias páginas só se não couber; cabeçalho da tabela repetido), cabeçalho azul-marinho, medalhas, barras e rodapé. Fontes Barlow embutidas (fonts.ts).
import { FONTS } from "./fonts";
import { CargaRanking, PaginaPdf, expandir, particionar, LIMITE_PAGINAS, PAG_H, TOPO1, TOPON, PASSO, LINHA_H, rodapeH, alturaImagem, escalasImagem, areaMaximaImagem, IMG_QUALIDADE, msgImagemGrande } from "./cargaRanking";
import { LOGO_PATHS, LOGO_W, LOGO_H, LOGO_DX, LOGO_DY } from "./logo";

const FONT_H = '"Barlow Condensed","Bahnschrift SemiBold Condensed","Arial Narrow","Segoe UI",sans-serif';
const FONT_B = '"Barlow","Segoe UI",system-ui,sans-serif';
const BADGE = ["#E0A526", "#8A96A3", "#B9773E"];
const brl = (n: number) => n.toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const int = (n: number) => n.toLocaleString("pt-BR");

interface Imagem { jpg: Uint8Array; w: number; h: number; pw: number; ph: number; }

let fontsReady: Promise<void> | null = null;
function loadFonts(): Promise<void> {
    if (fontsReady) return fontsReady;
    fontsReady = (async () => {
        try {
            const fonts: any = (document as any).fonts;
            if (!fonts || typeof FontFace === "undefined") return;
            await Promise.all(FONTS.map(async f => {
                const bin = atob(f.b64), buf = new Uint8Array(bin.length);
                for (let i = 0; i < bin.length; i++) buf[i] = bin.charCodeAt(i);
                const face = new FontFace(f.family, buf.buffer, { weight: f.weight });
                await face.load(); fonts.add(face);
            }));
            // garante que cada família/peso está pronto para o canvas (senão o primeiro desenho cairia na fonte de reserva)
            await Promise.all(FONTS.map(f => fonts.load(f.weight + " 16px \"" + f.family + "\"", "Ráç0123456789")));
            if (fonts.ready) await fonts.ready;
        } catch (e) { /* usa as fontes de reserva */ }
    })();
    return fontsReady;
}

function roundRect(c: CanvasRenderingContext2D, x: number, y: number, w: number, h: number, r: number) {
    c.beginPath(); c.moveTo(x + r, y); c.arcTo(x + w, y, x + w, y + h, r); c.arcTo(x + w, y + h, x, y + h, r); c.arcTo(x, y + h, x, y, r); c.arcTo(x, y, x + w, y, r); c.closePath();
}
function wrapText(c: CanvasRenderingContext2D, text: string, maxW: number) {
    const lines: string[] = []; let cur = "";
    text.split(" ").forEach(w => { const t = cur ? cur + " " + w : w; if (cur && c.measureText(t).width > maxW) { lines.push(cur); cur = w; } else cur = t; });
    if (cur) lines.push(cur); return lines;
}
function fitText(c: CanvasRenderingContext2D, text: string, maxW: number) {
    if (c.measureText(text).width <= maxW) return text;
    let t = text; while (t.length > 1 && c.measureText(t + "…").width > maxW) t = t.slice(0, -1); return t + "…";
}

// Desenha UMA página do ranking. rows = linhas desta página; ini = índice global da primeira (zebra contínua);
// n/total = número desta página e total; H = altura (fixa em A4 quando há várias páginas); ultima = leva o rodapé.
function desenharPagina(pg: PaginaPdf, rows: PaginaPdf["reps"], mxV: number, mxC: number, idx: number, total: number, H: number, lines: string[], S = 2, logo = false, alinhar = false): HTMLCanvasElement {
    const W = 880, isPos = pg.isPos, showV = pg.showV !== false, rowH = LINHA_H, pitch = PASSO;
    const primeira = idx === 0, ultima = idx === total - 1, top0 = primeira ? TOPO1 : TOPON;
    const sp = (c: any, v: string) => { try { c.letterSpacing = v; } catch (e) { /* sem suporte */ } };
    const cv = document.createElement("canvas"); cv.width = Math.floor(W * S); cv.height = Math.floor(H * S);
    const c = cv.getContext("2d"); if (!c) throw erroImagemGrande(); c.scale(S, S); c.textBaseline = "alphabetic"; c.imageSmoothingEnabled = true; c.imageSmoothingQuality = "high";
    c.fillStyle = "#F4F6F9"; c.fillRect(0, 0, W, H);
    c.textAlign = "left";
    let hy: number;
    if (primeira) {
        c.fillStyle = "#0B2A4A"; c.fillRect(0, 0, W, 292);
        c.fillStyle = "#F2A23A"; c.font = "600 15px " + FONT_B; sp(c, "3px"); c.fillText(fitText(c, pg.eyebrow, W - 104), 52, 62);
        const tw = logo ? W - 104 - LOGO_W - 40 : W - 104;
        let fs = 64; sp(c, "0px"); for (; fs > 36; fs -= 2) { c.font = "700 " + fs + "px " + FONT_H; if (c.measureText(pg.title).width <= tw) break; }
        c.fillStyle = "#fff"; c.fillText(pg.title, 52, 132);
        if (logo) { c.save(); c.translate(W - 52 - LOGO_W, 132 - LOGO_H); c.translate(LOGO_DX, LOGO_DY); c.fillStyle = "#fff"; LOGO_PATHS.forEach(d => c.fill(new Path2D(d))); c.restore(); }
        fs = 22; for (; fs > 13; fs -= 1) { c.font = "400 " + fs + "px " + FONT_B; if (c.measureText(pg.sub).width <= W - 104) break; }
        c.fillStyle = "#C5D3E6"; c.fillText(pg.sub, 52, 169);
        let sx = 52;
        pg.stats.forEach(s => {
            const L = s.l.toUpperCase();
            c.fillStyle = "#9FB3C8"; c.font = "500 12px " + FONT_B; sp(c, "2.4px"); c.fillText(L, sx, 210); const lw = c.measureText(L).width;
            c.fillStyle = "#fff"; c.font = "700 30px " + FONT_H; sp(c, "0px"); c.fillText(s.v, sx, 246); const vw = c.measureText(s.v).width;
            sx += Math.max(lw, vw) + 40;
        });
        hy = 328;
    } else {
        // faixa fina das páginas seguintes: mesmo azul-marinho, filtros e título
        c.fillStyle = "#0B2A4A"; c.fillRect(0, 0, W, 100);
        c.fillStyle = "#F2A23A"; c.font = "600 13px " + FONT_B; sp(c, "3px"); c.fillText(fitText(c, pg.eyebrow, W - 104), 52, 38);
        let fs = 40; sp(c, "0px"); for (; fs > 24; fs -= 2) { c.font = "700 " + fs + "px " + FONT_H; if (c.measureText(pg.title).width <= W - 104) break; }
        c.fillStyle = "#fff"; c.fillText(pg.title, 52, 78);
        hy = 135;
    }
    sp(c, "2px");
    c.font = "600 11px " + FONT_B; c.fillStyle = "#5A6B7E"; c.textAlign = "left"; c.fillText("POS.", 56, hy); c.fillText("REPRESENTANTE", 121, hy);
    c.textAlign = "right";
    if (showV) { c.font = (isPos ? "600" : "700") + " 11px " + FONT_B; c.fillStyle = isPos ? "#5A6B7E" : "#0B2A4A"; c.fillText("VENDIDO (R$)", 724, hy); }
    const cliOn = isPos || !showV;
    c.font = (cliOn ? "700" : "600") + " 11px " + FONT_B; c.fillStyle = cliOn ? "#0B2A4A" : "#5A6B7E"; c.fillText("CLIENTES", 827, hy);
    sp(c, "0px");
    rows.forEach((x, i) => {
        const y0 = top0 + i * pitch, y = alinhar ? Math.round(y0 * S) / S : y0, top = x.pos <= 3; // alinhar: topo da linha no pixel real (bordas e barras nítidas)
        c.save();
        if (top) { c.shadowColor = "rgba(11,42,74,.14)"; c.shadowBlur = 4; c.shadowOffsetY = 1; c.fillStyle = "#fff"; } else c.fillStyle = x.zebra % 2 ? "#EEF1F5" : "#F8FAFC";
        roundRect(c, 40, y, 801, rowH, 9); c.fill(); c.restore();
        c.fillStyle = top ? BADGE[x.pos - 1] : "#DDE3EA"; roundRect(c, 55, y + 7, 37, 27, 6); c.fill();
        c.fillStyle = top ? "#fff" : "#0B2A4A"; c.font = "700 16px " + FONT_H; c.textAlign = "center"; c.fillText(x.pos + "º", 73.5, y + 26);
        c.textAlign = "left"; c.fillStyle = "#0E1B2B"; c.font = "600 17px " + FONT_H; c.fillText(fitText(c, x.r, 470), 121, y + 21);
        const pct = Math.max(.012, isPos ? x.q / mxC : x.v / mxV);
        c.fillStyle = "#E3E8EF"; roundRect(c, 121, y + 31, 409, 4, 2); c.fill();
        c.fillStyle = top ? "#F2A23A" : "#0B2A4A"; roundRect(c, 121, y + 31, Math.max(4, 409 * pct), 4, 2); c.fill();
        c.textAlign = "right";
        if (showV) { c.font = isPos ? "400 16px " + FONT_B : "700 22px " + FONT_H; c.fillStyle = isPos ? "#5F6F80" : "#0E1B2B"; c.fillText(brl(x.v), 722, y + 28); }
        c.font = cliOn ? "700 24px " + FONT_H : "400 16px " + FONT_B; c.fillStyle = cliOn ? "#0E1B2B" : "#5F6F80"; c.fillText(int(x.q), 825, y + 28);
    });
    if (ultima) { c.textAlign = "left"; c.fillStyle = "#6B7A90"; c.font = "400 12px " + FONT_B; lines.forEach((t, i) => c.fillText(t, 52, H - 42 - (lines.length - 1 - i) * 18 + 4)); }
    if (total > 1) { c.textAlign = "center"; c.fillStyle = "#6B7A90"; c.font = "500 12px " + FONT_B; c.fillText("Página " + (idx + 1) + " de " + total, W / 2, H - 16); }
    return cv;
}
function erroImagemGrande(): Error { const e = new Error(msgImagemGrande(areaMaximaImagem())); e.name = "ImagemGrande"; return e; }

// Uma página do PDF: desenha (escala 2) e converte em JPEG para o PDF
function pageImage(pg: PaginaPdf, rows: PaginaPdf["reps"], mxV: number, mxC: number, idx: number, total: number, H: number, lines: string[]): Imagem {
    const W = 880, cv = desenharPagina(pg, rows, mxV, mxC, idx, total, H, lines, 2);
    const b64 = cv.toDataURL("image/jpeg", 0.92).split(",")[1], bin = atob(b64), jpg = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) jpg[i] = bin.charCodeAt(i);
    return { jpg, w: cv.width, h: cv.height, pw: 595.28, ph: Math.round(H * 595.28 / W * 100) / 100 };
}

// Um documento (lista contínua) -> uma ou mais páginas. Cabe numa página: altura dinâmica (como o PDF antigo); senão páginas A4.
function docPages(pg: PaginaPdf): { rows: PaginaPdf["reps"]; mxV: number; mxC: number; lines: string[]; sizes: number[] } {
    const mc = document.createElement("canvas").getContext("2d"); mc.font = "400 12px " + FONT_B;
    const lines = wrapText(mc, pg.foot, 880 - 52 * 2 - 20);
    return { rows: pg.reps, mxV: Math.max(...pg.reps.map(x => x.v)), mxC: Math.max(...pg.reps.map(x => x.q)), lines, sizes: particionar(pg.reps.length, lines.length) };
}

function buildPdf(pages: Imagem[]): Uint8Array {
    const enc = new TextEncoder(), chunks: Uint8Array[] = [], offs: number[] = []; let len = 0;
    const push = (d: string | Uint8Array) => { const u = typeof d === "string" ? enc.encode(d) : d; chunks.push(u); len += u.length; };
    const obj = (n: number, body: string) => { offs[n] = len; push(n + " 0 obj\n" + body + "\nendobj\n"); };
    push("%PDF-1.4\n");
    obj(1, "<< /Type /Catalog /Pages 2 0 R >>");
    obj(2, "<< /Type /Pages /Count " + pages.length + " /Kids [" + pages.map((p, i) => (3 + i * 3) + " 0 R").join(" ") + "] >>");
    pages.forEach((p, i) => {
        const a = 3 + i * 3, cn = "q " + p.pw + " 0 0 " + p.ph + " 0 0 cm /Im0 Do Q";
        obj(a, "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 " + p.pw + " " + p.ph + "] /Resources << /XObject << /Im0 " + (a + 2) + " 0 R >> >> /Contents " + (a + 1) + " 0 R >>");
        obj(a + 1, "<< /Length " + cn.length + " >>\nstream\n" + cn + "\nendstream");
        offs[a + 2] = len; push((a + 2) + " 0 obj\n<< /Type /XObject /Subtype /Image /Width " + p.w + " /Height " + p.h + " /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length " + p.jpg.length + " >>\nstream\n"); push(p.jpg); push("\nendstream\nendobj\n");
    });
    const N = 3 + pages.length * 3, xr = len;
    let x = "xref\n0 " + N + "\n0000000000 65535 f \n";
    for (let i = 1; i < N; i++) x += String(offs[i]).padStart(10, "0") + " 00000 n \n";
    push(x + "trailer\n<< /Size " + N + " /Root 1 0 R >>\nstartxref\n" + xr + "\n%%EOF");
    const out = new Uint8Array(len); let o = 0; chunks.forEach(ch => { out.set(ch, o); o += ch.length; });
    return out;
}

// Carga do visual -> bytes do PDF (progresso opcional: página atual / total)
export async function gerarPdfRanking(carga: CargaRanking, progresso?: (i: number, n: number) => void): Promise<Uint8Array> {
    const docs = expandir(carga);
    if (!docs.length) { throw new Error("Nada para exportar com os filtros escolhidos no painel."); }
    await loadFonts();
    const planos = docs.map(docPages);
    const totalPag = planos.reduce((a, p) => a + p.sizes.length, 0);
    if (totalPag > LIMITE_PAGINAS) { throw new Error("O PDF passaria de " + LIMITE_PAGINAS + " páginas. Filtre Supervisão ou UF e exporte de novo."); }
    const imgs: Imagem[] = [];
    let feitas = 0;
    for (let d = 0; d < docs.length; d++) {
        const pl = planos[d], pg = docs[d]; let ini = 0;
        for (let i = 0; i < pl.sizes.length; i++) {
            if (progresso) progresso(++feitas, totalPag);
            await new Promise(r => setTimeout(r, 0));
            const rows = pl.rows.slice(ini, ini + pl.sizes[i]); ini += pl.sizes[i];
            let H = PAG_H;
            if (pl.sizes.length === 1) H = Math.round(Math.max(946, TOPO1 + (rows.length - 1) * PASSO + LINHA_H + rodapeH(pl.lines.length)));
            imgs.push(pageImage(pg, rows, pl.mxV, pl.mxC, i, pl.sizes.length, H, pl.lines));
        }
    }
    return buildPdf(imgs);
}

// Ranking em UMA imagem (JPEG): a lista inteira, contínua, sem páginas (mesmo desenho do PDF, com a logo no cabeçalho).
// v1.0.9: escala-alvo 3 (depois 2,5 -> 2 -> 1,5 -> 1) conforme os limites do navegador; cada escala é VALIDADA (o canvas gigante precisa ter
// sido desenhado de verdade e o JPEG não pode sair vazio), senão tenta a próxima. Se nem na escala 1 couber, avisa (erro "ImagemGrande").
const perto = (a: Uint8ClampedArray, r: number, g: number, b: number) => Math.abs(a[0] - r) < 10 && Math.abs(a[1] - g) < 10 && Math.abs(a[2] - b) < 10 && a[3] === 255;
function desenhou(cv: HTMLCanvasElement): boolean {
    try {
        const c = cv.getContext("2d"); if (!c) return false;
        return perto(c.getImageData(3, 3, 1, 1).data, 0x0B, 0x2A, 0x4A) && perto(c.getImageData(3, cv.height - 4, 1, 1).data, 0xF4, 0xF6, 0xF9);   // cabeçalho azul-marinho e fundo no pé
    } catch (e) { return false; }
}
export interface OpcoesImagem { qualidade?: number; escalas?: number[]; escalaUsada?: (s: number) => void; }
export async function gerarJpegRanking(carga: CargaRanking, progresso?: (i: number, n: number) => void, op: OpcoesImagem = {}): Promise<Uint8Array> {
    const docs = expandir(carga);
    if (!docs.length) { throw new Error("Nada para exportar com os filtros escolhidos no painel."); }
    await loadFonts();
    const pg = docs[0], pl = docPages(pg);
    const escalas = op.escalas || escalasImagem(pg.reps.length, pl.lines.length, areaMaximaImagem());
    if (!escalas.length) { throw erroImagemGrande(); }
    if (progresso) progresso(1, 1);
    const H = alturaImagem(pg.reps.length, pl.lines.length);
    for (const S of escalas) {
        await new Promise(r => setTimeout(r, 0));
        let cv: HTMLCanvasElement | null = null;
        try {
            cv = desenharPagina(pg, pl.rows, pl.mxV, pl.mxC, 0, 1, H, pl.lines, S, true, true);
            if (!desenhou(cv)) { continue; }
            const blob: Blob | null = await new Promise(r => { try { cv.toBlob(b => r(b), "image/jpeg", op.qualidade || IMG_QUALIDADE); } catch (e) { r(null); } });
            if (!blob || blob.size < 1000) { continue; }
            if (op.escalaUsada) op.escalaUsada(S);
            return new Uint8Array(await blob.arrayBuffer());
        } catch (e) { /* tenta a próxima escala */ }
        finally { if (cv) { cv.width = 0; cv.height = 0; } }
    }
    throw erroImagemGrande();
}
