// Roda na página de download: recebe os dados (pequenos) pelo link, monta o .xlsx e baixa.
import { gerarXlsx, EntradaXlsx } from "../visual/src/xlsx";
import { gerarXlsxDia, EntradaDia } from "../painel-dia/src/xlsxDia";
import { gerarXlsxMes, EntradaMes, LinhaMes } from "../painel-mes/src/xlsxMes";
import { gerarXlsxPedido, EntradaPedido, LinhaPedido } from "../pedido-diario/src/xlsxPedido";
import { gerarXlsxReps, EntradaReps } from "../reps-sem-pedidos/src/xlsxReps";
import { ANTON_B64 } from "./fonte";

interface CargaDia extends Omit<EntradaDia, "imagens"> { tipo: "painelDia"; svgs: { svg: string; w: number; h: number }[]; }
// Painel Mês v1.1 / Painel Dia v1.2 mandam também periodo/filtros (rodapé do Excel); versões antigas não mandam
interface CargaMes { tipo: "painelMes"; titulo: string; geradoEm: string; periodo?: string; filtros?: string; linhas: LinhaMes[]; opcoes: EntradaMes["opcoes"]; card: { svg: string; w: number; h: number } | null; svgs: { svg: string; w: number; h: number }[]; }
// Pedido Diário: v2 (oficial) manda periodo/filtros e só tabela/dados; v1 mandava card/svgs (ignorados: o Excel sai sem imagens/topo)
interface CargaPedido { tipo: "pedidoDiario"; versao?: number; titulo: string; geradoEm: string; periodo?: string; filtros?: string; dias: number[]; linhas: LinhaPedido[]; opcoes: EntradaPedido["opcoes"]; card?: unknown; svgs?: { svg: string; w: number; h: number }[]; }
interface CargaReps extends EntradaReps { tipo: "repsSemPedidos"; svgs?: undefined; }
interface Carga extends Omit<EntradaXlsx, "imagens"> { svgs: { svg: string; w: number; h: number }[]; }

const $ = (s: string): HTMLElement => document.querySelector(s) as HTMLElement;
const erro = (m: string): void => {
    $("#st").textContent = "!"; $("#st").classList.add("erro");
    $("#tit").textContent = "Não foi possível preparar o arquivo"; $("#msg").textContent = m;
};
function b64u(s: string): Uint8Array {
    const b = s.replace(/-/g, "+").replace(/_/g, "/");
    const bin = atob(b.padEnd(Math.ceil(b.length / 4) * 4, "="));
    return Uint8Array.from(bin, c => c.charCodeAt(0));
}
async function inflar(b: Uint8Array): Promise<Uint8Array> {
    const ds = new DecompressionStream("deflate-raw");
    const w = ds.writable.getWriter(); void w.write(b as BufferSource); void w.close();
    return new Uint8Array(await new Response(ds.readable).arrayBuffer());
}
function svgParaPng(svg: string, w: number, h: number): Promise<Uint8Array> {
    return new Promise((resolve, reject) => {
        const img = new Image();
        img.onload = () => {
            const c = document.createElement("canvas"); c.width = w * 2; c.height = h * 2;
            const ctx = c.getContext("2d");
            if (!ctx) { reject(new Error("canvas indisponível")); return; }
            ctx.fillStyle = "#FFFFFF"; ctx.fillRect(0, 0, c.width, c.height);
            ctx.drawImage(img, 0, 0, c.width, c.height);
            c.toBlob(async b => { if (b) { resolve(new Uint8Array(await b.arrayBuffer())); } else { reject(new Error("falha ao gerar a imagem")); } }, "image/png");
        };
        img.onerror = () => reject(new Error("falha ao desenhar o gráfico"));
        const fonte = `@font-face{font-family:AntonPD;src:url(data:font/woff;base64,${ANTON_B64}) format("woff");}`;
        img.src = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(svg.replace("/*FONTE*/", fonte));
    });
}
function entregar(bytes: Uint8Array, nome: string, mime: string): void {
    const a = $("#bt") as HTMLAnchorElement;
    a.href = URL.createObjectURL(new Blob([bytes as BlobPart], { type: mime }));
    a.download = nome;
    $("#st").textContent = "✓"; $("#tit").textContent = "Arquivo pronto";
    $("#msg").textContent = "Se o download não começar, use o botão abaixo.";
    $("#nome").textContent = nome;
    $("#tam").textContent = bytes.length < 1048576 ? (bytes.length / 1024).toFixed(1) + " KB" : (bytes.length / 1048576).toFixed(1) + " MB";
    ["#arq", "#bt", "#nota"].forEach(s => $(s).classList.remove("oculto"));
    history.replaceState(null, "", location.pathname);
    a.click();
}

(async () => {
    const p = new URLSearchParams(location.hash.slice(1));
    const nome = p.get("n") || "painel.xlsx";
    const mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    try {
        const j = p.get("j"), d = p.get("d");
        if (j) {
            const bruto = p.get("c") === "1" ? await inflar(b64u(j)) : b64u(j);
            const carga = JSON.parse(new TextDecoder().decode(bruto)) as Carga | CargaDia | CargaMes | CargaPedido | CargaReps;
            const png = async (g: { svg: string; w: number; h: number }) => ({ png: await svgParaPng(g.svg, g.w, g.h), w: g.w, h: g.h });
            const imagens = [];
            const semImagens = (carga as CargaPedido).tipo === "pedidoDiario" || (carga as CargaReps).tipo === "repsSemPedidos";
            for (const g of semImagens ? [] : carga.svgs || []) { imagens.push(await png(g)); }
            if ((carga as CargaReps).tipo === "repsSemPedidos") {
                const m = carga as CargaReps;
                entregar(gerarXlsxReps({ titulo: m.titulo, faixa: m.faixa, geradoEm: m.geradoEm, periodo: m.periodo, filtros: m.filtros, linhas: m.linhas, opcoes: m.opcoes }), nome, mime);
            } else if ((carga as CargaMes).tipo === "painelMes") {
                const m = carga as CargaMes;
                const card = m.card ? await png(m.card) : null;
                entregar(gerarXlsxMes({ titulo: m.titulo, geradoEm: m.geradoEm, periodo: m.periodo || "", filtros: m.filtros || "", linhas: m.linhas, opcoes: m.opcoes, card, medidores: imagens }), nome, mime);
            } else if ((carga as CargaPedido).tipo === "pedidoDiario") {
                const m = carga as CargaPedido;
                const op = m.opcoes || { tabela: true, dados: false };
                entregar(gerarXlsxPedido({ titulo: m.titulo, geradoEm: m.geradoEm, periodo: m.periodo || "", filtros: m.filtros || "", dias: m.dias || [], linhas: m.linhas || [], opcoes: { tabela: !!op.tabela, dados: !!op.dados } }), nome, mime);
            } else if ((carga as CargaDia).tipo === "painelDia") { entregar(gerarXlsxDia({ ...(carga as CargaDia), imagens }), nome, mime); }
            else { entregar(gerarXlsx({ ...(carga as Carga), imagens }), nome, mime); }
        } else if (d) {
            const b = b64u(d);
            entregar(p.get("c") === "1" ? await inflar(b) : b, nome, p.get("t") || mime);
        } else { erro("Nenhum arquivo recebido (página v2.1; o endereço chegou com " + location.hash.length + " caracteres depois do #). Abra esta página pelo botão Exportar do painel."); }
    } catch (e) { erro("O arquivo chegou incompleto ou passou do limite do endereço. " + (e instanceof Error ? e.message : "")); }
})();
