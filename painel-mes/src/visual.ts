"use strict";
import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import "./../style/visual.less";
import { VisualSettings } from "./settings";
import { gerarXlsxMes, paraBase64, OpcoesMes } from "./xlsxMes";
import { Filtros, FiltroBasico, el, botao, dataBR } from "./filtros";

import IVisual = powerbi.extensibility.visual.IVisual;
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;

const NS = "http://www.w3.org/2000/svg";
const FAMILIAS = [
    { sigla: "FB", cor: "#2B9D8F", texto: "#FFFFFF" },
    { sigla: "FC", cor: "#F4B400", texto: "#FFFFFF" },
    { sigla: "CT", cor: "#BDBDBD", texto: "#FFFFFF" },
    { sigla: "PT", cor: "#FFFFFF", texto: "#1B1B1B" },
    { sigla: "AP", cor: "#F25C05", texto: "#FFFFFF" }
];
const MEDIDORES = [
    { rotulo: "FORRO E ACESS. BR", ind: 0 },
    { rotulo: "FORRO E ACESS. COLOR", ind: 1 },
    { rotulo: "CANTINHO", ind: 2 },
    { rotulo: "GERAL", ind: -1 },
    { rotulo: "PORTA PVC", ind: 3 },
    { rotulo: "ACESS. PARA PISO", ind: 4 }
];

interface Linha { nome: string; uf: string; sup: string; mes: number | null; fam: (number | null)[]; vend: number; }
interface Cab { vendMes: number; cliMes: number; repsMes: number; reps: number[]; repsPainel: number; }

function sv(tag: string, atributos: Record<string, string>): SVGElement {
    const e = document.createElementNS(NS, tag);
    Object.keys(atributos).forEach(k => e.setAttribute(k, atributos[k]));
    return e as SVGElement;
}
function milhar(n: number): string {
    return String(Math.round(n)).replace(/\B(?=(\d{3})+(?!\d))/g, ".");
}
function num(s: string): number | null {
    if (s === undefined || s === "") { return null; }
    const n = Number(s);
    return isNaN(n) ? null : n;
}

export class Visual implements IVisual {
    private host: IVisualHost;
    private raiz: HTMLElement;
    private fmtService = new FormattingSettingsService();
    private settings: VisualSettings = new VisualSettings();
    private linhas: Linha[] = [];
    private cab: Cab | null = null;
    private temDados = false;
    private filtros: Filtros;

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.raiz = el("div", "pd");
        options.element.appendChild(this.raiz);
        this.filtros = new Filtros(this.host, this.raiz, () => this.desenhar(), () => this.settings.filtros, () => this.periodo(), "Painel Mês (visual)");
        this.raiz.addEventListener("keydown", (ev: KeyboardEvent) => {
            if (ev.key !== "Escape") { return; }
            if (this.filtros.popAberto) { this.filtros.fechar(); ev.stopPropagation(); }
            else if (this.menuAberto) { this.menuAberto = false; this.desenhar(); }
        });
    }

    public update(options: VisualUpdateOptions): void {
        const dv = options.dataViews && options.dataViews[0];
        this.settings = this.fmtService.populateFormattingSettingsModel(VisualSettings, dv);
        this.lerDados(dv);
        if (Array.isArray(options.jsonFilters)) { this.filtros.lerFiltros(options.jsonFilters as unknown as FiltroBasico[]); }
        if (dv && dv.metadata && dv.metadata.segment) { this.host.fetchMoreData(true); }
        this.desenhar();
    }

    private lerDados(dv: powerbi.DataView | undefined): void {
        this.linhas = []; this.cab = null; this.temDados = false;
        const t = dv && dv.table;
        if (!t) { return; }
        let idx = t.columns.findIndex(c => c.roles && c.roles["dados"]);
        if (idx < 0) { idx = t.columns.length - 1; }
        for (const r of t.rows) {
            const bruto = r[idx];
            if (typeof bruto !== "string" || !bruto) { continue; }
            this.temDados = true;
            const partes = bruto.split("~");
            const c = partes[0].split("|");
            if (partes[1]) {
                // cabeçalho: 9 números (card e medidores) | dataRef | anos | meses | dias | supervisões | representantes | ufs | famílias | materiais
                const hs = partes[1].split("|");
                const h = hs.slice(0, 9).map(x => Number(x) || 0);
                if (h.length >= 9) { this.cab = { vendMes: h[0], cliMes: h[1], repsMes: h[2], reps: h.slice(3, 8), repsPainel: h[8] }; }
                this.filtros.lerCabecalho(hs, 9);
            }
            if (c.length < 10) { continue; }
            this.linhas.push({ nome: c[0], uf: c[1], sup: c[2], mes: num(c[3]), fam: [num(c[4]), num(c[5]), num(c[6]), num(c[7]), num(c[8])], vend: num(c[9]) || 0 });
        }
        this.linhas.sort((a, b) => (b.mes || 0) - (a.mes || 0) || b.vend - a.vend || a.nome.localeCompare(b.nome, "pt-BR"));
    }

    // ---------- medidor ----------
    private ponto(cx: number, cy: number, raio: number, f: number): [number, number] {
        const a = Math.PI * (1 - f);
        return [cx + raio * Math.cos(a), cy - raio * Math.sin(a)];
    }
    private setor(cx: number, cy: number, R: number, r: number, f0: number, f1: number): string {
        const [x0, y0] = this.ponto(cx, cy, R, f0), [x1, y1] = this.ponto(cx, cy, R, f1);
        const [x2, y2] = this.ponto(cx, cy, r, f1), [x3, y3] = this.ponto(cx, cy, r, f0);
        return `M${x0.toFixed(2)} ${y0.toFixed(2)}A${R} ${R} 0 0 1 ${x1.toFixed(2)} ${y1.toFixed(2)}L${x2.toFixed(2)} ${y2.toFixed(2)}A${r} ${r} 0 0 0 ${x3.toFixed(2)} ${y3.toFixed(2)}Z`;
    }
    private desenharMedidor(valor: number, max: number, rotulo: string, imagem: boolean): SVGElement {
        const cx = 95, cy = 72, R = 64, r = 40;
        const amarelo = Math.min(Math.max(this.settings.medidores.inicioAmarelo.value, 1), 98) / 100;
        const verde = Math.min(Math.max(this.settings.medidores.inicioVerde.value, amarelo * 100 + 1), 99) / 100;
        const svg = sv("svg", { viewBox: imagem ? "0 0 190 106" : "0 0 190 84", role: "img", "aria-label": `${rotulo}: ${valor} de ${max}` });
        if (imagem) {
            svg.setAttribute("xmlns", NS); svg.setAttribute("width", "200"); svg.setAttribute("height", "112");
            svg.appendChild(sv("rect", { x: "0", y: "0", width: "190", height: "106", fill: "#FFFFFF" }));
        }
        const trechos: [number, number, string][] = [[0, amarelo, "#C23B3B"], [amarelo, verde, "#E9E57F"], [verde, 1, "#47A863"]];
        trechos.forEach(([a, b, cor]) => svg.appendChild(sv("path", { d: this.setor(cx, cy, R, r, a, b), fill: cor, stroke: "#FFFFFF", "stroke-width": "0.6" })));
        const f = max > 0 ? Math.min(Math.max(valor / max, 0), 1) : 0;
        const a = Math.PI * (1 - f);
        const tx = cx + 61 * Math.cos(a), ty = cy - 61 * Math.sin(a);
        const px = 3.4 * Math.sin(a), py = 3.4 * Math.cos(a);
        svg.appendChild(sv("polygon", { points: `${(cx + px).toFixed(2)},${(cy + py).toFixed(2)} ${tx.toFixed(2)},${ty.toFixed(2)} ${(cx - px).toFixed(2)},${(cy - py).toFixed(2)}`, fill: "#707070" }));
        svg.appendChild(sv("circle", { cx: String(cx), cy: String(cy), r: "19", fill: "#FFFFFF" }));
        const txt = (x: number, y: number, ancora: string, tamanho: number, peso: number, cor: string, s: string, cls: string): void => {
            const t = sv("text", { x: String(x), y: String(y), "text-anchor": ancora, class: cls, "font-family": "Arial, sans-serif", "font-size": String(tamanho), "font-weight": String(peso), fill: cor });
            t.textContent = s; svg.appendChild(t);
        };
        txt(cx, cy - 5, "middle", 17, 700, "#1B1B1B", String(valor), "md-v");
        txt(27, 82, "end", 8, 400, "#666666", "0", "md-m");
        txt(163, 82, "start", 8, 400, "#666666", String(max), "md-m");
        if (imagem) { txt(cx, 100, "middle", 11, 700, "#1B1B1B", rotulo, "md-i"); }
        return svg;
    }
    private medidor(valor: number, max: number, rotulo: string): HTMLElement {
        const caixa = el("div", "md");
        caixa.appendChild(this.desenharMedidor(valor, max, rotulo, false));
        caixa.appendChild(el("div", "md-r", rotulo));
        return caixa;
    }

    // ---------- tela ----------
    private desenhar(): void {
        const raiz = this.raiz;
        while (raiz.firstChild) { raiz.removeChild(raiz.firstChild); }
        raiz.style.background = this.settings.fundo.mostrar.value ? this.settings.fundo.cor.value.value : "transparent";
        if (!this.cab) {
            // sem linhas: com filtro do botão ativo, deixa o botão para poder limpar
            if (this.filtros.ativos() || this.temDados) { raiz.appendChild(this.acoes()); }
            raiz.appendChild(el("div", "pd-vazio", this.filtros.ativos() || this.temDados ? "Sem representantes para os filtros escolhidos." : "Arraste o campo Representante e a medida Painel Mês (visual) para o visual."));
            if (this.filtros.popAberto) { raiz.appendChild(this.filtros.painel()); }
            return;
        }
        const c = this.cab;
        const topo = el("div", "pd-topo");
        const esq = el("div", "pd-kpis");
        esq.appendChild(el("div", "kpi kpi-vz"));
        esq.appendChild(this.kpi("VENDIDO NO MÊS", milhar(c.vendMes), c.cliMes));
        topo.appendChild(esq);
        const grade = el("div", "pd-medidores");
        MEDIDORES.forEach(m => grade.appendChild(this.medidor(m.ind < 0 ? c.repsMes : c.reps[m.ind], c.repsPainel, m.rotulo)));
        topo.appendChild(grade);
        raiz.appendChild(topo);
        raiz.appendChild(this.acoes());
        if (this.settings.titulo.mostrar.value) { raiz.appendChild(el("div", "pd-titulo", this.settings.titulo.texto.value)); }
        raiz.appendChild(this.tabela());
        if (this.filtros.popAberto) { raiz.appendChild(this.filtros.painel()); }
    }

    // botões Filtros e Exportar, no canto superior direito (onde já ficava o Exportar)
    private acoes(): HTMLElement {
        const a = el("div", "pd-acoes");
        if (this.settings.filtros.mostrarBotao.value) { a.appendChild(this.filtros.botao("ex-btn", () => { this.menuAberto = false; })); }
        if (this.settings.exportacao.mostrarBotao.value) { a.appendChild(this.menuExportar()); }
        return a;
    }

    private periodo(): string {
        const r = this.filtros.dataRef;
        return r ? `Período: ${dataBR(r.slice(0, 8) + "01")} a ${dataBR(r)}` : "";
    }

    private kpi(rotulo: string, valor: string, clientes: number): HTMLElement {
        const b = el("div", "kpi");
        b.appendChild(el("div", "kpi-r", rotulo));
        b.appendChild(el("div", "kpi-v", valor));
        b.appendChild(el("div", "kpi-c", `(${milhar(clientes)}) CLIENTES`));
        return b;
    }

    private tabela(): HTMLElement {
        const wrap = el("div", "pd-tab");
        const t = el("table");
        const larguras = ["31%", "7%", "14%", "10%", "7.6%", "7.6%", "7.6%", "7.6%", "7.6%"];
        const cg = el("colgroup");
        larguras.forEach(w => { const col = el("col"); col.style.width = w; cg.appendChild(col); });
        t.appendChild(cg);
        const th = el("thead");
        const f1 = el("tr", "faixa");
        const a = el("th", "", "POSITIVAÇÃO"); a.setAttribute("colspan", "4");
        const b = el("th", "", "MIX (MÊS)"); b.setAttribute("colspan", "5");
        f1.appendChild(a); f1.appendChild(b); th.appendChild(f1);
        const f2 = el("tr", "cols");
        ["REPRESENTANTES", "UF", "SUP"].forEach(x => f2.appendChild(el("th", "", x)));
        ["CLI MÊS"].forEach(x => { const h = el("th", "cli"); x.split(" ").forEach((p, i) => { if (i) { h.appendChild(document.createElement("br")); } h.appendChild(document.createTextNode(p)); }); f2.appendChild(h); });
        FAMILIAS.forEach(fm => { const h = el("th", "fam", fm.sigla); h.style.background = fm.cor; h.style.color = fm.texto; f2.appendChild(h); });
        th.appendChild(f2); t.appendChild(th);
        const tb = el("tbody");
        this.linhas.forEach(l => {
            const tr = el("tr");
            tr.appendChild(el("td", "nm", l.nome));
            tr.appendChild(el("td", "uf", l.uf));
            tr.appendChild(el("td", "sp", l.sup));
            tr.appendChild(el("td", l.mes === null ? "n x vz" : "n", l.mes === null ? "X" : String(l.mes)));
            l.fam.forEach(v => tr.appendChild(el("td", "mx" + (v === null ? " vz" : ""), v === null ? "" : String(v))));
            tb.appendChild(tr);
        });
        t.appendChild(tb);
        wrap.appendChild(t);
        return wrap;
    }

    // ---------- exportar ----------
    private escolhas: OpcoesMes = { cards: true, medidores: true, tabela: true, dados: false };
    private menuAberto = false;
    private aviso = "";
    private avisoErro = false;

    private menuExportar(): HTMLElement {
        const cx = el("div", "ex");
        const btn = el("button", "ex-btn", "Exportar ▾") as HTMLButtonElement;
        btn.type = "button";
        btn.setAttribute("aria-haspopup", "true");
        btn.setAttribute("aria-expanded", String(this.menuAberto));
        btn.addEventListener("click", () => { this.menuAberto = !this.menuAberto; if (this.menuAberto) { this.filtros.popAberto = false; } this.desenhar(); });
        cx.appendChild(btn);
        if (this.menuAberto) {
            const m = el("div", "ex-menu");
            m.appendChild(el("div", "ex-tit", "Incluir no Excel"));
            const itens: [keyof OpcoesMes, string][] = [["cards", "Card (vendido do mês)"], ["medidores", "Medidores (como imagem)"], ["tabela", "Tabela (células editáveis)"], ["dados", "Aba com dados brutos"]];
            itens.forEach(([k, rot]) => {
                const lb = el("label", "ex-op");
                const cb = document.createElement("input");
                cb.type = "checkbox"; cb.checked = this.escolhas[k];
                cb.addEventListener("change", () => { this.escolhas[k] = cb.checked; });
                lb.appendChild(cb); lb.appendChild(document.createTextNode(" " + rot));
                m.appendChild(lb);
            });
            m.appendChild(botao("ex-go", "Baixar Excel", () => { void this.exportarExcel(); }));
            cx.appendChild(m);
        }
        if (this.aviso) { cx.appendChild(el("div", "ex-msg" + (this.avisoErro ? " erro" : ""), this.aviso)); }
        return cx;
    }

    private mostrar(t: string, erro = false): void { this.aviso = t; this.avisoErro = erro; this.desenhar(); }

    private agora(): string {
        const d = new Date();
        const p = (n: number): string => (n < 10 ? "0" : "") + n;
        return `${p(d.getDate())}/${p(d.getMonth() + 1)}/${d.getFullYear()} ${p(d.getHours())}:${p(d.getMinutes())}`;
    }

    private svgsMedidores(): { svg: string; w: number; h: number }[] {
        const c = this.cab as Cab;
        const ser = new XMLSerializer();
        return MEDIDORES.map(m => ({
            svg: ser.serializeToString(this.desenharMedidor(m.ind < 0 ? c.repsMes : c.reps[m.ind], c.repsPainel, m.rotulo, true)),
            w: 200, h: 112
        }));
    }

    // Card VENDIDO NO MÊS como imagem. O marcador /*FONTE*/ é trocado pela fonte embutida na página de download.
    private svgCard(): { svg: string; w: number; h: number } {
        const c = this.cab as Cab;
        const fam = "AntonPD, Impact, 'Arial Narrow', sans-serif";
        const svg = `<svg xmlns="${NS}" width="200" height="100" viewBox="0 0 200 100"><style>/*FONTE*/</style><rect width="200" height="100" fill="#FFFFFF"/>` +
            `<text x="2" y="16" font-family="${fam}" font-size="14" fill="#003366">VENDIDO NO MÊS</text>` +
            `<text x="196" y="58" text-anchor="end" font-family="${fam}" font-size="38" fill="#003366">${milhar(c.vendMes)}</text>` +
            `<text x="196" y="76" text-anchor="end" font-family="Arial, sans-serif" font-size="10" fill="#8A8F98">(${milhar(c.cliMes)}) CLIENTES</text></svg>`;
        return { svg, w: 200, h: 100 };
    }

    private svgParaPng(svg: string, w: number, h: number): Promise<Uint8Array> {
        return new Promise((resolve, reject) => {
            const img = new Image();
            img.onload = () => {
                const cv = document.createElement("canvas"); cv.width = w * 2; cv.height = h * 2;
                const ctx = cv.getContext("2d");
                if (!ctx) { reject(new Error("canvas indisponível")); return; }
                ctx.fillStyle = "#FFFFFF"; ctx.fillRect(0, 0, cv.width, cv.height);
                ctx.drawImage(img, 0, 0, cv.width, cv.height);
                cv.toBlob(async b => { if (b) { resolve(new Uint8Array(await b.arrayBuffer())); } else { reject(new Error("falha ao gerar a imagem")); } }, "image/png");
            };
            img.onerror = () => reject(new Error("falha ao desenhar o medidor"));
            img.src = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(svg);
        });
    }

    private async exportarExcel(): Promise<void> {
        const o = this.escolhas;
        if (!this.cab || !this.linhas.length) { this.mostrar("Sem dados para exportar.", true); return; }
        if (!o.cards && !o.medidores && !o.tabela && !o.dados) { this.mostrar("Marque pelo menos um item para exportar.", true); return; }
        this.menuAberto = false;
        const ex = this.settings.exportacao;
        const nome = (ex.nomeArquivo.value || "painel_mes").replace(/[^\w\-]+/g, "_") + ".xlsx";
        const url = (ex.urlPagina.value || "").trim();
        try {
            if (ex.usarPagina.value && /^https:\/\//i.test(url)) { await this.abrirPaginaDownload(url, nome); return; }
            this.mostrar("Gerando Excel…");
            const medidores = [];
            for (const s of o.medidores ? this.svgsMedidores() : []) { medidores.push({ png: await this.svgParaPng(s.svg, s.w, s.h), w: s.w, h: s.h }); }
            let card = null;
            if (o.cards) { const sc = this.svgCard(); card = { png: await this.svgParaPng(sc.svg.replace("/*FONTE*/", ""), sc.w, sc.h), w: sc.w, h: sc.h }; }
            const bytes = gerarXlsxMes({ titulo: this.settings.titulo.texto.value || "Painel Mês", geradoEm: this.agora(), periodo: this.periodo(), filtros: this.filtros.resumo(), linhas: this.linhas, card, medidores, opcoes: o });
            let permitido = false;
            try { permitido = (await this.host.downloadService.exportStatus()) === powerbi.PrivilegeStatus.Allowed; } catch { permitido = false; }
            if (permitido) {
                const r = await this.host.downloadService.exportVisualsContentExtended(paraBase64(bytes), nome, "base64", "Exportação do Painel Mês para Excel");
                this.mostrar(r && (r as { downloadCompleted?: boolean }).downloadCompleted === false ? "Download cancelado." : "Excel exportado: " + nome);
            } else {
                const a = document.createElement("a");
                a.href = URL.createObjectURL(new Blob([bytes as BlobPart], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
                a.download = nome; this.raiz.appendChild(a); a.click(); this.raiz.removeChild(a);
                this.mostrar("Download direto (exportação do Power BI não liberada pelo administrador).");
            }
        } catch (e) { this.mostrar("Erro ao exportar: " + (e instanceof Error ? e.message : String(e)), true); }
    }

    // Abre a página de download levando só os dados (texto pequeno) no fragmento (#).
    private async abrirPaginaDownload(url: string, nome: string): Promise<void> {
        const o = this.escolhas;
        const carga = {
            tipo: "painelMes", titulo: this.settings.titulo.texto.value || "Painel Mês", geradoEm: this.agora(),
            periodo: this.periodo(), filtros: this.filtros.resumo(), linhas: this.linhas, opcoes: o, card: o.cards ? this.svgCard() : null, svgs: o.medidores ? this.svgsMedidores() : []
        };
        let dados: Uint8Array = new TextEncoder().encode(JSON.stringify(carga));
        let comp = "0";
        try {
            const CS = (window as unknown as { CompressionStream?: new (f: string) => { readable: ReadableStream; writable: WritableStream } }).CompressionStream;
            if (CS) {
                const cs = new CS("deflate-raw");
                const wr = cs.writable.getWriter();
                void wr.write(dados as BufferSource); void wr.close();
                dados = new Uint8Array(await new Response(cs.readable).arrayBuffer()); comp = "1";
            }
        } catch { comp = "0"; }
        let bin = "";
        for (let i = 0; i < dados.length; i += 0x8000) { bin += String.fromCharCode.apply(null, Array.from(dados.subarray(i, i + 0x8000))); }
        const b64 = btoa(bin).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
        const params = new URLSearchParams({ v: "2", n: nome, c: comp, j: b64 });
        const alvo = url.replace(/#.*$/, "") + "#" + params.toString();
        this.host.launchUrl(alvo);
        this.mostrar("Página de download aberta: " + nome + " (link de " + Math.round(alvo.length / 1024) + " mil caracteres)");
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.fmtService.buildFormattingModel(this.settings);
    }
}
