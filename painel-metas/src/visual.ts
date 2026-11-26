"use strict";
// Painel Metas (NovaForma) v1.0: réplica da aba PAINEL METAS (KPIs, Painel de Vendas - Mês por família, Vendido dia x dia),
// bloco total e um bloco por UF do cliente; botão Filtros (filtra a página) e Exportar (Excel igual à aba).
import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import "./../style/visual.less";
import { VisualSettings } from "./settings";
import { Filtros, FiltroBasico, el, botao, dataBR } from "./filtros";
import { lerCarga, textoAtualizacao } from "./dados";
import { gerarXlsxMetas, paraBase64, BlocoMetas, LinhaMetas, EntradaMetas, N } from "./xlsxMetas";

import IVisual = powerbi.extensibility.visual.IVisual;
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;

const MEDIDA = "Painel Metas (visual)";
const fmtInt = new Intl.NumberFormat("pt-BR", { maximumFractionDigits: 0 });
const fmt1 = new Intl.NumberFormat("pt-BR", { minimumFractionDigits: 1, maximumFractionDigits: 1 });
const fmt2 = new Intl.NumberFormat("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const fmtPct = new Intl.NumberFormat("pt-BR", { style: "percent", minimumFractionDigits: 1, maximumFractionDigits: 1 });
function f(v: N | undefined, tipo: "int" | "pct" | "dec1" | "dec2" = "int"): string {
    if (v === null || v === undefined || isNaN(v)) { return ""; }
    if (tipo === "pct") { return fmtPct.format(v); }
    if (tipo === "dec1") { return fmt1.format(v); }
    if (tipo === "dec2") { return fmt2.format(v); }
    return fmtInt.format(v);
}

// colunas do Painel de Vendas - Mês (cores da aba: t = verde-petróleo, n = azul-marinho, b = branco)
const COLS_TAB: { rot: string; cab: "b" | "t" | "n"; cel: string; tipo?: "pct" }[] = [
    { rot: "Objetivo", cab: "b", cel: "obj" },
    { rot: "Vendido mês", cab: "t", cel: "tl" },
    { rot: "Vendido Ontem", cab: "t", cel: "tl" },
    { rot: "Tendência", cab: "b", cel: "" },
    { rot: "Faturado", cab: "n", cel: "az" },
    { rot: "Obj x Fat", cab: "n", cel: "pc", tipo: "pct" },
    { rot: "Carteira", cab: "b", cel: "" },
    { rot: "Faturado + Carteira", cab: "b", cel: "" },
    { rot: "Cliente Mês", cab: "t", cel: "tl" },
    { rot: "Clientes Ontem", cab: "t", cel: "tl" },
    { rot: "Clientes Faturados", cab: "n", cel: "az" },
    { rot: "Clientes Carteira", cab: "b", cel: "" }
];

export class Visual implements IVisual {
    private host: IVisualHost;
    private raiz: HTMLElement;
    private fmtService = new FormattingSettingsService();
    private settings: VisualSettings = new VisualSettings();
    private filtros: Filtros;
    private temDados = false;
    private blocos: BlocoMetas[] = [];
    private atualizacao = "";
    private rolagem = 0;

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.raiz = el("div", "pm");
        options.element.appendChild(this.raiz);
        this.filtros = new Filtros(this.host, this.raiz, () => this.desenhar(), () => this.settings.filtros, () => this.periodo(), MEDIDA);
        this.raiz.addEventListener("keydown", (ev: KeyboardEvent) => {
            if (ev.key !== "Escape") { return; }
            if (this.filtros.popAberto) { this.filtros.fechar(); ev.stopPropagation(); }
            else if (this.menuAberto) { this.menuAberto = false; this.limparAviso(); this.desenhar(); }
        });
    }

    public update(options: VisualUpdateOptions): void {
        const dv = options.dataViews && options.dataViews[0];
        this.settings = this.fmtService.populateFormattingSettingsModel(VisualSettings, dv);
        this.lerDados(dv);
        if (Array.isArray(options.jsonFilters)) { this.filtros.lerFiltros(options.jsonFilters as unknown as FiltroBasico[]); }
        this.desenhar();
    }

    private lerDados(dv: powerbi.DataView | undefined): void {
        this.temDados = false; this.blocos = [];
        const t = dv && dv.table;
        if (!t || !t.rows.length) { return; }
        let idx = t.columns.findIndex(c => c.roles && c.roles["dados"]);
        if (idx < 0) { idx = t.columns.length - 1; }
        const texto = t.rows.map(r => r[idx]).filter(v => typeof v === "string" && v).join("\n");
        if (!texto) { return; }
        const c = lerCarga(texto);
        if (c.cabecalho.length) {
            this.atualizacao = textoAtualizacao(c.cabecalho[0]);
            this.filtros.lerCabecalho(c.cabecalho, 1);
        }
        this.blocos = c.blocos;
        this.temDados = c.blocos.length > 0;
    }

    private periodo(): string {
        const r = this.filtros.dataRef;
        return r ? `Período: ${dataBR(r.slice(0, 8) + "01")} a ${dataBR(r)}` : "";
    }

    // ---------- tela ----------
    private desenhar(): void {
        const raiz = this.raiz;
        const corpoAnt = raiz.querySelector(".pm-corpo") as HTMLElement | null;
        if (corpoAnt) { this.rolagem = corpoAnt.scrollTop; }
        while (raiz.firstChild) { raiz.removeChild(raiz.firstChild); }
        raiz.style.background = this.settings.fundo.mostrar.value ? this.settings.fundo.cor.value.value : "transparent";

        const topo = el("div", "pm-topo");
        if (this.settings.titulo.mostrar.value) { topo.appendChild(el("h2", "pm-titulo", this.settings.titulo.texto.value)); }
        else { topo.appendChild(el("div", "pm-espaco")); }
        if (this.atualizacao && this.temDados) { topo.appendChild(el("div", "pm-atu", this.atualizacao)); }
        const acoes = el("div", "pm-acoes");
        if (this.settings.filtros.mostrarBotao.value) { acoes.appendChild(this.filtros.botao("pm-btn", () => { this.menuAberto = false; this.limparAviso(); })); }
        if (this.settings.exportacao.mostrarBotao.value) { acoes.appendChild(this.menuExportar()); }
        topo.appendChild(acoes);
        raiz.appendChild(topo);

        const corpo = el("div", "pm-corpo");
        if (!this.temDados) {
            corpo.appendChild(el("div", "pm-vazio", this.filtros.ativos() ? "Sem dados para os filtros escolhidos." : `Arraste a medida ${MEDIDA} para o campo do visual.`));
        } else {
            this.blocos.forEach((b, i) => corpo.appendChild(this.bloco(b, i)));
        }
        raiz.appendChild(corpo);
        corpo.scrollTop = this.rolagem;
        if (this.filtros.popAberto) { raiz.appendChild(this.filtros.painel()); }
    }

    private bloco(b: BlocoMetas, i: number): HTMLElement {
        const s = el("section", "pm-bloco");
        const nome = b.uf ? "PAINEL METAS | " + b.uf : "Total";
        s.setAttribute("aria-label", b.uf ? nome : "Painel de metas: total");
        if (b.uf) { s.appendChild(el("h2", "pm-titulo pm-tuf", nome)); }
        s.appendChild(this.kpis(b, i === 0));
        s.appendChild(el("h3", "pm-sub", "PAINEL DE VENDAS - MÊS"));
        s.appendChild(this.tabela(b, nome));
        s.appendChild(el("h3", "pm-sub", "VENDIDO DIA x DIA"));
        s.appendChild(this.dias(b, nome));
        return s;
    }

    private kpis(b: BlocoMetas, total: boolean): HTMLElement {
        const k = b.k;
        type Item = { rot: string; cab: "c" | "t" | "n"; v: string; cls?: string };
        const grupos: Item[][] = [
            [{ rot: "Objetivo Clientes", cab: "c", v: f(k[0]) }, { rot: "Clientes(V)", cab: "t", v: f(k[1]) }, { rot: "Clientes (F)", cab: "n", v: f(k[2]) }, { rot: "%", cab: "n", v: f(k[3], "pct"), cls: "pc" }],
            [{ rot: "Clientes Carteira", cab: "c", v: f(k[4]) }, { rot: "Tendência (F+C)", cab: "c", v: f(k[5]) }],
            [{ rot: "Prazo (V)", cab: "c", v: f(k[6], "dec2") }, { rot: "Nota MIX (V)", cab: "c", v: f(k[7], "dec1") }]
        ];
        if (total) { grupos.push([{ rot: "Portas Qtd (V)", cab: "c", v: "", cls: "am" }]); }
        const w = el("div", "pm-kpis");
        grupos.forEach(g => {
            const t = el("table", "pm-kpi");
            const h = el("tr"), v = el("tr");
            g.forEach(it => {
                const th = el("th", "k" + it.cab, it.rot); th.setAttribute("scope", "col"); h.appendChild(th);
                const td = el("td", it.cls || "", it.v);
                if (!it.v) { td.setAttribute("aria-label", it.rot + ": sem valor"); }
                v.appendChild(td);
            });
            t.appendChild(h); t.appendChild(v);
            w.appendChild(t);
        });
        return w;
    }

    private tabela(b: BlocoMetas, nome: string): HTMLElement {
        const w = el("div", "pm-tw");
        const t = el("table", "pm-tab");
        t.setAttribute("aria-label", "Painel de vendas do mês por família, " + nome);
        const cg = el("colgroup");
        [13, 7.2, 8.2, 7.8, 7.2, 7.8, 6.2, 8.2, 8.2, 6.4, 6.2, 7.2, 6.4].forEach(x => { const c = el("col"); c.style.width = x + "%"; cg.appendChild(c); });
        t.appendChild(cg);
        const th = el("thead"), hr = el("tr");
        const h0 = el("th", "hb", "Família"); h0.setAttribute("scope", "col"); hr.appendChild(h0);
        COLS_TAB.forEach(c => { const h = el("th", "h" + c.cab, c.rot); h.setAttribute("scope", "col"); hr.appendChild(h); });
        th.appendChild(hr); t.appendChild(th);
        const tb = el("tbody");
        const linha = (l: LinhaMetas, tot: boolean): void => {
            const tr = el("tr", tot ? "tot" : "");
            const nm = el(tot ? "td" : "th", "nm", tot ? "" : l.nome);
            if (!tot) { nm.setAttribute("scope", "row"); } else { nm.setAttribute("aria-label", "Total"); }
            tr.appendChild(nm);
            COLS_TAB.forEach((c, i) => tr.appendChild(el("td", c.cel, f(l.v[i], c.tipo === "pct" ? "pct" : "int"))));
            tb.appendChild(tr);
        };
        b.fams.forEach(l => linha(l, false));
        linha(b.total, true);
        t.appendChild(tb);
        w.appendChild(t);
        return w;
    }

    private dias(b: BlocoMetas, nome: string): HTMLElement {
        const w = el("div", "pm-dw");
        w.setAttribute("tabindex", "0");
        w.setAttribute("role", "region");
        w.setAttribute("aria-label", "Vendido dia a dia, " + nome + " (role para os lados se precisar)");
        const t = el("table", "pm-dias");
        const th = el("thead"), hr = el("tr");
        const h0 = el("th", "nm", ""); h0.setAttribute("aria-label", "Família"); hr.appendChild(h0);
        for (let d = 1; d <= 31; d++) { const h = el("th", "", (d < 10 ? "0" : "") + d); h.setAttribute("scope", "col"); h.setAttribute("aria-label", "Dia " + d); hr.appendChild(h); }
        th.appendChild(hr); t.appendChild(th);
        const tb = el("tbody");
        const linha = (rot: string, vals: N[], cls: string, zero: boolean): void => {
            const tr = el("tr", cls);
            const nm = el("th", "nm", rot); nm.setAttribute("scope", "row"); tr.appendChild(nm);
            for (let d = 0; d < 31; d++) {
                const v = vals[d];
                tr.appendChild(el("td", "", v === null || v === undefined ? (zero ? "0" : "") : f(v)));
            }
            tb.appendChild(tr);
        };
        b.fams.forEach(l => linha(l.nome, l.d, "", false));
        linha("TOTAL", b.total.d, "tot", false);
        linha("CLIENTES", b.cli, "tot", true);
        t.appendChild(tb);
        w.appendChild(t);
        return w;
    }

    // ---------- exportar ----------
    private escolhas = { ufs: true, dados: false };
    private menuAberto = false;
    private aviso = "";
    private avisoErro = false;

    private menuExportar(): HTMLElement {
        const cx = el("div", "ex");
        const btn = botao("pm-btn ex-btn", "Exportar ▾", () => { this.menuAberto = !this.menuAberto; this.limparAviso(); if (this.menuAberto) { this.filtros.popAberto = false; } this.desenhar(); });
        btn.setAttribute("aria-haspopup", "true");
        btn.setAttribute("aria-expanded", String(this.menuAberto));
        cx.appendChild(btn);
        if (this.menuAberto) {
            const m = el("div", "ex-menu");
            m.appendChild(el("div", "ex-tit", "Incluir no Excel (aba igual à PAINEL METAS)"));
            const itens: ["ufs" | "dados", string][] = [["ufs", "Blocos por UF"], ["dados", "Aba com dados brutos"]];
            itens.forEach(([k, rot]) => {
                const lb = el("label", "ex-op");
                const cb = document.createElement("input");
                cb.type = "checkbox"; cb.checked = !!this.escolhas[k];
                cb.addEventListener("change", () => { this.escolhas[k] = cb.checked; });
                lb.appendChild(cb); lb.appendChild(document.createTextNode(" " + rot));
                m.appendChild(lb);
            });
            m.appendChild(botao("ex-go", "Baixar Excel", () => { void this.exportarExcel(); }));
            cx.appendChild(m);
            setTimeout(() => { const p = m.querySelector("input") as HTMLInputElement | null; if (p) { p.focus(); } }, 0);
        }
        if (this.aviso) { cx.appendChild(el("div", "ex-msg" + (this.avisoErro ? " erro" : ""), this.aviso)); }
        return cx;
    }

    private timerAviso: number | undefined;

    private limparAviso(): void {
        if (this.timerAviso !== undefined) { window.clearTimeout(this.timerAviso); this.timerAviso = undefined; }
        this.aviso = "";
        this.avisoErro = false;
    }

    private mostrar(t: string, erro = false): void {
        this.limparAviso();
        this.aviso = t; this.avisoErro = erro;
        this.desenhar();
        this.timerAviso = window.setTimeout(() => {
            this.timerAviso = undefined;
            this.aviso = ""; this.avisoErro = false;
            const m = this.raiz.querySelector(".ex-msg") as HTMLElement | null;
            if (!m) { return; }
            const reduz = typeof window.matchMedia === "function" && window.matchMedia("(prefers-reduced-motion: reduce)").matches;
            if (reduz) { m.remove(); return; }
            m.classList.add("sai");
            window.setTimeout(() => { m.remove(); }, 350);
        }, erro ? 10000 : 6000);
    }

    private agora(): string {
        const d = new Date();
        const p = (n: number): string => (n < 10 ? "0" : "") + n;
        return `${p(d.getDate())}/${p(d.getMonth() + 1)}/${d.getFullYear()} ${p(d.getHours())}:${p(d.getMinutes())}`;
    }

    private entrada(): EntradaMetas {
        return {
            titulo: this.settings.titulo.texto.value || "PAINEL DE METAS | Novaforma",
            atualizacao: this.atualizacao, geradoEm: this.agora(), periodo: this.periodo(), filtros: this.filtros.resumo(),
            blocos: this.escolhas.ufs ? this.blocos : this.blocos.slice(0, 1), opcoes: { dados: this.escolhas.dados }
        };
    }

    private async exportarExcel(): Promise<void> {
        if (!this.temDados) { this.mostrar("Sem dados para exportar.", true); return; }
        this.menuAberto = false;
        const ex = this.settings.exportacao;
        const nome = (ex.nomeArquivo.value || "painel_metas").replace(/[^\w-]+/g, "_") + ".xlsx";
        const url = (ex.urlPagina.value || "").trim();
        const entrada = this.entrada();
        try {
            if (ex.usarPagina.value && /^https:\/\//i.test(url)) { await this.abrirPaginaDownload(url, nome, entrada); return; }
            this.mostrar("Gerando Excel…");
            const bytes = gerarXlsxMetas(entrada);
            let permitido = false;
            try { permitido = (await this.host.downloadService.exportStatus()) === powerbi.PrivilegeStatus.Allowed; } catch { permitido = false; }
            if (permitido) {
                const r = await this.host.downloadService.exportVisualsContentExtended(paraBase64(bytes), nome, "base64", "Exportação do Painel Metas para Excel");
                this.mostrar(r && (r as { downloadCompleted?: boolean }).downloadCompleted === false ? "Download cancelado." : "Excel exportado: " + nome);
            } else {
                const a = document.createElement("a");
                a.href = URL.createObjectURL(new Blob([bytes as BlobPart], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
                a.download = nome; this.raiz.appendChild(a); a.click(); this.raiz.removeChild(a);
                this.mostrar("Download direto (exportação do Power BI não liberada pelo administrador).");
            }
        } catch (e) { this.mostrar("Erro ao exportar: " + (e instanceof Error ? e.message : String(e)), true); }
    }

    // Link curto: valores em R$ e quantidades vão arredondados a inteiro (a aba mostra #,##0); percentuais com 4 casas.
    public static compactar(e: EntradaMetas): object {
        const r0 = (v: N): N => (v === null || v === undefined ? null : Math.round(v));
        const r4 = (v: N): N => (v === null || v === undefined ? null : Math.round(v * 10000) / 10000);
        const corta = (a: N[]): N[] => { const b = a.slice(); while (b.length && (b[b.length - 1] === null || b[b.length - 1] === undefined)) { b.pop(); } return b; };
        const lin = (l: LinhaMetas): LinhaMetas => ({ nome: l.nome, v: l.v.map((v, i) => (i === 5 ? r4(v) : r0(v))), d: corta(l.d.map(r0)) });
        return {
            tipo: "painelMetas", versao: 1, titulo: e.titulo, atualizacao: e.atualizacao, geradoEm: e.geradoEm, periodo: e.periodo, filtros: e.filtros, opcoes: e.opcoes,
            blocos: e.blocos.map(b => ({ uf: b.uf, k: b.k.map(r4), total: lin(b.total), fams: b.fams.map(lin), cli: corta(b.cli) }))
        };
    }

    // Abre a página de download levando só os dados (texto pequeno) no fragmento (#).
    private async abrirPaginaDownload(url: string, nome: string, entrada: EntradaMetas): Promise<void> {
        let dados: Uint8Array = new TextEncoder().encode(JSON.stringify(Visual.compactar(entrada)));
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
        this.mostrar("Página de download aberta: " + nome);
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.fmtService.buildFormattingModel(this.settings);
    }
}
