"use strict";
// Painel Metas | Representante (NovaForma) v1.1: "Carregando…" até os dados chegarem; a medida v1.1 entrega todos os blocos numa célula.
// v1.0: réplica da aba "(Data do APP) PAINEL METAS | Re": um bloco por representante
// (PAINEL DE VENDAS - MÊS | NOME - UF, KPIs de clientes e tabela por família), regra Data do APP, chave = nome do representante.
// Desenho e Excel por molde vindos do visual Painel Metas; botão Filtros (filtra a página) e Exportar (Excel igual à aba:
// download direto, página de download pelo link ou área de transferência + página "Colar dados" quando os blocos não cabem no link).
import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import "./../style/visual.less";
import { VisualSettings } from "./settings";
import { Filtros, FiltroBasico, el, botao, dataBR } from "./filtros";
import { lerCarga, conferir, textoAtualizacao, Formato } from "./dados";
import { gerarXlsxMetasRep, paraBase64, compactar, BlocoRep, LinhaRep, EntradaRep, N, ABA, ROTULOS_V, ZERO_SE_VAZIO, IDX_PCT, tituloBloco } from "./xlsxMetasRep";

import IVisual = powerbi.extensibility.visual.IVisual;
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;

const MEDIDA = "Painel Metas Rep (visual)";
const ARQUIVO_DAX = "Painel_Metas_Rep_medida.dax";
// O Power BI recusa links longos ("link malformado"); acima disso os dados vão pela área de transferência.
export const LIMITE_LINK = 2500;
export const PREFIXO_COLA = "NFMR1.";
const fmtInt = new Intl.NumberFormat("pt-BR", { maximumFractionDigits: 0 });
const fmtPct = new Intl.NumberFormat("pt-BR", { style: "percent", minimumFractionDigits: 1, maximumFractionDigits: 1 });
function f(v: N | undefined, pct = false): string {
    if (v === null || v === undefined || isNaN(v)) { return ""; }
    return pct ? fmtPct.format(v) : fmtInt.format(Math.round(v));
}

// colunas da tabela por família (cores da aba: t = verde-petróleo, n = azul-marinho, b = branco)
const COLS_TAB: { cab: "b" | "t" | "n"; cel: string }[] = [
    { cab: "b", cel: "obj" }, { cab: "t", cel: "tl" }, { cab: "t", cel: "tl" }, { cab: "b", cel: "" },
    { cab: "n", cel: "az" }, { cab: "n", cel: "pc" }, { cab: "b", cel: "" }, { cab: "b", cel: "" },
    { cab: "t", cel: "tl" }, { cab: "t", cel: "tl" }, { cab: "n", cel: "az" }, { cab: "b", cel: "" }
];
// larguras relativas (colunas B:N da aba), ajustadas para texto de 15 px
const LARGURAS = [226, 92, 100, 104, 96, 100, 84, 98, 104, 84, 90, 98, 90];

export class Visual implements IVisual {
    private host: IVisualHost;
    private raiz: HTMLElement;
    private topo: HTMLElement;
    private corpo: HTMLElement;
    private fmtService = new FormattingSettingsService();
    private settings: VisualSettings = new VisualSettings();
    private filtros: Filtros;
    private blocos: BlocoRep[] = [];
    private atualizacao = "";
    private formato: Formato = "ok";
    private temDados = false;
    private carregando = false;
    private limiteAtingido = false;
    private corpoPronto = false;

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.raiz = el("div", "pm");
        this.topo = el("div", "pm-topo");
        this.corpo = el("div", "pm-corpo");
        this.raiz.appendChild(this.topo);
        this.raiz.appendChild(this.corpo);
        options.element.appendChild(this.raiz);
        // estado inicial até o Power BI entregar os dados da medida
        const car = el("div", "pm-vazio pm-carregando", "Carregando…");
        car.setAttribute("role", "status");
        car.setAttribute("aria-live", "polite");
        this.corpo.appendChild(car);
        this.filtros = new Filtros(this.host, this.raiz, () => this.desenhar(), () => this.settings.filtros, () => this.periodo(), MEDIDA);
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
        this.carregando = false;
        this.corpoPronto = false;
        if (dv && dv.metadata && dv.metadata.segment) {
            this.carregando = this.host.fetchMoreData(true);
            this.limiteAtingido = !this.carregando;
        } else { this.limiteAtingido = false; }
        this.desenhar();
    }

    private lerDados(dv: powerbi.DataView | undefined): void {
        this.temDados = false; this.formato = "ok"; this.blocos = [];
        const t = dv && dv.table;
        if (!t || !t.rows.length) { return; }
        let idx = t.columns.findIndex(c => c.roles && c.roles["dados"]);
        if (idx < 0) { idx = t.columns.length - 1; }
        const celulas = t.rows.map(r => r[idx]).filter(v => typeof v === "string" && v) as string[];
        if (!celulas.length) {
            if (t.rows.some(r => r[idx] !== null && r[idx] !== undefined)) { this.formato = "errada"; }
            return;
        }
        this.formato = conferir(celulas);
        if (this.formato !== "ok") { return; }
        const c = lerCarga(celulas);
        if (c.cabecalho.length) {
            this.atualizacao = textoAtualizacao(c.cabecalho[0]);
            this.filtros.lerCabecalho(c.cabecalho, 1);
        }
        this.blocos = c.blocos;
        this.temDados = true;
    }

    private periodo(): string {
        const r = this.filtros.dataRef;
        return r ? `Período: ${dataBR(r.slice(0, 8) + "01")} a ${dataBR(r)} (Data do APP)` : "";
    }

    // ---------- tela ----------
    private desenhar(): void {
        const s = this.settings;
        this.raiz.style.background = s.fundo.mostrar.value ? s.fundo.cor.value.value : "transparent";
        const topo = this.topo;
        while (topo.firstChild) { topo.removeChild(topo.firstChild); }
        if (s.titulo.mostrar.value) { topo.appendChild(el("h2", "pm-titulo", s.titulo.texto.value || "PAINEL METAS | REPRESENTANTE")); }
        else { topo.appendChild(el("div", "pm-espaco")); }
        if (this.atualizacao && this.temDados) { topo.appendChild(el("div", "pm-atu", this.atualizacao)); }
        const acoes = el("div", "pm-acoes");
        if (s.filtros.mostrarBotao.value) { acoes.appendChild(this.filtros.botao("pm-btn", () => { this.menuAberto = false; })); }
        if (s.exportacao.mostrarBotao.value) { acoes.appendChild(this.menuExportar()); }
        topo.appendChild(acoes);
        if (!this.corpoPronto) { this.corpoPronto = true; this.desenharCorpo(); }
        const pop = this.raiz.querySelector(".fp");
        if (pop) { this.raiz.removeChild(pop); }
        if (this.filtros.popAberto) { this.raiz.appendChild(this.filtros.painel()); }
    }

    private desenharCorpo(): void {
        const corpo = this.corpo;
        const rolagem = corpo.scrollTop;
        while (corpo.firstChild) { corpo.removeChild(corpo.firstChild); }
        if (this.formato === "outra") {
            corpo.appendChild(el("div", "pm-vazio erro", `A medida do campo é a Painel Metas (visual), do painel por UF. Este visual é o Painel Metas | Representante: troque pela medida ${MEDIDA} (arquivo ${ARQUIVO_DAX}).`));
            return;
        }
        if (this.formato === "errada") {
            corpo.appendChild(el("div", "pm-vazio erro", `A medida do campo não é a ${MEDIDA}. Troque pela medida do arquivo ${ARQUIVO_DAX} (campo "Painel Metas Rep (medida)").`));
            return;
        }
        if (!this.temDados) {
            corpo.appendChild(el("div", "pm-vazio", this.filtros.ativos() ? "Sem dados para os filtros escolhidos." : `Arraste o código do representante e a medida ${MEDIDA} para os campos do visual.`));
            return;
        }
        if (!this.blocos.length) {
            corpo.appendChild(el("div", "pm-vazio", "Sem representantes com meta para os filtros escolhidos."));
        }
        this.blocos.forEach(b => corpo.appendChild(this.bloco(b)));
        if (this.carregando) { corpo.appendChild(el("div", "pm-status", "Carregando mais representantes…")); }
        else if (this.limiteAtingido) { corpo.appendChild(el("div", "pm-status erro", "O Power BI parou de enviar linhas (limite de memória do visual). Use os filtros para ver o restante.")); }
        corpo.scrollTop = rolagem;
    }

    private bloco(b: BlocoRep): HTMLElement {
        const s = el("section", "pm-bloco");
        const nome = tituloBloco(b);
        s.setAttribute("aria-label", b.nome + (b.uf ? " - " + b.uf : ""));
        s.appendChild(el("h3", "pm-sub", nome));
        s.appendChild(this.kpis(b));
        s.appendChild(this.tabela(b));
        return s;
    }

    private kpis(b: BlocoRep): HTMLElement {
        const k = b.k;
        type Item = { rot: string; cab: "c" | "t" | "n"; v: string; cls?: string };
        const itens: Item[] = [
            { rot: "Objetivo Clientes", cab: "c", v: f(k[0]) }, { rot: "Clientes(V)", cab: "t", v: f(k[1]) },
            { rot: "Clientes (F)", cab: "n", v: f(k[2]) }, { rot: "%", cab: "n", v: f(k[3], true), cls: "pc" }
        ];
        const t = el("table", "pm-kpi");
        t.setAttribute("aria-label", "Clientes de " + b.nome);
        const h = el("tr"), v = el("tr");
        itens.forEach(it => {
            const th = el("th", "k" + it.cab, it.rot); th.setAttribute("scope", "col"); h.appendChild(th);
            const td = el("td", it.cls || "", it.v);
            if (!it.v) { td.setAttribute("aria-label", it.rot + ": sem valor"); }
            v.appendChild(td);
        });
        t.appendChild(h); t.appendChild(v);
        return t;
    }

    private tabela(b: BlocoRep): HTMLElement {
        const w = el("div", "pm-tw");
        const t = el("table", "pm-tab");
        t.setAttribute("aria-label", "Vendas do mês por família, " + b.nome);
        const cg = el("colgroup");
        const soma = LARGURAS.reduce((a, x) => a + x, 0);
        LARGURAS.forEach(x => { const c = el("col"); c.style.width = (x / soma * 100).toFixed(3) + "%"; cg.appendChild(c); });
        t.appendChild(cg);
        const th = el("thead"), hr = el("tr");
        const h0 = el("th", "hb", "Família"); h0.setAttribute("scope", "col"); hr.appendChild(h0);
        COLS_TAB.forEach((c, i) => { const h = el("th", "h" + c.cab, ROTULOS_V[i]); h.setAttribute("scope", "col"); hr.appendChild(h); });
        th.appendChild(hr); t.appendChild(th);
        const tb = el("tbody");
        const linha = (l: LinhaRep, tot: boolean): void => {
            const tr = el("tr", tot ? "tot" : "");
            const nm = el(tot ? "td" : "th", "nm", tot ? "" : l.nome);
            if (!tot) { nm.setAttribute("scope", "row"); } else { nm.setAttribute("aria-label", "Total"); }
            tr.appendChild(nm);
            COLS_TAB.forEach((c, i) => {
                let v = l.v[i];
                if ((v === null || v === undefined) && ZERO_SE_VAZIO.indexOf(i) >= 0) { v = 0; }
                const pct = i === IDX_PCT;
                const cls = c.cel + (pct && v !== null && v !== undefined && v >= 1 ? " meta" : "");
                tr.appendChild(el("td", cls, f(v, pct)));
            });
            tb.appendChild(tr);
        };
        b.fams.forEach(l => linha(l, false));
        linha(b.total, true);
        t.appendChild(tb);
        w.appendChild(t);
        return w;
    }

    // ---------- exportar ----------
    private escolhas = { tabela: true, dados: false };
    private menuAberto = false;
    private aviso = "";
    private avisoErro = false;

    private menuExportar(): HTMLElement {
        const cx = el("div", "ex");
        const btn = botao("pm-btn ex-btn", "Exportar ▾", () => { this.menuAberto = !this.menuAberto; this.aviso = ""; if (this.menuAberto) { this.filtros.popAberto = false; } this.desenhar(); });
        btn.setAttribute("aria-haspopup", "true");
        btn.setAttribute("aria-expanded", String(this.menuAberto));
        cx.appendChild(btn);
        if (this.menuAberto) {
            const m = el("div", "ex-menu");
            m.appendChild(el("div", "ex-tit", "Incluir no Excel"));
            const itens: ["tabela" | "dados", string][] = [["tabela", "Tabela (igual à aba PAINEL METAS | Representante)"], ["dados", "Aba com dados brutos"]];
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
        if (this.aviso) {
            const msg = el("div", "ex-msg" + (this.avisoErro ? " erro" : ""), this.aviso);
            msg.setAttribute("role", "status");
            const x = botao("ex-msg-x", "×", () => { this.aviso = ""; this.desenhar(); });
            x.setAttribute("aria-label", "Fechar aviso");
            msg.appendChild(x);
            cx.appendChild(msg);
        }
        return cx;
    }

    private mostrar(t: string, erro = false): void { this.aviso = t; this.avisoErro = erro; this.desenhar(); }

    private agora(): string {
        const d = new Date();
        const p = (n: number): string => (n < 10 ? "0" : "") + n;
        return `${p(d.getDate())}/${p(d.getMonth() + 1)}/${d.getFullYear()} ${p(d.getHours())}:${p(d.getMinutes())}`;
    }

    private entrada(): EntradaRep {
        return {
            titulo: this.settings.titulo.texto.value || "PAINEL METAS | REPRESENTANTE", aba: ABA,
            atualizacao: this.atualizacao, geradoEm: this.agora(), periodo: this.periodo(), filtros: this.filtros.resumo(),
            blocos: this.blocos, opcoes: { tabela: this.escolhas.tabela, dados: this.escolhas.dados }
        };
    }

    // Copia texto para a área de transferência: API do navegador e, se o iframe do Power BI bloquear, textarea + execCommand.
    private async copiar(texto: string): Promise<boolean> {
        try {
            const cb = navigator.clipboard;
            if (cb && cb.writeText) { await cb.writeText(texto); return true; }
        } catch { /* bloqueado pela política do iframe: tenta o modo antigo */ }
        const ta = document.createElement("textarea");
        ta.value = texto;
        ta.setAttribute("readonly", "");
        ta.className = "pm-copia";
        this.raiz.appendChild(ta);
        let ok = false;
        try { ta.focus(); ta.select(); ta.setSelectionRange(0, texto.length); ok = document.execCommand("copy"); } catch { ok = false; }
        this.raiz.removeChild(ta);
        return ok;
    }

    private async exportarExcel(): Promise<void> {
        const o = this.escolhas;
        if (!this.temDados || !this.blocos.length) { this.mostrar("Sem dados para exportar.", true); return; }
        if (!o.tabela && !o.dados) { this.mostrar("Marque pelo menos um item para exportar.", true); return; }
        this.menuAberto = false;
        const ex = this.settings.exportacao;
        const nome = (ex.nomeArquivo.value || "painel_metas_representante").replace(/[^\w-]+/g, "_") + ".xlsx";
        const url = (ex.urlPagina.value || "").trim();
        const entrada = this.entrada();
        try {
            // 1) download direto pelo Power BI (Desktop e serviço com a exportação liberada pelo administrador): arquivo completo
            let permitido = false;
            try { permitido = (await this.host.downloadService.exportStatus()) === powerbi.PrivilegeStatus.Allowed; } catch { permitido = false; }
            if (permitido) {
                this.mostrar("Gerando Excel…");
                const r = await this.host.downloadService.exportVisualsContentExtended(paraBase64(gerarXlsxMetasRep(entrada)), nome, "base64", "Exportação do Painel Metas | Representante para Excel");
                this.mostrar(r && (r as { downloadCompleted?: boolean }).downloadCompleted === false ? "Download cancelado." : "Excel exportado: " + nome);
                return;
            }
            // 2) página de download: dados no link quando cabem; senão pela área de transferência
            if (ex.usarPagina.value && /^https:\/\//i.test(url)) { await this.abrirPaginaDownload(url, nome, entrada); return; }
            // 3) sem página: tenta o download do navegador (funciona fora do iframe restrito do Power BI)
            const a = document.createElement("a");
            a.href = URL.createObjectURL(new Blob([gerarXlsxMetasRep(entrada) as BlobPart], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
            a.download = nome; this.raiz.appendChild(a); a.click(); this.raiz.removeChild(a);
            this.mostrar("Download direto (exportação do Power BI não liberada pelo administrador).");
        } catch (e) { this.mostrar("Erro ao exportar: " + (e instanceof Error ? e.message : String(e)), true); }
    }

    // carga compacta -> deflate-raw -> base64url
    public static async codificar(entrada: EntradaRep): Promise<{ b64: string; comp: string }> {
        let dados: Uint8Array = new TextEncoder().encode(JSON.stringify(compactar(entrada)));
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
        return { b64: paraBase64(dados).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, ""), comp };
    }

    private async abrirPaginaDownload(url: string, nome: string, entrada: EntradaRep): Promise<void> {
        const { b64, comp } = await Visual.codificar(entrada);
        const base = url.replace(/#.*$/, "") + "#";
        const alvo = base + new URLSearchParams({ v: "2", n: nome, c: comp, j: b64 }).toString();
        const qtd = fmtInt.format(entrada.blocos.length) + (entrada.blocos.length === 1 ? " representante" : " representantes");
        if (alvo.length <= LIMITE_LINK) {
            this.host.launchUrl(alvo);
            this.mostrar(`Página de download aberta: ${nome}`);
            return;
        }
        // Blocos demais para o link: os dados vão pela área de transferência e o link leva só o aviso (k=cola).
        const copiou = comp === "1" && await this.copiar(PREFIXO_COLA + b64);
        if (copiou) {
            this.host.launchUrl(base + new URLSearchParams({ v: "2", n: nome, k: "cola", q: String(entrada.blocos.length), u: "representantes" }).toString());
            this.mostrar(`Os ${qtd} não cabem no link. Os dados foram copiados: na página que abriu, clique em Colar dados.`);
        } else {
            this.mostrar(`Os ${qtd} não cabem no link e o Power BI bloqueou a cópia. Filtre por Supervisão ou Representante e exporte de novo.`, true);
        }
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.fmtService.buildFormattingModel(this.settings);
    }
}
