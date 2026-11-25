"use strict";
// Painel de Vendas (NovaForma) v1.0: réplica da aba (Data Aprovação | Data APP) EVOLUÇÃO VENDA ("PAINEL DE VENDAS - DIA"):
// vendido, faturado (3 meses), carteira em aberto, clientes (3 meses) e mix do mês atual por representante.
// O mesmo código atende os dois visuais; só src/config.ts (regra, medida, nomes) muda entre eles.
// Botão Filtros (filtra a página) e Exportar (Excel igual à aba: download direto, página de download pelo link ou
// área de transferência + página "Colar dados" quando a tabela não cabe no link).
import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import "./../style/visual.less";
import { VisualSettings } from "./settings";
import { Filtros, FiltroBasico, el, botao, dataBR } from "./filtros";
import { lerCarga, conferir, Ordem, Formato } from "./dados";
import { gerarXlsxPainelVendas, paraBase64, compactar, tsv, rotulos, LinhaPV, EntradaPV, N, NCOLS, FAMILIAS } from "./xlsxPainelVendas";
import { CONFIG } from "./config";

import IVisual = powerbi.extensibility.visual.IVisual;
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;

// O Power BI recusa links longos ("link malformado"); acima disso os dados vão pela área de transferência.
export const LIMITE_LINK = 2500;
// prefixo do texto copiado para a página de download (formato compacto, deflate-raw + base64url)
export const PREFIXO_COLA = "NFPV1.";
const fmtN = new Intl.NumberFormat("pt-BR", { maximumFractionDigits: 0 });
const fmtPct = new Intl.NumberFormat("pt-BR", { style: "percent", maximumFractionDigits: 0 });
// larguras relativas das colunas na tela (B, C, D, E:K, L:N, O:S da aba), ajustadas para texto de 15 px sem cortar números
const LARGURAS = [212, 38, 92, ...Array(7).fill(72), ...Array(3).fill(68), ...Array(5).fill(44)];
const GRUPO_CLS = ["g-vend", "g-fat", "g-cart", "g-cli", "g-mix"];
const GRUPO_SPAN = [3, 3, 1, 3, 5];
const GRUPO_ROT = ["VENDIDO", "FATURADO", "CARTEIRA", "CLIENTES (VENDIDO)", "MIX (VENDIDO %)"];

function fmt(v: N, i: number): string {
    if (v === null || v === undefined || isNaN(v)) { return ""; }
    return i >= 10 ? fmtPct.format(v) : fmtN.format(Math.round(v));
}

export class Visual implements IVisual {
    private host: IVisualHost;
    private raiz: HTMLElement;
    private topo: HTMLElement;
    private corpo: HTMLElement;
    private fmtService = new FormattingSettingsService();
    private settings: VisualSettings = new VisualSettings();
    private filtros: Filtros;
    private linhas: LinhaPV[] = [];
    private meses: string[] = [];
    private formato: Formato = "ok";
    private temDados = false;
    private carregando = false;
    private limiteAtingido = false;
    private corpoPronto = false;

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.raiz = el("div", "pv");
        this.topo = el("div", "pv-topo");
        this.corpo = el("div", "pv-corpo");
        this.raiz.appendChild(this.topo);
        this.raiz.appendChild(this.corpo);
        options.element.appendChild(this.raiz);
        this.filtros = new Filtros(this.host, this.raiz, () => this.desenhar(), () => this.settings.filtros, () => this.periodo(), CONFIG.medida);
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

    private ordem(): Ordem {
        const v = this.settings.tabela.ordem.value;
        const s = String((v && (v as { value?: string }).value) || "nome");
        return (s === "vendido" || s === "codigo" ? s : "nome") as Ordem;
    }

    private lerDados(dv: powerbi.DataView | undefined): void {
        this.temDados = false; this.formato = "ok"; this.linhas = [];
        const t = dv && dv.table;
        if (!t || !t.rows.length) { return; }
        let idx = t.columns.findIndex(c => c.roles && c.roles["dados"]);
        if (idx < 0) { idx = t.columns.length - 1; }
        const celulas = t.rows.map(r => r[idx]).filter(v => typeof v === "string" && v) as string[];
        if (!celulas.length) {
            // números ou outro tipo no campo: não é a medida de carga
            if (t.rows.some(r => r[idx] !== null && r[idx] !== undefined)) { this.formato = "errada"; }
            return;
        }
        this.formato = conferir(celulas, CONFIG.tipo, CONFIG.outroTipo);
        if (this.formato !== "ok") { return; }
        const c = lerCarga(celulas, this.ordem());
        if (c.cabecalho.length) {
            this.meses = c.cabecalho.slice(0, 3);
            this.filtros.lerCabecalho(c.cabecalho, 3);
        }
        this.linhas = c.linhas;
        this.temDados = true;
    }

    private periodo(): string {
        const m = this.meses.filter(x => !!x);
        const ref = this.filtros.dataRef ? ` · Data Ref ${dataBR(this.filtros.dataRef)}` : "";
        return m.length ? `Meses do painel: ${m[0]} a ${m[m.length - 1]}${ref} (Ano, Mês e Dia não mudam este painel)` : "";
    }

    // ---------- tela ----------
    private desenhar(): void {
        const s = this.settings;
        this.raiz.style.background = s.fundo.mostrar.value ? s.fundo.cor.value.value : "transparent";
        const topo = this.topo;
        while (topo.firstChild) { topo.removeChild(topo.firstChild); }
        if (s.titulo.mostrar.value) {
            const h = el("h2", "pv-titulo", s.titulo.texto.value || "PAINEL DE VENDAS - DIA");
            if (s.titulo.mostrarRegra.value) { h.appendChild(el("span", "pv-regra", CONFIG.regra)); }
            topo.appendChild(h);
        } else { topo.appendChild(el("div", "pv-espaco")); }
        const acoes = el("div", "pv-acoes");
        if (s.filtros.mostrarBotao.value) { acoes.appendChild(this.filtros.botao("pv-btn", () => { this.menuAberto = false; })); }
        if (s.exportacao.mostrarBotao.value) { acoes.appendChild(this.menuExportar()); }
        topo.appendChild(acoes);
        if (!this.corpoPronto) { this.corpoPronto = true; this.desenharCorpo(); }
        const pop = this.raiz.querySelector(".fp");
        if (pop) { this.raiz.removeChild(pop); }
        if (this.filtros.popAberto) { this.raiz.appendChild(this.filtros.painel()); }
    }

    private desenharCorpo(): void {
        const corpo = this.corpo;
        const rolagem = corpo.querySelector(".pv-tw") ? (corpo.querySelector(".pv-tw") as HTMLElement).scrollTop : 0;
        while (corpo.firstChild) { corpo.removeChild(corpo.firstChild); }
        if (this.formato === "outra") {
            corpo.appendChild(el("div", "pv-vazio erro", `A medida do campo é a ${CONFIG.outraMedida}, do outro painel. Este visual é o ${CONFIG.nomeVisual}: troque pela medida ${CONFIG.medida} (arquivo ${CONFIG.arquivoDax}).`));
            return;
        }
        if (this.formato === "errada") {
            corpo.appendChild(el("div", "pv-vazio erro", `A medida do campo não é a ${CONFIG.medida}. Troque pela medida do arquivo ${CONFIG.arquivoDax} (campo "${CONFIG.medida.replace(" (visual)", "")} (medida)").`));
            return;
        }
        if (!this.temDados) {
            corpo.appendChild(el("div", "pv-vazio", this.filtros.ativos() ? "Sem dados para os filtros escolhidos." : `Arraste o código do representante e a medida ${CONFIG.medida} para os campos do visual.`));
            return;
        }
        const area = el("div", "pv-tw");
        area.setAttribute("role", "region");
        area.setAttribute("tabindex", "0");
        area.setAttribute("aria-label", `Painel de vendas, ${CONFIG.regra}: ${fmtN.format(this.linhas.length)} representantes (role para ver todos)`);
        const t = el("table", "pv-tab");
        const cg = el("colgroup");
        const soma = LARGURAS.reduce((a, b) => a + b, 0);
        LARGURAS.forEach(w => { const c = el("col"); c.style.width = (w / soma * 100).toFixed(3) + "%"; cg.appendChild(c); });
        t.appendChild(cg);
        const th = el("thead");
        // linha 1: faixas (como a linha 4 da aba)
        const g = el("tr", "grupos");
        const g0 = el("td", "g-vazio"); g0.setAttribute("colspan", "3"); g0.setAttribute("aria-hidden", "true"); g.appendChild(g0);
        GRUPO_ROT.forEach((rot, i) => {
            const c = el("th", GRUPO_CLS[i], rot); c.setAttribute("colspan", String(GRUPO_SPAN[i])); c.setAttribute("scope", "colgroup"); g.appendChild(c);
        });
        th.appendChild(g);
        // linha 2: cabeçalho (meses dinâmicos)
        const hr = el("tr", "cab");
        ["REPRESENTANTES", "UF", "SUP"].forEach((x, i) => { const c = el("th", i === 0 ? "rep" : "", x); c.setAttribute("scope", "col"); hr.appendChild(c); });
        const rot = rotulos(this.meses);
        rot.forEach((x, i) => {
            const c = el("th", i >= 10 ? "fam f-" + FAMILIAS[i - 10].toLowerCase() : "");
            // "setembro 2026" em duas linhas (mês / ano), como o texto quebrado da aba
            const m = /^(\S+)\s+(\d{4})$/.exec(x);
            if (m) { c.appendChild(el("span", "l1", m[1])); c.appendChild(el("span", "l2", m[2])); } else { c.textContent = x; }
            c.setAttribute("scope", "col");
            const grupo = i < 3 ? "Vendido" : i < 6 ? "Faturado" : i === 6 ? "Carteira" : i < 10 ? "Clientes (vendido)" : "Mix (vendido %)";
            c.setAttribute("aria-label", grupo + " " + x);
            hr.appendChild(c);
        });
        th.appendChild(hr); t.appendChild(th);
        const tb = el("tbody");
        if (!this.linhas.length) {
            const tr = el("tr"); const td = el("td", "msg", "Sem representantes para os filtros escolhidos."); td.setAttribute("colspan", String(3 + NCOLS)); tr.appendChild(td); tb.appendChild(tr);
        }
        this.linhas.forEach(l => {
            const tr = el("tr");
            const rc = el("th", "rep", l.nome); rc.setAttribute("scope", "row"); rc.title = l.nome;
            tr.appendChild(rc);
            tr.appendChild(el("td", "uf", l.uf));
            tr.appendChild(el("td", "sup", l.sup));
            for (let i = 0; i < NCOLS; i++) {
                const v = l.v[i], txt = fmt(v, i);
                const cls = i >= 10 ? "mx" + (txt ? "" : " vz") + (i === 10 ? " ini" : i === 14 ? " fim" : "") : "vl";
                tr.appendChild(el("td", cls, txt));
            }
            tb.appendChild(tr);
        });
        t.appendChild(tb);
        area.appendChild(t);
        corpo.appendChild(area);
        if (this.carregando) { corpo.appendChild(el("div", "pv-status", "Carregando mais linhas…")); }
        else if (this.limiteAtingido) { corpo.appendChild(el("div", "pv-status erro", "O Power BI parou de enviar linhas (limite de memória do visual). Use os filtros para ver o restante.")); }
        area.scrollTop = rolagem;
    }

    // ---------- exportar ----------
    private escolhas = { tabela: true, dados: false };
    private menuAberto = false;
    private aviso = "";
    private avisoErro = false;

    private menuExportar(): HTMLElement {
        const cx = el("div", "ex");
        const btn = botao("pv-btn ex-btn", "Exportar ▾", () => { this.menuAberto = !this.menuAberto; this.aviso = ""; if (this.menuAberto) { this.filtros.popAberto = false; } this.desenhar(); });
        btn.setAttribute("aria-haspopup", "true");
        btn.setAttribute("aria-expanded", String(this.menuAberto));
        cx.appendChild(btn);
        if (this.menuAberto) {
            const m = el("div", "ex-menu");
            m.appendChild(el("div", "ex-tit", "Incluir no Excel"));
            const itens: ["tabela" | "dados", string][] = [["tabela", "Tabela (igual à aba " + CONFIG.aba + ")"], ["dados", "Aba com dados brutos"]];
            itens.forEach(([k, rot]) => {
                const lb = el("label", "ex-op");
                const cb = document.createElement("input");
                cb.type = "checkbox"; cb.checked = !!this.escolhas[k];
                cb.addEventListener("change", () => { this.escolhas[k] = cb.checked; });
                lb.appendChild(cb); lb.appendChild(document.createTextNode(" " + rot));
                m.appendChild(lb);
            });
            m.appendChild(botao("ex-go", "Baixar Excel", () => { void this.exportarExcel(); }));
            m.appendChild(botao("ex-copiar", "Copiar tabela (colar no Excel)", () => { void this.copiarTabela(); }));
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

    private entrada(): EntradaPV {
        return {
            titulo: this.settings.titulo.texto.value || "PAINEL DE VENDAS - DIA", aba: CONFIG.aba, regra: CONFIG.regra, meses: this.meses,
            geradoEm: this.agora(), periodo: this.periodo().replace(/ \(Ano, Mês e Dia não mudam este painel\)$/, ""), filtros: this.filtros.resumo(),
            linhas: this.linhas, opcoes: { tabela: this.escolhas.tabela, dados: this.escolhas.dados }
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
        ta.className = "pv-copia";
        this.raiz.appendChild(ta);
        let ok = false;
        try { ta.focus(); ta.select(); ta.setSelectionRange(0, texto.length); ok = document.execCommand("copy"); } catch { ok = false; }
        this.raiz.removeChild(ta);
        return ok;
    }

    private async copiarTabela(): Promise<void> {
        if (!this.temDados || !this.linhas.length) { this.mostrar("Sem dados para copiar.", true); return; }
        this.menuAberto = false;
        const ok = await this.copiar(tsv(this.entrada()));
        this.mostrar(ok ? `Tabela copiada (${fmtN.format(this.linhas.length)} representantes). Abra o Excel e cole com Ctrl+V.` : "O Power BI bloqueou a cópia neste relatório. Use Baixar Excel.", !ok);
    }

    private async exportarExcel(): Promise<void> {
        const o = this.escolhas;
        if (!this.temDados) { this.mostrar("Sem dados para exportar.", true); return; }
        if (!o.tabela && !o.dados) { this.mostrar("Marque pelo menos um item para exportar.", true); return; }
        this.menuAberto = false;
        const ex = this.settings.exportacao;
        const nome = (ex.nomeArquivo.value || CONFIG.arquivo).replace(/[^\w-]+/g, "_") + ".xlsx";
        const url = (ex.urlPagina.value || "").trim();
        const entrada = this.entrada();
        try {
            // 1) download direto pelo Power BI (Desktop e serviço com a exportação liberada pelo administrador): arquivo completo
            let permitido = false;
            try { permitido = (await this.host.downloadService.exportStatus()) === powerbi.PrivilegeStatus.Allowed; } catch { permitido = false; }
            if (permitido) {
                this.mostrar("Gerando Excel…");
                const r = await this.host.downloadService.exportVisualsContentExtended(paraBase64(gerarXlsxPainelVendas(entrada)), nome, "base64", "Exportação do " + CONFIG.nomeVisual + " para Excel");
                this.mostrar(r && (r as { downloadCompleted?: boolean }).downloadCompleted === false ? "Download cancelado." : "Excel exportado: " + nome);
                return;
            }
            // 2) página de download: dados no link quando cabem; senão pela área de transferência
            if (ex.usarPagina.value && /^https:\/\//i.test(url)) { await this.abrirPaginaDownload(url, nome, entrada); return; }
            // 3) sem página: tenta o download do navegador (funciona fora do iframe restrito do Power BI)
            const a = document.createElement("a");
            a.href = URL.createObjectURL(new Blob([gerarXlsxPainelVendas(entrada) as BlobPart], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
            a.download = nome; this.raiz.appendChild(a); a.click(); this.raiz.removeChild(a);
            this.mostrar("Download direto (exportação do Power BI não liberada pelo administrador). Se não baixar, use Copiar tabela.");
        } catch (e) { this.mostrar("Erro ao exportar: " + (e instanceof Error ? e.message : String(e)), true); }
    }

    // carga compacta -> deflate-raw -> base64url
    public static async codificar(entrada: EntradaPV): Promise<{ b64: string; comp: string }> {
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

    private async abrirPaginaDownload(url: string, nome: string, entrada: EntradaPV): Promise<void> {
        const { b64, comp } = await Visual.codificar(entrada);
        const base = url.replace(/#.*$/, "") + "#";
        const alvo = base + new URLSearchParams({ v: "2", n: nome, c: comp, j: b64 }).toString();
        const mil = (n: number): string => (n / 1000).toFixed(1).replace(".", ",");
        const qtd = fmtN.format(this.linhas.length);
        if (alvo.length <= LIMITE_LINK) {
            this.host.launchUrl(alvo);
            this.mostrar(`Página de download aberta: ${nome} (link de ${mil(alvo.length)} mil caracteres)`);
            return;
        }
        // Tabela grande demais para o link: os dados vão pela área de transferência e o link leva só o aviso (k=cola).
        const copiou = comp === "1" && await this.copiar(PREFIXO_COLA + b64);
        if (copiou) {
            this.host.launchUrl(base + new URLSearchParams({ v: "2", n: nome, k: "cola", q: String(this.linhas.length) }).toString());
            this.mostrar(`A tabela (${qtd} representantes) não cabe no link (${mil(alvo.length)} mil caracteres). Os dados foram copiados: na página que abriu, clique em Colar dados.`);
        } else {
            this.mostrar(`A tabela (${qtd} representantes) não cabe no link e o Power BI bloqueou a cópia. Filtre por Supervisão ou Representante e exporte de novo.`, true);
        }
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.fmtService.buildFormattingModel(this.settings);
    }
}
