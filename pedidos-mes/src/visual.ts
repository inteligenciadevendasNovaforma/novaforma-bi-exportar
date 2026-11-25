"use strict";
// Pedidos Mês (NovaForma) v1.0: réplica da aba PEDIDOS MÊS (vendido do mês por representante > cliente > família comercial),
// tabela rolável com cabeçalho fixo; botão Filtros (filtra a página) e Exportar (Excel igual à aba, ou cópia para colar no Excel).
// Dados grandes: uma linha do visual por representante (até 30 mil, com fetchMoreData), cada uma com as linhas cliente x família.
import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import "./../style/visual.less";
import { VisualSettings } from "./settings";
import { Filtros, FiltroBasico, el, botao, dataBR } from "./filtros";
import { lerCarga, Ordem } from "./dados";
import { gerarXlsxPedidos, paraBase64, compactar, tsv, RepPM, EntradaPedidos, CABECALHO } from "./xlsxPedidos";

import IVisual = powerbi.extensibility.visual.IVisual;
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;

const MEDIDA = "Pedidos Mês (visual)";
// O Power BI recusa links longos ("link malformado"); acima disso os dados vão pela área de transferência.
export const LIMITE_LINK = 2500;
// prefixo do texto copiado para a página de download (formato compacto, deflate-raw + base64url)
export const PREFIXO_COLA = "NFPM1.";
const fmtR = new Intl.NumberFormat("pt-BR", { style: "currency", currency: "BRL", minimumFractionDigits: 2, maximumFractionDigits: 2 });
const fmtN = new Intl.NumberFormat("pt-BR");
const LARGURAS = [42.6397, 12.9789, 42.6397, 23.2461, 13.1690333333333]; // colunas C:G da aba (Excel)

export class Visual implements IVisual {
    private host: IVisualHost;
    private raiz: HTMLElement;
    private topo: HTMLElement;
    private corpo: HTMLElement;
    private fmtService = new FormattingSettingsService();
    private settings: VisualSettings = new VisualSettings();
    private filtros: Filtros;
    private reps: RepPM[] = [];
    private nLinhas = 0;
    private temDados = false;
    private carregando = false;
    private limiteAtingido = false;
    private corpoPronto = false;

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.raiz = el("div", "pq");
        this.topo = el("div", "pq-topo");
        this.corpo = el("div", "pq-corpo");
        this.raiz.appendChild(this.topo);
        this.raiz.appendChild(this.corpo);
        options.element.appendChild(this.raiz);
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
            // busca o resto das linhas (representantes) em blocos; o Power BI chama update de novo com tudo junto
            this.carregando = this.host.fetchMoreData(true);
            this.limiteAtingido = !this.carregando;
        } else { this.limiteAtingido = false; }
        this.desenhar();
    }

    private ordem(): Ordem {
        const v = this.settings.tabela.ordem.value;
        const s = String((v && (v as { value?: string }).value) || "vendido");
        return (s === "nome" || s === "codigo" ? s : "vendido") as Ordem;
    }

    private lerDados(dv: powerbi.DataView | undefined): void {
        this.temDados = false; this.reps = []; this.nLinhas = 0;
        const t = dv && dv.table;
        if (!t || !t.rows.length) { return; }
        let idx = t.columns.findIndex(c => c.roles && c.roles["dados"]);
        if (idx < 0) { idx = t.columns.length - 1; }
        const celulas = t.rows.map(r => r[idx]).filter(v => typeof v === "string" && v) as string[];
        if (!celulas.length) { return; }
        const c = lerCarga(celulas, this.ordem());
        if (c.cabecalho.length) { this.filtros.lerCabecalho(c.cabecalho, 0); }
        this.reps = c.reps;
        this.nLinhas = c.linhas;
        this.temDados = true;
    }

    private periodo(): string {
        const r = this.filtros.dataRef;
        return r ? `Período: ${dataBR(r.slice(0, 8) + "01")} a ${dataBR(r)}` : "";
    }

    // ---------- tela ----------
    private desenhar(): void {
        const s = this.settings;
        this.raiz.style.background = s.fundo.mostrar.value ? s.fundo.cor.value.value : "transparent";
        // topo (título e botões): refeito a cada interação
        const topo = this.topo;
        while (topo.firstChild) { topo.removeChild(topo.firstChild); }
        if (s.titulo.mostrar.value) { topo.appendChild(el("h2", "pq-titulo", s.titulo.texto.value || "PEDIDOS MÊS")); }
        else { topo.appendChild(el("div", "pq-espaco")); }
        const acoes = el("div", "pq-acoes");
        if (s.filtros.mostrarBotao.value) { acoes.appendChild(this.filtros.botao("pq-btn", () => { this.menuAberto = false; })); }
        if (s.exportacao.mostrarBotao.value) { acoes.appendChild(this.menuExportar()); }
        topo.appendChild(acoes);
        // corpo (tabela): só é refeito quando os dados ou a ordem mudam (mantém a rolagem e fica leve com milhares de linhas)
        if (!this.corpoPronto) { this.corpoPronto = true; this.desenharCorpo(); }
        const pop = this.raiz.querySelector(".fp");
        if (pop) { this.raiz.removeChild(pop); }
        if (this.filtros.popAberto) { this.raiz.appendChild(this.filtros.painel()); }
    }

    private desenharCorpo(): void {
        const corpo = this.corpo;
        const rolagem = corpo.scrollTop;
        while (corpo.firstChild) { corpo.removeChild(corpo.firstChild); }
        if (!this.temDados) {
            corpo.appendChild(el("div", "pq-vazio", this.filtros.ativos() ? "Sem dados para os filtros escolhidos." : `Arraste o código do representante e a medida ${MEDIDA} para os campos do visual.`));
            return;
        }
        const area = el("div", "pq-tw");
        area.setAttribute("role", "region");
        area.setAttribute("tabindex", "0");
        area.setAttribute("aria-label", `Pedidos do mês: ${fmtN.format(this.nLinhas)} linhas de ${fmtN.format(this.reps.length)} representantes (role para ver todas)`);
        const t = el("table", "pq-tab");
        const cg = el("colgroup");
        const soma = LARGURAS.reduce((a, b) => a + b, 0);
        LARGURAS.forEach(w => { const c = el("col"); c.style.width = (w / soma * 100).toFixed(3) + "%"; cg.appendChild(c); });
        t.appendChild(cg);
        const th = el("thead"), hr = el("tr");
        CABECALHO.forEach((h, i) => { const c = el("th", i === 4 ? "vl" : "", h); c.setAttribute("scope", "col"); hr.appendChild(c); });
        th.appendChild(hr); t.appendChild(th);
        const tb = el("tbody");
        if (!this.reps.length) {
            const tr = el("tr"); const td = el("td", "msg", "Sem pedidos para os filtros escolhidos."); td.setAttribute("colspan", "5"); tr.appendChild(td); tb.appendChild(tr);
        }
        this.reps.forEach(rep => {
            let primeira = true;
            rep.clis.forEach(cli => cli.fams.forEach((f, i) => {
                const tr = el("tr", primeira ? "ini" : "");
                const rc = el(primeira ? "th" : "td", "rep", primeira ? rep.nome : "");
                if (primeira) { rc.setAttribute("scope", "rowgroup"); rc.title = rep.nome; }
                tr.appendChild(rc);
                tr.appendChild(el("td", "cod", i === 0 ? cli.cod : ""));
                const rz = el("td", "raz", i === 0 ? cli.razao : "");
                if (i === 0) { rz.title = cli.razao; }
                tr.appendChild(rz);
                tr.appendChild(el("td", "fam", f.nome));
                tr.appendChild(el("td", "vl", fmtR.format(Math.round(f.v * 100) / 100)));
                tb.appendChild(tr);
                primeira = false;
            }));
            const sep = el("tr", "sep"); const td = el("td"); td.setAttribute("colspan", "5"); td.setAttribute("aria-hidden", "true"); sep.appendChild(td); tb.appendChild(sep);
        });
        t.appendChild(tb);
        area.appendChild(t);
        corpo.appendChild(area);
        if (this.carregando) { corpo.appendChild(el("div", "pq-status", "Carregando mais linhas…")); }
        else if (this.limiteAtingido) { corpo.appendChild(el("div", "pq-status erro", "O Power BI parou de enviar linhas (limite de memória do visual). Use os filtros para ver o restante.")); }
        corpo.scrollTop = rolagem;
    }

    // ---------- exportar ----------
    private escolhas = { tabela: true, dados: false };
    private menuAberto = false;
    private aviso = "";
    private avisoErro = false;

    private menuExportar(): HTMLElement {
        const cx = el("div", "ex");
        const btn = botao("pq-btn ex-btn", "Exportar ▾", () => { this.menuAberto = !this.menuAberto; this.aviso = ""; if (this.menuAberto) { this.filtros.popAberto = false; } this.desenhar(); });
        btn.setAttribute("aria-haspopup", "true");
        btn.setAttribute("aria-expanded", String(this.menuAberto));
        cx.appendChild(btn);
        if (this.menuAberto) {
            const m = el("div", "ex-menu");
            m.appendChild(el("div", "ex-tit", "Incluir no Excel"));
            const itens: ["tabela" | "dados", string][] = [["tabela", "Tabela (igual à aba PEDIDOS MÊS)"], ["dados", "Aba com dados brutos"]];
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

    private entrada(): EntradaPedidos {
        return {
            titulo: this.settings.titulo.texto.value || "PEDIDOS MÊS", geradoEm: this.agora(), periodo: this.periodo(), filtros: this.filtros.resumo(),
            reps: this.reps, opcoes: { tabela: this.escolhas.tabela, dados: this.escolhas.dados }
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
        ta.className = "pq-copia";
        this.raiz.appendChild(ta);
        let ok = false;
        try { ta.focus(); ta.select(); ta.setSelectionRange(0, texto.length); ok = document.execCommand("copy"); } catch { ok = false; }
        this.raiz.removeChild(ta);
        return ok;
    }

    private async copiarTabela(): Promise<void> {
        if (!this.temDados || !this.reps.length) { this.mostrar("Sem dados para copiar.", true); return; }
        this.menuAberto = false;
        const ok = await this.copiar(tsv(this.entrada()));
        this.mostrar(ok ? `Tabela copiada (${fmtN.format(this.nLinhas)} linhas). Abra o Excel e cole com Ctrl+V.` : "O Power BI bloqueou a cópia neste relatório. Use Baixar Excel.", !ok);
    }

    private async exportarExcel(): Promise<void> {
        const o = this.escolhas;
        if (!this.temDados) { this.mostrar("Sem dados para exportar.", true); return; }
        if (!o.tabela && !o.dados) { this.mostrar("Marque pelo menos um item para exportar.", true); return; }
        this.menuAberto = false;
        const ex = this.settings.exportacao;
        const nome = (ex.nomeArquivo.value || "pedidos_mes").replace(/[^\w-]+/g, "_") + ".xlsx";
        const url = (ex.urlPagina.value || "").trim();
        const entrada = this.entrada();
        try {
            // 1) download direto pelo Power BI (Desktop e serviço com a exportação liberada pelo administrador): arquivo completo
            let permitido = false;
            try { permitido = (await this.host.downloadService.exportStatus()) === powerbi.PrivilegeStatus.Allowed; } catch { permitido = false; }
            if (permitido) {
                this.mostrar("Gerando Excel…");
                const r = await this.host.downloadService.exportVisualsContentExtended(paraBase64(gerarXlsxPedidos(entrada)), nome, "base64", "Exportação do Pedidos Mês para Excel");
                this.mostrar(r && (r as { downloadCompleted?: boolean }).downloadCompleted === false ? "Download cancelado." : "Excel exportado: " + nome);
                return;
            }
            // 2) página de download: dados no link quando cabem; senão pela área de transferência
            if (ex.usarPagina.value && /^https:\/\//i.test(url)) { await this.abrirPaginaDownload(url, nome, entrada); return; }
            // 3) sem página: tenta o download do navegador (funciona fora do iframe restrito do Power BI)
            const a = document.createElement("a");
            a.href = URL.createObjectURL(new Blob([gerarXlsxPedidos(entrada) as BlobPart], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
            a.download = nome; this.raiz.appendChild(a); a.click(); this.raiz.removeChild(a);
            this.mostrar("Download direto (exportação do Power BI não liberada pelo administrador). Se não baixar, use Copiar tabela.");
        } catch (e) { this.mostrar("Erro ao exportar: " + (e instanceof Error ? e.message : String(e)), true); }
    }

    // carga compacta -> deflate-raw -> base64url
    public static async codificar(entrada: EntradaPedidos): Promise<{ b64: string; comp: string }> {
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

    private async abrirPaginaDownload(url: string, nome: string, entrada: EntradaPedidos): Promise<void> {
        const { b64, comp } = await Visual.codificar(entrada);
        const base = url.replace(/#.*$/, "") + "#";
        const alvo = base + new URLSearchParams({ v: "2", n: nome, c: comp, j: b64 }).toString();
        const mil = (n: number): string => (n / 1000).toFixed(1).replace(".", ",");
        if (alvo.length <= LIMITE_LINK) {
            this.host.launchUrl(alvo);
            this.mostrar(`Página de download aberta: ${nome} (link de ${mil(alvo.length)} mil caracteres)`);
            return;
        }
        // Tabela grande demais para o link: os dados vão pela área de transferência e o link leva só o aviso (k=cola).
        const copiou = comp === "1" && await this.copiar(PREFIXO_COLA + b64);
        if (copiou) {
            this.host.launchUrl(base + new URLSearchParams({ v: "2", n: nome, k: "cola", q: String(this.nLinhas) }).toString());
            this.mostrar(`A tabela (${fmtN.format(this.nLinhas)} linhas) não cabe no link (${mil(alvo.length)} mil caracteres). Os dados foram copiados: na página que abriu, clique em Colar dados.`);
        } else {
            this.mostrar(`A tabela (${fmtN.format(this.nLinhas)} linhas) não cabe no link e o Power BI bloqueou a cópia. Filtre por Supervisão ou Representante e exporte de novo.`, true);
        }
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.fmtService.buildFormattingModel(this.settings);
    }
}
