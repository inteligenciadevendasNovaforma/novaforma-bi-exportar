"use strict";
import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import "./../style/visual.less";
import { VisualSettings, COLUNAS_PADRAO } from "./settings";
import { gerarXlsxReps, paraBase64, OpcoesReps, LinhaReps } from "./xlsxReps";

import IVisual = powerbi.extensibility.visual.IVisual;
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;

type ChaveColuna = keyof typeof COLUNAS_PADRAO;
interface Campo { id: string; rotulo: string; col: ChaveColuna; num: boolean; lista: number; }
// lista = posição da lista de valores no cabeçalho da medida (depois da data de referência)
const CAMPOS: Campo[] = [
    { id: "ano", rotulo: "Ano", col: "colAno", num: true, lista: 0 },
    { id: "mes", rotulo: "Mês", col: "colMes", num: false, lista: 1 },
    { id: "dia", rotulo: "Dia", col: "colDia", num: true, lista: 2 },
    { id: "sup", rotulo: "Supervisão", col: "colSup", num: false, lista: 3 },
    { id: "rep", rotulo: "Representante", col: "colRep", num: false, lista: 4 },
    { id: "fam", rotulo: "Família", col: "colFam", num: false, lista: 6 },
    { id: "mat", rotulo: "Material", col: "colMat", num: false, lista: 7 },
    { id: "uf", rotulo: "UF", col: "colUf", num: false, lista: 5 }
];
const MAX_ITENS = 400;
const NAO_VENDEU = "NÃO VENDEU";

interface Alvo { table: string; column: string; }
interface FiltroBasico { $schema: string; target: Alvo; filterType: number; operator: string; values: (string | number)[]; requireSingleSelection: boolean; }

function el(tag: string, cls?: string, texto?: string): HTMLElement {
    const e = document.createElement(tag);
    if (cls) { e.className = cls; }
    if (texto !== undefined) { e.textContent = texto; }
    return e;
}
function botao(cls: string, texto: string, acao: () => void): HTMLButtonElement {
    const b = el("button", cls, texto) as HTMLButtonElement;
    b.type = "button";
    b.addEventListener("click", acao);
    return b;
}
function semAcento(s: string): string {
    return s.normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase();
}
function dataBR(iso: string): string {
    const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso || "");
    return m ? `${m[3]}/${m[2]}/${m[1]}` : "";
}
// "Tabela[Coluna]; 'Outra Tabela'[Coluna]" -> alvos do filtro
export function lerAlvos(texto: string): Alvo[] {
    return (texto || "").split(";").map(p => /^\s*'?(.+?)'?\s*\[(.+)\]\s*$/.exec(p)).filter(m => !!m).map(m => ({ table: (m as RegExpExecArray)[1].trim(), column: (m as RegExpExecArray)[2].trim() }));
}

export class Visual implements IVisual {
    private host: IVisualHost;
    private raiz: HTMLElement;
    private fmtService = new FormattingSettingsService();
    private settings: VisualSettings = new VisualSettings();
    private temDados = false;
    private linhas: LinhaReps[] = [];
    private totalReps = 0;
    private dataRef = "";
    private listas = new Map<string, string[]>();
    private aplicados = new Map<string, Set<string>>();
    private pendentes = new Map<string, Set<string>>();
    private busca = new Map<string, string>();
    private popAberto = false;
    private campoAtivo = "ano";
    private btnFiltros: HTMLButtonElement | null = null;

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.raiz = el("div", "rs");
        options.element.appendChild(this.raiz);
        this.raiz.addEventListener("keydown", (ev: KeyboardEvent) => {
            if (ev.key !== "Escape") { return; }
            if (this.popAberto) { this.fecharPainel(); ev.stopPropagation(); }
            else if (this.menuAberto) { this.menuAberto = false; this.desenhar(); }
        });
    }

    public update(options: VisualUpdateOptions): void {
        const dv = options.dataViews && options.dataViews[0];
        this.settings = this.fmtService.populateFormattingSettingsModel(VisualSettings, dv);
        this.lerDados(dv);
        if (Array.isArray(options.jsonFilters)) { this.lerFiltros(options.jsonFilters as unknown as FiltroBasico[]); }
        if (dv && dv.metadata && dv.metadata.segment) { this.host.fetchMoreData(true); }
        this.desenhar();
    }

    private lerDados(dv: powerbi.DataView | undefined): void {
        this.linhas = []; this.totalReps = 0; this.temDados = false;
        const t = dv && dv.table;
        if (!t) { return; }
        let idx = t.columns.findIndex(c => c.roles && c.roles["dados"]);
        if (idx < 0) { idx = t.columns.length - 1; }
        for (const r of t.rows) {
            const bruto = r[idx];
            if (typeof bruto !== "string" || !bruto) { continue; }
            this.temDados = true;
            const partes = bruto.split("~");
            if (partes[1] !== undefined) {
                const h = partes[1].split("|");
                this.dataRef = h[0] || "";
                CAMPOS.forEach(cp => {
                    const v = (h[1 + cp.lista] || "").split("^").filter(x => x !== "");
                    this.listas.set(cp.id, v);
                });
            }
            const c = partes[0].split("|");
            if (c.length < 4) { continue; }
            this.totalReps++;
            const cli = Number(c[3]);
            if (c[3] === "" || !(cli > 0)) { this.linhas.push({ nome: c[0], uf: c[1], sup: c[2] }); }
        }
        this.linhas.sort((a, b) => a.nome.localeCompare(b.nome, "pt-BR"));
    }

    // filtros vindos do Power BI (marcador, reabertura do relatório): reconstrói a seleção aplicada
    private lerFiltros(fs: FiltroBasico[]): void {
        const novo = new Map<string, Set<string>>();
        for (const f of fs) {
            if (!f || !f.target || !Array.isArray(f.values)) { continue; }
            const alvo = f.target;
            const cp = CAMPOS.find(c => this.alvos(c).some(a => a.table === alvo.table && a.column === alvo.column));
            if (!cp) { continue; }
            const s = novo.get(cp.id) || new Set<string>();
            f.values.forEach(v => s.add(String(v)));
            novo.set(cp.id, s);
        }
        this.aplicados = novo;
    }

    private alvos(cp: Campo): Alvo[] {
        const txt = (this.settings.filtros[cp.col] as { value: string }).value;
        const a = lerAlvos(txt);
        return a.length ? a : lerAlvos(COLUNAS_PADRAO[cp.col]);
    }

    private filtrosAtivos(): number {
        let n = 0;
        this.aplicados.forEach(s => { if (s.size) { n++; } });
        return n;
    }

    // ---------- tela ----------
    private desenhar(): void {
        const raiz = this.raiz;
        while (raiz.firstChild) { raiz.removeChild(raiz.firstChild); }
        raiz.style.background = this.settings.fundo.mostrar.value ? this.settings.fundo.cor.value.value : "transparent";
        const topo = el("div", "rs-topo");
        if (this.settings.titulo.mostrar.value) { topo.appendChild(el("div", "rs-titulo", this.settings.titulo.texto.value)); }
        const acoes = el("div", "rs-acoes");
        this.btnFiltros = null;
        if (this.settings.filtros.mostrarBotao.value) { acoes.appendChild(this.botaoFiltros()); }
        if (this.settings.exportacao.mostrarBotao.value) { acoes.appendChild(this.menuExportar()); }
        topo.appendChild(acoes);
        raiz.appendChild(topo);
        if (!this.temDados) {
            raiz.appendChild(el("div", "rs-vazio", "Arraste o campo Representante (d_Representantes[Nome do Rep]) e a medida Reps Sem Pedidos (visual) para o visual."));
        } else {
            raiz.appendChild(this.tabela());
        }
        if (this.popAberto) { raiz.appendChild(this.painelFiltros()); }
    }

    private tabela(): HTMLElement {
        const wrap = el("div", "rs-tab");
        const t = el("table");
        t.setAttribute("aria-label", (this.settings.titulo.texto.value || "Reps sem pedidos") + ": " + this.linhas.length + " representantes sem pedido");
        const cg = el("colgroup");
        [46, 10.5, 15.5, 28].forEach(w => { const c = el("col"); c.style.width = w + "%"; cg.appendChild(c); });
        t.appendChild(cg);
        const th = el("thead");
        const f1 = el("tr", "faixa");
        const fx = el("th", "", this.settings.titulo.faixa.value || "POSITIVAÇÃO");
        fx.setAttribute("colspan", "4"); fx.setAttribute("scope", "colgroup");
        f1.appendChild(fx); th.appendChild(f1);
        const f2 = el("tr", "cols");
        ["REPRESENTANTES", "UF", "SUP", "CLI MÊS"].forEach(x => { const h = el("th", "", x); h.setAttribute("scope", "col"); f2.appendChild(h); });
        th.appendChild(f2); t.appendChild(th);
        const tb = el("tbody");
        if (!this.linhas.length) {
            const tr = el("tr", "rs-ok");
            const td = el("td", "", "Todos os representantes venderam no período.");
            td.setAttribute("colspan", "4"); tr.appendChild(td); tb.appendChild(tr);
        }
        this.linhas.forEach(l => {
            const tr = el("tr");
            tr.appendChild(el("td", "nm", l.nome));
            tr.appendChild(el("td", "uf", l.uf));
            tr.appendChild(el("td", "sp", l.sup));
            tr.appendChild(el("td", "nv", NAO_VENDEU));
            tb.appendChild(tr);
        });
        t.appendChild(tb);
        wrap.appendChild(t);
        return wrap;
    }

    // ---------- filtros ----------
    private botaoFiltros(): HTMLElement {
        const n = this.filtrosAtivos();
        const b = botao("rs-btn fl-btn" + (n ? " ativo" : ""), "Filtros", () => {
            if (this.popAberto) { this.fecharPainel(); return; }
            this.menuAberto = false;
            this.pendentes = new Map();
            this.aplicados.forEach((s, k) => this.pendentes.set(k, new Set(s)));
            this.popAberto = true;
            this.desenhar();
            const busca = this.raiz.querySelector(".fp-busca") as HTMLInputElement | null;
            if (busca) { busca.focus(); }
        });
        b.setAttribute("aria-haspopup", "dialog");
        b.setAttribute("aria-expanded", String(this.popAberto));
        if (n) {
            b.appendChild(el("span", "fl-n", String(n)));
            b.setAttribute("aria-label", `Filtros (${n} ativo${n > 1 ? "s" : ""})`);
        }
        this.btnFiltros = b;
        return b;
    }

    private fecharPainel(): void {
        this.popAberto = false;
        this.desenhar();
        if (this.btnFiltros) { this.btnFiltros.focus(); }
    }

    private qtdPendente(id: string): number { const s = this.pendentes.get(id); return s ? s.size : 0; }

    private painelFiltros(): HTMLElement {
        const p = el("div", "fp");
        p.setAttribute("role", "dialog");
        p.setAttribute("aria-label", "Filtros da página");
        const cab = el("div", "fp-cab");
        cab.appendChild(el("div", "fp-tit", "Filtros"));
        cab.appendChild(botao("fp-link", "Limpar tudo", () => { this.pendentes = new Map(); this.aplicar(); }));
        const x = botao("fp-x", "×", () => this.fecharPainel());
        x.setAttribute("aria-label", "Fechar filtros");
        cab.appendChild(x);
        p.appendChild(cab);

        const corpo = el("div", "fp-corpo");
        const nav = el("div", "fp-nav");
        nav.setAttribute("role", "tablist");
        nav.setAttribute("aria-label", "Campos");
        CAMPOS.forEach(cp => {
            const b = botao("fp-campo" + (cp.id === this.campoAtivo ? " sel" : ""), cp.rotulo, () => {
                this.campoAtivo = cp.id; this.desenhar();
                const s = this.raiz.querySelector(".fp-busca") as HTMLInputElement | null;
                if (s) { s.focus(); }
            });
            b.setAttribute("role", "tab");
            b.setAttribute("aria-selected", String(cp.id === this.campoAtivo));
            b.dataset.campo = cp.id;
            const q = this.qtdPendente(cp.id);
            const badge = el("span", "fp-q" + (q ? "" : " zero"), q ? String(q) : "");
            b.appendChild(badge);
            nav.appendChild(b);
        });
        corpo.appendChild(nav);
        corpo.appendChild(this.painelCampo(CAMPOS.find(c => c.id === this.campoAtivo) || CAMPOS[0], nav));
        p.appendChild(corpo);

        const rod = el("div", "fp-rod");
        const ini = this.dataRef ? dataBR(this.dataRef.slice(0, 8) + "01") : "";
        rod.appendChild(el("div", "fp-per", this.dataRef ? `Período do CLI MÊS: ${ini} a ${dataBR(this.dataRef)}` : ""));
        rod.appendChild(botao("rs-btn fp-aplicar", "Aplicar", () => this.aplicar()));
        p.appendChild(rod);
        return p;
    }

    private painelCampo(cp: Campo, nav: HTMLElement): HTMLElement {
        const box = el("div", "fp-pane");
        box.setAttribute("role", "tabpanel");
        const todos = this.listas.get(cp.id) || [];
        const sel = this.pendentes.get(cp.id) || new Set<string>();
        this.pendentes.set(cp.id, sel);
        // valores selecionados que não estão na lista atual continuam visíveis (para poder desmarcar)
        const valores = todos.slice();
        sel.forEach(v => { if (valores.indexOf(v) < 0) { valores.push(v); } });

        const topo = el("div", "fp-pane-top");
        const busca = document.createElement("input");
        busca.type = "search"; busca.className = "fp-busca";
        busca.placeholder = "Buscar " + cp.rotulo.toLowerCase() + "…";
        busca.setAttribute("aria-label", "Buscar em " + cp.rotulo);
        busca.value = this.busca.get(cp.id) || "";
        topo.appendChild(busca);
        box.appendChild(topo);

        const info = el("div", "fp-info");
        const lista = el("div", "fp-lista");
        lista.setAttribute("role", "group");
        lista.setAttribute("aria-label", cp.rotulo);
        const atualizarContagem = (): void => {
            const q = sel.size;
            const b = nav.querySelector(`[data-campo="${cp.id}"] .fp-q`) as HTMLElement | null;
            if (b) { b.textContent = q ? String(q) : ""; b.className = "fp-q" + (q ? "" : " zero"); }
            info.textContent = (q ? `${q} selecionado${q > 1 ? "s" : ""}` : "Nenhum selecionado (sem filtro)") + ` · ${todos.length} valores`;
        };
        let visiveis: string[] = [];
        const desenharLista = (): void => {
            while (lista.firstChild) { lista.removeChild(lista.firstChild); }
            const termo = semAcento(busca.value.trim());
            visiveis = termo ? valores.filter(v => semAcento(v).indexOf(termo) >= 0) : valores;
            if (!valores.length) { lista.appendChild(el("div", "fp-msg", "Sem valores: confira a medida Reps Sem Pedidos (visual).")); return; }
            if (!visiveis.length) { lista.appendChild(el("div", "fp-msg", "Nada encontrado.")); return; }
            visiveis.slice(0, MAX_ITENS).forEach(v => {
                const lb = el("label", "fp-op");
                const cb = document.createElement("input");
                cb.type = "checkbox"; cb.checked = sel.has(v);
                cb.addEventListener("change", () => { if (cb.checked) { sel.add(v); } else { sel.delete(v); } atualizarContagem(); });
                lb.appendChild(cb);
                lb.appendChild(el("span", "", v));
                lista.appendChild(lb);
            });
            if (visiveis.length > MAX_ITENS) { lista.appendChild(el("div", "fp-msg", `Mostrando ${MAX_ITENS} de ${visiveis.length}. Use a busca para achar os demais.`)); }
        };
        busca.addEventListener("input", () => { this.busca.set(cp.id, busca.value); desenharLista(); });
        const marcar = el("div", "fp-marcar");
        marcar.appendChild(botao("fp-link", "Marcar visíveis", () => { visiveis.forEach(v => sel.add(v)); desenharLista(); atualizarContagem(); }));
        marcar.appendChild(botao("fp-link", "Limpar", () => { sel.clear(); desenharLista(); atualizarContagem(); }));
        box.appendChild(marcar);
        box.appendChild(lista);
        box.appendChild(info);
        desenharLista();
        atualizarContagem();
        return box;
    }

    public montarFiltros(): FiltroBasico[] {
        const saida: FiltroBasico[] = [];
        CAMPOS.forEach(cp => {
            const s = this.aplicados.get(cp.id);
            if (!s || !s.size) { return; }
            const valores = Array.from(s).map(v => (cp.num && !isNaN(Number(v)) ? Number(v) : v));
            this.alvos(cp).forEach(a => saida.push({
                $schema: "https://powerbi.com/product/schema#basic", target: { table: a.table, column: a.column },
                filterType: 1, operator: "In", values: valores, requireSingleSelection: false
            }));
        });
        return saida;
    }

    private aplicar(): void {
        this.aplicados = new Map();
        this.pendentes.forEach((s, k) => { if (s.size) { this.aplicados.set(k, new Set(s)); } });
        const fs = this.montarFiltros();
        const carga = (fs.length ? fs : null) as unknown as powerbi.IFilter[];
        // "filter" filtra os outros visuais da página (como uma segmentação); "selfFilter" filtra este visual
        this.host.applyJsonFilter(carga, "general", "filter", powerbi.FilterAction.merge);
        try {
            this.host.applyJsonFilter(this.settings.filtros.filtrarVisual.value ? carga : null as unknown as powerbi.IFilter[], "general", "selfFilter", powerbi.FilterAction.merge);
        } catch { /* host sem selfFilter */ }
        this.popAberto = false;
        this.desenhar();
        if (this.btnFiltros) { this.btnFiltros.focus(); }
    }

    private resumoFiltros(): string {
        const p: string[] = [];
        CAMPOS.forEach(cp => {
            const s = this.aplicados.get(cp.id);
            if (!s || !s.size) { return; }
            const v = Array.from(s);
            p.push(cp.rotulo + " = " + (v.length > 4 ? v.slice(0, 3).join(", ") + ` e mais ${v.length - 3}` : v.join(", ")));
        });
        return p.length ? "Filtros: " + p.join("; ") : "";
    }

    // ---------- exportar ----------
    private escolhas: OpcoesReps = { tabela: true, dados: false };
    private menuAberto = false;
    private aviso = "";
    private avisoErro = false;

    private menuExportar(): HTMLElement {
        const cx = el("div", "ex");
        const btn = botao("rs-btn ex-btn", "Exportar ▾", () => { this.menuAberto = !this.menuAberto; if (this.menuAberto) { this.popAberto = false; } this.desenhar(); });
        btn.setAttribute("aria-haspopup", "true");
        btn.setAttribute("aria-expanded", String(this.menuAberto));
        cx.appendChild(btn);
        if (this.menuAberto) {
            const m = el("div", "ex-menu");
            m.appendChild(el("div", "ex-tit", "Incluir no Excel"));
            const itens: [keyof OpcoesReps, string][] = [["tabela", "Tabela (células editáveis)"], ["dados", "Aba com dados brutos"]];
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

    private periodo(): string {
        return this.dataRef ? `Período do CLI MÊS: ${dataBR(this.dataRef.slice(0, 8) + "01")} a ${dataBR(this.dataRef)}` : "";
    }

    private async exportarExcel(): Promise<void> {
        const o = this.escolhas;
        if (!this.temDados) { this.mostrar("Sem dados para exportar.", true); return; }
        if (!o.tabela && !o.dados) { this.mostrar("Marque pelo menos um item para exportar.", true); return; }
        this.menuAberto = false;
        const ex = this.settings.exportacao;
        const nome = (ex.nomeArquivo.value || "reps_sem_pedidos").replace(/[^\w-]+/g, "_") + ".xlsx";
        const url = (ex.urlPagina.value || "").trim();
        const entrada = {
            titulo: this.settings.titulo.texto.value || "Reps Sem Pedidos - Mês", faixa: this.settings.titulo.faixa.value || "POSITIVAÇÃO",
            geradoEm: this.agora(), periodo: this.periodo(), filtros: this.resumoFiltros(), linhas: this.linhas, opcoes: { tabela: o.tabela, dados: o.dados }
        };
        try {
            if (ex.usarPagina.value && /^https:\/\//i.test(url)) { await this.abrirPaginaDownload(url, nome, entrada); return; }
            this.mostrar("Gerando Excel…");
            const bytes = gerarXlsxReps(entrada);
            let permitido = false;
            try { permitido = (await this.host.downloadService.exportStatus()) === powerbi.PrivilegeStatus.Allowed; } catch { permitido = false; }
            if (permitido) {
                const r = await this.host.downloadService.exportVisualsContentExtended(paraBase64(bytes), nome, "base64", "Exportação do Reps Sem Pedidos para Excel");
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
    private async abrirPaginaDownload(url: string, nome: string, entrada: object): Promise<void> {
        const carga = { tipo: "repsSemPedidos", ...entrada };
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
