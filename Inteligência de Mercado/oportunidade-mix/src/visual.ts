"use strict";
// Oportunidade de Mix (NovaForma) v1.0.1.0 — 10º visual do Desempenho Venda.
// Réplica (desenho parecido, padrão NovaForma) do dashboard "Oportunidade de Mix por Representante": 5 colunas por família (FB, FC, PT, AP, CT)
// com os representantes SEM pedido válido da família no período; filtros de período (atalhos + intervalo), Supervisão, UF e busca;
// ordenação A–Z, UF e maior valor vendido nos últimos 30 dias; ficha do representante; "Sobre a base"; alerta de cobertura;
// Exportar só Excel pela página de download (link ou área de transferência + #k=cola). Nunca tenta download direto.
// Dados: medida "Oportunidade Mix (visual)" (arquivo Oportunidade_Mix_medida.dax), uma célula com todas as linhas + cabeçalho depois de "~".
// Período: o visual aplica um filtro de intervalo na coluna de data (d_Calendario[Date]); sem filtro a medida usa "Neste mês" (dia 1 até hoje).
import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import "./../style/visual.less";
import { VisualSettings, COLUNA_DATA } from "./settings";
import { lerCarga, conferir, Cabecalho, Formato } from "./dados";
import { ATALHOS, Preset, rotuloPreset, hojeSP, intervalo, validar, dataBR } from "./periodo";
import { FAMILIAS, RepMix, EntradaMix, compactar, paraBase64, diasSem, mixDe, PREFIXO_COLA } from "./xlsxMix";

import IVisual = powerbi.extensibility.visual.IVisual;
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;

const MEDIDA = "Oportunidade Mix (visual)";
const ARQUIVO_DAX = "Oportunidade_Mix_medida.dax";
// O Power BI recusa links longos ("link malformado"); acima disso os dados vão pela área de transferência.
export const LIMITE_LINK = 2500;
const MAX_ITENS = 400;
type Ordem = "nome" | "uf" | "valor";
const ORDENS: { id: Ordem; rotulo: string; dica: string; texto: string }[] = [
    { id: "nome", rotulo: "A–Z", dica: "Ordem alfabética do nome", texto: "Nome (A–Z)" },
    { id: "uf", rotulo: "UF", dica: "UF e depois o nome", texto: "UF e nome" },
    { id: "valor", rotulo: "Maior valor (30 dias)", dica: "Maior valor vendido nas cinco famílias nos últimos 30 dias até o fim do período", texto: "Maior valor vendido nos últimos 30 dias" }
];
const RELATIVOS: Preset[] = ["hoje", "ontem", "7d", "30d", "mesPassado"];

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
// SVG fixo (logo) convertido em nós sem innerHTML
function setHtml(alvo: Element, html: string): void {
    const body = new DOMParser().parseFromString(html, "text/html").body;
    while (alvo.firstChild) { alvo.removeChild(alvo.firstChild); }
    Array.from(body.childNodes).forEach(n => alvo.appendChild(document.importNode(n, true)));
}
const semAcento = (s: string): string => s.normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase();
const fmtBRL = new Intl.NumberFormat("pt-BR", { style: "currency", currency: "BRL" });
const fmtInt = new Intl.NumberFormat("pt-BR", { maximumFractionDigits: 0 });
const colador = new Intl.Collator("pt-BR", { sensitivity: "base", numeric: true });
interface Alvo { table: string; column: string; }
function lerAlvo(texto: string): Alvo | null {
    const m = /^\s*'?(.+?)'?\s*\[(.+)\]\s*$/.exec(texto || "");
    return m ? { table: m[1].trim(), column: m[2].trim() } : null;
}

export class Visual implements IVisual {
    private host: IVisualHost;
    private raiz: HTMLElement;
    private barra: HTMLElement;
    private sub: HTMLElement;
    private quadro: HTMLElement;
    private fmtService = new FormattingSettingsService();
    private settings: VisualSettings = new VisualSettings();
    private reps: RepMix[] = [];
    private cab: Cabecalho | null = null;
    private formato: Formato = "ok";
    private temDados = false;
    private carregando = true;
    private pendente: { ini: string; fim: string; preset: Preset } | null = null;
    private presetSalvo: Preset | "" = "";
    private iniSalvo = "";
    private fimSalvo = "";
    private reaplicado = "";
    // filtros locais (recortes do conjunto entregue pela medida)
    private sup = new Set<string>();
    private uf = new Set<string>();
    private busca = "";
    private ordem: Ordem = "nome";
    // painel Filtros
    private popAberto = false;
    private campoAtivo: "per" | "sup" | "uf" = "per";
    private pSup = new Set<string>();
    private pUf = new Set<string>();
    private pPreset: Preset = "mes";
    private pIni = "";
    private pFim = "";
    private buscaLista = new Map<string, string>();
    private btnFiltros: HTMLButtonElement | null = null;
    // diálogos
    private ficha: { cod: string; fam: number } | null = null;
    private sobre = false;
    private focoAntes: HTMLElement | null = null;
    // exportar
    private menuAberto = false;
    private aviso = "";
    private avisoErro = false;
    private timerAviso: number | undefined;
    private ocupado = false;

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.raiz = el("div", "om");
        this.barra = el("div", "om-barra");
        this.sub = el("div", "om-sub");
        this.quadro = el("div", "om-quadro");
        [this.barra, this.sub, this.quadro].forEach(x => this.raiz.appendChild(x));
        options.element.appendChild(this.raiz);
        const car = el("div", "om-vazio", "Carregando…");
        car.setAttribute("role", "status");
        car.setAttribute("aria-live", "polite");
        this.quadro.appendChild(car);
        this.raiz.addEventListener("keydown", (ev: KeyboardEvent) => {
            if (ev.key !== "Escape") { return; }
            if (this.ficha || this.sobre) { this.fecharDialogo(); ev.stopPropagation(); }
            else if (this.popAberto) { this.fecharPainel(); ev.stopPropagation(); }
            else if (this.menuAberto) { this.menuAberto = false; this.desenharBarra(); ev.stopPropagation(); }
        });
    }

    // ---------------- dados ----------------
    public update(options: VisualUpdateOptions): void {
        try { this.host.eventService.renderingStarted(options); } catch { /* host sem o serviço */ }
        const dv = options.dataViews && options.dataViews[0];
        this.settings = this.fmtService.populateFormattingSettingsModel(VisualSettings, dv);
        this.lerPeriodoSalvo(dv);
        this.lerDados(dv);
        this.carregando = false;
        if (dv && dv.metadata && dv.metadata.segment) {
            try { this.carregando = this.host.fetchMoreData(true); } catch { this.carregando = false; }
        }
        if (this.pendente && this.cab && this.periodoChegou()) { this.pendente = null; }
        this.conferirAtalhoVencido();
        this.desenhar();
        try { this.host.eventService.renderingFinished(options); } catch { /* host sem o serviço */ }
    }

    private lerPeriodoSalvo(dv: powerbi.DataView | undefined): void {
        const o = dv && dv.metadata && dv.metadata.objects && (dv.metadata.objects["periodo"] as { preset?: string; ini?: string; fim?: string } | undefined);
        const p = o && o.preset ? String(o.preset) : "";
        this.presetSalvo = (ATALHOS.some(a => a.id === p) ? p : "") as Preset | "";
        this.iniSalvo = o && o.ini ? String(o.ini) : "";
        this.fimSalvo = o && o.fim ? String(o.fim) : "";
    }

    private lerDados(dv: powerbi.DataView | undefined): void {
        this.temDados = false; this.formato = "ok"; this.reps = [];
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
        this.cab = c.cab;
        this.reps = c.reps;
        this.temDados = true;
        // filtros locais que não existem mais na lista são descartados
        if (this.cab) {
            const sups = new Set(this.cab.sups), ufs = new Set(this.cab.ufs);
            this.sup.forEach(v => { if (!sups.has(v)) { this.sup.delete(v); } });
            this.uf.forEach(v => { if (!ufs.has(v)) { this.uf.delete(v); } });
        }
    }

    private periodoChegou(): boolean {
        const p = this.pendente, c = this.cab;
        if (!p || !c) { return true; }
        if (p.preset === "mes") { return !c.temPeriodo; }
        return c.temPeriodo && c.ini === p.ini && c.fim === p.fim;
    }

    // Atalho relativo (Hoje, Ontem, 7/30 dias, Mês passado) salvo em outro dia: recalcula as datas e reaplica uma vez
    private conferirAtalhoVencido(): void {
        const p = this.presetSalvo, c = this.cab;
        if (!c || this.pendente || !p || RELATIVOS.indexOf(p as Preset) < 0) { return; }
        const [a, b] = intervalo(p as Preset, hojeSP());
        const chave = p + a + b;
        if (c.temPeriodo && c.ini === a && c.fim === b) { return; }
        if (this.reaplicado === chave) { return; }
        this.reaplicado = chave;
        this.pendente = { ini: a, fim: b, preset: p as Preset };
        this.aplicarFiltroPeriodo(a, b, p as Preset);
    }

    // ---------------- recortes ----------------
    private filtrados(): RepMix[] {
        const termo = semAcento(this.busca.trim());
        const lista = this.reps.filter(r => (!this.sup.size || this.sup.has(r.sup)) && (!this.uf.size || this.uf.has(r.uf)) && (!termo || semAcento(r.nome).indexOf(termo) >= 0));
        const o = this.ordem;
        return lista.sort((a, b) => (o === "uf" ? colador.compare(a.uf, b.uf) : o === "valor" ? b.v30 - a.v30 : 0) || colador.compare(a.nome, b.nome) || colador.compare(a.cod, b.cod));
    }
    private ativos(): number { return (this.sup.size ? 1 : 0) + (this.uf.size ? 1 : 0) + (this.cab && this.cab.temPeriodo ? 1 : 0); }

    private presetAtual(): Preset {
        if (this.pendente) { return this.pendente.preset; }
        if (!this.cab || !this.cab.temPeriodo) { return "mes"; }
        return (this.presetSalvo || "intervalo") as Preset;
    }
    private textoPeriodo(curto = false): string {
        const c = this.cab;
        const ini = this.pendente ? this.pendente.ini : c ? c.ini : "", fim = this.pendente ? this.pendente.fim : c ? c.fim : "";
        if (!ini) { return ""; }
        const rot = rotuloPreset(this.presetAtual());
        return `Período: ${dataBR(ini)} a ${dataBR(fim)}` + (rot ? (curto ? ` (${rot})` : ` · ${rot}`) : "") + (curto ? "" : " · data de entrada do pedido");
    }
    private cobertura(): string {
        const c = this.cab;
        if (!c || !c.fatMax || c.fatMax >= c.fim) { return ""; }
        return `Cobertura a conferir: maior data de pedido no faturamento ${dataBR(c.fatMax)}; carteira ${dataBR(c.cartMax)}. Veja "Sobre a base".`;
    }

    // ---------------- tela ----------------
    private desenhar(): void {
        const s = this.settings;
        this.raiz.style.background = s.fundo.mostrar.value ? s.fundo.cor.value.value : "transparent";
        this.raiz.classList.toggle("sem-fundo", !s.fundo.mostrar.value);
        this.desenharBarra();
        this.desenharSub();
        this.desenharQuadro();
        this.desenharCamadas();
    }

    private desenharBarra(): void {
        const b = this.barra;
        while (b.firstChild) { b.removeChild(b.firstChild); }
        const info = el("div", "om-info");
        const per = el("div", "om-per");
        const tp = this.textoPeriodo();
        if (tp) {
            const i = tp.indexOf(" · ");
            per.appendChild(el("strong", "", i > 0 ? tp.slice(0, i) : tp));
            if (i > 0) { per.appendChild(document.createTextNode(tp.slice(i))); }
        }
        info.appendChild(per);
        const cont = el("div", "om-cont");
        cont.setAttribute("aria-live", "polite");
        info.appendChild(cont);
        b.appendChild(info);
        this.atualizarContagem(cont);

        const acoes = el("div", "om-acoes");
        const bus = document.createElement("input");
        bus.type = "search"; bus.className = "om-busca";
        bus.placeholder = "Buscar representante…";
        bus.setAttribute("aria-label", "Buscar representante pelo nome");
        bus.value = this.busca;
        bus.addEventListener("input", () => { this.busca = bus.value; this.desenharQuadro(); this.atualizarContagem(); });
        acoes.appendChild(bus);
        acoes.appendChild(this.botaoFiltros());
        if (this.settings.exportacao.mostrarBotao.value) { acoes.appendChild(this.menuExportar()); }
        const sb = botao("om-btn", "Sobre a base", () => this.abrirSobre());
        sb.setAttribute("aria-haspopup", "dialog");
        acoes.appendChild(sb);
        b.appendChild(acoes);
    }

    private atualizarContagem(alvo?: HTMLElement): void {
        const cont = alvo || (this.barra.querySelector(".om-cont") as HTMLElement | null);
        if (!cont) { return; }
        while (cont.firstChild) { cont.removeChild(cont.firstChild); }
        if (!this.temDados) { return; }
        cont.appendChild(el("strong", "", fmtInt.format(this.filtrados().length)));
        cont.appendChild(document.createTextNode(" representantes no filtro"));
        if (this.cab) { cont.appendChild(el("span", "om-base", ` / ${fmtInt.format(this.cab.base)} na base`)); }
    }

    private desenharSub(): void {
        const s = this.sub;
        while (s.firstChild) { s.removeChild(s.firstChild); }
        s.hidden = !this.temDados;
        if (!this.temDados) { return; }
        const ord = el("div", "om-ord");
        ord.setAttribute("role", "group");
        ord.setAttribute("aria-label", "Ordenar as listas");
        ord.appendChild(el("span", "om-ord-rot", "Ordenar:"));
        ORDENS.forEach(o => {
            const b = botao("om-ord-b" + (this.ordem === o.id ? " sel" : ""), o.rotulo, () => { this.ordem = o.id; this.desenharSub(); this.desenharQuadro(); });
            b.title = o.dica;
            b.setAttribute("aria-pressed", String(this.ordem === o.id));
            ord.appendChild(b);
        });
        s.appendChild(ord);
        const st = el("div", "om-status");
        st.setAttribute("role", "status");
        if (this.pendente) { st.appendChild(el("span", "om-atual", "Atualizando o período…")); }
        else if (this.carregando) { st.appendChild(el("span", "om-atual", "Carregando mais representantes…")); }
        const cob = this.cobertura();
        if (cob && !this.pendente) { st.appendChild(el("span", "om-alerta", cob)); }
        s.appendChild(st);
    }

    private desenharQuadro(): void {
        const q = this.quadro;
        const rolagem = Array.from(q.querySelectorAll(".om-lista")).map(x => (x as HTMLElement).scrollTop);
        while (q.firstChild) { q.removeChild(q.firstChild); }
        q.classList.toggle("pendente", !!this.pendente);
        if (this.formato === "errada") {
            q.appendChild(el("div", "om-vazio erro", `A medida do campo não é a ${MEDIDA}. Troque pela medida do arquivo ${ARQUIVO_DAX}.`));
            return;
        }
        if (!this.temDados) {
            const v = el("div", "om-vazio", this.carregando ? "Carregando…" : `Arraste o campo Representante (d_Representantes[Código do Vendedor]) e a medida ${MEDIDA} para o visual.`);
            v.setAttribute("role", "status");
            q.appendChild(v);
            return;
        }
        const reps = this.filtrados();
        if (!reps.length) {
            q.appendChild(el("div", "om-vazio", "Nenhum representante no filtro. Limpe a busca ou os filtros de Supervisão e UF."));
            return;
        }
        FAMILIAS.forEach((f, k) => {
            const lista = reps.filter(r => r.f[k].n === 0);
            const sec = el("section", "om-fam");
            sec.style.setProperty("--fam", f.cor);
            sec.setAttribute("aria-label", `${f.nome}: ${lista.length} sem venda`);
            const hd = el("div", "om-fam-cab");
            const tt = el("div", "om-fam-tit");
            const bd = el("span", "om-badge", f.s);
            bd.style.background = f.cor; bd.style.color = f.texto;
            tt.appendChild(bd);
            tt.appendChild(el("span", "om-fam-nome", f.nome));
            tt.title = f.modelo;
            hd.appendChild(tt);
            const n = el("div", "om-fam-n");
            n.appendChild(el("strong", "", fmtInt.format(lista.length)));
            n.appendChild(el("small", "", "sem venda"));
            hd.appendChild(n);
            sec.appendChild(hd);
            const ul = el("div", "om-lista");
            ul.setAttribute("role", "list");
            if (!lista.length) { ul.appendChild(el("div", "om-ok", "Nenhum representante sem venda neste recorte.")); }
            lista.forEach(r => {
                const it = botao("om-rep", "", () => this.abrirFicha(r.cod, k));
                it.setAttribute("role", "listitem");
                it.setAttribute("aria-label", `Abrir ficha de ${r.nome} (${r.uf}), sem venda em ${f.nome}`);
                const nm = el("span", "om-rep-nome", r.nome); nm.title = r.nome;
                it.appendChild(nm);
                it.appendChild(el("span", "om-rep-uf", r.uf));
                ul.appendChild(it);
            });
            sec.appendChild(ul);
            q.appendChild(sec);
        });
        Array.from(q.querySelectorAll(".om-lista")).forEach((x, i) => { if (rolagem[i]) { (x as HTMLElement).scrollTop = rolagem[i]; } });
    }

    private desenharCamadas(): void {
        Array.from(this.raiz.querySelectorAll(".fp, .om-ov")).forEach(x => x.parentNode && x.parentNode.removeChild(x));
        if (this.popAberto) {
            const fp = this.painelFiltros();
            const topo = this.barra.offsetTop + this.barra.offsetHeight + 4;
            if (topo > 8) { fp.style.top = topo + "px"; fp.style.maxHeight = `calc(100% - ${topo + 8}px)`; }
            this.raiz.appendChild(fp);
        }
        if (this.ficha) { this.raiz.appendChild(this.dialogoFicha()); }
        else if (this.sobre) { this.raiz.appendChild(this.dialogoSobre()); }
    }

    // ---------------- ficha / sobre ----------------
    private abrirFicha(cod: string, fam: number): void {
        if (this.pendente) { return; }
        this.focoAntes = document.activeElement as HTMLElement;
        this.menuAberto = false; this.popAberto = false;
        this.ficha = { cod, fam };
        this.desenharCamadas();
        const d = this.raiz.querySelector(".om-dlg") as HTMLElement | null;
        if (d) { d.focus(); }
    }
    private abrirSobre(): void {
        this.focoAntes = document.activeElement as HTMLElement;
        this.menuAberto = false; this.popAberto = false;
        this.sobre = true;
        this.desenharBarra();
        this.desenharCamadas();
        const d = this.raiz.querySelector(".om-dlg") as HTMLElement | null;
        if (d) { d.focus(); }
    }
    private fecharDialogo(): void {
        this.ficha = null; this.sobre = false;
        this.desenharCamadas();
        if (this.focoAntes && this.raiz.contains(this.focoAntes)) { this.focoAntes.focus(); }
    }
    private moldura(titulo: string, eyebrow: string, meta: string): { ov: HTMLElement; corpo: HTMLElement } {
        const ov = el("div", "om-ov");
        ov.addEventListener("click", ev => { if (ev.target === ov) { this.fecharDialogo(); } });
        const d = el("div", "om-dlg");
        d.setAttribute("role", "dialog");
        d.setAttribute("aria-modal", "true");
        d.setAttribute("aria-labelledby", "om-dlg-t");
        d.tabIndex = -1;
        const hd = el("div", "om-dlg-cab");
        const tx = el("div", "om-dlg-tx");
        if (eyebrow) { tx.appendChild(el("div", "om-dlg-eye", eyebrow)); }
        const h = el("h3", "om-dlg-tit", titulo); h.id = "om-dlg-t";
        tx.appendChild(h);
        if (meta) { tx.appendChild(el("div", "om-dlg-meta", meta)); }
        hd.appendChild(tx);
        const x = botao("om-x", "×", () => this.fecharDialogo());
        x.setAttribute("aria-label", "Fechar");
        hd.appendChild(x);
        d.appendChild(hd);
        const corpo = el("div", "om-dlg-corpo");
        d.appendChild(corpo);
        // foco preso no diálogo
        d.addEventListener("keydown", (ev: KeyboardEvent) => {
            if (ev.key !== "Tab") { return; }
            const f = Array.from(d.querySelectorAll("button, input, [tabindex='0']")) as HTMLElement[];
            if (!f.length) { ev.preventDefault(); return; }
            const a = document.activeElement;
            if (ev.shiftKey && (a === f[0] || a === d)) { ev.preventDefault(); f[f.length - 1].focus(); }
            else if (!ev.shiftKey && a === f[f.length - 1]) { ev.preventDefault(); f[0].focus(); }
        });
        ov.appendChild(d);
        return { ov, corpo };
    }

    private dialogoFicha(): HTMLElement {
        const fi = this.ficha as { cod: string; fam: number };
        const r = this.reps.find(x => x.cod === fi.cod);
        const fim = this.cab ? this.cab.fim : "";
        if (!r) { this.ficha = null; return el("div"); }
        const { ov, corpo } = this.moldura(r.nome, "Oportunidade de mix · " + FAMILIAS[fi.fam].nome, `Código ${r.cod} · ${r.sup} · ${r.uf}`);
        if (this.cab) { corpo.appendChild(el("div", "om-dlg-per", `${dataBR(this.cab.ini)} a ${dataBR(this.cab.fim)} · data de entrada do pedido`)); }
        const mix = mixDe(r);
        const k = el("div", "om-kpis");
        [["Vendido no período", fmtBRL.format(r.vend)], ["Pedidos válidos", fmtInt.format(r.ped)], ["Ticket médio por pedido", r.ped ? fmtBRL.format(r.vend / r.ped) : "—"], ["Mix vendido", `${mix} de 5 famílias`]]
            .forEach(([a, b]) => { const c = el("div", "om-kpi"); c.appendChild(el("small", "", a)); c.appendChild(el("strong", "", b)); k.appendChild(c); });
        corpo.appendChild(k);
        const t = el("table", "om-tab");
        t.setAttribute("aria-label", "Venda por família de " + r.nome);
        const th = el("thead"), hr = el("tr");
        ["Família", "Última venda", "Dias sem vender", "No período"].forEach(x => { const h = el("th", "", x); h.setAttribute("scope", "col"); hr.appendChild(h); });
        th.appendChild(hr); t.appendChild(th);
        const tb = el("tbody");
        FAMILIAS.forEach((f, i) => {
            const d = r.f[i];
            const tr = el("tr", i === fi.fam ? "dest" : "");
            const c0 = el("th", "");
            c0.setAttribute("scope", "row");
            const bd = el("span", "om-badge pq", f.s); bd.style.background = f.cor; bd.style.color = f.texto;
            c0.appendChild(bd); c0.appendChild(document.createTextNode(" " + f.nome));
            tr.appendChild(c0);
            tr.appendChild(el("td", "", d.u ? dataBR(d.u) : "Sem registro"));
            const ds = diasSem(d.u, fim);
            tr.appendChild(el("td", "num", ds === null ? "—" : fmtInt.format(ds)));
            tr.appendChild(el("td", d.n ? "" : "falta", d.n ? `${fmtInt.format(d.n)} pedido${d.n > 1 ? "s" : ""} · ${fmtBRL.format(d.v)}` : "Sem venda"));
            tb.appendChild(tr);
        });
        t.appendChild(tb);
        const tw = el("div", "om-tw"); tw.appendChild(t);
        corpo.appendChild(tw);
        const faltam = FAMILIAS.filter((_, i) => r.f[i].n === 0).map(f => f.s);
        corpo.appendChild(el("p", "om-nota", (faltam.length ? `Faltam no período: ${faltam.join(", ")}. ` : "Vendeu nas 5 famílias no período. ") + (fim ? `Dias sem vender contados até ${dataBR(fim)} (fim do período).` : "")));
        return ov;
    }

    private dialogoSobre(): HTMLElement {
        const { ov, corpo } = this.moldura("Sobre a base", "Gestão Comercial · Mix de produtos", "");
        const c = this.cab;
        corpo.appendChild(el("p", "", "Representantes ativos e não bloqueados no SAP, sem Venda Direta. Cada coluna lista quem não teve nenhum pedido válido da família no período escolhido."));
        const ul = el("ul", "om-ul");
        [
            "Pedido válido: pedido faturado ou em carteira (sem recusa) com saldo positivo, pela data de entrada do pedido. O mesmo pedido nas duas bases conta uma vez.",
            "FB = FORRO E ACESS. BR; FC = FORRO E ACESS. COLOR; PT = PORTA PVC; AP = ACESS. PARA PISO; CT = CANTINHO.",
            "A carteira é a posição carregada agora; não reconstrói bloqueios de períodos anteriores.",
            "Ticket médio = vendido nas cinco famílias ÷ pedidos distintos do período. \"Maior valor (30 dias)\" usa os 30 dias até o fim do período.",
            "Última venda: maior data de entrada de pedido válido da família em todo o histórico até o fim do período."
        ].forEach(t => ul.appendChild(el("li", "", t)));
        corpo.appendChild(ul);
        if (c) {
            const p = el("p", "om-dlg-datas");
            p.appendChild(el("strong", "", "Maior data de pedido no faturamento: ")); p.appendChild(document.createTextNode(dataBR(c.fatMax) || "—"));
            p.appendChild(el("br"));
            p.appendChild(el("strong", "", "Maior data de pedido na carteira: ")); p.appendChild(document.createTextNode(dataBR(c.cartMax) || "—"));
            p.appendChild(el("br"));
            p.appendChild(el("strong", "", "Base elegível: ")); p.appendChild(document.createTextNode(fmtInt.format(c.base) + " representantes"));
            corpo.appendChild(p);
        }
        corpo.appendChild(el("p", "om-aviso-base", "Ausência de registro no histórico não comprova que o representante nunca vendeu a família."));
        return ov;
    }

    // ---------------- filtros ----------------
    private botaoFiltros(): HTMLElement {
        const n = this.ativos();
        const b = botao("om-btn fl-btn" + (n ? " ativo" : ""), "Filtros", () => {
            if (this.popAberto) { this.fecharPainel(); return; }
            this.menuAberto = false;
            this.pSup = new Set(this.sup); this.pUf = new Set(this.uf);
            this.pPreset = this.presetAtual();
            const c = this.cab;
            this.pIni = this.pendente ? this.pendente.ini : c ? c.ini : "";
            this.pFim = this.pendente ? this.pendente.fim : c ? c.fim : "";
            this.popAberto = true;
            this.desenharBarra();
            this.desenharCamadas();
            const alvo = this.raiz.querySelector(".fp-campo.sel") as HTMLElement | null;
            if (alvo) { alvo.focus(); }
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
        this.desenharBarra();
        this.desenharCamadas();
        if (this.btnFiltros) { this.btnFiltros.focus(); }
    }

    private painelFiltros(): HTMLElement {
        const p = el("div", "fp");
        p.setAttribute("role", "dialog");
        p.setAttribute("aria-label", "Filtros");
        const cab = el("div", "fp-cab");
        cab.appendChild(el("div", "fp-tit", "Filtros"));
        cab.appendChild(botao("fp-link", "Limpar tudo", () => {
            this.pSup = new Set(); this.pUf = new Set(); this.pPreset = "mes";
            this.busca = ""; this.aplicar();
        }));
        const x = botao("fp-x", "×", () => this.fecharPainel());
        x.setAttribute("aria-label", "Fechar filtros");
        cab.appendChild(x);
        p.appendChild(cab);

        const corpo = el("div", "fp-corpo");
        const nav = el("div", "fp-nav");
        nav.setAttribute("role", "tablist");
        nav.setAttribute("aria-label", "Campos");
        const campos: { id: "per" | "sup" | "uf"; rot: string; q: number }[] = [
            { id: "per", rot: "Período", q: this.pPreset === "mes" ? 0 : 1 },
            { id: "sup", rot: "Supervisão", q: this.pSup.size },
            { id: "uf", rot: "UF", q: this.pUf.size }
        ];
        campos.forEach(cp => {
            const b = botao("fp-campo" + (cp.id === this.campoAtivo ? " sel" : ""), cp.rot, () => {
                this.campoAtivo = cp.id; this.desenharCamadas();
                const s = this.raiz.querySelector(".fp-campo.sel") as HTMLElement | null;
                if (s) { s.focus(); }
            });
            b.setAttribute("role", "tab");
            b.setAttribute("aria-selected", String(cp.id === this.campoAtivo));
            b.dataset.campo = cp.id;
            b.appendChild(el("span", "fp-q" + (cp.q ? "" : " zero"), cp.q ? String(cp.q) : ""));
            nav.appendChild(b);
        });
        corpo.appendChild(nav);
        const erro = el("p", "fp-erro");
        erro.setAttribute("role", "alert");
        corpo.appendChild(this.campoAtivo === "per" ? this.painelPeriodo(erro) : this.painelLista(this.campoAtivo, nav));
        p.appendChild(corpo);
        const rod = el("div", "fp-rod");
        const txt = el("div", "fp-per");
        const [a, b] = this.pPreset === "intervalo" ? [this.pIni, this.pFim] : intervalo(this.pPreset, hojeSP());
        txt.appendChild(el("div", "", a && b ? `Período escolhido: ${dataBR(a)} a ${dataBR(b)}` : ""));
        txt.appendChild(erro);
        rod.appendChild(txt);
        rod.appendChild(botao("fp-aplicar", "Aplicar", () => {
            if (this.pPreset === "intervalo") {
                const m = validar(this.pIni, this.pFim, hojeSP(), this.cab ? this.cab.calMin : "");
                if (m) { erro.textContent = m; return; }
            }
            this.aplicar();
        }));
        p.appendChild(rod);
        return p;
    }

    private painelPeriodo(erro: HTMLElement): HTMLElement {
        const box = el("div", "fp-pane");
        box.setAttribute("role", "tabpanel");
        const g = el("div", "fp-chips");
        g.setAttribute("role", "group");
        g.setAttribute("aria-label", "Atalhos de período");
        ATALHOS.forEach(a => {
            const b = botao("fp-chip" + (this.pPreset === a.id ? " sel" : ""), a.rotulo, () => {
                this.pPreset = a.id;
                if (a.id === "intervalo" && (!this.pIni || !this.pFim)) { const c = this.cab; this.pIni = c ? c.ini : hojeSP(); this.pFim = c ? c.fim : hojeSP(); }
                this.desenharCamadas();
                const s = this.raiz.querySelector(".fp-chip.sel") as HTMLElement | null;
                if (s) { s.focus(); }
            });
            b.setAttribute("aria-pressed", String(this.pPreset === a.id));
            g.appendChild(b);
        });
        box.appendChild(g);
        if (this.pPreset === "intervalo") {
            const dt = el("div", "fp-datas");
            const campo = (rot: string, val: string, set: (v: string) => void): HTMLElement => {
                const lb = el("label", "fp-data");
                lb.appendChild(el("span", "", rot));
                const i = document.createElement("input");
                i.type = "date"; i.value = val; i.max = hojeSP();
                if (this.cab && this.cab.calMin) { i.min = this.cab.calMin; }
                i.addEventListener("change", () => { set(i.value); erro.textContent = ""; const r = this.raiz.querySelector(".fp-per div") as HTMLElement | null; if (r && this.pIni && this.pFim) { r.textContent = `Período escolhido: ${dataBR(this.pIni)} a ${dataBR(this.pFim)}`; } });
                lb.appendChild(i);
                return lb;
            };
            dt.appendChild(campo("De", this.pIni, v => { this.pIni = v; }));
            dt.appendChild(campo("Até", this.pFim, v => { this.pFim = v; }));
            box.appendChild(dt);
            box.appendChild(el("p", "fp-dica", "Para um único dia, use a mesma data nos dois campos."));
        }
        box.appendChild(el("p", "fp-dica", "Data de entrada do pedido, horário de Brasília. \"Neste mês\" vai do dia 1 até hoje."));
        return box;
    }

    private painelLista(id: "sup" | "uf", nav: HTMLElement): HTMLElement {
        const rot = id === "sup" ? "Supervisão" : "UF";
        const box = el("div", "fp-pane");
        box.setAttribute("role", "tabpanel");
        const todos = (this.cab ? (id === "sup" ? this.cab.sups : this.cab.ufs) : []).slice();
        const sel = id === "sup" ? this.pSup : this.pUf;
        sel.forEach(v => { if (todos.indexOf(v) < 0) { todos.push(v); } });
        const busca = document.createElement("input");
        busca.type = "search"; busca.className = "fp-busca";
        busca.placeholder = "Buscar " + rot.toLowerCase() + "…";
        busca.setAttribute("aria-label", "Buscar em " + rot);
        busca.value = this.buscaLista.get(id) || "";
        box.appendChild(busca);
        const info = el("div", "fp-info");
        const lista = el("div", "fp-lista");
        lista.setAttribute("role", "group");
        lista.setAttribute("aria-label", rot);
        const contar = (): void => {
            const q = sel.size;
            const b = nav.querySelector(`[data-campo="${id}"] .fp-q`) as HTMLElement | null;
            if (b) { b.textContent = q ? String(q) : ""; b.className = "fp-q" + (q ? "" : " zero"); }
            info.textContent = (q ? `${q} selecionado${q > 1 ? "s" : ""}` : "Nenhum selecionado (sem filtro)") + ` · ${todos.length} valores`;
        };
        let visiveis: string[] = [];
        const desenharLista = (): void => {
            while (lista.firstChild) { lista.removeChild(lista.firstChild); }
            const termo = semAcento(busca.value.trim());
            visiveis = termo ? todos.filter(v => semAcento(v).indexOf(termo) >= 0) : todos;
            if (!todos.length) { lista.appendChild(el("div", "fp-msg", `Sem valores: confira a medida ${MEDIDA}.`)); return; }
            if (!visiveis.length) { lista.appendChild(el("div", "fp-msg", "Nada encontrado.")); return; }
            visiveis.slice(0, MAX_ITENS).forEach(v => {
                const lb = el("label", "fp-op");
                const cb = document.createElement("input");
                cb.type = "checkbox"; cb.checked = sel.has(v);
                cb.addEventListener("change", () => { if (cb.checked) { sel.add(v); } else { sel.delete(v); } contar(); });
                lb.appendChild(cb);
                lb.appendChild(el("span", "", v));
                lista.appendChild(lb);
            });
        };
        busca.addEventListener("input", () => { this.buscaLista.set(id, busca.value); desenharLista(); });
        const marcar = el("div", "fp-marcar");
        marcar.appendChild(botao("fp-link", "Marcar visíveis", () => { visiveis.forEach(v => sel.add(v)); desenharLista(); contar(); }));
        marcar.appendChild(botao("fp-link", "Limpar", () => { sel.clear(); desenharLista(); contar(); }));
        box.appendChild(marcar);
        box.appendChild(lista);
        box.appendChild(info);
        desenharLista();
        contar();
        return box;
    }

    private aplicar(): void {
        this.sup = new Set(this.pSup);
        this.uf = new Set(this.pUf);
        const c = this.cab;
        const [a, b] = intervalo(this.pPreset, hojeSP(), this.pIni, this.pFim);
        const atual = this.presetAtual();
        const mudou = this.pPreset !== atual || (this.pPreset !== "mes" && (!c || !c.temPeriodo || c.ini !== a || c.fim !== b));
        this.popAberto = false;
        if (mudou) {
            this.pendente = { ini: a, fim: b, preset: this.pPreset };
            this.reaplicado = this.pPreset + a + b;
            this.persistirPeriodo(this.pPreset, a, b);
            this.aplicarFiltroPeriodo(a, b, this.pPreset);
        }
        this.desenhar();
        if (this.btnFiltros) { this.btnFiltros.focus(); }
    }

    private persistirPeriodo(preset: Preset, ini: string, fim: string): void {
        try {
            this.host.persistProperties({ merge: [{ objectName: "periodo", selector: null, properties: { preset, ini: preset === "mes" ? "" : ini, fim: preset === "mes" ? "" : fim } }] });
        } catch { /* host sem persistProperties */ }
    }

    // Filtro de intervalo (avançado) na coluna de data; "Neste mês" remove o filtro (a medida usa dia 1 até hoje).
    private aplicarFiltroPeriodo(ini: string, fim: string, preset: Preset): void {
        const alvo = lerAlvo(this.settings.periodoCfg.coluna.value) || (lerAlvo(COLUNA_DATA) as Alvo);
        const filtro = preset === "mes" ? null : {
            $schema: "https://powerbi.com/product/schema#advanced",
            target: { table: alvo.table, column: alvo.column },
            filterType: 0,
            logicalOperator: "And",
            conditions: [{ operator: "GreaterThanOrEqual", value: ini + "T00:00:00" }, { operator: "LessThanOrEqual", value: fim + "T00:00:00" }]
        };
        const carga = filtro as unknown as powerbi.IFilter;
        const so = this.settings.periodoCfg.soVisual.value;
        try {
            this.host.applyJsonFilter(carga, "general", so ? "selfFilter" : "filter", powerbi.FilterAction.merge);
        } catch {
            try { this.host.applyJsonFilter(carga, "general", "filter", powerbi.FilterAction.merge); } catch { this.pendente = null; this.mostrar("Não foi possível aplicar o período neste ambiente.", true); }
        }
    }

    private resumoFiltros(): string {
        const p: string[] = [];
        const lst = (rot: string, s: Set<string>): void => { if (s.size) { const v = Array.from(s); p.push(rot + " = " + (v.length > 4 ? v.slice(0, 3).join(", ") + ` e mais ${v.length - 3}` : v.join(", "))); } };
        lst("Supervisão", this.sup); lst("UF", this.uf);
        if (this.busca.trim()) { p.push(`Busca = "${this.busca.trim()}"`); }
        return p.length ? "Filtros: " + p.join("; ") : "";
    }

    // ---------------- exportar ----------------
    private menuExportar(): HTMLElement {
        const cx = el("div", "ex");
        const btn = botao("om-btn ex-btn", "Exportar ▾", () => { this.menuAberto = !this.menuAberto; this.limparAviso(); if (this.menuAberto) { this.popAberto = false; this.desenharCamadas(); } this.desenharBarra(); });
        btn.setAttribute("aria-haspopup", "true");
        btn.setAttribute("aria-expanded", String(this.menuAberto));
        btn.disabled = this.ocupado;
        cx.appendChild(btn);
        if (this.menuAberto) {
            const m = el("div", "ex-menu");
            m.appendChild(el("div", "ex-tit", "Excel com 3 abas"));
            m.appendChild(el("p", "ex-txt", "Resumo por família, Sem venda por família e Representantes (mix e última venda), com os filtros e a ordem da tela."));
            const go = botao("ex-go", "Baixar Excel", () => { void this.exportarExcel(); });
            m.appendChild(go);
            cx.appendChild(m);
            setTimeout(() => go.focus(), 0);
        }
        if (this.aviso) {
            const msg = el("div", "ex-msg" + (this.avisoErro ? " erro" : ""), this.aviso);
            msg.setAttribute("role", this.avisoErro ? "alert" : "status");
            const x = botao("ex-msg-x", "×", () => { this.limparAviso(); this.desenharBarra(); });
            x.setAttribute("aria-label", "Fechar aviso");
            msg.appendChild(x);
            cx.appendChild(msg);
        }
        return cx;
    }

    private limparAviso(): void {
        if (this.timerAviso !== undefined) { window.clearTimeout(this.timerAviso); this.timerAviso = undefined; }
        this.aviso = ""; this.avisoErro = false;
    }
    // Aviso que some sozinho: ~6 s (informação) ou ~10 s (erro), com fade; sem animação se o usuário pediu menos movimento.
    private mostrar(t: string, erro = false): void {
        this.limparAviso();
        this.aviso = t; this.avisoErro = erro;
        this.desenharBarra();
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

    public entrada(): EntradaMix {
        const c = this.cab;
        return {
            titulo: this.settings.titulo.texto.value || "Oportunidade de Mix", geradoEm: this.agora(), periodo: this.textoPeriodo(true) + " · data de entrada do pedido",
            ini: c ? c.ini : "", fim: c ? c.fim : "", filtros: this.resumoFiltros(), ordem: (ORDENS.find(o => o.id === this.ordem) || ORDENS[0]).texto,
            base: c ? c.base : 0, cobertura: this.cobertura().replace(` Veja "Sobre a base".`, ""), reps: this.filtrados()
        };
    }

    // Copia texto para a área de transferência: API do navegador e, se o iframe do Power BI bloquear, textarea + execCommand.
    private async copiar(texto: string): Promise<boolean> {
        try {
            const cb = navigator.clipboard;
            if (cb && cb.writeText) { await cb.writeText(texto); return true; }
        } catch { /* bloqueado pela política do iframe: tenta o modo antigo */ }
        const ta = document.createElement("textarea");
        ta.value = texto; ta.setAttribute("readonly", ""); ta.className = "om-copia";
        this.raiz.appendChild(ta);
        let ok = false;
        try { ta.focus(); ta.select(); ta.setSelectionRange(0, texto.length); ok = document.execCommand("copy"); } catch { ok = false; }
        this.raiz.removeChild(ta);
        return ok;
    }

    public static async codificar(e: EntradaMix): Promise<{ b64: string; comp: string }> {
        let dados: Uint8Array = new TextEncoder().encode(JSON.stringify(compactar(e)));
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

    // Exportar: sempre pela página de download (funciona no Power BI embutido). Dados no link quando cabem;
    // senão área de transferência + link curto (k=cola) e "Colar dados" na página. Nunca usa download direto.
    private async exportarExcel(): Promise<void> {
        if (this.ocupado) { return; }
        this.menuAberto = false;
        if (!this.temDados) { this.mostrar("Sem dados para exportar.", true); return; }
        if (this.pendente) { this.mostrar("Aguarde o período terminar de carregar e exporte de novo.", true); return; }
        const e = this.entrada();
        if (!e.reps.length) { this.mostrar("Nenhum representante no filtro para exportar.", true); return; }
        const ex = this.settings.exportacao;
        const url = (ex.urlPagina.value || "").trim();
        if (!/^https:\/\//i.test(url)) { this.mostrar("Informe o endereço da página de download em Formatar > Exportação.", true); return; }
        const nome = (ex.nomeArquivo.value || "oportunidade_mix").replace(/[^\w-]+/g, "_") + ".xlsx";
        this.ocupado = true;
        try {
            const { b64, comp } = await Visual.codificar(e);
            const base = url.replace(/#.*$/, "") + "#";
            const alvo = base + new URLSearchParams({ v: "2", n: nome, c: comp, j: b64 }).toString();
            if (alvo.length <= LIMITE_LINK) {
                this.host.launchUrl(alvo);
                this.mostrar("Página de download aberta: " + nome);
                return;
            }
            const copiou = comp === "1" && await this.copiar(PREFIXO_COLA + b64);
            if (copiou) {
                this.host.launchUrl(base + new URLSearchParams({ v: "2", n: nome, k: "cola", q: String(e.reps.length), u: "representantes" }).toString());
                this.mostrar("A lista não cabe no link. Os dados foram copiados: na página que abriu, clique em Colar dados.");
            } else {
                this.mostrar("A lista não cabe no link e o Power BI bloqueou a cópia. Filtre por Supervisão ou UF e exporte de novo.", true);
            }
        } catch (er) {
            this.mostrar("Não foi possível exportar: " + (er instanceof Error ? er.message : String(er)), true);
        } finally {
            this.ocupado = false;
            this.desenharBarra();
        }
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.fmtService.buildFormattingModel(this.settings);
    }
}
