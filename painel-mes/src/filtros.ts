// Botão Filtros (filtra a página toda): mesmo mecanismo do Pedido Diário v2.0 / Reps Sem Pedidos v1.0,
// separado numa classe para o Painel Mês e o Painel Dia usarem o mesmo código.
import powerbi from "powerbi-visuals-api";
import { FiltrosCard, COLUNAS_PADRAO } from "./settings";

import IVisualHost = powerbi.extensibility.visual.IVisualHost;

type ChaveColuna = keyof typeof COLUNAS_PADRAO;
interface Campo { id: string; rotulo: string; col: ChaveColuna; num: boolean; lista: number; }
// lista = posição da lista de valores no cabeçalho da medida (depois da data de referência)
export const CAMPOS: Campo[] = [
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

interface Alvo { table: string; column: string; }
export interface FiltroBasico { $schema: string; target: Alvo; filterType: number; operator: string; values: (string | number)[]; requireSingleSelection: boolean; }

export function el(tag: string, cls?: string, texto?: string): HTMLElement {
    const e = document.createElement(tag);
    if (cls) { e.className = cls; }
    if (texto !== undefined) { e.textContent = texto; }
    return e;
}
export function botao(cls: string, texto: string, acao: () => void): HTMLButtonElement {
    const b = el("button", cls, texto) as HTMLButtonElement;
    b.type = "button";
    b.addEventListener("click", acao);
    return b;
}
function semAcento(s: string): string {
    return s.normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase();
}
export function dataBR(iso: string): string {
    const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso || "");
    return m ? `${m[3]}/${m[2]}/${m[1]}` : "";
}
// "Tabela[Coluna]; 'Outra Tabela'[Coluna]" -> alvos do filtro
export function lerAlvos(texto: string): Alvo[] {
    return (texto || "").split(";").map(p => /^\s*'?(.+?)'?\s*\[(.+)\]\s*$/.exec(p)).filter(m => !!m).map(m => ({ table: (m as RegExpExecArray)[1].trim(), column: (m as RegExpExecArray)[2].trim() }));
}

export class Filtros {
    public dataRef = "";
    public popAberto = false;
    private listas = new Map<string, string[]>();
    private aplicados = new Map<string, Set<string>>();
    private pendentes = new Map<string, Set<string>>();
    private busca = new Map<string, string>();
    private campoAtivo = "ano";
    private btn: HTMLButtonElement | null = null;

    // cfg: cartão Formatar > Filtros; periodo: texto do rodapé do painel; medida: nome da medida (mensagem de lista vazia)
    constructor(private host: IVisualHost, private raiz: HTMLElement, private redesenhar: () => void,
        private cfg: () => FiltrosCard, private periodo: () => string, private medida: string) { }

    // h[ini] = dataRef (yyyy-mm-dd); h[ini + 1 + lista] = valores separados por "^"
    public lerCabecalho(h: string[], ini: number): void {
        this.dataRef = h[ini] || "";
        CAMPOS.forEach(cp => this.listas.set(cp.id, (h[ini + 1 + cp.lista] || "").split("^").filter(x => x !== "")));
    }

    // filtros vindos do Power BI (marcador, reabertura do relatório): reconstrói a seleção aplicada
    public lerFiltros(fs: FiltroBasico[]): void {
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
        const txt = (this.cfg()[cp.col] as { value: string }).value;
        const a = lerAlvos(txt);
        return a.length ? a : lerAlvos(COLUNAS_PADRAO[cp.col]);
    }

    public ativos(): number {
        let n = 0;
        this.aplicados.forEach(s => { if (s.size) { n++; } });
        return n;
    }

    // ---------- botão ----------
    public botao(cls: string, aoAbrir: () => void): HTMLElement {
        const n = this.ativos();
        const b = botao(cls + " fl-btn" + (n ? " ativo" : ""), "Filtros", () => {
            if (this.popAberto) { this.fechar(); return; }
            aoAbrir();
            this.pendentes = new Map();
            this.aplicados.forEach((s, k) => this.pendentes.set(k, new Set(s)));
            this.popAberto = true;
            this.redesenhar();
            const busca = this.raiz.querySelector(".fp-busca") as HTMLInputElement | null;
            if (busca) { busca.focus(); }
        });
        b.setAttribute("aria-haspopup", "dialog");
        b.setAttribute("aria-expanded", String(this.popAberto));
        if (n) {
            b.appendChild(el("span", "fl-n", String(n)));
            b.setAttribute("aria-label", `Filtros (${n} ativo${n > 1 ? "s" : ""})`);
        }
        this.btn = b;
        return b;
    }

    public fechar(): void {
        this.popAberto = false;
        this.redesenhar();
        if (this.btn) { this.btn.focus(); }
    }

    private qtdPendente(id: string): number { const s = this.pendentes.get(id); return s ? s.size : 0; }

    // ---------- painel ----------
    public painel(): HTMLElement {
        const p = el("div", "fp");
        p.setAttribute("role", "dialog");
        p.setAttribute("aria-label", "Filtros da página");
        const cab = el("div", "fp-cab");
        cab.appendChild(el("div", "fp-tit", "Filtros"));
        cab.appendChild(botao("fp-link", "Limpar tudo", () => { this.pendentes = new Map(); this.aplicar(); }));
        const x = botao("fp-x", "×", () => this.fechar());
        x.setAttribute("aria-label", "Fechar filtros");
        cab.appendChild(x);
        p.appendChild(cab);

        const corpo = el("div", "fp-corpo");
        const nav = el("div", "fp-nav");
        nav.setAttribute("role", "tablist");
        nav.setAttribute("aria-label", "Campos");
        CAMPOS.forEach(cp => {
            const b = botao("fp-campo" + (cp.id === this.campoAtivo ? " sel" : ""), cp.rotulo, () => {
                this.campoAtivo = cp.id; this.redesenhar();
                const s = this.raiz.querySelector(".fp-busca") as HTMLInputElement | null;
                if (s) { s.focus(); }
            });
            b.setAttribute("role", "tab");
            b.setAttribute("aria-selected", String(cp.id === this.campoAtivo));
            b.dataset.campo = cp.id;
            const q = this.qtdPendente(cp.id);
            b.appendChild(el("span", "fp-q" + (q ? "" : " zero"), q ? String(q) : ""));
            nav.appendChild(b);
        });
        corpo.appendChild(nav);
        const dica = el("div", "fp-dica");
        corpo.appendChild(this.painelCampo(CAMPOS.find(c => c.id === this.campoAtivo) || CAMPOS[0], nav, () => this.atualizarDica(dica)));
        p.appendChild(corpo);

        const rod = el("div", "fp-rod");
        const txt = el("div", "fp-per");
        txt.appendChild(el("div", "", this.periodo()));
        txt.appendChild(dica);
        this.atualizarDica(dica);
        rod.appendChild(txt);
        rod.appendChild(botao("fp-aplicar", "Aplicar", () => this.aplicar()));
        p.appendChild(rod);
        return p;
    }

    // A [Data Ref] só muda com um Dia, um Mês e um Ano: avisa quando a seleção de datas não forma uma data
    private atualizarDica(d: HTMLElement): void {
        const q = ["ano", "mes", "dia"].map(k => this.qtdPendente(k));
        const algum = q.some(n => n > 0), umaData = q.every(n => n === 1);
        d.textContent = algum && !umaData ? "Para mudar a data do painel, escolha um Dia, um Mês e um Ano." : "";
    }

    private painelCampo(cp: Campo, nav: HTMLElement, mudou: () => void): HTMLElement {
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
            mudou();
        };
        let visiveis: string[] = [];
        const desenharLista = (): void => {
            while (lista.firstChild) { lista.removeChild(lista.firstChild); }
            const termo = semAcento(busca.value.trim());
            visiveis = termo ? valores.filter(v => semAcento(v).indexOf(termo) >= 0) : valores;
            if (!valores.length) { lista.appendChild(el("div", "fp-msg", `Sem valores: confira a medida ${this.medida}.`)); return; }
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

    public montar(): FiltroBasico[] {
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
        const fs = this.montar();
        const carga = (fs.length ? fs : null) as unknown as powerbi.IFilter[];
        // "filter" filtra os outros visuais da página (como uma segmentação); "selfFilter" filtra este visual
        this.host.applyJsonFilter(carga, "general", "filter", powerbi.FilterAction.merge);
        try {
            this.host.applyJsonFilter(this.cfg().filtrarVisual.value ? carga : null as unknown as powerbi.IFilter[], "general", "selfFilter", powerbi.FilterAction.merge);
        } catch { /* host sem selfFilter */ }
        this.popAberto = false;
        this.redesenhar();
        if (this.btn) { this.btn.focus(); }
    }

    // texto para o rodapé do Excel
    public resumo(): string {
        const p: string[] = [];
        CAMPOS.forEach(cp => {
            const s = this.aplicados.get(cp.id);
            if (!s || !s.size) { return; }
            const v = Array.from(s);
            p.push(cp.rotulo + " = " + (v.length > 4 ? v.slice(0, 3).join(", ") + ` e mais ${v.length - 3}` : v.join(", ")));
        });
        return p.length ? "Filtros: " + p.join("; ") : "";
    }
}
