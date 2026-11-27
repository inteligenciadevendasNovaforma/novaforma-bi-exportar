"use strict";

import "./../style/visual.less";
import powerbi from "powerbi-visuals-api";
import { CargaRanking, PaginaPdf, compactar, particionar, LIMITE_PAGINAS, PREFIXO_COLA, escalaImagem, ehIos, MSG_IMAGEM_GRANDE } from "./cargaRanking";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import { VisualSettings } from "./settings";

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import DataView = powerbi.DataView;

const MARKUP = `<div class="card" id="app">
<div class="head">
<div class="head-top">
<div class="brand"><svg class="logo" viewBox="0 0 189.938 35.752" role="img" aria-labelledby="logoT"><title id="logoT">NovaForma</title><g transform="translate(-334.975 -30.824)"><path d="M364.3,54.658a10.71,10.71,0,0,0,.495,3.738q.492,1.252,1.483,1.253c.658,0,1.154-.418,1.483-1.253a10.737,10.737,0,0,0,.494-3.738,10.344,10.344,0,0,0-.494-3.681c-.329-.813-.825-1.221-1.483-1.221s-1.155.408-1.483,1.221a10.317,10.317,0,0,0-.495,3.681m-8.233.133q0-5.624,2.677-8.706a10.724,10.724,0,0,1,15.047,0q2.654,3.083,2.654,8.706t-2.654,8.7a10.736,10.736,0,0,1-15.059,0q-2.664-3.083-2.665-8.7" fill="#fff"/><path d="M376.807,43.635h7.867l2.6,13.617,2.6-13.617h7.867l-6.234,22.31h-8.469Z" fill="#fff"/><path d="M442.671,54.658a10.79,10.79,0,0,0,.493,3.738q.5,1.252,1.485,1.253t1.484-1.253a10.794,10.794,0,0,0,.492-3.738,10.4,10.4,0,0,0-.492-3.681q-.5-1.219-1.484-1.221t-1.485,1.221a10.394,10.394,0,0,0-.493,3.681m-8.234.133q0-5.624,2.678-8.706a10.724,10.724,0,0,1,15.047,0q2.654,3.083,2.655,8.706t-2.655,8.7a10.735,10.735,0,0,1-15.058,0q-2.666-3.083-2.667-8.7" fill="#fff"/><path d="M457.073,65.944V43.634h7.479v3.115a7.857,7.857,0,0,1,2.528-2.636A7.67,7.67,0,0,1,470.529,43v7.865h-.171q-2.945,0-4.117,1.121t-1.172,4v9.956Z" fill="#fff"/><path d="M472.485,65.944V43.634h7.654v3.313a7.766,7.766,0,0,1,2.633-2.834,6.715,6.715,0,0,1,3.537-.915,5.08,5.08,0,0,1,3.126.926,5.246,5.246,0,0,1,1.8,2.823,6.687,6.687,0,0,1,2.363-2.834,6.376,6.376,0,0,1,3.484-.915,4.8,4.8,0,0,1,3.836,1.525,6.6,6.6,0,0,1,1.323,4.444V65.944h-7.975V52.959a2.631,2.631,0,0,0-.366-1.546,1.29,1.29,0,0,0-1.118-.5,1.225,1.225,0,0,0-1.074.577,2.849,2.849,0,0,0-.386,1.6V65.944h-7.932V52.959a2.647,2.647,0,0,0-.365-1.546,1.294,1.294,0,0,0-1.119-.5,1.273,1.273,0,0,0-1.108.577,2.79,2.79,0,0,0-.4,1.6V65.944Z" fill="#fff"/><path d="M352.592,44.745a4.745,4.745,0,0,0-3.8-1.5,6.767,6.767,0,0,0-3.547.9,7.769,7.769,0,0,0-2.622,2.8V43.635h-7.653V59.7a8.9,8.9,0,0,0,7.909,6.241h0V53.09a2.681,2.681,0,0,1,.419-1.589,1.314,1.314,0,0,1,1.128-.589,1.363,1.363,0,0,1,1.16.511,2.536,2.536,0,0,1,.388,1.536V65.944h7.91V49.168a6.588,6.588,0,0,0-1.3-4.423" fill="#fff"/><path d="M431.02,38.743a1.628,1.628,0,0,1,1.441-.686,3.225,3.225,0,0,1,.6.056,3.042,3.042,0,0,1,.582.162V31.544a13.838,13.838,0,0,0-2.107-.534,12.887,12.887,0,0,0-2.214-.186q-3.594,0-5.2,2T422.507,39.3v4.335h-2.428V50.78h2.428v8.808a9.062,9.062,0,0,0,8.04,6.354l.014,0V55.695h-.014V50.78h3.1V43.635h-3.1V40.823a3.7,3.7,0,0,1,.473-2.08" fill="#fff"/><path d="M516.262,58.33q-.506,1.231-1.472,1.232t-1.452-1.2a10.083,10.083,0,0,1-.483-3.572,10.041,10.041,0,0,1,.495-3.6q.494-1.219,1.44-1.219.968,0,1.472,1.231a9.847,9.847,0,0,1,.505,3.585,9.684,9.684,0,0,1-.505,3.539m.677-14.7v2.418a8.532,8.532,0,0,0-2.428-2.322A5.366,5.366,0,0,0,511.759,43a5.732,5.732,0,0,0-5.17,3.084q-1.834,3.081-1.838,8.705,0,5.575,1.86,8.682a5.744,5.744,0,0,0,5.148,3.1,5.312,5.312,0,0,0,2.741-.739,8.484,8.484,0,0,0,2.439-2.353v2.461h.068a8.9,8.9,0,0,0,7.906-6.241V43.635Z" fill="#fff"/><path d="M409.921,58.33q-.5,1.231-1.472,1.232T407,58.363a10.082,10.082,0,0,1-.483-3.572,10.017,10.017,0,0,1,.495-3.6q.492-1.219,1.44-1.219.967,0,1.472,1.231a9.846,9.846,0,0,1,.505,3.585,9.683,9.683,0,0,1-.505,3.539m.677-14.7v2.418a8.545,8.545,0,0,0-2.428-2.322A5.369,5.369,0,0,0,405.418,43a5.733,5.733,0,0,0-5.171,3.084q-1.836,3.081-1.837,8.705,0,5.575,1.86,8.682a5.741,5.741,0,0,0,5.148,3.1,5.306,5.306,0,0,0,2.74-.739,8.486,8.486,0,0,0,2.44-2.353v2.461h.068a8.9,8.9,0,0,0,7.906-6.241V43.635Z" fill="#fff"/></g></svg><div style="min-width:0"><div class="eyebrow" id="eyebrow"></div><div class="title" id="title"></div></div></div>
<div class="hbtns"><button class="ftoggle fexport" id="fexport" type="button" title="Exporta o ranking em PDF como está na tela (filtros, métrica e valores)"><svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M8 2v8M4.8 7.2 8 10.4l3.2-3.2M2.5 13h11"/></svg><span>Exportar PDF</span></button><button class="ftoggle fexport" id="fexportjpg" type="button" title="Exporta o ranking em uma imagem única (JPEG) como está na tela (filtros, métrica e valores)"><svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="2" y="2.5" width="12" height="11" rx="1.6"/><circle cx="5.7" cy="6" r="1.1"/><path d="M2.5 12l3.6-3.4 2.6 2.4 2-1.8 3.3 2.8"/></svg><span>Exportar JPEG</span></button><button class="ftoggle" id="ftoggle" type="button" aria-haspopup="dialog" aria-expanded="false"><svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" aria-hidden="true"><path d="M2 3.5h12M4.5 8h7M7 12.5h2"/></svg>Filtros</button></div>
</div>
<div class="stats" id="stats"></div>
</div>
<dialog id="fdialog" class="fdialog" aria-labelledby="fdlgTitle">
<div class="dlg-head"><h2 id="fdlgTitle">Filtros</h2><button class="x" id="fclose" type="button" aria-label="Fechar">&times;</button></div>
<div class="filters" id="filters">
<div class="fg wide"><label>Período</label><div class="seg" data-key="per"></div></div>
<div class="fg wide"><label>Métrica</label><div class="seg" data-key="met"></div></div>
<div class="fg wide"><label>Valores</label><div class="seg" data-key="val"></div></div>
<div class="ms-grid"><div class="ms-box" data-ms="fam"><div class="ms-head"><label>Família <span class="ms-count" hidden></span></label><button type="button" class="ms-clear">Limpar</button></div><input type="search" class="ms-search" placeholder="Pesquisar" aria-label="Pesquisar família"><div class="ms-list" role="group" aria-label="Opções de Família"></div></div><div class="ms-box" data-ms="sup"><div class="ms-head"><label>Supervisão <span class="ms-count" hidden></span></label><button type="button" class="ms-clear">Limpar</button></div><input type="search" class="ms-search" placeholder="Pesquisar" aria-label="Pesquisar supervisão"><div class="ms-list" role="group" aria-label="Opções de Supervisão"></div></div><div class="ms-box" data-ms="uf"><div class="ms-head"><label>UF <span class="ms-count" hidden></span></label><button type="button" class="ms-clear">Limpar</button></div><input type="search" class="ms-search" placeholder="Pesquisar" aria-label="Pesquisar uf"><div class="ms-list" role="group" aria-label="Opções de UF"></div></div></div>
</div>
<div class="dlg-foot"><button class="reset" id="reset" type="button">Limpar filtros</button><button class="done" id="fdone" type="button">Concluído</button></div>
</dialog>
<div class="body" id="list"></div>
<div class="note" id="note"></div>
<p class="rk-state" id="rk-state" role="status" aria-live="polite" hidden></p>
</div>`;

// Linha do dataset interno: [mes, familia, supervisao, uf, cod_rep, nome_rep, vendido, clientes, tipo]
// tipo 0 = por familia | tipo 1 = total do representante (clientes distintos exatos, sem familia)
interface Row { m: number; f: string; s: string; u: string; c: string; r: string; v: number; q: number; t: number; }
interface Rep { c: string; r: string; v: number; q: number; pos?: number; zebra?: number; }
interface Page { reps: Rep[]; isPos: boolean; showV: boolean; eyebrow: string; title: string; sub: string; foot: string; stats: { l: string; v: string }[]; }

const MS: { [k: string]: string } = { fam: "f", sup: "s", uf: "u" };
const SEGS: { [k: string]: [any, string][] } = {
    per: [[0, "Mês atual"], [1, "Mês anterior"]],
    met: [["Vendas", "Venda"], ["Positivação", "Positivação"]],
    val: [["Mostrar", "Mostrar valores"], ["Ocultar", "Ocultar valores"]]
};
const UFN: { [k: string]: string } = { AC: "Acre", AL: "Alagoas", AP: "Amapá", AM: "Amazonas", BA: "Bahia", CE: "Ceará", DF: "Distrito Federal", ES: "Espírito Santo", GO: "Goiás", MA: "Maranhão", MT: "Mato Grosso", MS: "Mato Grosso do Sul", MG: "Minas Gerais", PA: "Pará", PB: "Paraíba", PR: "Paraná", PE: "Pernambuco", PI: "Piauí", RJ: "Rio de Janeiro", RN: "Rio Grande do Norte", RS: "Rio Grande do Sul", RO: "Rondônia", RR: "Roraima", SC: "Santa Catarina", SP: "São Paulo", SE: "Sergipe", TO: "Tocantins" };
const LIMITE_LINK = 2500; // o Power BI recusa links muito longos
const BADGE = ["#E0A526", "#8A96A3", "#B9773E"];

const setHtml = (el: Element, html: string) => {
    // Converte o texto em nos DOM sem usar innerHTML (os textos dinamicos ja passam por esc()).
    const body = new DOMParser().parseFromString(html, "text/html").body;
    while (el.firstChild) el.removeChild(el.firstChild);
    Array.from(body.childNodes).forEach(n => el.appendChild(document.importNode(n, true)));
};
const normalize = (s: any) => String(s).normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase();
const esc = (s: any) => String(s).replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "\"": "&quot;", "'": "&#39;" } as any)[c]);
const brl = (n: number) => n.toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const int = (n: number) => n.toLocaleString("pt-BR");
const num = (s: any) => { const n = Number(s); return isFinite(n) ? n : 0; };
const joinLabel = (arr: string[], all: string) => !arr.length ? all : arr.length <= 2 ? arr.join(", ") : arr.length + " selecionadas";

export class Visual implements IVisual {
    private host: IVisualHost;
    private root: HTMLElement;
    private DATA: Row[] = [];
    private GRP: { [k: string]: number[] } = {};
    private META: any = { atual: "Mês atual", anterior: "Mês anterior", ini: [], fim: [], parcial: [false, false], tot: [] };
    private state: any = { per: 0, met: "Vendas", val: "Mostrar", fam: [], sup: [], uf: [] };
    private hasData = false;
    private metaOk = false;
    private busy = false;
    private fmtService = new FormattingSettingsService();
    private fmt: VisualSettings = new VisualSettings();

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.root = document.createElement("div");
        this.root.className = "rk";
        setHtml(this.root, MARKUP);
        options.element.appendChild(this.root);
        this.bind();
        this.buildSegs();
        this.render();
    }

    private $(id: string): any { return this.root.querySelector("#" + id); }

    // ---------- dados ----------
    public update(options: VisualUpdateOptions) {
        try { this.host.eventService.renderingStarted(options); } catch (e) { /* ambiente sem o servico */ }
        const dv: DataView = options.dataViews && options.dataViews[0];
        if (dv && dv.metadata && dv.metadata.segment) {
            try { this.host.fetchMoreData(true); } catch (e) { /* segue com o que ja chegou */ }
        }
        if (dv) this.fmt = this.fmtService.populateFormattingSettingsModel(VisualSettings, dv);
        this.ingest(dv);
        this.render();
        this.applyFormat();
        try { this.host.eventService.renderingFinished(options); } catch (e) { /* ambiente sem o servico */ }
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.fmtService.buildFormattingModel(this.fmt);
    }

    private applyFormat() {
        const f = this.fmt, t = f.titulo, h = f.cabecalho;
        const title = this.$("title") as HTMLElement, eyebrow = this.$("eyebrow") as HTMLElement;
        title.hidden = !t.mostrar.value;
        eyebrow.hidden = !t.mostrarEyebrow.value;
        title.style.fontFamily = t.fontFamily.value || "";
        title.style.fontSize = (t.fontSize.value || 22) + "px";
        title.style.fontWeight = t.bold.value ? "700" : "400";
        title.style.fontStyle = t.italic.value ? "italic" : "normal";
        title.style.color = t.cor.value.value;
        const head = this.root.querySelector(".head") as HTMLElement;
        head.style.background = h.mostrarFundo.value ? h.cor.value.value : "";
        (this.$("stats") as HTMLElement).hidden = !h.mostrarTotais.value;
        this.root.classList.toggle("nohbg", !h.mostrarFundo.value);
        this.root.classList.toggle("nobg", !f.fundo.mostrar.value);
        (this.root.querySelector(".card") as HTMLElement).style.background = f.fundo.mostrar.value ? f.fundo.cor.value.value : "transparent";
    }

    private ingest(dv: DataView) {
        const table: any = dv && dv.table;
        this.DATA = []; this.GRP = {}; this.hasData = false; this.metaOk = false;
        if (!table || !table.rows || !table.columns) return;
        let li = -1;
        table.columns.forEach((col: any, i: number) => { if (li < 0 && col.roles && col.roles.linha) li = i; });
        if (li < 0) return;
        const reps = new Map<string, any>();
        const ymd = (t: string) => t && t.length === 8 ? t.slice(6) + "/" + t.slice(4, 6) + "/" + t.slice(0, 4) : "";
        const mes = (t: string) => {
            const d = new Date(Number(t.slice(0, 4)), Number(t.slice(4, 6)) - 1, 1);
            const n = d.toLocaleDateString("pt-BR", { month: "long" });
            return n.charAt(0).toUpperCase() + n.slice(1) + "/" + d.getFullYear();
        };
        table.rows.forEach((row: any[]) => {
            const linha = row[li];
            if (linha === null || linha === undefined || linha === "") return;
            const p = String(linha).split("|");
            if (p.length < 20) return;
            const cod = p[0], nome = p[1] || p[0], sup = p[2], uf = p[3], fam = p[4];
            const v = [num(p[5]), num(p[6])], c = [num(p[7]), num(p[8])], cr = [num(p[9]), num(p[10])], cg = [num(p[11]), num(p[12])];
            [0, 1].forEach(m => { if (v[m] > 0 || c[m] > 0) this.DATA.push({ m, f: fam, s: sup, u: uf, c: cod, r: nome, v: v[m], q: c[m], t: 0 }); });
            const key = cod + "\u0001" + sup + "\u0001" + uf;
            const o = reps.get(key) || { cod, nome, sup, uf, v: [0, 0], cr: [0, 0] };
            o.v[0] += v[0]; o.v[1] += v[1]; o.cr = cr;
            reps.set(key, o);
            this.GRP[sup + "\u0001" + uf] = cg;
            if (!this.metaOk) {
                this.META = {
                    atual: mes(p[15]), anterior: mes(p[17]),
                    ini: [ymd(p[15]), ymd(p[17])], fim: [ymd(p[16]), ymd(p[18])],
                    parcial: [p[19] === "1", false], tot: [num(p[13]), num(p[14])]
                };
                this.metaOk = true;
            }
        });
        reps.forEach(o => {
            [0, 1].forEach(m => { if (o.v[m] > 0 || o.cr[m] > 0) this.DATA.push({ m, f: "", s: o.sup, u: o.uf, c: o.cod, r: o.nome, v: o.v[m], q: o.cr[m], t: 1 }); });
        });
        this.hasData = this.DATA.length > 0;
    }

    // ---------- filtros ----------
    private pass(d: Row, skip: string | null) {
        const s = this.state;
        return d.m === s.per && Object.keys(MS).every(k => k === skip || !s[k].length || s[k].includes((d as any)[MS[k]]));
    }
    private optionsOf(key: string) {
        return Array.from(new Set(this.DATA.filter(d => d.t === 0 && this.pass(d, key)).map(d => (d as any)[MS[key]]))).sort((a: string, b: string) => a.localeCompare(b, "pt-BR"));
    }
    private renderOptions(key: string, soft: boolean) {
        const wrap = this.root.querySelector('[data-ms="' + key + '"]') as HTMLElement;
        const list = wrap.querySelector(".ms-list") as HTMLElement, cnt = wrap.querySelector(".ms-count") as HTMLElement;
        if (!soft) {
            const opts = this.optionsOf(key);
            this.state[key] = this.state[key].filter((v: string) => opts.includes(v));
            const q = normalize((wrap.querySelector(".ms-search") as HTMLInputElement).value);
            const vis = opts.filter((v: string) => normalize(v).includes(q));
            const top = list.scrollTop;
            setHtml(list, vis.length ? vis.map((v: string) => `<label class="ms-opt"><input type="checkbox" value="${esc(v)}"${this.state[key].includes(v) ? " checked" : ""}><span>${esc(v)}</span></label>`).join("") : '<div class="ms-empty">Nenhum item encontrado.</div>');
            list.scrollTop = top;
        }
        cnt.hidden = !this.state[key].length; cnt.textContent = String(this.state[key].length);
    }
    private buildSegs() {
        this.root.querySelectorAll(".seg").forEach((box: any) => {
            const key = box.dataset.key;
            setHtml(box, SEGS[key].map(([v, l]) => `<button type="button" data-v="${esc(v)}">${esc(l)}</button>`).join(""));
            box.addEventListener("click", (e: any) => {
                const b = e.target.closest("button"); if (!b) return;
                this.state[key] = key === "per" ? Number(b.dataset.v) : b.dataset.v; this.render();
            });
        });
    }
    private bind() {
        Object.keys(MS).forEach(key => {
            const wrap = this.root.querySelector('[data-ms="' + key + '"]') as HTMLElement;
            (wrap.querySelector(".ms-list") as HTMLElement).addEventListener("change", (e: any) => {
                const b = e.target.closest("input"); if (!b) return;
                const set = new Set<string>(this.state[key]); b.checked ? set.add(b.value) : set.delete(b.value);
                this.state[key] = Array.from(set); this.render(key);
            });
            (wrap.querySelector(".ms-search") as HTMLElement).addEventListener("input", () => this.renderOptions(key, false));
            (wrap.querySelector(".ms-clear") as HTMLElement).addEventListener("click", () => { this.state[key] = []; this.render(); });
        });
        this.$("reset").addEventListener("click", () => {
            Object.assign(this.state, { per: 0, met: "Vendas", val: "Mostrar", fam: [], sup: [], uf: [] });
            Object.keys(MS).forEach(k => { (this.root.querySelector('[data-ms="' + k + '"] .ms-search') as HTMLInputElement).value = ""; });
            this.render();
        });
        const dlg = this.$("fdialog");
        const setOpen = (open: boolean) => {
            if (open) { if (!dlg.open) { typeof dlg.showModal === "function" ? dlg.showModal() : dlg.setAttribute("open", ""); } this.$("ftoggle").setAttribute("aria-expanded", "true"); this.$("fdone").focus(); }
            else if (dlg.open) { typeof dlg.close === "function" ? dlg.close() : dlg.removeAttribute("open"); }
        };
        this.$("ftoggle").addEventListener("click", () => setOpen(true));
        this.$("fclose").addEventListener("click", () => setOpen(false));
        this.$("fdone").addEventListener("click", () => setOpen(false));
        dlg.addEventListener("close", () => { this.$("ftoggle").setAttribute("aria-expanded", "false"); this.$("ftoggle").focus(); });
        dlg.addEventListener("click", (e: any) => { if (e.target === dlg) { const r = dlg.getBoundingClientRect(); if (e.clientX < r.left || e.clientX > r.right || e.clientY < r.top || e.clientY > r.bottom) setOpen(false); } });
        this.$("fexport").addEventListener("click", () => this.exportar("pdf"));
        this.$("fexportjpg").addEventListener("click", () => this.exportar("jpg"));
    }

    // ---------- ranking ----------
    private compute(extra?: (d: Row) => boolean): Rep[] {
        const map = new Map<string, Rep>();
        const useTot = !this.state.fam.length;
        this.DATA.forEach(d => {
            if (d.t !== (useTot ? 1 : 0)) return;
            if (!this.pass(d, null)) return;
            if (extra && !extra(d)) return;
            const o = map.get(d.c) || { c: d.c, r: d.r, v: 0, q: 0 };
            o.v += d.v; o.q += d.q; map.set(d.c, o);
        });
        const reps = Array.from(map.values()).filter(x => x.v > 0 || x.q > 0);
        const isPos = this.state.met === "Positivação";
        reps.sort((a, b) => isPos ? (b.q - a.q) || (b.v - a.v) || a.r.localeCompare(b.r, "pt-BR") : (b.v - a.v) || a.r.localeCompare(b.r, "pt-BR"));
        reps.forEach((x, i) => { const p = reps[i - 1]; const tie = p && (isPos ? (p.q === x.q && p.v === x.v) : p.v === x.v); x.pos = tie ? p.pos : i + 1; x.zebra = i; });
        return reps;
    }

    private botoes() { ["fexport", "fexportjpg"].forEach(id => { const exp = this.$(id) as HTMLButtonElement; if (exp) exp.disabled = this.busy || !this.hasData; }); }

    private render(changed?: string) {
        const st = this.state;
        Object.keys(MS).forEach(k => this.renderOptions(k, k === changed));
        this.root.querySelectorAll(".seg").forEach((box: any) => box.querySelectorAll("button").forEach((b: any) => b.classList.toggle("active", String(st[box.dataset.key]) === b.dataset.v)));
        this.botoes();
        const list = this.$("list"), note = this.$("note");
        if (!this.hasData) {
            this.$("eyebrow").textContent = ""; this.$("title").textContent = "Ranking de Representantes"; setHtml(this.$("stats"), "");
            setHtml(list, '<div class="setup">Arraste para o visual: <b>Representante</b> (Cod_Representante), <b>Família</b> (Desc_Familia) e a medida <b>Ranking Linha</b>.</div>');
            note.textContent = ""; return;
        }
        const reps = this.compute();
        const isPos = st.met === "Positivação", showV = st.val === "Mostrar";
        const cols = showV ? "40px minmax(0,1fr) 130px 72px" : "40px minmax(0,1fr) 72px";
        const cliOn = isPos || !showV;
        const mes = st.per === 0 ? this.META.atual : this.META.anterior;
        this.$("eyebrow").textContent = [joinLabel(st.sup, "Todas as supervisões"), st.uf.length ? joinLabel(st.uf, "") : ""].filter(Boolean).join(" · ") + (st.fam.length ? " · " + joinLabel(st.fam, "") : "");
        this.$("title").textContent = "Ranking de " + (isPos ? "Positivação" : "Vendas") + " · " + mes;
        const totV = reps.reduce((s, x) => s + x.v, 0), totC0 = reps.reduce((s, x) => s + x.q, 0);
        const exact = !st.fam.length && !st.sup.length && !st.uf.length && this.META.tot && typeof this.META.tot[st.per] === "number";
        const totC = exact ? this.META.tot[st.per] : totC0;
        const qtd = reps.filter(x => !/VENDA DIRETA/i.test(x.r)).length;
        setHtml(this.$("stats"), `<div class="stat ${!isPos ? "on" : ""}"><div class="lbl">Vendido total</div><div class="val">R$ ${brl(totV)}</div></div>` +
            `<div class="stat ${isPos ? "on" : ""}"><div class="lbl">Clientes positivados</div><div class="val">${int(totC)}</div></div>` +
            `<div class="stat"><div class="lbl">Representantes</div><div class="val">${qtd}</div></div>`);
        if (!reps.length) { setHtml(list, '<div class="empty">Sem vendas para a combinação de filtros selecionada.</div>'); note.textContent = ""; return; }
        const mxV = Math.max(...reps.map(x => x.v)), mxC = Math.max(...reps.map(x => x.q));
        const head = `<div class="row hd" style="grid-template-columns:${cols}"><span>Pos.</span><span>Representante</span>` + (showV ? `<span class="num ${isPos ? "" : "on"}">Vendido (R$)</span>` : "") + `<span class="num ${cliOn ? "on" : ""}">Clientes</span></div>`;
        const rows = reps.map(x => {
            const top = x.pos <= 3;
            const pct = Math.max(1.5, (isPos ? x.q / mxC : x.v / mxV) * 100);
            const bg = top ? "background:#fff;box-shadow:0 1px 3px rgba(11,42,74,.12)" : `background:${x.zebra % 2 ? "#EEF1F5" : "#F8FAFC"}`;
            const bd = top ? `background:${BADGE[x.pos - 1]};color:#fff` : "background:#DDE3EA;color:#0B2A4A";
            return `<div class="row" style="grid-template-columns:${cols};${bg}"><div class="badge" style="${bd}">${x.pos}º</div>` +
                `<div style="min-width:0"><div class="name" title="${esc(x.r)}">${esc(x.r)}</div><div class="bar"><i style="width:${pct.toFixed(1)}%;background:${top ? "#F2A23A" : "#0B2A4A"}"></i></div></div>` +
                (showV ? `<div class="num ${isPos ? "off" : "on"}">${brl(x.v)}</div>` : "") + `<div class="num ${cliOn ? "on" : "off"}">${int(x.q)}</div></div>`;
        }).join("");
        setHtml(list, head + rows);
        note.textContent = (isPos ? "Positivação = clientes com pedido no mês (desempate pelo valor vendido). " : "") + (!showV && !isPos ? "Barras proporcionais às vendas; valores ocultos. " : "") + "Valor de produtos dos pedidos emitidos no mês, sem itens recusados." + (st.fam.length > 1 ? " Com mais de uma família, os clientes são somados por família: um cliente que comprou em duas famílias conta duas vezes." : "");
    }

    // ---------- PDF ----------
    // Mesma lista, ordem, filtros e modo da tela: um único documento (sem separar por supervisão ou UF)
    private pdfDoc(): Page | null {
        const st = this.state, isPos = st.met === "Positivação", fam = st.fam.length > 0, showV = st.val === "Mostrar";
        const reps = this.compute();
        if (!reps.length) return null;
        const M = this.META, mes = st.per === 0 ? M.atual : M.anterior;
        const ini = (M.ini && M.ini[st.per]) || "", fim = (M.fim && M.fim[st.per]) || "";
        const parcial = !!(M.parcial && M.parcial[st.per]);
        const periodo = ini && fim ? "de " + ini.slice(0, 5) + " a " + fim : "no mês";
        const totV = reps.reduce((a, x) => a + x.v, 0), totC0 = reps.reduce((a, x) => a + x.q, 0);
        const exact = !fam && !st.sup.length && !st.uf.length && M.tot && typeof M.tot[st.per] === "number";
        const totC = exact ? M.tot[st.per] : totC0;
        const qtd = reps.filter(x => !/VENDA DIRETA/i.test(x.r)).length;
        const eyebrow = ([joinLabel(st.sup, "Todas as supervisões"), st.uf.length ? joinLabel(st.uf, "") : ""].filter(Boolean).join(" · ") + (fam ? " · " + joinLabel(st.fam, "") : "")).toUpperCase();
        const sub = mes + (parcial && fim ? " · parcial até " + fim.slice(0, 5) : "");
        const famTxt = fam ? "Família: " + joinLabel(st.fam, "") + ". " : "";
        const foot = isPos
            ? "Positivação = clientes distintos com pedido emitido " + periodo + ", sem itens recusados; empate desfeito pelo valor vendido. " + (exact ? "O total conta cada cliente uma vez, mesmo que tenha comprado de mais de um representante. " : "O total soma os clientes de cada representante" + (st.fam.length > 1 ? " e de cada família" : "") + ". ") + famTxt + "Fonte: SAP · Power BI Comercial."
            : "Valor de produtos dos pedidos emitidos " + periodo + ", sem itens recusados. Clientes = clientes distintos com pedido no período." + (!showV ? " Valores ocultos; barras proporcionais às vendas." : "") + " " + famTxt + (exact ? "O total de clientes conta cada um uma vez. " : "") + "Fonte: SAP · Power BI Comercial.";
        return {
            reps, isPos, showV, eyebrow, title: "Ranking de " + (isPos ? "Positivação" : "Vendas"), sub, foot: foot.replace(/\s{2,}/g, " "),
            stats: [{ l: "Vendido total", v: "R$ " + brl(totV) }, { l: "Clientes positivados", v: int(totC) }, { l: "Representantes", v: String(qtd) }]
        };
    }

    private timerMsg: number | undefined;
    private limparMsg() {
        if (this.timerMsg !== undefined) { window.clearTimeout(this.timerMsg); this.timerMsg = undefined; }
        const el = this.$("rk-state") as HTMLElement;
        el.hidden = true; el.textContent = ""; el.classList.remove("erro", "sai");
    }
    // Aviso que some sozinho (6 s; erros 10 s) com fade; sem animação se o usuário pediu menos movimento. Um novo aviso reinicia o tempo.
    private setMsg(text: string, erro = false) {
        this.limparMsg();
        if (!text) return;
        const el = this.$("rk-state") as HTMLElement;
        el.textContent = text; el.classList.toggle("erro", erro); el.hidden = false;
        this.timerMsg = window.setTimeout(() => {
            this.timerMsg = undefined;
            const reduz = typeof window.matchMedia === "function" && window.matchMedia("(prefers-reduced-motion: reduce)").matches;
            if (reduz) { this.limparMsg(); return; }
            el.classList.add("sai");
            this.timerMsg = window.setTimeout(() => { this.timerMsg = undefined; this.limparMsg(); }, 350);
        }, erro ? 10000 : 6000);
    }

    // Copia texto para a área de transferência: API do navegador e, se o iframe do Power BI bloquear, textarea + execCommand.
    private async copiar(texto: string): Promise<boolean> {
        try {
            const cb = navigator.clipboard;
            if (cb && cb.writeText) { await cb.writeText(texto); return true; }
        } catch (e) { /* bloqueado pela política do iframe: tenta o modo antigo */ }
        const ta = document.createElement("textarea");
        ta.value = texto; ta.setAttribute("readonly", ""); ta.className = "rk-copia";
        this.root.appendChild(ta);
        let ok = false;
        try { ta.focus(); ta.select(); ta.setSelectionRange(0, texto.length); ok = document.execCommand("copy"); } catch (e) { ok = false; }
        this.root.removeChild(ta);
        return ok;
    }

    // JSON -> deflate-raw -> base64url
    private async codificar(carga: CargaRanking): Promise<{ b64: string; comp: string }> {
        let dados: Uint8Array = new TextEncoder().encode(JSON.stringify(carga));
        let comp = "0";
        try {
            const CS = (window as any).CompressionStream;
            if (CS) {
                const cs = new CS("deflate-raw"), wr = cs.writable.getWriter();
                void wr.write(dados); void wr.close();
                dados = new Uint8Array(await new Response(cs.readable).arrayBuffer()); comp = "1";
            }
        } catch (e) { comp = "0"; dados = new TextEncoder().encode(JSON.stringify(carga)); }
        let s = ""; const step = 0x8000;
        for (let i = 0; i < dados.length; i += step) s += String.fromCharCode.apply(null, dados.subarray(i, i + step) as any);
        return { b64: btoa(s).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, ""), comp };
    }

    // Exportar PDF / JPEG: o visual só monta os dados; a página de download desenha o PDF (ou a imagem única) e baixa (funciona no Power BI embutido).
    // Dados no link quando cabem; senão, área de transferência + link curto (k=cola) e "Colar dados" na página.
    private async exportar(saida: "pdf" | "jpg") {
        if (this.busy) return;
        const jpg = saida === "jpg", rotulo = jpg ? "Exportar JPEG" : "Exportar PDF", oQue = jpg ? "A imagem" : "O PDF";
        const doc = this.pdfDoc();
        if (!doc) { this.setMsg("Nada para exportar com os filtros atuais.", true); return; }
        const npag = jpg ? 1 : particionar(doc.reps.length, 3).length; // estimativa (o rodapé pode ocupar 2 a 4 linhas)
        if (jpg && !escalaImagem(doc.reps.length, 4, ehIos())) { this.setMsg(MSG_IMAGEM_GRANDE, true); return; }
        if (npag > LIMITE_PAGINAS) { this.setMsg("Essa seleção geraria " + npag + " páginas. Filtre Supervisão ou UF antes de exportar.", true); return; }
        const url = String((this.fmt.exportacao && this.fmt.exportacao.urlPagina.value) || "").trim();
        if (!/^https:\/\//i.test(url)) { this.setMsg("Informe o endereço da página de download em Formatar > Exportação.", true); return; }
        this.busy = true; this.setMsg("");
        const btn = this.$(jpg ? "fexportjpg" : "fexport") as HTMLButtonElement, lab = btn.querySelector("span"); this.botoes();
        lab.textContent = "Preparando…";
        try {
            const mes = String(this.state.per === 0 ? this.META.atual : this.META.anterior).normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/[^A-Za-z0-9]+/g, "_");
            const name = "Ranking_" + (this.state.met === "Positivação" ? "Positivacao" : "Vendas") + "_" + mes + (jpg ? ".jpg" : ".pdf");
            const { b64, comp } = await this.codificar(compactar(doc as PaginaPdf, jpg ? "jpg" : undefined));
            const base = url.replace(/#.*$/, "") + "#";
            const alvo = base + new URLSearchParams({ v: "2", n: name, c: comp, j: b64 }).toString();
            const pags = npag === 1 ? "1 página" : npag + " páginas";
            if (alvo.length <= LIMITE_LINK) {
                this.host.launchUrl(alvo);
                this.setMsg("Página de download aberta: " + name + (jpg ? "." : " (" + pags + ")."));
                return;
            }
            const copiou = comp === "1" && await this.copiar(PREFIXO_COLA + b64);
            if (copiou) {
                this.host.launchUrl(base + new URLSearchParams({ v: "2", n: name, k: "cola", q: String(doc.reps.length), u: "representantes" }).toString());
                this.setMsg(oQue + (jpg ? "" : " (" + pags + ")") + " não cabe no link. Os dados foram copiados: na página que abriu, clique em Colar dados.");
            } else {
                this.setMsg(oQue + (jpg ? "" : " (" + pags + ")") + " não cabe no link e o Power BI bloqueou a cópia. Filtre Supervisão ou UF e exporte de novo.", true);
            }
        } catch (e) { this.setMsg("Não foi possível exportar " + (jpg ? "a imagem" : "o PDF") + ": " + (e && (e as any).message || e), true); }
        finally { this.busy = false; lab.textContent = rotulo; this.botoes(); }
    }
}
