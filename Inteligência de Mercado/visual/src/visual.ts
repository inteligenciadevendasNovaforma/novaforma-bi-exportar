"use strict";
import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import "./../style/visual.less";
import { VisualSettings } from "./settings";
import { gerarXlsx, paraBase64, LinhaDado, OpcoesExport } from "./xlsx";

import IVisual = powerbi.extensibility.visual.IVisual;
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;

const PALETA = ["#0B2A4A", "#3E7CB1", "#F28C28", "#7FB7BE", "#8D99AE", "#C8553D"];
const NS_SVG = "http://www.w3.org/2000/svg";
const AZUL = "#0B2A4A";

function pad(n: number): string { return n < 10 ? "0" + n : String(n); }
function agora(): string {
    const d = new Date();
    return `${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}
function xe(s: string): string {
    return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
function fmt(n: number): string { return Math.round(n).toLocaleString("pt-BR"); }
function el(tag: string, cls?: string, texto?: string): HTMLElement {
    const e = document.createElement(tag);
    if (cls) { e.className = cls; }
    if (texto !== undefined) { e.textContent = texto; }
    return e;
}

interface Dados {
    meses: string[];
    unidades: string[];
    matriz: number[][];      // [unidade][mes]
    totMes: number[];
    totUni: number[];
    total: number;
    kpis: { rotulo: string; valor: string }[];
}

export class Visual implements IVisual {
    private host: IVisualHost;
    private raiz: HTMLElement;
    private cab: HTMLElement;
    private tituloEl: HTMLElement;
    private botoes: HTMLElement;
    private btnExportar!: HTMLButtonElement;
    private menu!: HTMLElement;
    private caixas: Partial<Record<keyof OpcoesExport, HTMLInputElement>> = {};
    private escolhas: OpcoesExport = { cards: true, graficos: true, matriz: true, dados: true };
    private padroes = "";
    private corpo: HTMLElement;
    private msg: HTMLElement;
    private fmtService = new FormattingSettingsService();
    private settings: VisualSettings = new VisualSettings();
    private linhas: LinhaDado[] = [];
    private d: Dados = { meses: [], unidades: [], matriz: [], totMes: [], totUni: [], total: 0, kpis: [] };
    private largura = 600;

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.raiz = el("div", "ex");
        this.cab = el("div", "ex-cab");
        this.tituloEl = el("div", "ex-titulo");
        this.botoes = el("div", "ex-botoes");
        this.btnExportar = el("button", "ex-btn forte", "Exportar ▾") as HTMLButtonElement;
        this.btnExportar.type = "button";
        this.btnExportar.setAttribute("aria-haspopup", "dialog");
        this.btnExportar.setAttribute("aria-expanded", "false");
        this.btnExportar.addEventListener("click", (ev) => { ev.stopPropagation(); this.alternarMenu(); });
        this.menu = el("div", "ex-menu");
        this.menu.setAttribute("role", "dialog");
        this.menu.setAttribute("aria-label", "Opções de exportação");
        this.menu.style.display = "none";
        this.menu.addEventListener("click", (ev) => ev.stopPropagation());
        this.menu.appendChild(el("div", "ex-menu-tit", "Incluir no Excel"));
        const itens: [keyof OpcoesExport, string][] = [["cards", "Cards (indicadores)"], ["graficos", "Gráficos (como imagem)"], ["matriz", "Matriz (células editáveis)"], ["dados", "Aba com dados brutos"]];
        for (const [chave, rotulo] of itens) {
            const lab = el("label", "ex-op");
            const cb = document.createElement("input");
            cb.type = "checkbox";
            cb.addEventListener("change", () => { this.escolhas[chave] = cb.checked; });
            this.caixas[chave] = cb;
            lab.append(cb, el("span", "", rotulo));
            this.menu.appendChild(lab);
        }
        const baixar = el("button", "ex-btn-go", "Baixar Excel") as HTMLButtonElement;
        baixar.type = "button";
        baixar.addEventListener("click", () => { void this.exportarExcel(); });
        this.menu.appendChild(baixar);
        this.botoes.append(this.btnExportar, this.menu);
        document.addEventListener("click", () => this.fecharMenu());
        document.addEventListener("keydown", (ev) => { if (ev.key === "Escape") { this.fecharMenu(); } });
        this.cab.append(this.tituloEl, this.botoes);
        this.corpo = el("div", "ex-corpo");
        this.msg = el("div", "ex-msg");
        this.raiz.append(this.cab, this.corpo, this.msg);
        options.element.appendChild(this.raiz);
    }

    public update(options: VisualUpdateOptions): void {
        const dv = options.dataViews && options.dataViews[0];
        this.settings = this.fmtService.populateFormattingSettingsModel(VisualSettings, dv);
        this.largura = options.viewport.width;
        if (dv && dv.metadata && dv.metadata.segment) { this.host.fetchMoreData(true); }
        this.lerDados(dv);
        this.calcular();
        const s = this.settings;
        this.raiz.style.background = s.fundo.mostrar.value ? s.fundo.cor.value.value : "transparent";
        this.tituloEl.style.visibility = s.titulo.mostrar.value ? "visible" : "hidden";
        this.tituloEl.textContent = s.titulo.texto.value || "";
        this.tituloEl.style.fontSize = Math.max(8, s.titulo.tamanho.value) + "px";
        this.tituloEl.style.color = s.titulo.cor.value.value;
        this.botoes.style.display = s.exportacao.mostrarBotao.value ? "flex" : "none";
        const pad = [s.exportacao.incluirCards.value, s.exportacao.incluirGraficos.value, s.exportacao.incluirMatriz.value, s.exportacao.incluirDados.value].join();
        if (pad !== this.padroes) {
            this.padroes = pad;
            this.escolhas = { cards: s.exportacao.incluirCards.value, graficos: s.exportacao.incluirGraficos.value, matriz: s.exportacao.incluirMatriz.value, dados: s.exportacao.incluirDados.value };
            this.sincronizarCaixas();
        }
        this.desenhar();
    }

    private lerDados(dv: powerbi.DataView | undefined): void {
        this.linhas = [];
        const tabela = dv && dv.table;
        if (!tabela) { return; }
        const idx = (papel: string) => tabela.columns.findIndex(c => c.roles && c.roles[papel]);
        const iMes = idx("mes"), iUni = idx("unidade"), iVal = idx("valor");
        if (iMes < 0 || iUni < 0 || iVal < 0) { return; }
        for (const r of tabela.rows) {
            const v = Number(r[iVal]);
            this.linhas.push({ mes: String(r[iMes] ?? ""), unidade: String(r[iUni] ?? ""), valor: isFinite(v) ? v : 0 });
        }
    }

    private calcular(): void {
        const meses = Array.from(new Set(this.linhas.map(l => l.mes)));
        const unidades = Array.from(new Set(this.linhas.map(l => l.unidade)));
        const matriz = unidades.map(u => meses.map(m =>
            this.linhas.filter(l => l.unidade === u && l.mes === m).reduce((a, l) => a + l.valor, 0)));
        const totMes = meses.map((_, j) => matriz.reduce((a, r) => a + r[j], 0));
        const totUni = matriz.map(r => r.reduce((a, b) => a + b, 0));
        const total = totUni.reduce((a, b) => a + b, 0);
        let variacao = "–";
        if (totMes.length >= 2 && totMes[totMes.length - 2] !== 0) {
            const v = (totMes[totMes.length - 1] - totMes[totMes.length - 2]) / totMes[totMes.length - 2] * 100;
            variacao = (v > 0 ? "+" : "") + v.toFixed(1).replace(".", ",") + "%";
        }
        const imax = totMes.indexOf(Math.max(...totMes));
        this.d = {
            meses, unidades, matriz, totMes, totUni, total,
            kpis: [
                { rotulo: "Total de devoluções", valor: fmt(total) },
                { rotulo: "Variação vs. mês anterior", valor: variacao },
                { rotulo: "Maior mês", valor: imax >= 0 ? meses[imax] : "–" }
            ]
        };
    }

    // ---------- gráficos (SVG) ----------
    private graficoMeses(w: number, h: number): string {
        const { meses, unidades, totMes } = this.d;
        const max = Math.max(1, ...totMes);
        const topo = 38, base = 46, esq = 44, dir = 12;
        const areaW = w - esq - dir, areaH = h - topo - base;
        const passo = areaW / Math.max(meses.length, 1);
        const larg = Math.min(60, passo * 0.62);
        let o = `<rect x="0.5" y="0.5" width="${w - 1}" height="${h - 1}" rx="6" fill="#FFFFFF" stroke="#E5E7EB"/>`;
        o += `<text x="12" y="22" font-size="13" font-weight="700" fill="${AZUL}">Devoluções por mês</text>`;
        for (let i = 0; i <= 4; i++) {
            const y = topo + areaH * (1 - i / 4);
            o += `<line x1="${esq}" x2="${w - dir}" y1="${y}" y2="${y}" stroke="#E5E7EB"/>`;
            o += `<text x="${esq - 6}" y="${y + 4}" text-anchor="end" font-size="11" fill="#6B7280">${Math.round(max * i / 4)}</text>`;
        }
        meses.forEach((m, i) => {
            const x = esq + passo * i + (passo - larg) / 2;
            let yAcc = topo + areaH;
            unidades.forEach((u, k) => {
                const v = this.d.matriz[k][i];
                const hh = areaH * v / max;
                yAcc -= hh;
                if (hh > 0) { o += `<rect x="${x}" y="${yAcc}" width="${larg}" height="${hh}" fill="${PALETA[k % PALETA.length]}"/>`; }
            });
            o += `<text x="${x + larg / 2}" y="${yAcc - 5}" text-anchor="middle" font-size="11" font-weight="600" fill="#111827">${Math.round(totMes[i])}</text>`;
            o += `<text x="${x + larg / 2}" y="${topo + areaH + 16}" text-anchor="middle" font-size="11" fill="#374151">${xe(m)}</text>`;
        });
        let lx = esq;
        unidades.forEach((u, k) => {
            o += `<rect x="${lx}" y="${h - 18}" width="9" height="9" fill="${PALETA[k % PALETA.length]}"/>`;
            o += `<text x="${lx + 14}" y="${h - 10}" font-size="11" fill="#374151">${xe(u)}</text>`;
            lx += 14 + u.length * 6.5 + 18;
        });
        return o;
    }

    private graficoUnidades(w: number, h: number): string {
        const { unidades, totUni } = this.d;
        const max = Math.max(1, ...totUni);
        const topo = 38, esq = 80, dir = 48;
        const n = Math.max(unidades.length, 1);
        const passo = (h - topo - 14) / n;
        const bh = Math.min(26, passo * 0.64);
        let o = `<rect x="0.5" y="0.5" width="${w - 1}" height="${h - 1}" rx="6" fill="#FFFFFF" stroke="#E5E7EB"/>`;
        o += `<text x="12" y="22" font-size="13" font-weight="700" fill="${AZUL}">Total por unidade</text>`;
        const ordem = unidades.map((u, i) => i).sort((a, b) => totUni[b] - totUni[a]);
        ordem.forEach((i, pos) => {
            const y = topo + passo * pos + (passo - bh) / 2;
            const len = (w - esq - dir) * totUni[i] / max;
            o += `<text x="${esq - 8}" y="${y + bh / 2 + 4}" text-anchor="end" font-size="11" fill="#374151">${xe(unidades[i])}</text>`;
            o += `<rect x="${esq}" y="${y}" width="${len}" height="${bh}" rx="3" fill="${PALETA[i % PALETA.length]}"/>`;
            o += `<text x="${esq + len + 6}" y="${y + bh / 2 + 4}" font-size="11" font-weight="600" fill="#111827">${Math.round(totUni[i])}</text>`;
        });
        return o;
    }

    private svgDoc(w: number, h: number, inner: string): string {
        return `<svg xmlns="${NS_SVG}" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}" font-family="Segoe UI, Arial, sans-serif"><rect width="${w}" height="${h}" fill="#FFFFFF"/>${inner}</svg>`;
    }

    // ---------- tela ----------
    private desenhar(): void {
        while (this.corpo.firstChild) { this.corpo.removeChild(this.corpo.firstChild); }
        if (!this.linhas.length) {
            this.corpo.appendChild(el("div", "ex-vazio", "Arraste Mês, Unidade e a medida de Devoluções para os campos do visual."));
            return;
        }
        const k = el("div", "ex-kpis");
        for (const i of this.d.kpis) {
            const c = el("div", "ex-kpi");
            c.append(el("span", "", i.rotulo), el("strong", "", i.valor));
            k.appendChild(c);
        }
        const g = el("div", "ex-graficos");
        const cw = Math.max(160, (this.largura - 24 - 8) / 2), ch = Math.round(cw * 0.5);
        for (const [fn, w, h] of [[() => this.graficoMeses(cw, ch), cw, ch], [() => this.graficoUnidades(cw, ch), cw, ch]] as [() => string, number, number][]) {
            const doc = new DOMParser().parseFromString(this.svgDoc(w, h, fn()), "image/svg+xml");
            g.appendChild(document.importNode(doc.documentElement, true));
        }
        const tab = el("table", "ex-tab");
        const cabTr = el("tr");
        for (const t of ["Unidade", ...this.d.meses, "Total"]) { cabTr.appendChild(el("th", "", t)); }
        tab.appendChild(cabTr);
        this.d.unidades.forEach((u, i) => {
            const tr = el("tr");
            tr.appendChild(el("td", "", u));
            this.d.matriz[i].forEach(v => tr.appendChild(el("td", "", fmt(v))));
            tr.appendChild(el("td", "", fmt(this.d.totUni[i])));
            tab.appendChild(tr);
        });
        const trT = el("tr", "tot");
        trT.appendChild(el("td", "", "Total"));
        this.d.totMes.forEach(v => trT.appendChild(el("td", "", fmt(v))));
        trT.appendChild(el("td", "", fmt(this.d.total)));
        tab.appendChild(trT);
        this.corpo.append(k, g, tab);
    }

    private mostrar(texto: string, erro = false): void {
        this.msg.textContent = texto;
        this.msg.className = "ex-msg" + (erro ? " erro" : "");
    }

    // ---------- imagens ----------
    private svgParaPng(svg: string, w: number, h: number): Promise<Blob> {
        return new Promise((resolve, reject) => {
            const img = new Image();
            img.onload = () => {
                const c = document.createElement("canvas");
                c.width = w * 2; c.height = h * 2;
                const ctx = c.getContext("2d");
                if (!ctx) { reject(new Error("canvas indisponível")); return; }
                ctx.fillStyle = "#FFFFFF";
                ctx.fillRect(0, 0, c.width, c.height);
                ctx.drawImage(img, 0, 0, c.width, c.height);
                c.toBlob(b => { if (b) { resolve(b); } else { reject(new Error("falha ao gerar a imagem")); } }, "image/png");
            };
            img.onerror = () => reject(new Error("falha ao desenhar a imagem"));
            img.src = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(svg);
        });
    }

    private blobParaDataUrl(b: Blob): Promise<string> {
        return new Promise((resolve, reject) => {
            const r = new FileReader();
            r.onload = () => resolve(String(r.result));
            r.onerror = () => reject(new Error("falha ao ler a imagem"));
            r.readAsDataURL(b);
        });
    }

    private async blobParaBytes(b: Blob): Promise<Uint8Array> {
        const url = await this.blobParaDataUrl(b);
        const bin = atob(url.split(",")[1]);
        const bytes = new Uint8Array(bin.length);
        for (let i = 0; i < bin.length; i++) { bytes[i] = bin.charCodeAt(i); }
        return bytes;
    }

    private async graficosPng(): Promise<{ blob: Blob; w: number; h: number }[]> {
        const w = 520, h = 300;
        const a = await this.svgParaPng(this.svgDoc(w, h, this.graficoMeses(w, h)), w, h);
        const b = await this.svgParaPng(this.svgDoc(w, h, this.graficoUnidades(w, h)), w, h);
        return [{ blob: a, w, h }, { blob: b, w, h }];
    }

    private sincronizarCaixas(): void {
        (Object.keys(this.escolhas) as (keyof OpcoesExport)[]).forEach(k => {
            const cb = this.caixas[k];
            if (cb) { cb.checked = this.escolhas[k]; }
        });
    }

    private alternarMenu(): void {
        const aberto = this.menu.style.display !== "none";
        if (aberto) { this.fecharMenu(); return; }
        this.sincronizarCaixas();
        this.menu.style.display = "block";
        this.btnExportar.setAttribute("aria-expanded", "true");
    }

    private fecharMenu(): void {
        this.menu.style.display = "none";
        this.btnExportar.setAttribute("aria-expanded", "false");
    }

    private baixar(blob: Blob, nome: string, aviso: string): void {
        const a = document.createElement("a");
        a.href = URL.createObjectURL(blob);
        a.download = nome;
        this.raiz.appendChild(a);
        a.click();
        this.raiz.removeChild(a);
        this.mostrar(aviso);
    }

    // ---------- Excel ----------
    private async exportarExcel(): Promise<void> {
        if (!this.linhas.length) { this.mostrar("Sem dados para exportar.", true); return; }
        const op: OpcoesExport = { ...this.escolhas };
        if (!op.cards && !op.graficos && !op.matriz && !op.dados) { this.mostrar("Marque pelo menos um item para exportar.", true); return; }
        this.fecharMenu();
        try {
            this.mostrar("Gerando Excel…");
            const url = (this.settings.exportacao.urlPagina.value || "").trim();
            const viaPagina = this.settings.exportacao.usarPagina.value && /^https:\/\//i.test(url);
            const base = (this.settings.exportacao.nomeArquivo.value || "painel").replace(/[^\w\-]+/g, "_");
            if (viaPagina) { await this.abrirPaginaDownload(url, base + ".xlsx", op); return; }
            const gr = op.graficos ? await this.graficosPng() : [];
            const imagens = await Promise.all(gr.map(async g => ({ png: await this.blobParaBytes(g.blob), w: g.w, h: g.h })));
            const bytes = gerarXlsx({
                titulo: this.settings.titulo.texto.value || "Painel", geradoEm: agora(),
                kpis: this.d.kpis, imagens, meses: this.d.meses, unidades: this.d.unidades,
                matriz: this.d.matriz, linhas: this.linhas, opcoes: op
            });
            const nome = base + ".xlsx";
            let permitido = false;
            try {
                const st = await this.host.downloadService.exportStatus();
                permitido = st === powerbi.PrivilegeStatus.Allowed;
            } catch { permitido = false; }
            if (permitido) {
                const r = await this.host.downloadService.exportVisualsContentExtended(paraBase64(bytes), nome, "base64", "Exportação do painel para Excel");
                this.mostrar(r && (r as { downloadCompleted?: boolean }).downloadCompleted === false ? "Download cancelado." : "Excel exportado: " + nome);
            } else {
                const blob = new Blob([bytes as BlobPart], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
                this.baixar(blob, nome, "Download direto (exportação do Power BI não liberada pelo administrador).");
            }
        } catch (e) { this.mostrar("Erro ao exportar: " + (e instanceof Error ? e.message : String(e)), true); }
    }

    // Abre a página de download levando só os dados (texto pequeno) no fragmento (#).
    // A página redesenha os gráficos e monta o .xlsx no navegador.
    private async abrirPaginaDownload(url: string, nome: string, op: OpcoesExport): Promise<void> {
        const w = 520, h = 300;
        const carga = {
            titulo: this.settings.titulo.texto.value || "Painel", geradoEm: agora(),
            kpis: this.d.kpis, meses: this.d.meses, unidades: this.d.unidades, matriz: this.d.matriz,
            linhas: op.dados ? this.linhas : [], opcoes: op,
            svgs: op.graficos ? [
                { svg: this.svgDoc(w, h, this.graficoMeses(w, h)), w, h },
                { svg: this.svgDoc(w, h, this.graficoUnidades(w, h)), w, h }
            ] : []
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
        for (let i = 0; i < dados.length; i += 0x8000) bin += String.fromCharCode.apply(null, Array.from(dados.subarray(i, i + 0x8000)));
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
