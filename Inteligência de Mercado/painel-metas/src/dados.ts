// Leitura da medida "Painel Metas (visual)": registros separados por quebra de linha.
//   V|1~atualizacao|dataRef|listas...         (1ª linha: cabeçalho depois do "~")
//   B|uf|objCli|cliV|cliF|pct|cliCart|tendFC|prazo|mix
//   F|uf|ordem|familia|12 valores|dias ("d:v,...")   ordem 0 = TOTAL do bloco
//   C|uf|dias (clientes por dia)
import { BlocoMetas, LinhaMetas, N } from "./xlsxMetas";

export interface Carga { cabecalho: string[]; blocos: BlocoMetas[]; }

function n(s: string | undefined): N {
    if (s === undefined || s === "") { return null; }
    const v = Number(s);
    return isNaN(v) ? null : v;
}
function dias(s: string | undefined): N[] {
    const a: N[] = [];
    for (let i = 0; i < 31; i++) { a.push(null); }
    (s || "").split(",").forEach(p => {
        const [d, v] = p.split(":");
        const di = Number(d), vv = Number(v);
        if (di >= 1 && di <= 31 && v !== undefined && v !== "" && !isNaN(vv)) { a[di - 1] = vv; }
    });
    return a;
}

export function lerCarga(texto: string): Carga {
    const cabecalho: string[] = [];
    const porUf = new Map<string, BlocoMetas & { ordem: Map<number, LinhaMetas> }>();
    const ordemUf: string[] = [];
    const bloco = (u: string): BlocoMetas & { ordem: Map<number, LinhaMetas> } => {
        let b = porUf.get(u);
        if (!b) {
            b = { uf: u === "*" ? "" : u, k: [], total: { nome: "TOTAL", v: [], d: dias("") }, fams: [], cli: dias(""), ordem: new Map() };
            porUf.set(u, b); ordemUf.push(u);
        }
        return b;
    };
    (texto || "").split(/\r?\n/).forEach(linha => {
        if (!linha) { return; }
        const [corpo, cab] = linha.split("~");
        if (cab !== undefined) { cabecalho.splice(0, cabecalho.length, ...cab.split("|")); }
        const c = corpo.split("|");
        if (c[0] === "B" && c.length >= 10) {
            bloco(c[1]).k = c.slice(2, 10).map(n);
        } else if (c[0] === "F" && c.length >= 17) {
            const b = bloco(c[1]);
            const l: LinhaMetas = { nome: c[3], v: c.slice(4, 16).map(n), d: dias(c[16]) };
            const o = Number(c[2]);
            if (o === 0) { b.total = l; } else { b.ordem.set(o, l); }
        } else if (c[0] === "C") {
            bloco(c[1]).cli = dias(c[2]);
        }
    });
    // total primeiro, depois as UFs em ordem alfabética
    const ufs = ordemUf.filter(u => u !== "*").sort((a, b) => a.localeCompare(b, "pt-BR"));
    const blocos = (porUf.has("*") ? ["*", ...ufs] : ufs).map(u => {
        const b = porUf.get(u) as BlocoMetas & { ordem: Map<number, LinhaMetas> };
        const fams = Array.from(b.ordem.keys()).sort((x, y) => x - y).map(o => b.ordem.get(o) as LinhaMetas);
        return { uf: b.uf, k: b.k, total: b.total, fams, cli: b.cli };
    });
    return { cabecalho, blocos };
}

// "2026-11-25 09:42:20" -> "*Última atualização: 25/11/26 - 09h42min20s GMT-03:00"
export function textoAtualizacao(iso: string): string {
    const m = /^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2}):(\d{2})/.exec(iso || "");
    if (!m) { return ""; }
    return `*Última atualização: ${m[3]}/${m[2]}/${m[1].slice(2)} - ${m[4]}h${m[5]}min${m[6]}s GMT-03:00`;
}
