// Leitura da medida "Painel Metas Rep (visual)": uma célula de texto por representante (linha do visual), registros separados por quebra de linha:
//   R|codigo|nome|uf|objetivoClientes|clientesV|clientesF|pct
//   F|codigo|ordem|familia|12 valores        (ordem 0 = TOTAL do bloco)
// No 1º representante vem o cabeçalho depois de "~": atualizacao|dataRef|anos|meses|dias|supervisoes|representantes|ufs|familias|materiais
// (a célula pode trazer só o cabeçalho, começando com "~", quando esse representante não tem meta).
import { BlocoRep, LinhaRep, N } from "./xlsxMetasRep";

export interface Carga { cabecalho: string[]; blocos: BlocoRep[]; }
// "ok"; "outra" = medida do visual Painel Metas (por UF), "V|1~..." / "B|..."; "errada" = outra medida qualquer
export type Formato = "ok" | "outra" | "errada";

function n(s: string | undefined): N {
    if (s === undefined || s === "") { return null; }
    const v = Number(s);
    return isNaN(v) ? null : v;
}
const NUM = /^-?\d+(\.\d+)?$/;

// confere a 1ª célula que tem linha de dados (a do 1º representante pode trazer só o cabeçalho, que começa com "~")
export function conferir(celulas: string[]): Formato {
    const amostra = celulas.map(x => x.split("~")[0]).find(x => x.trim().length > 0);
    if (!amostra) { return celulas.some(x => x.indexOf("~") >= 0) ? "ok" : "errada"; }
    const linhas = amostra.split(/\r?\n/).filter(x => !!x);
    const c = linhas[0].split("|");
    if (c[0] === "V" || c[0] === "B") { return "outra"; }
    if (c[0] !== "R" || c.length !== 8 || !c.slice(4).every(x => x === "" || NUM.test(x))) { return "errada"; }
    const f = (linhas[1] || "").split("|");
    if (linhas.length > 1 && (f[0] !== "F" || f.length !== 16)) { return "errada"; }
    return "ok";
}

export function lerCarga(celulas: string[]): Carga {
    let cabecalho: string[] = [];
    const porCod = new Map<string, BlocoRep & { ordem: Map<number, LinhaRep> }>();
    const bloco = (cod: string): BlocoRep & { ordem: Map<number, LinhaRep> } => {
        let b = porCod.get(cod);
        if (!b) { b = { cod, nome: "", uf: "", k: [], total: { nome: "TOTAL", v: [] }, fams: [], ordem: new Map() }; porCod.set(cod, b); }
        return b;
    };
    celulas.forEach(cel => {
        (cel || "").split(/\r?\n/).forEach(l => {
            if (!l) { return; }
            const til = l.indexOf("~");
            const corpo = til >= 0 ? l.slice(0, til) : l;
            if (til >= 0) { cabecalho = l.slice(til + 1).split("|"); }
            const c = corpo.split("|");
            if (c[0] === "R" && c.length >= 8) {
                const b = bloco(c[1]);
                b.nome = c[2]; b.uf = c[3]; b.k = c.slice(4, 8).map(n);
            } else if (c[0] === "F" && c.length >= 16) {
                const b = bloco(c[1]);
                const lin: LinhaRep = { nome: c[3], v: c.slice(4, 16).map(n) };
                const o = Number(c[2]);
                if (o === 0) { b.total = lin; } else { b.ordem.set(o, lin); }
            }
        });
    });
    const blocos = Array.from(porCod.values()).filter(b => !!b.nome).map(b => ({
        cod: b.cod, nome: b.nome, uf: b.uf, k: b.k, total: b.total,
        fams: Array.from(b.ordem.keys()).sort((x, y) => x - y).map(o => b.ordem.get(o) as LinhaRep)
    }));
    // como a aba: representantes em ordem alfabética do nome
    blocos.sort((a, b) => a.nome.localeCompare(b.nome, "pt-BR") || a.cod.localeCompare(b.cod));
    return { cabecalho, blocos };
}

// "2026-11-25 09:42:20" -> "*Última atualização: 25/11/26 - 09h42min20s GMT-03:00"
export function textoAtualizacao(iso: string): string {
    const m = /^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2}):(\d{2})/.exec(iso || "");
    if (!m) { return ""; }
    return `*Última atualização: ${m[3]}/${m[2]}/${m[1].slice(2)} - ${m[4]}h${m[5]}min${m[6]}s GMT-03:00`;
}
