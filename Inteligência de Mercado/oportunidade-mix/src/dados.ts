"use strict";
// Leitura da medida "Oportunidade Mix (visual)": uma célula de texto (a do 1º representante do visual) com todas as linhas.
//   Linhas separadas por quebra de linha; campos por "|":
//   codigo|nome|supervisao|uf|pedidos|vendido|vendido30|  (pedidos|vendido|ultimaVenda yyyymmdd) x 5 famílias (FB, FC, PT, AP, CT)
//   Depois de "~" vem o cabeçalho: MIX1|inicio|fim|hoje|temPeriodo|base|maxFaturamento|maxCarteira|minCalendario|supervisoes|ufs
//   (datas yyyy-mm-dd; listas separadas por "^"). As demais linhas do visual ficam vazias.
import { RepMix, FAMILIAS } from "./xlsxMix";

export interface Cabecalho {
    ini: string; fim: string; hoje: string; temPeriodo: boolean; base: number; fatMax: string; cartMax: string; calMin: string;
    sups: string[]; ufs: string[];
}
export interface Carga { cab: Cabecalho | null; reps: RepMix[]; }
// "ok"; "errada" = outra medida no campo (texto sem o cabeçalho MIX1 ou linhas com outro número de campos)
export type Formato = "ok" | "errada";

const NCAMPOS = 7 + FAMILIAS.length * 3;
const num = (s: string | undefined): number => { const n = Number(s); return isFinite(n) ? n : 0; };
const iso8 = (s: string): string => (/^\d{8}$/.test(s || "") ? `${s.slice(0, 4)}-${s.slice(4, 6)}-${s.slice(6)}` : "");

export function conferir(celulas: string[]): Formato {
    const c = celulas.find(x => x.indexOf("~MIX1|") >= 0);
    if (!c) { return "errada"; }
    const corpo = c.slice(0, c.indexOf("~MIX1|"));
    const l1 = corpo.split(/\r?\n/).find(x => !!x);
    if (l1 && l1.split("|").length !== NCAMPOS) { return "errada"; }
    return "ok";
}

export function lerCarga(celulas: string[]): Carga {
    const reps: RepMix[] = [];
    let cab: Cabecalho | null = null;
    const vistos = new Set<string>();
    celulas.forEach(cel => {
        const p = cel.indexOf("~MIX1|");
        const corpo = p >= 0 ? cel.slice(0, p) : cel;
        if (p >= 0) {
            const h = cel.slice(p + 1).split("|");
            cab = {
                ini: h[1] || "", fim: h[2] || "", hoje: h[3] || "", temPeriodo: h[4] === "1", base: num(h[5]),
                fatMax: h[6] || "", cartMax: h[7] || "", calMin: h[8] || "",
                sups: (h[9] || "").split("^").filter(x => !!x), ufs: (h[10] || "").split("^").filter(x => !!x)
            };
        }
        corpo.split(/\r?\n/).forEach(l => {
            const c = l.split("|");
            if (c.length < NCAMPOS || !c[0] || vistos.has(c[0])) { return; }
            vistos.add(c[0]);
            reps.push({
                cod: c[0], nome: c[1], sup: c[2], uf: c[3], ped: num(c[4]), vend: num(c[5]), v30: num(c[6]),
                f: FAMILIAS.map((_, k) => ({ n: num(c[7 + k * 3]), v: num(c[8 + k * 3]), u: iso8(c[9 + k * 3]) }))
            });
        });
    });
    return { cab, reps };
}
