"use strict";
// Período do Oportunidade de Mix: atalhos como no dashboard original (padrão "Neste mês"), datas no horário de São Paulo.
// O visual aplica o período como filtro de intervalo na coluna de data (d_Calendario[Date]); a medida lê MIN/MAX das datas filtradas.

export type Preset = "hoje" | "ontem" | "7d" | "mes" | "mesPassado" | "30d" | "intervalo";
export const ATALHOS: { id: Preset; rotulo: string }[] = [
    { id: "hoje", rotulo: "Hoje" }, { id: "ontem", rotulo: "Ontem" }, { id: "7d", rotulo: "Últimos 7 dias" },
    { id: "mes", rotulo: "Neste mês" }, { id: "mesPassado", rotulo: "Mês passado" }, { id: "30d", rotulo: "Últimos 30 dias" },
    { id: "intervalo", rotulo: "Intervalo personalizado" }
];
export const rotuloPreset = (p: string): string => (ATALHOS.find(a => a.id === p) || { rotulo: "" }).rotulo;

export function hojeSP(agora: Date = new Date()): string {
    try {
        return new Intl.DateTimeFormat("en-CA", { timeZone: "America/Sao_Paulo", year: "numeric", month: "2-digit", day: "2-digit" }).format(agora);
    } catch {
        const d = new Date(agora.getTime() - 3 * 3600000);
        return d.toISOString().slice(0, 10);
    }
}
export function somaDias(iso: string, dias: number): string {
    const d = new Date(iso + "T12:00:00Z");
    d.setUTCDate(d.getUTCDate() + dias);
    return d.toISOString().slice(0, 10);
}
// intervalo [início, fim] de cada atalho (intervalo personalizado usa as datas informadas)
export function intervalo(p: Preset, hoje: string, a = hoje, b = hoje): [string, string] {
    const prim = hoje.slice(0, 7) + "-01";
    switch (p) {
        case "hoje": return [hoje, hoje];
        case "ontem": return [somaDias(hoje, -1), somaDias(hoje, -1)];
        case "7d": return [somaDias(hoje, -6), hoje];
        case "30d": return [somaDias(hoje, -29), hoje];
        case "mesPassado": { const f = somaDias(prim, -1); return [f.slice(0, 7) + "-01", f]; }
        case "intervalo": return [a, b];
        default: return [prim, hoje];
    }
}
export function validar(a: string, b: string, hoje: string, min: string): string {
    const re = /^\d{4}-\d{2}-\d{2}$/;
    if (!re.test(a) || !re.test(b)) { return "Escolha a data inicial e a data final."; }
    if (a > b) { return "A data inicial precisa ser anterior ou igual à data final."; }
    if (b > hoje) { return "Escolha datas até hoje."; }
    if (min && a < min) { return "O calendário do modelo começa em " + dataBR(min) + "."; }
    return "";
}
export function dataBR(iso: string): string {
    const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso || "");
    return m ? `${m[3]}/${m[2]}/${m[1]}` : "";
}
