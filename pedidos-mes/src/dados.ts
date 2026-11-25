// Leitura da medida "Pedidos Mês (visual)": cada linha do visual (um representante) traz várias linhas de texto, uma por
// cliente x família:  codRep|representante|codCliente|razaoSocial|ordemFamilia|familia|vendido
// No 1º representante vem o cabeçalho depois de "~":  dataRef|anos|meses|dias|supervisoes|representantes|ufs|familias|materiais
import { RepPM, CliPM } from "./xlsxPedidos";

export type Ordem = "vendido" | "nome" | "codigo";
export interface Carga { cabecalho: string[]; reps: RepPM[]; linhas: number; }

const numCod = (s: string): number => (/^\d+$/.test(s) ? Number(s) : Number.MAX_SAFE_INTEGER);

export function lerCarga(celulas: string[], ordem: Ordem): Carga {
    let cabecalho: string[] = [];
    const porRep = new Map<string, RepPM & { porCli: Map<string, CliPM> }>();
    let linhas = 0;
    celulas.forEach(cel => {
        (cel || "").split(/\r?\n/).forEach(l => {
            if (!l) { return; }
            const til = l.indexOf("~");
            const corpo = til >= 0 ? l.slice(0, til) : l;
            if (til >= 0) { cabecalho = l.slice(til + 1).split("|"); }
            const c = corpo.split("|");
            if (c.length < 7) { return; }
            const [codRep, nome, cli, razao, ord, fam, val] = c;
            const v = Number(val);
            let rep = porRep.get(codRep);
            if (!rep) { rep = { cod: codRep, nome, total: 0, clis: [], porCli: new Map() }; porRep.set(codRep, rep); }
            let cl = rep.porCli.get(cli);
            if (!cl) { cl = { cod: cli, razao, fams: [] }; rep.porCli.set(cli, cl); rep.clis.push(cl); }
            const vv = isNaN(v) ? 0 : v;
            cl.fams.push({ ord: Number(ord) || 99, nome: fam, v: vv });
            rep.total += vv;
            linhas++;
        });
    });
    const reps: RepPM[] = Array.from(porRep.values()).map(r => {
        r.clis.sort((a, b) => numCod(a.cod) - numCod(b.cod) || a.cod.localeCompare(b.cod));
        r.clis.forEach(cl => cl.fams.sort((a, b) => a.ord - b.ord));
        return { cod: r.cod, nome: r.nome, total: r.total, clis: r.clis };
    });
    const porNome = (a: RepPM, b: RepPM): number => a.nome.localeCompare(b.nome, "pt-BR");
    if (ordem === "nome") { reps.sort(porNome); }
    else if (ordem === "codigo") { reps.sort((a, b) => numCod(a.cod) - numCod(b.cod) || porNome(a, b)); }
    else { reps.sort((a, b) => b.total - a.total || porNome(a, b)); }
    return { cabecalho, reps, linhas };
}
