// Configuração do visual: única diferença entre o Painel Vendas (Data Aprovação) e o Painel Vendas (Data do APP).
export const CONFIG = {
    tipo: "P",                                   // 1º campo de cada linha da medida (A = Data Aprovação, P = Data do APP)
    outroTipo: "A",
    medida: "Painel Vendas APP (visual)",        // nome da medida de carga
    outraMedida: "Painel Vendas Aprovação (visual)",
    arquivoDax: "Painel_Vendas_APP_medida.dax",
    regra: "Data do APP",                        // selo ao lado do título e rodapé do Excel
    aba: "(Data APP) EVOLUÇÃO VENDA",            // nome da aba do Excel (igual ao exemplo)
    arquivo: "painel_vendas_app",                // nome padrão do arquivo Excel
    nomeVisual: "Painel Vendas (Data do APP)"
};
