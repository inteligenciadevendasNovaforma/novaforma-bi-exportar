// Configuração do visual: única diferença entre o Painel Vendas (Data Aprovação) e o Painel Vendas (Data do APP).
export const CONFIG = {
    tipo: "A",                                   // 1º campo de cada linha da medida (A = Data Aprovação, P = Data do APP)
    outroTipo: "P",
    medida: "Painel Vendas Aprovação (visual)",  // nome da medida de carga
    outraMedida: "Painel Vendas APP (visual)",
    arquivoDax: "Painel_Vendas_Aprovacao_medida.dax",
    regra: "Data Aprovação",                     // selo ao lado do título e rodapé do Excel
    aba: "(Data Aprovação) EVOLUÇÃO VENDA",      // nome da aba do Excel (igual ao exemplo)
    arquivo: "painel_vendas_aprovacao",          // nome padrão do arquivo Excel
    nomeVisual: "Painel Vendas (Data Aprovação)"
};
