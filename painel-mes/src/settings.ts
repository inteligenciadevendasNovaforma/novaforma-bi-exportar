import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import Card = formattingSettings.SimpleCard;
import Model = formattingSettings.Model;

class TituloCard extends Card {
    name = "titulo";
    displayName = "Título";
    mostrar = new formattingSettings.ToggleSwitch({ name: "mostrar", displayName: "Mostrar", value: true });
    texto = new formattingSettings.TextInput({ name: "texto", displayName: "Texto", value: "PAINEL DE VENDAS - MÊS", placeholder: "Título" });
    slices = [this.mostrar, this.texto];
}

class MedidoresCard extends Card {
    name = "medidores";
    displayName = "Medidores";
    inicioAmarelo = new formattingSettings.NumUpDown({ name: "inicioAmarelo", displayName: "Início do amarelo (%)", value: 55 });
    inicioVerde = new formattingSettings.NumUpDown({ name: "inicioVerde", displayName: "Início do verde (%)", value: 72 });
    slices = [this.inicioAmarelo, this.inicioVerde];
}

class FundoCard extends Card {
    name = "fundo";
    displayName = "Fundo do visual";
    mostrar = new formattingSettings.ToggleSwitch({ name: "mostrar", displayName: "Mostrar", value: true });
    cor = new formattingSettings.ColorPicker({ name: "cor", displayName: "Cor", value: { value: "#FFFFFF" } });
    slices = [this.mostrar, this.cor];
}

// Colunas do modelo que cada campo do botão Filtros filtra. "Tabela[Coluna]"; mais de uma separada por ";".
export const COLUNAS_PADRAO = {
    colAno: "d_Calendario[Ano]",
    colMes: "d_Calendario[Mês]",
    colDia: "d_Calendario[Dia]",
    colSup: "d_Representantes[Supervisão]",
    colRep: "d_Representantes[Nome do Rep]",
    colFam: "f_Faturamento[Família]; f_Carteira[Família]",
    colMat: "f_Faturamento[Material]; f_Carteira[Material]",
    colUf: "d_Representantes[UF]"
};

export class FiltrosCard extends Card {
    name = "filtros";
    displayName = "Filtros";
    mostrarBotao = new formattingSettings.ToggleSwitch({ name: "mostrarBotao", displayName: "Mostrar botão Filtros", value: true });
    filtrarVisual = new formattingSettings.ToggleSwitch({ name: "filtrarVisual", displayName: "Filtrar também este visual", value: true });
    colAno = new formattingSettings.TextInput({ name: "colAno", displayName: "Ano (Tabela[Coluna])", value: COLUNAS_PADRAO.colAno, placeholder: COLUNAS_PADRAO.colAno });
    colMes = new formattingSettings.TextInput({ name: "colMes", displayName: "Mês (Tabela[Coluna])", value: COLUNAS_PADRAO.colMes, placeholder: COLUNAS_PADRAO.colMes });
    colDia = new formattingSettings.TextInput({ name: "colDia", displayName: "Dia (Tabela[Coluna])", value: COLUNAS_PADRAO.colDia, placeholder: COLUNAS_PADRAO.colDia });
    colSup = new formattingSettings.TextInput({ name: "colSup", displayName: "Supervisão (Tabela[Coluna])", value: COLUNAS_PADRAO.colSup, placeholder: COLUNAS_PADRAO.colSup });
    colRep = new formattingSettings.TextInput({ name: "colRep", displayName: "Representante (Tabela[Coluna])", value: COLUNAS_PADRAO.colRep, placeholder: COLUNAS_PADRAO.colRep });
    colFam = new formattingSettings.TextInput({ name: "colFam", displayName: "Família (Tabela[Coluna]; ...)", value: COLUNAS_PADRAO.colFam, placeholder: COLUNAS_PADRAO.colFam });
    colMat = new formattingSettings.TextInput({ name: "colMat", displayName: "Material (Tabela[Coluna]; ...)", value: COLUNAS_PADRAO.colMat, placeholder: COLUNAS_PADRAO.colMat });
    colUf = new formattingSettings.TextInput({ name: "colUf", displayName: "UF (Tabela[Coluna])", value: COLUNAS_PADRAO.colUf, placeholder: COLUNAS_PADRAO.colUf });
    slices = [this.mostrarBotao, this.filtrarVisual, this.colAno, this.colMes, this.colDia, this.colSup, this.colRep, this.colFam, this.colMat, this.colUf];
}

class ExportacaoCard extends Card {
    name = "exportacao";
    displayName = "Exportação";
    mostrarBotao = new formattingSettings.ToggleSwitch({ name: "mostrarBotao", displayName: "Mostrar botão Exportar", value: true });
    nomeArquivo = new formattingSettings.TextInput({ name: "nomeArquivo", displayName: "Nome do arquivo Excel", value: "painel_mes", placeholder: "nome do arquivo" });
    usarPagina = new formattingSettings.ToggleSwitch({ name: "usarPagina", displayName: "Baixar pela página de download (link)", value: true });
    urlPagina = new formattingSettings.TextInput({ name: "urlPagina", displayName: "Endereço da página de download", value: "https://gabrielpms2001.github.io/novaforma-exportar-bi/download-page/", placeholder: "https://.../download-page/" });
    slices = [this.mostrarBotao, this.nomeArquivo, this.usarPagina, this.urlPagina];
}

export class VisualSettings extends Model {
    titulo = new TituloCard();
    medidores = new MedidoresCard();
    fundo = new FundoCard();
    filtros = new FiltrosCard();
    exportacao = new ExportacaoCard();
    cards = [this.titulo, this.medidores, this.fundo, this.filtros, this.exportacao];
}
