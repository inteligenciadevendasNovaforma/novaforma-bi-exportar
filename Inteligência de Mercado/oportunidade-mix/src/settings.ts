"use strict";
// Aba Formatar do Oportunidade de Mix (NovaForma)
import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import Card = formattingSettings.SimpleCard;
import Model = formattingSettings.Model;

export const URL_PAGINA = "https://inteligenciadevendasnovaforma.github.io/novaforma-bi-exportar/download-page/";
export const COLUNA_DATA = "d_Calendario[Date]";

class TituloCard extends Card {
    name = "titulo";
    displayName = "Título";
    mostrar = new formattingSettings.ToggleSwitch({ name: "mostrar", displayName: "Mostrar título", value: true });
    texto = new formattingSettings.TextInput({ name: "texto", displayName: "Texto", value: "Oportunidade de Mix", placeholder: "Título" });
    subtitulo = new formattingSettings.TextInput({ name: "subtitulo", displayName: "Subtítulo", value: "Representantes sem venda por família de produto", placeholder: "Subtítulo" });
    slices = [this.mostrar, this.texto, this.subtitulo];
}

class CabecalhoCard extends Card {
    name = "cabecalho";
    displayName = "Cabeçalho";
    mostrar = new formattingSettings.ToggleSwitch({ name: "mostrar", displayName: "Mostrar faixa azul com logo", value: true });
    cor = new formattingSettings.ColorPicker({ name: "cor", displayName: "Cor da faixa", value: { value: "#0B2A4A" } });
    slices = [this.mostrar, this.cor];
}

class FundoCard extends Card {
    name = "fundo";
    displayName = "Fundo do visual";
    mostrar = new formattingSettings.ToggleSwitch({ name: "mostrar", displayName: "Mostrar fundo", value: true });
    cor = new formattingSettings.ColorPicker({ name: "cor", displayName: "Cor", value: { value: "#F2F4F7" } });
    slices = [this.mostrar, this.cor];
}

export class PeriodoCard extends Card {
    name = "periodoCfg";
    displayName = "Período";
    coluna = new formattingSettings.TextInput({ name: "coluna", displayName: "Coluna de data (Tabela[Coluna])", value: COLUNA_DATA, placeholder: COLUNA_DATA });
    soVisual = new formattingSettings.ToggleSwitch({ name: "soVisual", displayName: "Período filtra só este visual", value: true });
    slices = [this.coluna, this.soVisual];
}

class ExportacaoCard extends Card {
    name = "exportacao";
    displayName = "Exportação";
    mostrarBotao = new formattingSettings.ToggleSwitch({ name: "mostrarBotao", displayName: "Mostrar botão Exportar", value: true });
    nomeArquivo = new formattingSettings.TextInput({ name: "nomeArquivo", displayName: "Nome do arquivo Excel", value: "oportunidade_mix", placeholder: "nome do arquivo" });
    urlPagina = new formattingSettings.TextInput({ name: "urlPagina", displayName: "Endereço da página de download", value: URL_PAGINA, placeholder: "https://.../download-page/" });
    slices = [this.mostrarBotao, this.nomeArquivo, this.urlPagina];
}

export class VisualSettings extends Model {
    titulo = new TituloCard();
    cabecalho = new CabecalhoCard();
    fundo = new FundoCard();
    periodoCfg = new PeriodoCard();
    exportacao = new ExportacaoCard();
    cards = [this.titulo, this.cabecalho, this.fundo, this.periodoCfg, this.exportacao];
}
