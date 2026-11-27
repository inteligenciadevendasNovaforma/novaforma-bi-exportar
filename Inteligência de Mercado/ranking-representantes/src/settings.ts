"use strict";
import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import Card = formattingSettings.SimpleCard;
import Model = formattingSettings.Model;

class TituloCard extends Card {
    name = "titulo";
    displayName = "Título";
    mostrar = new formattingSettings.ToggleSwitch({ name: "mostrar", displayName: "Mostrar título", value: true });
    mostrarEyebrow = new formattingSettings.ToggleSwitch({ name: "mostrarEyebrow", displayName: "Mostrar linha de filtros", value: true });
    fontFamily = new formattingSettings.FontPicker({ name: "fontFamily", value: "Segoe UI, sans-serif" });
    fontSize = new formattingSettings.NumUpDown({ name: "fontSize", displayName: "Tamanho", value: 22 });
    bold = new formattingSettings.ToggleSwitch({ name: "bold", displayName: "Negrito", value: true });
    italic = new formattingSettings.ToggleSwitch({ name: "italic", displayName: "Itálico", value: false });
    cor = new formattingSettings.ColorPicker({ name: "cor", displayName: "Cor", value: { value: "#FFFFFF" } });
    slices = [this.mostrar, this.mostrarEyebrow, this.fontFamily, this.fontSize, this.bold, this.italic, this.cor];
}

class CabecalhoCard extends Card {
    name = "cabecalho";
    displayName = "Cabeçalho";
    mostrarFundo = new formattingSettings.ToggleSwitch({ name: "mostrarFundo", displayName: "Mostrar fundo", value: true });
    cor = new formattingSettings.ColorPicker({ name: "cor", displayName: "Cor do fundo", value: { value: "#0B2A4A" } });
    mostrarTotais = new formattingSettings.ToggleSwitch({ name: "mostrarTotais", displayName: "Mostrar totais", value: true });
    slices = [this.mostrarFundo, this.cor, this.mostrarTotais];
}

class FundoCard extends Card {
    name = "fundo";
    displayName = "Fundo do visual";
    mostrar = new formattingSettings.ToggleSwitch({ name: "mostrar", displayName: "Mostrar fundo", value: true });
    cor = new formattingSettings.ColorPicker({ name: "cor", displayName: "Cor", value: { value: "#F4F6F9" } });
    slices = [this.mostrar, this.cor];
}

class ExportacaoCard extends Card {
    name = "exportacao";
    displayName = "Exportação";
    urlPagina = new formattingSettings.TextInput({ name: "urlPagina", displayName: "Endereço da página de download", value: "https://inteligenciadevendasnovaforma.github.io/novaforma-bi-exportar/download-page/", placeholder: "https://.../download-page/" });
    slices = [this.urlPagina];
}

export class VisualSettings extends Model {
    titulo = new TituloCard();
    cabecalho = new CabecalhoCard();
    fundo = new FundoCard();
    exportacao = new ExportacaoCard();
    cards = [this.titulo, this.cabecalho, this.fundo, this.exportacao];
}
