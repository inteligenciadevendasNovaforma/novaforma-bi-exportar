import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import Card = formattingSettings.SimpleCard;
import Model = formattingSettings.Model;

class TituloCard extends Card {
    name = "titulo";
    displayName = "Título";
    mostrar = new formattingSettings.ToggleSwitch({ name: "mostrar", displayName: "Mostrar", value: true });
    texto = new formattingSettings.TextInput({ name: "texto", displayName: "Texto", value: "PAINEL DE DEVOLUÇÕES", placeholder: "Título" });
    tamanho = new formattingSettings.NumUpDown({ name: "tamanho", displayName: "Tamanho", value: 14 });
    cor = new formattingSettings.ColorPicker({ name: "cor", displayName: "Cor", value: { value: "#FFFFFF" } });
    slices = [this.mostrar, this.texto, this.tamanho, this.cor];
}

class FundoCard extends Card {
    name = "fundo";
    displayName = "Fundo do visual";
    mostrar = new formattingSettings.ToggleSwitch({ name: "mostrar", displayName: "Mostrar", value: true });
    cor = new formattingSettings.ColorPicker({ name: "cor", displayName: "Cor", value: { value: "#FFFFFF" } });
    slices = [this.mostrar, this.cor];
}

class BotoesCard extends Card {
    name = "exportacao";
    displayName = "Exportação";
    mostrarBotao = new formattingSettings.ToggleSwitch({ name: "mostrarBotao", displayName: "Mostrar botão Exportar", value: true });
    nomeArquivo = new formattingSettings.TextInput({ name: "nomeArquivo", displayName: "Nome do arquivo Excel", value: "painel_devolucoes", placeholder: "nome do arquivo" });
    incluirCards = new formattingSettings.ToggleSwitch({ name: "incluirCards", displayName: "Padrão: incluir cards", value: true });
    incluirGraficos = new formattingSettings.ToggleSwitch({ name: "incluirGraficos", displayName: "Padrão: incluir gráficos", value: true });
    incluirMatriz = new formattingSettings.ToggleSwitch({ name: "incluirMatriz", displayName: "Padrão: incluir matriz", value: true });
    incluirDados = new formattingSettings.ToggleSwitch({ name: "incluirDados", displayName: "Padrão: aba de dados brutos", value: false });
    usarPagina = new formattingSettings.ToggleSwitch({ name: "usarPagina", displayName: "Baixar pela página de download (link)", value: true });
    urlPagina = new formattingSettings.TextInput({ name: "urlPagina", displayName: "Endereço da página de download", value: "https://gabrielpms2001.github.io/novaforma-exportar-bi/download-page/", placeholder: "https://.../download-page/" });
    slices = [this.mostrarBotao, this.nomeArquivo, this.usarPagina, this.urlPagina, this.incluirCards, this.incluirGraficos, this.incluirMatriz, this.incluirDados];
}

export class VisualSettings extends Model {
    titulo = new TituloCard();
    fundo = new FundoCard();
    exportacao = new BotoesCard();
    cards = [this.titulo, this.fundo, this.exportacao];
}
