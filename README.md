# novaforma-bi-exportar

Exportação para Excel dos visuais personalizados do Power BI, mesmo com o relatório embutido (onde a exportação nativa de visual personalizado falha).

## Como funciona
1. O visual monta os dados e os gráficos em texto pequeno (JSON comprimido com deflate-raw e codificado em base64url) e coloca tudo no **fragmento** (`#...`) de um link. O link fica com poucos milhares de caracteres: o Power BI recusa links muito longos ("link malformado").
2. O visual chama o link com `host.launchUrl`. O Power BI pergunta se pode abrir o navegador.
3. A página `download-page/` lê o fragmento (o fragmento não é enviado ao servidor), descomprime os dados, redesenha os gráficos em imagem (SVG para canvas para PNG), monta o `.xlsx` no próprio navegador, sem bibliotecas, e baixa o arquivo.
4. Se a página não for usada, o visual tenta o download direto pelo `downloadService` do Power BI, que depende da configuração do tenant.

A página é única e atende todos os visuais. O tipo de arquivo vem no campo `tipo` da carga. Cargas antigas continuam aceitas.

## Publicar a página
Settings > Pages > Branch `main`, pasta `/ (root)`. O endereço será:
`https://<conta>.github.io/novaforma-bi-exportar/download-page/`

O endereço é configurado em cada visual, em Formatar > Exportação > "Endereço da página de download". Ele também precisa estar liberado nos privilégios de acesso web do visual (`https://*.github.io`).

## Pastas
- `download-page/` página publicada (um único `index.html`, já com o código embutido)
- `pagina/` código-fonte da página (`ponte.ts`, `template.html`, fonte). Para gerar de novo: `pagina/build.sh`
- `dist/` visuais empacotados (`.pbiviz`)
- `painel-dia/`, `painel-mes/`, `pedido-diario/`, `reps-sem-pedidos/`, `visual/` código-fonte de cada visual (`npm i && npx pbiviz package`) e a medida DAX
