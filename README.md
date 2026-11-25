# novaforma-bi-exportar

Exportação para Excel dos visuais personalizados do Power BI, mesmo com o relatório embutido (onde a exportação nativa de visual personalizado falha).

## Como funciona
1. O visual monta os dados e os gráficos em texto pequeno (JSON comprimido com deflate-raw e codificado em base64url) e coloca tudo no **fragmento** (`#...`) de um link. O link fica com poucos milhares de caracteres: o Power BI recusa links muito longos ("link malformado").
2. O visual chama o link com `host.launchUrl`. O Power BI pergunta se pode abrir o navegador.
3. A página `download-page/` lê o fragmento (o fragmento não é enviado ao servidor), descomprime os dados, redesenha os gráficos em imagem (SVG para canvas para PNG), monta o `.xlsx` no próprio navegador, sem bibliotecas, e baixa o arquivo.
4. Se a página não for usada, o visual tenta o download direto pelo `downloadService` do Power BI, que depende da configuração do tenant.

### Tabelas grandes (Pedidos Mês)
O Pedidos Mês tem milhares de linhas e não cabe no link. Ordem das tentativas ao clicar em **Baixar Excel**:
1. `downloadService` liberado (Desktop ou tenant com exportação permitida): o visual gera o `.xlsx` completo e baixa direto.
2. Sem permissão: monta o link com os dados compactados (famílias numa tabela, hierarquia representante > cliente > [família, centavos], deflate-raw + base64url). Se o link tiver até 2.500 caracteres (cerca de 100 linhas), abre a página com os dados no fragmento.
3. Link maior: o visual copia os dados compactados para a área de transferência (`NFPM1.` + base64url) e abre a página com um link curto (`#k=cola`). A página pede **Colar dados** (ou Ctrl+V), monta o Excel e baixa. Nada passa por servidor.
4. Cópia bloqueada: mensagem pedindo para filtrar (Supervisão/Representante) e exportar de novo.

O botão **Copiar tabela** copia a tabela em texto com tabulações, no layout da aba, para colar direto no Excel.

A página é única e atende todos os visuais. O tipo de arquivo vem no campo `tipo` da carga. Cargas antigas continuam aceitas.

## Publicar a página
Settings > Pages > Branch `main`, pasta `/ (root)`. O endereço será:
`https://<conta>.github.io/novaforma-bi-exportar/download-page/`

O endereço é configurado em cada visual, em Formatar > Exportação > "Endereço da página de download". Ele também precisa estar liberado nos privilégios de acesso web do visual (`https://*.github.io`).

## Pastas
- `download-page/` página publicada (um único `index.html`, já com o código embutido)
- `pagina/` código-fonte da página (`ponte.ts`, `template.html`, fonte). Para gerar de novo: `pagina/build.sh`
- `dist/` visuais empacotados (`.pbiviz`)
- `painel-dia/`, `painel-mes/`, `painel-metas/`, `pedido-diario/`, `pedidos-mes/`, `reps-sem-pedidos/`, `visual/` código-fonte de cada visual (`npm i && npx pbiviz package`) e a medida DAX
