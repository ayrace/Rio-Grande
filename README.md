# Mapa HFC — Rio Grande

Versão preparada para atualização da coleta XPERTrack por fonte compartilhada no Google Drive.

## Fluxo
1. A base geográfica (node, latitude, longitude e bairro) permanece fixa no projeto.
2. A equipe atualiza o CSV compartilhado com as colunas `Node`, `Pontuação`, `Impactado`, `Estressado`, `Total`.
3. Ao abrir/recarregar o mapa, o `index.html` busca a coleta externa.
4. Se a fonte estiver indisponível, o navegador usa a última coleta válida salva localmente.

## Configuração do Drive
No `index.html`, preencher `CONFIG.DRIVE_CSV_URL` com a URL HTTPS da fonte CSV/endpoint. O link será configurado quando a pasta compartilhada for fornecida.

> Observação: um link comum de pasta do Google Drive não entrega o CSV diretamente ao navegador. A integração final pode usar um arquivo publicado/endpoint Apps Script para transformar a pasta compartilhada na fonte automática do mapa.


## Atualização pelo Google Drive
O mapa está configurado para ler o arquivo `RIO GRANDE.csv` (ID fixo do Drive). Para manter o mesmo ID, substitua/atualize o conteúdo do arquivo existente, em vez de apagar e criar outro arquivo. Se um novo arquivo for criado, o ID muda e o `index.html` precisará ser atualizado. O arquivo precisa estar acessível para o navegador que abrir o GitHub Pages.
