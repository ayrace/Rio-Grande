# Painel Geográfico de Nodes – Rio Grande

Clone operacional do painel funcional de Porto Alegre, adaptado somente para a base geográfica/nodes de Rio Grande e coleta XPERTrack no Google Drive.

## Publicação
Main file path: `streamlit_app.py`

## Coleta
O app lê o arquivo fixo `RIO GRANDE.csv` do Google Drive. Para leitura pelo Streamlit Cloud, mantenha o arquivo acessível como **Qualquer pessoa com o link — Leitor** e preserve o mesmo arquivo/ID nas atualizações.

## Regras XPERTrack
- 0 = OFF
- 1–20 = porta crítica/degradada
- acima de 20 = online

A interface, cards, visão de crise, filtros, ranking, mapa, popups, tratativas, outages, histórico e demais recursos seguem a matriz de Porto Alegre.


## Correções validadas de topologia
- TRVABA = 1x4; `TRVABA3` é a porta 3 de TRVABA.
- TRVACD = 1x2 e é independente.
- TRVACC = 1x4 e é independente.
- TRVACD e TRVACC nunca são agrupados.


## Ajuste final
- CTAA removido da base geográfica por duplicidade confirmada.
- CNTAAA preservado no ponto correto, consolidando CNTAAA-1/2/3/4.
