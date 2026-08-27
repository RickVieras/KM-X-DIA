# KM X DIA Online

Painel Flask para processar a aba `PROGRAMADO` e gerar o resumo diário de KM.

## Atualização da base

No painel, clique em **Enviar nova planilha**, informe a senha e escolha o novo arquivo `.xlsx` ou `.xlsm`. Não é necessário alterar o HTML nem fazer um novo deploy.

O arquivo deve possuir a aba `PROGRAMADO` e manter a estrutura de colunas utilizada pelo relatório.

## Variáveis no Render

- `ADMIN_TOKEN`: senha usada para autorizar a atualização.
- `GOOGLE_SHEET_ID`: opcional. Mantém disponível o botão de atualização pelo Google Sheets.

## Render

- Root Directory: `km_x_dia_online`
- Build Command: `pip install -r requirements.txt`
- Start Command: `gunicorn app:app --timeout 180`

Observação: no plano sem disco persistente, os dados processados podem ser perdidos quando o serviço reiniciar. Nesse caso, basta reenviar a planilha. Para persistência total, adicione um Persistent Disk no Render e monte-o na pasta `data`.
