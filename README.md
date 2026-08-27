# KM X DIA Online

Painel Flask para processar a aba `PROGRAMADO` e gerar o resumo diário de KM.

## Atualização da base

No painel, escolha a data inicial e a data final, clique em **Enviar nova planilha**, informe a senha e escolha o novo arquivo `.xlsx` ou `.xlsm`. As datas alimentam os parâmetros BM1 e BN1 e delimitam diretamente o cálculo.

O arquivo deve possuir a aba `PROGRAMADO` e manter a estrutura de colunas utilizada pelo relatório.

## Histórico mensal no Supabase

Cada upload é identificado pelas datas da aba `PROGRAMADO`. O relatório processado é salvo no bucket privado `relatorios-km` e os indicadores são gravados na tabela `relatorios_mensais`. Se o mês já existir, ele é atualizado; os demais meses são preservados.

Na planilha gerada, não há coluna diária de KM Transporta. O valor aparece somente no fechamento ao final de cada aba de empresa, seguido pelo total mensal incluindo Transporta.

## Variáveis no Render

- `ADMIN_TOKEN`: senha usada para autorizar a atualização.
- `GOOGLE_SHEET_ID`: opcional. Mantém disponível o botão de atualização pelo Google Sheets.
- `SUPABASE_URL`: URL do projeto Supabase.
- `SUPABASE_SECRET_KEY`: chave secreta moderna iniciada por `sb_secret_`. Nunca coloque essa chave no GitHub ou no HTML.

## Render

- Root Directory: `km_x_dia_online`
- Build Command: `pip install -r requirements.txt`
- Start Command: `gunicorn app:app --timeout 180`

O histórico mensal fica persistido no Supabase e não é perdido quando o Render reinicia.
