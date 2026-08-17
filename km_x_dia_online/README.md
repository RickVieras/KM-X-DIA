# Painel KM x Dia

Aplicação online para recalcular o relatório **KM X DIA** a partir da aba `PROGRAMADO` de uma planilha Google.

## Como funciona

1. O botão **Atualizar dados** baixa a planilha de origem e aplica as regras do gerador atual: U/S/D, KM operacional, KM morta e exceção Transporta.
2. O painel lê o resumo processado.
3. **Baixar planilha** gera o Excel com abas por empresa e a aba `TOTAL POR EMPRESA`.

## Configuração para publicar no Render

Defina as variáveis de ambiente:

- `GOOGLE_SHEET_ID`: `1WSUarpfz868ogUWAVQTY6bSzdn4f6K3R5UwNxFiSkTc`
- `ADMIN_TOKEN`: uma senha escolhida por você para habilitar a atualização.

A planilha precisa estar acessível como **qualquer pessoa com o link — visualizador** para que o servidor consiga baixá-la.

Comando de inicialização:

```text
gunicorn app:app
```
