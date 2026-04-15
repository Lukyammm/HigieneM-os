# HigieneM-os

Web App em Google Apps Script para análise de adesão à higiene das mãos a partir de uma planilha Google Sheets.

## Configuração rápida

1. Abra o projeto no Apps Script.
2. Garanta que o arquivo HTML se chama **`index`** no projeto.
3. Configure a origem de dados por **uma** destas opções:
   - **Recomendado:** em *Project Settings > Script properties*, crie a chave `SPREADSHEET_ID` com o ID da planilha Google Sheets.
   - **Alternativa:** preencha `CONFIG.SPREADSHEET_ID` no `Code.gs`.
4. Implante como Web App (Deploy > New deployment > Web app), atualizando para a versão mais recente.

## Diagnóstico de conexão com base

No editor do Apps Script, execute:

- `testarAcesso()` para validar vínculo do projeto com a planilha e abas localizadas.
- `testarBase()` para validar leitura da base e amostra das primeiras linhas.
