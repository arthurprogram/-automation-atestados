# -automation-atestados
Script em Google Apps Script para automatizar o processamento de atestados médicos recebidos por e-mail no Gmail. Faz download automático dos anexos, organiza-os em pastas no Google Drive por ano e mês, registra em planilha e envia resposta automática ao remetente.
# Automação de Atestados Médicos com Google Apps Script 📄🤖

Este script automatiza o recebimento, organização e registro de atestados médicos enviados por e-mail no Gmail. Ele integra **Gmail**, **Google Drive** e **Google Sheets** para otimizar o fluxo de trabalho.

## 📌 Funcionalidades
- Pesquisa no Gmail por e-mails não lidos com anexos de atestados ou declarações.
- Criação automática de pastas no Google Drive organizadas por ano e mês.
- Salvamento de anexos na pasta correta.
- Registro dos dados do atestado em uma planilha do Google Sheets.
- Resposta automática ao remetente confirmando o recebimento.

## 🛠️ Tecnologias Utilizadas
- **Google Apps Script** (JavaScript)
- **GmailApp API**
- **DriveApp API**
- **SpreadsheetApp API**

## 🚀 Como Usar
1. Abra o [Google Apps Script](https://script.google.com/).
2. Crie um novo projeto e cole o conteúdo do arquivo `processarAtestadosMedicos.gs` na aba do editor.
3. Substitua o valor de:
   ```javascript
   const pastaAtestadosId = 'PASTE_AQUI_O_ID_DA_PASTA';
