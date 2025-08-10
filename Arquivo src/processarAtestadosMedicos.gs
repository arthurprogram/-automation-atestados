/**
 * Automação de Atestados Médicos com Google Apps Script
 * Desenvolvido por Arthur Santos
 * LinkedIn: https://www.linkedin.com/in/arthur-santos
 */

function processarAtestadosMedicos() {
  Logger.clear();
  Logger.log("▶️ Início do processamento");

  // ID da pasta principal no Google Drive - SUBSTITUA PELO SEU
  const pastaAtestadosId = 'PASTE_AQUI_O_ID_DA_PASTA';

  const planilha = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const mesesPT = ["JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO",
                   "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"];

  const dataAtual = new Date();
  const numeroMes = ("0" + (dataAtual.getMonth() + 1)).slice(-2);
  const nomeMes = mesesPT[dataAtual.getMonth()];
  const anoFixo = dataAtual.getFullYear();

  const nomeDaPastaDoAno = `${anoFixo}`;
  const nomeDaPastaDoMes = `${numeroMes} - ${nomeMes} ${anoFixo}`;

  Logger.log(`Procurando pasta do ano: "${nomeDaPastaDoAno}"`);
  Logger.log(`Procurando subpasta do mês: "${nomeDaPastaDoMes}"`);

  let pastaPrincipal;
  try {
    pastaPrincipal = DriveApp.getFolderById(pastaAtestadosId);
  } catch (e) {
    Logger.log(`❌ Erro ao acessar pasta principal: ${e.message}`);
    return;
  }

  let pastaDoAno = null;
  const pastasAno = pastaPrincipal.getFolders();
  while (pastasAno.hasNext()) {
    const pasta = pastasAno.next();
    Logger.log(`Encontrada pasta (nível ano): "${pasta.getName()}"`);
    if (pasta.getName() === nomeDaPastaDoAno) {
      pastaDoAno = pasta;
      Logger.log(`✅ Pasta do ano encontrada`);
      break;
    }
  }

  if (!pastaDoAno) {
    Logger.log(`ℹ️ Pasta do ano "${nomeDaPastaDoAno}" não encontrada. Criando...`);
    pastaDoAno = pastaPrincipal.createFolder(nomeDaPastaDoAno);
    Logger.log(`✅ Pasta do ano criada`);
  }

  let subpasta = null;
  const pastasMes = pastaDoAno.getFolders();
  while (pastasMes.hasNext()) {
    const pasta = pastasMes.next();
    Logger.log(`Encontrada pasta (nível mês): "${pasta.getName()}"`);
    if (pasta.getName() === nomeDaPastaDoMes) {
      subpasta = pasta;
      Logger.log(`✅ Subpasta encontrada`);
      break;
    }
  }

  if (!subpasta) {
    Logger.log(`ℹ️ Subpasta "${nomeDaPastaDoMes}" não encontrada. Criando...`);
    subpasta = pastaDoAno.createFolder(nomeDaPastaDoMes);
    Logger.log(`✅ Subpasta do mês criada`);
  }

  const palavrasChave = [
    "atestado médico", "declaração", "declarações", "declaracão",
    "declaraçao", "atestados", "declaracao de horas", "atestado"
  ];

  const etiquetaProcessado = GmailApp.getUserLabelByName('Processado') || GmailApp.createLabel('Processado');
  const query = `has:attachment is:unread newer_than:7d -label:Processado`;

  let threads;
  try {
    threads = GmailApp.search(query);
  } catch (e) {
    Logger.log(`❌ Erro ao buscar emails: ${e.message}`);
    return;
  }

  Logger.log(`Threads encontradas: ${threads.length}`);

  if (threads.length === 0) {
    Logger.log("ℹ️ Nenhum email recente com anexo.");
    return;
  }

  for (const t of threads) {
    const mensagens = t.getMessages();
    for (const m of mensagens) {
      if (!m.isUnread()) {
        Logger.log(`Mensagem já lida: "${m.getSubject()}"`);
        continue;
      }

      const etiquetasDaThread = t.getLabels();
      const jaProcessadoNaThread = etiquetasDaThread.some(et => et.getName() === 'Processado');
      if (jaProcessadoNaThread) {
        Logger.log(`Thread já processada`);
        continue;
      }

      const assuntoOriginal = m.getSubject();
      const assuntoLower = assuntoOriginal.toLowerCase();
      const corpo = m.getPlainBody().toLowerCase();
      const contemPalavraChave = palavrasChave.some(palavra => assuntoLower.includes(palavra) || corpo.includes(palavra));
      const anexos = m.getAttachments();

      if (!contemPalavraChave || anexos.length === 0) {
        Logger.log(`Ignorado: sem palavras-chave ou anexos`);
        continue;
      }

      const links = [];
      for (const a of anexos) {
        const dataHoje = new Date();
        const dia = dataHoje.getDate();
        const dataFormatada = `${dia}_${("0" + (dataHoje.getMonth() + 1)).slice(-2)}`;
        const assuntoLowerCase = assuntoOriginal.toLowerCase();
        let prefixo = "Atest";

        if (assuntoLowerCase.startsWith("declaração") || assuntoLowerCase.startsWith("declaracao")) {
          prefixo = "Declar";
        }

        const remetente = m.getFrom();
        const nomeRemetente = extrairNome(remetente);
        const nomeArquivo = `${dataFormatada}_${prefixo}.${nomeRemetente}`;
        const blob = a.copyBlob();
        const arquivo = subpasta.createFile(blob).setName(nomeArquivo);
        links.push(arquivo.getUrl());
      }

      Logger.log("Inserindo linha na planilha...");
      planilha.appendRow([
        Utilities.formatDate(m.getDate(), Session.getScriptTimeZone(), "dd/MM/yyyy"),
        m.getFrom(),
        m.getSubject(),
        'Atestado / Declaração',
        links.join(', '),
        'Pendente'
      ]);

      const remetente = m.getFrom();
      const remetenteEmail = extrairEmail(remetente);
      const nomeRemetente = extrairPrimeiroNome(remetente);

      const assuntoAtualizado = resumoAssunto(m.getSubject());
      const mensagemResposta = `Bom dia, ${nomeRemetente}!\nObrigado pelo envio das informações. Seu documento foi recebido e processado.`;

      const ccEmails = m.getCc();
      const ccList = ccEmails ? ccEmails.split(',').map(email => email.trim()) : [];

      if (remetenteEmail) {
        try {
          m.reply(
            mensagemResposta,
            {
              htmlBody: mensagemResposta.replace(/\n/g, '<br>'),
              subject: assuntoAtualizado,
              cc: ccList.join(',')
            }
          );
          Logger.log(`Resposta enviada`);
        } catch (e) {
          Logger.log(`Erro ao responder: ${e.message}`);
        }
      } else {
        Logger.log(`Não foi possível extrair email do remetente`);
      }

      m.markRead();
    }

    t.addLabel(etiquetaProcessado);
    Logger.log(`Thread marcada como processada`);
  }

  Logger.log("✅ Processamento concluído");
}

function extrairPrimeiroNome(remetente) {
  const nomeCompleto = remetente.replace(/<.*?>/g, '').trim();
  return nomeCompleto.split(' ')[0];
}

function extrairNome(remetente) {
  return remetente.replace(/<.*?>/g, '').trim().replace(/\s+/g, ' ');
}

function extrairEmail(remetente) {
  const match = remetente.match(/<(.+)>/);
  return match ? match[1] : remetente;
}

function resumoAssunto(assunto) {
  return assunto.toLowerCase().startsWith("re:") ? assunto : `Re: ${assunto}`;
}
