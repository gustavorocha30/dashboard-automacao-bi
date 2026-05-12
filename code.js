/**
 * AUTOMAÇÃO DE NOTIFICAÇÕES - DASHBOARD
 */

function enviarNotificacoesSemanais() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const aba = planilha.getSheetByName("Dados"); 
  
  if (!aba) {
    Logger.log("ERRO: Aba 'Dados' não encontrada!");
    return;
  }
  
  const ultimaLinha = aba.getLastRow();
  const primeiraLinhaDados = 10; 
  
  if (ultimaLinha < primeiraLinhaDados) {
    Logger.log("Nenhum dado encontrado abaixo da linha 10");
    return;
  }
  
  const dados = aba.getRange(primeiraLinhaDados, 1, ultimaLinha - primeiraLinhaDados + 1, 14).getValues();
  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0);
  
  const porResponsavel = {};
  
  dados.forEach((linha, index) => {
    const id = linha[0];          // Coluna A
    const descricao = linha[5];   // Coluna F
    const responsavel = linha[6]; // Coluna G
    const prazo = linha[7];       // Coluna H
    const status = linha[9];      // Coluna J
    const prioridade = linha[10]; // Coluna K
    const email = linha[13];      // Coluna N
    
    if (!email || email === "") return; 
    
    let dataPrazo;
    if (prazo instanceof Date) {
      dataPrazo = new Date(prazo);
    } else {
      dataPrazo = new Date(prazo);
      if (isNaN(dataPrazo.getTime())) return;
    }
    dataPrazo.setHours(0, 0, 0, 0);
    
    const diffTime = dataPrazo - hoje;
    const diasRestantes = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    
    let categoria = "vencendo"; 
    if (status === "Concluída") categoria = "concluidas";
    if (diasRestantes < 0 && status !== "Concluída") categoria = "atrasadas";

    if (!porResponsavel[email]) {
      porResponsavel[email] = {
        nome: responsavel || "Responsável",
        atrasadas: [], vencendo: [], concluidas: []
      };
    }
    
    porResponsavel[email][categoria].push({
      id: id || "S/ID",
      descricao: descricao || "Sem descrição",
      prazo: dataPrazo,
      diasRestantes: diasRestantes,
      prioridade: prioridade,
      status: status
    });
  });
  
  const emails = Object.keys(porResponsavel);
  emails.forEach(email => {
    Logger.log(`Tentando enviar e-mail para: ${email}`);
    // CHAMADA DA FUNÇÃO DE ENVIO
    processarEnvioEmail(email, porResponsavel[email], true);
  });
  
  Logger.log(`Processo finalizado. Responsáveis encontrados: ${emails.length}`);
}

// FUNÇÃO QUE ESTAVA FALTANDO (Renomeada para garantir que o script ache)
function processarEnvioEmail(destinatario, info, relatorioCompleto) {
  const assunto = info.atrasadas.length > 0 
    ? `⚠️ ALERTA: Recomendações Atrasadas`
    : `📋 Resumo de Recomendações - Dashboard`;
  
  let corpo = `<html><body style="font-family: Arial, sans-serif;">
    <h2 style="color: #1F3864;">Olá, ${info.nome}</h2>
    <p>Seguem as informações das suas recomendações no Dashboard:</p>`;
  
  if (info.atrasadas.length > 0) {
    corpo += `<h3 style="color: red;">🚨 ATRASADAS</h3>`;
    info.atrasadas.forEach(item => {
      corpo += `<p><strong>${item.id}</strong>: ${item.descricao}<br>Prazo: ${item.prazo.toLocaleDateString('pt-BR')}</p>`;
    });
  }

  if (info.vencendo.length > 0) {
    corpo += `<h3 style="color: orange;">⏰ EM ANDAMENTO / VENCENDO</h3>`;
    info.vencendo.forEach(item => {
      corpo += `<p><strong>${item.id}</strong>: ${item.descricao}<br>Prazo: ${item.prazo.toLocaleDateString('pt-BR')}</p>`;
    });
  }

  corpo += `<br><p>Acesse a planilha para atualizar: <a href="${SpreadsheetApp.getActiveSpreadsheet().getUrl()}">Clique aqui</a></p>
    </body></html>`;

  try {
    MailApp.sendEmail({
      to: destinatario,
      subject: assunto,
      htmlBody: corpo
    });
    Logger.log(`E-mail enviado com sucesso para ${destinatario}`);
  } catch (e) {
    Logger.log(`Erro crítico no envio: ${e}`);
  }
}

function testarEnvio() {
  Logger.log("=== INICIANDO TESTE ===");
  enviarNotificacoesSemanais();
  Logger.log("=== TESTE CONCLUÍDO ===");
}
