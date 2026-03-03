/*
 * ORQUESTRADOR_ALERTAS_OPERACOESPIPELINE.GS
 * ----------------------------------------
 * Pipeline de envio do relatório diário de operações em aberto.
 *

/**
 * Lê Config_Global para definir e-mail destino do relatório.
 * Usa "Email_Relatorio_Operacoes" ou, se vazio, "Email_Alerta_Padrao".
 */
function getConfigRelatorioOperacoes_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaConfig = ss.getSheetByName("Config_Global");

  if (!abaConfig) {
    log("PipelineOperacoes", "ERRO_CRITICO", "Aba Config_Global não encontrada", "");
    return { email: "" };
  }

  const ultima = abaConfig.getLastRow();
  if (ultima < 2) {
    return { email: "" };
  }

  const dados = abaConfig.getRange(2, 1, ultima - 1, 2).getValues();
  const map = {};
  dados.forEach(function (l) {
    if (l[0]) map[String(l[0]).trim()] = l[1];
  });

  var emailOp = String(map["Email_Relatorio_Operacoes"] || "").trim();
  var emailPadrao = String(map["Email_Alerta_Padrao"] || "").trim();

  var email = emailOp || emailPadrao;

  return { email: email };
}

/**
 * Função principal do pipeline das operações abertas.
 *
 * Ajuste solicitado:
 *   → REMOVIDA a sequência de pré-sync
 *       (SyncDadosAtivos, SyncDadosDetalhes, CalcGreeks, sleeps)
 *
 * Agora:
 *   → Apenas chama o envio do relatório usando dados já existentes.
 */
function rodarPipelineOperacoesAbertas() {
  const SERVICO = "PipelineOperacoes";

  log(
    SERVICO,
    "INFO",
    "Iniciando pipeline diário de operações abertas (somente envio do relatório)",
    ""
  );

  try {
    // Apenas envia o relatório
    enviarRelatorioOperacoesAbertas_();

    log(SERVICO, "SUCESSO", "Pipeline de operações abertas concluído (sem syncs)", "");

  } catch (e) {
    log(SERVICO, "ERRO_CRITICO", "Falha no pipeline de operações abertas", e.message);
    throw e;
  }
}

/**
 * Monta o DailyReport e envia o e-mail.
 */
function enviarRelatorioOperacoesAbertas_() {
  const SERVICO = "RelatorioOperacoes";

  const cfg = getConfigRelatorioOperacoes_();
  if (!cfg.email) {
    log(SERVICO, "ERRO", "Nenhum e-mail configurado para relatório de operações", "");
    return;
  }

  // 1) Carrega pernas abertas
  const pernas = carregarPernasAbertasDaCockpit();
  if (!pernas || pernas.length === 0) {
    log(SERVICO, "AVISO", "Nenhuma perna em aberto encontrada. Relatório não enviado.", "");
    return;
  }

  // 2) Agrupa em estruturas
  const estruturas = agruparPernasEmEstruturas(pernas);
  if (!estruturas || estruturas.length === 0) {
    log(SERVICO, "AVISO", "Nenhuma estrutura válida montada. Relatório não enviado.", "");
    return;
  }

  // 3) DailyReport
  const report = montarDailyReportOperacoes(estruturas);

  log(
    SERVICO,
    "DEBUG",
    "DailyReport de operações montado",
    JSON.stringify({
      totalEstruturas: report.resumo.totalEstruturas,
      lucroTotal: report.resumo.lucroTotal,
      deltaTotal: report.resumo.deltaTotal
    })
  );

  // 4) HTML
  const html = montarHtmlRelatorioOperacoesAbertas(report);

  // 5) Assunto
  const hoje = new Date();
  const dataStr = Utilities.formatDate(
    hoje,
    Session.getScriptTimeZone(),
    "dd/MM/yyyy"
  );

  const assunto = "Diário – Operações em Opções (" + dataStr + ")";

  // 6) Envio
  MailApp.sendEmail({
    to: cfg.email,
    subject: assunto,
    htmlBody: html
  });

  log(
    SERVICO,
    "SUCESSO_ENVIO",
    "Relatório diário de operações abertas enviado",
    JSON.stringify({
      destino: cfg.email,
      estruturas: report.resumo.totalEstruturas,
      timestamp: new Date().toISOString()
    })
  );
}

/**
 * Função de teste: monta e envia o relatório sem pipeline.
 */
function testarRelatorioOperacoesAbertas() {
  enviarRelatorioOperacoesAbertas_();
}
