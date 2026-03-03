/*
 * ORQUESTRADOR_ALERTAS.GS
 * Neste arquivo ficam:
 *  - Logs
 *  - Configuração de alertas (e-mail, DTE, tolerância)
 *  - Filtro com tolerância
 *  - Envio de e-mail
 *  - Pipeline final rodarPipelineStranglesComAlerta()
 */

/************************************************************
 * LOG COMPLETO — CONSOLE + Aba Logs
 ************************************************************/
function log(servico, tipo, mensagem, detalhe) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("Logs");

  const ts = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "dd/MM/yyyy HH:mm:ss"
  );

  const det = detalhe || "";

  console.log(`${ts}\t${servico}\t${tipo}\t${mensagem}\t${det}`);

  if (aba) aba.appendRow([ts, servico, tipo, mensagem, det]);
}

/************************************************************
 * LEITURA DAS CONFIGURAÇÕES DE ALERTA
 ************************************************************/
function getConfigAlertasStrangles(abaConfig) {
  const ultima = abaConfig.getLastRow();
  if (ultima < 2) {
    return { emailAlerta:"", diasAntes:45, tolerancia:0 };
  }

  const dados = abaConfig.getRange(2, 1, ultima - 1, 2).getValues();
  const map = {};
  dados.forEach(l => { if (l[0]) map[l[0]] = l[1]; });

  return {
    emailAlerta: String(map["Email_Alerta_Padrao"] || "").trim(),
    diasAntes: Number(map["Strangle_Alert_Dias_Antes"] || 45),
    tolerancia: Number(map["Strangle_Alert_DTE_Tolerance"] || 0)
  };
}

/************************************************************
 * ENVIO DE ALERTAS (COM TOLERÂNCIA DE DTE)
 ************************************************************/
function enviarAlertasStrangles(modoTeste) {
  const SERVICO = "AlertasStrangles";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaStr = ss.getSheetByName(ABA_SELECAO_STRANGLES);
  const abaCfg = ss.getSheetByName("Config_Global");

  if (!abaStr || !abaCfg) {
    log(SERVICO, "ERRO_CRITICO", "Abas necessárias não encontradas", "");
    return;
  }

  const cfg = getConfigAlertasStrangles(abaCfg);

  if (!cfg.emailAlerta) {
    log(SERVICO, "ERRO", "Email_Alerta_Padrao não configurado", "");
    return;
  }

  const ultima = abaStr.getLastRow();
  if (ultima <= 1) {
    log(SERVICO, "AVISO", "Selecao_Strangles está vazia", "");
    return;
  }

  const valores = abaStr.getRange(2, 1, ultima - 1, HEADERS_STRANGLES.length).getValues();

  const IDX = {
    ticker:0, due:1, spot:2, dte:3,
    ps:4, pst:5, pcl:6, pd:7,
    cs:8, cst:9, ccl:10, cd:11, tot:12
  };

  const distrib = {};
  valores.forEach(r => {
    var d = Number(r[IDX.dte]) || 0;
    distrib[d] = (distrib[d] || 0) + 1;
  });

  log(SERVICO, "DEBUG_DTE_DISTRIB", "Distribuição de DTE", JSON.stringify(distrib));

  const alvo = Number(cfg.diasAntes);
  const tol  = Number(cfg.tolerancia);
  const faixaMin = alvo - tol;
  const faixaMax = alvo + tol;

  log(SERVICO, "DEBUG_PARAM_DTE",
      "Filtro de DTE",
      JSON.stringify({ modoTeste, alvo, faixa:{min:faixaMin, max:faixaMax} })
  );

  const selecionados = [];

  valores.forEach(r => {
    var dteNum = Number(r[IDX.dte]) || 0;

    if (modoTeste || (dteNum >= faixaMin && dteNum <= faixaMax)) {

      var item = {
        ticker: r[IDX.ticker],
        dueDate: r[IDX.due],
        spot: Number(r[IDX.spot]),
        dte: dteNum,
        putSymbol: r[IDX.ps],
        putStrike: Number(r[IDX.pst]),
        putClose: Number(r[IDX.pcl]),
        putDist: Number(r[IDX.pd]),
        callSymbol: r[IDX.cs],
        callStrike: Number(r[IDX.cst]),
        callClose: Number(r[IDX.ccl]),
        callDist: Number(r[IDX.cd]),
        totalCredit: Number(r[IDX.tot])
      };

      selecionados.push(item);

      log(SERVICO, "DEBUG_SELECIONADO",
          "Linha selecionada",
          JSON.stringify({
            ticker:item.ticker,
            dte:item.dte,
            put:item.putSymbol,
            call:item.callSymbol
          })
      );
    }
  });

  if (!selecionados.length) {
    log(SERVICO, "AVISO", "Nenhum strangle dentro da tolerância", "");
    return;
  }

  log(SERVICO, "DEBUG_PRE_ENVIO",
      "Preparando envio",
      JSON.stringify({
        total: selecionados.length,
        destino: cfg.emailAlerta,
        modoTeste
      })
  );

  var assunto = (modoTeste ? "[TESTE] " : "") + "Oportunidades Short Strangles";
  var html = montarHtmlAlertaStrangles(selecionados, cfg.diasAntes, modoTeste);

  MailApp.sendEmail({
    to: cfg.emailAlerta,
    subject: assunto,
    htmlBody: html
  });

  log(SERVICO, "SUCESSO_ENVIO",
      "E-mail enviado",
      JSON.stringify({
        destino: cfg.emailAlerta,
        total: selecionados.length,
        modoTeste,
        ts:new Date().toISOString()
      })
  );
}

/************************************************************
 * PIPELINE FINAL
 ************************************************************/
function rodarPipelineStranglesComAlerta() {
  const SERVICO = "PipelineStrangles";

  log(SERVICO, "INFO",
      "Iniciando pipeline completo (Selecao_Opcoes → Strangles → Alerta)",
      ""
  );

  try {
    selecionarOpcoesParaVencimento();
    log(SERVICO, "DEBUG", "Selecao_Opcoes atualizada", "");

    gerarShortStrangles();
    log(SERVICO, "DEBUG", "Selecao_Strangles gerada", "");

    enviarAlertasStrangles(false);

    log(SERVICO, "SUCESSO", "Pipeline finalizado com sucesso", "");

  } catch(e){
    log(SERVICO, "ERRO_CRITICO", "Erro fatal no pipeline", e.message);
    throw e;
  }
}

/************************************************************
 * TESTE — IGNORA FILTRO DE DTE
 ************************************************************/
function testarEnvioEmailStrangles() {
  enviarAlertasStrangles(true);
}
