/*
 * SELECAO_STRANGLES.GS — GERAÇÃO DE SHORT STRANGLES
 * Neste arquivo ficam apenas:
 *  - Constantes globais
 *  - Geração da aba Selecao_Strangles
 *  - Seleção de melhores PUT/CALL
 *  - Funções auxiliares específicas de Strangles
 */

const ABA_SELECAO_STRANGLES = "Selecao_Strangles";

const HEADERS_STRANGLES = [
  "ticker",
  "due_date",
  "spot_price",
  "days_to_maturity",
  "put_symbol",
  "put_strike",
  "put_close",
  "put_dist_pct",
  "call_symbol",
  "call_strike",
  "call_close",
  "call_dist_pct",
  "total_credit"
];

/************************************************************
 * CONFIG DO RANGE DE DISTÂNCIA (Config_Global)
 ************************************************************/
function getConfigMap_Strangles(abaConfig) {
  const ultima = abaConfig.getLastRow();
  if (ultima < 2) {
    return {
      Strangle_Dist_PUT_Min: 0.05,
      Strangle_Dist_PUT_Max: 0.30,
      Strangle_Dist_CALL_Min: 0.05,
      Strangle_Dist_CALL_Max: 0.30
    };
  }

  const dados = abaConfig.getRange(2, 1, ultima - 1, 2).getValues();
  const map = {};
  dados.forEach(l => { if (l[0]) map[l[0]] = l[1]; });

  return {
    Strangle_Dist_PUT_Min: Number(String(map["Strangle_Dist_PUT_Min"]  || "0.05").replace(",", ".")),
    Strangle_Dist_PUT_Max: Number(String(map["Strangle_Dist_PUT_Max"]  || "0.30").replace(",", ".")),
    Strangle_Dist_CALL_Min: Number(String(map["Strangle_Dist_CALL_Min"] || "0.05").replace(",", ".")),
    Strangle_Dist_CALL_Max: Number(String(map["Strangle_Dist_CALL_Max"] || "0.30").replace(",", "."))
  };
}

/************************************************************
 * RESET DA ABA SELECAO_STRANGLES
 ************************************************************/
function resetAbaStrangles() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let aba = ss.getSheetByName(ABA_SELECAO_STRANGLES);
  if (!aba) aba = ss.insertSheet(ABA_SELECAO_STRANGLES);

  aba.clear();
  aba.getRange(1, 1, 1, HEADERS_STRANGLES.length).setValues([HEADERS_STRANGLES]);
  return aba;
}

/************************************************************
 * ESCOLHA DA MELHOR OPÇÃO (PUT/CALL)
 ************************************************************/
function escolherMelhorOpcaoPorDistancia(opcoes, min, max) {
  if (!opcoes || !opcoes.length) return null;

  const dentro = opcoes.filter(o => o.dist >= min && o.dist <= max);

  if (dentro.length > 0) {
    dentro.sort((a, b) => b.close - a.close);
    return dentro[0];
  }

  opcoes.sort((a, b) => b.dist - a.dist);
  return opcoes[0];
}

/************************************************************
 * FUNÇÃO PRINCIPAL — GERAÇÃO DOS SHORT STRANGLES
 ************************************************************/
function gerarShortStrangles() {
  const SERVICO = "SelecaoStrangles";

  try {
    log(SERVICO, "INFO", "Iniciando geração de Short Strangles...", "");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaOpcoes = ss.getSheetByName("Selecao_Opcoes");
    const abaConfig = ss.getSheetByName("Config_Global");

    if (!abaOpcoes || !abaConfig) {
      log(SERVICO, "ERRO_CRITICO", "Aba Selecao_Opcoes ou Config_Global não encontrada", "");
      throw new Error("Abas necessárias não encontradas.");
    }

    const cfg = getConfigMap_Strangles(abaConfig);
    log(SERVICO, "DEBUG_CONFIG", "Parâmetros de seleção", JSON.stringify(cfg));

    const ultimaLinha = abaOpcoes.getLastRow();
    if (ultimaLinha <= 1) {
      log(SERVICO, "AVISO", "Selecao_Opcoes vazia.", "");
      return;
    }

    const dados = abaOpcoes.getRange(1, 1, ultimaLinha, abaOpcoes.getLastColumn()).getValues();
    const header = dados[0];
    const linhas = dados.slice(1);

    const idxTicker    = header.indexOf("ticker");
    const idxSymbol    = header.indexOf("symbol");
    const idxCategory  = header.indexOf("category");
    const idxDueDate   = header.indexOf("due_date");
    const idxStrike    = header.indexOf("strike");
    const idxSpotPrice = header.indexOf("spot_price");
    const idxDTE       = header.indexOf("days_to_maturity");
    const idxClose     = header.indexOf("close");

    const grupos = {};

    linhas.forEach(l => {
      const ticker = l[idxTicker];
      const due    = l[idxDueDate];
      const spot   = Number(l[idxSpotPrice]);
      if (!ticker || !due || !spot) return;

      const key = ticker + "__" + due;

      if (!grupos[key]) {
        grupos[key] = {
          ticker,
          due_date: due,
          spot,
          days_to_maturity: l[idxDTE],
          rows: []
        };
      }
      grupos[key].rows.push(l);
    });

    const resultado = [];

    Object.keys(grupos).forEach(key => {
      const g = grupos[key];

      const puts  = [];
      const calls = [];

      g.rows.forEach(l => {
        const cat    = String(l[idxCategory]).toUpperCase();
        const strike = Number(l[idxStrike]);
        const close  = Number(l[idxClose]) || 0;

        if (cat === "PUT" && strike < g.spot)
          puts.push({ linha:l, strike, close, dist:(g.spot - strike)/g.spot });

        if (cat === "CALL" && strike > g.spot)
          calls.push({ linha:l, strike, close, dist:(strike - g.spot)/g.spot });
      });

      const bestPut  = escolherMelhorOpcaoPorDistancia(puts,  cfg.Strangle_Dist_PUT_Min,  cfg.Strangle_Dist_PUT_Max);
      const bestCall = escolherMelhorOpcaoPorDistancia(calls, cfg.Strangle_Dist_CALL_Min, cfg.Strangle_Dist_CALL_Max);

      if (!bestPut || !bestCall) {
        log(SERVICO, "DEBUG_DESCARTADO",
            "Sem PUT ou CALL válidas",
            JSON.stringify({ ticker: g.ticker, due: g.due_date }));
        return;
      }

      const pl = bestPut.linha;
      const cl = bestCall.linha;

      resultado.push([
        g.ticker,
        g.due_date,
        g.spot,
        g.days_to_maturity,
        pl[idxSymbol],
        bestPut.strike,
        bestPut.close,
        bestPut.dist,
        cl[idxSymbol],
        bestCall.strike,
        bestCall.close,
        bestCall.dist,
        bestPut.close + bestCall.close
      ]);
    });

    const abaStr = resetAbaStrangles();

    if (!resultado.length) {
      log(SERVICO, "AVISO", "Nenhum strangle encontrado.", "");
      return;
    }

    abaStr.getRange(2, 1, resultado.length, HEADERS_STRANGLES.length).setValues(resultado);

    log(SERVICO, "SUCESSO", "Short Strangles gerados", JSON.stringify({ total: resultado.length }));

  } catch (e) {
    log("SelecaoStrangles", "ERRO_CRITICO", "Erro no processo", e.message);
    throw e;
  }
}


/* ============================================================================ 
 * ORQUESTRADOR → GERAR SHORT STRANGLES
 * ============================================================================ */

/**
 * Função disparada pelo menu para gerar Short Strangles.
 */
function orq_GerarShortStrangles() {
  const SERVICO = "ORQUESTRADOR_STRANGLES";
  const inicio = Date.now();

  try {
    log(SERVICO, "INFO", "Iniciando serviço: gerarShortStrangles()", "");

    // Chama o módulo Selecao_Strangles.gs
    gerarShortStrangles();

    const duracao = Date.now() - inicio;
    log(
      SERVICO,
      "SUCESSO",
      "Short Strangles gerados com sucesso",
      JSON.stringify({
        duracao_ms: duracao,
        timestamp_fim: new Date().toISOString()
      })
    );


    // REMOVIDO PARA MANTER O BACKEND SILENCIOSO:
    /*
    safeAlert_(
      "Short Strangles Gerados!",
      "O relatório está disponível na aba Selecao_Strangles.\n\nDuração: " +
      duracao + "ms"
    );
    */

  } catch (erro) {
    const duracao = Date.now() - inicio;
    log(
      SERVICO,
      "ERRO_CRITICO",
      "Erro ao gerar Short Strangles",
      erro.message + " | STACK: " + erro.stack
    );

    safeAlert_(
      "Erro ao Gerar Short Strangles",
      erro.message
    );

    throw erro;
  }
}