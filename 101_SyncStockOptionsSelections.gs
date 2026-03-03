/******************************************************
 *  SELECAO_OPCOES.GS - Versão Final Corrigida
 * ----------------------------------------------------
 *  Objetivo:
 *  Selecionar opções de PUT (strike < spot) e CALL
 *  (strike > spot), filtradas pelo vencimento desejado
 *  e limitadas pela quantidade configurada.
 *
 *  Correções incluídas:
 *  ✔ Filtro correto de PUTs (mais próximas ao spot)
 *  ✔ Ordenação pela proximidade ao spot
 *  ✔ Ordenação final por strike crescente (para exibição)
 *  ✔ Manutenção do padrão CALL (mais próximas ao spot)
 *  ✔ Documentação e logs aprimorados
 *
 *  IMPORTANTE:
 *  Todas as constantes globais como:
 *    - OPLAB_TOKEN_NAME
 *    - OPLAB_API_URL
 *    - ABA_DADOS_ATIVOS
 *    - ABA_CONFIG
 *    - log()
 *  já existem no Código.gs e NÃO devem ser repetidas aqui.
 ******************************************************/

const ABA_SELECAO_OPCOES = "Selecao_Opcoes";

// Cabeçalhos fixos
const HEADERS_SELECAO_OPCOES = [
  "ticker","symbol","name","open","high","low","close","volume",
  "financial_volume","trades","bid","ask","category","due_date",
  "maturity_type","strike","contract_size","exchange_id","created_at",
  "updated_at","variation","spot_price","isin","security_category",
  "market_maker","block_date","days_to_maturity","cnpj","bid_volume",
  "ask_volume","time","type","last_trade_at","strike_eod"
];

/******************************************************
 * FUNÇÃO PRINCIPAL
 ******************************************************/
function selecionarOpcoesParaVencimento() {

  const SERVICO = "SelecaoOpcoes";
  log(SERVICO, "INFO", "Iniciando seleção de opções...", "");

  // 1) Ler parâmetros da Config_Global
  const config = lerParametrosSelecao();
  log(SERVICO, "INFO_PARAMETROS", "Parâmetros lidos", JSON.stringify(config));

  if (!config.vencimentoISO) {
    log(SERVICO, "ERRO", "Regra_Vencimento_Entrada_Opcoes não configurada", "");
    throw new Error("Regra_Vencimento_Entrada_Opcoes não configurada em Config_Global");
  }

  // 2) Ler tickers da aba Dados_Ativos
  const tickers = lerTickersDadosAtivos();
  if (tickers.length === 0) {
    log(SERVICO, "AVISO", "Nenhum ticker encontrado em Dados_Ativos", "");
    return;
  }

  // 3) Preparar aba de saída
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const abaSaida = planilha.getSheetByName(ABA_SELECAO_OPCOES);

  if (!abaSaida) {
    throw new Error("A aba Selecao_Opcoes não existe.");
  }

  const lastRow = abaSaida.getLastRow();
  if (lastRow > 1) {
    abaSaida.getRange(2, 1, lastRow - 1, HEADERS_SELECAO_OPCOES.length).clearContent();
  }

  // 4) Token OPLAB
  const token = PropertiesService.getScriptProperties().getProperty(OPLAB_TOKEN_NAME);
  if (!token) {
    log(SERVICO, "ERRO", "Token OPLAB não encontrado", "");
    throw new Error("Token OPLAB não encontrado em PropertiesService");
  }

  let linhasInseridas = [];

  // 5) Processar cada ticker
  tickers.forEach(ticker => {

    log(SERVICO, "INFO", "Consultando opções para " + ticker, "");

    const opcoes = buscarOpcoesTicker(ticker, token, SERVICO);
    if (!opcoes || opcoes.length === 0) {
      log(SERVICO, "AVISO", "Nenhuma opção retornada pela API para " + ticker, "");
      return;
    }

    log(SERVICO, "DEBUG", `API retornou ${opcoes.length} registros para ${ticker}`, "");

    /**********************
     * 5.1 Filtrar por vencimento
     **********************/
    const filtradasVenc = opcoes.filter(op => op.due_date === config.vencimentoISO);

    log(
      SERVICO,
      "DEBUG_FILTRO_VENC",
      `Filtro vencimento para ${ticker}`,
      `venc=${config.vencimentoISO}, antes=${opcoes.length}, depois=${filtradasVenc.length}`
    );

    if (filtradasVenc.length === 0) return;

    /**********************
     * 5.2 Definir spot
     **********************/
    const spot = parseFloat(filtradasVenc[0].spot_price || 0);
    if (!spot || isNaN(spot)) {
      log(SERVICO, "AVISO", `Spot inválido para ${ticker}`, String(spot));
      return;
    }

    /**********************
     * 5.3 Separar PUTs e CALLs
     **********************/
    const puts = filtradasVenc.filter(op => op.category === "PUT"  && op.strike < spot);
    const calls = filtradasVenc.filter(op => op.category === "CALL" && op.strike > spot);

    log(SERVICO, "DEBUG_FILTRO_PUTCALL",
        `PUT abaixo / CALL acima para ${ticker}`,
        `PUT=${puts.length} CALL=${calls.length}`);

    if (puts.length === 0 && calls.length === 0) return;

    /**********************
     * 5.4 Ordenação: proximidade ao spot
     **********************/
    puts.sort((a, b) => Math.abs(spot - a.strike) - Math.abs(spot - b.strike));
    calls.sort((a, b) => Math.abs(spot - a.strike) - Math.abs(spot - b.strike));

    /**********************
     * 5.5 Aplicar limites de quantidade
     **********************/
    const putsLimitadas  = puts.slice(0, config.qtdMaxPUT);
    const callsLimitadas = calls.slice(0, config.qtdMaxCALL);

    /**********************
     * 5.6 Ordenação final estética por strike crescente
     **********************/
    putsLimitadas.sort((a, b) => a.strike - b.strike);
    callsLimitadas.sort((a, b) => a.strike - b.strike);

    /**********************
     * 5.7 Montar linhas de saída
     **********************/
    const todas = [...putsLimitadas, ...callsLimitadas];

    const linhas = todas.map(op => mapearLinha(ticker, op));
    linhasInseridas.push(...linhas);
  });

  /**********************
   * 6) Gravar na planilha
   **********************/
  if (linhasInseridas.length > 0) {
    abaSaida.getRange(
      2, 1,
      linhasInseridas.length,
      HEADERS_SELECAO_OPCOES.length
    ).setValues(linhasInseridas);

    log(SERVICO, "SUCESSO", `Total de linhas inseridas: ${linhasInseridas.length}`, "");

  } else {
    log(SERVICO, "AVISO", "Nenhuma opção encontrada após os filtros.", "");
  }
}

/******************************************************
 * Ponte compatível com Código.gs
 ******************************************************/
function buscarOpcoesParaSelecao() {
  selecionarOpcoesParaVencimento();
}

/******************************************************
 * Ler parâmetros da aba Config_Global
 ******************************************************/
function lerParametrosSelecao() {

  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const aba = planilha.getSheetByName(ABA_CONFIG);

  const lastRow = aba.getLastRow();
  if (lastRow < 2) {
    return {
      vencimentoISO: "",
      qtdMaxPUT: 15,
      qtdMaxCALL: 15
    };
  }

  const dados = aba.getRange(2, 1, lastRow - 1, 2).getValues();

  let params = {};
  dados.forEach(l => {
    const chave = l[0];
    const valor = l[1];
    if (chave) params[chave] = valor;
  });

  return {
    vencimentoISO: params["Regra_Vencimento_Entrada_Opcoes"],
    qtdMaxPUT:   parseInt(params["Regra_Qtd_Max_PUT"]   || 15),
    qtdMaxCALL:  parseInt(params["Regra_Qtd_Max_CALL"]  || 15)
  };
}

/******************************************************
 * Ler tickers da aba Dados_Ativos
 ******************************************************/
function lerTickersDadosAtivos() {

  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_DADOS_ATIVOS);
  if (!aba) return [];

  const lastRow = aba.getLastRow();
  if (lastRow <= 1) return [];

  const valores = aba.getRange(2, 1, lastRow - 1, 1).getValues();

  return valores.map(v => v[0]).filter(t => t);
}

/******************************************************
 * API /market/options/{ticker}
 ******************************************************/
function buscarOpcoesTicker(ticker, token, SERVICO) {

  const url = `${OPLAB_API_URL}/market/options/${ticker}`;

  try {
    const response = UrlFetchApp.fetch(url, {
      method: "get",
      headers: { "Access-Token": token },
      muteHttpExceptions: true
    });

    const code = response.getResponseCode();
    const txt  = response.getContentText();

    if (code !== 200) {
      log(
        SERVICO,
        "ERRO_API",
        `Erro ao consultar ${ticker}`,
        `HTTP ${code} - ${txt}`
      );
      return [];
    }

    return JSON.parse(txt);

  } catch (e) {
    log(SERVICO, "ERRO", "Erro na chamada API", e.message);
    return [];
  }
}

/******************************************************
 * Mapear linha de saída
 ******************************************************/
function mapearLinha(ticker, op) {

  const base = Object.assign({}, op, { ticker: ticker });

  return HEADERS_SELECAO_OPCOES.map(col => {
    const v = base[col];
    return (v === undefined || v === null) ? "" : v;
  });
}


// =====================================================================
// ORQUESTRADOR: SELEÇÃO DE OPÇÕES
// =====================================================================

function orquestrarSelecaoOpcoes() {
  const SERVICO = "SELECAO_OPCOES";
  const registro = SERVICOS_REGISTRY[SERVICO];
  const inicio = new Date();
  
  logOrquestrador("INICIO", "Iniciando Seleção de Opções (DTE)", {
    servico: registro.nome,
    timestamp: inicio.toISOString()
  });

  try {
    const resultado = executarComTimeout(registro.funcao, registro.timeout);

    const fim = new Date();
    const duracao = fim - inicio;

    logOrquestrador("SUCESSO", "Seleção de opções concluída", {
      duracao_ms: duracao,
      timestamp_fim: fim.toISOString()
    });

    return resultado;

  } catch (erro) {
    const fim = new Date();

    logOrquestrador("ERRO_CRITICO", "Falha na Seleção de Opções", {
      erro: erro.message,
      stack: erro.stack,
      timestamp_fim: fim.toISOString()
    });

    safeAlert_(
      '❌ Erro na Seleção de Opções',
      'Ocorreu um erro durante a seleção de opções:\n\n' +
      erro.message +
      '\n\nVerifique a aba Logs para mais detalhes.'
    );
  }
}