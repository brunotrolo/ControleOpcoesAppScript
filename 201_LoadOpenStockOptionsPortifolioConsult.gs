/*
 * ORQUESTRADOR_ALERTAS_OPERACOESMODEL.GS
 * --------------------------------------
 * Modelos e funções de transformação para o relatório
 * de operações em aberto (aba Cockpit).
 *
 * Não altera nenhum fluxo existente.
 */

const ABA_COCKPIT = "Cockpit";
const LINHA_HEADER_COCKPIT = 10;
const LINHA_DADOS_INICIO_COCKPIT = 11;

/**
 * Carrega todas as pernas em aberto da aba Cockpit
 * (filtra por vencimento >= hoje e quantidade > 0).
 */
function carregarPernasAbertasDaCockpit() {
  const SERVICO = "OperacoesModel";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(ABA_COCKPIT);

  if (!aba) {
    log(SERVICO, "ERRO_CRITICO", "Aba Cockpit não encontrada", "");
    return [];
  }

  const ultimaLinha = aba.getLastRow();
  const ultimaColuna = aba.getLastColumn();

  if (ultimaLinha < LINHA_DADOS_INICIO_COCKPIT) {
    log(SERVICO, "AVISO", "Cockpit sem dados (menos que a linha de início)", "");
    return [];
  }


  // Lê o cabeçalho na linha 10
  const header = aba
    .getRange(LINHA_HEADER_COCKPIT, 1, 1, ultimaColuna)
    .getValues()[0];

  // Mapa: NOME_NORMALIZADO -> índice da coluna
  const headerIdx = {};
  header.forEach(function (nome, idx) {
    const key = normalizeHeaderName_(nome);
    if (key) headerIdx[key] = idx;
  });

  // Nomes "lógicos" das colunas obrigatórias (normalizados)
  const COL = {
    ID_TRADE:        normalizeHeaderName_("ID TRADE"),
    ID_ESTRATEGIA:   normalizeHeaderName_("ID ESTRATÉGIA"),
    CODIGO_OPCAO:    normalizeHeaderName_("CÓDIGO OPÇÃO"),
    TICKER:          normalizeHeaderName_("TICKER"),
    VENDA_COMPRA:    normalizeHeaderName_("VENDA/COMPRA"),
    VENCIMENTO:      normalizeHeaderName_("VENCIMENTO"),
    STRIKE:          normalizeHeaderName_("STRIKE"),
    TIPO:            normalizeHeaderName_("TIPO"),
    QTD:             normalizeHeaderName_("QTD"),
    PREMIO_PM:       normalizeHeaderName_("PRÊMIO (PM)"),
    PREMIO_ATUAL:    normalizeHeaderName_("PRÊMIO ATUAL"),
    PL_ATUAL:        normalizeHeaderName_("P/L ATUAL"),
    TICKER_SPOT:     normalizeHeaderName_("TICKER SPOT"),
    DTE_CORRIDOS:    normalizeHeaderName_("DTE CORRIDOS"),
    MONEYNESS_CODE:  normalizeHeaderName_("MONEYNESS_CODE"),
    DELTA:           normalizeHeaderName_("DELTA")
  };

  // Garantia mínima de colunas
  var missing = [];
  if (headerIdx[COL.ID_TRADE]        === undefined) missing.push("ID_TRADE");
  if (headerIdx[COL.ID_ESTRATEGIA]   === undefined) missing.push("ID_ESTRATÉGIA");
  if (headerIdx[COL.CODIGO_OPCAO]    === undefined) missing.push("CÓDIGO OPÇÃO");
  if (headerIdx[COL.TICKER]          === undefined) missing.push("TICKER");
  if (headerIdx[COL.VENDA_COMPRA]    === undefined) missing.push("VENDA/COMPRA");
  if (headerIdx[COL.VENCIMENTO]      === undefined) missing.push("VENCIMENTO");
  if (headerIdx[COL.STRIKE]          === undefined) missing.push("STRIKE");
  if (headerIdx[COL.TIPO]            === undefined) missing.push("TIPO");
  if (headerIdx[COL.QTD]             === undefined) missing.push("QTD");
  if (headerIdx[COL.PREMIO_PM]       === undefined) missing.push("PRÊMIO (PM)");
  if (headerIdx[COL.PREMIO_ATUAL]    === undefined) missing.push("PRÊMIO ATUAL");
  if (headerIdx[COL.PL_ATUAL]        === undefined) missing.push("P/L ATUAL");
  if (headerIdx[COL.TICKER_SPOT]     === undefined) missing.push("TICKER SPOT");
  if (headerIdx[COL.DTE_CORRIDOS]    === undefined) missing.push("DTE CORRIDOS");
  if (headerIdx[COL.MONEYNESS_CODE]  === undefined) missing.push("MONEYNESS_CODE");
  if (headerIdx[COL.DELTA]           === undefined) missing.push("DELTA");

  if (missing.length > 0) {
    log(
      SERVICO,
      "ERRO_CRITICO",
      "Estrutura da aba Cockpit incompatível. Colunas obrigatórias ausentes: " +
        missing.join(", "),
      header.join(" | ")
    );
    return [];
  }

  // Garantia mínima de colunas
  if (
    headerIdx[COL.ID_TRADE] === undefined ||
    headerIdx[COL.ID_ESTRATEGIA] === undefined ||
    headerIdx[COL.CODIGO_OPCAO] === undefined ||
    headerIdx[COL.TICKER] === undefined ||
    headerIdx[COL.VENDA_COMPRA] === undefined ||
    headerIdx[COL.VENCIMENTO] === undefined ||
    headerIdx[COL.STRIKE] === undefined ||
    headerIdx[COL.TIPO] === undefined ||
    headerIdx[COL.QTD] === undefined ||
    headerIdx[COL.PREMIO_PM] === undefined ||
    headerIdx[COL.PREMIO_ATUAL] === undefined ||
    headerIdx[COL.PL_ATUAL] === undefined ||
    headerIdx[COL.TICKER_SPOT] === undefined ||
    headerIdx[COL.DTE_CORRIDOS] === undefined ||
    headerIdx[COL.MONEYNESS_CODE] === undefined ||
    headerIdx[COL.DELTA] === undefined
  ) {
    log(
      SERVICO,
      "ERRO_CRITICO",
      "Estrutura da aba Cockpit incompatível",
      header.join(" | ")
    );
    return [];
  }

    const qtdLinhasDados = ultimaLinha - LINHA_HEADER_COCKPIT;
    const dados = aba
      .getRange(LINHA_DADOS_INICIO_COCKPIT, 1, qtdLinhasDados, ultimaColuna)
      .getValues();

    // 🔥 Normaliza hoje apenas 1 vez
    const hojeLimpo = truncarData_(new Date());

    const pernas = [];

    dados.forEach(function (linha) {

    const leg = parseLegFromCockpitRow_(linha, headerIdx, COL);
    if (!leg) return;

    // Considera aberta se vencimento >= hoje e quantidade > 0
    if (!leg.vencimento || leg.quantidade <= 0) return;

    // Normaliza corretamente
    var venc = truncarData_(new Date(leg.vencimento));

    // Mantém vencimentos HOJE ou futuramente (corrige DTE = 0)
    if (venc.getTime() < hojeLimpo.getTime()) return;

    pernas.push(leg);
  });

  log(
    SERVICO,
    "INFO",
    "Pernas em aberto carregadas da Cockpit",
    JSON.stringify({ totalPernasAbertas: pernas.length })
  );

  return pernas;
}

/**
 * Converte uma linha da Cockpit em um objeto OptionLeg.
 */
function parseLegFromCockpitRow_(linha, idx, COL) {
  try {
    const idTrade = linha[idx[COL.ID_TRADE]];
    const idEstrutura = linha[idx[COL.ID_ESTRATEGIA]];
    const ticker = linha[idx[COL.TICKER]];
    const codigoOpcao = linha[idx[COL.CODIGO_OPCAO]];
    const op = String(linha[idx[COL.VENDA_COMPRA]] || "").toUpperCase().trim();
    const tipo = String(linha[idx[COL.TIPO]] || "").toUpperCase().trim();

    const strike = opToNumber_(linha[idx[COL.STRIKE]]);
    const qtd = opToNumber_(linha[idx[COL.QTD]]);
    const premioPM = opToNumber_(linha[idx[COL.PREMIO_PM]]);
    const premioAtual = opToNumber_(linha[idx[COL.PREMIO_ATUAL]]);
    const plAtual = opToNumber_(linha[idx[COL.PL_ATUAL]]);
    const tickerSpot = opToNumber_(linha[idx[COL.TICKER_SPOT]]);
    const dte = opToNumber_(linha[idx[COL.DTE_CORRIDOS]]);
    const delta = opToNumber_(linha[idx[COL.DELTA]]);
    const moneyCode = String(linha[idx[COL.MONEYNESS_CODE]] || "").toUpperCase().trim();
    const vencimentoBruto = linha[idx[COL.VENCIMENTO]];

    if (!idEstrutura || !ticker || !codigoOpcao || !tipo || !op) {
      return null;
    }

    const vencimento = opParseDate_(vencimentoBruto);

    // Define direcao
    var direcao = "";
    if (op === "VENDA" && tipo === "PUT") direcao = "short_put";
    if (op === "VENDA" && tipo === "CALL") direcao = "short_call";
    if (op === "COMPRA" && tipo === "PUT") direcao = "long_put";
    if (op === "COMPRA" && tipo === "CALL") direcao = "long_call";

    return {
      idTrade: idTrade,
      idEstrutura: idEstrutura,
      ticker: ticker,
      tickerSpot: tickerSpot,
      codigoOpcao: codigoOpcao,
      tipo: tipo,                 // PUT/CALL
      operacao: op,               // VENDA/COMPRA
      strike: strike,
      quantidade: qtd,
      premioPM: premioPM,
      premioAtual: premioAtual,
      plAtual: plAtual,
      delta: delta,
      moneynessCode: moneyCode,   // ITM/ATM/OTM
      vencimento: vencimento,
      dte: dte,
      direcao: direcao
    };

  } catch (e) {
    log("OperacoesModel", "ERRO", "Falha ao parsear linha do Cockpit", e.message);
    return null;
  }
}

/**
 * Agrupa pernas em estruturas por ID_ESTRUTURA.
 */
function agruparPernasEmEstruturas(legs) {
  const SERVICO = "OperacoesModel";

  const mapa = {};
  legs.forEach(function (leg) {
    if (!mapa[leg.idEstrutura]) {
      mapa[leg.idEstrutura] = [];
    }
    mapa[leg.idEstrutura].push(leg);
  });

  const estruturas = [];
  Object.keys(mapa).forEach(function (idEstrutura) {
    const pernas = mapa[idEstrutura];
    const estrutura = gerarEstruturaAPartirDasPernas_(pernas);
    if (estrutura) estruturas.push(estrutura);
  });

  log(
    SERVICO,
    "INFO",
    "Estruturas geradas a partir das pernas",
    JSON.stringify({ totalEstruturas: estruturas.length })
  );

  // Ordena por vencimento ascendente (DTE menor primeiro)
  estruturas.sort(function(a, b) {
    return a.vencimento - b.vencimento;
  });

  return estruturas;

}

/**
 * Monta uma OptionStructure a partir de uma lista de pernas.
 */
function gerarEstruturaAPartirDasPernas_(pernas) {
  if (!pernas || pernas.length === 0) return null;

  var base = pernas[0];
  var ticker = base.ticker;
  var tickerSpot = base.tickerSpot;
  var vencimento = base.vencimento;

  var dteMin = (base.dte || 0);
  var premioTotal = 0;
  var custoZeragem = 0;
  var plTotal = 0;
  var deltaTotal = 0;

  pernas.forEach(function (p) {
    // DTE mínimo da estrutura
    if (!isNaN(p.dte) && (p.dte < dteMin || isNaN(dteMin))) {
      dteMin = p.dte;
    }

    // Sinal de prêmio: VENDA positivo, COMPRA negativo
    var sinalPremio = (p.operacao === "VENDA") ? 1 : -1;

    premioTotal += (p.premioPM * p.quantidade * sinalPremio);

    // Custo para zerar: VENDA precisa recomprar (custo positivo),
    // COMPRA receberia na venda (custo negativo)
    custoZeragem += (p.premioAtual * p.quantidade * (-sinalPremio));

    plTotal += p.plAtual;
    deltaTotal += (p.delta * p.quantidade);
  });

  var tipoEstrutura = classificarEstrutura_(pernas);

  return {
    idEstrutura: base.idEstrutura,
    ticker: ticker,
    tickerSpot: tickerSpot,
    vencimento: vencimento,
    dte: dteMin,
    pernas: pernas,
    premioTotal: premioTotal,
    custoZeragem: custoZeragem,
    plTotal: plTotal,
    deltaTotal: deltaTotal,
    qtdPernas: pernas.length,
    tipoEstrutura: tipoEstrutura,
    risco: "" // reservado para futuro
  };
}

/**
 * Classifica a estrutura de forma simples (opcional).
 */
function classificarEstrutura_(pernas) {
  var shortPut = 0, shortCall = 0, longPut = 0, longCall = 0;

  pernas.forEach(function (p) {
    if (p.direcao === "short_put") shortPut++;
    if (p.direcao === "short_call") shortCall++;
    if (p.direcao === "long_put") longPut++;
    if (p.direcao === "long_call") longCall++;
  });

  if (shortPut === 1 && shortCall === 0 && longPut === 0 && longCall === 0) {
    return "naked_put";
  }
  if (shortCall === 1 && shortPut === 0 && longPut === 0 && longCall === 0) {
    return "naked_call";
  }
  if (shortPut === 1 && shortCall === 1 && longPut === 0 && longCall === 0) {
    return "short_strangle";
  }
  if (longPut === 1 && shortPut === 1 && shortCall === 0 && longCall === 0) {
    return "vertical_put";
  }
  if (longCall === 1 && shortCall === 1 && shortPut === 0 && longPut === 0) {
    return "vertical_call";
  }
  if (pernas.length === 4 && shortPut && longPut && shortCall && longCall) {
    return "iron_condor";
  }

  return "multi_leg";
}

/**
 * Gera o resumo geral das operações abertas (DailySummary).
 */
function gerarResumoGeralOperacoes(estruturas) {
  var resumo = {
    totalEstruturas: estruturas.length,
    lucroTotal: 0,
    deltaTotal: 0,
    moneyness: {
      call: { ITM: 0, ATM: 0, OTM: 0, total: 0 },
      put:  { ITM: 0, ATM: 0, OTM: 0, total: 0 }
    },
    vencimentos: {
      curto:  { count: 0, pct: 0 },
      medio:  { count: 0, pct: 0 },
      longo:  { count: 0, pct: 0 }
    },
    concentracaoPorTicker: [],
    resultados: {
      noLucro:    { count: 0, pct: 0 },
      noPrejuizo: { count: 0, pct: 0 }
    }
  };

  if (estruturas.length === 0) return resumo;

  // LUCRO TOTAL / DELTA TOTAL / LUCRO VS PREJUÍZO
  estruturas.forEach(function (est) {
    resumo.lucroTotal += est.plTotal;
    resumo.deltaTotal += est.deltaTotal;

    if (est.plTotal > 0) resumo.resultados.noLucro.count++;
    if (est.plTotal < 0) resumo.resultados.noPrejuizo.count++;
  });

  // Pct lucro/prejuízo
  resumo.resultados.noLucro.pct =
    Math.round((resumo.resultados.noLucro.count / resumo.totalEstruturas) * 100);
  resumo.resultados.noPrejuizo.pct =
    Math.round((resumo.resultados.noPrejuizo.count / resumo.totalEstruturas) * 100);

  // Moneyness: apenas VENDAS
  estruturas.forEach(function (est) {
    est.pernas.forEach(function (p) {
      if (p.operacao !== "VENDA") return;
      var code = p.moneynessCode || "";
      if (!code) return;

      if (p.tipo === "CALL") {
        if (code === "ITM") resumo.moneyness.call.ITM++;
        if (code === "ATM") resumo.moneyness.call.ATM++;
        if (code === "OTM") resumo.moneyness.call.OTM++;
      }
      if (p.tipo === "PUT") {
        if (code === "ITM") resumo.moneyness.put.ITM++;
        if (code === "ATM") resumo.moneyness.put.ATM++;
        if (code === "OTM") resumo.moneyness.put.OTM++;
      }
    });
  });

  resumo.moneyness.call.total =
    resumo.moneyness.call.ITM + resumo.moneyness.call.ATM + resumo.moneyness.call.OTM;
  resumo.moneyness.put.total =
    resumo.moneyness.put.ITM + resumo.moneyness.put.ATM + resumo.moneyness.put.OTM;

  // VENCIMENTOS (DTE)
  var curto = 0, medio = 0, longo = 0;
  estruturas.forEach(function (est) {
    var d = est.dte || 0;
    if (d < 30) curto++;
    else if (d >= 31 && d <= 90) medio++;
    else if (d > 90) longo++;
  });

  resumo.vencimentos.curto.count = curto;
  resumo.vencimentos.medio.count = medio;
  resumo.vencimentos.longo.count = longo;

  resumo.vencimentos.curto.pct = Math.round((curto / resumo.totalEstruturas) * 100);
  resumo.vencimentos.medio.pct = Math.round((medio / resumo.totalEstruturas) * 100);
  resumo.vencimentos.longo.pct = Math.round((longo / resumo.totalEstruturas) * 100);

  // CONCENTRAÇÃO POR TICKER
  var mapaTicker = {};
  estruturas.forEach(function (est) {
    var t = est.ticker || "OUTROS";
    mapaTicker[t] = (mapaTicker[t] || 0) + 1;
  });

  var concArr = [];
  Object.keys(mapaTicker).forEach(function (t) {
    concArr.push({
      ticker: t,
      count: mapaTicker[t],
      pct: Math.round((mapaTicker[t] / resumo.totalEstruturas) * 100)
    });
  });

  concArr.sort(function (a, b) { return b.count - a.count; });
  resumo.concentracaoPorTicker = concArr;

  return resumo;
}

/**
 * Monta o objeto final DailyReport para o template.
 */
function montarDailyReportOperacoes(estruturas) {
  var resumo = gerarResumoGeralOperacoes(estruturas);

  return {
    dataReferencia: new Date(),
    resumo: resumo,
    estruturas: estruturas
  };
}

/* =====================================================================
 * AUXILIARES INTERNOS (somente deste módulo)
 * ===================================================================== */

function opToNumber_(v) {
  if (v === null || v === "" || v === undefined) return 0;
  if (typeof v === "number") return v;
  var s = String(v).replace(".", "").replace(",", ".");
  var n = Number(s);
  return isNaN(n) ? 0 : n;
}

function opParseDate_(valor) {
  if (!valor) return null;

  if (Object.prototype.toString.call(valor) === "[object Date]") {
    if (isNaN(valor.getTime())) return null;
    return valor;
  }

  var s = String(valor).trim();

  // dd/MM/yyyy
  var partes = s.split("/");
  if (partes.length === 3) {
    var d = parseInt(partes[0], 10);
    var m = parseInt(partes[1], 10) - 1;
    var y = parseInt(partes[2], 10);
    var dt = new Date(y, m, d);
    if (!isNaN(dt.getTime())) return dt;
  }

  // fallback
  var dt2 = new Date(s);
  if (!isNaN(dt2.getTime())) return dt2;

  return null;
}

  function truncarData_(d) {
    var x = new Date(d);
    x.setHours(0, 0, 0, 0);
    return x;
  }


/**
 * Normaliza o nome de uma coluna para comparação:
 * - Uppercase
 * - Remove acentos
 * - Troca _ / - ( ) por espaço
 * - Remove espaços duplicados
 */
function normalizeHeaderName_(valor) {
  if (!valor) return "";

  var s = String(valor).toUpperCase().trim();

  // Troca caracteres especiais por espaço
  s = s.replace(/[_\/\-]/g, " ");
  s = s.replace(/[()]/g, " ");

  // Remove acentos
  s = s
    .replace(/[ÁÀÃÂÄ]/g, "A")
    .replace(/[ÉÈÊË]/g, "E")
    .replace(/[ÍÌÎÏ]/g, "I")
    .replace(/[ÓÒÕÔÖ]/g, "O")
    .replace(/[ÚÙÛÜ]/g, "U")
    .replace(/[Ç]/g, "C");

  // Remove espaços duplicados
  s = s.replace(/\s+/g, " ");

  return s;
}