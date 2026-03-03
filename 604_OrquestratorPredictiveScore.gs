/**
 * SERVIÇO 604: Orquestrador de Score Preditivo
 * Responsabilidade: Consolidar pesos, calcular projeções e definir o Status Operacional formal.
 * Dependências: 002_CoreServiceLogger_DataExtractor
 */

const SCORE_ORCH_CONFIG = {
  servico: "604_OrquestratorPredictiveScore",
  abaDestino: "Pontuacao_Preditiva_Consolidada",
  pesos: {
    estatistico: 0.30,
    fundamentalista: 0.30,
    tecnico: 0.25,
    macro: 0.15
  }
};

/**
 * Executa a consolidação final da Análise Preditiva
 * @param {number} diasParam - (Opcional) Horizonte de tempo em dias (Ex: 5, 45)
 */
function orquestratorPredictiveScore_Execute(diasParam) {
  const t0 = Date.now();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetConsolidada = ss.getSheetByName(SCORE_ORCH_CONFIG.abaDestino);
  
  if (!sheetConsolidada) return;

  const horizonteDias = diasParam || 45; // Fallback caso rode de forma manual

  try {
    // 1. Captura de dados das abas especialistas
    const dadosEstatistica = getMapFromSheet_(ss, "Analise_Estatistica_Ativos");
    const dadosFundamentais = getMapFromSheet_(ss, "Analise_Fundamentalista_Ativos");
    const dadosTecnicos = getMapFromSheet_(ss, "Dados_Ativos_Historico_Tendencia");
    const dadosMacro = getMacroStatus_(ss); 

    const ativosParaProcessar = Object.keys(dadosEstatistica);
    const resultados = [];
    const agora = new Date();

    // 2. Cabeçalhos Oficiais da Aba Consolidada (Garante a estrutura pro Front-end)
    const headers = [
      "Ticker", "Timestamp_Calculo", "Preço_Spot", "Score_Estatistico", 
      "Score_Fundamentalista", "Score_Tecnico", "Score_Macro", "Config_Pesos", 
      "Score_Consolidado", "Probabilidade_Alta", "Probabilidade_Baixa", 
      "Expectativa_Matematica", "Spot_Min_Proj", "Spot_Max_Proj", 
      "Fator_Dominante", "Status_Operacional"
    ];

    ativosParaProcessar.forEach(ticker => {
      const estat = dadosEstatistica[ticker];
      const fund = dadosFundamentais[ticker] || {};
      const tec = dadosTecnicos[ticker] || {};

      // Normalização de Scores (Escala 0-10)
      const sMac = dadosMacro.scoreGlobal; 
      const sEst = normalizarScoreEstatistico_(estat);
      const sFun = parseFloat(fund.SCORE_FUNDAMENTALISTA || fund.SCORE_FUND || 0);
      const sTec = normalizarScoreTecnico_(tec, sMac);

      // Cálculo do Score Consolidado Ponderado
      const scoreFinal = (
        (sEst * SCORE_ORCH_CONFIG.pesos.estatistico) +
        (sFun * SCORE_ORCH_CONFIG.pesos.fundamentalista) +
        (sTec * SCORE_ORCH_CONFIG.pesos.tecnico) +
        (sMac * SCORE_ORCH_CONFIG.pesos.macro)
      );

      // Captura de métricas estatísticas e projeções calculadas no 602
      const spot = parseFloat(estat.PREÇO_SPOT || estat.SPOT || 0);
      const sigma = parseFloat(estat.SIGMA_45D || 0); // Mantém a chave de leitura antiga
      const spotMin = parseFloat(estat.PROB_1SIGMA_INFERIOR || 0);
      const spotMax = parseFloat(estat.PROB_1SIGMA_SUPERIOR || 0);

      // Probabilidades e Expectativa Matemática
      const probAlta = scoreFinal / 10;
      const probBaixa = 1 - probAlta;
      const alvoEstimado = spot * (1 + sigma);
      const riscoEstimado = spot * (1 - sigma);
      const expectativaMat = (probAlta * (alvoEstimado - spot)) - (probBaixa * (spot - riscoEstimado));

      // 🌟 NOVO: Identificação Automática do Fator Dominante
      let fatorDominante = "EQUILIBRIO_TECNICO_MACRO";
      if (sFun >= sEst && sFun >= sTec && sFun > 0) {
        const margem = parseFloat(fund.MARGEM_SEGURANCA_GRAHAM || 0);
        fatorDominante = "FUNDAMENTALISTA (MARGEM " + (margem * 100).toFixed(0) + "%)";
      } else if (sEst >= sFun && sEst >= sTec && sEst > 0) {
        fatorDominante = "ESTATISTICO (VOLATILIDADE ATRATIVA)";
      } else if (sTec >= sFun && sTec >= sEst && sTec > 0) {
        fatorDominante = "TECNICO (TENDENCIA E ALINHAMENTO)";
      } else if (sFun === 0 && sTec <= 4) {
        fatorDominante = "RISCO (FUNDAMENTO/TENDENCIA FRACA)";
      }

      // 🌟 NOVO: Geração do Status Operacional Formal
      const veredito = gerarVeredito_(scoreFinal);

      resultados.push([
        ticker, agora, spot, sEst, sFun, sTec, sMac,
        JSON.stringify(SCORE_ORCH_CONFIG.pesos),
        scoreFinal, probAlta, probBaixa, expectativaMat, 
        spotMin, spotMax, fatorDominante, veredito
      ]);

      log(SCORE_ORCH_CONFIG.servico, "DEBUG", `Consolidado: ${ticker} (${horizonteDias} dias)`, JSON.stringify({
        score: scoreFinal.toFixed(2),
        status: veredito,
        dominancia: fatorDominante
      }));
    });

    // 3. Gravação Final
    if (resultados.length > 0) {
      sheetConsolidada.clearContents(); // Limpa tudo
      sheetConsolidada.getRange(1, 1, 1, headers.length).setValues([headers]); // Força cabeçalhos atualizados
      sheetConsolidada.getRange(2, 1, resultados.length, resultados[0].length).setValues(resultados);
    }

    log(SCORE_ORCH_CONFIG.servico, "SUCESSO", `Score consolidado para ${resultados.length} ativos (Horizonte: ${horizonteDias} dias)`, JSON.stringify({ tempo: (Date.now() - t0) + "ms" }));

  } catch (e) {
    log(SCORE_ORCH_CONFIG.servico, "ERRO_FATAL", "Falha na orquestração", e.toString());
  }
}

function getMacroStatus_(ss) {
  const sheet = ss.getSheetByName("Dados_Macro_Setorial");
  const data = sheet.getRange("A2:B15").getValues();
  let vix = 20; 

  data.forEach(r => {
    if (r[0] === "Índice do Medo (VIX)") vix = parseFloat(r[1]);
  });

  let scoreGlobal = vix < 20 ? 8 : (vix > 30 ? 3 : 5);
  return { scoreGlobal: scoreGlobal };
}

function normalizarScoreEstatistico_(d) {
  let s = 5; 
  const ratio = parseFloat(d.HV_IV_RATIO || d.HVIV_RATIO || 1);
  const prob = parseFloat(d.PROB_PERMANECER_RANGE || d.PROB_PERM || 0.68);
  
  if (ratio > 1.1) s += 2; 
  if (prob > 0.70) s += 3; 
  return Math.min(s, 10);
}

function normalizarScoreTecnico_(d, scoreMacro) {
  const tend = String(d.VEREDITO_TENDENCIA || "").toUpperCase();
  let s = tend === "ALTA" ? 9 : (tend === "BAIXA" ? 2 : 5);

  const beta = parseFloat(d.BETA_IBOV_60D || 1);
  const correl = parseFloat(d.CORREL_IBOV_60D || 1);

  if (tend === "ALTA" && scoreMacro <= 5 && (correl > 0.7 || beta > 1.2)) {
    s -= 3; 
  }

  return Math.max(0, Math.min(s, 10));
}

/**
 * 🌟 NOVO: Gera o veredito técnico institucional baseado na faixa de pontuação.
 */
function gerarVeredito_(score) {
  if (score >= 8.0) return "ALTA_CONFLUENCIA_ALTA";
  if (score >= 6.0) return "TENDENCIA_ALTA_FAVORAVEL";
  if (score <= 4.0) return "RISCO_OPERACIONAL_ELEVADO";
  return "STATUS_NEUTRO_AGUARDAR";
}

function getMapFromSheet_(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => h.toString().toUpperCase().trim().replace(/ /g, "_"));
  const rows = data.slice(1);
  const map = {};

  rows.forEach(row => {
    const ticker = row[0];
    if (!ticker) return;
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    map[ticker] = obj;
  });
  return map;
}