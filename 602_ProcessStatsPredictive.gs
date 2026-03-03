/**
 * SERVIÇO 602: Processador de Estatística Preditiva
 * Responsabilidade: Calcular Volatilidade Histórica, Sigmas e Probabilidades dinâmicas.
 * Dependências: 002_CoreServiceLogger_DataExtractor
 */

const STATS_PRED_CONFIG = {
  servico: "602_ProcessStatsPredictive",
  abaOrigem: "Dados_Ativos",
  abaHistorico: "Dados_Ativos_Historico250d",
  abaDestino: "Analise_Estatistica_Ativos",
  prazoX: 45 // Valor padrão (Fallback)
};

/**
 * Executa o processamento estatístico para todos os ativos monitorados
 * @param {number} diasParam - (Opcional) Horizonte de tempo em dias enviado pelo Front-end
 */
function processStatsPredictive_Execute(diasParam) {
  const t0 = Date.now();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetOrigem = ss.getSheetByName(STATS_PRED_CONFIG.abaOrigem);
  const sheetEstatistica = ss.getSheetByName(STATS_PRED_CONFIG.abaDestino);
  
  // Define o prazo: Usa o parâmetro do front-end ou o padrão da configuração
  const diasProjecao = diasParam || STATS_PRED_CONFIG.prazoX;
  
  if (!sheetOrigem || !sheetEstatistica) {
    log(STATS_PRED_CONFIG.servico, "ERRO", "Abas necessárias não encontradas", "");
    return;
  }

  try {
    // 1. Extração manual da aba Dados_Ativos (Garante que pegamos os 10 ativos monitorados)
    const dataAtivos = sheetOrigem.getDataRange().getValues();
    const headersAtivos = dataAtivos[0].map(h => h.toString().toUpperCase().trim());
    const rowsAtivos = dataAtivos.slice(1);

    // Mapeamento dinâmico de colunas
    const col = {
      ticker: headersAtivos.indexOf("TICKER"),
      close: headersAtivos.indexOf("CLOSE"),
      iv: headersAtivos.indexOf("IV_CURRENT"),
      ivRank: headersAtivos.indexOf("IV_1Y_RANK")
    };

    const historicoData = ss.getSheetByName(STATS_PRED_CONFIG.abaHistorico).getDataRange().getValues();
    const resultados = [];
    const agora = new Date();

    rowsAtivos.forEach(row => {
      const ticker = row[col.ticker];
      const spot = parseNumber_(row[col.close]);
      const ivAtual = parseNumber_(row[col.iv]) / 100;
      const ivRank = row[col.ivRank];

      if (!ticker || isNaN(spot) || spot === 0 || isNaN(ivAtual)) return;

      // 2. Cálculo da Volatilidade Histórica (HV 250d)
      const hv250 = calculateHV_(ticker, historicoData);
      const hvIvRatio = hv250 > 0 ? ivAtual / hv250 : 0;

      // 3. Matemática de Sigmas (Ajuste Temporal Dinâmico para 'diasProjecao')
      const sigmaDiario = ivAtual / Math.sqrt(252);
      const sigmaPeriodo = sigmaDiario * Math.sqrt(diasProjecao);

      // 4. Canais de Probabilidade (Estes são os seus "Spots Projetados")
      const prob1Sup = spot * (1 + sigmaPeriodo); // Projeção Máxima (68%)
      const prob1Inf = spot * (1 - sigmaPeriodo); // Projeção Mínima (68%)
      const prob2Inf = spot * (1 - (sigmaPeriodo * 2));

      // 5. Z-Score e Probabilidade
      const probPermanecer = 0.6827; 
      const zScore = (spot > 0) ? (spot - (spot * (1 - sigmaDiario))) / (spot * sigmaDiario) : 0;

      // Mantemos a mesma estrutura de colunas para não quebrar a sua planilha
      resultados.push([
        ticker, agora, spot, hv250, ivAtual, ivRank, hvIvRatio, 
        ivAtual, sigmaDiario, sigmaPeriodo, prob1Sup, prob1Inf, prob2Inf, zScore, probPermanecer
      ]);

      // Memória de Cálculo no Log informando os dias
      log(STATS_PRED_CONFIG.servico, "DEBUG", `Cálculo: ${ticker} (${diasProjecao} dias)`, JSON.stringify({
        ticker: ticker,
        spot: spot,
        iv: (ivAtual * 100).toFixed(2) + "%",
        sigmaPeriodo: (sigmaPeriodo * 100).toFixed(2) + "%",
        spot_projetado_max: prob1Sup.toFixed(2),
        spot_projetado_min: prob1Inf.toFixed(2)
      }));
    });

    // 6. Gravação Final
    if (resultados.length > 0) {
      sheetEstatistica.getRange(2, 1, sheetEstatistica.getLastRow(), resultados[0].length).clearContent();
      sheetEstatistica.getRange(2, 1, resultados.length, resultados[0].length).setValues(resultados);
    }

    log(STATS_PRED_CONFIG.servico, "SUCESSO", `Processamento de ${resultados.length} ativos para horizonte de ${diasProjecao} dias concluído`, 
        JSON.stringify({ tempo: (Date.now() - t0) + "ms" }));

  } catch (e) {
    log(STATS_PRED_CONFIG.servico, "ERRO_FATAL", "Falha no motor estatístico", e.toString());
  }
}

/**
 * Auxiliar para converter strings financeiras em números reais
 */
function parseNumber_(val) {
  if (typeof val === 'number') return val;
  if (!val) return 0;
  let clean = val.toString().replace("R$ ", "").replace(/\./g, "").replace(",", ".");
  return parseFloat(clean) || 0;
}

/**
 * Calcula a Volatilidade Histórica Anualizada
 */
function calculateHV_(ticker, data) {
  const fechamentos = data
    .filter(row => row[0] === ticker)
    .map(row => parseNumber_(row[9])) // Coluna 'close' no histórico (ajuste se mudar)
    .filter(val => !isNaN(val) && val > 0);

  if (fechamentos.length < 2) return 0;

  const retornosLog = [];
  for (let i = 1; i < fechamentos.length; i++) {
    retornosLog.push(Math.log(fechamentos[i] / fechamentos[i - 1]));
  }

  const n = retornosLog.length;
  const media = retornosLog.reduce((a, b) => a + b, 0) / n;
  const variancia = retornosLog.reduce((a, b) => a + Math.pow(b - media, 2), 0) / (n - 1);
  const desvioPadrao = Math.sqrt(variancia);

  return desvioPadrao * Math.sqrt(252); 
}