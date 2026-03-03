/**
 * SERVIÇO 603: Processador de Fundamentos Preditivos
 * Responsabilidade: Calcular Preço Justo (Graham), Margem de Segurança e Saúde Financeira.
 * Dependências: 000_CoreServiceAPIClient, 002_CoreServiceLogger_DataExtractor
 */

const FUND_PRED_CONFIG = {
  servico: "603_ProcessFundamentalPredictive",
  abaOrigem: "Dados_Ativos", // 🌟 Alinhado com o 602
  abaDestino: "Analise_Fundamentalista_Ativos",
  selicReferencia: 0.1075 
};

/**
 * Executa o processamento fundamentalista utilizando dados REAIS
 */
function processFundamentalPredictive_Execute() {
  const t0 = Date.now();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetOrigem = ss.getSheetByName(FUND_PRED_CONFIG.abaOrigem);
  const sheetFund = ss.getSheetByName(FUND_PRED_CONFIG.abaDestino);
  
  if (!sheetOrigem || !sheetFund) {
    log(FUND_PRED_CONFIG.servico, "ERRO", "Abas de origem ou destino não encontradas", "");
    return;
  }

  try {
    // 1. Extração da lista de ativos via aba Dados_Ativos
    const dataAtivos = sheetOrigem.getDataRange().getValues();
    const headers = dataAtivos[0].map(h => h.toString().toUpperCase().trim());
    const rows = dataAtivos.slice(1);

    const col = {
      ticker: headers.indexOf("TICKER"),
      close: headers.indexOf("CLOSE")
    };

    const resultados = [];
    const agora = new Date();

  	rows.forEach(row => {
      const ticker = row[col.ticker];
      const spot = parseNumber_(row[col.close]);

      if (!ticker || isNaN(spot) || spot === 0) return;

      // 2. CAPTURA REAL via API BRAPI (Oficial)
      const dadosSI = getBrapiStockData(ticker);
      
      if (!dadosSI || dadosSI.lpa === 0) {
        log(FUND_PRED_CONFIG.servico, "AVISO", `Dados insuficientes para ${ticker}`, "Verificar BRAPI");
        return;
      }

      // 3. Cálculos de Valor (Fórmula de Graham)
      const lpa = dadosSI.lpa;
      const vpa = dadosSI.vpa;
      let precoGraham = 0;
      
      // Graham só funciona para empresas dando lucro e com patrimônio positivo
      if (lpa > 0 && vpa > 0) {
        precoGraham = Math.sqrt(22.5 * lpa * vpa);
      }

      const margemSeguranca = precoGraham > 0 ? (precoGraham - spot) / precoGraham : 0;

      // 4. Indicadores de Rentabilidade e Risco
      const earningsYield = spot > 0 ? lpa / spot : 0;
      const earningsVsSelic = earningsYield - FUND_PRED_CONFIG.selicReferencia;

      // 5. Score Fundamentalista (Escala 0 a 10)
      let scoreFund = 0;
      if (margemSeguranca > 0.20) scoreFund += 4;
      if (earningsVsSelic > 0) scoreFund += 3;   
      if (dadosSI.roe > 0.15) scoreFund += 3;    

      // 6. Montagem da Linha para a Planilha
      resultados.push([
        ticker, agora, lpa, vpa, precoGraham, margemSeguranca,
        dadosSI.pl, dadosSI.pvp, dadosSI.dy, earningsYield,
        earningsVsSelic, 0, dadosSI.roe, dadosSI.divida, scoreFund
      ]);

      // Log de memória de cálculo detalhada
      log(FUND_PRED_CONFIG.servico, "DEBUG", `Fundamentos: ${ticker}`, JSON.stringify({
        lpa: lpa,
        vpa: vpa,
        graham: precoGraham.toFixed(2),
        margem: (margemSeguranca * 100).toFixed(2) + "%"
      }));
    });

    // 7. Gravação em Lote
    if (resultados.length > 0) {
      sheetFund.getRange(2, 1, sheetFund.getLastRow(), resultados[0].length).clearContent();
      sheetFund.getRange(2, 1, resultados.length, resultados[0].length).setValues(resultados);
    }

    log(FUND_PRED_CONFIG.servico, "SUCESSO", `Processados ${resultados.length} ativos fundamentalistas`, 
        JSON.stringify({ tempo: (Date.now() - t0) + "ms" }));

  } catch (e) {
    log(FUND_PRED_CONFIG.servico, "ERRO_FATAL", "Falha no processador 603", e.toString());
  }
}

/**
 * Auxiliar para converter strings financeiras em números (Reutilizado do 602)
 */
function parseNumber_(val) {
  if (typeof val === 'number') return val;
  if (!val) return 0;
  let clean = val.toString().replace("R$ ", "").replace(/\./g, "").replace(",", ".");
  return parseFloat(clean) || 0;
}