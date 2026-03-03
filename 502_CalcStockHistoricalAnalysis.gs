/**
 * ═══════════════════════════════════════════════════════════════
 * CALC TENDÊNCIA DADOS ATIVOS - MOTOR TÉCNICO V3 (FULL LOGGING)
 * ═══════════════════════════════════════════════════════════════
 * RESPONSABILIDADES:
 * - Processar dump histórico de 250 dias.
 * - Calcular Médias (5, 20, 200), IFR14 e Bandas de Bollinger.
 * - Calcular Volume Relativo e Distância de Médias.
 * - Gerar vereditos técnicos para estratégia de Venda de PUT.
 * ═══════════════════════════════════════════════════════════════
 */

function calcularTendenciaMercado() {
  const SERVICO_NOME = "CalcTendencia_v3";
  log(SERVICO_NOME, "INFO", "Iniciando motor de análise técnica...", "Aba Fonte: Dados_Ativos_Historico250d");

  try {
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const abaFonte = planilha.getSheetByName("Dados_Ativos_Historico250d");
    const abaDestino = planilha.getSheetByName("Dados_Ativos_Historico_Tendencia");

    // Validação de infraestrutura
    if (!abaFonte || !abaDestino) {
      const erroAbas = "Abas necessárias não encontradas. Verifique se Dados_Ativos_Historico250d existe.";
      log(SERVICO_NOME, "ERRO_CRITICO", erroAbas, "");
      return;
    }

    const dadosBrutos = abaFonte.getDataRange().getValues();
    if (dadosBrutos.length <= 1) {
      log(SERVICO_NOME, "AVISO", "Aba fonte está vazia (apenas cabeçalho).", "Sincronize os dados históricos primeiro.");
      return;
    }

    // 1. Agrupar dados por Ticker (Otimização de Memória)
    log(SERVICO_NOME, "DEBUG", "Agrupando linhas por ticker...", `Total de linhas brutas: ${dadosBrutos.length}`);
    const historicoPorTicker = agruparPorTicker(dadosBrutos);
    const tickersEncontrados = Object.keys(historicoPorTicker);
    log(SERVICO_NOME, "INFO", `Iniciando cálculos para ${tickersEncontrados.length} ativos.`, "");

    const resultadosFinais = [];
    let sucessos = 0;
    let ignorados = 0;

    // 2. Processamento Individual por Ativo
    for (const ticker of tickersEncontrados) {
      const h = historicoPorTicker[ticker];
      const s = h.precos; // Array de fechamentos
      const v = h.volumes; // Array de volumes
      
      // Validação de profundidade histórica para MMA200
      if (s.length < 200) {
        log(SERVICO_NOME, "ALERTA", `Histórico insuficiente para ${ticker}`, `Encontrado: ${s.length} dias. Mínimo exigido: 200.`);
        ignorados++;
        continue;
      }

      try {
        // --- CÁLCULOS TÉCNICOS ---
        const close = s[s.length - 1];
        const mma5 = calcularMMA(s, 5);
        const mma20 = calcularMMA(s, 20);
        const mma200 = calcularMMA(s, 200);
        const ifr14 = calcularIFR(s, 14);
        
        // Bandas de Bollinger (Configuração Standard: 20p, 2 Desvios)
        const desvio = calcularDesvioPadrao(s, 20);
        const bSuperior = Number((mma20 + (desvio * 2)).toFixed(2));
        const bInferior = Number((mma20 - (desvio * 2)).toFixed(2));
        const larguraBanda = Number(((bSuperior - bInferior) / mma20 * 100).toFixed(2));

        // Volume Relativo (Volume de hoje vs Média Volume 20p)
        const volMedio20 = calcularMMA(v, 20);
        const volRelativo = volMedio20 > 0 ? Number((v[v.length - 1] / volMedio20).toFixed(2)) : 1;

        // --- LÓGICA DE TENDÊNCIA E STATUS ---
        const dist200 = Number(((close / mma200 - 1) * 100).toFixed(2));
        
        // Veredito Base: Preço e Médias
        let tendencia = "LATERAL";
        if (close > mma200 && mma20 > mma200) tendencia = "ALTA";
        else if (close < mma200 || mma20 < mma200) tendencia = "BAIXA";

        // Status IFR (RSI)
        let statusIFR = "NEUTRO";
        if (ifr14 < 30) statusIFR = "SOBREVENDIDO";
        else if (ifr14 > 70) statusIFR = "SOBRECOMPRADO";

        // Status Bollinger
        let statusBollinger = "DENTRO";
        if (close <= bInferior) statusBollinger = "FURA BANDA INF";
        else if (close >= bSuperior) statusBollinger = "FURA BANDA SUP";

        // 🌟 NOVO: Calcula Beta e Correlação se não for o próprio IBOV
        let beta = "";
        let correl = "";
        if (ticker !== "IBOV" && historicoPorTicker["IBOV"]) {
           const ibovHistory = historicoPorTicker["IBOV"].precos;
           const metricasEstat = calcularBetaCorrelacao_(s, ibovHistory, 60);
           beta = metricasEstat.beta;
           correl = metricasEstat.correl;
        }

        log(SERVICO_NOME, "DEBUG", `Processado: ${ticker}`, `Close: ${close} | Beta: ${beta} | Correl: ${correl}`);

        const info = h.lastInfo;
        
        // 3. Montar Array Final (Agora com 26 Colunas)
        // Ignora o IBOV na gravação final para não sujar a aba de ativos
        if (ticker !== "IBOV") {
          resultadosFinais.push([
            ticker, new Date(), info.symbol, info.name, info.resolution,
            info.time, info.open, info.high, info.low, close, 
            info.volume, info.fvolume, mma5, mma20, mma200, 
            ifr14, dist200, bSuperior, bInferior, larguraBanda, 
            tendencia, statusIFR, statusBollinger, volRelativo,
            beta, correl // 🌟 NOVO
          ]);
          sucessos++;
        }

      } catch (e) {
        log(SERVICO_NOME, "ERRO_TICKER", `Erro ao calcular indicadores de ${ticker}`, e.message);
      }
    }

    // 4. Persistência de Dados
    if (resultadosFinais.length > 0) {
      abaDestino.clearContents();
      const headers = [
        "ticker", "timestamp_sync", "symbol", "name", "resolution", "time", "open", "high", "low", "close", 
        "volume", "fvolume", "MMA5", "MMA20", "MMA200", "IFR14", "Dist_MMA200", "Banda_Sup", "Banda_Inf", 
        "Largura_Banda", "Veredito_Tendencia", "Status_IFR", "Status_Bollinger", "Vol_Relativo",
        "Beta_Ibov_60d", "Correl_Ibov_60d" // 🌟 NOVO
      ];
      
      abaDestino.getRange(1, 1, 1, headers.length).setValues([headers]);
      abaDestino.getRange(2, 1, resultadosFinais.length, headers.length).setValues(resultadosFinais);
      
      log(SERVICO_NOME, "SUCESSO", `Motor finalizado com sucesso.`, `Processados: ${sucessos} | Ignorados: ${ignorados}`);
    } else {
      log(SERVICO_NOME, "AVISO", "Nenhum resultado gerado após os cálculos.", "Verifique se os ativos possuem histórico suficiente.");
    }

  } catch (error) {
    log(SERVICO_NOME, "ERRO_CRITICO", "Falha catastrófica no motor de cálculo", error.stack);
  }
}

// --- BIBLIOTECA MATEMÁTICA ---

function calcularMMA(arr, p) {
  const slice = arr.slice(-p);
  const soma = slice.reduce((a, b) => a + b, 0);
  return Number((soma / p).toFixed(2));
}

function calcularDesvioPadrao(arr, p) {
  const slice = arr.slice(-p);
  const media = slice.reduce((a, b) => a + b, 0) / p;
  const variancia = slice.reduce((a, b) => a + Math.pow(b - media, 2), 0) / p;
  return Math.sqrt(variancia);
}

function calcularIFR(s, p) {
  if (s.length < p + 1) return 50;
  let ganhos = 0, perdas = 0;
  for (let i = s.length - p; i < s.length; i++) {
    const d = s[i] - s[i-1];
    d > 0 ? ganhos += d : perdas -= d;
  }
  if (perdas === 0) return 100;
  const rs = ganhos / perdas;
  return Number((100 - (100 / (1 + rs))).toFixed(2));
}

function agruparPorTicker(matriz) {
  const mapa = {};
  for (let i = 1; i < matriz.length; i++) {
    const t = matriz[i][0]; // Coluna A: ticker
    if (!mapa[t]) mapa[t] = { precos: [], volumes: [] };
    
    mapa[t].precos.push(matriz[i][9]); // Coluna J: close
    mapa[t].volumes.push(matriz[i][10]); // Coluna K: volume
    
    // Armazena metadados da última linha disponível (o "hoje")
    mapa[t].lastInfo = {
      symbol: matriz[i][2], name: matriz[i][3], resolution: matriz[i][4],
      time: matriz[i][5], open: matriz[i][6], high: matriz[i][7], low: matriz[i][8],
      volume: matriz[i][10], fvolume: matriz[i][11]
    };
  }
  return mapa;
}


/**
 * 🌟 NOVA FUNÇÃO: Calcula Beta e Correlação com base em dois vetores de preço
 */
function calcularBetaCorrelacao_(arrAtivo, arrIbov, periodos) {
  const pA = arrAtivo.slice(-periodos);
  const pI = arrIbov.slice(-periodos);
  
  if (pA.length < 2 || pI.length < 2 || pA.length !== pI.length) {
    return { beta: 1, correl: 1 };
  }

  const retA = [];
  const retI = [];
  
  for (let i = 1; i < pA.length; i++) {
    if (pA[i-1] > 0 && pI[i-1] > 0) {
      retA.push(Math.log(pA[i] / pA[i-1]));
      retI.push(Math.log(pI[i] / pI[i-1]));
    }
  }

  const meanA = retA.reduce((a, b) => a + b, 0) / retA.length;
  const meanI = retI.reduce((a, b) => a + b, 0) / retI.length;
  
  let cov = 0, varI = 0, varA = 0;
  
  for (let i = 0; i < retA.length; i++) {
    const diffA = retA[i] - meanA;
    const diffI = retI[i] - meanI;
    cov += (diffA * diffI);
    varI += Math.pow(diffI, 2);
    varA += Math.pow(diffA, 2);
  }
  
  const beta = varI > 0 ? (cov / varI) : 1;
  const correl = (varI > 0 && varA > 0) ? (cov / Math.sqrt(varI * varA)) : 1;

  return { 
    beta: Number(beta.toFixed(2)), 
    correl: Number(correl.toFixed(2)) 
  };
}