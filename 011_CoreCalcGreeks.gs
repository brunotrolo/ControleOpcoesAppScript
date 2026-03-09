/**
 * @fileoverview CoreCalcGreeks - v4.0 (Nativo - Black-Scholes Engine)
 * AÇÃO: Calcula Gregas e IV internamente via Black-Scholes (Newton-Raphson).
 * CORREÇÃO: Mapeado Código da Opção na coluna Ativo.
 * PADRÃO: Modo Silencioso e Contexto Serializado.
 */

const OptionMath = {
  DIAS_ANO: 252,
  T_MIN: 0.002, 

  pdf(x) { return Math.exp(-0.5 * x * x) / Math.sqrt(2 * Math.PI); },

  cdf(x) {
    const t = 1 / (1 + 0.2316419 * Math.abs(x));
    const d = 0.3989423 * Math.exp(-x * x / 2);
    const p = d * t * (0.3193815 + t * (-0.3565638 + t * (1.781478 + t * (-1.821256 + t * 1.330274))));
    return x > 0 ? 1 - p : p;
  },

  calculate(S, K, T, r, sigma, flag) {
    T = Math.max(T, this.T_MIN);
    const sqrtT = Math.sqrt(T);
    const d1 = (Math.log(S / K) + (r + 0.5 * sigma * sigma) * T) / (sigma * sqrtT);
    const d2 = d1 - (sigma * sqrtT);
    
    const nd1 = this.pdf(d1);
    const Nd1 = this.cdf(d1);
    const Nd2 = this.cdf(d2);
    const expRT = Math.exp(-r * T);

    const isCall = (flag.toLowerCase() === 'c' || flag.toLowerCase() === 'call');
    
    return {
      price: isCall ? (S * Nd1 - K * expRT * Nd2) : (K * expRT * this.cdf(-d2) - S * this.cdf(-d1)),
      delta: isCall ? Nd1 : Nd1 - 1,
      gamma: nd1 / (S * sigma * sqrtT),
      vega: (S * nd1 * sqrtT) / 100, 
      theta: (isCall ? 
              (-(S * nd1 * sigma) / (2 * sqrtT) - r * K * expRT * Nd2) :
              (-(S * nd1 * sigma) / (2 * sqrtT) + r * K * expRT * this.cdf(-d2))) / this.DIAS_ANO,
      rho: (isCall ? (K * T * expRT * Nd2) : (-K * T * expRT * this.cdf(-d2))) / 100,
      poe: isCall ? Nd2 : this.cdf(-d2)
    };
  },

  estimateIV(S, K, T, r, marketPrice, flag) {
    let sigma = 0.35; 
    for (let i = 0; i < 50; i++) {
      const g = this.calculate(S, K, T, r, sigma, flag);
      const diff = g.price - marketPrice;
      if (Math.abs(diff) < 0.0001) return sigma;
      const v = g.vega * 100; 
      if (v < 0.0001) break;
      sigma -= diff / v;
      if (sigma < 0.01) return 0.01;
      if (sigma > 5.0) return 5.0;
    }
    return sigma;
  },

  getMoneynessCode(S, K, flag) {
    const ratio = S / K;
    if (ratio >= 0.975 && ratio <= 1.025) return 'ATM';
    const isCall = (flag.toLowerCase() === 'c' || flag.toLowerCase() === 'call');
    if ((isCall && ratio > 1.025) || (!isCall && ratio < 0.975)) return 'ITM';
    return 'OTM';
  }
};

const GreeksCalculator = {
  _serviceName: "GreeksCalculator_v4.0",

  run() {
    const inicio = Date.now();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    const dataHoje = DataUtils.formatDateBR(new Date());
    const horaHoje = new Date().toLocaleTimeString('pt-BR');
    const timestampAtual = `${dataHoje} ${horaHoje}`;

    const cacheCalculos = {};
    const stats = { total: 0, calculos: 0, cache: 0, erros: 0 };

    SysLogger.log(this._serviceName, "START", ">>> INICIANDO CÁLCULO NATIVO (BLACK-SCHOLES) <<<", "");

    try {
      const abaGatilho = ss.getSheetByName(SYS_CONFIG.SHEETS.TRIGGER);
      const abaCalc = ss.getSheetByName("Calc_Greeks");
      const abaDetails = ss.getSheetByName(SYS_CONFIG.SHEETS.DETAILS);
      const abaAssets = ss.getSheetByName(SYS_CONFIG.SHEETS.ASSETS);

      if (!abaCalc) throw new Error("Aba 'Calc_Greeks' não encontrada.");

      const detailsMap = this._getMap(abaDetails, "ID_Trade_Unico");
      const assetsMap = this._getMap(abaAssets, "Ticker");
      const idToRowMap = this._getRowMap(abaCalc, "ID_Trade_Unico");

      const nectonData = abaGatilho.getDataRange().getValues();
      const headersN = nectonData[0];
      const idxID = headersN.indexOf("ID_Trade_Unico");
      const idxStatus = headersN.indexOf("Status Operação");

      const irate = 0.1075; // Selic

      for (let i = 1; i < nectonData.length; i++) {
        const idTrade = String(nectonData[i][idxID] || "").trim();
        const status = String(nectonData[i][idxStatus] || "").trim().toUpperCase();
        
        if (status !== "ATIVO" || idTrade.length < 10) continue;
        stats.total++;

        const detail = detailsMap[idTrade];
        const asset = detail ? assetsMap[detail.parent_symbol] : null;

        if (!detail || !asset) {
          stats.erros++;
          SysLogger.log(this._serviceName, "AVISO", `Insumos ausentes para ID: ${idTrade}`, `Procure por '${detail?.parent_symbol}' na aba de Ativos.`);
          continue;
        }

        const tickerOpcao = detail.symbol;
        let resBS;

        if (cacheCalculos[tickerOpcao]) {
          resBS = cacheCalculos[tickerOpcao];
          stats.cache++;
        } else {
          const S = Number(asset.close);
          const K = Number(detail.strike);
          const T_dias = Number(detail.days_to_maturity);
          const T_anos = T_dias / OptionMath.DIAS_ANO;
          const flag = detail.type.toLowerCase() === 'call' ? 'c' : 'p';
          const precoMercado = Number(detail.close);
          
          // Cálculo da IV Estimada (Onde a mágica acontece)
          const iv = OptionMath.estimateIV(S, K, T_anos, irate, precoMercado, flag);
          
          // Cálculo das Gregas
          resBS = OptionMath.calculate(S, K, T_anos, irate, iv, flag);
          resBS.volatility = iv;
          resBS.moneyness_code = OptionMath.getMoneynessCode(S, K, detail.type);
          resBS.moneyness_val = S / K;
          
          cacheCalculos[tickerOpcao] = resBS;
          stats.calculos++;

          // LOG DE AUDITORIA MATEMÁTICA (Relevante para conferência)
          const auditInfo = {
            spot: S,
            strike: K,
            dtm: T_dias,
            iv_calc: (iv * 100).toFixed(2) + "%",
            price_mkt: precoMercado,
            price_teo: resBS.price.toFixed(2),
            delta: resBS.delta.toFixed(4)
          };
          SysLogger.log(this._serviceName, "INFO", `Cálculo BS p/ ${tickerOpcao}`, JSON.stringify(auditInfo));
        }

        const dadosFinais = {
          ...resBS,
          ID_Trade_Unico: idTrade,
          Ativo: tickerOpcao,
          ID_Estrutura: detail.ID_Estrutura,
          Timestamp_Atualizacao: timestampAtual,
          moneyness: resBS.moneyness_val,
          moneyness_code: resBS.moneyness_code
        };

        const headersD = abaCalc.getRange(1, 1, 1, abaCalc.getLastColumn()).getValues()[0];
        const rowData = headersD.map(h => {
          const k = String(h).trim();
          return dadosFinais[k] !== undefined ? dadosFinais[k] : "";
        });

        const rowNum = idToRowMap[idTrade];
        if (rowNum) {
          abaCalc.getRange(rowNum, 1, 1, rowData.length).setValues([rowData]);
        } else {
          abaCalc.appendRow(rowData);
          idToRowMap[idTrade] = abaCalc.getLastRow();
        }
      }

      const duracao = ((Date.now() - inicio) / 1000).toFixed(1);
      SysLogger.log(this._serviceName, "FINISH", `>>> PROCESSAMENTO FINALIZADO <<<`, JSON.stringify({
        tempo: duracao + "s",
        ativos_analisados: stats.total,
        calculos_novos: stats.calculos,
        uso_cache: stats.cache,
        erros_insumo: stats.erros
      }));
      SysLogger.flush();

    } catch (e) {
      SysLogger.log(this._serviceName, "CRITICO", "Falha catastrófica no motor nativo", String(e.message));
      SysLogger.flush();
    }
  },

  _getMap(aba, pk) {
    if (!aba) return {};
    const data = aba.getDataRange().getValues();
    const headers = data[0];
    const pkIdx = headers.indexOf(pk);
    const map = {};
    for (let i = 1; i < data.length; i++) {
      const row = data[i], obj = {};
      headers.forEach((h, idx) => obj[h] = row[idx]);
      if (row[pkIdx]) map[String(row[pkIdx]).trim()] = obj;
    }
    return map;
  },

  _getRowMap(aba, pk) {
    const map = {};
    if (!aba || aba.getLastRow() < 2) return map;
    const data = aba.getDataRange().getValues();
    const pkIdx = data[0].indexOf(pk);
    for (let i = 1; i < data.length; i++) {
      if (data[i][pkIdx]) map[String(data[i][pkIdx]).trim()] = i + 1;
    }
    return map;
  }
};

// ============================================================================
// PONTO DE ENTRADA (Orquestrador / Menu)
// ============================================================================

/** Função para o Orquestrador rodar o cálculo nativo */
function calcularGregasNativo() { 
  GreeksCalculator.run(); 
}

// ============================================================================
// SUÍTE DE TESTES (011)
// ============================================================================

function testSuiteCalcGreeksInternal011() {
  console.log("=== INICIANDO AUDITORIA MATEMÁTICA: CALC GREEKS (011) ===");
  const tol = 0.001;
  
  // Teste 1: CDF
  const cdfZero = OptionMath.cdf(0);
  console.log(`[MATH] CDF(0): ${cdfZero} ${Math.abs(cdfZero - 0.5) < tol ? "✅" : "❌"}`);

  // Teste 2: Black-Scholes ATM Call
  const S = 100, K = 100, T = 1, r = 0.05, vol = 0.20;
  const res = OptionMath.calculate(S, K, T, r, vol, 'c');
  console.log(`[BS] Preço: ${res.price.toFixed(2)} (Esperado: ~10.45)`);

  console.log("--- Executando Carga Controlada ---");
  GreeksCalculator.run();

  console.log("=== FIM DA AUDITORIA ===");
}