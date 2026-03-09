/**
 * @fileoverview CoreSyncGreeks - v5.4 (Risk Engine - Absolute Mapping & Log Master)
 * AÇÃO: Calcula e sincroniza Gregas (BS) via API com mapeamento blindado.
 * CORREÇÃO: Fim das lacunas em colunas (Zero Holes) e Logs de Auditoria Master.
 * PADRÃO: Dicionário Universal de Dados (v5.0).
 */

const GreeksSync = {
  _serviceName: "GreeksSync_v5.4",

  run() {
    const inicio = Date.now();
    const cacheBS = {}; 
    const stats = { lidos: 0, ativos: 0, gravados: 0, skip_status: 0, erros: 0 };
    const statusEncontrados = {};
    
    // FORMATO DE DATA EXIGIDO: DD/MM/YYYY HH:MM:SS
    const dataBR = DataUtils.formatDateBR(new Date());
    const horaBR = new Date().toLocaleTimeString('pt-BR');
    const timestampSistema = `${dataBR} ${horaBR}`;

    SysLogger.log(this._serviceName, "START", ">>> INICIANDO MOTOR DE GREGAS (v5.4) <<<", "");

    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const abaImport   = ss.getSheetByName(SYS_CONFIG.SHEETS.IMPORT);
      const abaGreeks   = ss.getSheetByName(SYS_CONFIG.SHEETS.GREEKS_API);
      const abaDetails  = ss.getSheetByName(SYS_CONFIG.SHEETS.DETAILS);
      const abaAssets   = ss.getSheetByName(SYS_CONFIG.SHEETS.ASSETS);
      
      if (!abaImport || !abaGreeks) throw new Error("Abas críticas (IMPORT ou GREEKS) não encontradas.");

      // 1. CARREGAR MAPAS DE SUPORTE (Details por ID_TRADE e Assets por TICKER)
      const detailsMap = this._getDynamicMap(abaDetails, "ID_TRADE");
      const assetsMap  = this._getDynamicMap(abaAssets, "TICKER");

      // 2. SCAN DINÂMICO DE CABEÇALHOS
      const colI = this._getColMap(abaImport);
      const colG = this._getColMap(abaGreeks);

      // Mapa de ID_TRADE no destino para realizar UPDATE cirúrgico
      const idToRowMap = {};
      if (abaGreeks.getLastRow() > 1) {
        const ids = abaGreeks.getRange(2, colG.ID_TRADE + 1, abaGreeks.getLastRow() - 1, 1).getValues();
        ids.forEach((l, i) => { if (l[0]) idToRowMap[String(l[0]).trim()] = i + 2; });
      }

      const valoresImport = abaImport.getDataRange().getValues();

      // 3. LOOP DE PROCESSAMENTO
      for (let i = 1; i < valoresImport.length; i++) {
        const linha = valoresImport[i];
        const idTrade  = String(linha[colI.ID_TRADE] || "").trim();
        const optTicker = String(linha[colI.OPTION_TICKER] || "").trim();
        const statusRaw = String(linha[colI.STATUS_OP] || "").trim();
        const statusUpper = statusRaw.toUpperCase();

        if (!idTrade || idTrade.length < 5) continue;
        
        stats.lidos++;
        statusEncontrados[statusUpper] = (statusEncontrados[statusUpper] || 0) + 1;

        // FILTRO DE STATUS: Rigorosamente ATIVO
        if (statusUpper !== "ATIVO") { 
          stats.skip_status++; 
          continue; 
        }
        stats.ativos++;

        const detail = detailsMap[idTrade];
        const asset  = detail ? assetsMap[detail.TICKER] : null;

        if (!detail || !asset) {
          stats.erros++;
          // AVISO: Se faltar detalhe ou preço da ação, o cálculo de Gregas é impossível
          continue; 
        }

        // 4. CÁLCULO OU CACHE (BLACK-SCHOLES)
        let bsResult = null;
        if (cacheBS[optTicker]) {
          bsResult = cacheBS[optTicker];
          stats.cache++;
        } else {
          const params = {
            symbol: optTicker,
            irate: 10.75, // Selic base
            type: detail.OPTION_TYPE,
            spotprice: Number(asset.SPOT),
            strike: Number(detail.STRIKE),
            dtm: Number(detail.DTE_CALENDAR),
            vol: Number(asset.IV),
            amount: Math.abs(Number(linha[colI.QUANTITY] || 0))
          };

          bsResult = OplabService.calculateBS(params);
          if (bsResult) {
            cacheBS[optTicker] = bsResult;
            Utilities.sleep(850); // Proteção Rate Limit
          }
        }

        // 5. MAPEAMENTO ABSOLUTO (ZERO HOLES)
        if (bsResult) {
          const rowNum = idToRowMap[idTrade];
          // Se for update, carrega a linha atual para preservar colunas manuais (Ex: MARGIN)
          let linhaFinal = rowNum ? abaGreeks.getRange(rowNum, 1, 1, abaGreeks.getLastColumn()).getValues()[0] : new Array(abaGreeks.getLastColumn()).fill("");

          // Objeto de dados normalizado com as chaves exatas da Planilha (DUD)
          const dadosMapeados = {
            ID_TRADE: idTrade,
            OPTION_TICKER: optTicker,
            ID_STRATEGY: String(linha[colI.ID_STRATEGY] || "").trim(),
            UPDATED_AT: timestampSistema,
            DELTA: bsResult.delta,
            GAMMA: bsResult.gamma,
            VEGA: bsResult.vega,
            THETA: bsResult.theta,
            RHO: bsResult.rho,
            POE: bsResult.poe,
            PRICE: bsResult.price,
            IV_CALC: bsResult.volatility,
            MONEYNESS: bsResult.moneyness_code || bsResult.moneyness,
            MONEYNESS_RATIO: bsResult.moneyness_ratio || (Number(asset.SPOT) / Number(detail.STRIKE)),
            SPOT: Number(asset.SPOT),
            STRIKE: Number(detail.STRIKE)
          };

          // Injeção cirúrgica na linha: rótulo da coluna -> valor do objeto
          for (const label in colG) {
            const idx = colG[label];
            if (dadosMapeados[label] !== undefined) {
              linhaFinal[idx] = dadosMapeados[label];
            }
          }

          // Gravação Física
          if (rowNum) {
            abaGreeks.getRange(rowNum, 1, 1, linhaFinal.length).setValues([linhaFinal]);
          } else {
            abaGreeks.appendRow(linhaFinal);
            idToRowMap[idTrade] = abaGreeks.getLastRow();
          }
          stats.gravados++;
          SysLogger.log(this._serviceName, "SUCESSO", `Gregas de ${optTicker} OK`, `ID: ${idTrade}`);
        }
      }

      // 6. LOG DE FINALIZAÇÃO (EXATAMENTE COMO SOLICITADO)
      const duracaoFinal = ((Date.now() - inicio) / 1000).toFixed(1);
      const payloadLog = {
        duracao: duracaoFinal + "s",
        total_lido: stats.lidos,
        gravados: stats.gravados,
        pulados_status: stats.skip_status,
        diagnostico: statusEncontrados
      };

      SysLogger.log(this._serviceName, "FINISH", `>>> GREGAS ATUALIZADAS EM ${duracaoFinal}s <<<`, JSON.stringify(payloadLog));
      SysLogger.flush();

    } catch (e) {
      SysLogger.log(this._serviceName, "CRITICO", "Falha no motor 010", String(e.message));
      SysLogger.flush();
    }
  },

  /** Helper: Mapeia rótulos de uma aba para índices (0-based) */
  _getColMap(aba) {
    if (!aba) return {};
    const headers = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0];
    const map = {};
    headers.forEach((h, i) => { if(h) map[String(h).trim()] = i; });
    return map;
  },

  /** Helper: Cria mapa de dados de suporte por Chave Primária (Ex: TICKER ou ID_TRADE) */
  _getDynamicMap(aba, pkLabel) {
    if (!aba) return {};
    const data = aba.getDataRange().getValues();
    const headers = data[0];
    const pkIdx = headers.indexOf(pkLabel);
    if (pkIdx === -1) return {};
    const map = {};
    for (let i = 1; i < data.length; i++) {
      const obj = {};
      headers.forEach((h, idx) => obj[h] = data[i][idx]);
      if (data[i][pkIdx]) map[String(data[i][pkIdx]).trim()] = obj;
    }
    return map;
  }
};


// ============================================================================
// PONTO DE ENTRADA (Trigger Dinâmico / Menu)
// ============================================================================

/** Ponto de Entrada para o Motor de Risco */
function atualizarGregas() {
  GreeksSync.run();
}

// ============================================================================
// SUÍTE DE HOMOLOGAÇÃO (010)
// ============================================================================

/**
 * Testa o cálculo individual de Black-Scholes via API OpLab.
 */
function testSuiteGreeksSync010() {
  console.log("=== INICIANDO HOMOLOGAÇÃO: GREEKS SYNC (010) ===");
  
  const paramsTeste = {
    symbol: "PETRR315",
    irate: 10.75,      // Taxa Selic (em %)
    type: "PUT",       // Tipo da opção
    spotprice: 40.69,  // Preço atual do ativo objeto
    strike: 30.73,     // Preço de exercício
    dtm: 71,           // Dias úteis para o vencimento
    vol: 35.5,         // Volatilidade Implícita (em %)
    amount: 1000       // Quantidade da posição
  };

  console.log(`🚀 Solicitando cálculo Black-Scholes para ${paramsTeste.symbol}...`);
  
  const t0 = Date.now();
  const resultado = OplabService.calculateBS(paramsTeste);
  const t1 = Date.now();

  if (resultado && resultado.delta !== undefined) {
    console.log(`✅ SUCESSO: Resposta recebida em ${t1 - t0}ms.`);
    console.log(`📐 Delta: ${resultado.delta} (Exposição direcional)`);
    console.log(`📐 Gamma: ${resultado.gamma} (Aceleração do Delta)`);
    console.log(`📐 Theta: ${resultado.theta} (Decaimento temporal)`);
    console.log(`💰 Preço Teórico: R$ ${resultado.price}`);
  } else {
    console.error("❌ FALHA: A API não retornou cálculos válidos.");
  }
  
  console.log("--- Executando Carga Controlada ---");
  GreeksSync.run();

  console.log("=== TESTES CONCLUÍDOS ===");
}