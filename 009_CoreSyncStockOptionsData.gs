/**
 * @fileoverview CoreSyncStockOptionsData - v5.5 (Absolute Mapping & Timezone Fix)
 * AÇÃO: Sincroniza detalhes de opções com mapeamento rígido e formato BR.
 * CORREÇÃO: Impede que o 'updated_at' da API sobrescreva o timestamp BR do sistema.
 * PADRÃO: Dicionário Universal de Dados (v5.0).
 */

const OptionDetailsSync = {
  _serviceName: "OptionDetailsSync_v5.5",

  run() {
    const inicio = Date.now();
    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);

    const cacheAPI = {};
    const stats = { lidos: 0, processados: 0, skip_status: 0, api_calls: 0, erros: 0 };
    
    // FORMATO DE DATA COMBINADO: DD/MM/YYYY HH:MM:SS
    const dataBR = DataUtils.formatDateBR(new Date());
    const horaBR = new Date().toLocaleTimeString('pt-BR');
    const timestampSistema = `${dataBR} ${horaBR}`;

    SysLogger.log(this._serviceName, "START", ">>> INICIANDO SINCRONIZAÇÃO (FIX DATA/HORA) <<<", "");

    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const abaImport = ss.getSheetByName(SYS_CONFIG.SHEETS.IMPORT);
      const abaDetalhes = ss.getSheetByName(SYS_CONFIG.SHEETS.DETAILS);
      
      if (!abaImport || !abaDetalhes) throw new Error("Abas não encontradas.");
      
      const colI = this._getColMap(abaImport);
      const colD = this._getColMap(abaDetalhes);
      const idToRowMap = this._getIDRowMap(abaDetalhes, colD.ID_TRADE);
      
      const valoresImport = abaImport.getDataRange().getValues();

      // 1. MAPEAMENTO RÍGIDO (API Key -> Spreadsheet Label)
      const fieldMapper = {
        "symbol": "OPTION_TICKER",
        "parent_symbol": "TICKER",
        "name": "CONTRACT_DESC",
        "close": "CLOSE",
        "volume": "VOLUME_QTY",
        "financial_volume": "VOLUME_FIN",
        "trades": "TRADES_COUNT",
        "bid": "BID",
        "ask": "ASK",
        "due_date": "EXPIRY",
        "maturity_type": "MATURITY_TYPE",
        "contract_size": "LOT_SIZE",
        "exchange_id": "EXCHANGE_ID",
        "created_at": "CREATED_AT",
        "updated_at": "EDITED_AT", // Aqui redirecionamos o ISO da API para outra coluna
        "variation": "VARIATION",
        "spot_price": "SPOT",
        "isin": "ISIN",
        "security_category": "SECURITY_CAT",
        "market_maker": "MARKET_MAKER",
        "block_date": "BLOCK_DATE",
        "days_to_maturity": "DTE_CALENDAR",
        "cnpj": "CNPJ",
        "bid_volume": "BID_VOLUME",
        "ask_volume": "ASK_VOLUME",
        "time": "EXCH_TIMESTAMP",
        "type": "OPTION_TYPE",
        "last_trade_at": "LAST_TRADE_AT",
        "strike_eod": "STRIKE_EOD",
        "quotationForm": "QUOTATION_FORM",
        "lastUpdatedDividendsAt": "DIVIDEND_UPDATED_AT"
      };

      for (let i = 1; i < valoresImport.length; i++) {
        const linhaImport = valoresImport[i];
        const idTrade = String(linhaImport[colI.ID_TRADE] || "").trim();
        const optTicker = String(linhaImport[colI.OPTION_TICKER] || "").trim();
        const status = String(linhaImport[colI.STATUS_OP] || "").trim().toUpperCase();

        if (!idTrade || idTrade.length < 5) continue;
        stats.lidos++;

        if (status !== "ATIVO") { stats.skip_status++; continue; }

        let dadosAPI = cacheAPI[optTicker] || null;
        if (!dadosAPI) {
          dadosAPI = OplabService.getOptionDetails(optTicker);
          if (dadosAPI) { 
            cacheAPI[optTicker] = dadosAPI; 
            stats.api_calls++; 
            Utilities.sleep(1100); 
          }
        }

        if (dadosAPI) {
          const rowNum = idToRowMap[idTrade];
          let linhaFinal = rowNum ? abaDetalhes.getRange(rowNum, 1, 1, abaDetalhes.getLastColumn()).getValues()[0] : new Array(abaDetalhes.getLastColumn()).fill("");

          // 2. LOGICA DE EXTRAÇÃO COM PRIORIDADE DE SISTEMA
          for (const label in colD) {
            const idx = colD[label];
            let valorFinal;

            // A. Campos Controlados pelo Script (Prioridade Máxima)
            if (label === "UPDATED_AT") {
              valorFinal = timestampSistema; // SEMPRE BR
            } else if (label === "ID_TRADE") {
              valorFinal = idTrade;
            } else if (label === "ID_STRATEGY") {
              valorFinal = String(linhaImport[colI.ID_STRATEGY] || "").trim();
            } else {
              // B. Campos da API (via Mapper)
              const apiKey = Object.keys(fieldMapper).find(key => fieldMapper[key] === label);
              if (apiKey && dadosAPI[apiKey] !== undefined) {
                valorFinal = dadosAPI[apiKey];
              } else if (dadosAPI[label.toLowerCase()] !== undefined) {
                // C. Fallback para nomes idênticos em minúsculo (strike, bid, ask)
                valorFinal = dadosAPI[label.toLowerCase()];
              }
            }

            // Tratamento de tipos
            if (valorFinal !== undefined && valorFinal !== null) {
              if (label === "EXPIRY" && valorFinal.includes("-")) {
                linhaFinal[idx] = DataUtils.formatDateBR(valorFinal);
              } else {
                linhaFinal[idx] = typeof valorFinal === 'object' ? JSON.stringify(valorFinal) : valorFinal;
              }
            }
          }

          if (rowNum) {
            abaDetalhes.getRange(rowNum, 1, 1, linhaFinal.length).setValues([linhaFinal]);
          } else {
            abaDetalhes.appendRow(linhaFinal);
            idToRowMap[idTrade] = abaDetalhes.getLastRow();
          }
          stats.processados++;
        }
      }

      SysLogger.log(this._serviceName, "FINISH", ">>> SINCRONIA CONCLUÍDA <<<", JSON.stringify({
        tempo: ((Date.now() - inicio) / 1000).toFixed(1) + "s",
        total_lido: stats.lidos,
        atualizados: stats.processados,
        status_ignorado: stats.skip_status
      }));
      SysLogger.flush();

    } catch (e) {
      SysLogger.log(this._serviceName, "CRITICO", "Falha fatal no motor 009", String(e.message));
      SysLogger.flush();
    }
  },

  _getColMap(aba) {
    const headers = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0];
    const map = {};
    headers.forEach((h, i) => { if(h) map[String(h).trim()] = i; });
    return map;
  },

  _getIDRowMap(aba, colIdx) {
    const map = {};
    if (aba.getLastRow() < 2) return map;
    const ids = aba.getRange(2, colIdx + 1, aba.getLastRow() - 1, 1).getValues();
    ids.forEach((l, i) => { if (l[0]) map[String(l[0]).trim()] = i + 2; });
    return map;
  }
};

// ============================================================================
// PONTO DE ENTRADA (Trigger Dinâmico / Menu)
// ============================================================================

/**
 * Ponto de entrada para sincronizar detalhes (Gregas, Strikes, etc) das Opções ativas.
 */
function atualizarDetalhesOpcoes() {
  OptionDetailsSync.run();
}

// ============================================================================
// SUÍTE DE HOMOLOGAÇÃO (008)
// ============================================================================

function testSuiteOptionDetailsSync008() {
  console.log("=== INICIANDO HOMOLOGAÇÃO: OPTION DETAILS SYNC (008) ===");
  const tickerTeste = "PETRC425"; // Ajuste para um ticker válido de opção se necessário
  
  console.log(`--- Testando Fetch da API para ${tickerTeste} ---`);
  const dados = OplabService.getOptionDetails(tickerTeste);
  
  if (dados && dados.strike) {
    console.log(`✅ Dados da Opção recebidos. Strike: ${dados.strike}`);
    console.log(`   Data de Vencimento Original: ${dados.due_date}`);
  } else {
    console.error(`❌ Falha ao processar ${tickerTeste}. Talvez o ativo não exista mais ou a API falhou.`);
  }

  console.log("--- Executando Carga Controlada ---");
  OptionDetailsSync.run();

  console.log("=== TESTES CONCLUÍDOS ===");
}