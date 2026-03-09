/**
 * @fileoverview CoreSyncStockDataHistory - v4.0 (DUD Edition)
 * AÇÃO: Busca série histórica (250 dias) na OpLab e aplica Drop & Replace seguro.
 * PADRÃO: Modo Silencioso e Harmonização Universal de Cabeçalhos.
 */

const HistoricalDataSync = {
  _serviceName: "HistoricalDataSync_v4.0",
  _abaDestino: SYS_CONFIG.SHEETS.HIST_250D, // Puxa "DADOS_ATIVOS_HISTORICO250D" do 001
  _diasHistorico: 250,

  /**
   * Executa a extração em lote e substituição da base de dados históricos.
   */
  run() {
    const inicio = Date.now();
    
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const abaAtivos = ss.getSheetByName(SYS_CONFIG.SHEETS.ASSETS); // "DADOS_ATIVOS"
      
      if (!abaAtivos) throw new Error(`Aba origem (${SYS_CONFIG.SHEETS.ASSETS}) não encontrada.`);

      // 1. EXTRAÇÃO DINÂMICA DE TICKERS (Busca a coluna TICKER)
      const headersAtivos = abaAtivos.getRange(1, 1, 1, abaAtivos.getLastColumn()).getValues()[0];
      const colTickerIdx = headersAtivos.indexOf("TICKER");

      if (colTickerIdx === -1) throw new Error("Coluna 'TICKER' não encontrada em DADOS_ATIVOS.");

      const maxRows = abaAtivos.getLastRow();
      if (maxRows < 2) return;

      const tickersBrutos = abaAtivos.getRange(2, colTickerIdx + 1, maxRows - 1, 1).getValues().flat();
      const tickersAlvo = [...new Set(tickersBrutos.filter(t => t && String(t).trim() !== ""))];

      if (!tickersAlvo.includes("IBOV")) tickersAlvo.push("IBOV");

      SysLogger.log(this._serviceName, "INFO", `Fase 1: Mapeados ${tickersAlvo.length} ativos p/ histórico.`, JSON.stringify(tickersAlvo));

      // 2. BUSCA NA API
      const bufferDeDados = [];
      let contagemErro = 0;
      const timestampSync = DataUtils.formatDateBR(new Date()) + " " + new Date().toLocaleTimeString('pt-BR');

      tickersAlvo.forEach((ticker, index) => {
        const resAPI = OplabService.getHistoricalData(ticker, this._diasHistorico);
        
        if (resAPI && resAPI.data && resAPI.data.length > 0) {
          const symbolName = resAPI.symbol || ticker;
          const companyName = resAPI.name || "N/A";
          const resolution = resAPI.resolution || "1d";

          // Mapeia para as 12 colunas do novo DUD
          const rows = resAPI.data.map(item => [
            ticker,               // TICKER
            timestampSync,        // UPDATED_AT
            symbolName,           // SYMBOL_API
            companyName,          // COMPANY_NAME
            resolution,           // RESOLUTION
            item.time,            // CANDLE_TIME
            item.open,            // OPEN
            item.high,            // HIGH
            item.low,             // LOW
            item.close,           // SPOT (Antigo close)
            item.volume,          // VOLUME_QTY
            item.fvolume          // VOLUME_FIN
          ]);
          
          bufferDeDados.push(...rows);
        } else {
          contagemErro++;
          SysLogger.log(this._serviceName, "ERRO", `Falha no histórico de ${ticker}`);
        }
        if (index < tickersAlvo.length - 1) Utilities.sleep(800); 
      });

      if (bufferDeDados.length === 0) throw new Error("API retornou vazio. Abortando Drop & Replace.");

      // 3. DROP & REPLACE (Gravação com Novos Cabeçalhos)
      let abaDestino = ss.getSheetByName(this._abaDestino);
      if (!abaDestino) abaDestino = ss.insertSheet(this._abaDestino);

      // Cabeçalhos Oficiais DUD v5.0
      const headersDUD = [
        "TICKER", "UPDATED_AT", "SYMBOL_API", "COMPANY_NAME", "RESOLUTION", 
        "CANDLE_TIME", "OPEN", "HIGH", "LOW", "SPOT", "VOLUME_QTY", "VOLUME_FIN"
      ];

      abaDestino.clearContents(); 
      abaDestino.getRange(1, 1, 1, headersDUD.length).setValues([headersDUD]);
      abaDestino.getRange(2, 1, bufferDeDados.length, headersDUD.length).setValues(bufferDeDados);

      const duracao = ((Date.now() - inicio) / 1000).toFixed(1);
      SysLogger.log(this._serviceName, "FINISH", `>>> HISTÓRICO ATUALIZADO EM ${duracao}s <<<`, JSON.stringify({
        linhas: bufferDeDados.length,
        erros: contagemErro
      }));
      SysLogger.flush();

    } catch (e) {
      SysLogger.log(this._serviceName, "CRITICO", "Erro no Sync de Histórico", String(e.message));
      SysLogger.flush();
    }
  }
};

// ============================================================================
// PONTO DE ENTRADA (Trigger Dinâmico do Orquestrador)
// ============================================================================

/**
 * Ponto de entrada para varrer os ativos atuais e atualizar a série de 250 dias.
 */
function atualizarDadosHistoricos() {
  HistoricalDataSync.run();
}

// ============================================================================
// SUÍTE DE HOMOLOGAÇÃO 101% (009)
// ============================================================================

function testSuiteHistoricalSync009() {
  console.log("=== INICIANDO TESTE: HISTORICAL SYNC (009) ===");
  const tickerTeste = "IBOV"; 
  
  console.log(`--- Testando Conexão API para Histórico de ${tickerTeste} (5 dias) ---`);
  const t0 = Date.now();
  const resAPI = OplabService.getHistoricalData(tickerTeste, 5); 
  const t1 = Date.now();
  
  if (resAPI && resAPI.data && resAPI.data.length > 0) {
    console.log(`✅ [API] Retornou histórico em ${t1 - t0}ms.`);
    console.log(`   Último Candle (Data): ${resAPI.data[resAPI.data.length-1].time}`);
    console.log(`   Último Fechamento: ${resAPI.data[resAPI.data.length-1].close}`);
  } else {
    console.error(`❌ [API] Falha ao extrair histórico de ${tickerTeste}.`);
  }
  
  console.log("--- Executando Carga Total Controlada ---");
  HistoricalDataSync.run();
  
  console.log("=== TESTES CONCLUÍDOS ===");
}