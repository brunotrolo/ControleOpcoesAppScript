/**
 * ═══════════════════════════════════════════════════════════════
 * Sync_Ativos_Historicos
 * ═══════════════════════════════════════════════════════════════
 * RESPONSABILIDADE: Buscar série histórica (250 dias) na OpLab
 * ENDPOINT: /market/historical/{symbol}/1d
 * ABA DESTINO: Dados_Ativos_Historico250d
 * ═══════════════════════════════════════════════════════════════
 */

const ABA_HISTORICO_250D = "Dados_Ativos_Historico250d";
const OPLAB_BASE_URL_V3 = "https://api.oplab.com.br/v3";

/**
 * Função principal para sincronizar o histórico de 250 dias
 */
function sincronizarAtivosHistoricos() {
  const SERVICO_NOME = "SyncDadosAtivosHistoricos_v1";
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    // 1. Obter Token e Tickers (Reaproveitando a lógica do seu sistema)
    const token = PropertiesService.getScriptProperties().getProperty("OPLAB_ACCESS_TOKEN");
    if (!token) throw new Error("Token OPLAB não configurado.");

    const abaAtivos = planilha.getSheetByName("Dados_Ativos");
    if (!abaAtivos) throw new Error("Aba Dados_Ativos não encontrada.");
    
    // Extrai tickers únicos da Coluna A (Ticker)
    const tickers = abaAtivos.getRange(2, 1, abaAtivos.getLastRow() - 1, 1).getValues()
                             .flat()
                             .filter(t => t !== "" && t !== "Ticker");
    
    // 🌟 NOVO: Garante que o IBOV seja baixado para servir de base para Beta e Correlação
    if (!tickers.includes("IBOV")) {
      tickers.push("IBOV");
    }

    if (tickers.length === 0) return Logger.log("Nenhum ticker para processar.");

    // 2. Preparar Aba de Destino
    let abaDestino = planilha.getSheetByName(ABA_HISTORICO_250D);
    if (!abaDestino) {
      abaDestino = planilha.insertSheet(ABA_HISTORICO_250D);
    }
    
    // Define Headers Exatos conforme discutido (12 colunas)
    const headers = [
      "ticker", "timestamp_sync", "symbol", "name", "resolution", 
      "time", "open", "high", "low", "close", "volume", "fvolume"
    ];
    abaDestino.clearContents(); 
    abaDestino.getRange(1, 1, 1, headers.length).setValues([headers]);

    let bufferDeDados = [];
    const agora = new Date();

    // 3. Loop de busca por Ticker
    for (let i = 0; i < tickers.length; i++) {
      const tickerAlvo = tickers[i];
      Logger.log(`[${SERVICO_NOME}] Baixando: ${tickerAlvo} (${i+1}/${tickers.length})`);

      const resAPI = fetchHistoricoOpLab(tickerAlvo, token);
      
      if (resAPI && resAPI.data && resAPI.data.length > 0) {
        // Mapeia o JSON para o array de colunas da planilha
        const rows = resAPI.data.map(item => [
          tickerAlvo,           // A: ticker
          agora,                // B: timestamp_sync
          resAPI.symbol,        // C: symbol
          resAPI.name,          // D: name
          resAPI.resolution,    // E: resolution
          item.time,            // F: time
          item.open,            // G: open
          item.high,            // H: high
          item.low,             // I: low
          item.close,           // J: close
          item.volume,          // K: volume
          item.fvolume          // L: fvolume
        ]);
        
        bufferDeDados.push(...rows);
      }

      // Rate limit OpLab
      if (i < tickers.length - 1) Utilities.sleep(1000);
    }

    // 4. Escrita em lote para performance
    if (bufferDeDados.length > 0) {
      abaDestino.getRange(2, 1, bufferDeDados.length, headers.length).setValues(bufferDeDados);
      Logger.log(`✅ Sucesso: ${bufferDeDados.length} registros inseridos.`);
    }

  } catch (error) {
    Logger.log(`❌ Erro em ${SERVICO_NOME}: ${error.message}`);
  }
}

/**
 * Faz a chamada bruta ao endpoint historical
 */
function fetchHistoricoOpLab(ticker, token) {
  const url = `${OPLAB_BASE_URL_V3}/market/historical/${ticker}/1d?amount=250&smooth=true&df=iso`;
  const options = {
    "method": "get",
    "headers": { "Access-Token": token },
    "muteHttpExceptions": true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() === 200) {
      return JSON.parse(response.getContentText());
    }
    return null;
  } catch (e) {
    return null;
  }
}