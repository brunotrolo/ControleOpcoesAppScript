/**
 * @fileoverview CoreSyncStockDataHistory - v3.1 (Gold Standard Audit)
 * AÇÃO: Busca série histórica (250 dias) na OpLab e aplica Drop & Replace seguro.
 * ALVO: Aba Dados_Ativos_Historico250d (Crucial para Beta e Correlação).
 * PADRÃO: Modo Silencioso e Contexto Serializado.
 */

const HistoricalDataSync = {
  _serviceName: "HistoricalDataSync_v3.1",
  _abaDestino: "Dados_Ativos_Historico250d", 
  _diasHistorico: 250,

  /**
   * Executa a extração em lote e substituição da base de dados históricos.
   */
  run() {
    const inicio = Date.now();
    const metadadosExecucao = {
      aba_origem: SYS_CONFIG.SHEETS.ASSETS,
      aba_destino: this._abaDestino,
      dias_alvo: this._diasHistorico,
      timestamp_inicio: new Date().toISOString()
    };

    // MARCADOR DE TERRITÓRIO: INÍCIO
    SysLogger.log(this._serviceName, "START", ">>> INICIANDO SINCRONIZAÇÃO DE HISTÓRICO (250 DIAS) <<<", JSON.stringify(metadadosExecucao));

    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const abaAtivos = ss.getSheetByName(SYS_CONFIG.SHEETS.ASSETS); 
      
      if (!abaAtivos) throw new Error(`Aba origem (${SYS_CONFIG.SHEETS.ASSETS}) não encontrada.`);

      // 1. EXTRAÇÃO DE TICKERS ÚNICOS (DA ABA DADOS_ATIVOS)
      const maxRows = abaAtivos.getLastRow();
      if (maxRows < 2) {
        SysLogger.log(this._serviceName, "AVISO", "Aba Dados_Ativos vazia. Nenhum histórico a buscar.", "Linhas encontradas: " + maxRows);
        return;
      }

      // Lê a coluna A (Assumindo que os Tickers estão na primeira coluna)
      const tickersBrutos = abaAtivos.getRange(2, 1, maxRows - 1, 1).getValues().flat();
      const tickersAlvo = [...new Set(tickersBrutos.filter(t => t && String(t).trim() !== ""))];

      // Garante que o IBOV seja baixado para servir de benchmark (Beta/Correlação)
      if (!tickersAlvo.includes("IBOV")) {
        tickersAlvo.push("IBOV");
      }

      SysLogger.log(this._serviceName, "INFO", `Fase 1: Mapeados ${tickersAlvo.length} ativos para histórico (Incluindo IBOV).`, "Lista: " + JSON.stringify(tickersAlvo));
      
      if (tickersAlvo.length === 0) {
        SysLogger.log(this._serviceName, "FINISH", ">>> CICLO ENCERRADO: Nada para processar. <<<", "{}");
        SysLogger.flush();
        return;
      }

      // 2. BUSCA NA API (Em Lote com Rate Limit)
      const bufferDeDados = [];
      let contagemErro = 0;
      
      const dataHoje = DataUtils.formatDateBR(new Date());
      const horaHoje = new Date().toLocaleTimeString('pt-BR');
      const timestampSync = `${dataHoje} ${horaHoje}`;

      tickersAlvo.forEach((ticker, index) => {
        // Uso do Client Centralizado (Seguro e Tratado)
        const resAPI = OplabService.getHistoricalData(ticker, this._diasHistorico);
        
        if (resAPI && resAPI.data && resAPI.data.length > 0) {
          const symbolName = resAPI.symbol || ticker;
          const companyName = resAPI.name || "N/A";
          const resolution = resAPI.resolution || "1d";

          const rows = resAPI.data.map(item => [
            ticker,               // A: ticker pesquisado
            timestampSync,        // B: momento da sincronização
            symbolName,           // C: symbol (da API)
            companyName,          // D: name
            resolution,           // E: resolution
            item.time,            // F: time (Data do pregão, ex: 2026-03-08)
            item.open,            // G: open
            item.high,            // H: high
            item.low,             // I: low
            item.close,           // J: close
            item.volume,          // K: volume
            item.fvolume          // L: fvolume (Financeiro)
          ]);
          
          bufferDeDados.push(...rows);
          SysLogger.log(this._serviceName, "SUCESSO", `Histórico de ${ticker} baixado.`, `Candles: ${rows.length}`);
          
        } else {
          contagemErro++;
          SysLogger.log(this._serviceName, "ERRO", `Falha ou dados vazios para o histórico de ${ticker}.`, "Retorno da API vazio ou inválido.");
        }

        // Rate limit para não engasgar a Oplab com requisições pesadas (250 dias = payload grande)
        if (index < tickersAlvo.length - 1) Utilities.sleep(800); 
      });

      // FAIL-SAFE: Se a API cair para todos os ativos, não apagamos o banco antigo!
      if (bufferDeDados.length === 0) {
        throw new Error("Nenhum dado histórico retornado da API. Operação cancelada para proteger o banco atual.");
      }

      // 3. DROP & REPLACE (Gravação Segura em Bloco)
      SysLogger.log(this._serviceName, "INFO", `Fase 3: Iniciando substituição da base de dados.`, `Total de candles no buffer: ${bufferDeDados.length}`);
      
      let abaDestino = ss.getSheetByName(this._abaDestino);
      if (!abaDestino) {
        abaDestino = ss.insertSheet(this._abaDestino);
        SysLogger.log(this._serviceName, "AVISO", `Aba '${this._abaDestino}' não existia e foi criada automaticamente.`, "");
      }

      const headers = [
        "ticker", "timestamp_sync", "symbol", "name", "resolution", 
        "time", "open", "high", "low", "close", "volume", "fvolume"
      ];

      // Limpa os dados antigos apenas após o Fail-Safe passar
      abaDestino.clearContents(); 
      
      // Escreve Cabeçalho e Dados em um único pulso de I/O
      abaDestino.getRange(1, 1, 1, headers.length).setValues([headers]);
      abaDestino.getRange(2, 1, bufferDeDados.length, headers.length).setValues(bufferDeDados);

      // MARCADOR DE TERRITÓRIO: FINALIZAÇÃO
      const duracao = ((Date.now() - inicio) / 1000).toFixed(1);
      const resumoFinal = {
        ativos_processados: tickersAlvo.length - contagemErro,
        linhas_inseridas: bufferDeDados.length,
        erros_api: contagemErro,
        duracao_segundos: duracao
      };

      SysLogger.log(this._serviceName, "FINISH", `>>> CICLO FINALIZADO EM ${duracao}s <<<`, JSON.stringify(resumoFinal));
      SysLogger.flush();

    } catch (e) {
      SysLogger.log(this._serviceName, "CRITICO", "Falha fatal no Sync de Histórico", String(e.message));
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