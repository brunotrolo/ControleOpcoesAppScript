/**
 * @fileoverview CoreSyncStockData - v4.0 (DUD Edition)
 * AÇÃO: Sincroniza dados de mercado (Ações) na aba DADOS_ATIVOS.
 * PADRÃO: Mapeamento Dinâmico via DUD (v5.0) e Contexto Serializado.
 */

const StockDataSync = {
  _serviceName: "StockDataSync_v4.0",

  /**
   * Processa a sincronização de dados de ativos com a API.
   */
  run() {
    const inicio = Date.now();
    
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const abaImport = ss.getSheetByName(SYS_CONFIG.SHEETS.IMPORT);
      const abaAtivos = ss.getSheetByName(SYS_CONFIG.SHEETS.ASSETS);
      
      if (!abaImport || !abaAtivos) {
        throw new Error(`Abas não encontradas. Verifique ${SYS_CONFIG.SHEETS.IMPORT} e ${SYS_CONFIG.SHEETS.ASSETS}`);
      }

      // 1. OBTENÇÃO DINÂMICA DOS TICKERS ALVOS
      const headersImport = abaImport.getRange(1, 1, 1, abaImport.getLastColumn()).getValues()[0];
      const colTickerIdx = headersImport.indexOf("TICKER");

      if (colTickerIdx === -1) {
        throw new Error("Coluna 'TICKER' não encontrada na aba " + SYS_CONFIG.SHEETS.IMPORT);
      }

      const ultimaLinhaImport = abaImport.getLastRow();
      if (ultimaLinhaImport < 2) {
        SysLogger.log(this._serviceName, "AVISO", "Aba Import vazia. Nada a sincronizar.");
        return;
      }

      // Lê a coluna TICKER (Ação Objeto) preenchida pelo motor 006
      const valoresImport = abaImport.getRange(2, colTickerIdx + 1, ultimaLinhaImport - 1, 1).getValues();
      const tickersAlvo = [...new Set(valoresImport.flat().filter(t => t && String(t).trim() !== "" && t !== "ERRO_API" && t !== "N/A"))];
      
      SysLogger.log(this._serviceName, "INFO", `Fase 1: Identificados ${tickersAlvo.length} tickers únicos.`, JSON.stringify(tickersAlvo));
      
      if (tickersAlvo.length === 0) {
        SysLogger.log(this._serviceName, "FINISH", ">>> CICLO ENCERRADO: Nenhum ticker para buscar. <<<");
        SysLogger.flush();
        return;
      }

      // 2. MAPEAMENTO DINÂMICO DA ABA DADOS_ATIVOS
      const ultimaLinhaAtivos = abaAtivos.getLastRow();
      const ultimaColunaAtivos = abaAtivos.getLastColumn();
      const cabecalhosAtivos = abaAtivos.getRange(1, 1, 1, ultimaColunaAtivos).getValues()[0];
      
      const headerMap = {};
      cabecalhosAtivos.forEach((h, i) => { if(h) headerMap[String(h).trim()] = i; });

      // Mapeia ativos existentes para realizar UPDATE em vez de APPEND
      const idToRowMap = {};
      if (ultimaLinhaAtivos > 1) {
        const tickersExistentes = abaAtivos.getRange(2, 1, ultimaLinhaAtivos - 1, 1).getValues();
        tickersExistentes.forEach((linha, index) => {
          if (linha[0]) idToRowMap[String(linha[0]).trim()] = index + 2;
        });
      }

      // 3. PROCESSAMENTO VIA API COM TRADUÇÃO DUD
      const timestampAtual = DataUtils.formatDateBR(new Date()) + " " + new Date().toLocaleTimeString('pt-BR');
      
      // DE-PARA: Mapeia o rótulo da Planilha para a chave que a API OpLab retorna
      const tradutorAPI = {
        "SPOT": "close",
        "IV": "iv_current",
        "IV_RANK": "iv_1y_rank",
        "COMPANY_NAME": "name",
        "VARIATION": "variation",
        "UPDATED_AT": "manual_timestamp",
        "TICKER": "manual_ticker"
      };

      const listaParaNovos = [];
      const updatesEmLote = [];
      let contagemErro = 0;

      tickersAlvo.forEach((ticker, i) => {
        const dadosAPI = OplabService.getStockData(ticker);
        
        if (dadosAPI) {
          const linhaValores = new Array(ultimaColunaAtivos).fill("");
          
          for (const label in headerMap) {
            const index = headerMap[label];
            const apiKey = tradutorAPI[label] || label.toLowerCase(); // Tenta traduzir ou usa o label em minúsculo

            if (label === 'TICKER') {
              linhaValores[index] = ticker;
            } else if (label === 'UPDATED_AT') {
              linhaValores[index] = timestampAtual;
            } else if (dadosAPI[apiKey] !== undefined && dadosAPI[apiKey] !== null) {
              linhaValores[index] = dadosAPI[apiKey];
            }
          }

          if (idToRowMap[ticker]) {
            updatesEmLote.push({ linha: idToRowMap[ticker], dados: linhaValores });
          } else {
            listaParaNovos.push(linhaValores);
          }
        } else {
           contagemErro++;
           SysLogger.log(this._serviceName, "ERRO", `Falha na API para: ${ticker}`);
        }
        
        if (i % 5 === 0) Utilities.sleep(600); 
      });

      // 4. ESCRITA EM LOTE
      updatesEmLote.forEach(update => {
        abaAtivos.getRange(update.linha, 1, 1, ultimaColunaAtivos).setValues([update.dados]);
      });

      if (listaParaNovos.length > 0) {
        abaAtivos.getRange(ultimaLinhaAtivos + 1, 1, listaParaNovos.length, ultimaColunaAtivos).setValues(listaParaNovos);
      }

      // MARCADOR DE TERRITÓRIO: FINALIZAÇÃO
      const duracao = ((Date.now() - inicio) / 1000).toFixed(1);
      SysLogger.log(this._serviceName, "FINISH", `>>> SINCRONIA DE ATIVOS CONCLUÍDA EM ${duracao}s <<<`, JSON.stringify({
        atualizados: updatesEmLote.length,
        novos: listaParaNovos.length,
        erros: contagemErro
      }));
      SysLogger.flush();
      
    } catch (e) {
      SysLogger.log(this._serviceName, "CRITICO", "Falha no motor 007", String(e.message));
      SysLogger.flush();
    }
  }
};

// ============================================================================
// PONTO DE ENTRADA (Trigger Manual/Menu)
// ============================================================================

function atualizarDadosAtivos() {
  StockDataSync.run();
}

// ============================================================================
// SUÍTE DE HOMOLOGAÇÃO
// ============================================================================

function testSuiteStockDataSync007() {
  console.log("=== INICIANDO HOMOLOGAÇÃO: MOTOR 007 ===");
  
  // 1. TESTE DA API (Conexão e Resposta)
  console.log("--- Passo 1: Testando API OpLab para Stocks ---");
  const tickerTeste = "PETR4"; 
  const dados = OplabService.getStockData(tickerTeste);
  
  if (dados && dados.close !== undefined) {
    console.log(`✅ Conexão OK. Fechamento de ${tickerTeste}: R$ ${dados.close}`);
  } else {
    console.error(`❌ Falha Crítica: API não retornou cotação para ${tickerTeste}`);
  }

  // 2. TESTE DE AMBIENTE (Leitura Dinâmica)
  console.log("--- Passo 2: Validando Cabeçalhos da Aba de Destino ---");
  try {
    const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SYS_CONFIG.SHEETS.ASSETS);
    const colunas = aba.getLastColumn();
    const headers = aba.getRange(1, 1, 1, colunas).getValues()[0];
    
    if (headers.includes("Ticker") && headers.includes("Timestamp_Atualizacao")) {
      console.log("✅ Cabeçalhos obrigatórios encontrados.");
      console.log(`ℹ️ Total de colunas mapeadas: ${colunas}`);
    } else {
      console.error("❌ AVISO: A aba Dados_Ativos precisa das colunas 'Ticker' e 'Timestamp_Atualizacao'.");
    }
  } catch (e) {
    console.error("❌ Falha ao acessar aba Dados_Ativos:", e.message);
  }

  // 3. EXECUÇÃO REAL CONTROLADA
  console.log("--- Passo 3: Iniciando Sincronização Controlada ---");
  StockDataSync.run();
  
  console.log("=== FIM DA HOMOLOGAÇÃO 007 ===");
}