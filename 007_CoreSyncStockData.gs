/**
 * @fileoverview CoreSyncStockData - v3.1 (Dynamic Mapping & Gold Standard Audit)
 * AÇÃO: Sincroniza dados de mercado (Ações/Ativo Objeto) na aba Dados_Ativos.
 * PADRÃO: Respeita o contrato do SysLogger (003) com Contexto Serializado.
 */

const StockDataSync = {
  _serviceName: "StockDataSync_v3.1",

  /**
   * Processa a sincronização de dados de ativos com a API.
   */
  run() {
    const inicio = Date.now();
    const metadadosExecucao = {
      aba_gatilho: SYS_CONFIG.SHEETS.TRIGGER,
      aba_destino: SYS_CONFIG.SHEETS.ASSETS,
      timestamp_inicio: new Date().toISOString()
    };

    // MARCADOR DE TERRITÓRIO: INÍCIO
    SysLogger.log(this._serviceName, "START", ">>> INICIANDO SINCRONIZAÇÃO DE ATIVOS (STOCKS) <<<", JSON.stringify(metadadosExecucao));

    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const abaGatilho = ss.getSheetByName(SYS_CONFIG.SHEETS.TRIGGER);
      const abaAtivos = ss.getSheetByName(SYS_CONFIG.SHEETS.ASSETS);
      
      if (!abaGatilho || !abaAtivos) {
        throw new Error(`Abas não encontradas. Verifique ${SYS_CONFIG.SHEETS.TRIGGER} e ${SYS_CONFIG.SHEETS.ASSETS}`);
      }

      // 1. OBTENÇÃO DOS TICKERS ALVOS
      const ultimaLinhaGatilho = abaGatilho.getLastRow();
      if (ultimaLinhaGatilho < 2) {
        SysLogger.log(this._serviceName, "AVISO", "Aba Gatilho vazia. Nada a sincronizar.", "Linhas: " + ultimaLinhaGatilho);
        return;
      }

      // Lê a coluna onde os ativos objetos foram salvos pelo 006 (Coluna S)
      const valoresGatilho = abaGatilho.getRange(2, SYS_CONFIG.COLUMNS.OUTPUT_START, ultimaLinhaGatilho - 1, 1).getValues();
      
      // Filtra únicos, ignorando vazios, erros e N/A
      const tickersAlvo = [...new Set(valoresGatilho.flat().filter(t => t && String(t).trim() !== "" && t !== "ERRO_API" && t !== "N/A"))];
      
      SysLogger.log(this._serviceName, "INFO", `Fase 1: Identificados ${tickersAlvo.length} tickers únicos.`, "Lista: " + JSON.stringify(tickersAlvo));
      
      if (tickersAlvo.length === 0) {
        SysLogger.log(this._serviceName, "FINISH", ">>> CICLO ENCERRADO: Nenhum ativo válido para buscar. <<<", "{}");
        SysLogger.flush();
        return;
      }

      // 2. MAPEAMENTO DINÂMICO E ESTADO ATUAL
      const ultimaLinhaAtivos = abaAtivos.getLastRow();
      const ultimaColunaAtivos = abaAtivos.getLastColumn();
      
      if (ultimaColunaAtivos < 2) {
        throw new Error("Aba Dados_Ativos não possui cabeçalhos suficientes para mapeamento.");
      }

      const cabecalhosAtivos = abaAtivos.getRange(1, 1, 1, ultimaColunaAtivos).getValues()[0];
      const headerMap = {};
      cabecalhosAtivos.forEach((h, i) => { 
        if(h) headerMap[String(h).trim()] = i; 
      });

      SysLogger.log(this._serviceName, "INFO", "Fase 2: Mapeamento de colunas criado.", JSON.stringify(headerMap));

      // Mapeia ativos existentes
      const idToRowMap = {};
      if (ultimaLinhaAtivos > 1) {
        const tickersExistentes = abaAtivos.getRange(2, 1, ultimaLinhaAtivos - 1, 1).getValues();
        tickersExistentes.forEach((linha, index) => {
          if (linha[0]) idToRowMap[String(linha[0]).trim()] = index + 2;
        });
      }

      // 3. PROCESSAMENTO VIA API
      const dataHoje = DataUtils.formatDateBR(new Date());
      const horaHoje = new Date().toLocaleTimeString('pt-BR');
      const timestampAtual = `${dataHoje} ${horaHoje}`;
      
      const listaParaNovos = [];
      const updatesEmLote = [];
      let contagemErro = 0;

      tickersAlvo.forEach((ticker, i) => {
        const dadosAPI = OplabService.getStockData(ticker);
        
        if (dadosAPI) {
          const linhaValores = new Array(ultimaColunaAtivos).fill("");
          
          for (const key in headerMap) {
            const index = headerMap[key];
            if (key === 'Ticker') {
              linhaValores[index] = ticker;
            } else if (key === 'Timestamp_Atualizacao') {
              linhaValores[index] = timestampAtual;
            } else if (dadosAPI[key] !== undefined && dadosAPI[key] !== null) {
               linhaValores[index] = typeof dadosAPI[key] === 'object' ? JSON.stringify(dadosAPI[key]) : dadosAPI[key];
            }
          }

          if (idToRowMap[ticker]) {
            updatesEmLote.push({ linha: idToRowMap[ticker], dados: linhaValores });
          } else {
            listaParaNovos.push(linhaValores);
          }
          
          SysLogger.log(this._serviceName, "SUCESSO", `Dados formatados para ${ticker}`, "Payload: " + JSON.stringify(linhaValores));
          
        } else {
           contagemErro++;
           SysLogger.log(this._serviceName, "ERRO", `Falha na API para o ativo: ${ticker}`, "Retornou null");
        }
        
        if (i % 5 === 0) Utilities.sleep(600); // Rate Limit
      });

      // 4. ESCRITA EM LOTE (Batch Write)
      SysLogger.log(this._serviceName, "INFO", `Fase 4: Gravando ${updatesEmLote.length} atualizações e ${listaParaNovos.length} novos.`);

      // Atualiza os existentes
      updatesEmLote.forEach(update => {
        abaAtivos.getRange(update.linha, 1, 1, ultimaColunaAtivos).setValues([update.dados]);
      });

      // Insere os novos em bloco
      if (listaParaNovos.length > 0) {
        abaAtivos.getRange(ultimaLinhaAtivos + 1, 1, listaParaNovos.length, ultimaColunaAtivos).setValues(listaParaNovos);
      }

      // MARCADOR DE TERRITÓRIO: FINALIZAÇÃO
      const duracao = ((Date.now() - inicio) / 1000).toFixed(1);
      const resumoFinal = {
        tickers_solicitados: tickersAlvo.length,
        atualizados: updatesEmLote.length,
        inseridos_novos: listaParaNovos.length,
        erros: contagemErro,
        duracao_segundos: duracao
      };

      SysLogger.log(this._serviceName, "FINISH", `>>> CICLO FINALIZADO EM ${duracao}s <<<`, JSON.stringify(resumoFinal));
      SysLogger.flush();
      
    } catch (e) {
      // Passando String(e.message) para proteger a coluna de Timestamp
      SysLogger.log(this._serviceName, "CRITICO", "Falha na orquestração de sync dados ativos", String(e.message));
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