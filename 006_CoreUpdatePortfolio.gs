/**
 * @fileoverview CoreUpdatePortfolio - v4.0 (DUD Edition)
 * AÇÃO: Sincronia de ativos com Log de Contexto Profundo e Mapeamento Dinâmico.
 * PADRÃO: Respeita o Dicionário Universal de Dados (v5.0).
 */

const PortfolioUpdater = {
  _serviceName: "PortfolioUpdater_v4.0",

  /**
   * Processa o portfólio com auditoria e mapeamento dinâmico de colunas.
   */
  syncPortfolioData() {
    const inicio = Date.now();
    
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const aba = ss.getSheetByName(SYS_CONFIG.SHEETS.IMPORT);
      if (!aba) throw new Error(`Aba não encontrada: ${SYS_CONFIG.SHEETS.IMPORT}`);

      const maxRows = aba.getLastRow();
      if (maxRows < 2) {
        SysLogger.log(this._serviceName, "AVISO", "Aba vazia ou apenas cabeçalho.", "Linhas: " + maxRows);
        return;
      }

      // 1. SCAN DINÂMICO DE CABEÇALHOS (Mapeia rótulos para índices)
      const headers = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0];
      const col = {};
      headers.forEach((label, index) => {
        if (label) col[label.trim()] = index + 1;
      });

      // Validação de colunas obrigatórias via DUD
      const req = ["OPTION_TICKER", "ID_TRADE", "TICKER", "STATUS_OP"];
      req.forEach(key => {
        if (!col[key]) throw new Error(`Coluna obrigatória '${key}' não encontrada na aba ${SYS_CONFIG.SHEETS.IMPORT}`);
      });

      const dataFull = aba.getRange(2, 1, maxRows - 1, aba.getLastColumn()).getValues();
      const linhasParaProcessar = [];
      const tickersSucesso = [];
      let contagemErro = 0;

      // 2. FASE DE MAPEAMENTO (Usa os índices descobertos no Scan)
      for (let i = 0; i < dataFull.length; i++) {
        const linhaPlanilha = i + 2; 
        const rowData = dataFull[i];
        
        const optionTicker = String(rowData[col.OPTION_TICKER - 1]).trim();
        const idTrade      = rowData[col.ID_TRADE - 1];         
        const jaEnriquecido = rowData[col.TICKER - 1]; // Se a coluna TICKER (Ação) estiver vazia, precisa processar

        if (optionTicker && idTrade && !jaEnriquecido) {
          linhasParaProcessar.push({ linha: linhaPlanilha, optionTicker: optionTicker });
        }
      }

      SysLogger.log(this._serviceName, "INFO", `Mapeamento: ${linhasParaProcessar.length} ativos pendentes.`, `Total analisado: ${maxRows}`);

      if (linhasParaProcessar.length === 0) {
        SysLogger.log(this._serviceName, "FINISH", ">>> CICLO ENCERRADO: Nada para enriquecer. <<<");
        SysLogger.flush();
        return;
      }

      // 3. FASE DE EXECUÇÃO (API + ESCRITA)
      linhasParaProcessar.forEach((item) => {
        const dadosNovos = this._fetchOptionData(item.optionTicker);
        
        if (dadosNovos) {
          // Grava em bloco: TICKER, EXPIRY, STRIKE, CATEGORY, STATUS_OP
          aba.getRange(item.linha, col.TICKER, 1, 5).setValues([dadosNovos]);
          tickersSucesso.push(item.optionTicker);
          
          SysLogger.log(this._serviceName, "SUCESSO", `Linha ${item.linha}: ${item.optionTicker} enriquecida.`, JSON.stringify(dadosNovos));
        } else {
          aba.getRange(item.linha, col.TICKER, 1, 1).setValue("ERRO_API");
          contagemErro++;
          SysLogger.log(this._serviceName, "ERRO", `Falha na API para ${item.optionTicker}`, "Retornou null");
        }
        
        if (linhasParaProcessar.length > 5) Utilities.sleep(600); 
      });

      // MARCADOR DE TERRITÓRIO: FINALIZAÇÃO
      const duracao = ((Date.now() - inicio) / 1000).toFixed(1);
      SysLogger.log(this._serviceName, "FINISH", `>>> CICLO FINALIZADO EM ${duracao}s <<<`, JSON.stringify({
        total: linhasParaProcessar.length,
        sucesso: tickersSucesso.length,
        erros: contagemErro
      }));
      SysLogger.flush(); 

    } catch (e) {
      SysLogger.log(this._serviceName, "CRITICO", "FALHA NO MOTOR 006", String(e.message));
      SysLogger.flush();
    }
  },

  /**
   * Busca dados na API e formata para o bloco de saída.
   */
  _fetchOptionData(optionTicker) {
    try {
      const data = OplabService.getOptionDetails(optionTicker);
      if (!data) return null;

      // Payload ordenado conforme as colunas da planilha: TICKER, EXPIRY, STRIKE, CATEGORY, STATUS_OP
      return [
        String(data.parent_symbol || data.symbol).toUpperCase(),
        DataUtils.formatDateBR(data.due_date || data.expiration),
        data.strike ? Number(data.strike) : "N/A",
        String(data.category || data.type || "N/A").toUpperCase(),
        "ATIVO" // Novo Status Operacional
      ];
    } catch (e) {
      return null;
    }
  }
};

// ============================================================================
// PONTO DE ENTRADA (Trigger Manual/Menu)
// ============================================================================

/** Função global para atualizar a aba de importação da corretora */
function atualizarNecton() { 
  PortfolioUpdater.syncPortfolioData(); 
}

// ============================================================================
// SUÍTE DE HOMOLOGAÇÃO REAL
// ============================================================================

function testSuitePortfolio006() {
  console.log("=== INICIANDO HOMOLOGAÇÃO MOTOR 006 (v3.3) ===");
  
  // 1. TESTE DE CONECTIVIDADE E PARSER
  const ticker = "PETRC425"; 
  console.log(`--- Testando Fetch para ${ticker} ---`);
  const dados = PortfolioUpdater._fetchOptionData(ticker);
  
  if (dados) {
    console.log("✅ Parser OK:", dados.join(" | "));
  } else {
    console.error("❌ Falha no Fetch. Verifique conexão/token.");
  }

  // 2. TESTE DE LOGGING (Verificação do Timestamp)
  console.log("--- Testando Registro de Log ---");
  SysLogger.log("TEST_006", "INFO", "Validando se o timestamp aparece na coluna A");
  SysLogger.flush();
  
  console.log("⚠️ Verifique na aba 'Logs' se a primeira linha tem data/hora.");
  console.log("=== FIM DA HOMOLOGAÇÃO ===");
}