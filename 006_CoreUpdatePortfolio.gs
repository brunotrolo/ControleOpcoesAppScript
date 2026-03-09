/**
 * @fileoverview CoreUpdatePortfolio - v3.3 (Audit Trail Edition)
 * AÇÃO: Sincronia de ativos com Log de Contexto Profundo e marcadores de território.
 */

const PortfolioUpdater = {
  _serviceName: "PortfolioUpdater_v3.3",

  /**
   * Processa o portfólio com auditoria completa em todas as fases.
   */
  syncPortfolioData() {
    const inicio = Date.now();
    const metadadosExecucao = {
      planilha: SYS_CONFIG.SHEETS.TRIGGER,
      coluna_entrada: SYS_CONFIG.COLUMNS.TRIGGER_INPUT,
      coluna_saida: SYS_CONFIG.COLUMNS.OUTPUT_START,
      timestamp_inicio: new Date().toISOString()
    };

    // MARCADOR DE TERRITÓRIO: INÍCIO
    SysLogger.log(this._serviceName, "START", ">>> INICIANDO CICLO DE ATUALIZAÇÃO DO PORTFÓLIO <<<", JSON.stringify(metadadosExecucao));

    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const aba = ss.getSheetByName(SYS_CONFIG.SHEETS.TRIGGER);
      if (!aba) throw new Error(`Aba não encontrada: ${SYS_CONFIG.SHEETS.TRIGGER}`);

      const maxRows = aba.getLastRow();
      if (maxRows < 2) {
        SysLogger.log(this._serviceName, "AVISO", "Processamento abortado: Aba vazia ou apenas cabeçalho.", "Linhas encontradas: " + maxRows);
        return;
      }

      const dataFull = aba.getRange(2, 1, maxRows - 1, SYS_CONFIG.COLUMNS.OUTPUT_END).getValues();
      const linhasParaProcessar = [];
      const tickersSucesso = [];
      let contagemErro = 0;

      // 1. FASE DE MAPEAMENTO
      for (let i = 0; i < dataFull.length; i++) {
        const linhaPlanilha = i + 2; 
        const rowData = dataFull[i];
        const ticker = String(rowData[SYS_CONFIG.COLUMNS.TRIGGER_INPUT - 1]).trim();
        const idTrade = rowData[SYS_CONFIG.COLUMNS.ID_FORMULA - 1];        
        const jaProcessado = rowData[SYS_CONFIG.COLUMNS.OUTPUT_START - 1]; 

        if (ticker && idTrade && !jaProcessado) {
          linhasParaProcessar.push({ linha: linhaPlanilha, ticker: ticker });
        }
      }

      SysLogger.log(this._serviceName, "INFO", `Fase de Mapeamento: ${linhasParaProcessar.length} ativos pendentes encontrados.`, "Total linhas analisadas: " + maxRows);

      if (linhasParaProcessar.length === 0) {
        SysLogger.log(this._serviceName, "FINISH", ">>> CICLO ENCERRADO: Nada para processar. <<<");
        SysLogger.flush();
        return;
      }

      // 2. FASE DE EXECUÇÃO (API + ESCRITA)
      linhasParaProcessar.forEach((item) => {
        const dadosNovos = this._fetchOptionData(item.ticker);
        
        if (dadosNovos) {
          aba.getRange(item.linha, SYS_CONFIG.COLUMNS.OUTPUT_START, 1, 5).setValues([dadosNovos]);
          tickersSucesso.push(item.ticker);
          
          // Log detalhado com o payload que foi para a planilha
          SysLogger.log(this._serviceName, "SUCESSO", `Linha ${item.linha} atualizada: ${item.ticker}`, "Payload: " + JSON.stringify(dadosNovos));

        } else {
          aba.getRange(item.linha, SYS_CONFIG.COLUMNS.OUTPUT_START, 1, 1).setValue("ERRO_API");
          contagemErro++;
          SysLogger.log(this._serviceName, "ERRO", `Falha na atualização da linha ${item.linha} (${item.ticker})`, "API retornou null ou erro de parser.");
        }
        
        if (linhasParaProcessar.length > 5) Utilities.sleep(600); 
      });

      // MARCADOR DE TERRITÓRIO: FINALIZAÇÃO
      const duracao = ((Date.now() - inicio) / 1000).toFixed(1);
      const resumoFinal = {
        total_pendentes: linhasParaProcessar.length,
        sucesso: tickersSucesso.length,
        erros: contagemErro,
        duracao_segundos: duracao,
        lista_ativos: tickersSucesso
      };

      SysLogger.log(this._serviceName, "FINISH", `>>> CICLO FINALIZADO COM SUCESSO EM ${duracao}s <<<`, JSON.stringify(resumoFinal));
      SysLogger.flush(); 
      
    } catch (e) {
      // Passando String(e.message) para proteger a coluna de Timestamp
      SysLogger.log(this._serviceName, "CRITICO", "FALHA CATASTRÓFICA NO MOTOR 006", String(e.message));
      SysLogger.flush();
    }
  },

  /**
   * Helper privado com log de debug do payload bruto.
   */
  _fetchOptionData(ticker) {
    try {
      const data = OplabService.getOptionDetails(ticker);
      if (!data) return null;

      const payload = [
        String(data.parent_symbol || data.symbol || data.underlying).toUpperCase(),
        DataUtils.formatDateBR(data.due_date || data.expiration),
        data.strike ? Number(data.strike) : "N/A",
        String(data.category || data.type || "N/A").toUpperCase(),
        "ATIVO"
      ];

      return payload;
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