/**
 * @fileoverview CoreSyncStockOptionsData - v4.0 (Merge Logic & Gold Standard Audit)
 * AÇÃO: Sincroniza detalhes de opções ATIVAS preservando colunas manuais.
 * MELHORIA: Lógica de Merge não apaga dados que não vêm da API (ex: anotações).
 * PADRÃO: Modo Silencioso e Contexto Serializado.
 */

const OptionDetailsSync = {
  _serviceName: "OptionDetailsSync_v4.0",

  run() {
    const inicio = Date.now();
    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);

    const cacheAPI = {};
    const stats = { api: 0, cache: 0, vencidos: 0, invalidos: 0, atualizados: 0, novos: 0 };
    const metadadosExecucao = {
      aba_gatilho: SYS_CONFIG.SHEETS.TRIGGER,
      aba_destino: SYS_CONFIG.SHEETS.DETAILS,
      timestamp_inicio: new Date().toISOString()
    };

    // MARCADOR DE TERRITÓRIO: INÍCIO
    SysLogger.log(this._serviceName, "START", ">>> INICIANDO SINCRONIZAÇÃO DE DETALHES (OPÇÕES) <<<", JSON.stringify(metadadosExecucao));

    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const abaGatilho = ss.getSheetByName(SYS_CONFIG.SHEETS.TRIGGER);
      const abaDetalhes = ss.getSheetByName(SYS_CONFIG.SHEETS.DETAILS);
      
      if (!abaGatilho || !abaDetalhes) {
        throw new Error(`Abas ausentes: Verifique ${SYS_CONFIG.SHEETS.TRIGGER} e ${SYS_CONFIG.SHEETS.DETAILS}`);
      }
      
      // --- 1. SCAN DINÂMICO ---
      const headersGatilho = abaGatilho.getRange(1, 1, 1, abaGatilho.getLastColumn()).getValues()[0];
      const colMapG = {};
      headersGatilho.forEach((h, i) => { if(h) colMapG[String(h).trim()] = i; });

      const idxID = colMapG["ID_Trade_Unico"];
      const idxEstrutura = colMapG["ID_Estrutura"];
      const idxAtivo = colMapG["Ativo"] !== undefined ? colMapG["Ativo"] : 0;
      const idxVencimento = colMapG["Vencimento"];
      const idxStatus = colMapG["Status Operação"];

      const cabecalhosDestino = abaDetalhes.getRange(1, 1, 1, abaDetalhes.getLastColumn()).getValues()[0];
      const headerMapD = {};
      cabecalhosDestino.forEach((h, i) => { if(h) headerMapD[String(h).trim()] = i; });

      // Mapa de cache para evitar múltiplas leituras da mesma linha de destino
      const idToRowMap = {};
      if (abaDetalhes.getLastRow() > 1) {
        const ids = abaDetalhes.getRange(2, 1, abaDetalhes.getLastRow() - 1, 1).getValues();
        ids.forEach((l, i) => { if (l[0]) idToRowMap[String(l[0]).trim()] = i + 2; });
      }

      const valoresGatilho = abaGatilho.getDataRange().getValues();
      
      const dataHoje = DataUtils.formatDateBR(new Date());
      const horaHoje = new Date().toLocaleTimeString('pt-BR');
      const timestampAtual = `${dataHoje} ${horaHoje}`;

      // --- 2. PROCESSAMENTO E MERGE ---
      for (let i = 1; i < valoresGatilho.length; i++) {
        const linhaAtu = valoresGatilho[i];
        const status = String(linhaAtu[idxStatus] || "").trim().toUpperCase();
        const idTrade = String(linhaAtu[idxID] || "").trim();
        const ticker = String(linhaAtu[idxAtivo] || "").trim();

        // Filtro Primário (Mais rápido)
        if (status !== "ATIVO") continue;
        if (!idTrade || idTrade.length < 10 || !isNaN(idTrade)) { stats.invalidos++; continue; }

        // Filtro de Vencimento
        const dataVenc = this._parseDate(linhaAtu[idxVencimento]);
        if (dataVenc && dataVenc < hoje) { stats.vencidos++; continue; }

        // Chamada de API / Cache
        let dadosAPI = cacheAPI[ticker] || null;
        if (!dadosAPI) {
          dadosAPI = OplabService.getOptionDetails(ticker);
          if (dadosAPI) { 
            cacheAPI[ticker] = dadosAPI; 
            stats.api++; 
            Utilities.sleep(1300); // Rate Limit da API
          }
        } else {
          stats.cache++;
        }

        if (dadosAPI) {
          const rowNum = idToRowMap[idTrade];
          let linhaParaGravar;

          // --- LÓGICA DE MERGE (Preservação de Colunas Manuais) ---
          if (rowNum) {
            // Se já existe, lê a linha atual para não apagar colunas extras (Ex: Anotações)
            linhaParaGravar = abaDetalhes.getRange(rowNum, 1, 1, cabecalhosDestino.length).getValues()[0];
            stats.atualizados++;
          } else {
            // Se for novo, cria array vazio
            linhaParaGravar = new Array(cabecalhosDestino.length).fill("");
            stats.novos++;
          }

          // Injeta os dados da API apenas nas colunas mapeadas
          const snapshot = JSON.parse(JSON.stringify(dadosAPI));
          snapshot['ID_Trade_Unico'] = idTrade;
          snapshot['ID_Estrutura'] = String(linhaAtu[idxEstrutura] || "").trim();
          snapshot['Timestamp_Atualizacao'] = timestampAtual;
          
          if (snapshot['due_date']) snapshot['due_date'] = DataUtils.formatDateBR(snapshot['due_date']);

          for (const key in headerMapD) {
            const colIdx = headerMapD[key];
            if (snapshot[key] !== undefined) {
              linhaParaGravar[colIdx] = typeof snapshot[key] === 'object' ? JSON.stringify(snapshot[key]) : snapshot[key];
            }
          }

          // Gravação
          if (rowNum) {
            abaDetalhes.getRange(rowNum, 1, 1, cabecalhosDestino.length).setValues([linhaParaGravar]);
          } else {
            abaDetalhes.appendRow(linhaParaGravar);
            idToRowMap[idTrade] = abaDetalhes.getLastRow();
          }
          
          SysLogger.log(this._serviceName, "SUCESSO", `Opção ${ticker} processada.`, `ID_Trade: ${idTrade} | Tipo: ${rowNum ? 'Atualização' : 'Novo'}`);
        }
      }

      // MARCADOR DE TERRITÓRIO: FINALIZAÇÃO
      const duracao = ((Date.now() - inicio) / 1000).toFixed(1);
      stats.duracao_segundos = duracao;

      SysLogger.log(this._serviceName, "FINISH", `>>> CICLO FINALIZADO EM ${duracao}s <<<`, JSON.stringify(stats));
      SysLogger.flush();

    } catch (e) {
      SysLogger.log(this._serviceName, "CRITICO", "Falha fatal no motor de Detalhes de Opções", String(e.message));
      SysLogger.flush();
    }
  },

  _parseDate(raw) {
    if (raw instanceof Date) return raw;
    if (!raw) return null;
    const s = String(raw);
    const partes = s.split('/');
    if (partes.length === 3) return new Date(partes[2], partes[1] - 1, partes[0]);
    return new Date(s);
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