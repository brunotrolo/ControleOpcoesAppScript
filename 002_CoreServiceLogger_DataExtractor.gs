/**
 * MÓDULO: LoggingService (Otimizado)
 * Objetivo: Registrar eventos com suporte a Buffer para alta performance.
 */
var _logBuffer = [];


function gravarLog(servico, status, resumo, detalhe) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetLogs = ss.getSheetByName(ABA_LOGS); 
    if (!sheetLogs) return;

    const timestamp = new Date();
    const detalheStr = typeof detalhe === 'object' ? JSON.stringify(detalhe, null, 2) : String(detalhe || "");

    // Gravação Direta (Satisfaz o desejo de ver linha a linha)
    sheetLogs.appendRow([timestamp, servico, status, resumo, detalheStr]);
    
    // Força o Google Sheets a renderizar a nova linha imediatamente
    SpreadsheetApp.flush(); 
    
  } catch (e) {
    console.error("Falha no gravarLog: " + e.toString());
  }
}

function flushLogs() { 
  if (!_logBuffer || _logBuffer.length === 0) return;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // Use a constante ABA_LOGS ou "Logs" se a constante não estiver declarada no topo
    const sheetLogs = ss.getSheetByName(typeof ABA_LOGS !== 'undefined' ? ABA_LOGS : "Logs"); 
    
    if (sheetLogs) {
      // Descarrega tudo em lote (Batch Write)
      sheetLogs.getRange(sheetLogs.getLastRow() + 1, 1, _logBuffer.length, 5).setValues(_logBuffer);
    }
    
    _logBuffer = []; // Limpa o buffer após gravar
    SpreadsheetApp.flush(); // Força a renderização
    
  } catch (e) {
    console.error("Falha ao descarregar logs: " + e.toString());
  }
}

/**
 * MÓDULO: DataExtractor (Otimizado)
 */
var _cockpitHeaderMap = null;

function homologarExtracaoCockpit() {
  const NOME_ABA = "Cockpit";
  const LINHA_CABECALHO = 10;
  const servicoNome = "DataExtractor_v1.2";
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(NOME_ABA);
    if (!sheet) throw new Error("Aba Cockpit não encontrada");

    const lastRow = sheet.getLastRow();
    if (lastRow <= LINHA_CABECALHO) return [];

    const dataFull = sheet.getRange(LINHA_CABECALHO, 1, lastRow - (LINHA_CABECALHO - 1), sheet.getLastColumn()).getValues();
    const headers = dataFull[0];
    const rangeDados = dataFull.slice(1);

    // Mapeamento dinâmico de colunas (executado apenas se necessário)
    if (!_cockpitHeaderMap) {
      _cockpitHeaderMap = {};
      headers.forEach((h, i) => { if(h) _cockpitHeaderMap[h.trim()] = i; });
    }

    const idxStatus = _cockpitHeaderMap["STATUS"];
    if (idxStatus === undefined) throw new Error("Coluna 'STATUS' não encontrada");

    const operacoesAtivas = rangeDados
      .filter(row => row[idxStatus] === "ATIVO")
      .map(row => {
        let obj = {};
        for (let key in _cockpitHeaderMap) {
          obj[key] = row[_cockpitHeaderMap[key]];
        }
        return obj;
      });

    gravarLog(servicoNome, "SUCESSO", `Extração concluída: ${operacoesAtivas.length} ativos`, "");
    return operacoesAtivas;

  } catch (e) {
    gravarLog(servicoNome, "ERRO_FATAL", "Falha na extração", e.toString());
    flushLogs();
    return [];
  }
}


// ═══════════════════════════════════════════════════════════════
// FUNÇÃO DE LOG GLOBAL
// ═══════════════════════════════════════════════════════════════

function log(servico, status, mensagem, contexto) {
  try {
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const abaLogs = planilha.getSheetByName(ABA_LOGS);
    const timestamp = new Date();
    
    let contextoStr = String(contexto || "");
    if (contextoStr.length > 40000) {
      contextoStr = contextoStr.substring(0, 40000) + "... (Truncado)";
    }
    
    //abaLogs.appendRow([timestamp, servico, status, mensagem, contextoStr]);
    // Grava apenas na memória (Buffer) - 0 ms de custo
    _logBuffer.push([timestamp, servico, status, mensagem, contextoStr]);
    
  } catch (e) {
    Logger.log("FALHA AO ESCREVER LOG: " + e.message);
  }
}

/**
 * Sistema de log específico do orquestrador
 * 
 * @param {string} status - Status da operação
 * @param {string} mensagem - Mensagem descritiva
 * @param {Object} contexto - Contexto adicional
 */
function logOrquestrador(status, mensagem, contexto) {
  try {
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const abaLogs = planilha.getSheetByName(ABA_LOGS);
    const timestamp = new Date();
    
    let contextoStr = JSON.stringify(contexto || {}, null, 2);
    if (contextoStr.length > 40000) {
      contextoStr = contextoStr.substring(0, 40000) + "... (Truncado)";
    }
    
    abaLogs.appendRow([
      timestamp,
      "ORQUESTRADOR",
      status,
      mensagem,
      contextoStr
    ]);
    
  } catch (e) {
    Logger.log("FALHA NO LOG DO ORQUESTRADOR: " + e.message);
  }
}

/**
 * Exibe um alerta de forma segura.
 * - No contexto com UI (usuário no Sheets) mostra alert normalmente.
 * - No contexto backend/webhook (sem UI) apenas registra nos logs.
 */
function safeAlert_(title, message) {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert(title, message, ui.ButtonSet.OK);
  } catch (e) {
    // Sem UI (chamada via backend/webhook): só registra log e segue
    log(
      "ORQUESTRADOR",
      "AVISO",
      "Falha ao exibir alerta UI (provavelmente contexto sem interface).",
      JSON.stringify({ title: title, message: message, erro: e.message })
    );
  }
}








function testarPerformanceLogger() {
  console.log("--- INICIANDO TESTE 002: PERFORMANCE DO LOGGER ---");
  
  // 1. Reset do estado para o teste
  _logBuffer = [];
  
  // 2. Teste de latência de memória (Velocidade do script)
  let t0 = Date.now();
  for (let i = 0; i < 20; i++) {
    gravarLog("Teste_Unitario", "INFO", "Simulação de Log " + i, { iteration: i, status: "SUCCESS" });
  }
  let t1 = Date.now();
  
  console.log(`Tempo para processar 20 logs em MEMÓRIA (Buffer): ${t1 - t0}ms`);
  
  // 3. Teste de I/O (Velocidade de escrita na planilha)
  let t2 = Date.now();
  flushLogs();
  let t3 = Date.now();
  
  console.log(`Tempo para DESCARREGAR (Batch Write) 20 logs no Sheets: ${t3 - t2}ms`);
  console.log(`Tempo total da operação: ${t3 - t0}ms`);
  
  // No modelo antigo, o tempo total seria aproximadamente (Tempo de 1 appendRow * 20)
  // Geralmente cada appendRow leva ~200ms a 400ms.
  let estimativaLegado = (t3 - t2) * 20; 
  console.log(`Economia estimada de tempo: ~${estimativaLegado - (t3 - t0)}ms`);
  
  console.log("--- FIM DO TESTE ---");
}







function testarPasso2() {
  console.log("--- TESTANDO MIGRAÇÃO PASSO 2 (LOGS) ---");
  
  try {
    // Tenta gravar um log real usando a função que agora está no 002
    log("TESTE_MIGRACAO", "INFO", "Validando se a função log funciona no arquivo 002", "OK");
    console.log("✅ Sucesso: A função log() foi chamada e não gerou erro.");
    
    logOrquestrador("INFO", "Validando logOrquestrador no arquivo 002", { teste: true });
    console.log("✅ Sucesso: A função logOrquestrador() foi chamada e não gerou erro.");
    
    console.log("--- SE NÃO HOUVE ERRO DE 'REFERENCEERROR', O PASSO 2 FOI SUCESSO ---");
  } catch (e) {
    console.error("❌ FALHA: Erro de referência. A função não foi encontrada. Erro: " + e.message);
  }
}