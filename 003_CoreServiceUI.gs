/**
 * Detecta se a execução está vindo do backend (webhook / UrlFetch)
 * ou de um usuário clicando na interface do Sheets.
 *
 * - Se conseguir acessar SpreadsheetApp.getUi() → contexto COM UI
 *   (retorna false, ou seja, NÃO é backend).
 * - Se der erro ao acessar UI → contexto SEM UI (backend/webhook)
 *   (retorna true).
 */
function isBackendExecution() {
  try {
    // Em contexto com UI isso funciona
    SpreadsheetApp.getUi();
    return false; // execução a partir da planilha (menu / gatilho com UI)
  } catch (e) {
    // Em contexto sem UI (chamada via UrlFetch/webhook) isso lança erro
    return true; // backend / webhook
  }
}


// ═══════════════════════════════════════════════════════════════
// INTERFACE PARA O MENU (MANTÉM COMPATIBILIDADE)
// ═══════════════════════════════════════════════════════════════

/**
 * Função chamada pelo menu da planilha
 * Serve como ponte entre o menu e o orquestrador
 */
function AtualizarPortfolio_Menu() {
  try {
    // 1. Alerta inicial apenas se for no Sheets
    if (!isBackendExecution()) {
    //  SpreadsheetApp.getActiveSpreadsheet().toast("Iniciando orquestração...", "Automação");
    }

    // 2. Executa a orquestração
    orquestrarAtualizacaoPortfolio();

    // 3. RETORNO OBRIGATÓRIO para o Dashboard ler
    return { 
      sucesso: true, 
      msg: "Orquestração enviada com sucesso!" 
    };

  } catch (e) {
    // Se algo falhar, retorna o erro como texto
    return { 
      sucesso: false, 
      msg: "Falha na execução: " + e.toString() 
    };
  }
}



/**
 * Função chamada pelo menu da planilha para Sync Dados Ativos
 * Serve como ponte entre o menu e o orquestrador
 */
function SyncDadosAtivos_Menu() {
  if (!isBackendExecution()) {
  //  SpreadsheetApp.getUi().alert("Sincronizando dados de ativos...");
  }
  orquestrarSyncDadosAtivos();
}

/**
 * Função chamada pelo menu da planilha para Sync Dados Detalhes
 * Serve como ponte entre o menu e o orquestrador
 */
function SyncDadosDetalhes_Menu() {
  if (!isBackendExecution()) {
  //  SpreadsheetApp.getUi().alert("Sincronizando dados detalhados...");
  }
  orquestrarSyncDadosDetalhes();
}

/**
 * Função chamada pelo menu da planilha para Sync Greeks
 * Serve como ponte entre o menu e o orquestrador
 */
function SyncGreeks_Menu() {
  if (!isBackendExecution()) {
  //  SpreadsheetApp.getUi().alert("Calculando gregas via API...");
  }
  orquestrarSyncGreeks();
}

/**
 * Função chamada pelo menu da planilha para Calc Greeks (Nativo)
 * Serve como ponte entre o menu e o orquestrador
 */
function CalcGreeks_Menu() {
  if (!isBackendExecution()) {
  //  SpreadsheetApp.getUi().alert("Calculando gregas (nativo)...");
  }
  orquestrarCalcGreeks();
}

// =====================================================================
// >>> ADIÇÃO — ITEM DO MENU PARA NOVO SERVIÇO
// =====================================================================

function SelecaoOpcoes_Menu() {
  if (!isBackendExecution()) {
  //  SpreadsheetApp.getUi().alert("Executando seleção de opções (DTE)...");
  }
  orquestrarSelecaoOpcoes();
}

/**
 * Função chamada pelo menu da planilha para o fluxo de Tendência
 */
function TendenciaDadosAtivos_Menu() {
  if (!isBackendExecution()) {
  //  SpreadsheetApp.getUi().alert("Iniciando análise de tendência (Sync + Calc)...");
  }
  // Chama a função que criamos no arquivo Orquestrador_TendenciaDadosAtivos.gs
  orquestrarFluxoTendencia();
}