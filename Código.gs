/*
 * ═══════════════════════════════════════════════════════════════
 * CÓDIGO.GS - ORQUESTRADOR PRINCIPAL
 * ═══════════════════════════════════════════════════════════════
 * 
 * ARQUIVO PRINCIPAL - Este deve ser o primeiro arquivo do projeto!
 * 
 * RESPONSABILIDADES:
 * - Criar menu (onOpen) - EXECUTADO AUTOMATICAMENTE PELO GOOGLE SHEETS
 * 
 * VERSÃO: 1.0
 * DATA: 2025-11-13
 * 
 * ═══════════════════════════════════════════════════════════════
 */
executarFluxoSequencial

/**
 * Cria o menu "⚙️ Automação" quando a planilha é aberta
 */
function onOpen(e) {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('⚙️ Automação')
        .addItem('🧠 Atualizar Cockpit ', 'executarFluxoSequencial')
        .addSeparator()
        .addItem('📊 Atualizar Portfólio', 'AtualizarPortfolio_Menu')
        .addItem('📈 Sincronizar Dados Ativos', 'SyncDadosAtivos_Menu')
        .addItem('🕰️ Sincronizar Dados Ativos Históricos', 'sincronizarAtivosHistoricos')
        .addItem('🔍 Sincronizar Dados Detalhes', 'SyncDadosDetalhes_Menu')        
        .addSeparator()
        .addItem('🧮 Calcular Gregas (API)', 'SyncGreeks_Menu')
        .addItem('🔬 Calcular Gregas (Nativo)', 'CalcGreeks_Menu')
        .addItem('🔮 Calcular Análise Preditiva', 'executarPipelinePreditivo')
        .addSeparator()
        .addItem('🔎 Seleção de Opções (DTE)', 'SelecaoOpcoes_Menu')
        .addItem('🛡️ Gerar Short Strangles', "orq_GerarShortStrangles")
        .addItem('🧭 Gerar Consultoria Portifólio', "executarRotinaDiaria")
        .addItem('🔍 Gerar Scanner de Oportunidades', "orquestrarScannerComEmail")
        .addItem('📉 Gerar Analise Tendência (250d)', 'TendenciaDadosAtivos_Menu')
        .addToUi();
  } catch (err) {
    // Apenas registra no Logger interno se falhar ao criar o menu (comum em aberturas rápidas/mobile)
    console.warn("Não foi possível carregar a UI do menu: " + err.message);
  }
}


/**
 * Homologação de Carregamento de Arquivos
 */

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Stock Options | Intelligence')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Esta função é vital para o sistema de slots e componentes
function include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (e) {
    return ``;
  }
}




function testeFinalIntegridade() {
  console.log("--- INICIANDO TESTE FINAL DE ARQUITETURA ---");
  
  try {
    // 1. Testa se o motor (Core) acha as configurações
    console.log("🔍 Verificando Motor: " + CONFIG_ORQUESTRADOR.versao);
    
    // 2. Testa se o motor acha o Logger
    log("SISTEMA", "INFO", "Teste de integridade pós-fatiamento", "Sucesso");
    console.log("✅ Logger operacional.");

    // 3. Testa se o motor acha o Orquestrador
    const info = obterInformacoesOrquestrador();
    console.log("✅ Orquestrador operacional. Serviços mapeados: " + info.servicos_disponiveis.length);

    console.log("--- SISTEMA SAUDÁVEL E TOTALMENTE FATIADO ---");
  } catch (e) {
    console.error("❌ ERRO DE INTEGRIDADE: " + e.message);
  }
}

function FORCAR_AUTORIZACAO_GMAIL() {
  GmailApp.sendEmail("teste@exemplo.com", "Teste", "Forçando janela");
}