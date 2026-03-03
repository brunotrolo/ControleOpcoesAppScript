/**
 * ORQUESTRADOR DE INTEGRAÇÃO - SÉRIE 600
 * Objetivo: Rodar toda a cadeia preditiva em sequência.
 */

/**
 * Função de atalho para testes manuais no editor de script.
 * Executa o pipeline com o horizonte padrão de 45 dias.
 */
function testarPipelinePreditivoCompleto() {
  executarPipelinePreditivo(45);
}

/**
 * Executa o pipeline completo com horizonte de tempo dinâmico.
 * Essa função será o ponto de entrada principal quando o Web App acionar o Backend.
 * * @param {number} diasParam - Horizonte de tempo em dias (ex: 5, 21, 45)
 */
function executarPipelinePreditivo(diasParam) {
  const servico = "600_Pipeline_Integracao";
  const dias = diasParam || 45; // Garante um fallback de segurança
  
  log(servico, "START", `Iniciando homologação completa 601 -> 605 (Horizonte: ${dias} dias)`, "");

  try {
    // Passo 1: Sincronizar Macro (601) - Monitoramento de clima de mercado
    log(servico, "INFO", "Executando 601_SyncMacroData...", "");
    syncMacroData_Execute(); 
    SpreadsheetApp.flush(); // Garante que as fórmulas GOOGLEFINANCE carreguem

    // Passo 2: Processar Estatísticas (602) - Motor de Sigmas dinâmico
    log(servico, "INFO", `Executando 602_ProcessStatsPredictive (${dias} dias)...`, "");
    processStatsPredictive_Execute(dias);
    SpreadsheetApp.flush(); // Garante que a aba Analise_Estatistica esteja pronta

    // Passo 3: Processar Fundamentos (603) - Motor de Graham e Saúde Financeira
    log(servico, "INFO", "Executando 603_ProcessFundamentalPredictive...", "");
    processFundamentalPredictive_Execute();
    SpreadsheetApp.flush(); // Garante que a aba Analise_Fundamentalista esteja pronta

    // Passo 4: Consolidar Score (604) - Cérebro Preditivo e Veredito
    log(servico, "INFO", `Executando 604_OrquestratorPredictiveScore (${dias} dias)...`, "");
    orquestratorPredictiveScore_Execute(dias);
    SpreadsheetApp.flush();

    // Passo 5: Gerar Dashboard Visual (605) - Espelhamento para leitura humana
    log(servico, "INFO", `Executando 605_Analise_Preditiva_Heatmap (${dias} dias)...`, "");
    gerarAnalisePreditivaHeatmap(dias);
    SpreadsheetApp.flush();

    log(servico, "SUCESSO", `Pipeline finalizado com integridade de dados para ${dias} dias.`, "");

    flushLogs(); // <--- ADICIONE ESTA LINHA AQUI
    
  } catch (e) {
    log(servico, "ERRO_FATAL", "Falha na integração do pipeline", e.toString());

    flushLogs(); // <--- E ADICIONE ESTA LINHA AQUI TAMBÉM
  }
}