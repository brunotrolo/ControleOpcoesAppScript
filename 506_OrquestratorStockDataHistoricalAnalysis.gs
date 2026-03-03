/**
 * ═══════════════════════════════════════════════════════════════
 * ORQUESTRADOR - TENDÊNCIA E SCANNER DE OPORTUNIDADES (PADRÃO NEXO)
 * ═══════════════════════════════════════════════════════════════
 * VERSÃO: 2.0 (Sync + Calc + Scanner)
 * DATA: 2026-02-08
 * ═══════════════════════════════════════════════════════════════
 */

function orquestrarFluxoTendencia() {
  const OPERACAO_ID = Utilities.getUuid();
  const inicio = new Date();
  const SERVICO_NOME = "TendenciaScanner_Master";
  
  // 1. Log de Início (Telemetria Central)
  logOrquestrador("INICIO", "Iniciando orquestração completa: Tendência + Scanner", {
    operacao_id: OPERACAO_ID,
    servico: SERVICO_NOME,
    timestamp: inicio.toISOString()
  });

  try {
    // 2. Validação de Pré-condições
    const validacao = validarPreCondicoes();
    if (!validacao.sucesso) {
      throw new Error("Pré-condições não atendidas: " + validacao.mensagem);
    }

    // 3. PASSO 1: Sincronização Histórica
    logOrquestrador("INFO", "Passo 1/5: Sincronizando série histórica de 250 dias", {
      operacao_id: OPERACAO_ID,
      funcao: "sincronizarAtivosHistoricos"
    });
    
    sincronizarAtivosHistoricos();
    
    // Pequeno delay para estabilidade de escrita
    Utilities.sleep(2000);

    // 4. PASSO 2: Cálculo de Indicadores Técnicos
    logOrquestrador("INFO", "Passo 2/5: Calculando Médias, IFR, Bandas e Vereditos", {
      operacao_id: OPERACAO_ID,
      funcao: "calcularTendenciaMercado"
    });
    
    calcularTendenciaMercado();

    // Pequeno delay para estabilidade antes do cruzamento
    Utilities.sleep(1000);

    // 5. PASSO 3: Scanner de Oportunidades (Cruzamento Técnico + Macro)
    logOrquestrador("INFO", "Passo 3/5: Gerando dados técnicos e Score Nexo", { 
      operacao_id: OPERACAO_ID,
      funcao: "executarScannerOportunidades"
    });
    
    // Primeiro executamos o scanner para preparar a planilha
    executarScannerOportunidades();
    
    // Força o salvamento dos dados e espera a planilha "respirar"
    SpreadsheetApp.flush();
    Utilities.sleep(3000);


    // 6. PASSO 4: Consultoria IA (O Veredito Final)
    logOrquestrador("INFO", "Passo 4/5: Enviando Top 9 para análise estratégica do Gemini", { 
      operacao_id: OPERACAO_ID,
      funcao: "gerarAnaliseIA_Oportunidades"
    });
    
    gerarAnaliseIA_Oportunidades();

    // 🛑 PONTO CRÍTICO: Garante que a IA escreveu na planilha antes do e-mail ler
    SpreadsheetApp.flush(); 
    Utilities.sleep(3000); // 3 segundos para o Google "respirar"

    // 7. PASSO 5: Disparo de E-mail (Relatório Institucional)
    logOrquestrador("INFO", "Passo 5/5: Gerando e enviando Relatório Nexo Master", { 
      operacao_id: OPERACAO_ID,
      funcao: "enviarRelatorioNexoMaster"
    });

    enviarRelatorioNexoMaster();

    // 8. Registro de Sucesso e Métricas Finais
    const fim = new Date();
    const duracao = fim - inicio;
    
    const metricas = {
      operacao_id: OPERACAO_ID,
      duracao_segundos: (duracao / 1000).toFixed(2),
      timestamp_fim: fim.toISOString(),
      servico: SERVICO_NOME,
      status: "SUCESSO"
    };
    
    logOrquestrador("SUCESSO", "Fluxo Master (Scanner + IA + Email) concluído", metricas);
    
    // Notificação visual atualizada
    //safeAlert_('✅ Inteligência Nexo & Report Enviados', 
    //  `Fluxo completo finalizado com sucesso!\n\n` +
    //  `⏱️ Duração: ${metricas.duracao_segundos}s\n` +
    //  `📩 E-mail Institucional disparado.\n` +
    //  `📊 Planilha atualizada com Vereditos IA.`
    // );

  } catch (erro) {


    // 8. Tratamento de Erros Centralizado
    const fim = new Date();
    const erroDetalhado = {
      operacao_id: OPERACAO_ID,
      mensagem: erro.message,
      stack: erro.stack,
      timestamp: fim.toISOString()
    };
    
    logOrquestrador("ERRO_CRITICO", "Falha na orquestração do Fluxo Master", erroDetalhado);
    safeAlert_('❌ Erro no Sistema', 'Falha no processo: ' + erro.message);
  }
}