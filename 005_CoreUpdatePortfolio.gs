/**
 * ARQUITETURA V28 - OTIMIZADA
 * MÓDULO: CoreUpdatePortfolio
 * AÇÃO: Orquestra a atualização dos dados das opções (S, T, U, V) consumindo o Client 009.
 * MANDATO: NUNCA tocar nas colunas A-R. Sem coloração.
 */


/**
 * Notifica o usuário de forma segura.
 */
function notificarUsuario_(titulo, mensagem) {
  try {
    if (isUiAvailable()) {
      // CORREÇÃO: Usar as variáveis 'titulo' e 'mensagem', não strings fixas
      SpreadsheetApp.getUi().alert(titulo, mensagem, SpreadsheetApp.getUi().ButtonSet.OK);
    } else {
      console.warn("LOG WEB APP [" + titulo + "]: " + mensagem);
    }
  } catch (e) {
    console.error("Falha ao notificar: " + e.message);
  }
}


// -------------------------------------------------------------------
// FLUXO MESTRE (V28)
// -------------------------------------------------------------------

/**
 * Processa a planilha INTEIRA para enriquecer dados faltantes.
 * Chamado manualmente pelo menu.
 */
function FluxoMestreDeControle() {
  const SERVICO_NOME = "FluxoMestre_v28";
  log(SERVICO_NOME, "INFO", "Iniciando processamento em lote (Manual)...", "");
  
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const aba = planilha.getSheetByName(ABA_GATILHO);
  
  const range = aba.getDataRange();
  const valores = range.getValues();
  
  let linhasProcessadas = 0;
  let linhasComErro = 0;

  // Começa da linha 2 (i=1) para pular o cabeçalho
  for (let i = 1; i < valores.length; i++) {
    const linha = i + 1;
    const dadosLinha = valores[i];
    
    try {
      const tickerOpcao = dadosLinha[COLUNA_GATILHO_INPUT - 1]; // Coluna A
      const idTrade = dadosLinha[COLUNA_ID_FORMULA - 1];       // Coluna P
      const jaProcessado = dadosLinha[COLUNA_OUTPUT_INICIO - 1]; // Coluna S
      
      // FILTROS: Só processa se tiver Ticker e ID, e se S ainda estiver vazio
      if (tickerOpcao && idTrade && !jaProcessado) {
        log(SERVICO_NOME, "INFO", "Processando linha: " + linha + " | Ticker: " + tickerOpcao, "");
        
        // Agora o "Motor" utiliza o Client 009 internamente
        const sucesso = enriquecerLinha(aba, linha, tickerOpcao, SERVICO_NOME);
        
        if (sucesso) {
          linhasProcessadas++;
        } else {
          linhasComErro++;
        }
        
        // Respeita o rate limit da API
        Utilities.sleep(1000);
      }
    } catch (e) {
      log(SERVICO_NOME, "ERRO_CRITICO", "Erro no loop mestre, linha " + (i+1), e.message);
      linhasComErro++;
    }
  }
  
  // LOG de finalização sem mencionar coloração
  log(SERVICO_NOME, "SUCESSO", "Processamento concluído.", "Total: " + linhasProcessadas + " enriquecidas. Erros: " + linhasComErro);
  
  let mensagemFinal = linhasProcessadas + ' novas linhas foram enriquecidas com sucesso!';
  if (linhasComErro > 0) {
    mensagemFinal += '\n\n⚠️ ' + linhasComErro + ' linhas tiveram erro (verifique a aba Logs)';
  }
  
  // notificarUsuario_('Processamento Concluído!', mensagemFinal);
}

// -------------------------------------------------------------------
// O "MOTOR" ENRIQUECEDOR - INTEGRADO AO CLIENT 009
// -------------------------------------------------------------------

/**
 * Busca dados no Client 009 e escreve nas colunas S, T, U, V, W.
 */
function enriquecerLinha(aba, linha, tickerOpcao, servicoPai) {
  const SERVICO_NOME = "Motor_Enriquecer_v28_Ajustado";
  
  try {
    const data = getOpLabOptionDetails(tickerOpcao);

    if (!data) {
      // Se falhar, preenche as 5 colunas com o erro (S a W)
      aba.getRange(linha, COLUNA_OUTPUT_INICIO, 1, 5).setValues([["ERRO_API", "", "", "", ""]]);
      return false;
    }
    
    const parent_symbol = data.parent_symbol || data.symbol || data.underlying || "N/A";
    const vencimento = data.due_date || data.expiration || data.expiry_date || "N/A";
    
    let strike = "N/A";
    if (data.strike !== undefined && data.strike !== null && data.strike !== "") {
      strike = Number(data.strike);
      if (isNaN(strike)) strike = "N/A";
    }
    
    const categoria = data.category || data.type || data.option_type || "N/A";

    // --- NOVA REGRA: Mapeamento de 5 colunas ---
    const valoresParaEscrever = [[
      String(parent_symbol), // S - Ticker Objeto
      String(vencimento),    // T - Vencimento
      strike,                // U - Strike
      String(categoria),     // V - Categoria (CALL/PUT)
      "ATIVO",               // W - Status (Sua nova regra)
    ]];
    
    // Escrita em lote (S a W)
    aba.getRange(linha, COLUNA_OUTPUT_INICIO, 1, 5).setValues(valoresParaEscrever);
    
    log(servicoPai, "DEBUG", "Linha enriquecida com Status ATIVO", "Ticker: " + tickerOpcao);
    return true;

  } catch (erro) {
    log(servicoPai, "ERRO_SCRIPT", "Falha fatal no motor: " + tickerOpcao, erro.message);
    // Trata erro preenchendo as 5 colunas
    aba.getRange(linha, COLUNA_OUTPUT_INICIO, 1, 5).setValues([["ERRO_SCRIPT", erro.message.substring(0, 30), "", "", ""]]);
    return false;
  }
}



// ═══════════════════════════════════════════════════════════════
// ORQUESTRADOR: ATUALIZAR PORTFÓLIO
// ═══════════════════════════════════════════════════════════════

/**
 * Orquestra a atualização completa do portfólio
 * 
 * Este é o ponto de entrada principal que:
 * 1. Valida pré-condições
 * 2. Registra início da operação
 * 3. Chama o serviço especializado
 * 4. Registra resultado e métricas
 * 5. Trata erros de forma centralizada
 * 
 * @returns {Object} Resultado da operação com métricas
 */
function orquestrarAtualizacaoPortfolio() {
  const OPERACAO_ID = Utilities.getUuid();
  const inicio = new Date();
  
  logOrquestrador("INICIO", "Iniciando orquestração de atualização de portfólio", {
    operacao_id: OPERACAO_ID,
    timestamp: inicio.toISOString(),
    servico: SERVICOS_REGISTRY.ATUALIZACAO_PORTFOLIO.nome
  });
  
  try {
    // --- ETAPA 1: Validação de Pré-condições ---
    const validacao = validarPreCondicoes();
    if (!validacao.sucesso) {
      throw new Error("Pré-condições não atendidas: " + validacao.mensagem);
    }
    
    logOrquestrador("INFO", "Pré-condições validadas com sucesso", validacao);
    
    // --- ETAPA 2: Execução do Serviço Especializado ---
    logOrquestrador("INFO", "Chamando serviço: " + SERVICOS_REGISTRY.ATUALIZACAO_PORTFOLIO.funcao, {
      servico: SERVICOS_REGISTRY.ATUALIZACAO_PORTFOLIO.nome
    });
    
    const resultadoServico = executarComTimeout(
      SERVICOS_REGISTRY.ATUALIZACAO_PORTFOLIO.funcao,
      SERVICOS_REGISTRY.ATUALIZACAO_PORTFOLIO.timeout
    );
    
    // --- ETAPA 3: Registro de Métricas ---
    const fim = new Date();
    const duracao = fim - inicio;
    
    const metricas = {
      operacao_id: OPERACAO_ID,
      duracao_ms: duracao,
      duracao_segundos: (duracao / 1000).toFixed(2),
      timestamp_inicio: inicio.toISOString(),
      timestamp_fim: fim.toISOString(),
      servico: SERVICOS_REGISTRY.ATUALIZACAO_PORTFOLIO.nome,
      status: "SUCESSO"
    };
    
    logOrquestrador("SUCESSO", "Orquestração concluída com sucesso", metricas);
    
    return {
      sucesso: true,
      metricas: metricas,
      resultado: resultadoServico
    };
    
  } catch (erro) {
    // --- ETAPA 4: Tratamento Centralizado de Erros ---
    const fim = new Date();
    const duracao = fim - inicio;
    
    const erroDetalhado = {
      operacao_id: OPERACAO_ID,
      duracao_ms: duracao,
      mensagem: erro.message,
      stack: erro.stack,
      timestamp: fim.toISOString()
    };
    
    logOrquestrador("ERRO_CRITICO", "Falha na orquestração", erroDetalhado);
    
      // CORREÇÃO: Usar o nome correto da função de notificação
      // notificarUsuario_( '❌ Erro na Atualização','Ocorreu um erro durante a atualização do portfólio:\n\n' + erro.message);
      
      return { 
        sucesso: false, 
        erro: erroDetalhado 
        };
    }
  }

function testarEnriquecimentoUnitario() {
  console.log("--- INICIANDO TESTE UNITÁRIO 003 + 009 (v1.3) ---");
  
  // 1. Configuração do teste
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(ABA_GATILHO);
  const LINHA_TESTE = 2; 
  const tickerParaTeste = "BRKMO780"; // Forçando o ticker para isolar o problema
  
  console.log("🚀 Alvo: " + tickerParaTeste + " na linha: " + LINHA_TESTE);
  
  // 2. Execução
  const resultado = enriquecerLinha(aba, LINHA_TESTE, tickerParaTeste, "TESTE_DIAGNOSTICO");
  
  // 3. Força o descarregamento dos logs para podermos ler na planilha
  if (typeof flushLogs === 'function') flushLogs();

  if (resultado) {
    const dados = aba.getRange(LINHA_TESTE, COLUNA_OUTPUT_INICIO, 1, 4).getValues()[0];
    console.log("✅ SUCESSO! Dados na planilha: " + JSON.stringify(dados));
  } else {
    console.error("❌ O teste falhou. Vá até a aba 'Logs' AGORA e veja a última linha.");
  }
  
  console.log("--- FIM DO TESTE ---");
}