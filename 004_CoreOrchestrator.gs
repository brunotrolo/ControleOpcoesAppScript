// ═══════════════════════════════════════════════════════════════
// SERVIÇOS DISPONÍVEIS (REGISTRY)
// ═══════════════════════════════════════════════════════════════

/**
 * Verifica se a Interface de Usuário (UI) está disponível.
 */
function isUiAvailable() {
  try {
    // Tenta acessar a UI, se falhar, retorna false
    return !!SpreadsheetApp.getUi();
  } catch (e) {
    return false;
  }
}

/**
 * Notifica o usuário de forma segura, seja via UI ou via Log.
 * Blindado para Web App.
 */
function notificarUsuario_(titulo, mensagem) {
  try {
    // Se NÃO for backend (ou seja, se tiver UI), usa alert ou toast
    if (isUiAvailable()) {
      // Toast é menos intrusivo que Alert para Web Apps
      SpreadsheetApp.getActiveSpreadsheet().toast(mensagem, titulo);
    } else {
      // No Web App, apenas loga no console do servidor (Apps Script Dashboard)
      console.warn("LOG INTERNO [" + titulo + "]: " + mensagem);
    }
  } catch (e) {
    console.error("Erro ao tentar notificar: " + e.message);
  }
}



const SERVICOS_REGISTRY = {
  ATUALIZACAO_PORTFOLIO: {
    nome: "AtualizacaoPortfolio",
    funcao: "FluxoMestreDeControle",
    descricao: "Atualiza portfolio completo via API OpLab",
    timeout: 300000,
    requer_token: true
  },
  SYNC_DADOS_ATIVOS: {
    nome: "SyncDadosAtivos",
    funcao: "sincronizarDadosAtivos",
    descricao: "Sincroniza dados de ativos via API OpLab (/stocks)",
    timeout: 300000,
    requer_token: true
  },
  SYNC_DADOS_DETALHES: {
    nome: "SyncDadosDetalhes",
    funcao: "sincronizarDadosDetalhes",
    descricao: "Sincroniza dados detalhados de opções via API OpLab (/details)",
    timeout: 300000,
    requer_token: true
  },
  SYNC_GREEKS: {
    nome: "SyncGreeks",
    funcao: "sincronizarGreeks",
    descricao: "Calcula gregas de opções via API OpLab Black-Scholes",
    timeout: 300000,
    requer_token: true
  },
  CALC_GREEKS: {
    nome: "CalcGreeks",
    funcao: "calcularGreeksNativo",
    descricao: "Calcula gregas usando Black-Scholes implementado nativamente",
    timeout: 300000,
    requer_token: false
  },
  SYNC_TENDENCIA: {
    nome: "TendenciaDadosAtivos",
    funcao: "orquestrarFluxoTendencia",
    descricao: "Sincroniza histórico e calcula tendências (Médias/Bollinger)",
    timeout: 600000, // 10 minutos (mais tempo para série histórica)
    requer_token: true
  },

  // =====================================================================
  // >>> ADIÇÃO — NOVO SERVIÇO: SELEÇÃO DE OPÇÕES (DTE)
  // =====================================================================
  SELECAO_OPCOES: {
    nome: "SelecaoOpcoes",
    funcao: "buscarOpcoesParaSelecao",
    descricao: "Busca opções com filtro de DTE para seleção de estratégias",
    timeout: 300000,
    requer_token: true
  }
  // Aqui você poderá adicionar novos serviços no futuro:
  // ANALISE_RISCO: { ... }
  // CALCULO_GREGAS: { ... }
  // etc.
};



// ═══════════════════════════════════════════════════════════════
// FUNÇÕES AUXILIARES DO ORQUESTRADOR
// ═══════════════════════════════════════════════════════════════

/**
 * Valida se todas as pré-condições estão atendidas
 * 
 * @returns {Object} Resultado da validação
 */
function validarPreCondicoes() {
  const resultado = {
    sucesso: true,
    mensagem: "",
    detalhes: {}
  };
  
  try {
    // Validação 1: Planilha ativa
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    if (!planilha) {
      resultado.sucesso = false;
      resultado.mensagem = "Nenhuma planilha ativa encontrada";
      return resultado;
    }
    resultado.detalhes.planilha = planilha.getName();
    
    // Validação 2: Aba principal existe
    const abaGatilho = planilha.getSheetByName(ABA_GATILHO);
    if (!abaGatilho) {
      resultado.sucesso = false;
      resultado.mensagem = "Aba '" + ABA_GATILHO + "' não encontrada";
      return resultado;
    }
    resultado.detalhes.aba_principal = ABA_GATILHO;
    
    // Validação 3: Aba de logs existe
    const abaLogs = planilha.getSheetByName(ABA_LOGS);
    if (!abaLogs) {
      resultado.sucesso = false;
      resultado.mensagem = "Aba '" + ABA_LOGS + "' não encontrada";
      return resultado;
    }
    resultado.detalhes.aba_logs = ABA_LOGS;
    
    // Validação 4: Token configurado
    const token = PropertiesService.getScriptProperties().getProperty(OPLAB_TOKEN_NAME);
    if (!token) {
      resultado.sucesso = false;
      resultado.mensagem = "Token OPLAB não configurado nas Propriedades do Script";
      return resultado;
    }
    resultado.detalhes.token_configurado = true;
    
    // Validação 5: Há dados para processar
    const ultimaLinha = abaGatilho.getLastRow();
    if (ultimaLinha <= 1) {
      resultado.sucesso = false;
      resultado.mensagem = "Nenhum dado encontrado para processar (apenas cabeçalho)";
      return resultado;
    }
    resultado.detalhes.linhas_disponiveis = ultimaLinha - 1;
    
    resultado.mensagem = "Todas as pré-condições atendidas";
    return resultado;
    
  } catch (erro) {
    resultado.sucesso = false;
    resultado.mensagem = "Erro ao validar pré-condições: " + erro.message;
    return resultado;
  }
}



/**
 * Executa uma função com timeout e retorna um objeto de status
 * Blindado para não estourar erro bruto no Web App
 */
function executarComTimeout(nomeFuncao, timeout) {
  const inicio = new Date();
  
  try {
    const funcao = this[nomeFuncao];
    if (typeof funcao !== 'function') {
      throw new Error("Função '" + nomeFuncao + "' não encontrada");
    }
    
    // Executa a função
    const resultado = funcao();
    
    const duracao = new Date() - inicio;
    
    logOrquestrador("INFO", "Serviço executado com sucesso", {
      funcao: nomeFuncao,
      duracao_ms: duracao
    });
    
    return { sucesso: true, dado: resultado };
    
  } catch (erro) {
    // Loga o erro internamente
    logOrquestrador("ERRO", "Falha na execução da função: " + nomeFuncao, { erro: erro.message });
    
    // Em vez de dar THROW (que mata o Web App), retornamos um objeto de erro
    return { sucesso: false, erro: erro.message };
  }
}





/**
 * Função de diagnóstico - verifica saúde do sistema
 * 
 * @returns {Object} Status de saúde do sistema
 */
function diagnosticoSistema() {
  const diagnostico = {
    timestamp: new Date().toISOString(),
    sistema_operacional: true,
    problemas: []
  };
  
  try {
    // Testa conexão com planilha
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    diagnostico.planilha_ativa = !!planilha;
    
    // Testa abas necessárias
    diagnostico.aba_gatilho = !!planilha.getSheetByName(ABA_GATILHO);
    diagnostico.aba_logs = !!planilha.getSheetByName(ABA_LOGS);
    
    // Testa token
    diagnostico.token_configurado = !!PropertiesService.getScriptProperties().getProperty(OPLAB_TOKEN_NAME);
    
    // Identifica problemas
    if (!diagnostico.planilha_ativa) diagnostico.problemas.push("Planilha não está ativa");
    if (!diagnostico.aba_gatilho) diagnostico.problemas.push("Aba '" + ABA_GATILHO + "' não encontrada");
    if (!diagnostico.aba_logs) diagnostico.problemas.push("Aba '" + ABA_LOGS + "' não encontrada");
    if (!diagnostico.token_configurado) diagnostico.problemas.push("Token OPLAB não configurado");
    
    diagnostico.sistema_operacional = diagnostico.problemas.length === 0;
    
  } catch (erro) {
    diagnostico.sistema_operacional = false;
    diagnostico.problemas.push("Erro crítico: " + erro.message);
  }
  
  return diagnostico;
}




// ═══════════════════════════════════════════════════════════════
// FUNÇÕES UTILITÁRIAS DO ORQUESTRADOR
// ═══════════════════════════════════════════════════════════════

/**
 * Retorna informações sobre o orquestrador e serviços disponíveis
 * Útil para debugging e documentação
 * 
 * @returns {Object} Informações do sistema
 */
function obterInformacoesOrquestrador() {
  return {
    versao: CONFIG_ORQUESTRADOR.versao,
    ambiente: CONFIG_ORQUESTRADOR.ambiente,
    servicos_disponiveis: Object.keys(SERVICOS_REGISTRY).map(key => ({
      chave: key,
      nome: SERVICOS_REGISTRY[key].nome,
      funcao: SERVICOS_REGISTRY[key].funcao,
      descricao: SERVICOS_REGISTRY[key].descricao
    })),
    timestamp: new Date().toISOString()
  };
}



/**
 * Gatilho diário blindado
 */
function gatilhoDiarioStrangles() {
  const SERVICO = "GatilhoDiarioStrangles";
  log(SERVICO, "INFO", "Início do gatilho diário (08:00)", "");

  try {
    rodarPipelineStranglesComAlerta();
    log(SERVICO, "SUCESSO", "Gatilho diário concluído.", "");
    return { sucesso: true, msg: "Gatilho concluído" };

  } catch (e) {
    // CORREÇÃO: Usando 'e.message' consistentemente para evitar ReferenceError
    log(SERVICO, "ERRO_CRITICO", "Falha no gatilho", e.message);
    
    // CORREÇÃO: Passando os parâmetros REAIS para a função
    notificarUsuario_("❌ Erro Crítico", "Falha na orquestração diária: " + e.message);
    
    return { sucesso: false, erro: e.message };
  }
}



function testarPasso3() {
  console.log("--- TESTANDO MIGRAÇÃO PASSO 3 (ORQUESTRADOR) ---");
  
  try {
    // 1. Validar se o Registry está visível
    const servico = SERVICOS_REGISTRY.SYNC_GREEKS.nome;
    console.log("✅ Sucesso: Registry acessível. Exemplo: " + servico);
    
    // 2. Rodar diagnóstico do sistema que agora está no 003
    const diag = diagnosticoSistema();
    console.log("✅ Sucesso: Diagnóstico executado. Status: " + (diag.sistema_operacional ? "SAUDÁVEL" : "COM PROBLEMAS"));
    
    console.log("--- SE O DIAGNÓSTICO RODOU, O PASSO 3 FOI SUCESSO ---");
  } catch (e) {
    console.error("❌ FALHA: Erro no Orquestrador. Verifique os recortes. Erro: " + e.message);
  }
}

/**
 * ORQUESTRADOR DINÂMICO v2.0
 * Executa uma sequência de funções definida na Config_Global.
 */
function executarFluxoSequencial() {
  const servicoNome = "Orquestrador_Master";
  const configs = obterConfigsGlobais(); // Sua função que lê a Config_Global
  const sequenciaStr = configs['Orquestrador_Sequencia_Padrao'];

  if (!sequenciaStr) {
    gravarLog(servicoNome, "ERRO_CONFIG", "Sequência não encontrada", "Preencha 'Orquestrador_Sequencia_Padrao' na Config_Global");
    return;
  }

  // 1. Limpa e separa as funções (aceita vírgula ou ponto e vírgula)
  const listaFuncoes = sequenciaStr.split(/[;,]/).map(f => f.trim());

  gravarLog(servicoNome, "START", "Iniciando Fluxo Sequencial", "Passos: " + listaFuncoes.join(" -> "));

  listaFuncoes.forEach((nomeFuncao, index) => {
    const passo = `[Passo ${index + 1}/${listaFuncoes.length}]`;
    
    try {
      // 2. Verifica se a função existe no escopo global do projeto
      if (typeof this[nomeFuncao] === "function") {
        gravarLog(servicoNome, "INFO", `${passo} Executando: ${nomeFuncao}`, "");
        
        // EXECUÇÃO REAL
        this[nomeFuncao](); 
        
        gravarLog(servicoNome, "SUCESSO", `${passo} Finalizado: ${nomeFuncao}`, "");
      } else {
        throw new Error(`Função '${nomeFuncao}' não existe no script.`);
      }
    } catch (e) {
      gravarLog(servicoNome, "ERRO_CRITICO", `Falha no ${passo}: ${nomeFuncao}`, e.toString());
      // Opcional: interromper o fluxo se um passo falhar
      throw new Error(`Interrupção do fluxo: Falha em ${nomeFuncao}`);
    }
  });

  gravarLog(servicoNome, "FINISH", "Fluxo Master concluído", "");
}