/**
 * MÓDULO: SyncDadosAtivos - v1.5
 * CORREÇÃO: Alinhamento rigoroso com a estrutura de 19 colunas da folha.
 */

// A ordem aqui deve espelhar a sua folha Dados_Ativos da esquerda para a direita
const FIXED_HEADERS_ASSETS = [
  'Ticker', 
  'Timestamp_Atualizacao', 
  'name', 
  'sector', 
  'open', 
  'high', 
  'low', 
  'close', 
  'beta_ibov', 
  'iv_1y_rank', 
  'iv_current', 
  'variation', 
  'iv_1y_max', 
  'iv_1y_min', 
  'iv_1y_percentile', 
  'iv_6m_max', 
  'iv_6m_min', 
  'iv_6m_percentile', 
  'iv_6m_rank'
];

function sincronizarDadosAtivos() {
  const SERVICO_NOME = "SyncAtivos_v1.5";
  log(SERVICO_NOME, "INFO", "Iniciando sincronização com mapeamento de 19 colunas...", "");
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaNecton = ss.getSheetByName(ABA_GATILHO);
    const abaDadosAtivos = ss.getSheetByName(ABA_DADOS_ATIVOS);
    
    if (!abaNecton || !abaDadosAtivos) throw new Error("Abas não encontradas.");

    const ultimaLinhaNecton = abaNecton.getLastRow();
    if (ultimaLinhaNecton <= 1) return;

    // Obtém Tickers únicos da Coluna S (onde o 003 gravou o ativo objeto)
    const valoresNecton = abaNecton.getRange(2, COLUNA_OUTPUT_INICIO, ultimaLinhaNecton - 1, 1).getValues();
    const tickersUnicos = [...new Set(valoresNecton.flat().filter(t => t && t !== "ERRO_API" && t !== "N/A"))];
    
    if (tickersUnicos.length === 0) return;

    const { idToRowMap } = getAssetsMapAndHeaders(abaDadosAtivos);
    const timestamp = new Date().toISOString();
    const listaParaNovos = [];
    const updatesEmLote = {}; 

    tickersUnicos.forEach((ticker, i) => {
      const dadosAPI = getOpLabStockData(ticker); 
      
      if (dadosAPI) {
        // Mapeia os dados da API para as posições corretas das colunas
        const linhaValores = FIXED_HEADERS_ASSETS.map(campo => {
          if (campo === 'Ticker') return ticker;
          if (campo === 'Timestamp_Atualizacao') return timestamp;
          
          const valor = dadosAPI[campo];
          return valor !== undefined && valor !== null ? valor : "";
        });

        if (idToRowMap[ticker]) {
          updatesEmLote[idToRowMap[ticker]] = linhaValores;
        } else {
          listaParaNovos.push(linhaValores);
        }
      }
      if (i % 10 === 0) Utilities.sleep(100);
    });

    // Escrita em lote para evitar lentidão
    Object.keys(updatesEmLote).forEach(rowNum => {
      abaDadosAtivos.getRange(parseInt(rowNum), 1, 1, FIXED_HEADERS_ASSETS.length).setValues([updatesEmLote[rowNum]]);
    });

    if (listaParaNovos.length > 0) {
      abaDadosAtivos.getRange(abaDadosAtivos.getLastRow() + 1, 1, listaParaNovos.length, FIXED_HEADERS_ASSETS.length).setValues(listaParaNovos);
    }

    log(SERVICO_NOME, "SUCESSO", "Sincronização corrigida e concluída.", `Total: ${tickersUnicos.length}`);

  } catch (e) {
    log(SERVICO_NOME, "ERRO_FATAL", "Falha no Sync Ativos", e.toString());
  }
}

function getAssetsMapAndHeaders(aba) {
  const ultimaLinha = aba.getLastRow();
  const idToRowMap = {};
  if (ultimaLinha > 1) {
    const tickers = aba.getRange(2, 1, ultimaLinha - 1, 1).getValues();
    tickers.forEach((linha, index) => {
      if (linha[0]) idToRowMap[linha[0]] = index + 2;
    });
  }
  return { idToRowMap: idToRowMap };
}





// ═══════════════════════════════════════════════════════════════
// ORQUESTRADOR: SINCRONIZAR DADOS ATIVOS
// ═══════════════════════════════════════════════════════════════

/**
 * Orquestra a sincronização de dados de ativos
 * 
 * Este orquestrador:
 * 1. Valida pré-condições
 * 2. Registra início da operação
 * 3. Chama o serviço especializado (sincronizarDadosAtivos)
 * 4. Registra resultado e métricas
 * 5. Trata erros de forma centralizada
 * 
 * @returns {Object} Resultado da operação com métricas
 */
function orquestrarSyncDadosAtivos() {
  const OPERACAO_ID = Utilities.getUuid();
  const inicio = new Date();
  
  logOrquestrador("INICIO", "Iniciando orquestração de sincronização de dados ativos", {
    operacao_id: OPERACAO_ID,
    timestamp: inicio.toISOString(),
    servico: SERVICOS_REGISTRY.SYNC_DADOS_ATIVOS.nome
  });
  
  try {
    // --- ETAPA 1: Validação de Pré-condições ---
    const validacao = validarPreCondicoes();
    if (!validacao.sucesso) {
      throw new Error("Pré-condições não atendidas: " + validacao.mensagem);
    }
    
    logOrquestrador("INFO", "Pré-condições validadas com sucesso", validacao);
    
    // --- ETAPA 2: Execução do Serviço Especializado ---
    logOrquestrador("INFO", "Chamando serviço: " + SERVICOS_REGISTRY.SYNC_DADOS_ATIVOS.funcao, {
      servico: SERVICOS_REGISTRY.SYNC_DADOS_ATIVOS.nome
    });
    
    const resultadoServico = executarComTimeout(
      SERVICOS_REGISTRY.SYNC_DADOS_ATIVOS.funcao,
      SERVICOS_REGISTRY.SYNC_DADOS_ATIVOS.timeout
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
      servico: SERVICOS_REGISTRY.SYNC_DADOS_ATIVOS.nome,
      status: "SUCESSO"
    };
    
    logOrquestrador("SUCESSO", "Orquestração de sync dados ativos concluída com sucesso", metricas);
    
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
    
    logOrquestrador("ERRO_CRITICO", "Falha na orquestração de sync dados ativos", erroDetalhado);
    
    // Notifica usuário sobre o erro (seguro p/ backend)
    safeAlert_(
      '❌ Erro na Sincronização',
      'Ocorreu um erro durante a sincronização de dados ativos:\n\n' +
      erro.message +
      '\n\nVerifique a aba Logs para mais detalhes.'
    );
    
    return {
      sucesso: false,
      erro: erroDetalhado
    };
  }
}













function testarSyncAtivoUnitario() {
  console.log("--- INICIANDO TESTE UNITÁRIO 004: DATA STOCK ---");
  
  // Usaremos a CSAN3 como teste, que é o ativo objeto da CSANL670 que você testou antes
  const tickerTeste = "CSAN3";
  console.log("🚀 Buscando dados de mercado para: " + tickerTeste);
  
  const dados = getOpLabStockData(tickerTeste);
  
  if (dados && dados.close) {
    console.log("✅ SUCESSO: Dados recebidos via Client 009");
    console.log("📊 Preço Atual: R$ " + dados.close);
    console.log("📈 Volatilidade Implícita (IV): " + dados.iv_current + "%");
    console.log("🏢 Setor: " + dados.sector);
    
    // Validação de um campo que o FIXED_HEADERS_ASSETS costuma usar
    if (dados.name) {
      console.log("🏆 Nome da Empresa: " + dados.name);
    }
  } else {
    console.error("❌ FALHA: O Client 009 não retornou dados válidos para o ativo.");
    console.warn("Dica: Verifique se o ticker '" + tickerTeste + "' está correto e se o token tem permissão para Stocks.");
  }
  
  console.log("--- FIM DO TESTE ---");
}