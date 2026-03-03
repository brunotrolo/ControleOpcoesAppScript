/**
 * MÓDULO: SyncDadosDetalhes - v1.5
 * CORREÇÃO: Alinhamento de colunas de origem (ID na Col Q, Estrutura na Col R).
 */

const FIXED_HEADERS_DETAILS = [
  'ID_Trade_Unico', 'ID_Estrutura', 'Timestamp_Atualizacao', 'symbol', 'name',
  'open', 'high', 'low', 'close', 'volume', 'financial_volume', 'trades', 'bid', 'ask',
  'parent_symbol', 'category', 'due_date', 'maturity_type', 'strike', 'contract_size',
  'exchange_id', 'created_at', 'updated_at', 'variation', 'spot_price', 'isin',
  'security_category', 'market_maker', 'block_date', 'days_to_maturity', 'cnpj',
  'bid_volume', 'ask_volume', 'time', 'type', 'last_trade_at', 'strike_eod',
  'quotationForm', 'lastUpdatedDividendsAt'
];

function sincronizarDadosDetalhes() {
  const SERVICO_NOME = "SyncDetalhes_v1.5";
  log(SERVICO_NOME, "INFO", "Iniciando sincronização com correção de mapeamento...", "");
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaNecton = ss.getSheetByName(ABA_GATILHO);
    const abaDadosDetalhes = ss.getSheetByName(ABA_DADOS_DETALHES);
    
    if (!abaNecton || !abaDadosDetalhes) throw new Error("Abas não encontradas.");

    const ultimaLinha = abaNecton.getLastRow();
    if (ultimaLinha <= 1) return;
    
    // MAPEAMENTO RÍGIDO:
    const colTicker      = 1;  // Coluna A
    const colIdTrade     = 17; // Coluna Q (O ID Único concatenado)
    const colIdEstrutura = 18; // Coluna R (O ID da Estrutura)
    const colVencimento  = 20; // Coluna T (Data de Vencimento)

    const valoresNecton = abaNecton.getRange(2, 1, ultimaLinha - 1, colVencimento).getValues();
    const { idToRowMap } = getDetailsMapAndHeaders(abaDadosDetalhes);
    
    const timestamp = new Date().toISOString();
    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0); // Zera as horas para comparar apenas os dias

    const listaParaNovos = [];
    const updatesEmLote = {};
    
    // Contadores para o Log
    let processadosValidos = 0;
    let ignoradosPorVencimento = 0;

    valoresNecton.forEach((linha, i) => {
      const tickerOpcao   = linha[colTicker - 1];
      const idTrade       = linha[colIdTrade - 1]; 
      const idEstrutura   = linha[colIdEstrutura - 1];
      const vencimentoStr = linha[colVencimento - 1];

      // Ignora linhas em branco
      if (!tickerOpcao || !idTrade || idTrade === "") return;

      // --- NOVA LÓGICA DE FILTRAGEM POR VENCIMENTO ---
      let dataVencimento;
      if (vencimentoStr instanceof Date) {
        dataVencimento = vencimentoStr;
      } else if (typeof vencimentoStr === 'string' && vencimentoStr !== "") {
        // Tenta converter string (ex: "2026-04-17") para Date
        dataVencimento = new Date(vencimentoStr); 
      }
      
      if (dataVencimento && !isNaN(dataVencimento.getTime())) {
        dataVencimento.setHours(0, 0, 0, 0);
        if (dataVencimento < hoje) {
          ignoradosPorVencimento++;
          return; // Pula esta iteração (não chama a API)
        }
      } else {
        // Se a data for inválida ou vazia, podemos ignorar para evitar falhas ou processar.
        // Optamos por ignorar por segurança.
        ignoradosPorVencimento++;
        return; 
      }
      // -----------------------------------------------

      processadosValidos++;

      log(SERVICO_NOME, "DEBUG_DATA", `Validado p/ API: ${tickerOpcao}`, `Vencimento: ${dataVencimento.toISOString().split('T')[0]}`);
      
      const dadosAPI = getOpLabOptionDetails(tickerOpcao);
      
      if (dadosAPI) {
        // Alimenta o objeto com as chaves de controle
        dadosAPI.ID_Trade_Unico = idTrade;
        dadosAPI.ID_Estrutura = idEstrutura;
        dadosAPI.Timestamp_Atualizacao = timestamp;

        // Gera a linha seguindo o array FIXED_HEADERS_DETAILS
        const linhaValores = FIXED_HEADERS_DETAILS.map(campo => {
          const v = dadosAPI[campo];
          return (v !== undefined && v !== null) ? v : "";
        });

        if (idToRowMap[idTrade]) {
          updatesEmLote[idToRowMap[idTrade]] = linhaValores;
        } else {
          listaParaNovos.push(linhaValores);
        }
      }
      if (i % 10 === 0) Utilities.sleep(50);
    });

    // Escrita em Lote
    Object.keys(updatesEmLote).forEach(rowNum => {
      abaDadosDetalhes.getRange(parseInt(rowNum), 1, 1, FIXED_HEADERS_DETAILS.length).setValues([updatesEmLote[rowNum]]);
    });

    if (listaParaNovos.length > 0) {
      abaDadosDetalhes.getRange(abaDadosDetalhes.getLastRow() + 1, 1, listaParaNovos.length, FIXED_HEADERS_DETAILS.length).setValues(listaParaNovos);
    }

    // --- LOG ENRIQUECIDO ---
    log(
      SERVICO_NOME, 
      "SUCESSO", 
      "Sincronização de detalhes finalizada.", 
      `Total lido: ${valoresNecton.length} | Processados (API): ${processadosValidos} | Ignorados (Vencidos): ${ignoradosPorVencimento}`
    );

  } catch (e) {
    log(SERVICO_NOME, "ERRO_FATAL", "Falha no Sync Detalhes", e.toString());
  }
}

function getDetailsMapAndHeaders(aba) {
  const ultimaLinha = aba.getLastRow();
  const idToRowMap = {};
  if (ultimaLinha > 1) {
    const ids = aba.getRange(2, 1, ultimaLinha - 1, 1).getValues();
    ids.forEach((linha, index) => {
      if (linha[0]) idToRowMap[linha[0]] = index + 2;
    });
  }
  return { idToRowMap: idToRowMap };
}






// ═══════════════════════════════════════════════════════════════
// ORQUESTRADOR: SINCRONIZAR DADOS DETALHES
// ═══════════════════════════════════════════════════════════════

/**
 * Orquestra a sincronização de dados detalhados de opções
 * 
 * Este orquestrador:
 * 1. Valida pré-condições
 * 2. Registra início da operação
 * 3. Chama o serviço especializado (sincronizarDadosDetalhes)
 * 4. Registra resultado e métricas
 * 5. Trata erros de forma centralizada
 * 
 * @returns {Object} Resultado da operação com métricas
 */
function orquestrarSyncDadosDetalhes() {
  const OPERACAO_ID = Utilities.getUuid();
  const inicio = new Date();
  
  logOrquestrador("INICIO", "Iniciando orquestração de sincronização de dados detalhados", {
    operacao_id: OPERACAO_ID,
    timestamp: inicio.toISOString(),
    servico: SERVICOS_REGISTRY.SYNC_DADOS_DETALHES.nome
  });
  
  try {
    // --- ETAPA 1: Validação de Pré-condições ---
    const validacao = validarPreCondicoes();
    if (!validacao.sucesso) {
      throw new Error("Pré-condições não atendidas: " + validacao.mensagem);
    }
    
    logOrquestrador("INFO", "Pré-condições validadas com sucesso", validacao);
    
    // --- ETAPA 2: Execução do Serviço Especializado ---
    logOrquestrador("INFO", "Chamando serviço: " + SERVICOS_REGISTRY.SYNC_DADOS_DETALHES.funcao, {
      servico: SERVICOS_REGISTRY.SYNC_DADOS_DETALHES.nome
    });
    
    const resultadoServico = executarComTimeout(
      SERVICOS_REGISTRY.SYNC_DADOS_DETALHES.funcao,
      SERVICOS_REGISTRY.SYNC_DADOS_DETALHES.timeout
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
      servico: SERVICOS_REGISTRY.SYNC_DADOS_DETALHES.nome,
      status: "SUCESSO"
    };
    
    logOrquestrador("SUCESSO", "Orquestração de sync dados detalhados concluída com sucesso", metricas);
    
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
    
    logOrquestrador("ERRO_CRITICO", "Falha na orquestração de sync dados detalhados", erroDetalhado);
    
    // Notifica usuário sobre o erro (seguro p/ backend)
    safeAlert_(
      '❌ Erro na Sincronização',
      'Ocorreu um erro durante a sincronização de dados detalhados:\n\n' +
      erro.message +
      '\n\nVerifique a aba Logs para mais detalhes.'
    );
    
    return {
      sucesso: false,
      erro: erroDetalhado
    };
  }
}








function testarSyncDetalhesUnitario() {
  console.log("--- INICIANDO TESTE UNITÁRIO 005: DETALHES DA OPÇÃO ---");
  
  // Usando um ticker que sabemos que está ativo (da sua lista de sucesso anterior)
  const tickerTeste = "BRKMO780"; 
  console.log("🚀 Buscando detalhes completos para: " + tickerTeste);
  
  const dados = getOpLabOptionDetails(tickerTeste);
  
  if (dados && dados.strike) {
    console.log(`✅ SUCESSO: Conexão com OpLab via Client 009 estável.`);
    console.log(`📊 Strike: ${dados.strike} | Tipo: ${dados.type}`);
    console.log(`⏳ Dias para Vencimento: ${dados.days_to_maturity}`);
    console.log(`🏛️ Market Maker: ${dados.market_maker ? "Sim" : "Não"}`);
    console.log(`🔍 ISIN: ${dados.isin}`);
    
    // Teste de alinhamento de array
    const testeMapeamento = FIXED_HEADERS_DETAILS.slice(0, 5).map(campo => {
      return campo + ": " + (dados[campo] || "N/A");
    });
    console.log("📋 Amostra de Mapeamento (Primeiros 5 campos):", JSON.stringify(testeMapeamento));
    
  } else {
    console.error("❌ FALHA: Não foi possível recuperar os detalhes. Verifique o ticker.");
  }
  
  console.log("--- FIM DO TESTE ---");
}
