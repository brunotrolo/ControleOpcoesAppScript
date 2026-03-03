/**
 * MÓDULO: SyncGreeks - v1.5
 * STATUS: Homologado via Swagger
 */

const FIXED_HEADERS_GREEKS = [
  'ID_Trade_Unico', 'Ativo', 'ID_Estrutura', 'Timestamp_Atualizacao',
  'moneyness', 'price', 'delta', 'gamma', 'vega', 'theta', 'rho',
  'volatility', 'poe', 'spotprice', 'strike', 'margin'
];

function sincronizarGreeks() {
  const SERVICO_NOME = "SyncGreeks_v1.5";
  log(SERVICO_NOME, "INFO", "Iniciando sincronização via Black-Scholes...", "");
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaNecton = ss.getSheetByName(ABA_GATILHO);
    const abaGreeks = ss.getSheetByName(ABA_GREEKS);
    const abaDetails = ss.getSheetByName(ABA_DADOS_DETALHES);
    const abaAssets = ss.getSheetByName(ABA_DADOS_ATIVOS);
    
    const detailsMap = getMapFromSheet_(abaDetails, 'ID_Trade_Unico');
    const assetsMap = getMapFromSheet_(abaAssets, 'Ticker');
    const { idToRowMap } = getGreeksMapAndHeaders_(abaGreeks);

    const nectonData = abaNecton.getDataRange().getValues();
    const timestamp = new Date().toISOString();
    
    const updatesEmLote = {};
    const listaParaNovos = [];
    let contadorSucesso = 0;

    for (let i = 1; i < nectonData.length; i++) {
      const linha = nectonData[i];
      const tickerOpcao = linha[0];  
      const idTrade     = linha[16]; 
      const idEstrutura = linha[17]; 
      const qtde        = Math.abs(linha[15] || 0); 

      const detail = detailsMap[idTrade];
      const asset = detail ? assetsMap[detail.parent_symbol] : null;

      if (!idTrade || !detail || !asset) continue;

      const bsParams = {
        symbol: tickerOpcao,
        irate: 10.75, 
        type: detail.type, 
        spotprice: parseFloat(asset.close),
        strike: parseFloat(detail.strike),
        dtm: parseInt(detail.days_to_maturity),
        vol: parseFloat(asset.iv_current),
        amount: qtde
      };

      const bsResult = calculateOpLabBS(bsParams);

      if (bsResult) {
        const dadosCompletos = {
          ...bsResult,
          ID_Trade_Unico: idTrade,
          Ativo: detail.parent_symbol,
          ID_Estrutura: idEstrutura,
          Timestamp_Atualizacao: timestamp
        };

        const linhaGreeks = FIXED_HEADERS_GREEKS.map(campo => dadosCompletos[campo] ?? "");

        if (idToRowMap[idTrade]) {
          updatesEmLote[idToRowMap[idTrade]] = linhaGreeks;
        } else {
          listaParaNovos.push(linhaGreeks);
        }
        contadorSucesso++;
      }
      if (i % 15 === 0) Utilities.sleep(50);
    }

    // Escrita Final
    Object.keys(updatesEmLote).forEach(rowNum => {
      abaGreeks.getRange(parseInt(rowNum), 1, 1, FIXED_HEADERS_GREEKS.length).setValues([updatesEmLote[rowNum]]);
    });

    if (listaParaNovos.length > 0) {
      abaGreeks.getRange(abaGreeks.getLastRow() + 1, 1, listaParaNovos.length, FIXED_HEADERS_GREEKS.length).setValues(listaParaNovos);
    }

    log(SERVICO_NOME, "SUCESSO", "Cálculos concluídos.", `Processados: ${contadorSucesso}`);

  } catch (e) {
    log(SERVICO_NOME, "ERRO_FATAL", "Falha no Sync Greeks", e.toString());
  }
}

// Helpers necessários
function getMapFromSheet_(aba, primaryKey) {
  const data = aba.getDataRange().getValues();
  const headers = data[0];
  const keyIndex = headers.indexOf(primaryKey);
  const map = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const obj = {};
    headers.forEach((h, idx) => obj[h] = row[idx]);
    if (row[keyIndex]) map[row[keyIndex]] = obj;
  }
  return map;
}

function getGreeksMapAndHeaders_(aba) {
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
// ORQUESTRADOR: SINCRONIZAR GREEKS
// ═══════════════════════════════════════════════════════════════

/**
 * Orquestra o cálculo de gregas (Black-Scholes)
 * 
 * Este orquestrador:
 * 1. Valida pré-condições
 * 2. Registra início da operação
 * 3. Chama o serviço especializado (sincronizarGreeks)
 * 4. Registra resultado e métricas
 * 5. Trata erros de forma centralizada
 * 
 * @returns {Object} Resultado da operação com métricas
 */
function orquestrarSyncGreeks() {
  const OPERACAO_ID = Utilities.getUuid();
  const inicio = new Date();
  
  logOrquestrador("INICIO", "Iniciando orquestração de cálculo de gregas", {
    operacao_id: OPERACAO_ID,
    timestamp: inicio.toISOString(),
    servico: SERVICOS_REGISTRY.SYNC_GREEKS.nome
  });
  
  try {
    // --- ETAPA 1: Validação de Pré-condições ---
    const validacao = validarPreCondicoes();
    if (!validacao.sucesso) {
      throw new Error("Pré-condições não atendidas: " + validacao.mensagem);
    }
    
    logOrquestrador("INFO", "Pré-condições validadas com sucesso", validacao);
    
    // --- ETAPA 2: Execução do Serviço Especializado ---
    logOrquestrador("INFO", "Chamando serviço: " + SERVICOS_REGISTRY.SYNC_GREEKS.funcao, {
      servico: SERVICOS_REGISTRY.SYNC_GREEKS.nome
    });
    
    const resultadoServico = executarComTimeout(
      SERVICOS_REGISTRY.SYNC_GREEKS.funcao,
      SERVICOS_REGISTRY.SYNC_GREEKS.timeout
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
      servico: SERVICOS_REGISTRY.SYNC_GREEKS.nome,
      status: "SUCESSO"
    };
    
    logOrquestrador("SUCESSO", "Orquestração de cálculo de gregas concluída com sucesso", metricas);
    
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
    
    logOrquestrador("ERRO_CRITICO", "Falha na orquestração de cálculo de gregas", erroDetalhado);
    
    // Notifica usuário sobre o erro (seguro p/ backend)
    safeAlert_(
      '❌ Erro no Cálculo de Gregas',
      'Ocorreu um erro durante o cálculo de gregas:\n\n' +
      erro.message +
      '\n\nVerifique a aba Logs para mais detalhes.'
    );
    
    return {
      sucesso: false,
      erro: erroDetalhado
    };
  }
}















function testarCalculoGreeksSwagger() {
  console.log("--- TESTE UNITÁRIO 006: VALIDAÇÃO SWAGGER ---");
  
  const paramsTeste = {
    symbol: "BRKMO780",
    irate: 10.75,      // Em % conforme Swagger
    type: "PUT",
    spotprice: 9.80,   // Preço da BRKM5
    strike: 7.8,
    dtm: 26,
    vol: 54.99         // IV em % conforme Swagger
  };
  
  console.log("🚀 Chamando BS com parâmetros do Swagger...");
  const resultado = calculateOpLabBS(paramsTeste);
  
  if (resultado && resultado.delta !== null) {
    console.log("✅ SUCESSO!");
    console.log("📐 Delta: " + resultado.delta);
    console.log("📐 Gamma: " + resultado.gamma);
    console.log("💰 Preço Teórico: " + resultado.price);
    console.log("🎯 Moneyness: " + resultado.moneyness);
  } else {
    console.error("❌ A API ainda retornou dados nulos. Verifique o Access-Token ou os valores de entrada.");
  }
}