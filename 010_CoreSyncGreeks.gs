/**
 * @fileoverview CoreSyncGreeks - v4.1 (Risk Engine & BS Cache)
 * AÇÃO: Calcula e sincroniza Gregas (BS) apenas para operações ATIVAS.
 * OTIMIZAÇÃO: Cache por ativo para evitar cálculos redundantes.
 * PADRÃO: Modo Silencioso e Contexto Serializado.
 */

const GreeksSync = {
  _serviceName: "GreeksSync_v4.1",

  run() {
    const inicio = Date.now();
    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);

    const cacheBS = {}; // Cache para não recalcular a mesma opção
    const stats = { calculos: 0, cache: 0, pulados: 0, gravados: 0 };
    const metadadosExecucao = {
      aba_gatilho: SYS_CONFIG.SHEETS.TRIGGER,
      aba_destino: SYS_CONFIG.SHEETS.GREEKS,
      timestamp_inicio: new Date().toISOString()
    };

    // MARCADOR DE TERRITÓRIO: INÍCIO
    SysLogger.log(this._serviceName, "START", ">>> INICIANDO MOTOR DE RISCO (GREGAS) <<<", JSON.stringify(metadadosExecucao));

    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const abaGatilho = ss.getSheetByName(SYS_CONFIG.SHEETS.TRIGGER);
      const abaGreeks = ss.getSheetByName(SYS_CONFIG.SHEETS.GREEKS);
      
      if (!abaGatilho || !abaGreeks) {
        throw new Error(`Abas ausentes: Verifique ${SYS_CONFIG.SHEETS.TRIGGER} e ${SYS_CONFIG.SHEETS.GREEKS}`);
      }
      
      // 1. CARREGAR MAPAS DE SUPORTE (Details e Assets)
      const detailsMap = this._getDynamicMap(ss.getSheetByName(SYS_CONFIG.SHEETS.DETAILS), "ID_Trade_Unico");
      const assetsMap = this._getDynamicMap(ss.getSheetByName(SYS_CONFIG.SHEETS.ASSETS), "Ticker");

      // 2. SCAN DINÂMICO DE CABEÇALHOS
      const headersG = abaGatilho.getRange(1, 1, 1, abaGatilho.getLastColumn()).getValues()[0];
      const colMapG = {};
      headersG.forEach((h, i) => { if(h) colMapG[String(h).trim()] = i; });

      const cabecalhosD = abaGreeks.getRange(1, 1, 1, abaGreeks.getLastColumn()).getValues()[0];
      const headerMapD = {};
      cabecalhosD.forEach((h, i) => { if(h) headerMapD[String(h).trim()] = i; });

      // Mapear linhas existentes no destino para UPDATE
      const idToRowMap = {};
      if (abaGreeks.getLastRow() > 1) {
        const ids = abaGreeks.getRange(2, 1, abaGreeks.getLastRow() - 1, 1).getValues();
        ids.forEach((l, i) => { if (l[0]) idToRowMap[String(l[0]).trim()] = i + 2; });
      }

      const valoresGatilho = abaGatilho.getDataRange().getValues();
      
      const dataHoje = DataUtils.formatDateBR(new Date());
      const horaHoje = new Date().toLocaleTimeString('pt-BR');
      const timestampAtual = `${dataHoje} ${horaHoje}`;

      // 3. LOOP DE PROCESSAMENTO
      for (let i = 1; i < valoresGatilho.length; i++) {
        const linhaAtu = valoresGatilho[i];
        const status = String(linhaAtu[colMapG["Status Operação"]] || "").trim().toUpperCase();
        const idTrade = String(linhaAtu[colMapG["ID_Trade_Unico"]] || "").trim();
        const tickerOpcao = String(linhaAtu[0]).trim();
        const qtde = Math.abs(Number(linhaAtu[colMapG["Qtd. exec"]] || 0));

        // FILTROS DE SEGURANÇA E PERFORMANCE
        if (status !== "ATIVO") { stats.pulados++; continue; }
        if (!idTrade || idTrade.length < 10 || isNaN(qtde)) continue;

        const detail = detailsMap[idTrade];
        const asset = detail ? assetsMap[detail.parent_symbol] : null;

        if (!detail || !asset) {
          SysLogger.log(this._serviceName, "AVISO", `Dados insuficientes para ${tickerOpcao} (ID: ${idTrade}).`, "A opção precisa existir nas abas Dados_Ativos e Dados_Detalhes.");
          continue;
        }

        // 4. CÁLCULO BLACK-SCHOLES (COM CACHE)
        let bsResult = null;
        if (cacheBS[tickerOpcao]) {
          bsResult = cacheBS[tickerOpcao];
          stats.cache++;
        } else {
          const params = {
            symbol: tickerOpcao,
            irate: 10.75, // TODO: No futuro, extrair da aba SYS_CONFIG
            type: detail.type,
            spotprice: Number(asset.close),
            strike: Number(detail.strike),
            dtm: Number(detail.days_to_maturity),
            vol: Number(asset.iv_current),
            amount: qtde
          };

          bsResult = OplabService.calculateBS(params);
          if (bsResult) {
            cacheBS[tickerOpcao] = bsResult;
            stats.calculos++;
            Utilities.sleep(800); // Rate Limit
          }
        }

        if (bsResult) {
          // CORREÇÃO: Gravando exatamente o código da opção (tickerOpcao) na coluna Ativo
          const dadosCompletos = {
            ...bsResult,
            ID_Trade_Unico: idTrade,
            Ativo: tickerOpcao, 
            ID_Estrutura: String(linhaAtu[colMapG["ID_Estrutura"]] || "").trim(),
            Timestamp_Atualizacao: timestampAtual
          };

          // Construção da linha preservando colunas manuais (Merge)
          const rowNum = idToRowMap[idTrade];
          let linhaFinal = rowNum ? abaGreeks.getRange(rowNum, 1, 1, cabecalhosD.length).getValues()[0] : new Array(cabecalhosD.length).fill("");

          for (const key in headerMapD) {
            const colIdx = headerMapD[key];
            if (dadosCompletos[key] !== undefined) {
              linhaFinal[colIdx] = typeof dadosCompletos[key] === 'object' ? JSON.stringify(dadosCompletos[key]) : dadosCompletos[key];
            }
          }

          // GRAVAÇÃO
          if (rowNum) {
            abaGreeks.getRange(rowNum, 1, 1, cabecalhosD.length).setValues([linhaFinal]);
          } else {
            abaGreeks.appendRow(linhaFinal);
            idToRowMap[idTrade] = abaGreeks.getLastRow();
          }
          stats.gravados++;
          
          SysLogger.log(this._serviceName, "SUCESSO", `Gregas de ${tickerOpcao} processadas.`, `Cacheado: ${cacheBS[tickerOpcao] ? 'Sim' : 'Não'}`);
        }
        
        if (stats.gravados % 5 === 0) SysLogger.flush();
      }

      // MARCADOR DE TERRITÓRIO: FINALIZAÇÃO
      const duracao = ((Date.now() - inicio) / 1000).toFixed(1);
      stats.duracao_segundos = duracao;

      SysLogger.log(this._serviceName, "FINISH", `>>> CICLO FINALIZADO EM ${duracao}s <<<`, JSON.stringify(stats));
      SysLogger.flush();

    } catch (e) {
      SysLogger.log(this._serviceName, "CRITICO", "Falha fatal no motor de Gregas", String(e.message));
      SysLogger.flush();
    }
  },

  /** Helper para criar mapa dinâmico de abas de suporte */
  _getDynamicMap(aba, pk) {
    if (!aba) return {};
    const data = aba.getDataRange().getValues();
    const headers = data[0];
    const pkIdx = headers.indexOf(pk);
    if (pkIdx === -1) return {};
    
    const map = {};
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const obj = {};
      headers.forEach((h, idx) => obj[h] = row[idx]);
      if (row[pkIdx]) map[String(row[pkIdx]).trim()] = obj;
    }
    return map;
  }
};

// ============================================================================
// PONTO DE ENTRADA (Trigger Dinâmico / Menu)
// ============================================================================

/** Ponto de Entrada para o Motor de Risco */
function atualizarGregas() {
  GreeksSync.run();
}

// ============================================================================
// SUÍTE DE HOMOLOGAÇÃO (010)
// ============================================================================

/**
 * Testa o cálculo individual de Black-Scholes via API OpLab.
 */
function testSuiteGreeksSync010() {
  console.log("=== INICIANDO HOMOLOGAÇÃO: GREEKS SYNC (010) ===");
  
  const paramsTeste = {
    symbol: "PETRR315",
    irate: 10.75,      // Taxa Selic (em %)
    type: "PUT",       // Tipo da opção
    spotprice: 40.69,  // Preço atual do ativo objeto
    strike: 30.73,     // Preço de exercício
    dtm: 71,           // Dias úteis para o vencimento
    vol: 35.5,         // Volatilidade Implícita (em %)
    amount: 1000       // Quantidade da posição
  };

  console.log(`🚀 Solicitando cálculo Black-Scholes para ${paramsTeste.symbol}...`);
  
  const t0 = Date.now();
  const resultado = OplabService.calculateBS(paramsTeste);
  const t1 = Date.now();

  if (resultado && resultado.delta !== undefined) {
    console.log(`✅ SUCESSO: Resposta recebida em ${t1 - t0}ms.`);
    console.log(`📐 Delta: ${resultado.delta} (Exposição direcional)`);
    console.log(`📐 Gamma: ${resultado.gamma} (Aceleração do Delta)`);
    console.log(`📐 Theta: ${resultado.theta} (Decaimento temporal)`);
    console.log(`💰 Preço Teórico: R$ ${resultado.price}`);
  } else {
    console.error("❌ FALHA: A API não retornou cálculos válidos.");
  }
  
  console.log("--- Executando Carga Controlada ---");
  GreeksSync.run();

  console.log("=== TESTES CONCLUÍDOS ===");
}