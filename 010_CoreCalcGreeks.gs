/*
 * ═══════════════════════════════════════════════════════════════
 * CALC GREEKS - CÁLCULO DE GREGAS VIA BLACK-SCHOLES
 * ═══════════════════════════════════════════════════════════════
 * 
 * RESPONSABILIDADES:
 * - Calcular gregas usando fórmulas Black-Scholes nativas
 * - Calcular Implied Volatility usando Newton-Raphson
 * - Calcular POE (Probability of Expiring ITM)
 * - Fazer MERGE na aba Calc_Greeks
 * 
 * VERSÃO: 1.2 (clamping de T + ajuste near-expiration)
 * DATA: 2025-11-14
 * 
 * ═══════════════════════════════════════════════════════════════
 */

const ABA_CALC_GREEKS = "Calc_Greeks";


// Headers fixos (14 campos) - colunas simplificadas
const FIXED_HEADERS_CALC_GREEKS = [
  'ID_Trade_Unico',
  'Ativo',
  'ID_Estrutura',
  'Timestamp_Atualizacao',
  'moneyness_code',
  'moneyness',
  'price',
  'delta',
  'gamma',
  'vega',
  'theta',
  'rho',
  'volatility',
  'poe'
];

// Constantes para cálculos
const DIAS_ANO = 252;              // Dias úteis
const DEFAULT_VOL_INICIAL = 0.35;  // 35% para Newton-Raphson
const MAX_ITER_IV = 100;
const TOLERANCIA_IV = 0.0001;


// Solução A: T mínimo para estabilidade numérica (~0.5 dia útil)
const T_MIN = 0.002;

// Solução B: janela de "near expiration" em dias corridos
const NEAR_EXPIRATION_DAYS = 3;

// ═══════════════════════════════════════════════════════════════
// MÓDULO: FUNÇÕES MATEMÁTICAS (DISTRIBUIÇÃO NORMAL)
// ═══════════════════════════════════════════════════════════════

/**
 * Função de densidade de probabilidade normal padrão (PDF)
 */
function normPDF(x) {
  return Math.exp(-0.5 * x * x) / Math.sqrt(2 * Math.PI);
}

/**
 * Função de distribuição acumulada normal padrão (CDF)
 * Aproximação de Abramowitz e Stegun
 */
function normCDF(x) {
  const t = 1 / (1 + 0.2316419 * Math.abs(x));
  const d = 0.3989423 * Math.exp(-x * x / 2);
  const prob = d * t * (0.3193815 + t * (-0.3565638 + t * (1.781478 + t * (-1.821256 + t * 1.330274))));
  
  return x > 0 ? 1 - prob : prob;
}


/**
 * Valida se a data de vencimento é válida e futura (ou hoje).
 * @param {Date|string} dateValue Valor da célula de vencimento
 * @return {boolean}
 */
function checkDueDateFromColT(dateValue) {
  if (!dateValue) return false;
  
  try {
    const dataVencimento = (dateValue instanceof Date) ? dateValue : new Date(dateValue);
    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);
    dataVencimento.setHours(0, 0, 0, 0);
    
    // Retorna true se o vencimento for hoje ou no futuro
    return dataVencimento.getTime() >= hoje.getTime();
  } catch (e) {
    return false;
  }
}



// ═══════════════════════════════════════════════════════════════
/**
 * BLACK-SCHOLES CORE
 */
// ═══════════════════════════════════════════════════════════════

/**
 * Calcula d1 e d2 do modelo Black-Scholes
 */
function calcularD1D2(S, K, T, r, sigma) {
  if (T <= 0 || sigma <= 0) {
    return { d1: 0, d2: 0 };
  }
  
  const sqrtT = Math.sqrt(T);
  const sigmaSqrtT = sigma * sqrtT;
  
  const d1 = (Math.log(S / K) + (r + 0.5 * sigma * sigma) * T) / sigmaSqrtT;
  const d2 = d1 - sigmaSqrtT;
  
  return { d1, d2 };
}

/**
 * Calcula o preço teórico usando Black-Scholes
 */
function calcularPrecoBS(S, K, T, r, sigma, flag) {
  if (T <= 0) return Math.max(flag === 'c' ? S - K : K - S, 0);
  
  const { d1, d2 } = calcularD1D2(S, K, T, r, sigma);
  
  if (flag === 'c') {
    // Call
    return S * normCDF(d1) - K * Math.exp(-r * T) * normCDF(d2);
  } else {
    // Put
    return K * Math.exp(-r * T) * normCDF(-d2) - S * normCDF(-d1);
  }
}





/**
 * BLACK-SCHOLES V2 (Renomeada para evitar conflito com versões antigas)
 */
function calcularGregasV2(S, K, T, r, sigma, flag) {
  // Retorno seguro se dados inválidos
  if (T <= 0 || sigma <= 0) {
    return { price: 0, delta: 0, gamma: 0, vega: 0, theta: 0, rho: 0, poe: 0 };
  }
  
  const { d1, d2 } = calcularD1D2(S, K, T, r, sigma);
  const sqrtT = Math.sqrt(T);
  const nd1 = normPDF(d1);
  const Nd1 = normCDF(d1);
  const Nd2 = normCDF(d2);
  const discountFactor = Math.exp(-r * T); // Fator de desconto explícito
  
  // 1. Preço (Lógica corrigida e explícita)
  let price;
  if (flag === 'c') {
    price = S * Nd1 - K * discountFactor * Nd2;
  } else {
    // Put: K * e^(-rT) * N(-d2) - S * N(-d1)
    price = K * discountFactor * normCDF(-d2) - S * normCDF(-d1);
  }
  
  // 2. Delta
  const delta = flag === 'c' ? Nd1 : Nd1 - 1;
  
  // 3. Gamma
  const gamma = nd1 / (S * sigma * sqrtT);
  
  // 4. Vega (DIVIDIDO POR 100)
  const vega = (S * nd1 * sqrtT) / 100;
  
  // 5. Theta (DIVIDIDO POR 252 - DIAS ÚTEIS)
  let theta;
  if (flag === 'c') {
    theta = (-(S * nd1 * sigma) / (2 * sqrtT) - r * K * discountFactor * Nd2) / DIAS_ANO;
  } else {
    theta = (-(S * nd1 * sigma) / (2 * sqrtT) + r * K * discountFactor * normCDF(-d2)) / DIAS_ANO;
  }

  // 6. Rho (DIVIDIDO POR 100)
  let rho;
  if (flag === 'c') {
    rho = (K * T * discountFactor * Nd2) / 100;
  } else {
    rho = (-K * T * discountFactor * normCDF(-d2)) / 100;
  }
  
  // 7. POE
  const poe = flag === 'c' ? Nd2 : normCDF(-d2);
  
  return { price, delta, gamma, vega, theta, rho, poe };
}







/**
 * Ajuste near-expiration (Solução B)
 * - Não altera price, delta, poe
 * - Suaviza gamma, vega, theta e rho à medida que DTE → 0
 */
function ajustarGregasNearVencimento(gregas, alpha) {
  // alpha ∈ (0, 1] - fator de suavização
  return {
    price: gregas.price,
    delta: gregas.delta,
    gamma: gregas.gamma * alpha,
    vega: gregas.vega * alpha,
    theta: gregas.theta * alpha,
    rho: gregas.rho * alpha,
    poe: gregas.poe
  };
}

/**
 * Calcula Implied Volatility usando Newton-Raphson
 */
function calcularImpliedVolatility(S, K, T, r, precoMercado, flag) {
  if (T <= 0 || precoMercado <= 0) return 0;
  
  let sigma = DEFAULT_VOL_INICIAL;
  
  for (let i = 0; i < MAX_ITER_IV; i++) {
    const precoBS = calcularPrecoBS(S, K, T, r, sigma, flag);
    const diff = precoBS - precoMercado;
    
    if (Math.abs(diff) < TOLERANCIA_IV) {
      return sigma;
    }
    
    // Vega para Newton-Raphson
    const { d1 } = calcularD1D2(S, K, T, r, sigma);
    const vega = S * normPDF(d1) * Math.sqrt(T);
    
    if (vega < 0.0001) break; // Evita divisão por zero
    
    sigma = sigma - diff / vega;
    
    // Limites razoáveis para IV
    if (sigma < 0.01) sigma = 0.01;
    if (sigma > 5.0) sigma = 5.0;
  }
  
  return sigma;
}

/**
 * Calcula moneyness code (ATM/ITM/OTM)
 */
function calcularMoneynessCode(S, K, flag) {
  const moneyness = S / K;
  
  // ATM: entre 97.5% e 102.5%
  if (moneyness >= 0.975 && moneyness <= 1.025) {
    return 'ATM';
  }
  
  // ITM: Call > 102.5% ou Put < 97.5%
  if ((flag === 'c' && moneyness > 1.025) || (flag === 'p' && moneyness < 0.975)) {
    return 'ITM';
  }
  
  return 'OTM';
}

// ═══════════════════════════════════════════════════════════════
// FUNÇÃO PRINCIPAL
// ═══════════════════════════════════════════════════════════════

/**
 * Calcula gregas usando Black-Scholes implementado nativamente
 */
function calcularGreeksNativo() {
  const SERVICO_NOME = "CalcGreeks_v1";
  log(SERVICO_NOME, "INFO", "Iniciando cálculo de gregas (nativo)...", "");
  
  try {
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    
    // --- PASSO 1: Abrir abas ---
    const abaNecton = planilha.getSheetByName(ABA_GATILHO);
    const abaDetails = planilha.getSheetByName(ABA_DADOS_DETALHES);
    const abaAssets = planilha.getSheetByName(ABA_DADOS_ATIVOS);
    const abaConfig = planilha.getSheetByName(ABA_CONFIG);
    const abaCalcGreeks = planilha.getSheetByName(ABA_CALC_GREEKS);
    
    if (!abaNecton || !abaDetails || !abaAssets || !abaConfig || !abaCalcGreeks) {
      log(SERVICO_NOME, "ERRO_CRITICO", "Uma ou mais abas não encontradas", "");
      throw new Error("Abas necessárias não encontradas");
    }
    
    // --- PASSO 2: Ler todos os dados ---
    log(SERVICO_NOME, "INFO", "Lendo fontes: Necton, Details, Assets, Config...", "");
    
    const nectonTrades = getAllRecordsSafely(abaNecton);
    const detailsRecords = getAllRecordsSafely(abaDetails);
    const assetsRecords = getAllRecordsSafely(abaAssets);
    const configRecords = getAllRecordsSafely(abaConfig);
    
    // Criar mapas
    const configMap = {};
    configRecords.forEach(row => {
      const chave = row['Chave'];
      if (chave) configMap[chave] = row['Valor'];
    });
    
    const assetsMap = {};
    assetsRecords.forEach(row => {
      const ticker = row['Ticker'];
      if (ticker) assetsMap[ticker] = row;
    });
    
    const detailsMap = {};
    detailsRecords.forEach(row => {
      const idTrade = row['ID_Trade_Unico'];
      if (idTrade) detailsMap[idTrade] = row;
    });
    
    const { headers: calcHeaders, idToRowMap } = getCalcGreeksMapAndHeaders(abaCalcGreeks);
    const idsExistentes = Object.keys(idToRowMap);
    log(SERVICO_NOME, "INFO", "IDs existentes na Calc_Greeks", `${idsExistentes.length} registros`);
    
    // Garantir headers corretos
    if (!arraysEqual(calcHeaders, FIXED_HEADERS_CALC_GREEKS)) {
      const headerRange = abaCalcGreeks.getRange(1, 1, 1, FIXED_HEADERS_CALC_GREEKS.length);
      headerRange.setValues([FIXED_HEADERS_CALC_GREEKS]);
      log(SERVICO_NOME, "INFO_HEADER", "Cabeçalho Calc_Greeks atualizado (14 colunas)", "");
    }
    
    // Obter taxa de juros
    const irate = safeFloatConvert(configMap['Taxa_Selic_Anual'] || '0.10');

    const updatesInPlace = [];
    const appendNew = [];
    const timestamp = new Date().toISOString();
    
    // --- PASSO 3: PRÉ-FILTRAGEM ---
    log(SERVICO_NOME, "INFO", "Iniciando pré-filtragem...", "");
    
    const tradesFiltrados = [];
    let totalNecton = nectonTrades.length;
    let ignoradosVencidos = 0;
    let ignoradosManual = 0;
    
    for (let i = 0; i < nectonTrades.length; i++) {
      const trade = nectonTrades[i];
      
      // Filtro 1: Vencimento
      if (!checkDueDateFromColT(trade[COLUNA_VENCIMENTO])) {
        ignoradosVencidos++;
        continue;
      }
      
      // Filtro 2: Status manual
      if (trade[COLUNA_STATUS_ROBO] === 'IGNORAR') {
        ignoradosManual++;
        continue;
      }
      
      tradesFiltrados.push(trade);
    }
    
    const tradesValidos = tradesFiltrados.length;
    const contextoLog = `Total: ${totalNecton} | Vencidos: ${ignoradosVencidos} | Manual: ${ignoradosManual}`;
    log(SERVICO_NOME, "INFO_SUMARIO", `Trades válidos: ${tradesValidos}`, contextoLog);
    
    // --- PASSO 4: CALCULAR GREGAS ---
    let processados = 0;
    let erros = 0;
    
    for (let i = 0; i < tradesFiltrados.length; i++) {
      const trade = tradesFiltrados[i];
      const idTrade = trade['ID_Trade_Unico'];
      const ativo = trade['Ativo'];
      
      // Buscar dados
      const detail = detailsMap[idTrade];
      if (!detail) {
        log(SERVICO_NOME, "AVISO_DADOS", `Details faltando: ${idTrade}`, "");
        continue;
      }
      
      const parentTicker = detail['parent_symbol'];
      if (!parentTicker) continue;
      
      if (!assetsMap[parentTicker]) {
        log(SERVICO_NOME, "AVISO_DADOS", `Asset faltando: ${parentTicker}`, "");
        continue;
      }
      
      const assetData = assetsMap[parentTicker];
      
      try {
        // --- PREPARAR PARÂMETROS ---
        // Data de vencimento
        let dueDate;
        const dueDateRaw = detail['due_date'];
        
        if (dueDateRaw instanceof Date) {
          dueDate = new Date(dueDateRaw);
        } else if (typeof dueDateRaw === 'string') {
          const dtmStr = dueDateRaw.split('T')[0];
          dueDate = new Date(dtmStr);
        } else {
          dueDate = new Date(dueDateRaw);
        }
        
        const hoje = new Date();
        hoje.setHours(0, 0, 0, 0);
        dueDate.setHours(0, 0, 0, 0);
        
        const daysToMaturity = Math.floor((dueDate - hoje) / (1000 * 60 * 60 * 24));
        if (daysToMaturity < 0) continue; // Já venceu
        
        // Tempo em anos
        let T = daysToMaturity / DIAS_ANO; // Anos
        const T_original = T;
        
        // 🔒 Solução A — Clamping mínimo para evitar instabilidade numérica
        if (T < T_MIN) {
          log(
            SERVICO_NOME,
            "DEBUG_T_CLAMP",
            `T clamped para ${T_MIN} no trade ${idTrade}`,
            `T original = ${T_original}`
          );
          T = T_MIN;
        }
        
        // Parâmetros
        const S = safeFloatConvert(assetData['close']); // Spot
        const K = safeFloatConvert(detail['strike']);   // Strike
        const precoMercado = safeFloatConvert(detail['close']); // Preço da opção
        const qtd = safeFloatConvert(trade['Qtd']) || 1;
        const precoTrade = safeFloatConvert(trade['Preco']);
        
        // Log detalhado para debug (primeiro trade apenas)
        if (processados === 0) {
          log(SERVICO_NOME, "DEBUG_PARAMS", `Parâmetros primeiro trade: ${idTrade}`, 
              `S=${S}, K=${K}, PrecoMercado=${precoMercado}, T=${T.toFixed(4)}, r=${irate}`);
        }
        
        // Tipo
        let flag;
        let optionTypeRaw = trade['Categoria'] || "";
        let optionType = optionTypeRaw.toUpperCase();
        
        if (optionType === 'CALL') {
          flag = 'c';
        } else if (optionType === 'PUT') {
          flag = 'p';
        } else {
          // Fallback
          if (idTrade.indexOf('_V_') !== -1) {
            flag = 'p';
          } else if (idTrade.indexOf('_C_') !== -1) {
            flag = 'c';
          } else {
            erros++;
            continue;
          }
        }
        
        // --- CALCULAR IMPLIED VOLATILITY ---
        let iv = calcularImpliedVolatility(S, K, T, irate, precoMercado, flag);
        
        // Fallback: Se IV inválido, usa volatilidade histórica do ativo
        if (iv <= 0 || iv > 5) {
          const ivHistorica = safeFloatConvert(assetData['iv_current']) / 100;
          
          if (ivHistorica > 0 && ivHistorica <= 5) {
            iv = ivHistorica;
            log(SERVICO_NOME, "INFO_FALLBACK", `Usando IV histórica para: ${idTrade}`, `IV=${iv.toFixed(4)}`);
          } else {
            log(SERVICO_NOME, "AVISO_DADOS", `IV inválido: ${idTrade}`, `IV calculado=${iv}, IV histórico=${ivHistorica}`);
            erros++;
            continue;
          }
        }
        
      // --- CALCULAR GREGAS (modelo base) ---

      // [NOVO] Bloco de Inspeção de Dados
       // Loga os primeiros 5 trades ou qualquer um que tenha erro potencial
       if (processados < 3) { 
         log(SERVICO_NOME, "DEBUG_RAIO_X", 
             `Trade: ${idTrade} (${flag})`, 
             `ENTRADA -> S: ${S} | K: ${K} | T: ${T.toFixed(4)} | r: ${irate} | Vol: ${iv.toFixed(4)}`);
       }

        let gregas = calcularGregasV2(S, K, T, irate, iv, flag);

        // [NOVO] COLE ESTE BLOCO IMEDIATAMENTE ABAIXO DA LINHA ACIMA:
        if (processados < 3) {
             log(SERVICO_NOME, "DEBUG_SAIDA", 
                 `Resultados Math`, 
                 `Price: ${gregas.price} | Rho: ${gregas.rho} | Vega: ${gregas.vega}`);
        }


        // --- Solução B: ajuste near-expiration quando DTE < NEAR_EXPIRATION_DAYS ---
        if (daysToMaturity > 0 && daysToMaturity < NEAR_EXPIRATION_DAYS) {
          const alpha = daysToMaturity / NEAR_EXPIRATION_DAYS;
          gregas = ajustarGregasNearVencimento(gregas, alpha);
          log(
            SERVICO_NOME,
            "DEBUG_NEAR_EXP",
            `Ajuste near-expiration aplicado: ${idTrade}`,
            `DTE=${daysToMaturity}, alpha=${alpha.toFixed(3)}`
          );
        }
        
        // --- CALCULAR MÉTRICAS ADICIONAIS ---
        const moneyness = S / K;
        const moneynessCode = calcularMoneynessCode(S, K, flag);
        
        // --- PREPARAR DADOS PARA MERGE ---
        const dadosCalc = {
          ID_Trade_Unico: idTrade,
          Ativo: ativo,
          ID_Estrutura: detail['ID_Estrutura'],
          Timestamp_Atualizacao: timestamp,
          moneyness_code: moneynessCode,
          moneyness: moneyness,
          price: gregas.price,
          delta: gregas.delta,
          gamma: gregas.gamma,
          vega: gregas.vega,
          theta: gregas.theta,
          rho: gregas.rho,
          volatility: iv,
          poe: gregas.poe
        };
        
        const linhaDados = FIXED_HEADERS_CALC_GREEKS.map(campo => {
          const valor = dadosCalc[campo];
          return valor !== undefined && valor !== null ? valor : "";
        });
        
        // --- LÓGICA DE MERGE ---
        if (idToRowMap[idTrade]) {
          // UPDATE: ID já existe
          const rowIndex = idToRowMap[idTrade];
          for (let col = 0; col < linhaDados.length; col++) {
            updatesInPlace.push({
              row: rowIndex,
              col: col + 1,
              value: linhaDados[col]
            });
          }
          log(SERVICO_NOME, "DEBUG", `Trade atualizado: ${idTrade}`, `Linha: ${rowIndex}`);
        } else {
          // APPEND: ID novo
          appendNew.push(linhaDados);
          log(SERVICO_NOME, "DEBUG", `Trade novo adicionado: ${idTrade}`, "Será inserido no append");
        }
        
        processados++;
        
      } catch (e) {
        log(SERVICO_NOME, "ERRO_DADOS", `Falha ao calcular: ${idTrade}`, e.message);
        erros++;
        continue;
      }
      
      // Rate limit
      if (i < tradesFiltrados.length - 1 && i % 10 === 0) {
        Utilities.sleep(100);
      }
    }
    
    // --- PASSO 5: ATUALIZAR PLANILHA ---
    let celulasAtualizadas = 0;
    let linhasAdicionadas = 0;
    let updatesPorLinha = {};
    
    if (updatesInPlace.length > 0) {
      updatesInPlace.forEach(update => {
        if (!updatesPorLinha[update.row]) {
          updatesPorLinha[update.row] = {};
        }
        updatesPorLinha[update.row][update.col] = update.value;
      });
      
      Object.keys(updatesPorLinha).forEach(row => {
        const rowNum = parseInt(row);
        const valores = updatesPorLinha[row];
        
        const linhaCompleta = [];
        for (let col = 1; col <= FIXED_HEADERS_CALC_GREEKS.length; col++) {
          linhaCompleta.push(valores[col] !== undefined ? valores[col] : "");
        }
        
        abaCalcGreeks.getRange(rowNum, 1, 1, FIXED_HEADERS_CALC_GREEKS.length).setValues([linhaCompleta]);
        celulasAtualizadas += FIXED_HEADERS_CALC_GREEKS.length;
      });
    }
    
    if (appendNew.length > 0) {
      abaCalcGreeks.getRange(
        abaCalcGreeks.getLastRow() + 1,
        1,
        appendNew.length,
        FIXED_HEADERS_CALC_GREEKS.length
      ).setValues(appendNew);
      
      linhasAdicionadas = appendNew.length;
    }
    
    // --- LOG FINAL ---
    const mensagemFinal = `Cálculo concluído! ${processados} gregas calculadas (${Object.keys(updatesPorLinha).length} atualizados, ${linhasAdicionadas} novos). Erros: ${erros}`;
    
    log(SERVICO_NOME, "SUCESSO", mensagemFinal,
        `Células atualizadas: ${celulasAtualizadas}, Linhas adicionadas: ${linhasAdicionadas}`);
    
    // Use o console para registrar o resultado sem interromper a execução automática
    console.log(mensagemFinal);

    /* // REMOVIDO PARA EVITAR ERRO DE PERMISSÃO:
    if (MailApp.getRemainingDailyQuota() > 0 && processados === 0) {
      // Opcional: enviar e-mail apenas em caso de erro crítico
    }
    */


  } catch (erro) {
      const msgErro = "Erro fatal no cálculo: " + erro.message;
      log(SERVICO_NOME, "ERRO_CRITICO", msgErro, erro.stack);
      console.error(msgErro);
      safeAlert('❌ Erro no Cálculo', msgErro);
      throw erro;
    }

}

// ═══════════════════════════════════════════════════════════════
// FUNÇÕES AUXILIARES
// ═══════════════════════════════════════════════════════════════

function getCalcGreeksMapAndHeaders(aba) {
  const ultimaLinha = aba.getLastRow();
  
  let headers = [];
  if (ultimaLinha >= 1) {
    headers = aba.getRange(1, 1, 1, FIXED_HEADERS_CALC_GREEKS.length).getValues()[0];
  }
  
  if (!headers || headers.length < 3) {
    headers = FIXED_HEADERS_CALC_GREEKS;
  }
  
  if (ultimaLinha <= 1) {
    return { headers: headers, idToRowMap: {} };
  }
  
  const ids = aba.getRange(2, 1, ultimaLinha - 1, 1).getValues();
  const idToRowMap = {};
  
  for (let i = 0; i < ids.length; i++) {
    const idTrade = ids[i][0];
    if (idTrade) {
      idToRowMap[idTrade] = i + 2;
    }
  }
  
  return { headers: headers, idToRowMap: idToRowMap };
}

/**
 * Função Auxiliar para comparar arrays (Evita o erro ReferenceError)
 */
function arraysEqual(a, b) {
  if (a === b) return true;
  if (a == null || b == null) return false;
  if (a.length !== b.length) return false;

  for (var i = 0; i < a.length; ++i) {
    if (a[i] !== b[i]) return false;
  }
  return true;
}




/**
 * FUNÇÃO AUXILIAR: Alerta Seguro
 * Evita o erro "Cannot call SpreadsheetApp.getUi() from this context"
 */
function safeAlert(titulo, mensagem) {
  try {
    const ui = SpreadsheetApp.getUi();
    if (ui) {
      ui.alert(titulo, mensagem, ui.ButtonSet.OK);
    }
  } catch (e) {
    // Se estiver em modo trigger (sem UI), apenas loga
    console.log(`${titulo}: ${mensagem}`);
  }
}

/**
 * FUNÇÃO AUXILIAR: Conversão de Números BR
 */
function safeFloatConvert(val) {
  if (val === undefined || val === null || val === "") return 0;
  if (typeof val === 'number') return val;
  
  // Remove R$, espaços e troca vírgula por ponto
  let limpo = val.toString()
    .replace("R$", "")
    .replace(/\s/g, "")
    .replace(".", "") // Remove separador de milhar se houver
    .replace(",", "."); // Troca decimal
    
  let num = parseFloat(limpo);
  return isNaN(num) ? 0 : num;
}





// ═══════════════════════════════════════════════════════════════
// ORQUESTRADOR: CALCULAR GREEKS (NATIVO)
// ═══════════════════════════════════════════════════════════════

/**
 * Orquestra o cálculo de gregas (implementação nativa)
 */
function orquestrarCalcGreeks() {
  const OPERACAO_ID = Utilities.getUuid();
  const inicio = new Date();
  
  logOrquestrador("INICIO", "Iniciando orquestração de cálculo de gregas (nativo)", {
    operacao_id: OPERACAO_ID,
    timestamp: inicio.toISOString(),
    servico: SERVICOS_REGISTRY.CALC_GREEKS.nome
  });
  
  try {
    const validacao = validarPreCondicoes();
    if (!validacao.sucesso) {
      throw new Error("Pré-condições não atendidas: " + validacao.mensagem);
    }
    
    logOrquestrador("INFO", "Pré-condições validadas com sucesso", validacao);
    
    logOrquestrador("INFO", "Chamando serviço: " + SERVICOS_REGISTRY.CALC_GREEKS.funcao, {
      servico: SERVICOS_REGISTRY.CALC_GREEKS.nome
    });
    
    const resultadoServico = executarComTimeout(
      SERVICOS_REGISTRY.CALC_GREEKS.funcao,
      SERVICOS_REGISTRY.CALC_GREEKS.timeout
    );
    
    const fim = new Date();
    const duracao = fim - inicio;
    
    const metricas = {
      operacao_id: OPERACAO_ID,
      duracao_ms: duracao,
      duracao_segundos: (duracao / 1000).toFixed(2),
      timestamp_inicio: inicio.toISOString(),
      timestamp_fim: fim.toISOString(),
      servico: SERVICOS_REGISTRY.CALC_GREEKS.nome,
      status: "SUCESSO"
    };
    
    logOrquestrador("SUCESSO", "Orquestração de cálculo de gregas (nativo) concluída", metricas);
    
    return {
      sucesso: true,
      metricas: metricas,
      resultado: resultadoServico
    };
    
  } catch (erro) {
    const fim = new Date();
    const duracao = fim - inicio;
    
    const erroDetalhado = {
      operacao_id: OPERACAO_ID,
      duracao_ms: duracao,
      mensagem: erro.message,
      stack: erro.stack,
      timestamp: fim.toISOString()
    };
    
    logOrquestrador("ERRO_CRITICO", "Falha na orquestração de cálculo de gregas (nativo)", erroDetalhado);
    
    safeAlert_(
      '❌ Erro no Cálculo de Gregas',
      'Ocorreu um erro:\n\n' +
      erro.message +
      '\n\nVerifique a aba Logs.'
    );
    
    return {
      sucesso: false,
      erro: erroDetalhado
    };
  }
}