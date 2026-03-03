/**
 * NEXO MARKET SCANNER - v2.2 (REVISÃO BRAVO - LOGS INTEGRADOS)
 * Foco: Manter lógica homologada e aplicar logs institucionais detalhados.
 */

function runMarketScannerEngine() {
  const SERVICO_NOME = "Scanner_Engine_v2.2";
  log(SERVICO_NOME, "INICIO", "Iniciando varredura institucional...", "");

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configData = getRecordsAsMap_(ss.getSheetByName("Config_Global"), "Chave", "Valor");
    
    const ivRankMin = safeFloatConvert(configData['Regra_IV_Rank_Min'] || "0,5");
    const dMin = safeFloatConvert(configData['Regra_Delta_PUT_Entrada_Min'] || "-0,10");
    const dMax = safeFloatConvert(configData['Regra_Delta_PUT_Entrada_Max'] || "-0,50");
    const selic = safeFloatConvert(configData['Taxa_Selic_Anual'] || "0,15");
    let targetVenc = configData['Regra_Vencimento_Entrada_Opcoes'];
    let targetVencStr = (targetVenc instanceof Date) ? Utilities.formatDate(targetVenc, "GMT-3", "yyyy-MM-dd") : targetVenc;

    // 1. TRIAGEM DE ATIVOS (PRESERVADO: Proteção CSNA3)
    const ativosRecords = getAllRecordsSafely(ss.getSheetByName("Dados_Ativos"));
    const ativosAlvoMap = {};
    let aceitosArr = []; let rejeitadosArr = [];

    ativosRecords.forEach(row => {
      let ticker = row['Ticker'];
      if (!ticker) return;

      let ivRank = safeFloatConvert(row['iv_1y_rank']);
      if (ivRank > 1) ivRank = ivRank / 100;

      let spotRaw = row['close'];
      // CORREÇÃO: Usa nxNum para bloquear datas (46032) transformando-as em 0
      let spotFinal = nxNum(spotRaw);


      if (ivRank >= ivRankMin && spotFinal > 0) {
        ativosAlvoMap[ticker] = { 
          iv: safeFloatConvert(row['iv_current']) / 100, 
          spot: spotFinal,
          ivRankPerc: (ivRank * 100).toFixed(0) + "%"
        };
        aceitosArr.push(`${ticker} (${(ivRank*100).toFixed(0)}%)`);
      } else {
        rejeitadosArr.push(`${ticker} (${(ivRank*100).toFixed(0)}%)`);
      }
    });

    // LOG: Triagem com Auditoria de Tickers
    log(SERVICO_NOME, "AUDITORIA_TICKERS", "Triagem concluída", 
        `✅ ACEITOS: ${aceitosArr.join(", ")} | ❌ REJEITADOS: ${rejeitadosArr.slice(0,8).join(", ")}`);

    // 2. VARREDURA DE OPÇÕES
    const opcoesRecords = getAllRecordsSafely(ss.getSheetByName("Selecao_Opcoes"));
    let achados = [];
    let stats = { analisadas: 0, deltaFora: 0 };

    opcoesRecords.forEach(opt => {
      if (!ativosAlvoMap[opt['ticker']]) return;
      let rowVenc = (opt['due_date'] instanceof Date) ? Utilities.formatDate(opt['due_date'], "GMT-3", "yyyy-MM-dd") : opt['due_date'];

      if (opt['type'] === 'PUT' && rowVenc === targetVencStr) {
        stats.analisadas++;
        const details = ativosAlvoMap[opt['ticker']];
        const K = safeFloatConvert(opt['strike']);
        const premio = safeFloatConvert(opt['close']);
        const dte = safeFloatConvert(opt['days_to_maturity']);
        
        const g = calcularGregas(details.spot, K, dte/252, selic, details.iv, 'p');
        
 
      if (g.delta <= dMin && g.delta >= dMax) {
          
        // --- CORREÇÃO: CÁLCULOS EM DECIMAL PURO ---
        // Removemos o "* 100" daqui. O Google Sheets vai multiplicar visualmente depois.
        const roi = (premio / K); 
        const dist = ((details.spot - K) / details.spot);
        const taxaDiaria = roi / dte; 
        
        // Ajuste no Score: Multiplicamos por 100 AQUI DENTRO para manter a escala do Score correta
        let absDelta = Math.abs(g.delta);
        const score = (absDelta > 0) ? ((roi*100) / absDelta) * ((dist*100) / 10) : 0;

        achados.push({ 
          ticker: opt['ticker'], spot: details.spot, ivRank: details.ivRankPerc,
          symbol: opt['symbol'], strike: K, premio: premio, venc: rowVenc, dte: dte,
          delta: g.delta, gamma: g.gamma, theta: g.theta, vega: g.vega,
          roi: roi, taxaDiaria: taxaDiaria, breakeven: K - premio,
          dist: dist, score: score
          });
        } else {
          stats.deltaFora++;
        }
      }
    });

    log(SERVICO_NOME, "SUCESSO", "Varredura finalizada", `Total Analisadas: ${stats.analisadas} | Sucessos: ${achados.length}`);

    // LOG: Relatório Estratégico Detalhado
    let dataRef = "Vencimento";
    try { 
      let p = targetVencStr.split('-');
      dataRef = Utilities.formatDate(new Date(p[0], p[1]-1, p[2]), "GMT-3", "MMM/yy"); 
    } catch(e) {}
    
    let relatorio = `📊 ${stats.analisadas} Opções Analisadas (Puts de ${Object.keys(ativosAlvoMap).join(", ")} para ${dataRef}).\n` +
                    `✅ ${achados.length} Sucessos: Delta entre ${dMin} e ${dMax}.\n` +
                    `⚠️ ${stats.deltaFora} Delta Fora: Opções descartadas por risco.\n\n` +
                    `🧠 Auditoria Nexo: Rejeitados ${rejeitadosArr.slice(0,6).join(", ")}... por IV Rank < ${(ivRankMin*100).toFixed(0)}%.`;
    
    log(SERVICO_NOME, "SUCESSO", "Relatório Estratégico Gerado", relatorio);

    // 3. ATUALIZAÇÃO UI
    atualizarAbaScanner(achados);
    
    // 4. CHAMADA IA
    gerarVereditosScanner();

    return achados;

  } catch (e) {
    log(SERVICO_NOME, "ERRO_CRITICO", "Falha na Engine v2.2", e.message);
  }
}


function atualizarAbaScanner(oportunidades) {
  const SERVICO_NOME = "Scanner_UI_v2.2";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Scanner_Oportunidades");
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  
  // 1. Limpeza Segura (Deleta dados antigos)
  if (lastRow > 1) {
    try {
      sheet.deleteRows(2, lastRow - 1);
    } catch(e) {
      // Se falhar ao deletar (raro), tenta limpar
      sheet.getRange(2, 1, lastRow, 20).clearContent();
    }
  }

  if (!oportunidades || oportunidades.length === 0) {
    log(SERVICO_NOME, "AVISO", "Nenhuma oportunidade para salvar.", "");
    return;
  }

  const dadosFinais = oportunidades.map(op => {
    let perfil = (Math.abs(op.delta) <= 0.16) ? "🟢 Conservador" : (Math.abs(op.delta) >= 0.25) ? "🔴 Agressivo" : "🟡 Equilibrado";
    return [
      0, op.ticker, op.spot, op.ivRank, op.symbol, op.strike,
      op.premio, op.venc, op.dte, perfil,
      op.delta, op.gamma, op.theta, op.vega,
      op.roi, op.taxaDiaria, op.breakeven, op.dist, op.score,
      "Análise pendente..."
    ];
  });

  dadosFinais.sort((a, b) => b[18] - a[18]).forEach((row, i) => row[0] = i + 1);
  
  const numLinhas = dadosFinais.length;
  const range = sheet.getRange(2, 1, numLinhas, 20);
  
  // 2. Salva os dados (O MAIS IMPORTANTE)
  range.setValues(dadosFinais);
  
  // 3. Formatação em Bloco TRY/CATCH (Se falhar, o programa NÃO PARA)
  try {
    // Tenta limpar validações antigas que podem estar travando a coluna
    range.clearDataValidations(); 
    
    sheet.getRange(2, 3, numLinhas, 1).setNumberFormat("R$ #,##0.00"); // Spot
    sheet.getRange(2, 6, numLinhas, 2).setNumberFormat("R$ #,##0.00"); // Strike, Premio
    sheet.getRange(2, 17, numLinhas, 1).setNumberFormat("R$ #,##0.00"); // Breakeven
    sheet.getRange(2, 15, numLinhas, 2).setNumberFormat("0.00%"); // ROI, Taxa
    sheet.getRange(2, 18, numLinhas, 1).setNumberFormat("0.00%"); // Dist
    sheet.getRange(2, 11, numLinhas, 4).setNumberFormat("0.0000"); // Gregas
    sheet.getRange(2, 19, numLinhas, 1).setNumberFormat("0.00");   // Score
    sheet.getRange(2, 8, numLinhas, 1).setNumberFormat("dd/MM/yyyy"); // Vencimento
    
    log(SERVICO_NOME, "SUCESSO", "Dashboard atualizado e formatado.", "");
  } catch (e) {
    // Se der erro de "Coluna com Tipo", apenas avisamos no log, mas o código segue!
    log(SERVICO_NOME, "ALERTA_VISUAL", "Dados salvos, mas o Google Sheets bloqueou a formatação.", e.message);
  }
}



// --- UTILITÁRIOS (PRESERVADOS) ---

function log(servico, status, acao, detalhe) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName("Logs"); 
  const timestamp = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm:ss");
  if (logSheet) logSheet.appendRow([timestamp, servico, status, acao, detalhe]);
  console.log(`${timestamp} | ${servico} | ${status} | ${acao} | ${detalhe}`);
}

function getRecordsAsMap_(sheet, k, v) {
  const d = sheet.getDataRange().getValues();
  const h = d[0]; const ki = h.indexOf(k); const vi = h.indexOf(v);
  const m = {};
  for (let i = 1; i < d.length; i++) { if (d[i][ki]) m[d[i][ki]] = d[i][vi]; }
  return m;
}

function getAllRecordsSafely(sheet) {
  if (!sheet) return [];
  const d = sheet.getDataRange().getValues();
  const h = d[0];
  return d.slice(1).map(r => {
    let o = {}; h.forEach((header, i) => o[header] = r[i]);
    return o;
  });
}

function safeFloatConvert(v) {
  if (!v) return 0;
  if (typeof v === 'number') return v;
  let c = v.toString().replace(/[R$\s]/g, '').replace(/\./g, '').replace(',', '.');
  return parseFloat(c) || 0;
}


function calcularGregas(S, K, T, r, sigma, type) {
  if (T <= 0 || sigma <= 0) return { delta: 0, gamma: 0, theta: 0, vega: 0 };
  
  const sqrtT = Math.sqrt(T);
  const d1 = (Math.log(S / K) + (r + Math.pow(sigma, 2) / 2) * T) / (sigma * sqrtT);
  const d2 = d1 - sigma * sqrtT; // d2 é essencial para o Theta

  // nD1 = PDF (densidade) de d1
  const nD1 = (1 / Math.sqrt(2 * Math.PI)) * Math.exp(-Math.pow(d1, 2) / 2);
  
  // Função auxiliar para a CDF (acumulada) usando sua errorFunction
  const cdf = (x) => 0.5 * (1 + errorFunction(x / Math.sqrt(2)));
  const Nd2 = cdf(d2);

  let delta, theta;

  if (type === 'p') {
    delta = cdf(d1) - 1;
    // Cálculo do Theta para PUT
    const termo1 = -(S * nD1 * sigma) / (2 * sqrtT);
    const termo2 = r * K * Math.exp(-r * T) * (1 - Nd2);
    theta = (termo1 + termo2) / 252; // Dividido por 252 para valor DIÁRIO
  } else {
    delta = cdf(d1);
    // Cálculo do Theta para CALL
    const termo1 = -(S * nD1 * sigma) / (2 * sqrtT);
    const termo2 = -r * K * Math.exp(-r * T) * Nd2;
    theta = (termo1 + termo2) / 252;
  }

  const gamma = nD1 / (S * sigma * sqrtT);
  const vega = (S * sqrtT * nD1) / 100;

  return { delta, gamma, theta, vega };
}


function errorFunction(x) {
  const t = 1 / (1 + 0.5 * Math.abs(x));
  const ans = 1 - t * Math.exp(-x * x - 1.26551223 + t * (1.00002368 + t * (0.37409196 + t * (0.09678418 + t * (-0.18628806 + t * (0.27886807 + t * (-1.13520398 + t * (1.48851587 + t * (-0.82215223 + t * 0.17087277)))))))));
  return x >= 0 ? ans : -ans;
}



// ═══════════════════════════════════════════════════════════════
// UTILITÁRIOS MESTRES (PREFIXO NX - PROTEÇÃO GLOBAL)
// ═══════════════════════════════════════════════════════════════

/**
 * Log Institucional Blindado
 */
function nxLog(servico, status, acao, detalhe) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName("Logs"); 
    const timestamp = new Date();
    
    if (logSheet) {
      logSheet.appendRow([timestamp, servico, status, acao, detalhe || ""]);
    }
    console.log(`${servico} | ${status} | ${acao}`);
  } catch (e) {
    Logger.log("Falha no log nx: " + e.message);
  }
}

/**
 * Vacina Anti-Bug 46366 (Conversão Numérica com Proteção)
 */
function nxNum(v) {
  if (!v && v !== 0) return 0;
  if (v instanceof Date) return 0; // Bloqueia data serial
  if (typeof v === 'number') return (v > 30000) ? 0 : v; // Proteção contra datas convertidas em número
  
  let c = v.toString().replace(/[R$\s]/g, '').replace(/\./g, '').replace(',', '.');
  let n = parseFloat(c);
  return (n > 30000) ? 0 : (n || 0);
}

/**
 * Utilitários de Mapeamento de Dados
 */
function nxGetMap(sheet, k, v) {
  const d = sheet.getDataRange().getValues();
  const h = d[0]; const ki = h.indexOf(k); const vi = h.indexOf(v);
  const m = {};
  for (let i = 1; i < d.length; i++) { if (d[i][ki]) m[d[i][ki]] = d[i][vi]; }
  return m;
}

function nxGetRecords(sheet) {
  if (!sheet) return [];
  const d = sheet.getDataRange().getValues();
  const h = d[0];
  return d.slice(1).map(r => {
    let o = {}; h.forEach((header, i) => o[header] = r[i]);
    return o;
  });
}