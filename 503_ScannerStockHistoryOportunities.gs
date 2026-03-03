/**
 * ═══════════════════════════════════════════════════════════════
 * SCANNER TENDÊNCIA OPORTUNIDADES - MOTOR DE CONFLUÊNCIA
 * ═══════════════════════════════════════════════════════════════
 * RESPONSABILIDADES:
 * - Cruzar Selecao_Opcoes com Dados_Ativos_Historico_Tendencia
 * - Calcular Gregas (Delta, Theta, IV) nativamente
 * - Gerar Score Nexo (0-10) e Veredito IA
 * - Identificar Proteções Técnicas (MMA200 e Bollinger)
 * ═══════════════════════════════════════════════════════════════
 */



function executarScannerOportunidades() {
  const SERVICO_NOME = "ScannerOportunidades_v2";
  const OPERACAO_ID = Utilities.getUuid();
  logOrquestrador("INICIO", "Iniciando Scanner V2 - Mapeamento de API detectado", { operacao_id: OPERACAO_ID });

  try {
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const abaTendencia = planilha.getSheetByName("Dados_Ativos_Historico_Tendencia");
    const abaSelecao = planilha.getSheetByName("Selecao_Opcoes");
    const abaDestino = planilha.getSheetByName("Scanner_Tendencia_Oportunidades");
    const abaConfig = planilha.getSheetByName("Config_Global");

    const abaMacro = planilha.getSheetByName("Dados_Macro_Setorial");
    const dadosMacroRaw = abaMacro.getDataRange().getValues();
    
    // Mapeamento dos sensores (ajustado para a estrutura da sua aba)
    const macro = {
      ibovVar:  parseFloat(dadosMacroRaw[2][1]),  // Linha 3 (B3): Ibovespa %
      vix:      parseFloat(dadosMacroRaw[10][1]), // Linha 11 (B11): VIX
      vale:     parseFloat(dadosMacroRaw[6][1]),  // Linha 7 (B7): Vale %
      petroleo: parseFloat(dadosMacroRaw[7][1]), // Linha 8 (B8): Petróleo %
      juros:    parseFloat(dadosMacroRaw[9][1])  // Linha 10 (B10): Juros %
    };

    const configData = abaConfig.getDataRange().getValues();
    const irate = buscarValorConfig(configData, "Taxa_Selic_Anual") || 0.12;

    const dadosTendencia = abaTendencia.getDataRange().getValues();
    const mapaTendencia = mapearDadosTendencia(dadosTendencia);

    const dadosSelecao = abaSelecao.getDataRange().getValues();
    const headers = dadosSelecao[0].map(h => h.toString().toLowerCase().trim());
    
    // --- MAPEAMENTO REAL BASEADO NO SEU EXEMPLO ---
    const colIdx = {
      ticker: headers.indexOf("ticker"),
      symbol: headers.indexOf("symbol"), // Antes era 'opcao'
      strike: headers.indexOf("strike"),
      bid: headers.indexOf("bid"),       // Seu preço de venda
      due_date: headers.indexOf("due_date"),
      days_to_maturity: headers.indexOf("days_to_maturity"),
      type: headers.indexOf("type")
    };

    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);
    const resultadosScanner = [];


    for (let i = 1; i < dadosSelecao.length; i++) {
      const rowArr = dadosSelecao[i];
      const ticker = rowArr[colIdx.ticker];
      const tendencia = mapaTendencia[ticker];
      if (!tendencia) continue;

      let dte = parseInt(rowArr[colIdx.days_to_maturity]);
      if (isNaN(dte) || dte <= 0) continue; 

      const S = tendencia.close;
      const K = parseFloat(rowArr[colIdx.strike].toString().replace(",", "."));
      
      let premioClose = parseFloat(rowArr[headers.indexOf("close")].toString().replace(",", "."));
      let premioBid = parseFloat(rowArr[colIdx.bid].toString().replace(",", "."));
      let premio = (!isNaN(premioClose) && premioClose > 0) ? premioClose : premioBid;
      if (isNaN(premio) || premio <= 0) premio = 0.01;

      const T = dte / 252;
      const symbol = rowArr[colIdx.symbol];
      const flag = rowArr[colIdx.type].toString().toLowerCase() === "put" ? "p" : "c";

      // 🛑 FILTRO DE SEGURANÇA: IGNORAR CALLS 🛑
      if (flag === 'c') continue; 



    // --- CÁLCULO GREGAS ---
      let iv = calcularImpliedVolatility(S, K, T, irate, premio, flag);
      if (iv <= 0 || iv > 5) iv = 0.35; 

      const gregas = calcularGregas(S, K, T, irate, iv, flag);

      // --- CORREÇÃO DE ACESSO E FALLBACK ---
      // Tentamos pegar o theta do objeto. Se vier 0 ou undefined, calculamos o decaimento base.
      let thetaBase = gregas.theta;

      if (!thetaBase || thetaBase === 0) {
        // Recalcula apenas o componente de decaimento (Theta de Black-Scholes) para o log e planilha
        const d1_f = (Math.log(S / K) + (irate + 0.5 * iv * iv) * T) / (iv * Math.sqrt(T));
        const nd1_f = Math.exp(-0.5 * d1_f * d1_f) / Math.sqrt(2 * Math.PI);
        thetaBase = (-(S * nd1_f * iv) / (2 * Math.sqrt(T))) / 365; // Converte para diário
      }


      // --- CÁLCULO DE APOIO ---
      const distSpot = ((K / S - 1) * 100); // Garante que a variável existe
      
      // --- SEGURANÇA E SCORE ---
      let protecao = "SEM PROTEÇÃO";
      if (K < tendencia.banda_inf && K < tendencia.mma200) protecao = "🛡️ PROTEÇÃO DUPLA";
      else if (K < tendencia.banda_inf) protecao = "📊 ABAIXO BANDA";
      else if (K < tendencia.mma200) protecao = "📉 ABAIXO MMA200";

      const scoreObj = calcularScoreNexo(tendencia, gregas, distSpot, protecao);

      // --- AJUSTE MACRO E INFORMAÇÃO QUALITATIVA ---
      let scoreFinal = scoreObj.score;
      let motivos = [];

      // 1. Penalidade VIX (Medo Global)
      if (macro.vix > 25) {
        scoreFinal -= 2.0;
        motivos.push("⚠️ VIX ALTO");
      }

      // 2. Confluência Setorial (Bônus para Petróleo e Siderurgia)
      if ((ticker === "PETR4" || ticker === "BRAV3") && macro.petroleo > 0.005) {
        scoreFinal += 1.0;
        motivos.push("🛢️ PETRÓLEO+");
      }
      if ((ticker === "CSNA3" || ticker === "USIM5") && macro.vale > 0.005) {
        scoreFinal += 1.0;
        motivos.push("🏗️ VALE/MINÉRIO+");
      }

      // 3. Confluência IBOV (Humor do Mercado)
      if (macro.ibovVar > 0.01) {
        scoreFinal += 0.5;
        motivos.push("📈 IBOV BULL");
      } else if (macro.ibovVar < -0.01) {
        scoreFinal -= 1.0;
        motivos.push("📉 IBOV BEAR");
      }

      // Garante que o score fique entre 0 e 10
      scoreFinal = Math.min(10, Math.max(0, scoreFinal));

      // Define a nova informação qualitativa (Coluna 24)
      let detalheScore = motivos.length > 0 ? motivos.join(" | ") : "📊 ANÁLISE TÉCNICA";


      // --- PUSH ATUALIZADO (LINHA A LINHA COM 25 COLUNAS) ---
      resultadosScanner.push([
        ticker,                 // Coluna 1: Ticker do Ativo
        new Date(),             // Coluna 2: Timestamp do Scanner
        symbol,                 // Coluna 3: Código da Opção
        S,                      // Coluna 4: Preço Spot
        K,                      // Coluna 5: Strike
        tendencia.veredito,     // Coluna 6: Tendência
        tendencia.ifr14,        // Coluna 7: IFR14
        tendencia.dist_mma200,  // Coluna 8: Distância MMA200 %
        tendencia.banda_inf,    // Coluna 9: Banda Inferior
        tendencia.largura_banda,// Coluna 10: Largura de Banda
        gregas.delta,           // Coluna 11: Delta
        (thetaBase * 365 * 100),// Coluna 12: Theta Real
        iv,                     // Coluna 13: IV
        rowArr[colIdx.due_date],// Coluna 14: Vencimento
        dte,                    // Coluna 15: DTE
        distSpot,               // Coluna 16: Distância Strike %
        protecao,               // Coluna 17: Proteção Técnica
        (K - premio),           // Coluna 18: Breakeven
        premio,                 // Coluna 19: Prêmio Ref.
        (premio / K),           // Coluna 20: ROI %
        (premio / K / dte),     // Coluna 21: Taxa Diária
        tendencia.vol_relativo, // Coluna 22: Vol. Relativo
        scoreFinal,             // Coluna 23 (W): Score Final Ajustado
        detalheScore,           // Coluna 24 (X): Detalhe Qualitativo (NOVA!)
        "",                     // Coluna 25 (Y): ESPAÇO VAZIO PARA 'perfil_risco'
        scoreObj.vereditoIA     // Coluna 26 (Z): Veredito Técnico IA
      ]);

    }


    if (resultadosScanner.length > 0) {
      abaDestino.clearContents();
      
      // 1. Lista de Cabeçalhos Atualizada (26 colunas)
      const headersScanner = [
        "ticker", "timestamp_scanner", "opcao", "spot_price", "strike", 
        "veredito_tendencia", "ifr14", "dist_mma200_%", "banda_inf", "largura_banda", 
        "delta", "theta_real", "iv_rank", "vencimento", "dte", 
        "dist_spot_%", "protecao_tecnica", "breakeven", "premio_ref", "roi_percent", 
        "taxa_diaria", "vol_relativo", "score", "detalhe_score", 
        "perfil_risco", "veredito_ia" // <-- AGORA SÃO 26 COLUNAS
      ];
      
      // 2. Gravação dos dados na Planilha (Atualizado para 26 colunas)
      abaDestino.getRange(1, 1, 1, headersScanner.length).setValues([headersScanner]);
      abaDestino.getRange(2, 1, resultadosScanner.length, headersScanner.length).setValues(resultadosScanner);
      
      // 3. Formatação das Colunas
      // Delta e Theta (Colunas 11 e 12) -> 4 casas decimais
      abaDestino.getRange(2, 11, resultadosScanner.length, 2).setNumberFormat("0.0000"); 
      
      // IV (Coluna 13) -> Porcentagem
      abaDestino.getRange(2, 13, resultadosScanner.length, 1).setNumberFormat("0.00%"); 
      
      // Prêmio, ROI e Taxa Diária (Colunas 19, 20 e 21) -> Porcentagem
      abaDestino.getRange(2, 19, resultadosScanner.length, 3).setNumberFormat("0.00%"); 
      
      // Score (Coluna 23) -> 1 casa decimal
      abaDestino.getRange(2, 23, resultadosScanner.length, 1).setNumberFormat("0.0"); 

      // NOVO: Perfil de Risco (Coluna 25) -> Centralizar
      abaDestino.getRange(2, 25, resultadosScanner.length, 1).setHorizontalAlignment("center");

      // NOVO: Veredito IA (Coluna 26) -> Quebra de texto automática (Wrap)
      // Isso ajuda a ler o texto denso que a IA vai gerar sem cortar as palavras
      abaDestino.getRange(2, 26, resultadosScanner.length, 1).setWrap(true);
      
      logOrquestrador("SUCESSO", "Scanner finalizado com Nova Coluna Qualitativa e Ajustes Macro.");
    }

  } catch (e) {
    logOrquestrador("ERRO_CRITICO", "Erro no Scanner: " + e.message, { stack: e.stack });
  }
}



/**
 * Lógica de Pontuação Nexo - Versão de Auditoria
 */
function calcularScoreNexo(tend, greg, dist, prot) {
  let s = 0;
  
  // 1. Tendência (Peso 3pts)
  // Se o veredito for ALTA, ganha 3. Se for BAIXA, ganha 0.
  if (tend.veredito === "ALTA") s += 3;
  else if (tend.veredito === "LATERAL") s += 1.5;

  // 2. Proteção Técnica (Peso 3pts)
  if (prot === "🛡️ PROTEÇÃO DUPLA") s += 3;
  else if (prot !== "SEM PROTEÇÃO") s += 1.5;

  // 3. Delta (Peso 2pts) - Sweet Spot para Venda de PUT
  const absDelta = Math.abs(greg.delta);
  if (absDelta >= 0.10 && absDelta <= 0.18) s += 2;
  else if (absDelta < 0.10) s += 1;

  // 4. IFR (Peso 2pts) - Indica sobrevenda
  if (tend.ifr14 < 40) s += 2;
  else if (tend.ifr14 < 55) s += 1;

  // Limite de segurança
  s = Math.min(10, Math.max(0, s));

  // Determinação do Veredito IA baseado no Score
  let v = "AGUARDAR";
  if (s >= 8) v = "💎 COMPRA FORTE";
  else if (s >= 6) v = "✅ BOA OPORTUNIDADE";
  else if (s >= 4) v = "⚠️ RISCO MODERADO";
  else v = "❌ EVITAR";

  // Retornamos o Score como número explícito para evitar o bug da data
  return { 
    score: parseFloat(s.toFixed(1)), 
    vereditoIA: v 
  };
}









// --- FUNÇÕES AUXILIARES DE MAPEAMENTO ---

function mapearDadosTendencia(matriz) {
  const mapa = {};
  const h = matriz[0];
  for (let i = 1; i < matriz.length; i++) {
    const ticker = matriz[i][0];
    mapa[ticker] = {
      close: matriz[i][9],
      veredito: matriz[i][20],
      ifr14: matriz[i][15],
      dist_mma200: matriz[i][16],
      banda_inf: matriz[i][18],
      mma200: matriz[i][14],
      largura_banda: matriz[i][19],
      vol_relativo: matriz[i][23]
    };
  }
  return mapa;
}

function converterLinhaParaObjeto(linha, headers) {
  const obj = {};
  headers.forEach((h, i) => obj[h] = linha[i]);
  return obj;
}

function buscarValorConfig(data, chave) {
  const row = data.find(r => r[0] === chave);
  return row ? parseFloat(row[1]) : null;
}