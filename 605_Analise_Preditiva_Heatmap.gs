/**
 * SERVIÇO 605: Analise Preditiva Heatmap (Rastreamento Profundo)
 * Foco total em descobrir de ONDE está vindo o erro: "Faça uma seleção em uma coluna"
 */
function gerarAnalisePreditivaHeatmap(diasParam) {
  const SERVICO = "605_Analise_Heatmap";
  
  try {
    console.log("605_FLAG_01: Iniciou função");
    
    let diasPadrao = 45;
    try {
      if (typeof obterConfigsGlobais === "function") {
        const configGlobal = obterConfigsGlobais();
        if (configGlobal && configGlobal["Regra_Dias_Horizonte_Preditivo"]) {
           diasPadrao = parseInt(configGlobal["Regra_Dias_Horizonte_Preditivo"], 10);
        }
      }
    } catch(e) {
      console.warn("605_FLAG_02: Fallback Config_Global -> 45 dias.");
    }

    const dias = diasParam || diasPadrao; 
    console.log("605_FLAG_03: Dias = " + dias);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const NOME_ABA = "Analise_Preditiva_Heatmap";
    
    console.log("605_FLAG_04: Buscando aba " + NOME_ABA);
    let sheet = ss.getSheetByName(NOME_ABA);
    if (!sheet) {
      console.log("605_FLAG_05: Criando aba nova " + NOME_ABA);
      sheet = ss.insertSheet(NOME_ABA);
    }
    
    console.log("605_FLAG_06: Buscando aba consolidada");
    const abaConsolidada = ss.getSheetByName("Pontuacao_Preditiva_Consolidada");
    if (!abaConsolidada) throw new Error("Aba Pontuacao_Preditiva_Consolidada não encontrada");

    console.log("605_FLAG_07: Lendo dados");
    const dados = abaConsolidada.getDataRange().getValues();
    if (dados.length <= 1) {
      console.log("605_FLAG_08: Sem dados na aba, abortando.");
      return;
    }
    
    const rows = dados.slice(1);
    console.log("605_FLAG_09: Dados lidos com sucesso. Qtd linhas: " + rows.length);

    const dashboard = [
      [
        "Ticker", "Timestamp_Calculo", "Prob_Alta", "Prob_Baixa", "Spot_Atual", 
        `Spot_Min_${dias}d`, `Spot_Max_${dias}d`, "Score", "Fator_Dominante", "Status"
      ]
    ];

    const forceNumber = (val) => {
      if (val === "" || val === null || val === undefined) return 0;
      if (typeof val === 'number') return val;
      let str = String(val).replace("R$", "").replace(/%/g, "").replace(/\./g, "").replace(",", ".").trim();
      return parseFloat(str) || 0;
    };

    console.log("605_FLAG_10: Iniciando Loop ForEach");
    rows.forEach(r => {
      const ticker = String(r[0]);
      let pAlta = forceNumber(r[9]);
      let pBaixa = forceNumber(r[10]);
      
      if (pAlta > 1) pAlta = pAlta / 100;
      if (pBaixa > 1) pBaixa = pBaixa / 100;
      
      dashboard.push([
        ticker, r[1], pAlta, pBaixa, forceNumber(r[2]), forceNumber(r[12]),  
        forceNumber(r[13]), forceNumber(r[8]), String(r[14]), String(r[15])        
      ]);
    });
    console.log("605_FLAG_11: Fim do Loop. " + dashboard.length + " arrays montados.");

    console.log("605_FLAG_12: Limpando a aba destino com clearContents");
    sheet.clearContents();
    
    console.log("605_FLAG_13: Gravando dados na planilha");
    sheet.getRange(1, 1, dashboard.length, dashboard[0].length).setValues(dashboard);

    console.log("605_FLAG_14: Aplicando number formats");
    sheet.getRange(2, 3, rows.length, 2).setNumberFormat("0.00%");      
    sheet.getRange(2, 5, rows.length, 3).setNumberFormat('"R$ "#,##0.00'); 
    sheet.getRange(2, 8, rows.length, 1).setNumberFormat("0.00");         

    console.log("605_FLAG_15: Indo chamar o condicional formatter");
    aplicarFormatacaoHeatmap_(sheet, rows.length);
    
    console.log("605_FLAG_16: Processo finalizado no 605, tentando gravar Log SUCESSO");
    
    // Tenta usar a função global "gravarLog" (evitando a palavra log solta)
    if (typeof gravarLog === "function") {
      gravarLog(SERVICO, "SUCESSO", `Dashboard gerado para ${rows.length} ativos.`, `Horizonte: ${dias}d`);
    }

  } catch (e) {
    // 💥 AQUI ESTÁ O QUE NÓS QUEREMOS
    const errorMsg = `🔥 EXPLOSÃO NO 605 🔥\nErro: ${e.message}\nStack: ${e.stack}`;
    
    console.error(errorMsg);
    
    // Tenta avisar o logger global sobre o erro detalhado
    if (typeof gravarLog === "function") {
       gravarLog(SERVICO, "ERRO_FATAL", "Exception capturada dentro do 605", errorMsg);
    }
    
    // Dispara o erro de volta pro 600
    throw new Error(errorMsg); 
  }
}

function aplicarFormatacaoHeatmap_(sheet, rowCount) {
  try {
    console.log("605_FLAG_C1: Dentro do formatter. Limpando regras.");
    sheet.clearConditionalFormatRules();
    
    const regras = [];
    console.log("605_FLAG_C2: Regra gradient");
    const rangeProb = sheet.getRange(2, 3, rowCount, 1);
    regras.push(SpreadsheetApp.newConditionalFormatRule()
      .setGradientMinpoint("#E67C73") 
      .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, "0.5") 
      .setGradientMaxpoint("#57BB8A") 
      .setRanges([rangeProb])
      .build()
    );

    console.log("605_FLAG_C3: Regra status");
    const rangeStatus = sheet.getRange(2, 10, rowCount, 1);
    const coresStatus = {
      "ALTA_CONFLUENCIA_ALTA": "#b7e1cd",
      "TENDENCIA_ALTA_FAVORAVEL": "#d9ead3",
      "STATUS_NEUTRO_AGUARDAR": "#fff2cc",
      "RISCO_OPERACIONAL_ELEVADO": "#f4cccc"
    };

    for (const [texto, cor] of Object.entries(coresStatus)) {
      regras.push(SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains(texto)
        .setBackground(cor)
        .setRanges([rangeStatus])
        .build()
      );
    }

    console.log("605_FLAG_C4: Aplicando setConditionalFormatRules");
    sheet.setConditionalFormatRules(regras);
    console.log("605_FLAG_C5: Condicionais aplicados com sucesso.");

  } catch (err) {
     console.error("605_FLAG_C_ERRO: Quebrou DENTRO do formatter.", err);
     throw err;
  }
}