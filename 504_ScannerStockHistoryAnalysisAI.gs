/**
 * MÓDULO: 504_ScannerStockHistoryAnalysisAI
 * OBJETIVO: Gerar análise qualitativa em lote (JSON) usando o Motor Central 010.
 */

function gerarAnaliseIA_Oportunidades() {
  const SERVICO_NOME = "Scanner_AI_v2.0";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaScanner = ss.getSheetByName("Scanner_Tendencia_Oportunidades");
  const abaMacro = ss.getSheetByName("Dados_Macro_Setorial");

  const dados = abaScanner.getDataRange().getValues();
  if (dados.length <= 1) return;

  const rows = dados.slice(1);

  // 1. CAPTURA CONTEXTO MACRO (Enriquecimento da tese)
  const dadosMacro = abaMacro.getDataRange().getValues();
  const macro = {
    ibov: dadosMacro[2][1],
    dolar: dadosMacro[3][1],
    vix: dadosMacro[10][1],
    setor_vale: dadosMacro[6][1],
    setor_petro: dadosMacro[7][1]
  };

  // 2. SELEÇÃO DO TOP 9 (3 Conservadores, 3 Moderados, 3 Agressivos)
  let selecionados = [];
  const filtrar = (minDelta, maxDelta, qtd) => {
    return rows.filter(r => {
      let d = Math.abs(parseFloat(r[10])); 
      return d >= minDelta && d < maxDelta;
    }).sort((a, b) => b[22] - a[22]).slice(0, qtd);
  };

  selecionados = selecionados.concat(filtrar(0, 0.12, 3));    // Conservadores
  selecionados = selecionados.concat(filtrar(0.12, 0.18, 3)); // Moderados
  selecionados = selecionados.concat(filtrar(0.18, 0.25, 3)); // Agressivos

  if (selecionados.length === 0) return;

  // 3. MONTAGEM DOS DADOS PARA O BATCH
  const listaAtivosParaIA = selecionados.map(r => 
    `Ativo: ${r[0]} | Strike: ${r[4]} | Delta: ${r[10]} | ROI: ${r[19]} | Score: ${r[22]} | Contexto: ${r[23]}`
  ).join("\n");

  // 4. PERSONA E PROMPT ESTRATÉGICO (Sua versão lapidada)
  const personaIA = `Você é o Head de Estratégia de Derivativos. Analise oportunidades de Short Put. 
    REGRAS: 1. NÃO use saudações. 2. NÃO se apresente. 3. Comece direto no ponto.`;

  const promptFinal = `Analise estas 9 oportunidades de Short Put (Venda de Put) com foco em geração de renda e proteção de capital.
    
    CENÁRIO MACRO: Ibovespa ${macro.ibov}, VIX ${macro.vix}, Câmbio ${macro.dolar}.

    TAREFA:
    Para cada ativo, forneça um racional denso (3 a 5 linhas) relacionando o ROI oferecido com a segurança do Strike. 
    Seja cético. Se o Score for alto, valide se o cenário macro (Petróleo/Minério) realmente apoia a tese.
    Identifique se é uma oportunidade de "Alta Convicção" ou um "Risco Assimétrico Elevado".

    FORMATO OBRIGATÓRIO (JSON):
    Retorne apenas um array de objetos: [{"perfil": "🟢 Conservador", "texto": "..."}, {"perfil": "🟡 Moderado", "texto": "..."}, {"perfil": "🔴 Agressivo", "texto": "..."}] 
    Mantenha EXATAMENTE a mesma ordem dos ativos fornecidos.

    DADOS:
    ${listaAtivosParaIA}`;

  // 5. CHAMADA AO MOTOR CENTRAL (010) - Solicitando JSON estruturado
  const vereditos = callGeminiAI(promptFinal, personaIA, true);

  // 6. ESCRITA DOS RESULTADOS NA PLANILHA
  if (vereditos && Array.isArray(vereditos)) {
    selecionados.forEach((linhaOriginal, index) => {
      const tickerOpcao = linhaOriginal[2]; // Coluna C
      const idxPlanilha = dados.findIndex(r => r[2] === tickerOpcao);
      
      if (idxPlanilha !== -1 && vereditos[index]) {
        // Coluna 25 (Y): Perfil | Coluna 26 (Z): Racional
        const rangePerfil = abaScanner.getRange(idxPlanilha + 1, 25);
        const rangeTexto = abaScanner.getRange(idxPlanilha + 1, 26);

        rangePerfil.setValue(vereditos[index].perfil).setFontWeight("bold").setHorizontalAlignment("center");
        rangeTexto.setValue(vereditos[index].texto).setBackground("#fdf7e3").setFontStyle("italic").setWrap(true);
      }
    });
    log(SERVICO_NOME, "SUCESSO", `Lote de ${vereditos.length} análises processado.`, "");
  } else {
    log(SERVICO_NOME, "ERRO", "IA não retornou o formato JSON esperado ou falhou.", "");
  }
}