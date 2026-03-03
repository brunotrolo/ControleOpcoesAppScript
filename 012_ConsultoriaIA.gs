/**
 * MÓDULO: 011_ConsultoriaIA
 * OBJETIVO: Middleman entre o Front-end e o Motor Gemini. 
 * Formata os dados, aplica a Persona Dinâmica (Config_Global) e roteia o JSON.
 */

// =======================================================
// 1. LEITURA DE CONFIGURAÇÕES (MOTOR DE PERSONA)
// =======================================================
function lerConfiguracoesIA() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Config_Global");
    if (!sheet) return {};

    const data = sheet.getDataRange().getValues();
    let config = {};
    
    // Transforma a aba (Chave e Valor) em um dicionário blindado
    data.forEach(row => {
      if (row[0] && typeof row[0] === 'string') {
        // Salva a chave sempre em CAIXA ALTA e sem espaços para evitar bugs de digitação
        config[row[0].trim().toUpperCase()] = row[1];
      }
    });
    
    return config;
  } catch (e) {
    gravarLog("ConsultoriaIA", "ERRO", "Falha ao ler Config_Global", e.toString());
    return {};
  }
}

// =======================================================
// 2. FUNÇÃO PRINCIPAL DO MIDDLEMAN
// =======================================================
function apiAnalisarOperacoesAtivas(operacoes) {
  const SERVICO = "ConsultoriaIA";
  
  if (!operacoes || operacoes.length === 0) {
    gravarLog(SERVICO, "AVISO", "Nenhuma operação recebida para análise.", "");
    flushLogs();
    return { success: false, error: "Nenhuma operação recebida." };
  }

  gravarLog(SERVICO, "INFO", `Iniciando montagem do Super-Prompt para ${operacoes.length} ativos.`, "");

  try {

    // Função auxiliar para ignorar maiúsculas/minúsculas na matriz de dados
    const getSafe = (obj, key) => {
      if (!obj) return "N/D";
      const foundKey = Object.keys(obj).find(k => k.toUpperCase() === key.toUpperCase());
      return foundKey ? obj[foundKey] : "N/D";
    };

    // 1. LIMPEZA E FORMATAÇÃO (Token Optimization)
    const carteiraLimpa = operacoes.map((op, index) => {
      return {
        _id_temp: index, 
        Ativo: getSafe(op, "Ticker") !== "N/D" ? getSafe(op, "Ticker") : getSafe(op, "Código"),
        Tipo: `${getSafe(op, "Venda/Compra")} ${getSafe(op, "Tipo")}`,
        Vencimento_Dias: getSafe(op, "DTE"),
        Moneyness: getSafe(op, "Moneyness"),
        Strike: getSafe(op, "Strike"),
        Spot: getSafe(op, "Spot"),
        PL_Porcentagem: getSafe(op, "P/L TOTAL %"),
        Tendencia: getSafe(op, "Veredito_Tendencia"),
        IFR14: getSafe(op, "IFR14"),
        Dist_M200: getSafe(op, "Dist_MMA200")
      };
    });

    // 2. A PERSONA DINÂMICA (Lida direto da Config_Global)
    const configs = lerConfiguracoesIA();
    
    // Descobre qual é o perfil ativo na planilha (Padrão: EQUILIBRADO)
    const perfilAtivo = String(configs["IA_PERFIL_CONSULTOR"] || "EQUILIBRADO").trim().toUpperCase();
    
    // Puxa as diretrizes baseadas no perfil
    const regrasGerais = configs["PROMPT_REGRAS_GERAIS"] || "Atue como um Gestor de Risco focado em proteção de capital.";
    const promptPerfil = configs[`PROMPT_SISTEMA_${perfilAtivo}`] || "";


// Monta o cérebro da IA (A arquitetura dinâmica e estruturada)
    const systemInstruction = `
${regrasGerais}

${promptPerfil}

REGRA DE ISOLAMENTO E ESTRUTURAÇÃO:
1. PROIBIDO AGRUPAR: Avalie cada "_id_temp" de forma 100% isolada.
2. ESTRUTURA OBRIGATÓRIA DA ANÁLISE: Para cada ativo, você deve obrigatoriamente seguir este modelo de texto:
   - O QUE: [Ação clara: MANTER, RECOMPRAR, ROLAR ou ASSUMIR]
   - QUANDO: [Timing exato: AGORA, PRÓXIMOS DIAS ou NO VENCIMENTO]
   - POR QUE: [Racional técnico denso citando IFR, Tendência ou DTE]

3. FORMATO DE SAÍDA: Retorne estritamente um array de objetos JSON mapeando o "id" ao "_id_temp" e "analise" contendo a estrutura acima.
   Exemplo: [{"id": 0, "analise": "- O QUE: RECOMPRAR\\n- QUANDO: AGORA\\n- POR QUE: Lucro de 94% atingido..."}]
    `.trim();


    // 3. O PROMPT DO USUÁRIO
    const promptUser = `
Audite a seguinte carteira de opções e retorne o JSON mapeado pelo "_id_temp".
Carteira Atual:
${JSON.stringify(carteiraLimpa, null, 2)}
    `.trim();

    // 4. CHAMADA AO MOTOR V8 (010_CoreDebugGemini)
    gravarLog(SERVICO, "INFO", `Disparando requisição Gemini. Perfil Ativo: ${perfilAtivo}`, "");
    
    const respostaIA = callGeminiAI(promptUser, systemInstruction, true);

    if (!respostaIA) {
      throw new Error("O motor Gemini retornou nulo ou falhou no parsing.");
    }

    // 5. SUCESSO E RETORNO
    gravarLog(SERVICO, "SUCESSO", "Análise recebida e parseada com sucesso.", JSON.stringify(respostaIA));
    flushLogs(); 
    
    return { 
      success: true, 
      data: respostaIA 
    };

  } catch (e) {
    gravarLog(SERVICO, "ERRO_FATAL", "Falha na geração da consultoria IA.", e.toString());
    flushLogs();
    return { success: false, error: e.toString() };
  }
}