/**
 * MÓDULO: 010_CoreServiceIA
 * OBJETIVO: Motor centralizado para chamadas ao Gemini 2.5 Flash
 */

const IA_CONFIG = {
  MODELO: "gemini-2.5-flash", // Centralizado: mude aqui e mude em todo o sistema
  TEMPERATURE: 0.2,
  TOP_P: 0.8
};

/**
 * Função Mestre para todos os módulos (402, 504, 301, etc.)
 * @param {string} promptText - O prompt do usuário
 * @param {string} systemInstruction - (Opcional) Define quem a IA é
 * @param {boolean} isJsonResponse - Se true, força e limpa o retorno para JSON
 */
function callGeminiAI(promptText, systemInstruction = "", isJsonResponse = false) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  const SERVICO = "CoreServiceIA";
  
  if (!apiKey) {
    log(SERVICO, "ERRO_CONFIG", "API Key não encontrada nas propriedades do script.", "");
    return null;
  }

  const url = `https://generativelanguage.googleapis.com/v1beta/models/${IA_CONFIG.MODELO}:generateContent?key=${apiKey}`;
  
  // Estrutura o payload (Gemini aceita instruções de sistema separadas para melhor performance)
  const payload = {
    contents: [{ parts: [{ text: promptText }] }],
    generationConfig: {
      temperature: IA_CONFIG.TEMPERATURE,
      topP: IA_CONFIG.TOP_P,
      responseMimeType: isJsonResponse ? "application/json" : "text/plain"
    }
  };

  if (systemInstruction) {
    payload.systemInstruction = { parts: [{ text: systemInstruction }] };
  }

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const resText = response.getContentText();
    
    if (response.getResponseCode() !== 200) {
      log(SERVICO, "ERRO_API", `Status ${response.getResponseCode()}`, resText);
      return null;
    }

    const json = JSON.parse(resText);
    let rawContent = json.candidates[0].content.parts[0].text;

    if (isJsonResponse) {
      // Limpeza blindada de Markdown que a IA costuma enviar
      rawContent = rawContent.replace(/```json/g, "").replace(/```/g, "").trim();
      try {
        return JSON.parse(rawContent);
      } catch (e) {
        log(SERVICO, "ERRO_PARSING", "IA não retornou JSON válido. Tentando extração por RegExp.", rawContent);
        const match = rawContent.match(/\[.*\]|\{.*\}/s);
        return match ? JSON.parse(match[0]) : null;
      }
    }

    return rawContent;

  } catch (e) {
    log(SERVICO, "ERRO_FATAL", "Falha na comunicação IA", e.toString());
    return null;
  }
}








function testeUniversalIA() {
  console.log("--- TESTE 1: Resposta em Texto ---");
  const txt = callGeminiAI("Responda apenas 'SISTEMA OK'", "Você é um monitor de sistema.");
  console.log("Resultado Texto: " + txt);

  console.log("--- TESTE 2: Resposta em JSON (Estruturado) ---");
  const promptJson = "Retorne um objeto com 'status': 'ativo' e 'data': 'hoje'";
  const obj = callGeminiAI(promptJson, "Retorne apenas JSON.", true);
  
  if (obj && obj.status) {
    console.log("✅ SUCESSO: Motor central processou JSON corretamente!");
    console.log(obj);
  } else {
    console.error("❌ FALHA: O motor não conseguiu parsear o JSON.");
  }
}