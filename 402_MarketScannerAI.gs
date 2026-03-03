/**
 * NEXO MARKET SCANNER - AI v1.5 (NEXO VOICE)
 * OBJETIVO: Vereditos técnicos sincronizados com o Motor Central 010.
 */

function gerarVereditosScanner() {
  const SERVICO_NOME = "Nexo_Voice_AI_v1.5";
  
  // Usa o log centralizado que agora mora no 002
  log(SERVICO_NOME, "INICIO", "Iniciando análise das pérolas do dia...", "");

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Busca configurações globais (usando o helper nxGetMap do seu Engine)
    const configData = nxGetMap(ss.getSheetByName("Config_Global"), "Chave", "Valor");
    const qtdPorPerfil = nxNum(configData['Regra_Qtd_Analises_IA_Por_Perfil'] || "3");

    const sheet = ss.getSheetByName("Scanner_Oportunidades");
    const dados = sheet.getDataRange().getValues();
    const headers = dados.shift(); // Remove cabeçalho

    const perfisValidos = ["🟢 Conservador", "🟡 Equilibrado", "🔴 Agressivo"];
    let perolas = [];
    
    // Filtragem das melhores oportunidades por perfil
    perfisValidos.forEach(p => {
      let filtrados = dados.filter(r => r[9] === p).slice(0, qtdPorPerfil);
      perolas = perolas.concat(filtrados);
    });

    if (perolas.length === 0) {
      log(SERVICO_NOME, "AVISO", "Nenhuma oportunidade qualificada para análise IA.", "");
      return;
    }

    // Definição da Identidade da IA (Persona) - Enviada como System Instruction no 010
    const personaIA = `Você é o Head de Estratégia de Opções da Nexo. 
    Seu objetivo é dar vereditos técnicos e secos. 
    REGRAS: 1. NÃO use saudações. 2. NÃO se apresente. 3. Comece direto no ponto.`;

    perolas.forEach(linha => {
      // 1. MAPEAMENTO DE DADOS
      const rank   = linha[0];
      const ticker = linha[1];
      const spot   = linha[2];
      const ivRank = linha[3];
      const strike = linha[5];
      const perfil = linha[9];
      const delta  = linha[10];
      const roi    = linha[14];
      const dist   = linha[17];
      const score  = linha[18];


     // 2. PROMPT "NEXO INSTITUCIONAL" (Detalhado e Analítico)
      const promptIA = `Você é o Head de Estratégia de Opções.
      Analise analiticamente esta operação de VENDA DE PUT (Short Put).

      DADOS TÉCNICOS:
      - Ativo: ${ticker} (Spot R$ ${spot} | Strike R$ ${strike})
      - Perfil: ${perfil}
      - Volatilidade (IV Rank): ${ivRank}
      - Delta: ${delta} | ROI: ${(roi*100).toFixed(2)}% | Margem: ${(dist*100).toFixed(2)}%
      - Nexo Score: ${(score).toFixed(2)} (Eficiência Matemática)

      REGRAS DE ESTILO (CRÍTICO):
      1. NÃO use saudações (ex: "Prezado investidor", "Olá").
      2. NÃO se apresente (ex: "Como Head de Estratégia...", "Minha análise é...").
      3. COMECE O TEXTO IMEDIATAMENTE com a análise técnica.

      DIRETRIZES DE ANÁLISE:
      1. Se CONSERVADOR: Valide se a Margem de Segurança compensa o retorno.
      2. Se AGRESSIVO: Avalie se o IV Rank e Prêmio justifica o risco assumido.
      3. Use o Score para validar a eficiência.

      RESPOSTA:
      Escreva um veredito analítico de 3 linhas.
      Explique a relação Risco vs. Retorno usando os dados acima e a eficiência da alocação.
      Termine com uma conclusão técnica sobre a qualidade da operação.`;



      // 3. CHAMADA AO MOTOR CENTRAL (Arquivo 010)
      // callGeminiAI(prompt, systemInstruction, isJsonResponse)
      const veredito = callGeminiAI(promptIA, personaIA, false); 

      if (veredito) {
        // Localiza a linha correta na planilha para salvar o veredito
        const rowIdx = dados.findIndex(r => r[0] === rank) + 2; 
        sheet.getRange(rowIdx, 20)
             .setValue(veredito)
             .setBackground("#fdf7e3")
             .setFontStyle("italic")
             .setWrap(true);
      }
    });

    log(SERVICO_NOME, "SUCESSO", `Vereditos gerados para ${perolas.length} opções.`, "");

  } catch (e) {
    log(SERVICO_NOME, "ERRO", "Falha no Nexo Voice v1.5: " + e.message, e.stack);
  }
}