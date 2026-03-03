/**
 * MÓDULO: Servico_GeminiService - v2.0 (Estratégico & Parametrizado)
 * OBJETIVO: Gerar consultoria de nível institucional baseada em dogmas operacionais.
 */

function gerarConsultoriaEstrategica() {
  const servicoNome = "GeminiService_v2.0";
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  
  if (!apiKey) {
    gravarLog(servicoNome, "ERRO_CONFIG", "API Key ausente", "Configure 'GEMINI_API_KEY' nas Propriedades");
    return;
  }

  // 1. Obtém diagnóstico e configurações
  const diagnostico = processarAnaliseDeRisco();
  if (!diagnostico) return;

  const configs = diagnostico.configsUsadas;
  gravarLog(servicoNome, "INICIO", "Enviando dados para Gemini 2.5", `Alertas: ${diagnostico.alertas.length}`);

// 2. PROMPT DE ALTA PERFORMANCE v3.1 (Ajuste Racional Alpha)
  const prompt = `
    Atue como Gestor Sénior Nexo Opções. Sem introduções, sem "Prezado Cliente". Vá direto aos dados.
    
    FILOSOFIA: Foco em Theta/Vega, evitar Gamma (< ${configs['Regra_DTE_Saida_Alvo']} dias), proteger capital.

    REGRAS TÉCNICAS:
    - Lucro Alvo: ${configs['Regra_Lucro_Alvo_Min']*100}% a ${configs['Regra_Lucro_Alvo_Max']*100}%
    - Saída por Tempo: < ${configs['Regra_DTE_Saida_Alvo']} DTE.

    ESTRUTURA DO RELATÓRIO (SIGA RIGOROSAMENTE):
    IMPORTANTE: Mantenha a ordem cronológica dos dados fornecidos (vencimentos mais próximos primeiro) em todas as seções.

    1. PRIORIDADE CRÍTICA (Ações Imediatas):
    Para cada ativo que atingiu lucro máximo ou DTE crítico, use este formato exato:
    [Código do Ativo] | MOTIVO: [Racional técnico denso] | AÇÃO: [Comando Direto]

    // Exemplo que a IA deve seguir:
    BRAVM170W5 | MOTIVO: Risco de Gamma acentuado... | AÇÃO: FECHAR

    2. PRIORIDADE DE ALERTA (Defesa e Monitoramento):
    Para ativos ITM ou próximos do DTE limite, use este formato exato:
    TICKER: [Código] | MOTIVO: [Análise do risco direcional, sensibilidade do prêmio ao movimento do spot e por que a rolagem/defesa é a melhor estratégia.] | AÇÃO: [Sugestão de Rolagem/Defesa]

    3. MONITORIZAÇÃO REGULAR:
    Para todos os ativos remanescentes (OTM com DTE seguro), use este formato (o mesmo dos itens 1 e 2):
    [Código do Ativo] | MOTIVO: [Racional ultra-curto: cite apenas que o Delta está sob controle e o DTE é confortável] | AÇÃO: [Comando Curto, ex: MANTER]

    4. ANÁLISE DE NOTIONAL:
    Apresente no formato: EXPOSIÇÃO: R$ ${diagnostico.notionalTotal.toLocaleString('pt-BR', {minimumFractionDigits: 2})} | VEREDITO: [Análise de risco macro sobre a alavancagem atual e a saúde do patrimônio]

    5. STATUS DO SISTEMA:
    Uma frase curta confirmando que todos os dados foram processados sob as diretrizes Nexo 2026.

    DADOS:
    ALERTAS: ${JSON.stringify(diagnostico.alertas)}
    PORTFÓLIO: ${JSON.stringify(diagnostico.dadosParaIA)}
  `;


  try {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;
    
    const payload = { contents: [{ parts: [{ text: prompt }] }] };
    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true 
    };

    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() !== 200) {
      gravarLog(servicoNome, "ERRO_API", `Status: ${response.getResponseCode()}`, response.getContentText());
      return;
    }

    const json = JSON.parse(response.getContentText());
    const textoConsultoria = json.candidates[0].content.parts[0].text;

    // LINHA PARA ADICIONAR AGORA:
    console.log("TEXTO_BRUTO_GEMINI: " + textoConsultoria);
    
    gravarLog(servicoNome, "SUCESSO", "Consultoria Gerada", "Pronto para envio");
    return textoConsultoria;

  } catch (e) {
    gravarLog(servicoNome, "ERRO_FATAL", "Falha no Gemini 2.5", e.toString());
  }
}

/**
 * NOVO: Servico_GeminiScanner - v1.0
 * OBJETIVO: Gerar o "Veredito Nexo" para as oportunidades encontradas pelo Scanner.
 */
function callNexoScannerAI(promptEstrategico) {
  const servicoNome = "GeminiScanner_v1.0";
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  
  if (!apiKey) return "API Key ausente nas configurações.";

  try {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;
    
    const payload = { 
      contents: [{ 
        parts: [{ 
          text: promptEstrategico 
        }] 
      }] 
    };

    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true 
    };

    const response = UrlFetchApp.fetch(url, options);
    
    if (response.getResponseCode() === 200) {
      const json = JSON.parse(response.getContentText());
      return json.candidates[0].content.parts[0].text;
    } else {
      return "Análise técnica indisponível.";
    }

  } catch (e) {
    return "Erro na conexão com Nexo Voice.";
  }
}