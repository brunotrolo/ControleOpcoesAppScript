/**
 * MÓDULO: OpLabClient - v1.4
 * OBJETIVO: Centralizar requisições para a API OpLab (Options e Stocks).
 */
function oplaBCall_(endpoint) {
  const servico = "OpLabClient_v1.4";
  
  // 1. Recuperação segura de constantes globais/propriedades
  let chaveToken = "OPLAB_ACCESS_TOKEN"; 
  try { if (typeof OPLAB_TOKEN_NAME !== 'undefined') chaveToken = OPLAB_TOKEN_NAME; } catch(e){}
  
  const propToken = PropertiesService.getScriptProperties().getProperty(chaveToken);
  
  if (!propToken) {
    console.error("❌ OpLabClient: Token não encontrado na chave: " + chaveToken);
    return null;
  }

  let urlBase = "https://api.oplab.com.br/v3";
  try { if (typeof OPLAB_API_URL !== 'undefined') urlBase = OPLAB_API_URL; } catch(e){}
  
  const urlFinal = urlBase + endpoint;

  // 2. Configuração da requisição
  const options = {
    "method": "get",
    "headers": { "Access-Token": String(propToken).trim() },
    "muteHttpExceptions": true
  };

  try {
    const response = UrlFetchApp.fetch(urlFinal, options);
    const code = response.getResponseCode();
    
    if (code !== 200) {
      // 204 é comum para ativos vencidos ou sem dados no momento
      console.warn(`⚠️ OpLab API [${code}] em: ${endpoint}`);
      return null;
    }
    
    return JSON.parse(response.getContentText());

  } catch (e) {
    console.error("FALHA CRÍTICA NO FETCH: " + e.toString());
    return null;
  }
}

/**
 * Interface para Detalhes de OPÇÕES
 */
function getOpLabOptionDetails(ticker) {
  if (!ticker) return null;
  return oplaBCall_("/market/options/details/" + ticker);
}

/**
 * Interface para Dados de ATIVOS (Stocks)
 */
function getOpLabStockData(ticker) {
  if (!ticker) return null;
  return oplaBCall_("/market/stocks/" + ticker);
}

/**
 * Interface para CALCULAR GREGAS (Black-Scholes)
 * Fiel ao Swagger OpLab v3
 */
function calculateOpLabBS(params) {
  // Ajustando para os nomes EXATOS do Swagger
  const queryParams = {
    "symbol": params.symbol,
    "irate": params.irate,     // Ex: 10.75 (em %)
    "type": params.type,       // "CALL" ou "PUT"
    "spotprice": params.spotprice,
    "strike": params.strike,
    "premium": params.premium || 0,
    "dtm": params.dtm,
    "vol": params.vol,         // Ex: 44.27 (em %)
    "amount": params.amount || 0
  };

  const queryString = Object.keys(queryParams)
    .map(k => encodeURIComponent(k) + '=' + encodeURIComponent(queryParams[k]))
    .join('&');

  return oplaBCall_("/market/options/bs?" + queryString);
}

function testarTokenOpLab() {
  const token = PropertiesService.getDocumentProperties().getProperty("OPLAB_ACCESS_TOKEN");
  Logger.log("TOKEN LIDO: " + token);
}


/**
 * Interface para Dados FUNDAMENTALISTAS via BRAPI (Oficial e sem bloqueios)
 * Retorna JSON puro da B3. Utiliza módulos avançados para VPA e ROE.
 */
function getBrapiStockData(ticker) {
  if (!ticker) return null;

  // Dica de segurança: Lembre-se de rotacionar ou ocultar este token no futuro!
  const BRAPI_TOKEN = "49jViAQnD5LmPuZmdCZgM8"; 
  
  try {
    // 🌟 NOVO: Adicionamos os módulos avançados que trazem o VPA e o ROE
    const url = `https://brapi.dev/api/quote/${ticker}?modules=defaultKeyStatistics,financialData,summaryDetail&token=${BRAPI_TOKEN}`;
    
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    
    if (response.getResponseCode() !== 200) {
      console.warn(`BRAPI falhou para ${ticker}. HTTP: ${response.getResponseCode()}`);
      return null;
    }

    const json = JSON.parse(response.getContentText());
    if (!json.results || json.results.length === 0) return null;
    
    const data = json.results[0];
    const stats = data.defaultKeyStatistics || {};
    const finData = data.financialData || {};
    const summary = data.summaryDetail || {};

    // Helper: A API entrega os dados soltos ou dentro de objetos { raw: 123 }
    const getVal = (obj, key) => {
      if (!obj || obj[key] === undefined) return 0;
      if (typeof obj[key] === 'object' && obj[key] !== null) return parseFloat(obj[key].raw) || 0;
      return parseFloat(obj[key]) || 0;
    };

    return {
      lpa: getVal(stats, 'trailingEps') || parseFloat(data.earningsPerShare) || 0,
      vpa: getVal(stats, 'bookValue'),
      pl: getVal(summary, 'trailingPE') || parseFloat(data.priceEarnings) || 0,
      pvp: getVal(stats, 'priceToBook'),
      dy: getVal(summary, 'dividendYield') || (parseFloat(data.dividendsYield) || 0) / 100, 
      roe: getVal(finData, 'returnOnEquity'), 
      divida: 0 
    };

  } catch (e) {
    console.error(`Erro na API BRAPI (${ticker}): ` + e.toString());
    return null;
  }
}