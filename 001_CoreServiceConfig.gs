/**
 * MÓDULO: Servico_Config - v1.2 (Otimizado para Performance)
 * OBJETIVO: Carregamento ultra-rápido de configurações com camada de Cache.
 */

// Variável global para evitar múltiplas leituras na mesma execução (Memoization)
var _configCache = null;

function obterConfigsGlobais() {
  const servicoNome = "ConfigService_v1.2";
  
  // 1. Tenta recuperar da memória global da instância atual
  if (_configCache !== null) {
    return _configCache;
  }

  // 2. Tenta recuperar do CacheService (útil para Web Apps e execuções sucessivas)
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get("global_configs");
  if (cachedData) {
    _configCache = JSON.parse(cachedData);
    return _configCache;
  }

  // 3. Se não houver cache, lê da planilha (fallback lento, mas necessário)
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Config_Global");
    
    if (!sheet) {
      gravarLog(servicoNome, "ERRO", "Aba Config_Global não encontrada", "");
      return {};
    }

    const data = sheet.getDataRange().getValues();
    let configs = {};
    
    for (let i = 1; i < data.length; i++) {
      let chave = data[i][0];
      let valor = data[i][1];
      
      if (chave && !chave.toString().startsWith("//")) { // Ignora comentários
        // Normalização de números (Padrão BR)
        if (typeof valor === 'string' && valor.includes(',')) {
          valor = valor.replace(',', '.');
        }
        
        let valorNumerico = parseFloat(valor);
        configs[chave.trim()] = isNaN(valorNumerico) ? valor : valorNumerico;
      }
    }

    // 4. Salva no CacheService por 25 minutos (máximo permitido é 6 horas)
    cache.put("global_configs", JSON.stringify(configs), 1500);
    
    // 5. Salva na variável global para esta execução
    _configCache = configs;

    gravarLog(servicoNome, "SUCESSO", "Configs carregadas via Planilha", `Total: ${Object.keys(configs).length}`);
    return configs;

  } catch (e) {
    gravarLog(servicoNome, "ERRO_CRITICO", "Falha ao carregar configs", e.toString());
    return {};
  }
}

/**
 * Função para limpar o cache manualmente (usar após alterar a aba Config_Global)
 */
function limparCacheConfig() {
  CacheService.getScriptCache().remove("global_configs");
  _configCache = null;
  console.log("Cache de configurações limpo com sucesso.");
}


// ═══════════════════════════════════════════════════════════════
// CONFIGURAÇÕES GLOBAIS DO ORQUESTRADOR
// ═══════════════════════════════════════════════════════════════

const CONFIG_ORQUESTRADOR = {
  versao: "1.0.0",
  ambiente: "PRODUCAO",
  timeout_padrao: 300000, // 5 minutos em ms
  retry_config: {
    max_tentativas: 3,
    delay_base: 1000, // 1 segundo
    delay_multiplicador: 2
  }
};


// ═══════════════════════════════════════════════════════════════
// CONSTANTES GLOBAIS DO PROJETO
// ═══════════════════════════════════════════════════════════════

// --- ABAS ---
const ABA_GATILHO = "Necton_Import";
const ABA_LOGS = "Logs";
const ABA_DADOS_DETALHES = "Dados_Detalhes";
const ABA_DADOS_ATIVOS = "Dados_Ativos";
const ABA_GREEKS = "Dados_Greeks";
const ABA_CONFIG = "Config_Global";


// --- COLUNAS DE INPUT/OUTPUT ---
const COLUNA_GATILHO_INPUT = 1;     // Coluna A (Ativo)
const COLUNA_ID_FORMULA = 16;       // Coluna P (ID_Trade_Unico)
const COLUNA_ID_ESTRUTURA = 18;     // Coluna R (ID_Estrutura)
const COLUNA_OUTPUT_INICIO = 19;    // Coluna S (Início da escrita do Script)
const COLUNA_OUTPUT_FIM = 23;       // Coluna W (Fim da escrita do Script)
const COLUNA_VENCIMENTO_NUM = 20;   // Coluna T (Vencimento - número)

// --- COLUNAS DE FILTRO (nomes dos headers) ---
const COLUNA_STATUS_ROBO = 'Status_Robo';  // Coluna Z
const COLUNA_VENCIMENTO = 'Vencimento';    // Coluna T (nome do header)

// --- OPLAB API ---
const OPLAB_API_URL = "https://api.oplab.com.br/v3";
const OPLAB_TOKEN_NAME = "OPLAB_ACCESS_TOKEN";

// --- PALETA DE CORES ---
const PALETA_CORES = [
  '#FFE5E5', '#E5F5FF', '#FFF9E5', '#E5FFE5', '#FFE5F5',
  '#F5E5FF', '#E5FFFF', '#FFFFE5', '#FFE5CC', '#E5CCFF',
  '#CCFFE5', '#FFDAE5', '#E5E5FF', '#FFCCE5', '#E5FFCC',
  '#FFD5E5', '#D5E5FF', '#FFFFE5', '#E5FFD5', '#FFE5D5'
];




function testarPerformanceConfig() {
  console.log("--- INICIANDO TESTE DE PERFORMANCE 001 ---");
  
  // Limpa o cache para garantir que o primeiro teste seja "frio"
  limparCacheConfig();
  
  let t0 = new Date().getTime();
  let c1 = obterConfigsGlobais();
  let t1 = new Date().getTime();
  console.log(`Teste 1 (Cache Frio - Planilha): ${t1 - t0}ms`);
  
  let t2 = new Date().getTime();
  let c2 = obterConfigsGlobais();
  let t3 = new Date().getTime();
  console.log(`Teste 2 (Cache Quente - Memória): ${t3 - t2}ms`);
  
  // Validação de Integridade
  if (Object.keys(c1).length === Object.keys(c2).length && c1["Taxa_Selic_Anual"]) {
    console.log("✅ SUCESSO: Dados íntegros e Selic encontrada: " + c1["Taxa_Selic_Anual"]);
  } else {
    console.warn("⚠️ AVISO: Verifique se as chaves foram carregadas corretamente ou se o nome da aba está correto.");
  }
  
  const tempoFrio = t1 - t0;
  const tempoQuente = t3 - t2;
  console.log(`Ganho de performance: ${Math.round(tempoFrio / (tempoQuente + 0.1))}x mais rápido`);
  console.log("--- FIM DO TESTE ---");
}




/**
 * TESTE UNITÁRIO: validarMigracaoInfra
 * Objetivo: Garantir que as constantes no 001 e os logs no 002 estão acessíveis.
 */
function testarMigracaoInfraUnitario() {
  console.log("--- INICIANDO TESTE DE MIGRAÇÃO ---");
  
  try {
    // 1. Teste de Acesso ao 001 (Config)
    const corTeste = PALETA_CORES[0];
    if (corTeste) {
      console.log("✅ Sucesso: Constantes do arquivo 001 acessíveis. Cor lida: " + corTeste);
    }

    // 2. Teste de Acesso ao 002 (Logger)
    // Vamos disparar um log de teste
    log("TESTE_MIGRACAO", "INFO", "Validando acesso ao logger após fatiamento", "Contexto de teste");
    console.log("✅ Sucesso: Função de log no arquivo 002 executada.");

    // 3. Teste de Configuração
    console.log("📊 Aba de Logs configurada como: " + ABA_LOGS);
    
    console.log("--- TESTE CONCLUÍDO COM SUCESSO ---");
  } catch (e) {
    console.error("❌ FALHA NO TESTE: Alguma referência foi perdida na migração. Erro: " + e.toString());
  }
}



function testarPasso1() {
  console.log("--- TESTANDO MIGRAÇÃO PASSO 1 ---");
  console.log("Aba de Logs: " + ABA_LOGS);
  console.log("URL API: " + OPLAB_API_URL);
  console.log("Coluna Gatilho: " + COLUNA_GATILHO_INPUT);
  console.log("Cor 1: " + PALETA_CORES[0]);
  console.log("--- SE APARECEU TUDO ACIMA, O PASSO 1 FOI SUCESSO ---");
}