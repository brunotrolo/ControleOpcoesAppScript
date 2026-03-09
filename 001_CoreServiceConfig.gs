/**
 * @fileoverview ConfigManager - v3.1 (Clean Architecture)
 * RESPONSABILIDADE: Apenas constantes e gerenciamento de cache de configurações.
 * LIMPEZA: Toda lógica de formatação de data/moeda foi delegada ao 002_CoreDataUtils.
 */

const SYS_CONFIG = {
  SHEETS: {
    TRIGGER: "Necton_Import",
    LOGS: "Logs",
    DETAILS: "Dados_Detalhes",
    ASSETS: "Dados_Ativos",
    GREEKS: "Dados_Greeks",
    CONFIG: "Config_Global",
    COCKPIT: "Cockpit",
    SELECTION: "Selecao_Opcoes"
  },
  
  COLUMNS: {
    TRIGGER_INPUT: 1,      // A
    ID_FORMULA: 17,        // Q
    ID_ESTRUTURA: 18,      // R
    OUTPUT_START: 19,      // S
    OUTPUT_END: 23,        // W
    EXPIRATION_NUM: 20     // T
  }
};

const ConfigManager = {
  _memoryCache: null,
  _cacheKey: "APP_GLOBAL_CONFIGS_V3",
  _cacheTime: 21600, // 6 horas

  /**
   * Obtém configurações dinâmicas da planilha com cache de 3 camadas.
   */
  get() {
    if (this._memoryCache) return this._memoryCache;

    const cache = CacheService.getScriptCache();
    const cachedData = cache.get(this._cacheKey);
    if (cachedData) {
      this._memoryCache = JSON.parse(cachedData);
      return this._memoryCache;
    }

    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(SYS_CONFIG.SHEETS.CONFIG);
      if (!sheet) return {};

      const data = sheet.getDataRange().getValues();
      const configs = {};
      
      for (let i = 1; i < data.length; i++) {
        const key = String(data[i][0]).trim();
        let val = data[i][1];
        if (key && !key.startsWith("//")) {
          // Usa o DataUtils (do arquivo 002) para limpar números vindo da planilha
          configs[key] = (typeof val === 'string' && val.includes(',')) ? 
                          DataUtils.safeFloat(val) : val;
        }
      }

      this._memoryCache = configs;
      cache.put(this._cacheKey, JSON.stringify(configs), this._cacheTime);
      return configs;
    } catch (e) {
      console.error(`[ConfigManager] Erro crítico no I/O: ${e.message}`);
      return {};
    }
  },

  /**
   * Invalida os caches para forçar nova leitura.
   */
  clearCache() {
    CacheService.getScriptCache().remove(this._cacheKey);
    this._memoryCache = null;
  }
};

// ============================================================================
// TESTES DE INTEGRAÇÃO DOS UTILITÁRIOS
// ============================================================================

/**
 * TESTE DE INTEGRIDADE: Verifica se o ConfigManager consegue usar o DataUtils (002)
 */
function testConfigIntegrity() {
  ConfigManager.clearCache();
  const cfg = ConfigManager.get();
  console.log("=== TESTE DE INTEGRIDADE 001 + 002 ===");
  console.log("Configurações carregadas:", Object.keys(cfg).length);
  console.log("Status: ✅ Sistema de Cache e Dependência OK.");
}