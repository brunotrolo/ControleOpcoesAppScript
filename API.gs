/**
 * API.gs - The "State of the Art" Universal Backend Repository
 * Padrão: Agnóstico. O Servidor não conhece regras de negócio, apenas executa comandos de dados.
 */

// ==========================================
// 1. READ (Leitura Universal de Todo o Banco)
// ==========================================

function getInitialData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    
    const data = {
      success: true,
      timestamp: new Date().toLocaleString('pt-BR'),
      raw: {}
    };

    // O Backend entrega a matriz bruta EXATA. O Frontend decide onde está o cabeçalho.
    sheets.forEach(sheet => {
      const name = sheet.getName();
      const lastRow = sheet.getLastRow();
      
      if (lastRow === 0) {
        data.raw[name] = []; // Previne erros em abas recém-criadas e vazias
      } else {
        // Retorna a aba inteira, da linha 1 até o final.
        data.raw[name] = sheet.getDataRange().getDisplayValues();
      }
    });

    return data;
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ==========================================
// 2. CREATE (Inserção em Lote)
// ==========================================
function apiAdicionarLinhas(nomeAba, dadosMatriz) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nomeAba);
    if (!sheet) throw new Error(`Aba [${nomeAba}] não existe no banco de dados.`);

    const maxRows = sheet.getMaxRows();
    const valuesColA = sheet.getRange(1, 1, maxRows, 1).getValues();
    let startRow = 1;

    // Busca a primeira linha estritamente vazia
    for (let i = 0; i < valuesColA.length; i++) {
      if (valuesColA[i][0] === "") { startRow = i + 1; break; }
    }
    if (startRow === 1 && valuesColA[0][0] !== "") startRow = maxRows + 1;

    sheet.getRange(startRow, 1, dadosMatriz.length, dadosMatriz[0].length).setValues(dadosMatriz);
    return { success: true, message: `${dadosMatriz.length} linhas adicionadas em [${nomeAba}].` };
  } catch (e) { return { success: false, error: e.message }; }
}

// ==========================================
// 3. UPDATE (Atualização de Chave-Valor)
// Substitui a antiga "saveGlobalConfig"
// ==========================================
function apiAtualizarChaveValor(nomeAba, dicionarioAtualizacoes, colChave = 1, colValor = 2) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nomeAba);
    if (!sheet) throw new Error(`Aba [${nomeAba}] não existe no banco de dados.`);

    const data = sheet.getDataRange().getValues();
    
    const newVals = data.map(row => {
      const key = row[colChave - 1]; // Índice da chave (ex: A)
      // Se a chave veio no payload, atualiza. Se não, devolve o valor que já estava lá.
      return [dicionarioAtualizacoes[key] !== undefined ? dicionarioAtualizacoes[key] : row[colValor - 1]];
    });

    sheet.getRange(1, colValor, newVals.length, 1).setValues(newVals);
    return { success: true, message: `Dados atualizados com sucesso em [${nomeAba}].` };
  } catch (e) { return { success: false, error: e.message }; }
}

// ==========================================
// 4. DELETE & TRUNCATE (Exclusão e Limpeza)
// ==========================================
function apiExcluirLinhaSegura(nomeAba, numeroLinha, valorEsperadoColunaA) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nomeAba);
    if (!sheet) throw new Error(`Aba [${nomeAba}] não existe no banco de dados.`);

    const valorPlanilha = (sheet.getRange(numeroLinha, 1).getDisplayValue() || "").toString().trim().toUpperCase();
    const valorSeguro = (valorEsperadoColunaA || "").toString().trim().toUpperCase();

    // Trava anti-dessincronização
    if (valorPlanilha !== valorSeguro) {
      return { success: false, error: `Falha de sincronia: Esperava encontrar [${valorSeguro}], mas encontrou [${valorPlanilha}].` };
    }

    sheet.deleteRow(numeroLinha);
    return { success: true, message: `Registro removido de [${nomeAba}].` };
  } catch (e) { return { success: false, error: e.message }; }
}


/**
 * EXCLUSÃO EM LOTE (BACKEND)
 * Recebe: nome da aba (String) e lista de linhas (Array de Números)
 */
function apiExcluirLinhasEmLote(nomeAba, listaLinhas) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(nomeAba);
    if (!sheet) throw new Error("Aba não encontrada: " + nomeAba);

    // FILTRO DE SEGURANÇA: Remove valores nulos, garante que são números inteiros
    var linhasOrdenadas = listaLinhas
      .filter(function(l) { return l !== null && l !== undefined; }) 
      .map(function(l) { return parseInt(l, 10); })
      .sort(function(a, b) { return b - a; });

    // 2. Executa as exclusões em loop
    linhasOrdenadas.forEach(function(linha) {
      // Evita tentar deletar linha 0 ou negativa
      if (linha >= 1) { 
        sheet.deleteRow(linha);
      }
    });

    SpreadsheetApp.flush();
    return { success: true, count: linhasOrdenadas.length };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// Substitui a antiga "apiLimparLogsSistema"
function apiLimparAba(nomeAba, manterLinhasTop = 1, mensagemAuditoria = null) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(nomeAba);
    if (!sheet) throw new Error(`Aba [${nomeAba}] não existe no banco de dados.`);

    const lastRow = sheet.getLastRow();
    if (lastRow > manterLinhasTop) {
      sheet.getRange(manterLinhasTop + 1, 1, lastRow - manterLinhasTop, sheet.getLastColumn()).clearContent();
    }

    // Se o frontend pediu para deixar um rastro de auditoria, ele o faz.
    if (mensagemAuditoria) {
      const ts = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "dd/MM/yyyy HH:mm:ss");
      sheet.getRange(manterLinhasTop + 1, 1, 1, 4).setValues([[ts, "SYSTEM", "AVISO", mensagemAuditoria]]);
    }

    return { success: true, message: `Aba [${nomeAba}] foi limpa.` };
  } catch (e) { return { success: false, error: e.message }; }
}

// ==========================================
// 5. EXTERNAL API BRIDGE (Integrações de Terceiros)
// ==========================================
function apiIntegracaoOpLab(ticker) {
  if (!ticker || ticker.trim() === '') return { success: false, error: "Ticker não fornecido." };
  try {
    // Wrapper seguro para a função nativa getOpLabOptionDetails()
    const data = getOpLabOptionDetails(ticker.toUpperCase().trim());
    if (!data) return { success: false, error: "Ativo não encontrado." };
    
    return {
      success: true,
      data: {
        symbol: data.symbol, 
        category: data.category, 
        strike: parseFloat(data.strike || 0),
        premioAtual: parseFloat(data.close > 0 ? data.close : (data.bid || 0)), 
        spotPrice: parseFloat(data.spot_price || 0),
        dte: parseInt(data.days_to_maturity || 0)
      }
    };
  } catch (e) { return { success: false, error: e.toString() }; }
}

// ==========================================
// 6. ATOMIC UPDATE (Escrita em Célula Única)
// ==========================================
function apiSetCellValue(nomeAba, linha, coluna, valor) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(nomeAba);
    if (!sheet) throw new Error("Aba não encontrada.");

    // Operação atômica: sem leitura prévia, apenas escrita direta por coordenada
    sheet.getRange(linha, coluna).setValue(valor);
    
    return { success: true, timestamp: new Date().toLocaleTimeString() };
  } catch (e) {
    return { success: false, error: e.message };
  }
}


/**
 * ═══════════════════════════════════════════════════════════════
 * SERVIÇO DE INTERFACE: apiSimularHorizontePreditivo
 * ═══════════════════════════════════════════════════════════════
 * RESPONSABILIDADE: Receber o novo DTE do Front-end, persistir na
 * Config_Global e engatilhar o motor de recálculo (Pipeline 600).
 * ═══════════════════════════════════════════════════════════════
 */
function apiSimularHorizontePreditivo(diasParam) {
  try {
    // 1. Validação de Segurança
    const dias = parseInt(diasParam, 10);
    if (isNaN(dias) || dias < 1 || dias > 45) {
      throw new Error("Horizonte inválido. O parâmetro deve ser um número entre 1 e 45.");
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaConfig = ss.getSheetByName("Config_Global");
    if (!abaConfig) throw new Error("Aba Config_Global não encontrada no banco de dados.");

    // 2. Persistência de Estado (On-Demand)
    // Busca a chave Regra_Dias_Horizonte_Preditivo e atualiza o valor
    const dados = abaConfig.getDataRange().getValues();
    let configuracaoAtualizada = false;
    
    for (let i = 0; i < dados.length; i++) {
      if (dados[i][0] === "Regra_Dias_Horizonte_Preditivo") {
        // i + 1 porque as linhas no Google Sheets começam em 1, e a coluna Valor é a 2 (B)
        abaConfig.getRange(i + 1, 2).setValue(dias);
        configuracaoAtualizada = true;
        break;
      }
    }

    // Se a chave não existir por algum motivo, ele a cria automaticamente seguindo o nosso Dicionário
    if (!configuracaoAtualizada) {
      abaConfig.appendRow([
        "Regra_Dias_Horizonte_Preditivo", 
        dias, 
        "[PREDICAO] Horizonte de simulação do Cone de Probabilidade | 1 a 45 dias | Módulos 602/605"
      ]);
    }

    // Limpa o cache para forçar a leitura fresca no próximo passo (caso você use o sistema de CacheService)
    if (typeof CacheService !== 'undefined') {
       CacheService.getScriptCache().remove("CONFIG_GLOBAL_CACHE");
    }

    // 3. Orquestração: Aciona o recálculo do Pipeline
    // Ele tenta chamar a sua função mestre que roda o 600_Pipeline_Integracao
    if (typeof executarFluxoSequencial === "function") {
      executarFluxoSequencial(); 
    } else if (typeof gerarAnalisePreditivaHeatmap === "function") {
      // Fallback: Se o mestre não for encontrado, roda direto o 605
      gerarAnalisePreditivaHeatmap(dias);
    } else {
      throw new Error("Pipeline de cálculo preditivo não encontrado no servidor.");
    }

    // 4. Retorno de Sucesso para o Web App
    return { 
      success: true, 
      mensagem: `Simulação para ${dias} dias concluída com sucesso.`,
      horizonte: dias 
    };

  } catch (error) {
    console.error("❌ ERRO em apiSimularHorizontePreditivo:", error);
    
    // Se você tiver o Logger global configurado
    if (typeof gravarLog === "function") {
      gravarLog("API_Simulador", "ERRO", "Falha ao recalcular horizonte preditivo", error.message);
    }

    return { 
      success: false, 
      error: error.message 
    };
  }
}

/**
 * 🛡️ BUSCA DE PLANILHA DINÂMICA
 * Encontra a aba independentemente de como foi digitada
 */
function getPlanilhaDinamica(planilhaAtiva, nomeProcurado) {
  const abas = planilhaAtiva.getSheets();
  const abaEncontrada = abas.find(aba => 
    aba.getName().toUpperCase() === String(nomeProcurado).toUpperCase()
  );
  return abaEncontrada || null;
}

function getAbaDinamica(payloadRaw, nomeProcurado) {
  const chaveReal = Object.keys(payloadRaw).find(k => 
    String(k).toUpperCase() === String(nomeProcurado).toUpperCase()
  );
  return chaveReal ? payloadRaw[chaveReal] : null;
}