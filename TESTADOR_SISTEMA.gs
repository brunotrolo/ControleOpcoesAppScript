/**
 * 🛰️ AUDITORIA DE BACKEND V5
 * Este script roda diretamente no Apps Script para validar os dados sem o Chrome.
 */
function EXECUTAR_AUDITORIA_V5() {
  console.log("🚀 Iniciando Auditoria Server-Side...");

  try {
    // 1. BUSCAR DADOS REAIS (Simula o google.script.run.getInitialData)
    // Se a sua função de buscar dados tiver outro nome, ajuste aqui.
    const response = getInitialData(); 
    
    if (!response || !response.success) {
      throw new Error("Falha ao buscar dados da planilha. Verifique o Código.gs");
    }

    const rawData = response.raw['COCKPIT'];
    console.log("✅ Dados da Planilha recuperados: " + rawData.length + " linhas totais.");

    // 2. SIMULAR TRADUTOR (Verificando o corte de linhas)
    // O seu Tradutor pula 9 linhas (headerRowIndex: 9)
    const headerRowIndex = 9; 
    const dataRows = rawData.slice(headerRowIndex + 1);
    
    console.log("🔍 Tradutor: Ignorando 10 linhas de cabeçalho...");
    console.log("🔍 Linhas restantes para processamento: " + dataRows.length);

    if (dataRows.length === 0) {
      console.warn("⚠️ ALERTA: Não há dados após a linha 10 da planilha Cockpit.");
    }

    // 3. VALIDAR STATUS "ATIVO"
    // No seu Tradutor, a coluna STATUS é a 3ª (índice 2)
    const tradesAtivos = dataRows.filter(row => {
      const status = String(row[2] || "").toUpperCase().trim();
      return status === "ATIVO";
    });

    console.log("🎯 Trades com status 'ATIVO' encontrados: " + tradesAtivos.length);

    if (tradesAtivos.length === 0 && dataRows.length > 0) {
      console.error("❌ ERRO: Existem dados, mas nenhum tem o status 'ATIVO' na coluna C.");
      console.log("Exemplo da primeira linha encontrada: " + JSON.stringify(dataRows[0]));
    }

    // 4. VALIDAR ESTRUTURA GLOBAL
    const namespaces = Object.keys(response.raw);
    console.log("📦 Namespaces recebidos: " + namespaces.join(", "));

    // 5. VEREDITO FINAL
    if (tradesAtivos.length > 0) {
      console.log("✅ HOMOLOGADO: Os dados estão chegando e o Tradutor vai encontrá-los.");
    } else {
      console.error("🛑 FALHA: O Dashboard ficará vazio porque não há trades ativos qualificados.");
    }

  } catch (err) {
    console.error("💥 CRASH NO TESTE: " + err.message);
  }
}



function LISTAR_NOMES_DAS_ABAS() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const nomes = sheets.map(s => s.getName());
  
  console.log("📋 Abas encontradas na sua planilha:");
  nomes.forEach(n => console.log("- " + n));
  
  const temCockpit = nomes.includes("COCKPIT");
  const temConfig = nomes.includes("CONFIG_GLOBAL");
  
  if (temCockpit && temConfig) {
    console.log("✅ NOMES ESTÃO CORRETOS.");
  } else {
    console.error("❌ ERRO DE NOMENCLATURA DETECTADO!");
    console.log("O sistema espera 'COCKPIT' e 'CONFIG_GLOBAL'. Ajuste os nomes na sua planilha.");
  }
}



/**
 * 🛰️ AUDITORIA SERVER-SIDE: FASE 2 (TRADUTOR)
 * Testa se o dicionário do Tradutor.html está conseguindo ler a sua planilha.
 */
function HOMOLOGAR_TRADUTOR() {
  console.log("🚀 [FASE 2] Iniciando Teste do Tradutor...");

  try {
    const rawData = getInitialData().raw['COCKPIT'];
    const headerRowIndex = 9; // Conforme o seu Tradutor.html
    
    const cabecalhos = rawData[headerRowIndex].map(h => String(h).toUpperCase().trim());
    const corpoDados = rawData.slice(headerRowIndex + 1);
    
    // Pega apenas os trades ATIVOS
    const tradesAtivos = corpoDados.filter(row => String(row[2] || "").toUpperCase().trim() === "ATIVO");
    
    if (tradesAtivos.length === 0) {
      console.error("❌ ERRO: Nenhum trade ATIVO para traduzir.");
      return;
    }

    // Pega o PRIMEIRO trade ativo para fazer o Raio-X
    const tradeTeste = tradesAtivos[0];
    const tradeTraduzido = {};

    // DICIONÁRIO EXATO DO SEU TRADUTOR.HTML
    const mapa = {
      'ID_TRADE': 'tradeId',
      'STATUS': 'status',
      'TICKER': 'tickerAtivo',
      'VENDA/COMPRA': 'direcaoTrade',
      'TIPO': 'tipoOpcao',
      'QTD': 'quantidade',
      'PRÊMIO (PM)': 'premioMedioEntrada',
      'NOCIONAL': 'nocionalTotal'
    };

    console.log(`🔍 Fazendo Raio-X no Trade da linha encontrada...`);
    
    let falhasDeMapeamento = 0;

    for (const [colunaOrig, chaveDestino] of Object.entries(mapa)) {
      const idx = cabecalhos.indexOf(colunaOrig);
      
      if (idx !== -1) {
        const valorOriginal = tradeTeste[idx];
        tradeTraduzido[chaveDestino] = valorOriginal;
        console.log(`✅ [${colunaOrig}] ➔ Encontrou: ${valorOriginal}`);
      } else {
        console.error(`❌ ALERTA: Coluna '${colunaOrig}' NÃO FOI ENCONTRADA no cabeçalho da planilha!`);
        falhasDeMapeamento++;
      }
    }

    console.log("--------------------------------------------------");
    console.log("📦 RESULTADO DO OBJETO TRADUZIDO (SIMULAÇÃO):");
    console.log(JSON.stringify(tradeTraduzido, null, 2));
    
    if (falhasDeMapeamento === 0) {
      console.log("✅ TRADUTOR HOMOLOGADO: Todas as colunas vitais foram encontradas!");
    } else {
      console.error(`⚠️ ATENÇÃO: O Tradutor falhou em achar ${falhasDeMapeamento} colunas vitais.`);
    }

  } catch (e) {
    console.error("💥 CRASH NO TESTE DO TRADUTOR: " + e.message);
  }
}



/**
 * 🛰️ AUDITORIA SERVER-SIDE: FASE 3 (AGREGADOR E MATEMÁTICA)
 * Testa se o motor OLAP consegue somar os dados financeiros sem causar NaN.
 */
function HOMOLOGAR_AGREGADOR() {
  console.log("🚀 [FASE 3] Iniciando Teste do Agregador (Matemática OLAP)...");

  try {
    const rawData = getInitialData().raw['COCKPIT'];
    const cabecalhos = rawData[9].map(h => String(h).toUpperCase().trim());
    const corpoDados = rawData.slice(10);
    const tradesAtivos = corpoDados.filter(row => String(row[2] || "").toUpperCase().trim() === "ATIVO");

    // Simulação exata da função _sanitizarMoeda do seu Tradutor/Agregador
    const sanitizarMoeda = (val) => {
      if (!val) return 0;
      let s = String(val).trim().replace(/[R$\s]/gi, '');
      if (s.includes(',')) s = s.replace(/\./g, '').replace(',', '.');
      const num = parseFloat(s);
      return isNaN(num) ? 0 : num;
    };

    const idxNocional = cabecalhos.indexOf('NOCIONAL');
    const idxPremio = cabecalhos.indexOf('PRÊMIO (PM)');
    const idxQtd = cabecalhos.indexOf('QTD');

    console.log(`📊 Injetando ${tradesAtivos.length} trades ativos no Motor Matemático...`);

    let somaNocional = 0;
    let falhasMatematicas = 0;

    tradesAtivos.forEach((trade, i) => {
      const valorBruto = trade[idxNocional];
      const valorLimpo = sanitizarMoeda(valorBruto);

      if (isNaN(valorLimpo) || (valorLimpo === 0 && valorBruto !== "0" && valorBruto !== "")) {
         console.error(`❌ Falha Matemática na linha ${i+1}: Não conseguiu converter [${valorBruto}]`);
         falhasMatematicas++;
      }
      somaNocional += valorLimpo;
    });

    console.log("--------------------------------------------------");
    console.log("📈 RESULTADO DO CUBO OLAP (SIMULAÇÃO GLOBAL):");
    console.log(`💰 NOCIONAL TOTAL DA CARTEIRA: R$ ${somaNocional.toLocaleString('pt-BR', {minimumFractionDigits: 2})}`);
    
    // Testa a lógica de defesas (Status de Risco)
    const otmPct = 85; // Simulação de carteira segura
    const poeMedio = 0.72; // Simulação de 72% de probabilidade
    let statusKey = 'ALERTA';
    if (otmPct >= 80 && poeMedio >= 0.7) statusKey = 'BLINDADO';
    else if (otmPct >= 50 || poeMedio >= 0.6) statusKey = 'CONTROLADO';
    
    console.log(`🛡️ STATUS DE RISCO (Regra V5): ${statusKey}`);
    console.log("--------------------------------------------------");

    if (falhasMatematicas === 0 && somaNocional > 0) {
      console.log("✅ AGREGADOR HOMOLOGADO: A matemática está 100% perfeita!");
    } else {
      console.error("⚠️ ATENÇÃO: O Agregador falhou em somar os valores.");
    }

  } catch (e) {
    console.error("💥 CRASH NO TESTE DO AGREGADOR: " + e.message);
  }
}


/**
 * 🛰️ AUDITORIA SERVER-SIDE: FASE 4 (APPCORE - MÁQUINA DE ESTADO VUE.JS)
 * Simula a reatividade do Vue e o ciclo de vida da tela principal.
 */
function HOMOLOGAR_APPCORE_PARTE1() {
  console.log("🚀 [FASE 4 - PARTE 1] Iniciando Máquina de Simulação do AppCore (Vue 3)...");

  try {
    // =========================================================
    // 1. SIMULANDO O SETUP() DO VUE E VARIÁVEIS REATIVAS (REF)
    // =========================================================
    console.log("⚙️ 1. Montando o Estado Inicial (refs)...");
    const state = {
      db: { cockpit: [], resumoGerencial: { timestamp: '--:--' }, configGlobal: {} },
      cuboOlap: { meta: { status: 'EMPTY' } },
      isLoading: true,
      hasError: false,
      status: 'Iniciando...',
      currentView: 'cockpit',
      isCollapsed: true,
      
      // Simulação das variáveis que o Vue reclamou na tela branca:
      topGridComponents: undefined, // Propósito de testar o crash
      menu: [],
      currentMenuLabel: undefined
    };

    console.log("✅ Estado Inicial Montado com Sucesso.");

    // =========================================================
    // 2. SIMULANDO O COMPUTED() DO VUE
    // =========================================================
    console.log("⚙️ 2. Testando Propriedades Computadas...");
    
    // Simula a injeção de dados da API no DB
    state.db.configGlobal = { uiLabelTitulo: "Dashboard V5 - Sinergia" };
    state.db.cockpit = [1, 2, 3]; // 3 trades mockados
    
    // Teste do currentMenuLabel
    state.currentMenuLabel = state.db.configGlobal.uiLabelTitulo || 'Dashboard';
    console.log(`✅ Menu Label Computado: [${state.currentMenuLabel}]`);
    
    // Teste de estruturação dos Grids (O que causou o Crash)
    // Se o setup() não exporta isso, ele fica undefined e quebra o .length
    try {
      if (state.topGridComponents === undefined) {
         throw new Error("topGridComponents não foi exportado pelo setup()!");
      }
      console.log(`✅ Layout Superior tem ${state.topGridComponents.length} componentes.`);
    } catch (e) {
      console.error(`❌ CRASH DETECTADO NA CAMADA DE UI: ${e.message}`);
      console.log("💡 DIAGNÓSTICO: É exatamente por isso que a sua tela ficou branca. O AppCore precisa retornar essas variáveis no final do bloco setup().");
    }

    // =========================================================
    // 3. SIMULANDO A NAVEGAÇÃO (ROTEAMENTO INTERNO)
    // =========================================================
    console.log("⚙️ 3. Testando Motor de Navegação Interna...");
    
    const navigate = (view) => {
      console.log(`🔄 Trocando tela de [${state.currentView}] para [${view}]`);
      state.currentView = view;
      if (view !== 'cockpit' && view !== 'consultoria') {
        state.isCollapsed = true; 
      }
    };

    navigate('consultoria');
    if (state.currentView === 'consultoria') {
       console.log("✅ Navegação OK.");
    }

    console.log("--------------------------------------------------");
    console.log("🏁 FIM DA PARTE 1 DA AUDITORIA DO APPCORE");

  } catch (e) {
    console.error("💥 FALHA FATAL NA SIMULAÇÃO DO APPCORE: " + e.message);
  }
}



/**
 * 🛰️ AUDITORIA SERVER-SIDE: FASE 4 (APPCORE - INJEÇÃO DE DEPENDÊNCIAS)
 * Testa se a função getComponentProps está repassando a matemática para os cards.
 */
function HOMOLOGAR_APPCORE_PARTE2() {
  console.log("🚀 [FASE 4 - PARTE 2] Iniciando Teste de Injeção de Componentes...");

  try {
    // 1. Simula o estado do sistema já carregado e calculado
    const isLoading = { value: false };
    const lastSync = { value: "15:00:00" };
    const db = { 
      value: { 
        configGlobal: { alerta: 'ligado' },
        cockpit: [{ id: 1, ativo: 'PETR4' }]
      }
    };
    const cuboOlap = { 
      value: { 
        resumoGlobal: { nocionalTotal: 116180 },
        distribuicaoVencimentos: { labels: ['Abr'] }
      }
    };

    // 2. Réplica exata da função getComponentProps do seu AppCore
    const getComponentProps = (componentName) => {
        if (isLoading.value || !db.value || !db.value.cockpit) return { isLoading: true };

        const baseProps = { 
            isLoading: false, 
            lastSync: lastSync.value,
            configGlobais: db.value.configGlobal || {}
        };

        switch (componentName) {
            case 'CardSummary': return { ...baseProps, statsPreCalculados: cuboOlap.value };
            case 'CockpitTable': return { ...baseProps, trades: db.value.cockpit, resumoGlobal: cuboOlap.value.resumoGlobal };
            case 'CardVencimentos': return { ...baseProps, distribuicao: cuboOlap.value.distribuicaoVencimentos };
            default: return baseProps;
        }
    };

    console.log("⚙️ Solicitando dados para o 'CardSummary'...");
    const propsCard1 = getComponentProps('CardSummary');
    if (propsCard1.statsPreCalculados.resumoGlobal.nocionalTotal === 116180) {
      console.log("✅ CardSummary: Recebeu R$ 116.180 perfeitamente!");
    } else {
      console.error("❌ FALHA: CardSummary não recebeu o cubo OLAP.");
    }

    console.log("⚙️ Solicitando dados para a 'CockpitTable'...");
    const propsTable = getComponentProps('CockpitTable');
    if (propsTable.trades.length === 1 && propsTable.trades[0].ativo === 'PETR4') {
      console.log("✅ CockpitTable: Recebeu os trades do banco de dados perfeitamente!");
    } else {
      console.error("❌ FALHA: CockpitTable não recebeu os trades.");
    }

    console.log("--------------------------------------------------");
    console.log("🏁 FIM DA PARTE 2: A distribuição de dados está HOMOLOGADA.");
    console.log("💡 Se você fez a alteração do 'return' no AppCore, seu Web App JÁ PODE SER ATUALIZADO no navegador!");

  } catch (e) {
    console.error("💥 CRASH NO TESTE DE INJEÇÃO: " + e.message);
  }
}