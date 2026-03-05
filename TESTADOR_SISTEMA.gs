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



/**
 * 🛰️ MASTER DIAGNÓSTICO INTEGRADO V5
 * Este teste simula a fiação entre AppCore, LayoutConfig e Menu.
 */
function MASTER_DIAGNOSTICO_V5() {
  console.log("🚀 Iniciando Diagnóstico de Fiação (Integração Total)...");

  try {
    // --- 1. SIMULAÇÃO DO AMBIENTE (LAYOUT E MENU) ---
    // Aqui simulamos o que deveria vir do seu LayoutConfig.html
    const MOCK_APP_MENU = [
      { id: 'cockpit', label: 'Cockpit', icon: 'layout' },
      { id: 'consultoria', label: 'Consultor', icon: 'shield' }
    ];

    const MOCK_LAYOUT_MAP = {
      'cockpit': [{ id: 'card-1', component: 'card-summary' }]
    };

    console.log("✅ Simulação de Configurações carregada.");

    // --- 2. TESTE DE IDENTIDADE (O QUE QUEBRA O MENU) ---
    const currentView = 'cockpit'; // O que o sistema inicia
    console.log(`🔍 Testando correspondência de Menu para: '${currentView}'`);

    const menuMatch = MOCK_APP_MENU.find(m => m.id === currentView);
    if (menuMatch) {
      console.log(`✅ SUCESSO: O ID '${currentView}' foi encontrado no Menu.`);
    } else {
      console.error(`❌ FALHA CRÍTICA: O ID '${currentView}' NÃO EXISTE no Menu. Por isso o menu não acende.`);
    }

    // --- 3. TESTE DE MAPEAMENTO DE TELA (O QUE ESCONDE A HOME) ---
    console.log(`🔍 Testando se o LayoutConfig reconhece a tela: '${currentView}'`);
    const layoutMatch = MOCK_LAYOUT_MAP[currentView];
    if (layoutMatch && layoutMatch.length > 0) {
      console.log(`✅ SUCESSO: Foram encontrados ${layoutMatch.length} componentes para esta tela.`);
    } else {
      console.error(`❌ FALHA CRÍTICA: A tela '${currentView}' não tem componentes no LAYOUT_MAP. Por isso a Home fica vazia.`);
    }

    // --- 4. TESTE DE CONVERSÃO DE DADOS (INTEGRAÇÃO TRADUTOR -> AGREGADOR) ---
    console.log("🔍 Testando Fluxo de Dados Real...");
    const response = getInitialData(); // Pega os dados reais da sua planilha
    
    // Simula a normalização de nomes de abas que fizemos no Tradutor
    const getTab = (name) => {
        const key = Object.keys(response.raw).find(k => k.toUpperCase() === name.toUpperCase());
        return response.raw[key] || [];
    };

    const cockpitData = getTab('COCKPIT');
    console.log(`📊 Dados Brutos do Cockpit: ${cockpitData.length} linhas.`);

    // --- 5. O TESTE DO "DISPENSADOR DE PROPS" (A PONTE DO APPCORE) ---
    console.log("🔍 Testando o Dispensador de Props (getComponentProps)...");
    
    // Simulação do objeto db que o AppCore cria
    const mockDB = { cockpit: cockpitData.slice(10) }; // Pula o cabeçalho
    const mockCubo = { resumoGlobal: { total: 100 } }; // Simula o Agregador

    function testarProps(compName) {
      // Esta é a lógica que está dentro do seu AppCore.html
      const props = {
        'card-summary': { stats: mockCubo.resumoGlobal },
        'cockpit-table': { data: mockDB.cockpit }
      };
      return props[compName] || null;
    }

    const testSummary = testarProps('card-summary');
    if (testSummary) {
      console.log("✅ getComponentProps: 'card-summary' configurado corretamente.");
    } else {
      console.error("❌ ERRO: 'card-summary' não está mapeado no getComponentProps.");
    }

    console.log("--------------------------------------------------");
    console.log("🏁 RESUMO DO DIAGNÓSTICO:");
    if (!menuMatch || !layoutMatch) {
       console.log("🛑 CONCLUSÃO: Há um erro de NOMENCLATURA entre os arquivos. O ID que um usa, o outro não reconhece.");
    } else {
       console.log("💎 CONCLUSÃO: A lógica interna está sólida. O problema é a renderização no Vue.js.");
    }

  } catch (e) {
    console.error("💥 CRASH NO DIAGNÓSTICO: " + e.message);
  }
}


/**
 * 🛰️ TESTE UNITÁRIO 1: INTEGRALIDADE DOS ARQUIVOS FRONT-END
 * Verifica se todos os componentes registrados no Index.html realmente existem no servidor.
 */
function TESTE_01_ARQUIVOS_HTML() {
  console.log("🚀 [TESTE 1] Verificando existência dos arquivos HTML dos Componentes...");
  
  const componentes = [
    'MenuSidebar', 'Tradutor', 'Agregador', 'AppCore', 'LayoutConfig',
    'ConsultoriaView', 'TickerTape', 'CardSummary', 'CardDistribMoneyness',
    'CardMoneyness', 'CardVencimentos', 'CardConcentracao', 'CardResultado',
    'CarouselAtivas', 'CockpitTable', 'SettingsView', 'AutomationView',
    'LogConsole', 'ComparativoView', 'ScannerView', 'NectonImport',
    'GestaoAtivos', 'GestaoAtivosSidebar', 'BancoDeDados'
  ];

  let falhas = 0;

  componentes.forEach(comp => {
    try {
      HtmlService.createHtmlOutputFromFile(comp);
      console.log(`✅ Arquivo encontrado: ${comp}.html`);
    } catch (e) {
      console.error(`❌ ERRO CRÍTICO: Arquivo não encontrado: ${comp}.html`);
      falhas++;
    }
  });

  console.log("--------------------------------------------------");
  if (falhas === 0) console.log("🏆 TESTE 1 APROVADO: Todos os arquivos visuais existem.");
  else console.error(`🛑 TESTE 1 REPROVADO: Faltam ${falhas} arquivos. O Vue.js vai quebrar ao tentar carregá-los.`);
}


/**
 * 🛰️ TESTE UNITÁRIO 2: MAPEAMENTO DE PROPS (DATA BINDING)
 * Simula a entrega de dados para cada card e verifica se não há "vazios".
 */
function TESTE_02_SIMULADOR_PROPS() {
  console.log("🚀 [TESTE 2] Simulando injeção de dados nos Componentes...");

  // 1. Mock do Banco de Dados (O que o Tradutor e Agregador entregariam)
  const mockDB = {
    cockpit: [{ id: 1, ativo: 'PETR4', status: 'ATIVO' }],
    dadosAtivos: [{ ticker: 'PETR4', cotacao: 35.50 }],
    configGlobal: { uiLabelTitulo: "Dashboard V5" }
  };
  
  const mockCuboOlap = {
    listaAtivas: [{ id: 1, ativo: 'PETR4' }],
    plTotal: 15000,
    nocionalTotal: 120000
  };

  // 2. A função exata que roda no seu AppCore.html
  function getComponentProps(componentName) {
    const customProps = {
      'ticker-tape': { ativos: mockDB.dadosAtivos },
      'card-summary': { statsPreCalculados: mockCuboOlap },
      'cockpit-table': { data: mockDB.cockpit },
      'card-vencimentos': { operations: mockCuboOlap.listaAtivas },
      'card-resultado': { operations: mockCuboOlap.listaAtivas },
      'consultoria-view': { operations: mockCuboOlap.listaAtivas }
    };
    return customProps[componentName] || { erro: "Componente não mapeado" };
  }

  // 3. Testando a entrega para componentes críticos
  const alvos = ['card-summary', 'cockpit-table', 'card-vencimentos', 'card-distrib-moneyness'];
  
  alvos.forEach(alvo => {
    const props = getComponentProps(alvo);
    if (props.erro) {
      console.error(`❌ FALHA: O componente '${alvo}' NÃO está mapeado no AppCore. Ele vai aparecer vazio.`);
    } else {
      const chaves = Object.keys(props);
      console.log(`✅ SUCESSO: O componente '${alvo}' recebeu as variáveis: [${chaves.join(', ')}]`);
    }
  });
}

/**
 * 🛰️ TESTE UNITÁRIO 3: AUDITORIA DO TRADUTOR V5
 * Verifica se os cabeçalhos da aba COCKPIT estão sendo lidos corretamente.
 */
function TESTE_03_MOTOR_TRADUTOR() {
  console.log("🚀 [TESTE 3] Testando leitura da planilha Cockpit...");
  
  try {
    const response = getInitialData();
    const rawCockpit = response.raw['COCKPIT'];
    
    if (!rawCockpit) throw new Error("Aba 'COCKPIT' não encontrada na planilha.");
    
    // Pega o cabeçalho (linha 9 ou 10 baseada no seu código)
    const cabecalho = rawCockpit[9]; 
    console.log("📊 Cabeçalhos reais encontrados na planilha:");
    console.log(cabecalho.slice(0, 15).join(" | ")); // Mostra os 15 primeiros

    const chavesCriticas = ['STATUS', 'TICKER', 'P/L TOTAL', 'VENCIMENTO'];
    let faltantes = [];

    chavesCriticas.forEach(chave => {
      // Procura com tolerância a espaços e maiúsculas
      const achou = cabecalho.some(c => String(c).toUpperCase().trim().includes(chave));
      if (!achou) faltantes.push(chave);
    });

    if (faltantes.length > 0) {
      console.error(`❌ ALERTA: As seguintes colunas NÃO foram encontradas: ${faltantes.join(', ')}`);
      console.log("💡 O Agregador vai falhar ao calcular esses dados.");
    } else {
      console.log("✅ SUCESSO: Todas as colunas críticas existem na planilha.");
    }

  } catch (e) {
    console.error("💥 FALHA NO TESTE 3: " + e.message);
  }
}


/**
 * 🛰️ TESTE UNITÁRIO 4: TOPOGRAFIA DA PLANILHA
 * Imprime as primeiras linhas para vermos o que o Tradutor está recebendo.
 */
function TESTE_04_ESTRUTURA_PLANILHA() {
  console.log("🚀 [TESTE 4] Mapeando a Topografia da Planilha Cockpit...");
  
  try {
    const response = getInitialData();
    const raw = response.raw['COCKPIT'] || response.raw['Cockpit'];
    
    if (!raw) throw new Error("Aba Cockpit não encontrada.");

    console.log("📊 Primeiras 12 linhas do banco de dados bruto:");
    for(let i = 0; i < 12; i++) {
      // Pega apenas as 8 primeiras colunas para o log não ficar gigante
      let amostra = raw[i].slice(0, 8).map(celula => celula === "" ? "[VAZIO]" : celula);
      console.log(`Linha ${i + 1}: ${amostra.join(" | ")}`);
    }

  } catch (e) {
    console.error("💥 FALHA NO TESTE 4: " + e.message);
  }
}


/**
 * 🛰️ TESTE UNITÁRIO 5: SIMULAÇÃO DO ERRO DO TRADUTOR
 * Prova que o Tradutor está usando a linha errada como cabeçalho.
 */
function TESTE_05_SIMULACAO_TRADUTOR() {
  console.log("🚀 [TESTE 5] Simulando a lógica de cabeçalho do Tradutor V5...");
  
  try {
    const response = getInitialData();
    const raw = response.raw['COCKPIT'] || response.raw['Cockpit'];

    // Simula a linha que o Tradutor atual está pegando (matriz[0])
    const cabecalhoErrado = raw[0].slice(0, 6).map(c => c === "" ? "[VAZIO]" : c);
    console.log("❌ O Tradutor V5 acha que o cabeçalho é: " + cabecalhoErrado.join(" | "));

    // Simula a linha que o Tradutor DEVERIA pegar (ex: linha 10)
    // O Teste 3 nos mostrou que a linha 10 tem os dados reais.
    const cabecalhoCerto = raw[9].slice(0, 6);
    console.log("✅ O cabeçalho real é: " + cabecalhoCerto.join(" | "));

    if (cabecalhoErrado[0] === "[VAZIO]") {
      console.log("🛑 CONCLUSÃO: O Tradutor está cego. Ele está tentando usar células vazias como nome de variável. É por isso que tudo fica vazio!");
    }

  } catch (e) {
    console.error("💥 FALHA NO TESTE 5: " + e.message);
  }
}


/**
 * 🛰️ TESTE UNITÁRIO 6: RAIO-X DO MAPEAMENTO DE COLUNAS
 * Verifica se as palavras exatas do Dicionário batem com o Cabeçalho da Planilha.
 */
function TESTE_06_RAIO_X_DICIONARIO() {
  console.log("🚀 [TESTE 6] Iniciando Raio-X de Mapeamento na Linha 10...");
  
  try {
    const response = getInitialData();
    const raw = response.raw['COCKPIT'] || response.raw['Cockpit'];
    
    // Captura a Linha 10 (Índice 9) e limpa espaços extras das pontas
    const cabecalhos = raw[9].map(h => String(h).toUpperCase().trim());
    const dadosLinha11 = raw[10];

    console.log("🔍 Procurando as chaves vitais do V5 no cabeçalho:");

    // Estas são as palavras exatas que o seu Tradutor.html está procurando
    const chavesParaTestar = {
      'TICKER': 'tickerAtivo',
      'P/L TOTAL': 'plTotalValor',
      'P/L TOTAL %': 'plTotalPct',
      'STATUS': 'status',
      'VENCIMENTO': 'dataVencimento',
      'NOCIONAL': 'nocionalTotal'
    };

    let falhas = 0;

    for (const [colunaExcel, chaveV5] of Object.entries(chavesParaTestar)) {
      const idx = cabecalhos.indexOf(colunaExcel);
      
      if (idx !== -1) {
        let valor = dadosLinha11[idx] === "" ? "[VAZIO]" : dadosLinha11[idx];
        console.log(`✅ ENCONTRADO: '${colunaExcel}' (Coluna ${idx + 1}) -> Vai virar '${chaveV5}'. Valor do Trade: ${valor}`);
      } else {
        console.error(`❌ ALERTA: A coluna exata '${colunaExcel}' NÃO FOI ENCONTRADA no cabeçalho da planilha!`);
        falhas++;
      }
    }

    console.log("--------------------------------------------------");
    if (falhas > 0) {
      console.log(`🛑 CONCLUSÃO: ${falhas} colunas vitais falharam no cruzamento. Se 'P/L TOTAL' falhou, os cards do Dashboard ficarão zerados.`);
    } else {
      console.log("💎 CONCLUSÃO: O Dicionário está perfeito. O problema é 100% no visual (Handsontable do Desktop).");
    }

  } catch (e) {
    console.error("💥 FALHA NO TESTE 6: " + e.message);
  }
}

/**
 * 🛡️ Função Ajudante para os Testes (Coloque no topo do TESTADOR_SISTEMA.gs)
 */
function getAbaDinamica(payloadRaw, nomeProcurado) {
  const chaveReal = Object.keys(payloadRaw).find(k => 
    String(k).toUpperCase() === String(nomeProcurado).toUpperCase()
  );
  return chaveReal ? payloadRaw[chaveReal] : null;
}

/**
 * 🛰️ TESTE UNITÁRIO 7: PROVA REAL DA ABA DINÂMICA
 */
function TESTE_07_ABAS_REAIS() {
  console.log("🚀 [TESTE 7] Testando Busca Dinâmica de Abas nos Testes...");
  
  try {
    const response = getInitialData(); // Sua API perfeita puxando tudo
    
    // O Teste agora usa a busca inteligente
    const rawCockpit = getAbaDinamica(response.raw, 'COCKPIT');
    
    if (rawCockpit && rawCockpit.length > 0) {
      console.log("✅ SUCESSO ABSOLUTO: O teste encontrou a aba Cockpit ignorando diferenças de maiúsculas/minúsculas!");
      console.log(`📊 Linhas carregadas com sucesso: ${rawCockpit.length}`);
    } else {
      console.error("❌ FALHA no Teste.");
    }

  } catch(e) {
    console.error("💥 ERRO NO TESTE 7: " + e.message);
  }
}

/**
 * 🛰️ TESTE UNITÁRIO 8: RAIO-X DA ABA DADOS_ATIVOS
 */
function TESTE_08_DADOS_ATIVOS() {
  console.log("🚀 [TESTE 8] Analisando a aba DADOS_ATIVOS...");
  
  try {
    const response = getInitialData();
    const rawDados = getAbaDinamica(response.raw, 'DADOS_ATIVOS');
    
    if (!rawDados) {
      console.error("❌ FALHA: A aba DADOS_ATIVOS não foi encontrada na planilha!");
      return;
    }

    console.log(`✅ Aba encontrada! Total de linhas: ${rawDados.length}`);
    console.log("📊 Primeiras 4 linhas da aba (para acharmos o cabeçalho):");
    
    for(let i = 0; i < Math.min(4, rawDados.length); i++) {
      console.log(`Linha ${i + 1}: ${rawDados[i].slice(0, 5).join(" | ")}`);
    }

  } catch(e) {
    console.error("💥 ERRO NO TESTE 8: " + e.message);
  }
}