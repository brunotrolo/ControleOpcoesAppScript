/**
 * MÓDULO: Orquestrador_Consultoria
 * OBJETIVO: Executar o fluxo completo (Extrair -> Analisar -> Consultar -> Notificar)
 */

function executarRotinaDiaria() {
  const servicoNome = "Orquestrador_v1.2";
  try {
    gravarLog(servicoNome, "START", "Iniciando ciclo rico");
    
    const diagnostico = processarAnaliseDeRisco();
    const relatorioIA = gerarConsultoriaEstrategica(diagnostico);

    if (relatorioIA) {
      // Passamos o texto da IA E os dados brutos para o e-mail
      enviarEmailConsultoria(relatorioIA, diagnostico.dadosParaIA);
      gravarLog(servicoNome, "FINISH", "Ciclo concluído com sucesso");
    }
  } catch (e) {
    gravarLog(servicoNome, "ERRO_CRITICO", "Falha na orquestração", e.toString());
  }
}


/**
 * FUNÇÃO PARA TRIGGER DIÁRIO (8:00 AM)
 * Orquestra: Motor Existente -> Novo Serviço de E-mail
 */
function orquestrarScannerComEmail() {
  const SERVICO = "Orquestrador_Scanner";
  
  // Usamos o log institucional para monitorar o gatilho
  nxLog(SERVICO, "INICIO", "Iniciando ciclo automático de varredura...", "");

  try {
    // 1. Roda sua Engine (Sem alterações no arquivo dela)
    runMarketScannerEngine(); 

    // 2. Dispara o novo serviço de e-mail blindado
    EnviarRelatorioEmail(); 

    nxLog(SERVICO, "SUCESSO", "Ciclo de varredura e notificação finalizado.", "");
  } catch (e) {
    nxLog(SERVICO, "ERRO_CRITICO", "Falha no disparo automático: " + e.toString(), "");
  }
}