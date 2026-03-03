/**
 * SERVIÇO 601: Monitor de Dados Macro
 * Responsabilidade: Apenas ler os dados calculados pelas fórmulas e registrar log.
 */
function syncMacroData_Execute() {
  const t0 = Date.now();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Dados_Macro_Setorial");
  
  try {
    // 1. Apenas lê os valores que as fórmulas do GoogleFinance já calcularam
    const data = sheet.getRange("A2:B15").getValues();
    
    // Filtra apenas itens que têm valor para o log
    const resumo = data.filter(r => r[0] !== "").map(r => `${r[0]}: ${r[1]}`);

    // 2. Grava o log de sucesso (Consumindo o arquivo 002)
    log("601_SyncMacroData", "SUCESSO", "Monitoramento Macro registrado", {
      duracao: (Date.now() - t0) + "ms",
      snapshot: resumo
    });

    console.log("✅ 601: Monitoramento macro concluído silenciosamente.");

  } catch (e) {
    log("601_SyncMacroData", "ERRO", "Falha ao ler dados macro", e.toString());
  }
}