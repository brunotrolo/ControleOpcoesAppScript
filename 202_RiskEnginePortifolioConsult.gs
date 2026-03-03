/**
 * MÓDULO: Servico_RiskEngine - v1.3
 */
function processarAnaliseDeRisco() {
  const servicoNome = "RiskEngine_v1.3";
  const configs = obterConfigsGlobais();
  const operacoes = homologarExtracaoCockpit(); 
  if (!operacoes || !configs) return null;

  // ORDENAÇÃO POR DTE: Garante que tudo (logs, alertas e IA) siga a ordem cronológica
  operacoes.sort((a, b) => (a['DTE CORRIDOS'] || 0) - (b['DTE CORRIDOS'] || 0));

  // Normalização das metas com log de conferência
  const lucroMin = (configs['Regra_Lucro_Alvo_Min'] !== undefined ? configs['Regra_Lucro_Alvo_Min'] : 0.6) * 100;
  const lucroMax = (configs['Regra_Lucro_Alvo_Max'] !== undefined ? configs['Regra_Lucro_Alvo_Max'] : 0.7) * 100;
  const dteAlvo = configs['Regra_DTE_Saida_Alvo'] || 21;

  gravarLog(servicoNome, "PARAMETROS", "Regras Aplicadas nesta rodada", 
            `Lucro Alvo: ${lucroMin}% a ${lucroMax}% | DTE Alvo: < ${dteAlvo}`);

  let alertasEncontrados = [];
  let resumoNotional = 0;

  operacoes.forEach((op, index) => {
    const ticker = op['CÓDIGO OPÇÃO'] || "N/A";
    const lucroPercentual = (op['P/L TOTAL %'] || 0) * 100;
    const dte = op['DTE CORRIDOS'] || 0;
    const notionalOp = op['NOCIONAL'] || 0;
    const moneyness = op['MONEYNESS_CODE'] || "N/A";
    
    resumoNotional += notionalOp;

    gravarLog(servicoNome, "DEBUG", `Analisando ${ticker} [${index + 1}/${operacoes.length}]`, 
              `Lucro: ${lucroPercentual.toFixed(2)}% | DTE: ${dte} | Notional: R$ ${notionalOp.toFixed(2)}`);

    // VERIFICAÇÃO DE LUCRO (Corrigida)
    if (lucroPercentual >= lucroMin) {
      const isCritico = lucroPercentual >= lucroMax;
      const statusLucro = isCritico ? "LUCRO_CRITICO" : "LUCRO_ALVO";
      const acaoLucro = isCritico ? "ENCERRAR IMEDIATO" : "CONSIDERAR FECHAMENTO";
      
      alertasEncontrados.push({ticker: ticker, motivo: "Lucro", valor: lucroPercentual.toFixed(2), acao: acaoLucro});
      gravarLog(servicoNome, statusLucro, `${ticker}: Meta atingida!`, `Valor: ${lucroPercentual.toFixed(2)}% | Regra: >${lucroMin}%`);
    }
    
    // ... (restante do código: DTE e ITM permanecem iguais)
    if (dte <= dteAlvo && dte > 0) {
      alertasEncontrados.push({ticker: ticker, motivo: "DTE", valor: dte, acao: "Rolar ou Fechar"});
      gravarLog(servicoNome, "ALERTA_TEMPO", `${ticker}: Risco de Gamma`, `DTE Atual: ${dte} | Regra: <${dteAlvo}`);
    }
    if (moneyness === "ITM") {
      alertasEncontrados.push({ticker: ticker, motivo: "ITM", valor: "ITM", acao: "Defesa/Monitorar"});
      gravarLog(servicoNome, "ALERTA_RISCO", `${ticker} está ITM`, `Strike: ${op['STRIKE']} | Spot: ${op['TICKER SPOT']}`);
    }
  });

  gravarLog(servicoNome, "SUCESSO", "Análise concluída e ordenada por DTE", `Alertas: ${alertasEncontrados.length}`);

  return {
    alertas: alertasEncontrados,
    notionalTotal: resumoNotional,
    dadosParaIA: operacoes,
    configsUsadas: configs
  };
}