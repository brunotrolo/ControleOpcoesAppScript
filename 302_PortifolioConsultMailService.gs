/**
 * MÓDULO: Servico_MailService - v4.8 (Versão Master Final)
 */

var DESIGN_COLORS = {
  primary: '#2563eb', success: '#10b981', danger: '#ef4444', warning: '#f59e0b',
  gray100: '#f3f4f6', gray200: '#e5e7eb', gray700: '#374151', gray900: '#111827',
  info: '#06b6d4'
};

// --- FUNÇÕES DE APOIO ---
function fmtM_(v) { 
  return "R$ " + Number(v).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ".").replace(".", ","); 
}

function badge_(t, tp) {
  var c = { 
    danger: ['#fee2e2','#991b1b'], 
    warning: ['#fef3c7','#92400e'], 
    success: ['#d1fae5','#065f46'], 
    info: ['#dbeafe','#1e40af'] 
  }[tp] || ['#f3f4f6','#374151'];
  return `<span style="display:inline-block;padding:2px 8px;border-radius:10px;font-size:10px;font-weight:600;background:${c[0]};color:${c[1]};white-space:nowrap;margin-right:4px;">${t}</span>`;
}

// --- FUNÇÃO PRINCIPAL DE ENVIO ---
function enviarEmailConsultoria(textoIA, dadosBrutos) {
  const servicoNome = "MailService_v4.8"; 
  const configs = obterConfigsGlobais();
  const destinatario = configs['Email_Alerta_Padrao'] || Session.getActiveUser().getEmail();
  const dataRef = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm");
  const assunto = `Consultoria Portifólio ${dataRef}`;

  let htmlCards = "";
  let linhas = textoIA.split('\n');
  let sentinelaNotional = false; 
  let sentinelaStatus = false;
  // --- CORREÇÃO 1: Variável Global do Loop ---
  let corContexto = DESIGN_COLORS.primary;


  linhas.forEach(linha => {
    linha = linha.trim();
    if (linha === "" || linha.length < 5) return;

    // 1. Identificação de Cabeçalhos
    if (linha.match(/PRIORIDADE|MONITORIZAÇÃO|NOTIONAL|INSIGHT|STATUS/i)) {
      
      // --- CORREÇÃO 2: Lógica de Cores (Vermelho, Laranja e Verde) ---
      if (linha.includes("CRÍTICA")) {
        corContexto = DESIGN_COLORS.danger;   
      } else if (linha.includes("ALERTA")) {
        corContexto = DESIGN_COLORS.warning;  
      } else if (linha.includes("MONITORIZAÇÃO")) {
        corContexto = DESIGN_COLORS.success;  // Verde para Monitorização
      } else {
        corContexto = DESIGN_COLORS.primary;  // Azul para o resto
      }

      // Caso A: Título de Status
      if (linha.includes("STATUS DO SISTEMA")) {
        sentinelaStatus = true;
        htmlCards += `<div style="font-size:12px; font-weight:800; color:${corContexto}; margin:25px 0 10px; border-bottom:2px solid ${DESIGN_COLORS.gray200}; text-transform:uppercase;">${linha.replace(/\*/g,'')}</div>`;
        htmlCards += `
          <div style="margin-top:10px; padding:15px; border:1px solid ${DESIGN_COLORS.gray200}; border-radius:8px; background-color: #f9fafb;">
            <div style="font-size:11px; font-weight:800; color:${DESIGN_COLORS.gray900}; margin-bottom:10px; text-transform:uppercase;">📚 Glossário de Campo</div>
            <div style="font-size:10px; color:#555; line-height:1.6; border-bottom:1px solid #eee; padding-bottom:10px; margin-bottom:10px;">
              <strong>• Delta:</strong> Direção. Quanto a opção oscila para cada R$ 1 do ativo.<br>
              <strong>• Gamma:</strong> Aceleração. O risco de mudança brusca perto do vencimento.<br>
              <strong>• Theta:</strong> Tempo. Seu lucro diário pelo "derretimento" da opção.<br>
              <strong>• Vega:</strong> Volatilidade. O impacto do medo/incerteza no prêmio.<br>
              <strong>• Notional:</strong> Exposição total controlada pela operação.
            </div>`;
        return; 
      }


      // Caso B: Título de Notional (Versão Auto-Extração)
      // TRAVA 1: Só entra se for título mesmo (sem barra de dados)
      if (linha.toUpperCase().includes("NOTIONAL") && !linha.includes("|")) {
        htmlCards += `<div style="font-size:12px; font-weight:800; color:${corContexto}; margin:25px 0 10px; border-bottom:2px solid ${DESIGN_COLORS.gray200}; text-transform:uppercase;">${linha.replace(/\*/g,'')}</div>`;

        // Se a linha da IA já tiver o valor (como no seu exemplo), tentamos extrair agora
        let valorDetectado = "Consultando...";
        if (linha.includes("R$")) {
           valorDetectado = "R$ " + linha.split("R$")[1].split("|")[0].trim();
        }

        htmlCards += `
          <div style="background:#ffffff; border:1px solid ${DESIGN_COLORS.gray200}; border-left:5px solid ${DESIGN_COLORS.primary}; border-radius:8px; padding:15px; margin:20px 0; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">
            <div style="font-size:10px; text-transform:uppercase; color:${DESIGN_COLORS.gray700}; margin-bottom:5px; font-weight:700; letter-spacing:1px;">💰 Exposição Total (Notional)</div>
            <div id="notional-val" style="font-size:22px; font-weight:800; color:${DESIGN_COLORS.primary}; margin-bottom:10px;">${valorDetectado}</div>
            <div id="veredito-space" style="font-size:12px; line-height:1.5; color:${DESIGN_COLORS.gray700}; border-top:1px solid ${DESIGN_COLORS.gray100}; padding-top:10px;">
              <strong>Veredito Estratégico:</strong> ...carregando...
            </div>
          </div>`;
        
        sentinelaNotional = true; 
        return; 
      }


      // Caso C: Outros Títulos (Prioridades e Insights)
      htmlCards += `<div style="font-size:12px; font-weight:800; color:${corContexto}; margin:25px 0 10px; border-bottom:2px solid ${DESIGN_COLORS.gray200}; text-transform:uppercase;">${linha.replace(/\*/g,'')}</div>`;


      return; // Mata a linha aqui para evitar o duplicado no final
    }
    
    // 2. Captura do Veredito e Correção de Valor
    if (sentinelaNotional) {
      let textoLimpo = linha.replace(/\*/g,'');

      // --- TRAVA 2 (ADICIONE ESTA LINHA): Ignora linhas de "papo furado" da IA ---
      if (!textoLimpo.includes("R$") && !textoLimpo.includes("|")) return;
      
      // Se detectarmos o valor de R$ no texto que a IA mandou agora
      if (textoLimpo.includes("R$")) {
        let valorExtraido = "R$ " + textoLimpo.split("R$")[1].split("|")[0].trim();
        htmlCards = htmlCards.replace("Consultando...", valorExtraido);
      }
      
      // Limpa o texto do veredito para não repetir a palavra "EXPOSIÇÃO" no campo de baixo
      let vereditoFinal = textoLimpo.includes("|") ? textoLimpo.split("|")[1].replace(/VEREDITO:/i, '').trim() : textoLimpo;

      htmlCards = htmlCards.replace("...carregando...", vereditoFinal);
      sentinelaNotional = false;
      return;
    }

    // AJUSTE: Captura o texto da IA (ex: "Todos os dados...") para dentro da mesma div
    else if (sentinelaStatus) {
      htmlCards += `<p style="font-size:11px; color:#666; margin:0; line-height:1.4;">${linha.replace(/\*/g,'')}</p>`;
      return; 
    }


    // 3. Lógica de Ativos (Cards Individuais - v4.9 Enriquecida)
    let temMarcador = linha.includes("|") || linha.includes("MOTIVO:") || linha.includes("AÇÃO:");

    if (temMarcador) {
      let tickerAlvo = linha.split(/MOTIVO:|\|/i)[0].replace(/TICKER:|\*|:/gi, '').trim();
      let d = dadosBrutos.find(op => op['CÓDIGO OPÇÃO'].toString().trim().toUpperCase() === tickerAlvo.toUpperCase());

    if (d) {
        let partesSecoes = linha.split(/MOTIVO:|AÇÃO:/i);
        let motivoIA = (partesSecoes[1] || "").split("|")[0].trim();
        let acaoIA = (partesSecoes[2] || "MONITORAR").replace(/[|*]/g, '').trim();

        // Se for crítico é Vermelho, senão usa a cor da seção (que pode ser Verde, Laranja ou Azul)
        // CORREÇÃO: Removemos 'DEFESA' e 'ROLAR' do vermelho forçado. 
        // Agora eles respeitam a cor do contexto (Laranja no item 2).
        let corAcao = acaoIA.match(/ENCERRAR|IMEDIATO|FECHAR/i) ? DESIGN_COLORS.danger : corContexto;


        let venc = Utilities.formatDate(new Date(d['VENCIMENTO']), "GMT-3", "dd/MM");
        
        // --- MAPEAMENTO INTELIGENTE (Tenta vários nomes possíveis) ---
        let qtde = d['QTDE'] || d['QTD'] || d['QUANTIDADE'] || 0;
        let tickerAtivo = d['TICKER ATIVO'] || d['ATIVO'] || d['TICKER'] || "";
        let delta = d['DELTA'] ? Number(d['DELTA']).toFixed(2) : "0.00";
        let theta = d['THETA'] ? "R$ " + Number(d['THETA']).toFixed(2) : "N/A";
        let notionalVal = d['NOCIONAL'] || d['NOTIONAL'] || 0;

        htmlCards += `
          <div style="background:#ffffff; border:1px solid ${DESIGN_COLORS.gray200}; border-left:5px solid ${corAcao}; border-radius:8px; padding:15px; margin-bottom:15px; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">
            <div style="margin-bottom:8px;">
              <span style="font-size:16px; font-weight:800; color:${DESIGN_COLORS.gray900};">${tickerAtivo} › ${tickerAlvo}</span>
              <span style="float:right; font-size:11px; color:#666;">📅 ${venc} (${d['DTE CORRIDOS']}d)</span>
              <div style="margin-top:4px;">
                ${badge_(qtde + " un", 'info')} 
                ${badge_(d['MONEYNESS_CODE'], d['MONEYNESS_CODE']=='ITM'?'danger':'success')}
                ${badge_("Δ " + delta, 'gray')}
              </div>
            </div>

            <div style="font-size:10px; color:${DESIGN_COLORS.gray700}; background:${DESIGN_COLORS.gray100}; padding:10px; border-radius:4px; margin-bottom:10px; line-height:1.5;">
              <table style="width:100%; border-collapse:collapse;">
                <tr>
                  <td style="padding:2px 0;"><strong>Spot:</strong> ${fmtM_(d['TICKER SPOT'])}</td>
                  <td style="padding:2px 0;"><strong>Strike:</strong> ${fmtM_(d['STRIKE'])}</td>
                  <td style="padding:2px 0;"><strong>PM:</strong> ${fmtM_(d['PRÊMIO (PM)'])}</td>
                  <td style="padding:2px 0;"><strong>Atual:</strong> ${fmtM_(d['PRÊMIO ATUAL'])}</td>
                </tr>
                <tr>
                  <td style="padding:2px 0;"><strong>Lucro:</strong> <span style="color:${d['P/L TOTAL'] >= 0 ? DESIGN_COLORS.success : DESIGN_COLORS.danger}; font-weight:bold;">${fmtM_(d['P/L TOTAL'])}</span></td>
                  <td style="padding:2px 0;"><strong>L. Máx:</strong> ${(Number(d['LUCRO MAX']) * 100).toFixed(2)}%</td>
                  <td style="padding:2px 0;" colspan="2"><strong>💵 Notional:</strong> ${fmtM_(notionalVal)}</td>
                </tr>
              </table>
            </div>

            <div style="font-size:12px; color:#444; margin-bottom:12px; padding-left:5px; border-left:2px solid #ddd;">
              <strong>Racional:</strong> ${motivoIA}
            </div>
            
            <div style="font-size:11px; font-weight:700; color:${corAcao}; background:${corAcao}15; padding:6px 12px; border-radius:4px; display:inline-block; border:1px solid ${corAcao}40; text-transform:uppercase;">
              🎯 ${acaoIA}
            </div>
          </div>`;
        return; 
      }
    }
    
    // 4. Fallback para textos simples
    htmlCards += `<p style="font-size:13px; color:#555; line-height:1.6; margin:10px 0;">${linha.replace(/\*/g,'')}</p>`;
  });

  if (sentinelaStatus) { htmlCards += `</div>`; } // ADICIONE ESTA LINHA

  const htmlFinal = `
    <!DOCTYPE html><html><body style="margin:0; padding:0; background:#f3f4f6; font-family:-apple-system, sans-serif;">
      <div style="max-width:600px; margin:0 auto; background:#ffffff; min-height:100vh;">
        <div style="padding:20px;">
          <h2 style="color:${DESIGN_COLORS.gray900}; border-bottom:2px solid ${DESIGN_COLORS.gray200}; padding-bottom:10px; font-size:18px; margin-top:0;">🚀 Resumo - ${dataRef}</h2>
          ${htmlCards}
          <div style="margin-top:30px; padding:15px; background:#f9fafb; text-align:center; font-size:9px; color:#999; border-radius:8px;">
            Gerado automaticamente · Nexo Opções 2026
          </div>
        </div>
      </div>
    </body></html>`;

  MailApp.sendEmail({ to: destinatario, subject: assunto, htmlBody: htmlFinal });
  gravarLog(servicoNome, "SUCESSO", "Relatório Rico enviado", destinatario);
}