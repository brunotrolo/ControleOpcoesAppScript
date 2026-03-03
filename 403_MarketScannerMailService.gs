/**
 * ═══════════════════════════════════════════════════════════════
 * SERVIÇO: MarketScanner_MailService - v5.0 (PREMIUM REPORT)
 * OBJETIVO: Notificação com todos os indicadores técnicos solicitados.
 * ═══════════════════════════════════════════════════════════════
 */

var SCANNER_DESIGN = {
  primary: '#2563eb', success: '#10b981', warning: '#f59e0b', danger: '#ef4444',
  gray100: '#f3f4f6', gray200: '#e5e7eb', gray700: '#374151', gray900: '#111827'
};

function EnviarRelatorioEmail() {
  const SERVICO_NOME = "MailService_Scanner";
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Scanner_Oportunidades");
    const dados = sheet.getDataRange().getValues();
    dados.shift(); // Remove cabeçalho

    // Filtra apenas as que a IA já analisou
    const perolas = dados.filter(r => r[19] && r[19] !== "" && r[19] !== "Análise pendente...");

    if (perolas.length === 0) return;

    const htmlFinal = templateScannerV50(perolas);

  // Pega o Ticker da primeira pérola (que é a de maior Score pois já está ordenada)
    const topPick = perolas[0][1]; 

    MailApp.sendEmail({
      to: Session.getActiveUser().getEmail(),
      subject: `MarketScanner: ${perolas.length} Oportunidades (Top Pick: ${topPick})`, // Assunto Inteligente
      htmlBody: htmlFinal
    });

    nxLog(SERVICO_NOME, "SUCESSO", "Relatório detalhado enviado com sucesso.", "");
  } catch (e) {
    nxLog(SERVICO_NOME, "ERRO", e.message, "");
  }
}

/**
 * TEMPLATE v5.0: Inclusão de Spot, Opção, Prêmio, Venc, DTE, Taxa Diária e Breakeven.
 */
function templateScannerV50(perolas) {
  const dataRef = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm");
  let htmlCards = "";

  perolas.forEach(p => {
    /* MAPEAMENTO DE ÍNDICES:
       p[1]=Ticker, p[2]=Spot, p[4]=Opção, p[5]=Strike, p[6]=Prêmio, 
       p[7]=Venc, p[8]=DTE, p[9]=Perfil, p[10]=Delta, p[14]=ROI, 
       p[15]=Taxa Diária, p[16]=Breakeven, p[17]=Dist, p[18]=Score, p[19]=IA
    */

  // Formata o vencimento
    let vencLimpo = (p[7] instanceof Date) ? Utilities.formatDate(p[7], "GMT-3", "dd/MM/yy") : p[7];

    // --- FORMATAÇÃO PARA E-MAIL (ARREDONDAMENTOS) ---
    let valSpot = p[2];
    let txtSpot = (valSpot instanceof Date) ? "R$ --" : "R$ " + Number(valSpot).toFixed(2).replace('.', ',');

    // Multiplica decimais por 100 para virar %
    let txtROI = (Number(p[14]) * 100).toFixed(2) + "%";
    let txtMargem = (Number(p[17]) * 100).toFixed(2) + "%";
    
    // Taxa Diária: Se vier decimal pequeno (0.006), multiplica.
    let rawTaxa = Number(p[15]);
    let txtTaxa = (rawTaxa < 1) ? (rawTaxa * 100).toFixed(2) + "%" : rawTaxa.toFixed(2) + "%";

    // Arredonda valores numéricos
    let txtScore = Number(p[18]).toFixed(2);
    // Lógica visual de estrelas baseada no Score Nexo
    let rawScore = Number(p[18]);
    let estrelas = "";
    if (rawScore >= 25) estrelas = "⭐⭐⭐⭐⭐ (Elite)";
    else if (rawScore >= 20) estrelas = "⭐⭐⭐⭐ (Muito Forte)";
    else if (rawScore >= 15) estrelas = "⭐⭐⭐ (Forte)";
    else if (rawScore >= 10) estrelas = "⭐⭐ (Moderado)";
    else estrelas = "⭐ (Especulativo)";

    let txtDelta = Number(p[10]).toFixed(2);
    let txtBreakeven = "R$ " + Number(p[16]).toFixed(2).replace('.', ',');

    let perfil = p[9];


    let corPerfil = perfil.includes("Conservador") ? SCANNER_DESIGN.success : (perfil.includes("Agressivo") ? SCANNER_DESIGN.danger : SCANNER_DESIGN.warning);

    htmlCards += `
      <div style="background:#ffffff; border:1px solid ${SCANNER_DESIGN.gray200}; border-left:6px solid ${corPerfil}; border-radius:10px; padding:18px; margin-bottom:20px; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">
        
        <div style="margin-bottom:12px; border-bottom:1px solid #f0f0f0; padding-bottom:8px;">
          <span style="font-size:18px; font-weight:800; color:${SCANNER_DESIGN.gray900};">${p[1]} (Spot: ${txtSpot})</span>
          <span style="float:right; font-size:12px; font-weight:bold; color:${corPerfil};">${perfil.toUpperCase()}</span>
          <div style="margin-top:4px;">
            <span style="background:${SCANNER_DESIGN.gray100}; color:${SCANNER_DESIGN.gray700}; padding:3px 8px; border-radius:5px; font-size:11px; font-weight:bold; border:1px solid #ddd;">${p[4]}</span>
            <span style="background:#dbeafe; color:#1e40af; padding:3px 8px; border-radius:5px; font-size:11px; font-weight:bold; margin-left:8px;">Venc: ${vencLimpo} (${p[8]} DTE)</span>
          </div>
        </div>

        <div style="display: table; width: 100%; margin-bottom: 15px;">
          <div style="display: table-row;">
            <div style="display: table-cell; padding: 5px; font-size: 11px; color: #666;">
              <b>STRIKE:</b><br><span style="font-size:13px; color:#333;">R$ ${p[5]}</span>
            </div>
            <div style="display: table-cell; padding: 5px; font-size: 11px; color: #666;">
              <b>PRÊMIO:</b><br><span style="font-size:13px; color:#333;">R$ ${p[6]}</span>
            </div>
            <div style="display: table-cell; padding: 5px; font-size: 11px; color: #666;">
              <b>BREAKEVEN:</b><br><span style="font-size:13px; color:#333;">${txtBreakeven}</span>
            </div>
          </div>
          <div style="display: table-row;">
            <div style="display: table-cell; padding: 5px; font-size: 11px; color: #666;">
              <b>DELTA:</b><br><span style="font-size:13px; color:#333;">${txtDelta}</span>
            </div>
            <div style="display: table-cell; padding: 5px; font-size: 11px; color: #666;">
              <b>ROI:</b><br><span style="font-size:13px; color:#059669; font-weight:bold;">${txtROI}</span>
            </div>
            <div style="display: table-cell; padding: 5px; font-size: 11px; color: #666;">
              <b>TAXA DIÁRIA:</b><br><span style="font-size:13px; color:#059669; font-weight:bold;">${txtTaxa}</span>
            </div>
          </div>
        </div>

        <div style="font-size:12px; color:#374151; background:#fffdf2; padding:12px; border-radius:6px; border-left:3px solid #f1c40f; line-height:1.6;">
          <strong>🎯 Veredito Estratégico:</strong><br>
          ${p[19].replace(/\n/g, '<br>')}
        </div>

        <div style="margin-top:10px; font-size:11px; color:#555; text-align:right; background-color:#f8f9fa; padding:5px; border-radius:4px;">
          Score: <b>${txtScore}</b> <span style="margin-left:5px;">${estrelas}</span>
          <span style="color:#ccc; margin:0 8px;">|</span>
          Margem: <b>${txtMargem}</b>
        </div>

      </div>`;
  });

  return `
    <!DOCTYPE html><html><body style="margin:0; padding:0; background:#f3f4f6; font-family: 'Segoe UI', Arial, sans-serif;">
      <div style="max-width:650px; margin:0 auto; background:#ffffff; min-height:100vh; padding:25px; border:1px solid #eee;">
        <div style="text-align:center; margin-bottom:30px;">
        </div>
        
        <p style="font-size:14px; color:#444;">🎯 Relatório gerado em <b>${dataRef}</b>. As seguintes oportunidades atendem aos critérios institucionais de Volatilidade e Delta:</p>
        
        ${htmlCards}
        
        <div style="margin-top:40px; padding:20px; background:#f9fafb; border-radius:8px; border:1px solid #eee; text-align:center;">
          <p style="font-size:11px; color:#999; margin:0;">
            <b>Nota Técnica:</b> ROI e Taxa Diária baseados no Strike. Breakeven calculado subtraindo o prêmio recebido do Strike.
            Monitore o DTE para a regra de encerramento em 21 dias.
          </p>
        </div>
      </div>
    </body></html>`;
}