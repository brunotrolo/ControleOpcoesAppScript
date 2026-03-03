/**
 * ═══════════════════════════════════════════════════════════════
 * ARQUIVO: ScannerTendenciaOportunidades_MailService
 * VERSÃO: 7.0 (The Hedge Fund Report)
 * OBJETIVO: Relatório de Inteligência Institucional via E-mail
 * ═══════════════════════════════════════════════════════════════
 */

// Paleta de Cores Institucional (Dark Mode Header + Clean Cards)
var NEXO_THEME = {
  headerBg: '#0f172a',
  headerText: '#f8fafc',
  cardBg: '#ffffff',
  textMain: '#334155',
  textMuted: '#64748b',
  border: '#e5e7eb', // gray200 da v5.0
  // Perfis
  conservative: { color: '#10b981', bg: '#ecfdf5', border: '#10b981' }, 
  moderate:     { color: '#f59e0b', bg: '#fffbeb', border: '#f59e0b' }, 
  aggressive:   { color: '#ef4444', bg: '#fff1f2', border: '#ef4444' }, 
  // IA Box (Estilo v5.0)
  aiBg: '#fffdf2', 
  aiBorder: '#f1c40f' 
};

// --- CONFIGURAÇÃO DE FONTE (PADRÃO v5.0 PREMIUM) ---
const FONT_FAMILY = "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif";
const FONT_PREMIUM = "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif";

// --- FORMATAÇÃO DE DADOS ---
function fmtMoney_(v) { return "R$ " + Number(v).toFixed(2).replace('.', ','); }
function fmtPercent_(v) { return (Number(v) * 100).toFixed(2) + "%"; }
function fmtDate_(d) { return (d instanceof Date) ? Utilities.formatDate(d, "GMT-3", "dd/MM") : d; }

function getStarRating_(score) {
  const s = Number(score);
  // Estrelas visuais para e-mail
  if (s >= 9.0) return `<span style="color:#eab308">★★★★★</span> (Elite)`;
  if (s >= 7.5) return `<span style="color:#eab308">★★★★☆</span> (Forte)`;
  if (s >= 6.0) return `<span style="color:#eab308">★★★☆☆</span> (Bom)`;
  return `<span style="color:#94a3b8">★★☆☆☆</span> (Neutro)`;
}

/**
 * 🚀 FUNÇÃO PRINCIPAL
 * Dispara o relatório lendo destinatário da Config_Global
 */
function enviarRelatorioNexoMaster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaScanner = ss.getSheetByName("Scanner_Tendencia_Oportunidades");
  const abaConfig = ss.getSheetByName("Config_Global");

  // 1. OBTEM CONFIGURAÇÃO DE E-MAIL
  const dadosConfig = abaConfig.getDataRange().getValues();
  const rowEmail = dadosConfig.find(r => r[0] === "Email_Alerta_Padrao");
  const destinatario = (rowEmail && rowEmail[1]) ? rowEmail[1] : Session.getActiveUser().getEmail();

  // 2. OBTEM DADOS DO SCANNER
  const dados = abaScanner.getDataRange().getValues();
  dados.shift(); // Remove cabeçalho
  
  // Filtra Top 9: Onde a Coluna Z (IA) não está vazia e não é "AGUARDAR"
  // Coluna 25 do array = Coluna Z na planilha
  const opportunities = dados.filter(r => r[25] && r[25].length > 10 && r[25] !== "AGUARDAR");

  if (opportunities.length === 0) {
    console.log("Nenhuma oportunidade analisada pela IA para enviar.");
    return;
  }

  const dataHora = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm");

  let htmlBodyContent = "";

  // 3. CONSTRUÇÃO DOS CARDS POR PERFIL
  const profiles = [
    { key: "Conservador", label: "🛡️ PERFIL CONSERVADOR", style: NEXO_THEME.conservative },
    { key: "Moderado",    label: "⚖️ PERFIL MODERADO",    style: NEXO_THEME.moderate },
    { key: "Agressivo",   label: "🚀 PERFIL AGRESSIVO",   style: NEXO_THEME.aggressive }
  ];

  profiles.forEach(profile => {
    // Filtra ativos que contenham a chave (ex: "Conservador") na Coluna Y (Índice 24)
    const ativosPerfil = opportunities.filter(r => r[24].toString().includes(profile.key));
    
    if (ativosPerfil.length > 0) {
      // Título da Seção
      htmlBodyContent += `
        <div style="margin: 30px 0 15px 0; border-bottom: 2px solid ${profile.style.border}; padding-bottom: 5px;">
          <span style="font-family:'Helvetica', sans-serif; font-size:12px; font-weight:900; color:${profile.style.color}; letter-spacing:1px;">
            ${profile.label}
          </span>
        </div>`;


      // Cards dos Ativos

      ativosPerfil.forEach(atv => {
        htmlBodyContent += `
          <div style="background:#ffffff; border:1px solid ${NEXO_THEME.border}; border-left:6px solid ${profile.style.border}; border-radius:10px; padding:18px; margin-bottom:20px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); font-family:${FONT_PREMIUM};">
            
            <div style="margin-bottom:12px; border-bottom:1px solid #f0f0f0; padding-bottom:8px;">
              <span style="font-size:18px; font-weight:800; color:#111827;">${atv[0]} (Spot: ${fmtMoney_(atv[3])})</span>
              <span style="float:right; font-size:12px; font-weight:bold; color:${profile.style.color};">${profile.key.toUpperCase()}</span>
              <div style="margin-top:4px;">
                <span style="background:#f3f4f6; color:#374151; padding:3px 8px; border-radius:5px; font-size:11px; font-weight:bold; border:1px solid #ddd;">${atv[2]}</span>
                <span style="background:#dbeafe; color:#1e40af; padding:3px 8px; border-radius:5px; font-size:11px; font-weight:bold; margin-left:8px;">Venc: ${fmtDate_(atv[13])} (${atv[14]} DTE)</span>
              </div>
            </div>

            <div style="display: table; width: 100%; margin-bottom: 15px;">
              <div style="display: table-row;">
                <div style="display: table-cell; padding: 5px; font-size: 11px; color: #666; width:33%;">
                  <b>STRIKE:</b><br><span style="font-size:13px; color:#333;">${fmtMoney_(atv[4])}</span>
                </div>
                <div style="display: table-cell; padding: 5px; font-size: 11px; color: #666; width:33%;">
                  <b>PRÊMIO:</b><br><span style="font-size:13px; color:#333;">${fmtMoney_(atv[18])}</span>
                </div>
                <div style="display: table-cell; padding: 5px; font-size: 11px; color: #666; width:33%;">
                  <b>BREAKEVEN:</b><br><span style="font-size:13px; color:#333;">${fmtMoney_(atv[17])}</span>
                </div>
              </div>
              <div style="display: table-row;">
                <div style="display: table-cell; padding: 5px; font-size: 11px; color: #666;">
                  <b>DELTA:</b><br><span style="font-size:13px; color:#333;">${Number(atv[10]).toFixed(3)}</span>
                </div>
                <div style="display: table-cell; padding: 5px; font-size: 11px; color: #666;">
                  <b>ROI:</b><br><span style="font-size:13px; color:#059669; font-weight:bold;">${fmtPercent_(atv[19])}</span>
                </div>
                <div style="display: table-cell; padding: 5px; font-size: 11px; color: #666;">
                  </div>
              </div>
            </div>

            <div style="font-size:12px; color:#374151; background:#fffdf2; padding:12px; border-radius:6px; border-left:3px solid #f1c40f; line-height:1.6;">
              <strong>🎯 Veredito Estratégico:</strong><br>
              ${atv[25].replace(/\n/g, '<br>')}
            </div>

            <div style="margin-top:10px; font-size:11px; color:#555; text-align:right; background-color:#f8f9fa; padding:5px; border-radius:4px;">
              Score: <b>${atv[22]}</b> <span style="margin-left:5px;">${getStarRating_(atv[22])}</span>
              <span style="color:#ccc; margin:0 8px;">|</span>
              Contexto: <b>${atv[23]}</b>
            </div>

          </div>`;
      });

    }
  });



// --- GLOSSÁRIO COMPLETO (COM PETRÓLEO E IBOV DETALHADO) ---
  htmlBodyContent += `
    <div style="margin-top:20px; padding:18px; border:1px solid ${NEXO_THEME.border}; border-radius:10px; background-color: #f9fafb; font-family:${FONT_PREMIUM};">
      <div style="font-size:11px; font-weight:800; color:${NEXO_THEME.textMain}; margin-bottom:12px; text-transform:uppercase; letter-spacing:1px; border-bottom:1px solid #eee; padding-bottom:5px;">
        📚 Legenda de Contexto & Sinais
      </div>
      
      <div style="font-size:10px; color:${NEXO_THEME.textMuted}; line-height:1.6;">
        <div style="margin-bottom:12px;">
          <strong style="color:${NEXO_THEME.textMain};">Sinais de Mercado:</strong><br>
          • <strong>📈 IBOV BULL:</strong> Tendência de Alta no índice. Força compradora predominante (Otimismo).<br>
          • <strong>📉 IBOV BEAR:</strong> Tendência de Baixa no índice. Pressão vendedora ou correção (Cautela).<br>
          • <strong>🏗️ VALE/MINÉRIO+:</strong> Alta correlação com Minério de Ferro/China. Vetor positivo para Vale/Siderúrgicas.<br>
          • <strong>🛢️ PETRÓLEO+:</strong> Alta correlação com Brent/WTI. Vetor positivo para Petrobras, Prio e 3R.<br>
          • <strong>⚡ VOL+:</strong> Volatilidade Implícita alta. Prêmios "gordos", ideal para estratégias de venda (Theta).
        </div>

        <div>
          <strong style="color:${NEXO_THEME.textMain};">Indicadores Técnicos:</strong><br>
          • <strong>Delta:</strong> Sensibilidade da opção. Um Delta 0.30 indica ~30% de chance de exercício.<br>
          • <strong>BreakEven:</strong> Ponto de empate. O preço que o ativo precisa atingir para a operação sair no zero a zero.<br>
          • <strong>ROI:</strong> Retorno potencial sobre a margem/capital em risco (se levado ao vencimento).
        </div>
      </div>
    </div>`;



// 4. MONTAGEM FINAL DO E-MAIL
  const emailTemplate = `
  <!DOCTYPE html>
  <html>
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
  </head>
  <body style="margin:0; padding:0; background-color:#f1f5f9; font-family:${FONT_FAMILY};">
    <div style="max-width:600px; margin:0 auto; background-color:#ffffff; min-height:100vh;">
      
      <div style="padding:20px;">
        
        <p style="font-size:14px; color:#444; line-height:1.6; text-align:center; margin-bottom:25px; font-family:${FONT_PREMIUM};">
          🎯 Relatório gerado em <b>${dataHora}</b>. O sistema identificou <b>${opportunities.length} oportunidades de alta confluência</b> hoje, validado pelo modelo Gemini AI.
        </p>

        ${htmlBodyContent}


        <div style="margin-top:40px; border-top:1px solid ${NEXO_THEME.border}; padding-top:20px; text-align:center;">
          <p style="font-size:10px; color:#94a3b8; line-height:1.4; margin:0;">
            Este relatório é gerado automaticamente pelo Ecossistema Nexo v6.0.<br>
            As análises (Score, Delta, ROI) são quantitativas. A validação de texto é gerada por IA.<br>
            Não constitui recomendação direta de investimento. Gestão de risco é mandatória.
          </p>
        </div>

      </div>
    </div>
  </body>
  </html>
  `;

  // 5. DISPARO
  try {
    MailApp.sendEmail({
      to: destinatario,
      subject: `Stock Options Intelligence: ${opportunities.length} Oportunidades (${dataHora})`,
      htmlBody: emailTemplate
    });
    console.log(`Relatório enviado com sucesso para: ${destinatario}`);
  } catch (e) {
    console.error("Erro ao enviar e-mail: " + e.message);
  }
}