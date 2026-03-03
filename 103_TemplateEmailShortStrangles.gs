/*
 * TEMPLATE_EMAIL_STRANGLES.GS
 * Versão harmonizada com DESIGN_COLORS do Template_Email_Operacoes
 * (somente design; nenhuma funcionalidade foi removida)
 */

/************************************************************
 * PALETA — reutiliza DESIGN_COLORS do outro template
 ************************************************************/
var STRANGLE_COLORS = (typeof DESIGN_COLORS !== 'undefined')
  ? DESIGN_COLORS
  : {
      primary: '#2563eb',
      success: '#10b981',
      danger: '#ef4444',
      warning: '#f59e0b',
      info: '#06b6d4',
      gray50: '#f9fafb',
      gray100: '#f3f4f6',
      gray200: '#e5e7eb',
      gray300: '#d1d5db',
      gray400: '#9ca3af',
      gray500: '#6b7280',
      gray600: '#4b5563',
      gray700: '#374151',
      gray800: '#1f2937',
      gray900: '#111827',
      callColor: '#3b82f6',
      putColor: '#f59e0b',
      spotColor: '#10b981',
      cardBg: '#ffffff'
    };

/************************************************************
 * FORMATADORES
 ************************************************************/
function formatarDataCurtaStrangle(d) {
  try {
    if (!d) return "";
    if (Object.prototype.toString.call(d) === "[object Date]") {
      return Utilities.formatDate(d, Session.getScriptTimeZone(), "dd/MM/yyyy");
    }
    var dt = new Date(d);
    if (!isNaN(dt.getTime())) {
      return Utilities.formatDate(dt, Session.getScriptTimeZone(), "dd/MM/yyyy");
    }
    var iso = String(d).substring(0, 10);
    var p = iso.split("-");
    if (p.length === 3 && p[0].length === 4)
      return p[2] + "/" + p[1] + "/" + p[0];
    return d;
  } catch (e) {
    return "";
  }
}

function formatarMoedaBR(v) {
  if (isNaN(v)) return "-";
  return "R$ " + Number(v).toFixed(2).replace(".", ",");
}

function formatarPercentualBR(v) {
  if (isNaN(v)) return "-";
  return (v * 100).toFixed(1).replace(".", ",") + "%";
}

/************************************************************
 * BADGE
 ************************************************************/
function criarBadgeStrangle_(texto, tipo) {
  tipo = tipo || 'neutral';
  var colors = {
    success: { bg: '#d1fae5', text: '#065f46' },
    danger:  { bg: '#fee2e2', text: '#991b1b' },
    warning: { bg: '#fef3c7', text: '#92400e' },
    info:    { bg: '#dbeafe', text: '#1e40af' },
    neutral: { bg: '#f3f4f6', text: '#374151' }
  };
  var c = colors[tipo] || colors.neutral;

  return ''
    + '<span style="display:inline-block;padding:2px 8px;border-radius:10px;'
    + 'font-size:10px;font-weight:600;'
    + 'background:' + c.bg + ';color:' + c.text + ';white-space:nowrap;">'
    + texto
    + '</span>';
}

/************************************************************
 * SPARKLINE — estilo alinhado ao Template_Email_Operacoes
 * Assinatura mantida: (min, max, put, spot, call)
 ************************************************************/
function gerarSparklineUltraPremium(min, max, put, spot, call) {
  // Normalização de min/max para não quebrar
  var valores = [];
  [put, spot, call].forEach(function(v) {
    if (typeof v === 'number' && isFinite(v)) valores.push(v);
  });

  if (!isFinite(min) || !isFinite(max) || min === max) {
    if (valores.length) {
      min = Math.min.apply(null, valores);
      max = Math.max.apply(null, valores);
    } else {
      min = 0;
      max = 1;
    }
  }
  if (max === min) max = min + 1;

  var N = 16;
  function scale(v) { return (v - min) / (max - min); }
  function idx(v) {
    if (!isFinite(v)) return -1;
    return Math.max(0, Math.min(N - 1, Math.floor(scale(v) * (N - 1))));
  }

  var idxPut  = idx(put);
  var idxSpot = idx(spot);
  var idxCall = idx(call);

  var html = '';
  html += '<div style="margin:8px 0;padding:8px;background:' + STRANGLE_COLORS.gray50 + ';border-radius:6px;border:1px solid ' + STRANGLE_COLORS.gray200 + ';">';
  html += '<table cellpadding="0" cellspacing="0" style="border-collapse:collapse;width:100%;"><tr>';

  for (var i = 0; i < N; i++) {
    var bg = STRANGLE_COLORS.gray200;
    var height = 4;

    if (i === idxSpot) {
      bg = STRANGLE_COLORS.spotColor;
      height = 6;
    } else if (i === idxPut) {
      bg = STRANGLE_COLORS.putColor;
      height = 5;
    } else if (i === idxCall) {
      bg = STRANGLE_COLORS.callColor;
      height = 5;
    }

    html += '<td style="padding:0 1px;"><div style="height:'+height+'px;background:'+bg+';border-radius:2px;"></div></td>';
  }

  html += '</tr></table>';

  // Legenda harmonizada com DESIGN_COLORS
  html += '<div style="text-align:center;font-size:9px;color:' + STRANGLE_COLORS.gray500 + ';margin-top:4px;">';
  html += '<span style="display:inline-block;width:10px;height:4px;background:'+STRANGLE_COLORS.putColor+';border-radius:2px;margin-right:3px;vertical-align:middle;"></span>PUT ';
  html += '<span style="display:inline-block;width:10px;height:5px;background:'+STRANGLE_COLORS.spotColor+';border-radius:2px;margin:0 3px;vertical-align:middle;"></span>SPOT ';
  html += '<span style="display:inline-block;width:10px;height:4px;background:'+STRANGLE_COLORS.callColor+';border-radius:2px;margin-left:3px;vertical-align:middle;"></span>CALL';
  html += '</div></div>';

  return html;
}

/************************************************************
 * TEMPLATE HTML — harmonizado com Template_Email_Operacoes
 ************************************************************/
function montarHtmlAlertaStrangles(strangles, diasAntes, modoTeste) {

  // Dourado passa a usar a cor "warning" da paleta
  var dourado = STRANGLE_COLORS.warning;

  if (!strangles || !strangles.length)
    return "<p>Sem dados.</p>";

  var primeiro = strangles[0];
  var venc = formatarDataCurtaStrangle(primeiro.dueDate);
  var dte  = primeiro.dte;

  // Agrupar por ticker
  var grupos = {};
  strangles.forEach(function(s) {
    if (!grupos[s.ticker]) {
      grupos[s.ticker] = {
        lista: [],
        spot: s.spot,
        total: 0
      };
    }
    grupos[s.ticker].lista.push(s);
    grupos[s.ticker].total += (s.totalCredit || 0);
  });

  // Ordenar tickers pelo crédito médio (mais "interessantes" primeiro)
  var tickers = Object.keys(grupos).sort(function(a, b) {
    var ga = grupos[a], gb = grupos[b];
    var ma = ga.total / Math.max(1, ga.lista.length);
    var mb = gb.total / Math.max(1, gb.lista.length);
    return mb - ma;
  });

  var titulo = modoTeste
    ? "📈 Short Strangles (TESTE)"
    : "📈 Short Strangles";

  var html = '';
  html += '<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0"></head>';
  html += '<body style="margin:0;padding:0;font-family:-apple-system,BlinkMacSystemFont,\'Segoe UI\',Roboto,Arial,sans-serif;background:'+STRANGLE_COLORS.gray100+';">';
  html += '<div style="max-width:640px;margin:0 auto;background:'+STRANGLE_COLORS.cardBg+';">';

  // Padding principal
  html += '<div style="padding:16px;">';

  /***********************
   * HEADER / RESUMO
   ***********************/
  html += '<div style="margin-bottom:16px;">';
  html +=   '<div style="font-size:13px;font-weight:700;color:'+STRANGLE_COLORS.gray900+';margin-bottom:8px;padding-bottom:4px;border-bottom:2px solid '+STRANGLE_COLORS.gray200+';">'
         +   titulo
         + '</div>';

  html +=   '<div style="background:'+STRANGLE_COLORS.cardBg+';border-radius:6px;padding:10px;border:1px solid '+STRANGLE_COLORS.gray200+';border-left:3px solid '+STRANGLE_COLORS.primary+';">';
  html +=     '<table cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;font-size:11px;">';
  html +=       '<tr>';

  html +=         '<td style="padding:4px;vertical-align:top;">'
         +           '<div><strong>Strangles:</strong> ' + strangles.length + '</div>'
         +           '<div><strong>Vencimento:</strong> ' + venc
         +             ' <span style="color:'+STRANGLE_COLORS.gray500+';">(DTE: ' + dte + 'd)</span></div>'
         +         '</td>';

  html +=         '<td style="padding:4px;vertical-align:top;text-align:right;">'
         +           '<div style="margin-bottom:4px;font-size:10px;color:'+STRANGLE_COLORS.gray500+';">'
         +             'Gerado a D-' + diasAntes
         +           '</div>';

  if (modoTeste) {
    html +=         criarBadgeStrangle_('MODO TESTE', 'warning');
  }

  html +=         '</td>';
  html +=       '</tr>';
  html +=     '</table>';
  html +=   '</div>'; // card resumo
  html += '</div>';   // bloco header

  /***********************
   * BLOCOS POR TICKER
   ***********************/
  tickers.forEach(function(ticker) {
    var grupo = grupos[ticker];
    var spot  = formatarMoedaBR(grupo.spot);

    // Min/max para o sparkline
    var min = null, max = null;
    grupo.lista.forEach(function(s) {
      var localMin = Math.min(s.putStrike, s.callStrike, s.spot);
      var localMax = Math.max(s.putStrike, s.callStrike, s.spot);
      if (min === null || localMin < min) min = localMin;
      if (max === null || localMax > max) max = localMax;
    });
    if (!isFinite(min) || !isFinite(max) || max === min) {
      min = min || 0;
      max = min + 1;
    }

    // Usa a primeira estrutura como referência do sparkline
    var s0 = grupo.lista[0];

    // Ordena as estruturas deste ticker por crédito total
    grupo.lista.sort(function(a, b) {
      return (b.totalCredit || 0) - (a.totalCredit || 0);
    });

    // Card do ticker
    html += '<div style="background:'+STRANGLE_COLORS.gray50+';border-radius:8px;padding:10px;margin-bottom:12px;border:1px solid '+STRANGLE_COLORS.gray200+';border-left:3px solid '+STRANGLE_COLORS.info+';">';

    // Header do ticker
    html +=   '<div style="margin-bottom:6px;">'
         +      '<table cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;">'
         +        '<tr>';

    html +=         '<td style="vertical-align:middle;">'
         +            '<span style="font-size:16px;font-weight:800;color:'+STRANGLE_COLORS.gray900+';margin-right:6px;">'
         +              ticker
         +            '</span>'
         +            criarBadgeStrangle_(spot, 'info')
         +         '</td>';

    html +=         '<td style="vertical-align:middle;text-align:right;font-size:10px;color:'+STRANGLE_COLORS.gray500+';">'
         +            grupo.lista.length + ' strangle' + (grupo.lista.length > 1 ? 's' : '')
         +         '</td>';

    html +=        '</tr>'
         +      '</table>'
         +    '</div>';

    // Sparkline PUT / SPOT / CALL
    html += gerarSparklineUltraPremium(min, max, s0.putStrike, s0.spot, s0.callStrike);

    // Cartõezinhos de cada strangle
    for (var i = 0; i < grupo.lista.length; i++) {
      var s = grupo.lista[i];
      var rowBg = (i % 2 === 0) ? STRANGLE_COLORS.cardBg : STRANGLE_COLORS.gray50;

      html += '<div style="padding:8px 10px;margin-top:8px;border-radius:6px;background:'+rowBg+';border:1px solid '+STRANGLE_COLORS.gray200+';">';

      html +=   '<table cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;font-size:11px;">'
           +      '<tr>';

      // Coluna PUT
      html +=       '<td style="vertical-align:top;padding-right:6px;border-right:1px dashed '+STRANGLE_COLORS.gray200+';">'
           +          '<div style="margin-bottom:2px;font-weight:700;color:'+STRANGLE_COLORS.putColor+';">PUT</div>'
           +          '<div style="font-size:10px;color:'+STRANGLE_COLORS.gray700+';">'
           +            s.putSymbol + ' · <strong>' + formatarMoedaBR(s.putStrike) + '</strong>'
           +          '</div>'
           +          '<div style="font-size:9px;color:'+STRANGLE_COLORS.gray500+';margin-top:2px;">'
           +            'Prêmio: <strong>' + formatarMoedaBR(s.putClose) + '</strong>'
           +            ' · Dist.: <strong>' + formatarPercentualBR(s.putDist) + '</strong>'
           +          '</div>'
           +        '</td>';

      // Coluna CALL
      html +=       '<td style="vertical-align:top;padding-left:6px;">'
           +          '<div style="margin-bottom:2px;font-weight:700;color:'+STRANGLE_COLORS.callColor+';">CALL</div>'
           +          '<div style="font-size:10px;color:'+STRANGLE_COLORS.gray700+';">'
           +            s.callSymbol + ' · <strong>' + formatarMoedaBR(s.callStrike) + '</strong>'
           +          '</div>'
           +          '<div style="font-size:9px;color:'+STRANGLE_COLORS.gray500+';margin-top:2px;">'
           +            'Prêmio: <strong>' + formatarMoedaBR(s.callClose) + '</strong>'
           +            ' · Dist.: <strong>' + formatarPercentualBR(s.callDist) + '</strong>'
           +          '</div>'
           +        '</td>';

      html +=      '</tr>'
           +    '</table>';

      html +=    '<div style="margin-top:6px;font-size:11px;font-weight:700;color:'+dourado+';text-align:right;">'
           +      'Crédito total: ' + formatarMoedaBR(s.totalCredit)
           +    '</div>';

      html +=  '</div>'; // cartão strangle
    }

    html += '</div>'; // bloco ticker
  });

  /***********************
   * DISCLAIMER / FOOTER
   ***********************/
  html += '<div style="margin-top:16px;font-size:10px;color:'+STRANGLE_COLORS.gray500+';text-align:left;">'
       +   '⚠️ Este alerta é gerado automaticamente e não constitui recomendação de investimento.'
       + '</div>';

  html += '<div style="margin-top:8px;font-size:9px;color:'+STRANGLE_COLORS.gray400+';text-align:center;">'
       +   'Nexo Opções © 2025'
       + '</div>';

  html += '</div>';  // padding
  html += '</div>';  // wrapper
  html += '</body></html>';

  return html;
}