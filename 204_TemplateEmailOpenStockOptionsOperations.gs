/* ============================================================
 *  TEMPLATE_EMAIL_OPERACOES.GS — MODERN COMPACT EDITION
 *  Versão: 3.5 — Ajustes finais
 * ============================================================
 */

var DESIGN_COLORS = {
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

function fmtDataOp_(dt) {
  if (!dt) return "";
  try {
    if (Object.prototype.toString.call(dt) === "[object Date]") {
      if (isNaN(dt.getTime())) return "";
      return Utilities.formatDate(dt, Session.getScriptTimeZone(), "dd/MM/yyyy");
    }
    var d2 = new Date(dt);
    if (!isNaN(d2.getTime()))
      return Utilities.formatDate(d2, Session.getScriptTimeZone(), "dd/MM/yyyy");
    return String(dt);
  } catch (_) { return ""; }
}

function fmtMoedaOp_(v) {
  if (isNaN(v)) return "-";
  var num = Number(v).toFixed(2);
  var parts = num.split(".");
  parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ".");
  return "R$ " + parts.join(",");
}

function fmtPctOp_(v) {
  if (isNaN(v)) return "-";
  return Number(v).toFixed(0) + "%";
}

function criarBadge_(texto, tipo) {
  tipo = tipo || 'neutral';
  var colors = {
    success: { bg: '#d1fae5', text: '#065f46' },
    danger: { bg: '#fee2e2', text: '#991b1b' },
    warning: { bg: '#fef3c7', text: '#92400e' },
    info: { bg: '#dbeafe', text: '#1e40af' },
    neutral: { bg: '#f3f4f6', text: '#374151' }
  };
  var c = colors[tipo] || colors.neutral;
  
  return '<span style="display:inline-block;padding:2px 8px;border-radius:10px;font-size:10px;font-weight:600;background:' + c.bg + ';color:' + c.text + ';white-space:nowrap;">' + texto + '</span>';
}

function criarProgressBar_(porcentagem, cor) {
  cor = cor || DESIGN_COLORS.primary;
  porcentagem = Math.max(0, Math.min(100, porcentagem));
  
  return '<div style="width:100%;height:4px;background:' + DESIGN_COLORS.gray200 + ';border-radius:2px;overflow:hidden;"><div style="width:' + porcentagem + '%;height:100%;background:' + cor + ';"></div></div>';
}

function calcularNocional_(strike, quantidade) {
  return (strike || 0) * (quantidade || 0);
}

function gerarSparklineOperacaoHeatmap_(est) {
  var spot = est.tickerSpot || 0;
  var strikes = est.pernas.map(function(p){ return p.strike || 0; });
  var todos = strikes.slice();
  if (spot) todos.push(spot);

  var min = Math.min.apply(null, todos);
  var max = Math.max.apply(null, todos);
  if (!isFinite(min) || !isFinite(max) || min === max) {
    min = min || 0;
    max = min + 1;
  }

  var N = 16;
  function scale(v) { return (v - min) / (max - min); }

  var posPUT = [], posCALL = [];
  est.pernas.forEach(function(p) {
    var idx = Math.floor(scale(p.strike || 0) * (N - 1));
    if (p.tipo === "PUT")  posPUT.push(idx);
    if (p.tipo === "CALL") posCALL.push(idx);
  });

  var idxSpot = Math.floor(scale(spot) * (N - 1));

  var html = '<div style="margin:8px 0;padding:8px;background:' + DESIGN_COLORS.gray50 + ';border-radius:6px;"><table cellpadding="0" cellspacing="0" style="border-collapse:collapse;width:100%;"><tr>';

  for (var i = 0; i < N; i++) {
    var bg = DESIGN_COLORS.gray200;
    var height = 4;

    if (i === idxSpot) {
      bg = DESIGN_COLORS.spotColor;
      height = 6;
    } else if (posPUT.indexOf(i) !== -1) {
      bg = DESIGN_COLORS.putColor;
      height = 5;
    } else if (posCALL.indexOf(i) !== -1) {
      bg = DESIGN_COLORS.callColor;
      height = 5;
    }

    html += '<td style="padding:0 1px;"><div style="height:'+height+'px;background:'+bg+';border-radius:2px;"></div></td>';
  }

  html += '</tr></table>';
  html += '<div style="text-align:center;font-size:9px;color:' + DESIGN_COLORS.gray500 + ';margin-top:4px;">';
  html += '<span style="display:inline-block;width:10px;height:4px;background:' + DESIGN_COLORS.putColor + ';border-radius:2px;margin-right:3px;vertical-align:middle;"></span>PUT ';
  html += '<span style="display:inline-block;width:10px;height:5px;background:' + DESIGN_COLORS.spotColor + ';border-radius:2px;margin:0 3px;vertical-align:middle;"></span>SPOT ';
  html += '<span style="display:inline-block;width:10px;height:4px;background:' + DESIGN_COLORS.callColor + ';border-radius:2px;margin-left:3px;vertical-align:middle;"></span>CALL';
  html += '</div></div>';

  return html;
}

function gerarSparklinePernaHeatmap_(p, spot) {
  var strike = p.strike || 0;
  if (!spot || !strike) return "";

  var N = 16;
  var mid = Math.floor(N / 2);
  var dist = strike - spot;
  var range = spot * 0.15;
  var step = range / (mid - 1);
  var rawOffset = dist / step;
  var idxStrike = Math.round(mid + rawOffset);

  if (p.tipo === "PUT") idxStrike = Math.min(idxStrike, mid - 1);
  else if (p.tipo === "CALL") idxStrike = Math.max(idxStrike, mid + 1);

  idxStrike = Math.max(0, Math.min(N - 1, idxStrike));

  var html = '<div style="width:100%;margin:4px 0;"><table cellpadding="0" cellspacing="0" style="border-collapse:collapse;width:100%;"><tr>';

  for (var i = 0; i < N; i++) {
    var height = "2px";
    var bg = DESIGN_COLORS.gray300;

    if (i === idxStrike) {
      bg = (p.tipo === "PUT") ? DESIGN_COLORS.putColor : DESIGN_COLORS.callColor;
      height = "3px";
    }
    if (i === mid) {
      bg = DESIGN_COLORS.spotColor;
      height = "3px";
    }

    html += '<td style="padding:0 0.5px;"><div style="height:'+height+';background:'+bg+';border-radius:1px;"></div></td>';
  }

  html += '</tr></table></div>';
  return html;
}

function gerarHtmlMoneynessVendas_(resumo) {
  var call = resumo.moneyness?.call || { ITM:0, ATM:0, OTM:0, total:0 };
  var put  = resumo.moneyness?.put  || { ITM:0, ATM:0, OTM:0, total:0 };

  return ''
    + '<table cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;margin-top:6px;">'
    + '<tr>'
    + '<td style="width:50%;padding:10px;background:' + DESIGN_COLORS.gray50 + ';border-radius:6px;border-left:3px solid ' + DESIGN_COLORS.callColor + ';vertical-align:top;">'
    + '<div style="font-size:11px;font-weight:700;color:' + DESIGN_COLORS.callColor + ';margin-bottom:6px;">CALL</div>'
    + '<div style="font-size:10px;color:' + DESIGN_COLORS.gray700 + ';margin-bottom:4px;">'
    + criarBadge_('ITM: ' + call.ITM, call.ITM > 0 ? 'danger' : 'neutral') + ' '
    + criarBadge_('ATM: ' + call.ATM, call.ATM > 0 ? 'warning' : 'neutral') + ' '
    + criarBadge_('OTM: ' + call.OTM, call.OTM > 0 ? 'success' : 'neutral')
    + '</div>'
    + '<div style="font-size:10px;color:' + DESIGN_COLORS.gray600 + ';">Total: <strong>' + call.total + '</strong></div>'
    + '</td>'
    + '<td style="width:8px;"></td>'
    + '<td style="width:50%;padding:10px;background:' + DESIGN_COLORS.gray50 + ';border-radius:6px;border-left:3px solid ' + DESIGN_COLORS.putColor + ';vertical-align:top;">'
    + '<div style="font-size:11px;font-weight:700;color:' + DESIGN_COLORS.putColor + ';margin-bottom:6px;">PUT</div>'
    + '<div style="font-size:10px;color:' + DESIGN_COLORS.gray700 + ';margin-bottom:4px;">'
    + criarBadge_('ITM: ' + put.ITM, put.ITM > 0 ? 'danger' : 'neutral') + ' '
    + criarBadge_('ATM: ' + put.ATM, put.ATM > 0 ? 'warning' : 'neutral') + ' '
    + criarBadge_('OTM: ' + put.OTM, put.OTM > 0 ? 'success' : 'neutral')
    + '</div>'
    + '<div style="font-size:10px;color:' + DESIGN_COLORS.gray600 + ';">Total: <strong>' + put.total + '</strong></div>'
    + '</td>'
    + '</tr>'
    + '</table>';
}

function gerarHtmlDistribuicaoVencimento_(resumo) {
  var curto = resumo.vencimentos?.curto  || { pct: 0 };
  var medio = resumo.vencimentos?.medio  || { pct: 0 };
  var longo = resumo.vencimentos?.longo  || { pct: 0 };

  var html = '<div style="margin-top:6px;">';
  
  html += '<div style="margin-bottom:8px;">';
  html += '<div style="display:table;width:100%;margin-bottom:3px;">';
  html += '<div style="display:table-cell;font-size:11px;font-weight:600;color:' + DESIGN_COLORS.gray700 + ';">Curto Prazo (0-15 dias)</div>';
  html += '<div style="display:table-cell;text-align:right;font-size:12px;font-weight:700;color:' + DESIGN_COLORS.danger + ';">' + fmtPctOp_(curto.pct) + '</div>';
  html += '</div>';
  html += criarProgressBar_(curto.pct, DESIGN_COLORS.danger);
  html += '</div>';
  
  html += '<div style="margin-bottom:8px;">';
  html += '<div style="display:table;width:100%;margin-bottom:3px;">';
  html += '<div style="display:table-cell;font-size:11px;font-weight:600;color:' + DESIGN_COLORS.gray700 + ';">Médio Prazo (16-45 dias)</div>';
  html += '<div style="display:table-cell;text-align:right;font-size:12px;font-weight:700;color:' + DESIGN_COLORS.warning + ';">' + fmtPctOp_(medio.pct) + '</div>';
  html += '</div>';
  html += criarProgressBar_(medio.pct, DESIGN_COLORS.warning);
  html += '</div>';
  
  html += '<div>';
  html += '<div style="display:table;width:100%;margin-bottom:3px;">';
  html += '<div style="display:table-cell;font-size:11px;font-weight:600;color:' + DESIGN_COLORS.gray700 + ';">Longo Prazo (46+ dias)</div>';
  html += '<div style="display:table-cell;text-align:right;font-size:12px;font-weight:700;color:' + DESIGN_COLORS.success + ';">' + fmtPctOp_(longo.pct) + '</div>';
  html += '</div>';
  html += criarProgressBar_(longo.pct, DESIGN_COLORS.success);
  html += '</div>';
  
  html += '</div>';
  return html;
}

function gerarHtmlConcentracaoAtivo_(resumo) {
  var conc = (resumo.concentracaoPorTicker || []).slice();
  if (!conc.length) return '<div style="font-size:11px;color:' + DESIGN_COLORS.gray500 + ';">Sem dados.</div>';

  conc.sort(function(a,b){ return (b.pct || 0) - (a.pct || 0); });

  var html = '<div style="margin-top:6px;">';

  conc.forEach(function(item) {
    var cor = DESIGN_COLORS.primary;
    if (item.pct >= 30) cor = DESIGN_COLORS.danger;
    else if (item.pct >= 20) cor = DESIGN_COLORS.warning;
    else if (item.pct >= 10) cor = DESIGN_COLORS.info;
    else cor = DESIGN_COLORS.success;
    
    html += '<div style="margin-bottom:8px;">';
    html += '<div style="display:table;width:100%;margin-bottom:3px;">';
    html += '<div style="display:table-cell;font-size:12px;font-weight:700;color:' + DESIGN_COLORS.gray900 + ';">' + item.ticker + '</div>';
    html += '<div style="display:table-cell;text-align:center;font-size:10px;color:' + DESIGN_COLORS.gray500 + ';">' + (item.count || 0) + ' estrutura' + (item.count > 1 ? 's' : '') + '</div>';
    html += '<div style="display:table-cell;text-align:right;font-size:12px;font-weight:700;color:' + cor + ';">' + fmtPctOp_(item.pct) + '</div>';
    html += '</div>';
    html += criarProgressBar_(item.pct, cor);
    html += '</div>';
  });

  html += '</div>';
  return html;
}

function gerarHtmlLucroPrejuizo_(resumo) {
  var lucro = resumo.resultados?.noLucro?.pct || 0;
  var prejuizo = resumo.resultados?.noPrejuizo?.pct || 0;
  var breakeven = 100 - lucro - prejuizo;

  return ''
    + '<div style="margin-top:6px;">'
    + '<table cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;margin-bottom:10px;">'
    + '<tr>'
    + '<td style="text-align:center;padding:8px 4px;">'
    + '<div style="font-size:9px;color:' + DESIGN_COLORS.gray500 + ';margin-bottom:2px;text-transform:uppercase;">Lucro</div>'
    + '<div style="font-size:18px;font-weight:700;color:' + DESIGN_COLORS.success + ';">' + fmtPctOp_(lucro) + '</div>'
    + '</td>'
    + '<td style="text-align:center;padding:8px 4px;">'
    + '<div style="font-size:9px;color:' + DESIGN_COLORS.gray500 + ';margin-bottom:2px;text-transform:uppercase;">Breakeven</div>'
    + '<div style="font-size:18px;font-weight:700;color:' + DESIGN_COLORS.gray600 + ';">' + fmtPctOp_(breakeven) + '</div>'
    + '</td>'
    + '<td style="text-align:center;padding:8px 4px;">'
    + '<div style="font-size:9px;color:' + DESIGN_COLORS.gray500 + ';margin-bottom:2px;text-transform:uppercase;">Prejuízo</div>'
    + '<div style="font-size:18px;font-weight:700;color:' + DESIGN_COLORS.danger + ';">' + fmtPctOp_(prejuizo) + '</div>'
    + '</td>'
    + '</tr>'
    + '</table>'
    + '<div style="display:table;width:100%;height:8px;border-radius:4px;overflow:hidden;">'
    + '<div style="display:table-cell;width:' + lucro + '%;background:' + DESIGN_COLORS.success + ';"></div>'
    + '<div style="display:table-cell;width:' + breakeven + '%;background:' + DESIGN_COLORS.gray400 + ';"></div>'
    + '<div style="display:table-cell;width:' + prejuizo + '%;background:' + DESIGN_COLORS.danger + ';"></div>'
    + '</div>'
    + '</div>';
}

function montarHtmlRelatorioOperacoesAbertas(report) {
  var r = report.resumo || {};
  var estruturas = report.estruturas || [];
  var dataRef = fmtDataOp_(report.dataReferencia);

  // CALCULAR NOCIONAL CALL E PUT SE NÃO EXISTIR
  var nocionalCallTotal = 0;
  var nocionalPutTotal = 0;
  
  estruturas.forEach(function(est) {
    var nocionalEst = 0;
    (est.pernas || []).forEach(function(perna) {
      var nocional = calcularNocional_(perna.strike, perna.quantidade);
      perna.nocional = nocional;
      nocionalEst += nocional;
      
      if (perna.tipo === "CALL") {
        nocionalCallTotal += nocional;
      } else if (perna.tipo === "PUT") {
        nocionalPutTotal += nocional;
      }
    });
    est.nocional = nocionalEst;
  });


  // ORDENAR ESTRUTURAS POR VENCIMENTO (mais próximo primeiro)
  estruturas.sort(function(a, b) {
    function normalizarDte(v) {
      // Se for número válido (inclusive 0), usa o próprio valor
      if (typeof v === "number" && !isNaN(v)) return v;
      // Se estiver vazio/undefined/null, joga lá pra frente
      return 999999;
    }

    var dteA = normalizarDte(a.dte);
    var dteB = normalizarDte(b.dte);

    return dteA - dteB;
  });


  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0"></head>';
  html += '<body style="margin:0;padding:0;font-family:-apple-system,BlinkMacSystemFont,\'Segoe UI\',Roboto,Arial,sans-serif;background:#f3f4f6;">';
  html += '<div style="max-width:640px;margin:0 auto;background:#ffffff;">';

  html += '<div style="padding:16px;">';

  // RESUMO EXECUTIVO COMPACTO
  var lucroTotal = r.lucroTotal || 0;
  var corLucro = lucroTotal >= 0 ? DESIGN_COLORS.success : DESIGN_COLORS.danger;
  var totalEst = r.totalEstruturas || 0;
  
  var vencProx = "";
  var diasProx = 999;
  var existeVencHoje = false;
  var qtdVencHoje = 0;

  estruturas.forEach(function(est) {
    var dte = est.dte;

    // Normaliza DTE
    if (dte === "" || dte === null || isNaN(dte)) dte = 999;
    if (dte < 0) dte = 0;

    // Marca estruturas que vencem hoje
    if (dte === 0) {
      existeVencHoje = true;
      qtdVencHoje++;
    }

    // Vencimento mais próximo
    if (dte < diasProx) {
      diasProx = dte;
      vencProx = fmtDataOp_(est.vencimento);
    }
  });


  var maiorExposicao = "";
  var maiorPct = 0;
  (r.concentracaoPorTicker || []).forEach(function(item) {
    if (item.pct > maiorPct) {
      maiorPct = item.pct;
      maiorExposicao = item.ticker + " (" + fmtPctOp_(item.pct) + ")";
    }
  });

  html += '<div style="margin-bottom:16px;">';
  html += '<div style="font-size:13px;font-weight:700;color:' + DESIGN_COLORS.gray900 + ';margin-bottom:8px;padding-bottom:4px;border-bottom:2px solid ' + DESIGN_COLORS.gray200 + ';">💼 Resumo · ' + dataRef + '</div>';
  html += '<div style="background:#ffffff;border-radius:6px;padding:10px;border:1px solid ' + DESIGN_COLORS.gray200 + ';border-left:3px solid ' + corLucro + ';">';
  var badgeVenceHoje = "";
  if (existeVencHoje) {
    var textoBadge = (qtdVencHoje > 1
      ? "VENCEM HOJE · " + qtdVencHoje + " estruturas"
      : "VENCE HOJE");

    badgeVenceHoje = " " + criarBadge_(textoBadge, "warning");
  }
  html += '<table cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;font-size:11px;">';
  html += '<tr>';
  html += '<td style="padding:4px;"><strong>Resultado:</strong> <span style="color:' + corLucro + ';font-weight:700;">' + fmtMoedaOp_(lucroTotal) + '</span> <span style="color:' + DESIGN_COLORS.gray500 + ';">(' + totalEst + ' estruturas)</span></td>';
  html += '<td style="padding:4px;text-align:right;"><strong>Exposição:</strong> <span style="color:' + DESIGN_COLORS.warning + ';font-weight:700;">' + maiorExposicao + '</span></td>';
  html += '</tr>';
  html += '<tr>';
  html += '<td style="padding:4px;"><strong>Vencimento:</strong> '
     + vencProx
     + ' <span style="color:' + DESIGN_COLORS.gray500 + ';">(' + diasProx + 'd)</span>'
     + badgeVenceHoje
     + '</td>';
  html += '<td style="padding:4px;text-align:right;"><strong>💵 CALL:</strong> ' + fmtMoedaOp_(nocionalCallTotal) + '</td>';
  html += '</tr>';
  html += '<tr>';
  html += '<td colspan="2" style="padding:4px;text-align:right;"><strong>💵 PUT:</strong> ' + fmtMoedaOp_(nocionalPutTotal) + '</td>';
  html += '</tr>';
  html += '</table>';
  html += '</div>';
  html += '</div>';

  // MONEYNESS
  html += '<div style="margin-bottom:16px;">';
  html += '<div style="font-size:13px;font-weight:700;color:' + DESIGN_COLORS.gray900 + ';margin-bottom:8px;padding-bottom:4px;border-bottom:2px solid ' + DESIGN_COLORS.gray200 + ';">🎯 Moneyness</div>';
  html += gerarHtmlMoneynessVendas_(r);
  html += '</div>';

  // DISTRIBUIÇÃO POR VENCIMENTO
  html += '<div style="margin-bottom:16px;">';
  html += '<div style="font-size:13px;font-weight:700;color:' + DESIGN_COLORS.gray900 + ';margin-bottom:8px;padding-bottom:4px;border-bottom:2px solid ' + DESIGN_COLORS.gray200 + ';">📅 Vencimentos</div>';
  html += '<div style="background:#ffffff;border-radius:6px;padding:10px;border:1px solid ' + DESIGN_COLORS.gray200 + ';">';
  html += gerarHtmlDistribuicaoVencimento_(r);
  html += '</div>';
  html += '</div>';

  // CONCENTRAÇÃO POR ATIVO
  html += '<div style="margin-bottom:16px;">';
  html += '<div style="font-size:13px;font-weight:700;color:' + DESIGN_COLORS.gray900 + ';margin-bottom:8px;padding-bottom:4px;border-bottom:2px solid ' + DESIGN_COLORS.gray200 + ';">🏦 Concentração</div>';
  html += '<div style="background:#ffffff;border-radius:6px;padding:10px;border:1px solid ' + DESIGN_COLORS.gray200 + ';">';
  html += gerarHtmlConcentracaoAtivo_(r);
  html += '</div>';
  html += '</div>';

  // LUCRO VS PREJUÍZO
  html += '<div style="margin-bottom:16px;">';
  html += '<div style="font-size:13px;font-weight:700;color:' + DESIGN_COLORS.gray900 + ';margin-bottom:8px;padding-bottom:4px;border-bottom:2px solid ' + DESIGN_COLORS.gray200 + ';">💰 Resultado</div>';
  html += '<div style="background:#ffffff;border-radius:6px;padding:10px;border:1px solid ' + DESIGN_COLORS.gray200 + ';">';
  html += gerarHtmlLucroPrejuizo_(r);
  html += '</div>';
  html += '</div>';

  // ESTRUTURAS EM DETALHE
  html += '<div style="margin-bottom:16px;">';
  html += '<div style="font-size:13px;font-weight:700;color:' + DESIGN_COLORS.gray900 + ';margin-bottom:8px;padding-bottom:4px;border-bottom:2px solid ' + DESIGN_COLORS.gray200 + ';">📂 Estruturas</div>';

  estruturas.forEach(function(est) {
    var spot = est.tickerSpot || 0;
    var venc = fmtDataOp_(est.vencimento);
    var dte = est.dte || 0;
    var premio = fmtMoedaOp_(est.premioTotal || 0);
    var custoZerar = fmtMoedaOp_(est.custoZeragem || 0);
    var pl = est.plTotal || 0;
    var plFmt = fmtMoedaOp_(pl);
    var nocional = est.nocional || 0;

    var corBorda = DESIGN_COLORS.gray400;
    var textoStatus = '⚪';
    
    if (pl > 0) {
      corBorda = DESIGN_COLORS.success;
      textoStatus = '🟢';
    } else if (pl < 0) {
      corBorda = DESIGN_COLORS.danger;
      textoStatus = '🔴';
    }

    // ORDENAR PERNAS: por strike, depois por código
    var pernasOrdenadas = (est.pernas || []).slice().sort(function(a, b) {
      var strikeA = a.strike || 0;
      var strikeB = b.strike || 0;
      if (strikeA !== strikeB) return strikeA - strikeB;
      var codigoA = (a.codigoOpcao || "").toUpperCase();
      var codigoB = (b.codigoOpcao || "").toUpperCase();
      if (codigoA < codigoB) return -1;
      if (codigoA > codigoB) return 1;
      return 0;
    });

    html += '<div style="background:#ffffff;border-radius:6px;padding:10px;margin-bottom:10px;border:1px solid ' + DESIGN_COLORS.gray200 + ';border-left:3px solid ' + corBorda + ';">';
    
    // Header compacto
    html += '<div style="margin-bottom:8px;">';
    html += '<table cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;">';
    html += '<tr>';
    html += '<td style="vertical-align:middle;">';
    html += '<span style="font-size:16px;font-weight:800;color:' + DESIGN_COLORS.gray900 + ';margin-right:6px;">' + est.ticker + '</span>';
    html += criarBadge_(fmtMoedaOp_(spot), 'info') + ' ';
    html += '<span style="font-size:14px;">' + textoStatus + '</span>';
    html += '</td>';
    html += '<td style="text-align:right;vertical-align:middle;">';
    html += '<span style="font-size:10px;color:' + DESIGN_COLORS.gray500 + ';">📅 ' + venc + ' (' + dte + 'd)</span>';
    html += '</td>';
    html += '</tr>';
    html += '</table>';
    html += '</div>';

    // Métricas inline (SEM DELTA)
    html += '<div style="font-size:10px;color:' + DESIGN_COLORS.gray700 + ';margin-bottom:8px;padding:6px;background:' + DESIGN_COLORS.gray50 + ';border-radius:4px;">';
    html += '<strong>PM:</strong> ' + premio + ' · ';
    html += '<strong>Zerar:</strong> ' + custoZerar + ' · ';
    html += '<strong>Lucro:</strong> <span style="color:' + (pl >= 0 ? DESIGN_COLORS.success : DESIGN_COLORS.danger) + ';font-weight:700;">' + plFmt + '</span> · ';
    html += '<strong>💵:</strong> ' + fmtMoedaOp_(nocional);
    html += '</div>';

    // Sparkline
    html += gerarSparklineOperacaoHeatmap_(est);

    // Pernas compactas (ordenadas)
    html += '<div style="margin-top:8px;">';
    html += '<div style="font-size:10px;font-weight:700;color:' + DESIGN_COLORS.gray700 + ';margin-bottom:6px;">Pernas (' + est.qtdPernas + ')</div>';

    pernasOrdenadas.forEach(function(perna, idxPerna) {
      var operacao = perna.operacao || "";
      var tipo = perna.tipo || "";
      var codigo = perna.codigoOpcao || "";
      var strike = fmtMoedaOp_(perna.strike || 0);
      var qtd = perna.quantidade || 0;
      var moneyness = perna.moneynessCode || "";
      var deltaPerna = (perna.delta || 0).toFixed(4);
      var pm = fmtMoedaOp_(perna.premioPM || 0);
      var atual = fmtMoedaOp_(perna.premioAtual || 0);
      var plPerna = perna.plAtual || 0;
      var plPernaFmt = fmtMoedaOp_(plPerna);
      var nocionalPerna = perna.nocional || 0;

      var corPerna = tipo === "CALL" ? DESIGN_COLORS.callColor : DESIGN_COLORS.putColor;
      var corMoneyness = 'neutral';
      if (moneyness === 'ITM') corMoneyness = 'danger';
      else if (moneyness === 'ATM') corMoneyness = 'warning';
      else if (moneyness === 'OTM') corMoneyness = 'success';

      var bgPerna = idxPerna % 2 === 0 ? '#ffffff' : DESIGN_COLORS.gray50;

      html += '<div style="padding:8px;background:' + bgPerna + ';border-radius:4px;margin-bottom:4px;border-left:2px solid ' + corPerna + ';">';
      
      html += '<div style="margin-bottom:4px;font-size:10px;">';
      html += '<span style="font-weight:700;color:' + DESIGN_COLORS.gray900 + ';margin-right:4px;">' + strike + '</span>';
      html += '<span style="font-weight:700;color:' + corPerna + ';margin-right:4px;">' + operacao.charAt(0) + ' ' + tipo + '</span>';
      html += criarBadge_(codigo, 'neutral') + ' ';
      html += criarBadge_(moneyness, corMoneyness) + ' ';
      html += '<span style="float:right;font-weight:700;color:' + (plPerna >= 0 ? DESIGN_COLORS.success : DESIGN_COLORS.danger) + ';">' + plPernaFmt + '</span>';
      html += '</div>';
      
      html += '<div style="font-size:9px;color:' + DESIGN_COLORS.gray600 + ';">';
      html += 'Qtd: <strong>' + qtd + '</strong> · ';
      html += 'Δ: <strong>' + deltaPerna + '</strong> · ';
      html += 'PM: <strong>' + pm + '</strong> · ';
      html += '↻: <strong>' + atual + '</strong> · ';
      html += '💵: <strong>' + fmtMoedaOp_(nocionalPerna) + '</strong>';
      html += '</div>';
      
      html += gerarSparklinePernaHeatmap_(perna, spot);
      
      html += '</div>';
    });

    html += '</div>';
    html += '</div>';
  });

  html += '</div>';

  // GUIA DE ESTRATÉGIAS
  html += '<div style="margin-bottom:16px;">';
  html += '<div style="font-size:13px;font-weight:700;color:' + DESIGN_COLORS.gray900 + ';margin-bottom:8px;padding-bottom:4px;border-bottom:2px solid ' + DESIGN_COLORS.gray200 + ';">📘 Estratégias</div>';
  html += '<div style="background:#ffffff;border-radius:6px;padding:10px;border:1px solid ' + DESIGN_COLORS.gray200 + ';">';
  
  html += '<table cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;font-size:10px;line-height:1.5;">';
  html += '<tr>';
  html += '<td style="width:50%;padding-right:8px;vertical-align:top;">';
  
  html += '<div style="font-weight:700;color:' + DESIGN_COLORS.gray700 + ';margin-bottom:4px;font-size:9px;text-transform:uppercase;">Neutras / Theta</div>';
  html += '<div style="margin-bottom:6px;color:' + DESIGN_COLORS.gray700 + ';">• <strong>Short Strangle</strong> 🟡<br><span style="font-size:9px;color:' + DESIGN_COLORS.gray600 + ';">PUT OTM + CALL OTM</span></div>';
  html += '<div style="margin-bottom:6px;color:' + DESIGN_COLORS.gray700 + ';">• <strong>Iron Condor</strong> 🟡<br><span style="font-size:9px;color:' + DESIGN_COLORS.gray600 + ';">Strangle protegido</span></div>';
  
  html += '<div style="font-weight:700;color:' + DESIGN_COLORS.gray700 + ';margin:8px 0 4px;font-size:9px;text-transform:uppercase;">Direcionais</div>';
  html += '<div style="margin-bottom:6px;color:' + DESIGN_COLORS.gray700 + ';">• <strong>Vertical Call</strong> 🟡<br><span style="font-size:9px;color:' + DESIGN_COLORS.gray600 + ';">Alta com proteção</span></div>';
  html += '<div style="margin-bottom:6px;color:' + DESIGN_COLORS.gray700 + ';">• <strong>Vertical Put</strong> 🟡<br><span style="font-size:9px;color:' + DESIGN_COLORS.gray600 + ';">Baixa com proteção</span></div>';
  
  html += '</td>';
  html += '<td style="width:50%;padding-left:8px;vertical-align:top;">';
  
  html += '<div style="font-weight:700;color:' + DESIGN_COLORS.gray700 + ';margin-bottom:4px;font-size:9px;text-transform:uppercase;">Renda & Entrada</div>';
  html += '<div style="margin-bottom:6px;color:' + DESIGN_COLORS.gray700 + ';">• <strong>Covered Call</strong> 🟢<br><span style="font-size:9px;color:' + DESIGN_COLORS.gray600 + ';">CALL com lastro</span></div>';
  html += '<div style="margin-bottom:6px;color:' + DESIGN_COLORS.gray700 + ';">• <strong>Naked Put</strong> 🟡<br><span style="font-size:9px;color:' + DESIGN_COLORS.gray600 + ';">Entrada com desconto</span></div>';
  
  html += '<div style="font-weight:700;color:' + DESIGN_COLORS.gray700 + ';margin:8px 0 4px;font-size:9px;text-transform:uppercase;">Multi-Leg</div>';
  html += '<div style="color:' + DESIGN_COLORS.gray700 + ';">• <strong>Custom</strong><br><span style="font-size:9px;color:' + DESIGN_COLORS.gray600 + ';">3+ pernas</span></div>';
  
  html += '</td>';
  html += '</tr>';
  html += '</table>';
  
  html += '</div>';
  html += '</div>';

  html += '</div>'; // Fecha padding

  // FOOTER
  html += '<div style="background:' + DESIGN_COLORS.gray100 + ';padding:12px 16px;text-align:center;border-top:1px solid ' + DESIGN_COLORS.gray200 + ';">';
  html += '<div style="font-size:9px;color:' + DESIGN_COLORS.gray500 + ';">Nexo Opções © 2025 · v3.5</div>';
  html += '</div>';

  html += '</div></body></html>';

  return html;
}

function getReportForPreview_() {
  return mockReportForPreview_();
}

function mockReportForPreview_() {
  var hoje = new Date();

  return {
    dataReferencia: hoje,
    resumo: {
      totalEstruturas: 11,
      lucroTotal: -692.00,
      deltaTotal: -3080.4246,
      moneyness: {
        call: { ITM: 0, ATM: 0, OTM: 8, total: 8 },
        put:  { ITM: 4, ATM: 2, OTM: 7, total: 13 }
      },
      vencimentos: {
        curto: { pct: 45 },
        medio: { pct: 36 },
        longo: { pct: 19  }
      },
      concentracaoPorTicker: [
        { ticker: "CSAN3", pct: 36, count: 4 },
        { ticker: "NATU3", pct: 18, count: 2 },
        { ticker: "BBDC4", pct: 9,  count: 1 },
        { ticker: "USIM5", pct: 9,  count: 1 },
        { ticker: "CSNA3", pct: 9,  count: 1 },
        { ticker: "PSSA3", pct: 9,  count: 1 },
        { ticker: "BRAV3", pct: 9,  count: 1 }
      ],
      resultados: {
        noLucro:    { pct: 18 },
        noPrejuizo: { pct: 82 }
      }
    },
    estruturas: [
      {
        ticker: "BBDC4",
        tickerSpot: 19.09,
        vencimento: new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate() + 2),
        dte: 2,
        premioTotal: 83.00,
        custoZeragem: -50.00,
        plTotal: 33.00,
        deltaTotal: -0.1487,
        tipoEstrutura: "multi_leg",
        qtdPernas: 3,
        pernas: [
          {
            operacao: "COMPRA",
            tipo: "PUT",
            codigoOpcao: "BBDCW164",
            strike: 16.07,
            quantidade: 100,
            moneynessCode: "OTM",
            delta: -0.0000,
            premioPM: 0.01,
            premioAtual: 0.01,
            plAtual: 0.00
          },
          {
            operacao: "VENDA",
            tipo: "PUT",
            codigoOpcao: "BBDCW164",
            strike: 16.07,
            quantidade: 100,
            moneynessCode: "OTM",
            delta: -0.0000,
            premioPM: 0.24,
            premioAtual: 0.01,
            plAtual: 23.00
          },
          {
            operacao: "VENDA",
            tipo: "PUT",
            codigoOpcao: "BBDCW191",
            strike: 18.57,
            quantidade: 1000,
            moneynessCode: "OTM",
            delta: -0.1636,
            premioPM: 0.06,
            premioAtual: 0.05,
            plAtual: 10.00
          }
        ]
      },
      {
        ticker: "USIM5",
        tickerSpot: 5.20,
        vencimento: new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate() + 2),
        dte: 2,
        premioTotal: 36.00,
        custoZeragem: -45.00,
        plTotal: -45.00,
        deltaTotal: -0.4334,
        tipoEstrutura: "multi_leg",
        qtdPernas: 4,
        pernas: [
          {
            operacao: "VENDA",
            tipo: "CALL",
            codigoOpcao: "USIMK560",
            strike: 5.60,
            quantidade: 100,
            moneynessCode: "OTM",
            delta: 0.0833,
            premioPM: 0.19,
            premioAtual: 0.01,
            plAtual: 18.00
          },
          {
            operacao: "COMPRA",
            tipo: "PUT",
            codigoOpcao: "USIMW560",
            strike: 5.60,
            quantidade: 100,
            moneynessCode: "ITM",
            delta: -0.7618,
            premioPM: 0.28,
            premioAtual: 0.46,
            plAtual: -18.00
          },
          {
            operacao: "VENDA",
            tipo: "PUT",
            codigoOpcao: "USIMW520",
            strike: 5.20,
            quantidade: 400,
            moneynessCode: "ATM",
            delta: -0.4804,
            premioPM: 0.06,
            premioAtual: 0.11,
            plAtual: -20.00
          },
          {
            operacao: "VENDA",
            tipo: "PUT",
            codigoOpcao: "USIMW560",
            strike: 5.60,
            quantidade: 100,
            moneynessCode: "ITM",
            delta: -0.7618,
            premioPM: 0.21,
            premioAtual: 0.46,
            plAtual: -25.00
          }
        ]
      },
      {
        ticker: "NATU3",
        tickerSpot: 7.75,
        vencimento: new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate() + 9),
        dte: 9,
        premioTotal: 140.00,
        custoZeragem: -290.00,
        plTotal: -150.00,
        deltaTotal: -1.0000,
        tipoEstrutura: "naked_put",
        qtdPernas: 1,
        pernas: [
          {
            operacao: "VENDA",
            tipo: "PUT",
            codigoOpcao: "NATUW810",
            strike: 8.10,
            quantidade: 1000,
            moneynessCode: "ITM",
            delta: -1.0000,
            premioPM: 0.14,
            premioAtual: 0.29,
            plAtual: -150.00
          }
        ]
      }
    ]
  };
}

function previewRelatorioOperacoesAbertas() {
  var report = getReportForPreview_();
  var html = montarHtmlRelatorioOperacoesAbertas(report);

  Logger.log(html);

  var output = HtmlService.createHtmlOutput(html)
    .setTitle("Preview — Relatório v3.5")
    .setWidth(1200)
    .setHeight(900);

  SpreadsheetApp.getUi().showModalDialog(output, "Preview — Relatório de Operações Abertas");
}