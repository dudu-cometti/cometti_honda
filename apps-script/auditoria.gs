/**
 * Auditoria da planilha do Meta Lead Ads — onde o lead se perde.
 *
 * Não usa a API: lê só a planilha. Serve para comparar três números e
 * descobrir em que ponto o lead desaparece:
 *
 *   leads no Meta  >  linhas na planilha  >  leads no TEIA
 *
 * Se o Meta entregou e a planilha tem a linha, mas o TEIA não tem o lead,
 * o problema é o sync — não é falta de demanda.
 *
 * COMO USAR
 *   Cole num arquivo novo do mesmo projeto Apps Script e rode
 *   auditarPlanilha(). Resultado em Ver > Registros de execução.
 */

/** Nome exato da aba. Vazio = primeira aba. */
var AUDITORIA_ABA = '';

/** Quantos meses mostrar, contando o corrente. */
var AUDITORIA_MESES = 3;

function auditarPlanilha() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var aba = AUDITORIA_ABA ? ss.getSheetByName(AUDITORIA_ABA) : ss.getSheets()[0];
  if (!aba) throw new Error('Aba "' + AUDITORIA_ABA + '" não encontrada.');

  var dados = aba.getDataRange().getValues();
  if (dados.length < 2) { Logger.log('Planilha vazia.'); return; }

  var cab = dados[0];
  var iCriado = acharColuna_(cab, ['created_time', 'criado']);
  var iLoja   = acharColuna_(cab, ['loja']);
  var iLeadId = acharColuna_(cab, ['lead_id']);
  var iStatus = acharColuna_(cab, ['status teia', 'status']);

  Logger.log('Colunas usadas: criado=' + rotulo_(cab, iCriado) +
             ' | loja=' + rotulo_(cab, iLoja) +
             ' | lead_id=' + rotulo_(cab, iLeadId) +
             ' | status=' + rotulo_(cab, iStatus));
  if (iCriado < 0 || iLeadId < 0) {
    Logger.log('ERRO: não achei a coluna de data de preenchimento ou a de lead_id. ' +
               'Cabeçalhos: ' + cab.join(' | '));
    return;
  }

  var meses = ultimosMeses_(AUDITORIA_MESES);
  var porMes = {};
  meses.forEach(function (m) {
    porMes[m] = { total: 0, semLead: 0, loja: {}, situacao: {}, erros: {}, lojaSemLead: {} };
  });

  for (var i = 1; i < dados.length; i++) {
    var mes = mesDe_(dados[i][iCriado]);
    if (!porMes[mes]) continue;

    var b = porMes[mes];
    var loja = iLoja >= 0 ? normalizarLoja_(dados[i][iLoja]) : '(sem coluna)';
    var temLead = String(dados[i][iLeadId] || '').trim() !== '';
    var status = iStatus >= 0 ? String(dados[i][iStatus] || '').trim() : '';

    b.total++;
    b.loja[loja] = (b.loja[loja] || 0) + 1;

    var sit = classificar_(status, temLead);
    b.situacao[sit] = (b.situacao[sit] || 0) + 1;

    if (sit === 'ERRO') {
      var msg = status.slice(0, 70);
      b.erros[msg] = (b.erros[msg] || 0) + 1;
    }

    if (!temLead && sit !== 'teste') {
      b.semLead++;
      b.lojaSemLead[loja] = (b.lojaSemLead[loja] || 0) + 1;
    }
  }

  meses.forEach(function (m) {
    var b = porMes[m];
    Logger.log('');
    Logger.log('================ ' + m + ' ================');
    Logger.log('Linhas no formulário : ' + b.total);
    Logger.log('Por loja escolhida   : ' + emLinha_(b.loja));
    Logger.log('Situação no TEIA     : ' + emLinha_(b.situacao));
    Logger.log('SEM lead no TEIA     : ' + b.semLead +
               (b.total ? '  (' + (b.semLead / b.total * 100).toFixed(1) + '% do mês)' : ''));
    if (b.semLead) Logger.log('   por loja          : ' + emLinha_(b.lojaSemLead));
    if (Object.keys(b.erros).length) {
      Logger.log('Erros:');
      Object.keys(b.erros).sort(function (a, c) { return b.erros[c] - b.erros[a]; })
        .forEach(function (k) { Logger.log('   ' + b.erros[k] + 'x  ' + k); });
    }
  });

  Logger.log('');
  Logger.log('COMO LER: "SEM lead no TEIA" é lead que o Meta entregou, caiu na');
  Logger.log('planilha e NUNCA chegou no CRM. Ninguém ligou para essa pessoa.');
  Logger.log('Compare "Por loja escolhida" entre os meses: se São Gabriel caiu aí,');
  Logger.log('é mix do formulário; se não caiu, a perda é do sync.');
}

// ------------------------------------------------------------- helpers

function classificar_(status, temLead) {
  var s = status.toLowerCase();
  if (s.indexOf('linha de teste') === 0 || s.indexOf('teste') === 0) return 'teste';
  if (s.indexOf('erro') === 0) return 'ERRO';
  if (s.indexOf('já existia') === 0 || s.indexOf('ja existia') === 0) return 'duplicado';
  if (s.indexOf('simula') === 0) return 'simulação';
  if (temLead) return 'no TEIA';
  if (!status) return 'nunca processada';
  return 'outro: ' + status.slice(0, 40);
}

/** Acha a coluna pelo pedaço do nome — o texto da pergunta do Meta muda. */
function acharColuna_(cab, pedacos) {
  for (var p = 0; p < pedacos.length; p++) {
    for (var i = 0; i < cab.length; i++) {
      if (String(cab[i] || '').toLowerCase().indexOf(pedacos[p]) !== -1) return i;
    }
  }
  return -1;
}

function rotulo_(cab, i) { return i >= 0 ? '"' + cab[i] + '"' : 'NÃO ACHEI'; }

function mesDe_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'America/Sao_Paulo', 'yyyy-MM');
  var s = String(v || '').trim();
  if (/^\d{4}-\d{2}/.test(s)) return s.slice(0, 7);
  var d = new Date(s);
  return isNaN(d.getTime()) ? '?' : Utilities.formatDate(d, 'America/Sao_Paulo', 'yyyy-MM');
}

function ultimosMeses_(n) {
  var out = [];
  var ref = new Date();
  for (var i = 0; i < n; i++) {
    out.push(Utilities.formatDate(ref, 'America/Sao_Paulo', 'yyyy-MM'));
    ref.setMonth(ref.getMonth() - 1, 1);
  }
  return out;
}

function normalizarLoja_(v) {
  var s = String(v || '').toLowerCase()
    .replace(/[áàãâä]/g, 'a').replace(/[éêë]/g, 'e').replace(/[íï]/g, 'i')
    .replace(/[óõôö]/g, 'o').replace(/[úü]/g, 'u').replace(/ç/g, 'c');
  if (!s.trim()) return '(vazio)';
  if (s.indexOf('colatina') !== -1) return 'COLATINA';
  if (s.indexOf('mantena') !== -1) return 'MANTENA';
  if (s.indexOf('gabriel') !== -1) return 'SAO GABRIEL';
  if (s.indexOf('afonso') !== -1) return 'AFONSO CLAUDIO';
  return 'NAO MAPEADA: ' + String(v).slice(0, 30);
}

function emLinha_(c) {
  var ks = Object.keys(c);
  if (!ks.length) return '-';
  return ks.sort(function (a, b) { return c[b] - c[a] || (a < b ? -1 : 1); })
           .map(function (k) { return k + '=' + c[k]; }).join('  ');
}
