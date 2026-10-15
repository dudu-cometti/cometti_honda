/**
 * Consultas pontuais no TEIA.
 *
 * COMO USAR
 *   1. No MESMO projeto Apps Script da planilha (Extensões > Apps Script),
 *      crie um arquivo novo e cole isto. Ele reaproveita o TEIA_TOKEN das
 *      Propriedades do script — não precisa configurar nada.
 *   2. Escolha a função no seletor do topo e clique em Executar.
 *   3. Leia o resultado em Ver > Registros de execução.
 *
 * Só lê. Nenhuma destas funções escreve no TEIA nem na planilha.
 */

var CONSULTA_URL_LISTA = 'https://teiacrm.com.br/api/newapi/leads/1161';

/**
 * Janela de CRIAÇÃO a varrer. A API só filtra por data de criação, então
 * tudo que for contado aqui é "lead que ENTROU neste período".
 */
var CONSULTA_DE  = '2026-01-01';
var CONSULTA_ATE = '2026-10-15';

var CONSULTA_MAX_PAGINAS = 60;

// ------------------------------------------------------------- a pergunta

/** Quantos clientes da Izabelly (Afonso Cláudio) com interesse em Biz. */
function contarBizIzabelly() {
  contarModeloPorVendedor_('biz', 'izabelly');
}

/** Mesmo recorte, outros exemplos — troque à vontade. */
function contarPopIzabelly()      { contarModeloPorVendedor_('pop', 'izabelly'); }
function contarBizRedeInteira()   { contarModeloPorVendedor_('biz', ''); }

// ------------------------------------------------------------- motor

function contarModeloPorVendedor_(modelo, vendedor) {
  var leads = puxarLeads_(CONSULTA_DE, CONSULTA_ATE);
  Logger.log('=== ' + modelo.toUpperCase() + (vendedor ? ' / ' + vendedor.toUpperCase() : ' / REDE') + ' ===');
  Logger.log(leads.length + ' lead(s) criados entre ' + CONSULTA_DE + ' e ' + CONSULTA_ATE + '.');

  if (!leads.length) return;

  // Os nomes dos campos da resposta não estão documentados. Isto mostra um
  // registro inteiro para conferir onde ficam produto e comentário antes de
  // confiar na conta.
  Logger.log('Exemplo de registro (confira os nomes dos campos):');
  Logger.log(JSON.stringify(leads[0]));

  var alvo = semAcentoConsulta_(vendedor);
  var daVendedora = vendedor
    ? leads.filter(function (l) {
        return semAcentoConsulta_(l.user_name || '').indexOf(alvo) !== -1;
      })
    : leads;

  if (vendedor && !daVendedora.length) {
    // Se o nome estiver escrito de outra forma no CRM, a conta daria zero e
    // pareceria falta de demanda. Então mostra quem existe de verdade.
    Logger.log('NENHUM lead com user_name contendo "' + vendedor + '".');
    Logger.log('Nomes que aparecem no período: ' + nomesDistintos_(leads).join(' | '));
    return;
  }

  // \b + modelo pega "Biz 125", "BIZ110i", "biz-125", e não pega palavra
  // que só contenha as letras no meio.
  var re = new RegExp('\\b' + modelo, 'i');

  var noProduto = [];
  var soNoComentario = [];

  daVendedora.forEach(function (l) {
    var produto = String(l.product_name || l.product || l.model || '');
    var texto   = String(l.comment || l.observation || l.description || '');
    var p = re.test(produto);
    var c = re.test(texto);
    if (p) noProduto.push(l);
    else if (c) soNoComentario.push(l);
  });

  var todos = noProduto.concat(soNoComentario);

  Logger.log('--------------------------------------------------');
  Logger.log('Leads da ' + (vendedor || 'rede') + ' no período : ' + daVendedora.length);
  Logger.log('COM INTERESSE EM ' + modelo.toUpperCase() + '      : ' + todos.length);
  Logger.log('   com o produto gravado      : ' + noProduto.length);
  Logger.log('   só citado no comentário    : ' + soNoComentario.length +
             '   <- o POST ignora product_id, por isso essa coluna existe');

  var fechados = todos.filter(function (l) { return !!l.dealed_at; });
  var taxa = todos.length ? (fechados.length / todos.length * 100).toFixed(1) + '%' : '-';
  Logger.log('Negócios fechados              : ' + fechados.length + '  (conversão ' + taxa + ')');

  var porMes = {};
  todos.forEach(function (l) {
    var m = String(l.created_at || '').slice(0, 7) || '????-??';
    porMes[m] = (porMes[m] || 0) + 1;
  });
  Logger.log('Por mês de entrada: ' +
             Object.keys(porMes).sort().map(function (m) { return m + '=' + porMes[m]; }).join('  '));

  Logger.log('--------------------------------------------------');
  todos.sort(function (a, b) {
    return String(a.created_at || '') < String(b.created_at || '') ? -1 : 1;
  }).forEach(function (l) {
    Logger.log([
      String(l.created_at || '').slice(0, 10),
      l.name || '?',
      l.phone || '?',
      String(l.product_name || l.product || l.model || '-'),
      l.dealed_at ? 'FECHADO ' + String(l.dealed_at).slice(0, 10) : (l.status || '')
    ].join('  |  '));
  });
}

// ------------------------------------------------------------- helpers

/** Puxa TODAS as páginas do período. Parar na primeira trunca sem avisar. */
function puxarLeads_(de, ate) {
  var token = PropertiesService.getScriptProperties().getProperty('TEIA_TOKEN');
  if (!token) throw new Error('TEIA_TOKEN não configurado nas Propriedades do script.');

  var todos = [];
  var pagina = 1;
  var ultima = 1;

  do {
    var url = CONSULTA_URL_LISTA + '?createdFrom=' + de + '&createdTo=' + ate +
              '&perPage=500&page=' + pagina + '&token=' + encodeURIComponent(token);

    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (resp.getResponseCode() >= 400) {
      throw new Error('HTTP ' + resp.getResponseCode() + ' na página ' + pagina + ': ' +
                      resp.getContentText().slice(0, 300));
    }

    var d = JSON.parse(resp.getContentText());
    var lote = [];
    if (Array.isArray(d)) lote = d;
    else if (Array.isArray(d.data)) lote = d.data;
    else if (d.data && Array.isArray(d.data.data)) lote = d.data.data;

    todos = todos.concat(lote);
    ultima = Number(d.last_page || (d.data && d.data.last_page) || 1);
    if (!lote.length) break;
    pagina++;
  } while (pagina <= ultima && pagina <= CONSULTA_MAX_PAGINAS);

  if (ultima > CONSULTA_MAX_PAGINAS) {
    throw new Error('O período tem ' + ultima + ' páginas, acima do teto de ' +
                    CONSULTA_MAX_PAGINAS + '. Reduza a janela — seguir truncado daria número errado.');
  }

  return todos;
}

function nomesDistintos_(leads) {
  var vistos = {};
  leads.forEach(function (l) { if (l.user_name) vistos[l.user_name] = true; });
  return Object.keys(vistos).sort();
}

function semAcentoConsulta_(v) {
  return String(v || '').toLowerCase()
    .replace(/[áàãâä]/g, 'a').replace(/[éêë]/g, 'e').replace(/[íï]/g, 'i')
    .replace(/[óõôö]/g, 'o').replace(/[úü]/g, 'u').replace(/ç/g, 'c');
}
