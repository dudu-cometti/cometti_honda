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
var CONSULTA_MAX_PAGINAS = 60;

/**
 * Janela usada pelas consultas de MODELO. A API só filtra por data de
 * criação, então o que for contado é "lead que ENTROU neste período".
 * As consultas de mês calculam a janela sozinhas.
 */
var CONSULTA_DE  = '2026-01-01';
var CONSULTA_ATE = '2026-10-20';

// --------------------------------------------------- quantos leads no mês

/** Quantos leads a Giovana (São Gabriel) recebeu no mês corrente. */
function leadsDoMesGiovana()  { contarLeadsDoMes_('giovana'); }

/** Mesma conta para quem você quiser. */
function leadsDoMesIzabelly() { contarLeadsDoMes_('izabelly'); }
function leadsDoMesMaik()     { contarLeadsDoMes_('maik'); }
function leadsDoMesRede()     { contarLeadsDoMes_(''); }

/** Mês anterior, para comparar. */
function leadsMesPassadoGiovana() {
  var ref = new Date();
  ref.setMonth(ref.getMonth() - 1, 1);
  var mes = Utilities.formatDate(ref, 'America/Sao_Paulo', 'yyyy-MM');
  var ultimo = new Date(ref.getFullYear(), ref.getMonth() + 1, 0);
  contarLeadsNoPeriodo_('giovana', mes + '-01',
                        Utilities.formatDate(ultimo, 'America/Sao_Paulo', 'yyyy-MM-dd'));
}

// --------------------------------------------------- quantos por modelo

/** Quantos clientes da Izabelly (Afonso Cláudio) com interesse em Biz. */
function contarBizIzabelly() {
  contarModeloPorVendedor_('biz', 'izabelly');
}

function contarPopIzabelly()    { contarModeloPorVendedor_('pop', 'izabelly'); }
function contarBizRedeInteira() { contarModeloPorVendedor_('biz', ''); }

// ------------------------------------------------------------- motor

function contarLeadsDoMes_(vendedor) {
  var hoje = new Date();
  var mes = Utilities.formatDate(hoje, 'America/Sao_Paulo', 'yyyy-MM');
  contarLeadsNoPeriodo_(vendedor, mes + '-01',
                        Utilities.formatDate(hoje, 'America/Sao_Paulo', 'yyyy-MM-dd'));
}

function contarLeadsNoPeriodo_(vendedor, de, ate) {
  var leads = puxarLeads_(de, ate);

  Logger.log('=== LEADS RECEBIDOS | ' + (vendedor ? vendedor.toUpperCase() : 'REDE') +
             ' | ' + de + ' a ' + ate + ' ===');
  Logger.log(leads.length + ' lead(s) entraram no período (rede inteira).');
  if (!leads.length) return;

  var dela = filtrarPorVendedor_(leads, vendedor);
  if (vendedor && !dela.length) {
    // Zero silencioso aqui pareceria "não recebeu nada" quando o problema é
    // só a grafia do nome no CRM.
    Logger.log('NENHUM lead com user_name parecido com "' + vendedor + '".');
    Logger.log('Nomes no período: ' + nomesDistintos_(leads).join(' | '));
    return;
  }

  avisarSeJuntouPessoas_(dela, vendedor);

  Logger.log('--------------------------------------------------');
  Logger.log('LEADS RECEBIDOS : ' + dela.length);

  var fechados = dela.filter(function (l) { return !!l.dealed_at; });
  var taxa = dela.length ? (fechados.length / dela.length * 100).toFixed(1) + '%' : '-';
  // Atenção ao recorte: estes são os leads que ENTRARAM no período e já
  // fecharam. Não é "vendas do mês" — venda de lead antigo não está aqui.
  Logger.log('Desses, já fecharam : ' + fechados.length + '  (' + taxa + ')');

  Logger.log('Por fonte : ' + contarPor_(dela, function (l) {
    return l.source_name || l.source || ('fonte ' + (l.source_id || '?'));
  }));
  Logger.log('Por status : ' + contarPor_(dela, function (l) { return l.status || '-'; }));

  Logger.log('--------------------------------------------------');
  listar_(dela);
}

function contarModeloPorVendedor_(modelo, vendedor) {
  var leads = puxarLeads_(CONSULTA_DE, CONSULTA_ATE);

  Logger.log('=== ' + modelo.toUpperCase() + (vendedor ? ' / ' + vendedor.toUpperCase() : ' / REDE') +
             ' | ' + CONSULTA_DE + ' a ' + CONSULTA_ATE + ' ===');
  Logger.log(leads.length + ' lead(s) criados no período (rede inteira).');
  if (!leads.length) return;

  // Os nomes dos campos da resposta não estão documentados. Isto mostra um
  // registro inteiro para conferir onde ficam produto e comentário antes de
  // confiar na conta.
  Logger.log('Exemplo de registro (confira os nomes dos campos):');
  Logger.log(JSON.stringify(leads[0]));

  var dela = filtrarPorVendedor_(leads, vendedor);
  if (vendedor && !dela.length) {
    Logger.log('NENHUM lead com user_name parecido com "' + vendedor + '".');
    Logger.log('Nomes no período: ' + nomesDistintos_(leads).join(' | '));
    return;
  }
  avisarSeJuntouPessoas_(dela, vendedor);

  // \b + modelo pega "Biz 125", "BIZ110i", "biz-125", e não pega palavra
  // que só contenha as letras no meio.
  var re = new RegExp('\\b' + modelo, 'i');

  var noProduto = [];
  var soNoComentario = [];

  dela.forEach(function (l) {
    var produto = String(l.product_name || l.product || l.model || '');
    var texto   = String(l.comment || l.observation || l.description || '');
    if (re.test(produto)) noProduto.push(l);
    else if (re.test(texto)) soNoComentario.push(l);
  });

  var todos = noProduto.concat(soNoComentario);

  Logger.log('--------------------------------------------------');
  Logger.log('Leads de ' + (vendedor || 'toda a rede') + ' no período : ' + dela.length);
  Logger.log('COM INTERESSE EM ' + modelo.toUpperCase() + '      : ' + todos.length);
  Logger.log('   com o produto gravado      : ' + noProduto.length);
  Logger.log('   só citado no comentário    : ' + soNoComentario.length +
             '   <- o POST ignora product_id, por isso essa coluna existe');

  var fechados = todos.filter(function (l) { return !!l.dealed_at; });
  var taxa = todos.length ? (fechados.length / todos.length * 100).toFixed(1) + '%' : '-';
  Logger.log('Negócios fechados              : ' + fechados.length + '  (conversão ' + taxa + ')');

  Logger.log('Por mês de entrada : ' + contarPor_(todos, function (l) {
    return String(l.created_at || '').slice(0, 7) || '????-??';
  }));

  Logger.log('--------------------------------------------------');
  listar_(todos);
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

/**
 * Chave de comparação de nome: sem acento, caixa baixa e com letra repetida
 * colapsada. É o que faz "giovanna" achar a "Giovana" do CRM e "vanessa"
 * achar a "VANESA" do dropdown — a grafia nunca bate de primeira.
 */
function chaveNome_(v) {
  return String(v || '').toLowerCase()
    .replace(/[áàãâä]/g, 'a').replace(/[éêë]/g, 'e').replace(/[íï]/g, 'i')
    .replace(/[óõôö]/g, 'o').replace(/[úü]/g, 'u').replace(/ç/g, 'c')
    .replace(/(.)\1+/g, '$1');
}

/** Filtra por vendedor usando user_name (team_id vem nulo na maioria dos leads). */
function filtrarPorVendedor_(leads, vendedor) {
  if (!vendedor) return leads;
  var alvo = chaveNome_(vendedor);
  return leads.filter(function (l) {
    return chaveNome_(l.user_name || '').indexOf(alvo) !== -1;
  });
}

/** Busca parcial pode juntar duas pessoas; se juntar, mostra o rateio. */
function avisarSeJuntouPessoas_(leads, vendedor) {
  if (!vendedor) return;
  var nomes = nomesDistintos_(leads);
  if (nomes.length > 1) {
    Logger.log('ATENÇÃO: "' + vendedor + '" casou com ' + nomes.length + ' pessoas — ' +
               contarPor_(leads, function (l) { return l.user_name || '?'; }));
  } else if (nomes.length === 1) {
    Logger.log('Vendedor no CRM : ' + nomes[0]);
  }
}

function contarPor_(leads, chave) {
  var c = {};
  leads.forEach(function (l) {
    var k = String(chave(l));
    c[k] = (c[k] || 0) + 1;
  });
  return Object.keys(c)
    .sort(function (a, b) { return c[b] - c[a] || (a < b ? -1 : 1); })
    .map(function (k) { return k + '=' + c[k]; })
    .join('  ');
}

function listar_(leads) {
  leads.slice().sort(function (a, b) {
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

function nomesDistintos_(leads) {
  var vistos = {};
  leads.forEach(function (l) { if (l.user_name) vistos[l.user_name] = true; });
  return Object.keys(vistos).sort();
}
