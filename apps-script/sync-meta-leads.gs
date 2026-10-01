/**
 * Sincroniza a planilha do Meta Lead Ads com o TEIA CRM.
 *
 * A cada execução: procura linhas ainda não enviadas, cria o lead no TEIA
 * roteando para a equipe Vendas Digital da loja escolhida no formulário, e
 * escreve de volta na planilha quem o rodízio designou.
 *
 * INSTALAÇÃO
 *   1. Na planilha: Extensões > Apps Script, cole este arquivo.
 *   2. Configurações do projeto > Propriedades do script:
 *        TEIA_TOKEN = <token de Diretor>
 *   3. Rode instalarGatilho() uma vez (autoriza e agenda de 5 em 5 min).
 *
 * Roda de novo com segurança: linha que já tem lead_id preenchido é pulada.
 */

// ----------------------------------------------------------------- config

var URL_LEAD = 'https://teiacrm.com.br/api/lead';
var URL_LISTA = 'https://teiacrm.com.br/api/newapi/leads/1161';

/** Nome exato da aba. Vazio = primeira aba (depende da ordem — preencha). */
var ABA = '';

/** Resposta do formulário => loja + equipe Vendas Digital. */
var LOJAS = {
  'colatina':             { store: 1395, team: 2081 },
  'mantena':              { store: 1396, team: 2082 },
  'sao gabriel da palha': { store: 1397, team: 2083 },
  'afonso claudio':       { store: 1398, team: 2084 }
};

var FONTES = { ig: 4571, fb: 4573 };            // Instagram/Facebook - Concessionária
var PLATAFORMAS = { ig: 'Instagram', fb: 'Facebook' };
var MODELOS = { 'tornado_300': 1422 };          // só o que resolve para UM produto
var PRODUTO_NAO_INFORMADO = 35;                 // "Não informado"

/**
 * "PRECOS E CONDIÇOES" — o cliente ainda não escolheu como vai pagar.
 * Precisa ser enviado explicitamente: omitido, o TEIA preenche sozinho com
 * 853 (CNH - CONSÓRCIO NACIONAL HONDA) e o lead nasce marcado como consórcio
 * mesmo quando o formulário diz "financiamento" ou "à vista".
 */
var PAGAMENTO_A_DEFINIR = 937;

/**
 * Precisa existir, com este nome EXATO, nas configurações do TEIA. Se não
 * existir, o TEIA cria o lead e descarta o array de tags inteiro sem erro —
 * e aí não há como separar o que veio de formulário na hora de medir o Meta.
 * Rode conferirTag() para checar antes de confiar no número.
 */
var TAG = 'Formularios META 2.0';

/** Colunas de controle criadas por este script, se ainda não existirem. */
var COL_VENDEDOR = 'Vendedor';
var COL_LEAD_ID  = 'lead_id TEIA';
var COL_STATUS   = 'Status TEIA';

/**
 * Janela de duplicidade, em dias. Cliente que preencheu de novo depois disso
 * é oportunidade nova, não repetido — janela larga silencia justamente quem
 * voltou a demonstrar interesse.
 */
var DIAS_DUPLICADO = 15;

/**
 * Teto de páginas da janela de duplicidade. Estourar o teto ABORTA a execução
 * em vez de seguir com índice parcial: índice incompleto = duplicata na fila
 * do vendedor. Se estourar, reduza DIAS_DUPLICADO.
 */
var MAX_PAGINAS = 20;

/**
 * O POST ignora product_id: o TEIA lê o modelo do texto do comment e, sem
 * menção a modelo, assume HONDA ADV (id 10). Só o PATCH respeita o valor.
 *
 * Desligado porque o formato do PATCH ainda não foi conferido neste ambiente.
 * Para ativar: ligue, deixe entrar UM lead de Tornado, confira o produto no
 * TEIA, e só então deixe ligado de vez.
 */
var CORRIGIR_PRODUTO = false;

/**
 * Em simulação o script faz tudo — lê, mapeia, monta o payload, checa
 * duplicidade — mas NÃO envia nada ao TEIA e não escreve na planilha.
 * O resultado sai no log (Ver > Registros de execução).
 *
 * Já validado nesta planilha em 25/08/2026: fica false para operar de
 * verdade. Só volte para true se for testar alguma mudança grande.
 */
var SIMULAR = false;

/**
 * Linhas de teste do formulário. Se a coluna Vendedor tiver um destes
 * valores, a linha nunca vira lead — senão o preenchimento de teste entra
 * na fila de um vendedor de verdade, que liga para ninguém.
 */
var MARCADORES_TESTE = ['TEST', 'TESTE'];

/**
 * user_id no TEIA => rótulo EXATO aceito pela validação de dados da coluna
 * Vendedor. A planilha só admite os valores do dropdown; escrever qualquer
 * outra coisa (o nome completo do CRM, por exemplo) faz a gravação falhar e
 * aborta a execução inteira.
 *
 * Vendedor que não estiver aqui não quebra nada: o lead é criado, o lead_id
 * é gravado, e o user_id vai para a coluna Status. Acrescente o rótulo ao
 * dropdown e a linha aqui quando isso acontecer.
 */
var ROTULO_VENDEDOR = {
  // Colatina — Time COL - Vendas Digital (2081)
  9076:  'MAIK',
  9078:  'LARISSA',
  9083:  'GEOVANE',
  9084:  'DAIANE',      // no CRM é "Daiana"; o dropdown usa DAIANE
  10681: 'VIVIANE',
  10682: 'LEIDIANE',
  11104: 'MATHEUS',

  // São Gabriel da Palha — Time SGP - Vendas Digital (2083)
  9089:  'GIOVANA',

  // Afonso Cláudio — Time AFC - Vendas Digital (2084)
  10516: 'IZABELLY',
  9786:  'KAILANY',

  // Mantena — Time MAN - Vendas Digital (2082)
  9585:  'VANESA'       // no CRM é "Vanessa"; o dropdown usa VANESA
};

// ----------------------------------------------------------------- gatilho

function instalarGatilho() {
  if (SIMULAR) {
    throw new Error('SIMULAR ainda está true. Confira o log de sincronizar() ' +
                    'e só então mude para false e instale o gatilho.');
  }

  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sincronizar') ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger('sincronizar').timeBased().everyMinutes(5).create();
  Logger.log('Gatilho instalado: sincronizar() a cada 5 minutos.');
}

// ----------------------------------------------------------------- principal

function sincronizar() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    Logger.log('Outra execução em andamento — saindo.');
    return;
  }

  try {
    var aba = abaDeTrabalho_();
    var dados = aba.getDataRange().getValues();
    if (dados.length < 2) return;

    var cab = dados[0];
    var col = garantirColunas_(aba, cab);

    var pendentes = [];
    for (var i = 1; i < dados.length; i++) {
      if (String(dados[i][col.leadId] || '').trim() === '') {
        pendentes.push({ linha: i + 1, valores: dados[i] });
      }
    }

    if (!pendentes.length) {
      Logger.log('Nada pendente.');
      return;
    }
    Logger.log(pendentes.length + ' linha(s) pendente(s).');

    // Antes do loop e fora do try de cada linha: se o índice de duplicidade
    // não carregar, a execução inteira para. Seguir sem ele criaria duplicata.
    var jaExistem = telefonesRecentes_();

    pendentes.forEach(function (p) {
      try {
        processarLinha_(aba, p, cab, col, jaExistem);
      } catch (e) {
        // Uma linha problemática não pode impedir as outras de rodarem.
        // marcar_ não apaga lead_id já gravado (só escreve se vier valor),
        // então registrar o erro aqui é seguro mesmo após criar o lead.
        Logger.log('Linha ' + p.linha + ' falhou: ' + e);
        try {
          marcar_(aba, p.linha, col, '', '', 'ERRO: ' + (e.message || e));
        } catch (e2) {
          Logger.log('Não deu nem para marcar a linha ' + p.linha + ': ' + e2);
        }
      }
    });

  } finally {
    lock.releaseLock();
  }
}

function processarLinha_(aba, p, cab, col, jaExistem) {
  var r = objetoDaLinha_(cab, p.valores);

  var marcador = String(r[COL_VENDEDOR] || '').trim().toUpperCase();
  if (MARCADORES_TESTE.indexOf(marcador) !== -1) {
    marcar_(aba, p.linha, col, '', 'ignorado', 'Linha de teste — não enviada');
    return;
  }

  var telefone = normalizarTelefone_(r['phone_number']);
  if (telefone.length !== 11) {
    marcar_(aba, p.linha, col, '', '', 'ERRO: telefone com ' + telefone.length + ' dígitos');
    return;
  }

  var destino = mapearLoja_(r['qual_loja_fica_mais_perto_de_você?']);
  if (!destino) {
    marcar_(aba, p.linha, col, '', '', 'ERRO: loja não mapeada (' + r['qual_loja_fica_mais_perto_de_você?'] + ')');
    return;
  }

  var existente = jaExistem[chaveTelefone_(telefone)];
  if (existente) {
    // Grava o lead_id do que JÁ existe. Sem isso a linha continuaria pendente
    // e, quando o lead original saísse da janela de duplicidade, o script
    // criaria uma cópia — duplicata com semanas de atraso.
    //
    // Não mexe na coluna Vendedor: quem já preencheu foi uma pessoa, e o
    // nome do CRM não passa na validação do dropdown.
    var idExistente = idDoLead_(existente);
    var quando = String(existente.created_at || '').slice(0, 10) || 'data ?';

    if (!idExistente) {
      marcar_(aba, p.linha, col, '', '',
              'Já existe no TEIA (' + quando + ') mas a API não devolveu o id — resolva na mão');
      return;
    }
    marcar_(aba, p.linha, col, '', idExistente,
            'Já existia (lead de ' + quando + ', ' + (existente.user_name || '?') + ')');
    return;
  }

  var plataforma = String(r['platform'] || '').toLowerCase().trim();

  var payload = {
    token: token_(),
    name: String(r['full_name']).trim(),
    phone: telefone,
    store_id: destino.store,
    source_id: FONTES[plataforma] || FONTES.ig,
    product_id: MODELOS[String(r['qual_modelo_você_tem_interesse?'] || '').trim()] || PRODUTO_NAO_INFORMADO,
    payment_method_id: PAGAMENTO_A_DEFINIR,
    comment: montarComentario_(r),
    tags: [TAG],
    // auto_assignment é obrigatório: sem ele o TEIA ignora o team_id em silêncio.
    team_id: destino.team,
    auto_assignment: 'team'
  };

  if (SIMULAR) {
    // Escreve na planilha (sem lead_id, para reprocessar quando sair da
    // simulação). Simulação silenciosa parece script quebrado.
    marcar_(aba, p.linha, col, '', '',
            'SIMULAÇÃO — não enviado (loja ' + payload.store_id + ', equipe ' + payload.team_id + ')');
    Logger.log('[SIMULAÇÃO] linha ' + p.linha + ': ' + payload.name + ' / ' + payload.phone +
               ' -> loja ' + payload.store_id + ', equipe ' + payload.team_id);
    return;
  }

  var resp = UrlFetchApp.fetch(URL_LEAD, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  var codigo = resp.getResponseCode();
  var corpo;
  try {
    corpo = JSON.parse(resp.getContentText());
  } catch (e) {
    marcar_(aba, p.linha, col, '', '', 'ERRO HTTP ' + codigo + ': resposta não é JSON');
    return;
  }

  if (codigo >= 400) {
    var msg = corpo.message || corpo.error || 'erro desconhecido';
    marcar_(aba, p.linha, col, '', '', 'ERRO HTTP ' + codigo + ': ' + msg);
    return;
  }

  var lead = corpo.lead || corpo.data || corpo;
  var leadId = idDoLead_(lead);

  if (!leadId) {
    // Sem id na planilha a linha segue pendente e é reenviada a cada 5 min.
    // Avisar alto é melhor que duplicar em silêncio.
    marcar_(aba, p.linha, col, '', '',
            'ERRO: lead criado mas a resposta não trouxe id — confira no TEIA antes da próxima rodada');
    return;
  }

  var rotulo = ROTULO_VENDEDOR[lead.user_id];

  if (rotulo) {
    marcar_(aba, p.linha, col, rotulo, leadId, 'OK ' + agora_());
  } else {
    // Vendedor fora do dropdown: grava o lead normalmente e reporta no Status.
    // Escrever na coluna Vendedor abortaria a gravação.
    marcar_(aba, p.linha, col, '', leadId,
            'OK ' + agora_() + ' — user ' + lead.user_id + ' fora do dropdown (' +
            nomeDoVendedor_(lead.user_id) + '?)');
  }

  if (CORRIGIR_PRODUTO && payload.product_id !== PRODUTO_NAO_INFORMADO) {
    corrigirProduto_(leadId, payload.product_id);
  }

  lead.lead_id = leadId;
  jaExistem[chaveTelefone_(telefone)] = lead;

  Logger.log('Linha ' + p.linha + ': lead ' + leadId + ' -> ' +
             (rotulo || 'user ' + lead.user_id) + ' (equipe ' + lead.team_id + ')');
}

// ----------------------------------------------------------------- TEIA

function token_() {
  var t = PropertiesService.getScriptProperties().getProperty('TEIA_TOKEN');
  if (!t) throw new Error('TEIA_TOKEN não configurado nas Propriedades do script.');
  return t;
}

/** O id do lead, seja qual for o nome do campo na resposta. */
function idDoLead_(lead) {
  if (!lead) return '';
  return lead.lead_id || lead.id || '';
}

/** As linhas de uma página, tolerando paginador aninhado. */
function linhasDaPagina_(d) {
  if (!d) return [];
  if (Array.isArray(d)) return d;
  if (Array.isArray(d.data)) return d.data;
  if (d.data && Array.isArray(d.data.data)) return d.data.data;
  return [];
}

/** Telefones já cadastrados na janela de duplicidade. */
function telefonesRecentes_() {
  var fim = new Date();
  var ini = new Date(fim.getTime() - DIAS_DUPLICADO * 86400000);
  var fmt = function (d) { return Utilities.formatDate(d, 'America/Sao_Paulo', 'yyyy-MM-dd'); };

  var indice = {};
  var pagina = 1;
  var ultima = 1;

  do {
    // perPage explícito: no padrão da API a página vem pequena e a janela
    // inteira não caberia em MAX_PAGINAS. Se a API ignorar o parâmetro, o
    // único efeito é paginar mais.
    var url = URL_LISTA + '?createdFrom=' + fmt(ini) + '&createdTo=' + fmt(fim) +
              '&perPage=500&page=' + pagina + '&token=' + encodeURIComponent(token_());

    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (resp.getResponseCode() >= 400) {
      throw new Error('Falha ao ler leads recentes (HTTP ' + resp.getResponseCode() +
                      ') — execução abortada para não duplicar. O gatilho tenta de novo em 5 min.');
    }

    var d = JSON.parse(resp.getContentText());
    var lote = linhasDaPagina_(d);

    lote.forEach(function (l) {
      var t = normalizarTelefone_(l.phone);
      // Guarda o lead inteiro, não só um "true": quem já existe precisa ter
      // o lead_id anotado na planilha para nunca mais ser reprocessado.
      var k = chaveTelefone_(t);
      if (k && !indice[k]) indice[k] = l;
    });

    ultima = Number(d.last_page || (d.data && d.data.last_page) || 1);
    if (!lote.length) break;
    pagina++;
  } while (pagina <= ultima && pagina <= MAX_PAGINAS);

  if (ultima > MAX_PAGINAS) {
    throw new Error('A janela de duplicidade tem ' + ultima + ' páginas, acima do teto de ' +
                    MAX_PAGINAS + '. Reduza DIAS_DUPLICADO — seguir com índice parcial duplicaria leads.');
  }

  Logger.log('Índice de duplicidade: ' + Object.keys(indice).length + ' telefone(s) dos últimos ' +
             DIAS_DUPLICADO + ' dias.');
  return indice;
}

/**
 * Grava o produto que o POST ignorou. Ver CORRIGIR_PRODUTO.
 * O id vai no corpo; se este ambiente esperar na rota, troque para
 * URL_LEAD + '/' + leadId e remova o id do payload.
 */
function corrigirProduto_(leadId, produtoId) {
  var resp = UrlFetchApp.fetch(URL_LEAD, {
    method: 'patch',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token_() },
    payload: JSON.stringify({ id: leadId, product_id: produtoId }),
    muteHttpExceptions: true
  });

  if (resp.getResponseCode() >= 400) {
    // Não derruba a linha: o lead já está criado e roteado, só o produto
    // ficou com o que o TEIA inferiu do comment.
    Logger.log('PATCH do produto falhou no lead ' + leadId + ': HTTP ' +
               resp.getResponseCode() + ' ' + resp.getContentText().slice(0, 200));
  }
}

/**
 * Nome do vendedor a partir do user_id.
 *
 * ATENÇÃO: /api/user documentado não recebe parâmetro — ele devolve o dono do
 * token. Se for esse o caso, isto retorna sempre o nome do Diretor, e é por
 * isso que o Status marca o resultado com "?" e sempre registra o user_id.
 * A fonte confiável é ROTULO_VENDEDOR; confirme com conferirUser() antes de
 * tratar este nome como verdade.
 */
function nomeDoVendedor_(userId) {
  if (!userId) return '';

  var cache = CacheService.getScriptCache();
  var chave = 'user_' + userId;
  var salvo = cache.get(chave);
  if (salvo) return salvo;

  var resp = UrlFetchApp.fetch('https://teiacrm.com.br/api/user?user_id=' + userId, {
    headers: { Authorization: 'Bearer ' + token_() },
    muteHttpExceptions: true
  });

  if (resp.getResponseCode() >= 400) return String(userId);

  var d = JSON.parse(resp.getContentText());
  var u = Array.isArray(d) ? d[0] : d;

  // Primeiro nome do campo "name", em maiúsculas — é a convenção da planilha
  // (MAIK, LARISSA, GEOVANE). NÃO usar o nickname: cada usuário preencheu do
  // seu jeito e alguns não são nome nenhum (o do Geovane é "Honda").
  var nome = String(userId);
  if (u && u.name) {
    nome = String(u.name).trim().split(/\s+/)[0].toUpperCase();
  }

  cache.put(chave, nome, 21600);   // 6 horas
  return nome;
}

// ----------------------------------------------------------- conferências

/**
 * Roda uma vez, na mão, para saber se a TAG existe de verdade.
 * Cria um lead de teste no SEU usuário, lê as tags que voltaram e avisa.
 * Apague o lead no TEIA depois.
 */
function conferirTag() {
  var resp = UrlFetchApp.fetch(URL_LEAD, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      token: token_(),
      name: 'TESTE TAG - APAGAR',
      phone: '27999999999',
      store_id: 1395,
      payment_method_id: PAGAMENTO_A_DEFINIR,
      comment: 'Lead de teste criado por conferirTag(). Pode apagar.',
      tags: [TAG],
      team_id: 2081,
      auto_assignment: 'team'
    }),
    muteHttpExceptions: true
  });

  Logger.log('HTTP ' + resp.getResponseCode());
  Logger.log(resp.getContentText());
  Logger.log('Se o campo tags voltou vazio, a tag "' + TAG + '" NÃO existe no TEIA ' +
             'com esse nome exato — e todos os leads estão entrando sem tag.');
}

/** Roda uma vez para ver o que /api/user devolve de verdade. */
function conferirUser() {
  [9076, 9585].forEach(function (id) {
    var resp = UrlFetchApp.fetch('https://teiacrm.com.br/api/user?user_id=' + id, {
      headers: { Authorization: 'Bearer ' + token_() },
      muteHttpExceptions: true
    });
    Logger.log('user_id=' + id + ' -> HTTP ' + resp.getResponseCode() + ' ' +
               resp.getContentText().slice(0, 300));
  });
  Logger.log('Se os dois devolveram a MESMA pessoa, /api/user ignora o parâmetro ' +
             'e nomeDoVendedor_() não serve — confie só em ROTULO_VENDEDOR.');
}

// ----------------------------------------------------------------- planilha

function abaDeTrabalho_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ABA ? ss.getSheetByName(ABA) : ss.getSheets()[0];
  if (!aba) throw new Error('Aba "' + ABA + '" não encontrada.');
  return aba;
}

/** Cria as colunas de controle se faltarem e devolve os índices. */
function garantirColunas_(aba, cab) {
  var idx = function (nome) { return cab.indexOf(nome); };

  [COL_VENDEDOR, COL_LEAD_ID, COL_STATUS].forEach(function (nome) {
    if (idx(nome) === -1) {
      cab.push(nome);
      aba.getRange(1, cab.length).setValue(nome);
    }
  });

  return { vendedor: idx(COL_VENDEDOR), leadId: idx(COL_LEAD_ID), status: idx(COL_STATUS) };
}

function marcar_(aba, linha, col, vendedor, leadId, status) {
  // A coluna Vendedor tem validação de dados. Um valor fora da lista lança
  // exceção e derruba a gravação — daí o try: uma linha problemática não pode
  // impedir as outras de serem processadas.
  if (vendedor) {
    try {
      aba.getRange(linha, col.vendedor + 1).setValue(vendedor);
    } catch (e) {
      status += ' [vendedor "' + vendedor + '" recusado pela validação]';
    }
  }

  if (leadId) aba.getRange(linha, col.leadId + 1).setValue(leadId);
  aba.getRange(linha, col.status + 1).setValue(status);
}

function objetoDaLinha_(cab, valores) {
  var o = {};
  cab.forEach(function (nome, i) { o[nome] = valores[i]; });
  return o;
}

// ----------------------------------------------------------------- helpers

/** Meta manda "p:+5527996406008"; o TEIA quer "27996406008". */
function normalizarTelefone_(v) {
  var d = String(v || '').replace(/^[a-z]:/i, '').replace(/\D/g, '');
  if (d.length > 11 && d.slice(0, 2) === '55') d = d.slice(2);
  return d;
}

/**
 * Chave de duplicidade: DDD + os 8 dígitos finais.
 *
 * Ignorar o DDD (slice(-9) puro) junta clientes diferentes de cidades
 * diferentes — a rede tem DDD 27 no ES e 33 em Mantena, então a colisão é
 * possível e o efeito é o pior: lead novo marcado como repetido e ninguém
 * liga. Os 8 finais absorvem o nono dígito, que é o que varia de verdade.
 */
function chaveTelefone_(d) {
  var n = String(d || '');
  if (n.length < 10) return '';
  return n.slice(0, 2) + n.slice(-8);
}

/** Tira emoji e underscore, preserva / - : ( ) % para não quebrar datas. */
function limpar_(v) {
  return String(v || '')
    .replace(/[^\wÀ-ÿ\s\?\$,\.\/\-:()%]/g, '')
    .replace(/_/g, ' ')
    .replace(/\s{2,}/g, ' ')
    .trim();
}

function semAcento_(v) {
  return String(v || '').toLowerCase()
    .replace(/[áàãâä]/g, 'a').replace(/[éêë]/g, 'e').replace(/[íï]/g, 'i')
    .replace(/[óõôö]/g, 'o').replace(/[úü]/g, 'u').replace(/ç/g, 'c');
}

function mapearLoja_(resposta) {
  var alvo = semAcento_(limpar_(resposta));
  for (var nome in LOJAS) {
    if (alvo.indexOf(semAcento_(nome)) !== -1) return LOJAS[nome];
  }
  return null;
}

/** Texto que o vendedor lê antes de ligar. */
function montarComentario_(r) {
  var perguntas = [
    ['Modelo de interesse', 'qual_modelo_você_tem_interesse?'],
    ['Já possui moto',      'você_já_possui_uma_moto?'],
    ['Prazo de compra',     'quando_você_pretende_comprar_ou_trocar_sua_moto?'],
    ['Forma de aquisição',  'como_você_pretende_adquirir_sua_próxima_moto?'],
    ['Tem entrada',         'você_possui_valor_disponível_para_dar_de_entrada?']
  ];

  var plat = String(r['platform'] || '').toLowerCase().trim();
  var linhas = ['Meta Lead Ads - ' + limpar_(r['form_name']),
                'Plataforma : ' + (PLATAFORMAS[plat] || plat.toUpperCase())];

  perguntas.forEach(function (p) {
    var v = limpar_(r[p[1]]);
    if (v) linhas.push(p[0] + ' : ' + v);
  });

  // A pergunta do consórcio tem enunciado longo; acha pelo trecho estável.
  for (var chave in r) {
    if (chave.indexOf('consórcio') !== -1 && String(r[chave] || '').trim()) {
      linhas.push('Consórcio (a partir de R$ 200/mês) : ' + limpar_(r[chave]));
      break;
    }
  }

  [['Loja mais próxima', 'qual_loja_fica_mais_perto_de_você?'],
   ['Campanha', 'campaign_name'],
   ['Anúncio', 'ad_name']].forEach(function (p) {
    if (r[p[1]]) linhas.push(p[0] + ' : ' + limpar_(r[p[1]]));
  });

  if (r['created_time']) linhas.push('Preenchido em : ' + String(r['created_time']).slice(0, 16));

  return linhas.join('\n');
}

function agora_() {
  return Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM HH:mm');
}
