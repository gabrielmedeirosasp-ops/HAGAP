/**
 * HAGAP — CONTROLE DE OBRAS (Apps Script)
 * Independente do site HAGAP atual.
 * Gmail -> Base -> Cruzamentos -> Painel Web.
 */

const HAGAP = {
  INICIO_HISTORICO: '2025-01-01',
  INTERVALO_MINUTOS: 5,
  LIMITE_MS: 4.5 * 60 * 1000,
  LOTE_THREADS: 20,
  LOGO_URL: 'https://raw.githubusercontent.com/gabrielmedeirosasp-ops/HAGAP/main/logo_hagap.png',
  ABAS: {
    EVENTOS: 'EVENTOS',
    EMAILS: 'EMAILS_PROCESSADOS',
    LOG: 'LOG',
    CONFIG: 'CONFIG'
  },
  HEAD_EVENTOS: [
    'CHAVE_EVENTO','ID_EMAIL','DATA_EMAIL','TIPO','PROJETO','PROJETO_BASE',
    'REFERENCIA','PDE_RELACIONADO','STATUS','MUNICIPIO',
    'DATA_SERVICO','HORA_INICIO','HORA_FIM',
    'ASSUNTO','ANEXO','URL_EMAIL','DATA_PROCESSAMENTO','OBSERVACAO'
  ],
  HEAD_EMAILS: [
    'ID_EMAIL','DATA_EMAIL','ASSUNTO','TIPO','STATUS','DATA_PROCESSAMENTO'
  ]
};


function configurarSistema() {
  const props = PropertiesService.getScriptProperties();
  let ss;

  const id = props.getProperty('HAGAP_PLANILHA_ID');
  if (id) {
    ss = SpreadsheetApp.openById(id);
  } else {
    ss = SpreadsheetApp.create('HAGAP - CONTROLE DE OBRAS');
    props.setProperty('HAGAP_PLANILHA_ID', ss.getId());
  }

  garantirAba_(ss, HAGAP.ABAS.EVENTOS, HAGAP.HEAD_EVENTOS);
  garantirAba_(ss, HAGAP.ABAS.EMAILS, HAGAP.HEAD_EMAILS);
  garantirAba_(ss, HAGAP.ABAS.LOG, ['DATA','NIVEL','ACAO','DETALHE']);
  garantirAba_(ss, HAGAP.ABAS.CONFIG, ['CHAVE','VALOR']);

  const cfg = ss.getSheetByName(HAGAP.ABAS.CONFIG);
  if (cfg.getLastRow() <= 1) {
    cfg.getRange(2,1,5,2).setValues([
      ['INICIO_HISTORICO', HAGAP.INICIO_HISTORICO],
      ['INTERVALO_MINUTOS', String(HAGAP.INTERVALO_MINUTOS)],
      ['ALERTA_PEDIDO_DIAS', '1'],
      ['PLANILHA_ID', ss.getId()],
      ['PLANILHA_URL', ss.getUrl()]
    ]);
  }

  if (!props.getProperty('BACKFILL_MES')) {
    props.setProperty('BACKFILL_MES', HAGAP.INICIO_HISTORICO.substring(0,7));
    props.setProperty('BACKFILL_OFFSET', '0');
    props.setProperty('BACKFILL_CONCLUIDO', '0');
  }

  // Migração v1: PEDIDOS passaram a fazer parte da rotina.
  // Reinicia a varredura histórica uma única vez; os IDs já processados
  // continuam deduplicados, então não duplica os eventos existentes.
  if (props.getProperty('MIGRACAO_PEDIDO_V1') !== '1') {
    props.setProperty('BACKFILL_MES', HAGAP.INICIO_HISTORICO.substring(0,7));
    props.setProperty('BACKFILL_OFFSET', '0');
    props.setProperty('BACKFILL_CONCLUIDO', '0');
    props.setProperty('MIGRACAO_PEDIDO_V1', '1');
  }

  criarTriggerAutomatico_();
  log_('CONFIRMADO','CONFIGURAR','Sistema configurado: ' + ss.getUrl());

  return {
    ok:true,
    planilhaUrl:ss.getUrl(),
    mensagem:'Configurado. Execute sincronizarAgora() uma vez.'
  };
}


function criarTriggerAutomatico_() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'rotinaAutomatica') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('rotinaAutomatica')
    .timeBased()
    .everyMinutes(HAGAP.INTERVALO_MINUTOS)
    .create();
}


function rotinaAutomatica() {
  return sincronizarGmail_();
}


function sincronizarAgora() {
  return sincronizarGmail_();
}


function sincronizarGmail_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) {
    return {ok:false,status:'EM_EXECUCAO'};
  }

  const inicio = Date.now();
  const prazo = inicio + HAGAP.LIMITE_MS;
  const ss = getSS_();

  try {
    const processados = carregarIdsProcessados_(ss);
    const chaves = carregarChavesEventos_(ss);

    const stats = {
      ok:true,
      emailsNovos:0,
      eventosNovos:0,
      erros:0,
      incremental:0,
      historico:0
    };

    // Sempre prioriza novidades recentes.
    if (Date.now() < prazo) {
      const inc = processarBusca_(
        ss,
        'newer_than:7d {in:sent subject:DOCUMENTOS subject:"Confirmação de Recebimento: Solicitação para o Projeto" subject:"PDE número:" subject:"PLV número:" subject:"OMB Enviado Integração" subject:BMD subject:FFO}',
        processados,
        chaves,
        prazo,
        null
      );
      somarStats_(stats, inc);
      stats.incremental += inc.emailsNovos;
    }

    // Depois continua a carga histórica desde 01/2025.
    if (Date.now() < prazo && !backfillConcluido_()) {
      const hist = processarBackfill_(ss, processados, chaves, prazo);
      somarStats_(stats, hist);
      stats.historico += hist.emailsNovos;
    }

    resolverOmbs_(ss);

    stats.backfillConcluido = backfillConcluido_();
    PropertiesService.getScriptProperties()
      .setProperty('ULTIMA_SYNC_OK', new Date().toISOString());

    log_(
      stats.erros ? 'ATENCAO' : 'CONFIRMADO',
      'SINCRONIZAR',
      JSON.stringify(stats)
    );

    return stats;

  } finally {
    lock.releaseLock();
  }
}


function processarBackfill_(ss, processados, chaves, prazo) {
  const props = PropertiesService.getScriptProperties();
  const mes = props.getProperty('BACKFILL_MES') || HAGAP.INICIO_HISTORICO.substring(0,7);
  const offset = Number(props.getProperty('BACKFILL_OFFSET') || '0');
  const janela = janelaMes_(mes);

  if (janela.inicio > new Date()) {
    props.setProperty('BACKFILL_CONCLUIDO','1');
    return {emailsNovos:0,eventosNovos:0,erros:0,threadsEncontradas:0,threadsProcessadas:0};
  }

  const query =
    'after:' + fmtQueryDate_(janela.inicio) +
    ' before:' + fmtQueryDate_(janela.fim) +
    ' {in:sent subject:DOCUMENTOS subject:"Confirmação de Recebimento: Solicitação para o Projeto" subject:"PDE número:" subject:"PLV número:" subject:"OMB Enviado Integração" subject:BMD subject:FFO}';

  const r = processarBusca_(
    ss, query, processados, chaves, prazo,
    {inicio:janela.inicio,fim:janela.fim,offset:offset}
  );

  if (r.threadsEncontradas < HAGAP.LOTE_THREADS && Date.now() < prazo) {
    const prox = addMes_(janela.inicio,1);
    props.setProperty('BACKFILL_MES',fmtMes_(prox));
    props.setProperty('BACKFILL_OFFSET','0');

    if (prox > new Date()) {
      props.setProperty('BACKFILL_CONCLUIDO','1');
      log_('CONFIRMADO','BACKFILL','Carga histórica concluída.');
    }
  } else {
    props.setProperty(
      'BACKFILL_OFFSET',
      String(offset + r.threadsProcessadas)
    );
  }

  return r;
}


function processarBusca_(ss, query, processados, chaves, prazo, janela) {
  const offset = janela ? janela.offset : 0;
  const threads = GmailApp.search(query, offset, HAGAP.LOTE_THREADS);

  const stats = {
    emailsNovos:0,
    eventosNovos:0,
    erros:0,
    threadsEncontradas:threads.length,
    threadsProcessadas:0
  };

  for (let ti=0; ti<threads.length; ti++) {
    if (Date.now() >= prazo) break;

    const mensagens = threads[ti].getMessages();

    for (let mi=0; mi<mensagens.length; mi++) {
      if (Date.now() >= prazo) break;

      const msg = mensagens[mi];

      if (janela) {
        const d = msg.getDate();
        if (d < janela.inicio || d >= janela.fim) continue;
      }

      const id = msg.getId();
      if (processados.has(id)) continue;

      try {
        const eventos = processarMensagem_(msg);
        if (!eventos.length) continue;

        const novos = gravarEventos_(ss,eventos,chaves);
        registrarEmailProcessado_(ss,msg,eventos[0].tipo,'OK');
        processados.add(id);

        stats.emailsNovos++;
        stats.eventosNovos += novos;

      } catch (e) {
        stats.erros++;
        log_(
          'PENDENTE',
          'EMAIL ' + id,
          msg.getSubject() + ' | ' + String(e && e.message ? e.message : e)
        );
      }
    }

    stats.threadsProcessadas++;
  }

  return stats;
}


function processarMensagem_(msg) {
  const assunto = msg.getSubject() || '';
  const tipo = classificarAssunto_(assunto);
  if (!tipo) return [];

  const anexos = msg.getAttachments({
    includeInlineImages:false,
    includeAttachments:true
  });

  const base = {
    idEmail:msg.getId(),
    dataEmail:msg.getDate(),
    assunto:assunto,
    urlEmail:'https://mail.google.com/mail/u/0/#all/' + msg.getId()
  };

  if (tipo === 'PEDIDO') {
    const corpo = msg.getPlainBody() || '';
    const projeto = extrairProjetoPedido_(assunto, corpo);

    if (!projeto) {
      throw new Error('[PENDENTE] Pedido sem projeto confirmado.');
    }

    const info = extrairPedido_(corpo);

    return [evento_(base,{
      tipo:'PEDIDO',
      projeto:projeto,
      status:info.tipoDocumento || 'PEDIDO',
      municipio:info.municipio || '',
      dataServico:info.data || '',
      horaInicio:info.horaInicio || '',
      horaFim:info.horaFim || '',
      observacao:info.servico || ''
    })];
  }

  if (tipo === 'DOC_FINAL' || tipo === 'DOC_PARCIAL') {
    const nomes = anexos.map(a => a.getName()).join(' ');
    const projetos = extrairProjetosLivres_(
      assunto + '\n' + (msg.getPlainBody() || '') + '\n' + nomes
    );

    if (!projetos.length) {
      throw new Error('[PENDENTE] DOCUMENTOS sem projeto confirmado.');
    }

    return projetos.map(projeto => evento_(base,{
      tipo:tipo,
      projeto:projeto,
      status:tipo === 'DOC_PARCIAL' ? 'PARCIAL' : 'ENVIADO'
    }));
  }

  if (tipo === 'PDE' || tipo === 'PLV' || tipo === 'OMB') {
    const pdf = primeiroPdf_(anexos);
    if (!pdf) {
      throw new Error('[PENDENTE] ' + tipo + ' sem PDF anexado.');
    }

    const texto = extrairTextoPdf_(pdf.copyBlob());
    if (!texto || texto.trim().length < 20) {
      throw new Error('[PENDENTE] Texto do PDF não extraído: ' + pdf.getName());
    }

    if (tipo === 'PDE') {
      const projeto = extrairProjetoRotulado_(texto);
      if (!projeto) {
        throw new Error('[PENDENTE] PDE sem Nº Projeto confirmado.');
      }

      return [evento_(base,{
        tipo:'PDE',
        projeto:projeto,
        referencia:extrairReferencia_(assunto,texto,'PDE'),
        status:extrairStatusAssunto_(assunto,'PDE'),
        municipio:extrairMunicipioPde_(assunto,texto),
        dataServico:extrairData_(texto,'Data Confirmada') || extrairData_(texto,'Data Solicitada'),
        anexo:pdf.getName()
      })];
    }

    if (tipo === 'PLV') {
      const projeto = extrairProjetoRotulado_(texto);
      if (!projeto) {
        throw new Error('[PENDENTE] PLV sem Nº Projeto confirmado.');
      }

      return [evento_(base,{
        tipo:'PLV',
        projeto:projeto,
        referencia:extrairReferencia_(assunto,texto,'PLV'),
        status:extrairStatusAssunto_(assunto,'PLV'),
        municipio:extrairMunicipioAssunto_(assunto),
        dataServico:extrairData_(texto,'Data Confirmada') || extrairData_(texto,'Data Solicitada'),
        anexo:pdf.getName()
      })];
    }

    const pdeRel = extrairReferencia_('',texto,'PDE');
    if (!pdeRel) {
      throw new Error('[PENDENTE] OMB sem PDE relacionado confirmado.');
    }

    const periodo = extrairPeriodoOmb_(texto);

    return [evento_(base,{
      tipo:'OMB',
      projeto:'',
      referencia:extrairReferencia_(assunto,texto,'OMB'),
      pdeRelacionado:pdeRel,
      status:'EMITIDA',
      municipio:extrairMunicipioAssunto_(assunto),
      dataServico:periodo.data || '',
      horaInicio:periodo.inicio || '',
      horaFim:periodo.fim || '',
      anexo:pdf.getName()
    })];
  }

  if (tipo === 'BMD' || tipo === 'FFO' || tipo === 'BMD_FFO') {
    const eventos = [];

    anexos.forEach(anexo => {
      const nome = anexo.getName() || '';
      const n = normalizar_(nome);

      let t = '';
      if (/\bFFO\b|\bFF0\b/.test(n)) t = 'FFO';
      else if (/\bBMD\b/.test(n)) t = 'BMD';
      else return; // MULTA e outros não viram BMD.

      const projetos = extrairProjetosLivres_(nome);
      if (!projetos.length) return;

      projetos.forEach(projeto => {
        const conjunto = normalizar_(assunto + ' ' + nome);
        eventos.push(evento_(base,{
          tipo:t,
          projeto:projeto,
          status:t === 'FFO'
            ? 'FFO'
            : (conjunto.indexOf('PARCIAL') >= 0
                ? 'PARCIAL'
                : (conjunto.indexOf('FINAL') >= 0 ? 'FINAL' : 'BMD')),
          anexo:nome
        }));
      });
    });

    if (!eventos.length && tipo !== 'BMD_FFO') {
      const projetos = extrairProjetosLivres_(assunto);
      if (projetos.length === 1) {
        eventos.push(evento_(base,{
          tipo:tipo,
          projeto:projetos[0],
          status:normalizar_(assunto).indexOf('PARCIAL') >= 0 ? 'PARCIAL' : tipo
        }));
      }
    }

    if (!eventos.length) {
      throw new Error('[PENDENTE] BMD/FFO sem projeto confirmado.');
    }

    return eventos;
  }

  return [];
}


function extrairTextoPdf_(blob) {
  // PDF -> Google Docs temporário -> texto -> lixeira.
  let arq = null;

  try {
    arq = Drive.Files.create(
      {
        name:'HAGAP_TMP_' + Date.now(),
        mimeType:'application/vnd.google-apps.document'
      },
      blob,
      {
        ocrLanguage:'pt',
        fields:'id,name'
      }
    );

    Utilities.sleep(250);
    return DocumentApp.openById(arq.id).getBody().getText() || '';

  } finally {
    if (arq && arq.id) {
      try {
        Drive.Files.update({trashed:true},arq.id);
      } catch (e) {
        log_('ATENCAO','LIMPAR_TEMP',arq.id + ' | ' + e.message);
      }
    }
  }
}


function primeiroPdf_(anexos) {
  for (let i=0;i<anexos.length;i++) {
    const nome = (anexos[i].getName() || '').toLowerCase();
    const mime = (anexos[i].getContentType() || '').toLowerCase();

    if (nome.endsWith('.pdf') || mime === 'application/pdf') {
      return anexos[i];
    }
  }
  return null;
}


function classificarAssunto_(assunto) {
  const s = normalizar_(assunto);

  if (s.indexOf('CONFIRMACAO DE RECEBIMENTO: SOLICITACAO PARA O PROJETO') >= 0) return 'PEDIDO';
  if (s.indexOf('DOCUMENTOS PARCIAL') >= 0 || s.indexOf('DOCUMENTO PARCIAL') >= 0) return 'DOC_PARCIAL';
  if (s.indexOf('DOCUMENTOS') >= 0 || s.indexOf('DOCUMENTO ') === 0) return 'DOC_FINAL';
  if (s.indexOf('PDE NUMERO:') >= 0 || s.indexOf('PDE ') === 0) return 'PDE';
  if (s.indexOf('PLV NUMERO:') >= 0 || s.indexOf('PLV ') === 0) return 'PLV';
  if (s.indexOf('OMB ENVIADO INTEGRACAO') >= 0) return 'OMB';

  const bmd = /\bBMD\b/.test(s);
  const ffo = /\bFFO\b|\bFF0\b/.test(s);

  if (bmd && ffo) return 'BMD_FFO';
  if (ffo) return 'FFO';
  if (bmd) return 'BMD';

  return '';
}


function extrairProjetoPedido_(assunto,corpo) {
  const fonte = String(assunto || '') + '\n' + String(corpo || '');
  const m = fonte.match(/(?:PROJETO|projeto)\s+(\d{7}[A-Za-z]?)/);
  return m ? m[1].toUpperCase() : '';
}


function extrairPedido_(corpo) {
  const t = String(corpo || '');
  const municipio = (t.match(/Munic[ií]pio\s*:\s*([^\n\r]+)/i) || [,''])[1].trim();
  const dataMatch = t.match(/Data\s*:\s*(\d{2}\/\d{2}\/\d{4})(?:\s+das?\s+(\d{1,2}:\d{2})\s+(?:à|a)s?\s+(\d{1,2}:\d{2}))?/i);
  const servico = (t.match(/Servi[cç]o\s+Solicitado\s*:\s*([\s\S]*?)(?:\n\s*\n|A\s+Copel\s+entrar[aá]|$)/i) || [,''])[1]
    .replace(/\s+/g,' ')
    .trim();
  const tipoDocumento = (t.match(/documento\s+do\s+tipo\s+(PDE|PLV)/i) || [,''])[1].toUpperCase();

  return {
    municipio:municipio,
    data:dataMatch ? dataMatch[1] : '',
    horaInicio:dataMatch && dataMatch[2] ? dataMatch[2] : '',
    horaFim:dataMatch && dataMatch[3] ? dataMatch[3] : '',
    servico:servico,
    tipoDocumento:tipoDocumento
  };
}


function extrairProjetoRotulado_(texto) {
  const patterns = [
    /N[º°]\s*Projeto\s*[:.]?\s*(\d{7}[A-Za-z]?)/i,
    /N[º°]\s*PROJETO\s*[:.]?\s*(\d{7}[A-Za-z]?)/i,
    /FISCAL\s+N[º°]\s*PROJETO[^\n]*\n[^\n]*?\b(\d{7}[A-Za-z]?)\b/i
  ];

  for (let i=0;i<patterns.length;i++) {
    const m = String(texto || '').match(patterns[i]);
    if (m) return m[1].toUpperCase();
  }

  const candidatos = extrairProjetosLivres_(texto);
  return candidatos.length === 1 ? candidatos[0] : '';
}


function extrairProjetosLivres_(texto) {
  const re = /(^|[^\d])(\d{7})([A-Za-z])?(?!\d)/g;
  const vistos = {};
  const out = [];
  let m;

  while ((m = re.exec(String(texto || ''))) !== null) {
    const p = m[2] + (m[3] ? m[3].toUpperCase() : '');
    if (!vistos[p]) {
      vistos[p] = true;
      out.push(p);
    }
  }

  return out;
}


function extrairReferencia_(assunto,texto,rotulo) {
  const fonte = String(assunto || '') + '\n' + String(texto || '');
  const p = new RegExp(
    '\\b' + rotulo + '\\s*(?:n[º°o.]*)?\\s*[:\\-]?\\s*0*(\\d{1,5})\\s*[\\/\\-]\\s*(20\\d{2})',
    'i'
  );
  const m = fonte.match(p);
  return m ? String(Number(m[1])) + '/' + m[2] : '';
}


function extrairStatusAssunto_(assunto,rotulo) {
  const p = new RegExp(
    rotulo + '\\s*(?:n[uú]mero\\s*:)?\\s*\\d+\\s*[\\/\\-]\\s*\\d{4}\\s*-\\s*([^-]+)',
    'i'
  );
  const m = String(assunto || '').match(p);
  return m ? limpar_(m[1]) : '';
}


function extrairMunicipioAssunto_(assunto) {
  const partes = String(assunto || '').split(' - ').map(limpar_);
  const ignorar = {
    ABERTO:true,APROVADO:true,CANCELADO:true,REPROVADO:true,
    ENCERRADO:true,FECHADO:true
  };

  for (let i=1;i<partes.length;i++) {
    const n = normalizar_(partes[i]);
    if (!partes[i] || ignorar[n]) continue;
    if (n.indexOf('RESP') >= 0 || n.indexOf('ENVIADO POR') >= 0) continue;
    if (/^\d{1,2}\/\d{1,2}\/\d{4}/.test(partes[i])) continue;
    return partes[i];
  }

  return '';
}


function extrairMunicipioPde_(assunto,texto) {
  const a = extrairMunicipioAssunto_(assunto);
  if (a) return a;

  const m = String(texto || '').match(
    /C[oó]d\.\s*\/\s*Localidade\s*\/\s*Munic[ií]pio\s*:\s*[^\n]*\/\s*([^\n]+)/i
  );
  return m ? limpar_(m[1]) : '';
}


function extrairData_(texto,rotulo) {
  const p = new RegExp(rotulo + '[^\\d]*(\\d{2}[\\/.]\\d{2}[\\/.]\\d{4})','i');
  const m = String(texto || '').match(p);
  return m ? m[1].replace(/\./g,'/') : '';
}


function extrairPeriodoOmb_(texto) {
  const m = String(texto || '').match(
    /(\d{2}\/\d{2}\/\d{4}).{0,100}?(\d{1,2}:\d{2}).{0,60}?(?:ÀS|AS)\s*(\d{1,2}:\d{2})/is
  );

  return m ? {data:m[1],inicio:m[2],fim:m[3]} : {};
}


function normalizar_(v) {
  return String(v || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g,'')
    .replace(/\s+/g,' ')
    .trim()
    .toUpperCase();
}


function limpar_(v) {
  return String(v || '').replace(/\s+/g,' ').trim();
}


function evento_(base,dados) {
  const projeto = String(dados.projeto || '').toUpperCase();

  const ev = {
    idEmail:base.idEmail,
    dataEmail:base.dataEmail,
    tipo:dados.tipo || '',
    projeto:projeto,
    projetoBase:projeto ? projeto.substring(0,7) : '',
    referencia:dados.referencia || '',
    pdeRelacionado:dados.pdeRelacionado || '',
    status:dados.status || '',
    municipio:dados.municipio || '',
    dataServico:dados.dataServico || '',
    horaInicio:dados.horaInicio || '',
    horaFim:dados.horaFim || '',
    assunto:base.assunto || '',
    anexo:dados.anexo || '',
    urlEmail:base.urlEmail || '',
    dataProcessamento:new Date(),
    observacao:dados.observacao || ''
  };

  ev.chaveEvento = [
    ev.idEmail,ev.tipo,ev.projeto,ev.referencia,ev.anexo
  ].join('|');

  return ev;
}


function gravarEventos_(ss,eventos,chaves) {
  const aba = ss.getSheetByName(HAGAP.ABAS.EVENTOS);
  const linhas = [];

  eventos.forEach(ev => {
    if (chaves.has(ev.chaveEvento)) return;

    linhas.push([
      ev.chaveEvento,ev.idEmail,ev.dataEmail,ev.tipo,ev.projeto,ev.projetoBase,
      ev.referencia,ev.pdeRelacionado,ev.status,ev.municipio,
      ev.dataServico,ev.horaInicio,ev.horaFim,
      ev.assunto,ev.anexo,ev.urlEmail,ev.dataProcessamento,ev.observacao
    ]);

    chaves.add(ev.chaveEvento);
  });

  if (linhas.length) {
    aba.getRange(
      aba.getLastRow()+1,1,linhas.length,HAGAP.HEAD_EVENTOS.length
    ).setValues(linhas);
  }

  return linhas.length;
}


function registrarEmailProcessado_(ss,msg,tipo,status) {
  ss.getSheetByName(HAGAP.ABAS.EMAILS).appendRow([
    msg.getId(),msg.getDate(),msg.getSubject(),tipo,status,new Date()
  ]);
}


function carregarIdsProcessados_(ss) {
  const aba = ss.getSheetByName(HAGAP.ABAS.EMAILS);
  const set = new Set();

  if (aba.getLastRow() <= 1) return set;

  aba.getRange(2,1,aba.getLastRow()-1,5).getValues().forEach(r => {
    if (r[0] && String(r[4]).toUpperCase() === 'OK') {
      set.add(String(r[0]));
    }
  });

  return set;
}


function carregarChavesEventos_(ss) {
  const aba = ss.getSheetByName(HAGAP.ABAS.EVENTOS);
  const set = new Set();

  if (aba.getLastRow() <= 1) return set;

  aba.getRange(2,1,aba.getLastRow()-1,1).getValues().forEach(r => {
    if (r[0]) set.add(String(r[0]));
  });

  return set;
}


function resolverOmbs_(ss) {
  const aba = ss.getSheetByName(HAGAP.ABAS.EVENTOS);
  if (aba.getLastRow() <= 1) return;

  const range = aba.getRange(
    2,1,aba.getLastRow()-1,HAGAP.HEAD_EVENTOS.length
  );
  const dados = range.getValues();

  const pdes = {};

  dados.forEach(r => {
    if (String(r[3]).toUpperCase() === 'PDE' && r[6] && r[4]) {
      pdes[String(r[6])] = String(r[4]);
    }
  });

  let alterou = false;

  dados.forEach(r => {
    if (
      String(r[3]).toUpperCase() === 'OMB' &&
      !r[4] &&
      r[7]
    ) {
      const projeto = pdes[String(r[7])];
      if (projeto) {
        r[4] = projeto;
        r[5] = projeto.substring(0,7);
        alterou = true;
      }
    }
  });

  if (alterou) range.setValues(dados);
}


function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('HAGAP — Controle de Obras');
}


function getPainelData() {
  const ss = getSS_();
  resolverOmbs_(ss);

  const eventos = lerEventos_(ss);
  const mapa = {};

  eventos.forEach(e => {
    if (!e.projeto) return;

    // Projeto I/C/S permanece separado. Não mistura silenciosamente.
    if (!mapa[e.projeto]) {
      mapa[e.projeto] = {
        projeto:e.projeto,
        projetoBase:e.projetoBase,
        municipio:'',
        docFinal:false,
        docParcial:false,
        pedido:false,
        pedidos:[],
        pdes:{},plvs:{},ombs:{},
        bmds:0,ffos:0,
        ultimoMovimento:''
      };
    }

    const p = mapa[e.projeto];

    if (e.municipio && !p.municipio) p.municipio = e.municipio;
    if (e.tipo === 'PEDIDO') {
      p.pedido = true;
      p.pedidos.push({
        dataEmail:dataIso_(e.dataEmail),
        dataServico:e.dataServico,
        horaInicio:e.horaInicio,
        horaFim:e.horaFim,
        tipoDocumento:e.status,
        observacao:e.observacao,
        urlEmail:e.urlEmail
      });
    }
    if (e.tipo === 'DOC_FINAL') p.docFinal = true;
    if (e.tipo === 'DOC_PARCIAL') p.docParcial = true;
    if (e.tipo === 'PDE' && e.referencia) p.pdes[e.referencia] = true;
    if (e.tipo === 'PLV' && e.referencia) p.plvs[e.referencia] = true;
    if (e.tipo === 'OMB' && e.referencia) p.ombs[e.referencia] = true;
    if (e.tipo === 'BMD') p.bmds++;
    if (e.tipo === 'FFO') p.ffos++;

    const iso = dataIso_(e.dataEmail);
    if (iso > p.ultimoMovimento) p.ultimoMovimento = iso;
  });

  const projetos = Object.keys(mapa).map(k => {
    const p = mapa[k];

    p.qtdPde = Object.keys(p.pdes).length;
    p.qtdPlv = Object.keys(p.plvs).length;
    p.qtdOmb = Object.keys(p.ombs).length;
    p.programada = p.qtdPde > 0 || p.qtdPlv > 0 || p.qtdOmb > 0;
    p.temLiberacao = p.qtdPde > 0 || p.qtdPlv > 0;

    p.pedidos.sort((a,b) => String(b.dataServico || '').localeCompare(String(a.dataServico || '')));
    p.ultimoPedido = p.pedidos.length ? p.pedidos[0] : null;
    p.alertaPedido = calcularAlertaPedido_(p.ultimoPedido, p.temLiberacao);

    delete p.pdes;
    delete p.plvs;
    delete p.ombs;

    return p;
  });

  projetos.sort((a,b) =>
    b.ultimoMovimento.localeCompare(a.ultimoMovimento) ||
    b.projeto.localeCompare(a.projeto)
  );

  return {
    resumo:{
      total:projetos.length,
      comPedido:projetos.filter(p => p.pedido).length,
      pedidosSemLiberacao:projetos.filter(p => p.pedido && !p.temLiberacao).length,
      alertasPedido:projetos.filter(p => p.alertaPedido && (p.alertaPedido.nivel === 'ALERTA' || p.alertaPedido.nivel === 'HOJE' || p.alertaPedido.nivel === 'ATRASADO')).length,
      documentosEnviados:projetos.filter(p => p.docFinal).length,
      documentosParciais:projetos.filter(p => p.docParcial && !p.docFinal).length,
      programadas:projetos.filter(p => p.programada).length,
      comBmd:projetos.filter(p => p.bmds > 0).length,
      comFfo:projetos.filter(p => p.ffos > 0).length
    },
    projetos:projetos,
    ultimaSync:PropertiesService.getScriptProperties()
      .getProperty('ULTIMA_SYNC_OK') || '',
    backfillConcluido:backfillConcluido_(),
    planilhaUrl:ss.getUrl(),
    logoUrl:HAGAP.LOGO_URL
  };
}


function getProjetoDetalhes(projeto) {
  const alvo = String(projeto || '').toUpperCase();

  return {
    projeto:alvo,
    eventos:lerEventos_(getSS_())
      .filter(e => e.projeto === alvo)
      .sort((a,b) => dataIso_(b.dataEmail).localeCompare(dataIso_(a.dataEmail)))
      .map(e => ({
        tipo:e.tipo,
        dataEmail:dataIso_(e.dataEmail),
        referencia:e.referencia,
        pdeRelacionado:e.pdeRelacionado,
        status:e.status,
        municipio:e.municipio,
        dataServico:e.dataServico,
        horaInicio:e.horaInicio,
        horaFim:e.horaFim,
        assunto:e.assunto,
        anexo:e.anexo,
        observacao:e.observacao,
        urlEmail:e.urlEmail
      }))
  };
}


function lerEventos_(ss) {
  const aba = ss.getSheetByName(HAGAP.ABAS.EVENTOS);
  if (aba.getLastRow() <= 1) return [];

  return aba.getRange(
    2,1,aba.getLastRow()-1,HAGAP.HEAD_EVENTOS.length
  ).getValues().map(r => ({
    chaveEvento:r[0],
    idEmail:r[1],
    dataEmail:r[2],
    tipo:String(r[3] || ''),
    projeto:String(r[4] || ''),
    projetoBase:String(r[5] || ''),
    referencia:String(r[6] || ''),
    pdeRelacionado:String(r[7] || ''),
    status:String(r[8] || ''),
    municipio:String(r[9] || ''),
    dataServico:String(r[10] || ''),
    horaInicio:String(r[11] || ''),
    horaFim:String(r[12] || ''),
    assunto:String(r[13] || ''),
    anexo:String(r[14] || ''),
    urlEmail:String(r[15] || ''),
    observacao:String(r[17] || '')
  }));
}


function calcularAlertaPedido_(pedido,temLiberacao) {
  if (!pedido) return {nivel:'SEM_PEDIDO',texto:'SEM PEDIDO'};
  if (temLiberacao) return {nivel:'OK',texto:'PDE/PLV RECEBIDO'};
  if (!pedido.dataServico) return {nivel:'PEDIDO',texto:'PEDIDO — DATA NÃO LOCALIZADA'};

  const alvo = parseDataBr_(pedido.dataServico);
  if (!alvo) return {nivel:'PEDIDO',texto:'PEDIDO — DATA NÃO LOCALIZADA'};

  const hoje = new Date();
  hoje.setHours(0,0,0,0);
  alvo.setHours(0,0,0,0);

  const dias = Math.round((alvo.getTime() - hoje.getTime()) / 86400000);

  if (dias < 0) return {nivel:'ATRASADO',texto:'DATA PASSOU — SEM PDE/PLV',dias:dias};
  if (dias === 0) return {nivel:'HOJE',texto:'HOJE — SEM PDE/PLV',dias:dias};
  if (dias <= 1) return {nivel:'ALERTA',texto:'FALTA 1 DIA — SEM PDE/PLV',dias:dias};
  return {nivel:'PEDIDO',texto:'PEDIDO',dias:dias};
}


function parseDataBr_(v) {
  const m = String(v || '').match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (!m) return null;
  return new Date(Number(m[3]),Number(m[2])-1,Number(m[1]));
}


function getSS_() {
  const id = PropertiesService.getScriptProperties()
    .getProperty('HAGAP_PLANILHA_ID');

  if (!id) {
    throw new Error('Execute configurarSistema() uma vez.');
  }

  return SpreadsheetApp.openById(id);
}


function garantirAba_(ss,nome,headers) {
  let aba = ss.getSheetByName(nome);
  if (!aba) aba = ss.insertSheet(nome);

  if (aba.getLastRow() === 0) {
    aba.getRange(1,1,1,headers.length).setValues([headers]);
    aba.setFrozenRows(1);
  } else {
    const atuais = aba.getRange(1,1,1,Math.max(aba.getLastColumn(),headers.length)).getValues()[0];
    const diferentes = headers.some((h,i) => String(atuais[i] || '') !== h);
    if (diferentes || aba.getLastColumn() < headers.length) {
      aba.getRange(1,1,1,headers.length).setValues([headers]);
      aba.setFrozenRows(1);
    }
  }

  return aba;
}


function log_(nivel,acao,detalhe) {
  try {
    const aba = getSS_().getSheetByName(HAGAP.ABAS.LOG);
    aba.appendRow([new Date(),nivel,acao,String(detalhe || '')]);
  } catch (e) {
    console.log(nivel + ' | ' + acao + ' | ' + detalhe);
  }
}


function somarStats_(dest,src) {
  dest.emailsNovos += Number(src.emailsNovos || 0);
  dest.eventosNovos += Number(src.eventosNovos || 0);
  dest.erros += Number(src.erros || 0);
}


function backfillConcluido_() {
  return PropertiesService.getScriptProperties()
    .getProperty('BACKFILL_CONCLUIDO') === '1';
}


function janelaMes_(mes) {
  const p = mes.split('-').map(Number);
  return {
    inicio:new Date(p[0],p[1]-1,1,0,0,0,0),
    fim:new Date(p[0],p[1],1,0,0,0,0)
  };
}


function addMes_(d,q) {
  return new Date(d.getFullYear(),d.getMonth()+q,1);
}


function fmtMes_(d) {
  return d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0');
}


function fmtQueryDate_(d) {
  return d.getFullYear() + '/' +
    String(d.getMonth()+1).padStart(2,'0') + '/' +
    String(d.getDate()).padStart(2,'0');
}


function dataIso_(v) {
  if (!v) return '';

  try {
    const d = v instanceof Date ? v : new Date(v);
    if (isNaN(d.getTime())) return String(v);

    return Utilities.formatDate(
      d,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd'T'HH:mm:ss"
    );
  } catch (e) {
    return String(v);
  }
}
