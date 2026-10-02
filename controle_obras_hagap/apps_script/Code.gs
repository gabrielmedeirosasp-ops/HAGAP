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
  SITE_HAGAP: 'https://hagap.onrender.com',
  API_DADOS_PC: 'https://hagap.onrender.com/api/dados',
  ABAS: {
    EVENTOS: 'EVENTOS',
    EMAILS: 'EMAILS_PROCESSADOS',
    BASE_PC: 'BASE_PC',
    AJUSTES: 'AJUSTES_MANUAIS',
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
  ],
  HEAD_PC: [
    'PROJETO','PROJETO_BASE','AES','PRAZO_AES','LOCAL','STATUS_PC','ARQUIVO_AES',
    'PASTA_PROJETO','PDFS_PROJETO_JSON','BMDS_JSON','FFOS_JSON','ORIGEM_JSON','ATUALIZADO'
  ],
  HEAD_AJUSTES: ['PROJETO','MUNICIPIO','PRAZO_AES','DATA_AJUSTE']
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
  garantirAba_(ss, HAGAP.ABAS.BASE_PC, HAGAP.HEAD_PC);
  garantirAba_(ss, HAGAP.ABAS.AJUSTES, HAGAP.HEAD_AJUSTES);
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


// Executar UMA VEZ manualmente quando a integração BASE_PC for ativada.
// A chamada fica fora do try/catch de sincronizarGmail_ para forçar o Google
// a solicitar a permissão de acesso externo (UrlFetchApp).
function autorizarBasePc() {
  const resp = UrlFetchApp.fetch(HAGAP.API_DADOS_PC, {
    method:'get',
    muteHttpExceptions:true,
    followRedirects:true
  });

  const code = resp.getResponseCode();
  if (code !== 200) {
    throw new Error('BASE_PC respondeu HTTP ' + code + '.');
  }

  const dados = JSON.parse(resp.getContentText('UTF-8'));
  if (!Array.isArray(dados)) {
    throw new Error('BASE_PC não retornou uma lista de obras.');
  }

  const r = sincronizarBasePc_(getSS_());
  log_('CONFIRMADO','AUTORIZAR_BASE_PC','Base externa autorizada e carregada: ' + (r.registros || 0) + ' registro(s).');
  return r;
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
      historico:0,
      basePc:0
    };

    // Base mestre: lê, sem alterar, os dados já consolidados pelo HAGAP antigo.
    // Isso traz AES/prazo, projetos do PC, BMD e FFO para o novo painel.
    if (Date.now() < prazo) {
      try {
        const pc = sincronizarBasePc_(ss);
        stats.basePc = pc.registros || 0;
      } catch (e) {
        stats.erros++;
        log_('PENDENTE','BASE_PC',String(e && e.message ? e.message : e));
      }
    }

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


function sincronizarBasePc_(ss) {
  const resp = UrlFetchApp.fetch(HAGAP.API_DADOS_PC, {
    method:'get',
    muteHttpExceptions:true,
    followRedirects:true
  });

  const code = resp.getResponseCode();
  if (code !== 200) {
    throw new Error('[PENDENTE] API HAGAP retornou HTTP ' + code);
  }

  const texto = resp.getContentText('UTF-8');
  const hash = Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, texto)
  );

  const props = PropertiesService.getScriptProperties();
  const aba = getSS_().getSheetByName(HAGAP.ABAS.BASE_PC);

  if (props.getProperty('BASE_PC_HASH') === hash && aba.getLastRow() > 1) {
    return {alterou:false,registros:aba.getLastRow()-1};
  }

  const dados = JSON.parse(texto);
  if (!Array.isArray(dados)) {
    throw new Error('[PENDENTE] /api/dados não retornou lista.');
  }

  const linhas = [];
  const agora = new Date();

  dados.forEach(r => {
    const projeto = String(r.projeto || '').trim().toUpperCase();
    // Mantém projetos COPEL e seus sufixos (I/C/S/II etc.) sem misturar.
    if (!/^\d{7}[A-Z]{0,3}$/.test(projeto)) return;

    const bmds = (Array.isArray(r.bmds) ? r.bmds : []).filter(x => {
      const nome = normalizar_((x && x.arquivo) || '');
      return nome.indexOf('BMD') >= 0 && nome.indexOf('MULTA') < 0;
    });

    const ffos = (Array.isArray(r.ffos) ? r.ffos : []).filter(x => {
      const nome = normalizar_((x && x.arquivo) || '');
      return nome.indexOf('FFO') >= 0 || nome.indexOf('FF0') >= 0;
    });

    linhas.push([
      projeto,
      projeto.substring(0,7),
      String(r.ae || ''),
      formatarDataBr_(r.prazo || ''),
      String(r.local || ''),
      String(r.status || ''),
      String(r.arquivo_ae || ''),
      String(r.pasta_projeto || ''),
      JSON.stringify(Array.isArray(r.pdfs_projeto) ? r.pdfs_projeto : []),
      JSON.stringify(bmds),
      JSON.stringify(ffos),
      JSON.stringify(Array.isArray(r.origem) ? r.origem : []),
      agora
    ]);
  });

  if (aba.getLastRow() > 1) {
    aba.getRange(2,1,aba.getLastRow()-1,aba.getLastColumn()).clearContent();
  }

  if (linhas.length) {
    aba.getRange(2,1,linhas.length,HAGAP.HEAD_PC.length).setValues(linhas);
  }

  props.setProperty('BASE_PC_HASH',hash);
  props.setProperty('BASE_PC_ATUALIZADA',agora.toISOString());

  return {alterou:true,registros:linhas.length};
}


function lerBasePc_(ss) {
  const aba = ss.getSheetByName(HAGAP.ABAS.BASE_PC);
  if (!aba || aba.getLastRow() <= 1) return [];

  return aba.getRange(2,1,aba.getLastRow()-1,HAGAP.HEAD_PC.length).getValues().map(r => ({
    projeto:String(r[0] || ''),
    projetoBase:String(r[1] || ''),
    aes:String(r[2] || ''),
    prazoAes:formatarDataBr_(r[3]),
    local:String(r[4] || ''),
    statusPc:String(r[5] || ''),
    arquivoAes:String(r[6] || ''),
    pastaProjeto:String(r[7] || ''),
    pdfsProjeto:jsonArraySeguro_(r[8]),
    bmdsPc:jsonArraySeguro_(r[9]),
    ffosPc:jsonArraySeguro_(r[10]),
    origemPc:jsonArraySeguro_(r[11]),
    atualizado:r[12]
  }));
}


function jsonArraySeguro_(v) {
  try {
    const x = JSON.parse(String(v || '[]'));
    return Array.isArray(x) ? x : [];
  } catch (e) {
    return [];
  }
}


function lerAjustes_(ss) {
  const aba = ss.getSheetByName(HAGAP.ABAS.AJUSTES);
  const mapa = {};
  if (!aba || aba.getLastRow() <= 1) return mapa;

  aba.getRange(2,1,aba.getLastRow()-1,HAGAP.HEAD_AJUSTES.length).getValues().forEach(r => {
    const projeto = String(r[0] || '').trim().toUpperCase();
    if (!projeto) return;
    mapa[projeto] = {
      municipio:String(r[1] || '').trim(),
      prazoAes:formatarDataBr_(r[2]),
      dataAjuste:r[3]
    };
  });

  return mapa;
}


function salvarAjusteProjeto(projeto,campo,valor) {
  const p = String(projeto || '').trim().toUpperCase();
  const c = String(campo || '').trim().toUpperCase();
  let v = String(valor == null ? '' : valor).trim();

  if (!/^\d{7}[A-Z]{0,3}$/.test(p)) {
    throw new Error('Projeto inválido.');
  }

  if (c !== 'MUNICIPIO' && c !== 'PRAZO_AES') {
    throw new Error('Campo de ajuste inválido.');
  }

  if (c === 'MUNICIPIO') {
    v = v.replace(/\s+/g,' ').trim().toUpperCase();
  }

  if (c === 'PRAZO_AES' && v) {
    if (!/^\d{2}\/\d{2}\/\d{4}$/.test(v) || !parseDataBr_(v)) {
      throw new Error('Prazo inválido. Use DD/MM/AAAA.');
    }
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const ss = getSS_();
    const aba = garantirAba_(ss,HAGAP.ABAS.AJUSTES,HAGAP.HEAD_AJUSTES);
    const qtd = Math.max(0,aba.getLastRow()-1);
    let linha = -1;
    let atual = ['', '', '', ''];

    if (qtd > 0) {
      const dados = aba.getRange(2,1,qtd,HAGAP.HEAD_AJUSTES.length).getValues();
      for (let i=0;i<dados.length;i++) {
        if (String(dados[i][0] || '').trim().toUpperCase() === p) {
          linha = i + 2;
          atual = dados[i].slice();
          break;
        }
      }
    }

    atual[0] = p;
    if (c === 'MUNICIPIO') atual[1] = v;
    if (c === 'PRAZO_AES') atual[2] = v;
    atual[3] = new Date();

    const semAjuste = !String(atual[1] || '').trim() && !String(atual[2] || '').trim();

    if (linha > 0 && semAjuste) {
      aba.deleteRow(linha);
    } else if (linha > 0) {
      aba.getRange(linha,1,1,HAGAP.HEAD_AJUSTES.length).setValues([atual]);
    } else if (!semAjuste) {
      aba.appendRow(atual);
    }

    log_('CONFIRMADO','AJUSTE_MANUAL',p + ' | ' + c + ' = ' + (v || '[REVERTER BASE]'));
    return {ok:true,projeto:p,campo:c,valor:v};

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
  const basePc = lerBasePc_(ss);
  const ajustes = lerAjustes_(ss);
  const mapa = {};

  function novoProjeto_(projeto, projetoBase) {
    return {
      projeto:projeto,
      projetoBase:projetoBase || projeto.substring(0,7),
      municipio:'',
      municipioFonte:'',
      aes:'',
      prazoAes:'',
      prazoAesFonte:'',
      ajusteMunicipio:false,
      ajustePrazo:false,
      statusPc:'',
      arquivoAes:'',
      pastaProjeto:'',
      pdfsProjeto:[],
      bmdsPc:[],
      ffosPc:[],
      origemPc:[],
      docFinal:false,
      docParcial:false,
      docUrl:'',
      docParcialUrl:'',
      pedido:false,
      pedidos:[],
      pdes:{},plvs:{},ombs:{},
      bmdLinks:[],
      ffoLinks:[],
      datasServico:[],
      ultimoMovimento:''
    };
  }

  // A base do PC/AES é a lista mestre. O Gmail entra por cima.
  basePc.forEach(r => {
    if (!r.projeto) return;
    const p = novoProjeto_(r.projeto,r.projetoBase);
    p.municipio = r.local || '';
    p.municipioFonte = r.local || '';
    p.aes = r.aes || '';
    p.prazoAes = r.prazoAes || '';
    p.prazoAesFonte = r.prazoAes || '';
    p.statusPc = r.statusPc || '';
    p.arquivoAes = r.arquivoAes || '';
    p.pastaProjeto = r.pastaProjeto || '';
    p.pdfsProjeto = r.pdfsProjeto || [];
    p.bmdsPc = r.bmdsPc || [];
    p.ffosPc = r.ffosPc || [];
    p.origemPc = r.origemPc || [];
    mapa[r.projeto] = p;
  });

  eventos.forEach(e => {
    if (!e.projeto) return;

    if (!mapa[e.projeto]) {
      mapa[e.projeto] = novoProjeto_(e.projeto,e.projetoBase);
    }

    const p = mapa[e.projeto];

    if (e.municipio && !p.municipio) p.municipio = e.municipio;
    if (e.dataServico) p.datasServico.push(e.dataServico);

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

    if (e.tipo === 'DOC_FINAL') {
      p.docFinal = true;
      if (e.urlEmail) p.docUrl = e.urlEmail;
    }

    if (e.tipo === 'DOC_PARCIAL') {
      p.docParcial = true;
      if (e.urlEmail) p.docParcialUrl = e.urlEmail;
    }

    if (e.tipo === 'PDE' && e.referencia) {
      p.pdes[e.referencia] = {
        referencia:e.referencia,status:e.status,urlEmail:e.urlEmail,dataServico:e.dataServico
      };
    }

    if (e.tipo === 'PLV' && e.referencia) {
      p.plvs[e.referencia] = {
        referencia:e.referencia,status:e.status,urlEmail:e.urlEmail,dataServico:e.dataServico
      };
    }

    if (e.tipo === 'OMB' && e.referencia) {
      p.ombs[e.referencia] = {
        referencia:e.referencia,status:e.status,urlEmail:e.urlEmail,
        dataServico:e.dataServico,pdeRelacionado:e.pdeRelacionado
      };
    }

    if (e.tipo === 'BMD') {
      p.bmdLinks.push({status:e.status,urlEmail:e.urlEmail,anexo:e.anexo});
    }

    if (e.tipo === 'FFO') {
      p.ffoLinks.push({status:e.status,urlEmail:e.urlEmail,anexo:e.anexo});
    }

    const iso = dataIso_(e.dataEmail);
    if (iso > p.ultimoMovimento) p.ultimoMovimento = iso;
  });

  const projetos = Object.keys(mapa).map(k => {
    const p = mapa[k];

    p.qtdPde = Object.keys(p.pdes).length;
    p.qtdPlv = Object.keys(p.plvs).length;
    p.qtdOmb = Object.keys(p.ombs).length;
    p.temLiberacao = p.qtdPde > 0 || p.qtdPlv > 0;

    p.pedidos.sort((a,b) =>
      dataBrParaChave_(b.dataServico).localeCompare(dataBrParaChave_(a.dataServico))
    );
    p.ultimoPedido = p.pedidos.length ? p.pedidos[0] : null;
    p.alertaPedido = calcularAlertaPedido_(p.ultimoPedido,p.temLiberacao);

    p.linksPde = Object.values(p.pdes);
    p.linksPlv = Object.values(p.plvs);
    p.linksOmb = Object.values(p.ombs);

    p.dataSolicitada = p.ultimoPedido && p.ultimoPedido.dataServico
      ? p.ultimoPedido.dataServico
      : escolherDataPrincipal_(p.datasServico);

    p.servicoResumo = p.ultimoPedido && p.ultimoPedido.observacao
      ? p.ultimoPedido.observacao
      : '';

    const ajuste = ajustes[p.projeto] || {};
    if (ajuste.municipio) {
      p.municipio = ajuste.municipio;
      p.ajusteMunicipio = true;
    }
    if (ajuste.prazoAes) {
      p.prazoAes = ajuste.prazoAes;
      p.ajustePrazo = true;
    }

    p.temProjetoPdf = p.pdfsProjeto.length > 0;
    p.temBmd = p.bmdsPc.length > 0 || p.bmdLinks.length > 0;
    p.temFfo = p.ffosPc.length > 0 || p.ffoLinks.length > 0;
    p.temAes = !!(p.aes || p.prazoAes || p.arquivoAes);

    p.aesUrl = p.arquivoAes
      ? HAGAP.SITE_HAGAP + '/pdfs/aes/' + encodeURIComponent(p.arquivoAes)
      : '';

    p.projetoUrls = p.pdfsProjeto.map(nome => ({
      nome:nome,
      url:HAGAP.SITE_HAGAP + '/pdfs/projetos/' +
        encodeURIComponent(p.pastaProjeto || p.projeto) + '/' +
        String(nome).split('/').map(encodeURIComponent).join('/')
    }));

    p.bmdsPc = p.bmdsPc.map(x => ({
      arquivo:String((x && x.arquivo) || ''),
      status:String((x && x.status) || ''),
      data:String((x && x.data) || ''),
      url:(x && x.arquivo)
        ? HAGAP.SITE_HAGAP + '/pdfs/medicoes/' + encodeURIComponent(x.arquivo)
        : ''
    }));

    p.ffosPc = p.ffosPc.map(x => ({
      arquivo:String((x && x.arquivo) || ''),
      status:String((x && x.status) || ''),
      data:String((x && x.data) || ''),
      url:(x && x.arquivo)
        ? HAGAP.SITE_HAGAP + '/pdfs/medicoes/' + encodeURIComponent(x.arquivo)
        : ''
    }));

    delete p.pdes;
    delete p.plvs;
    delete p.ombs;
    delete p.datasServico;

    return p;
  });

  // ORDEM PRINCIPAL: vencimento da AES.
  // Sem AES fica no fim, claramente separado.
  projetos.sort((a,b) => {
    const pa = dataBrParaChave_(a.prazoAes);
    const pb = dataBrParaChave_(b.prazoAes);

    if (pa && pb && pa !== pb) return pa.localeCompare(pb);
    if (pa && !pb) return -1;
    if (!pa && pb) return 1;

    const da = dataBrParaChave_(a.dataSolicitada);
    const db = dataBrParaChave_(b.dataSolicitada);
    if (da && db && da !== db) return da.localeCompare(db);
    if (da && !db) return -1;
    if (!da && db) return 1;

    return a.projeto.localeCompare(b.projeto);
  });

  return {
    resumo:{
      total:projetos.length,
      comAes:projetos.filter(p => p.temAes).length,
      comProjetoPdf:projetos.filter(p => p.temProjetoPdf).length,
      comPedido:projetos.filter(p => p.pedido).length,
      pedidosSemLiberacao:projetos.filter(p => p.pedido && !p.temLiberacao).length,
      alertasPedido:projetos.filter(p =>
        p.alertaPedido && ['ALERTA','HOJE','ATRASADO'].indexOf(p.alertaPedido.nivel) >= 0
      ).length,
      documentosEnviados:projetos.filter(p => p.docFinal).length,
      documentosParciais:projetos.filter(p => p.docParcial && !p.docFinal).length,
      comBmd:projetos.filter(p => p.temBmd).length,
      comFfo:projetos.filter(p => p.temFfo).length
    },
    projetos:projetos,
    ultimaSync:PropertiesService.getScriptProperties().getProperty('ULTIMA_SYNC_OK') || '',
    basePcAtualizada:PropertiesService.getScriptProperties().getProperty('BASE_PC_ATUALIZADA') || '',
    backfillConcluido:backfillConcluido_(),
    planilhaUrl:ss.getUrl(),
    logoUrl:HAGAP.LOGO_URL
  };
}


function getProjetoDetalhes(projeto) {
  const alvo = String(projeto || '').toUpperCase();
  const ss = getSS_();
  const pc = lerBasePc_(ss).find(x => x.projeto === alvo) || null;
  const ajuste = lerAjustes_(ss)[alvo] || {};

  const pcDetalhes = pc ? {
    aes:pc.aes,
    prazoAes:ajuste.prazoAes || pc.prazoAes,
    prazoAesFonte:pc.prazoAes,
    ajustePrazo:!!ajuste.prazoAes,
    local:ajuste.municipio || pc.local,
    localFonte:pc.local,
    ajusteMunicipio:!!ajuste.municipio,
    statusPc:pc.statusPc,
    arquivoAes:pc.arquivoAes,
    aesUrl:pc.arquivoAes
      ? HAGAP.SITE_HAGAP + '/pdfs/aes/' + encodeURIComponent(pc.arquivoAes)
      : '',
    projetos:(pc.pdfsProjeto || []).map(nome => ({
      nome:nome,
      url:HAGAP.SITE_HAGAP + '/pdfs/projetos/' +
        encodeURIComponent(pc.pastaProjeto || alvo) + '/' +
        String(nome).split('/').map(encodeURIComponent).join('/')
    })),
    bmds:(pc.bmdsPc || []).map(x => ({
      arquivo:String((x && x.arquivo) || ''),
      url:(x && x.arquivo)
        ? HAGAP.SITE_HAGAP + '/pdfs/medicoes/' + encodeURIComponent(x.arquivo)
        : ''
    })),
    ffos:(pc.ffosPc || []).map(x => ({
      arquivo:String((x && x.arquivo) || ''),
      url:(x && x.arquivo)
        ? HAGAP.SITE_HAGAP + '/pdfs/medicoes/' + encodeURIComponent(x.arquivo)
        : ''
    }))
  } : null;

  return {
    projeto:alvo,
    pc:pcDetalhes,
    eventos:lerEventos_(ss)
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
    dataServico:formatarDataBr_(r[10]),
    horaInicio:formatarHora_(r[11]),
    horaFim:formatarHora_(r[12]),
    assunto:String(r[13] || ''),
    anexo:String(r[14] || ''),
    urlEmail:String(r[15] || ''),
    observacao:String(r[17] || '')
  }));
}


function formatarHora_(v) {
  if (!v) return '';

  if (v instanceof Date && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'HH:mm');
  }

  const s = String(v || '').trim();

  // Já está no formato correto.
  const direto = s.match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);
  if (direto) {
    return String(Number(direto[1])).padStart(2,'0') + ':' + direto[2];
  }

  // Corrige valores de horário que o Google Sheets converte para
  // datas-base de 1899/1900 (ex.: Sat Dec 30 1899 10:00:00...).
  const d = new Date(s);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'HH:mm');
  }

  return s;
}


function formatarDataBr_(v) {
  if (!v) return '';

  if (v instanceof Date && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }

  const s = String(v || '').trim();

  if (/^\d{2}\/\d{2}\/\d{4}$/.test(s)) return s;

  const d = new Date(s);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }

  return s;
}


function dataBrParaChave_(v) {
  const m = String(v || '').match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  return m ? (m[3] + '-' + m[2] + '-' + m[1]) : '';
}


function escolherDataPrincipal_(datas) {
  const validas = (datas || [])
    .map(formatarDataBr_)
    .filter(d => /^\d{2}\/\d{2}\/\d{4}$/.test(d));

  if (!validas.length) return '';

  const hoje = new Date();
  hoje.setHours(0,0,0,0);

  const itens = validas.map(d => ({texto:d,data:parseDataBr_(d)}))
    .filter(x => x.data);

  const futuras = itens.filter(x => x.data >= hoje)
    .sort((a,b) => a.data - b.data);

  if (futuras.length) return futuras[0].texto;

  itens.sort((a,b) => b.data - a.data);
  return itens.length ? itens[0].texto : '';
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

  const dia = Number(m[1]);
  const mes = Number(m[2]);
  const ano = Number(m[3]);
  const d = new Date(ano,mes-1,dia);

  if (
    d.getFullYear() !== ano ||
    d.getMonth() !== mes-1 ||
    d.getDate() !== dia
  ) {
    return null;
  }

  return d;
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
