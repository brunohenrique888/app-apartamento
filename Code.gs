/**
 * App do Apartamento — Despesas, Contas fixas e Vencimentos/garantias
 * Fica no projeto do Apps Script da planilha do apartamento (Extensões → Apps Script).
 * Separado do app de lançamentos dos pais: não lê nem grava em nenhuma outra planilha.
 * Acesso por link pessoal (?k=CHAVE), configurado na aba APP_ACESSOS. Quem está lá vê e lança tudo.
 * Primeira vez: rode configurar() pelo editor (cria as abas, as chaves e os links).
 */
const ID_PLANILHA = ''; // vazio = a planilha a que este script está ligado
const ABA_DESP = 'DESPESAS';
const ABA_FIXAS = 'CONTAS FIXAS';
const ABA_VENC = 'VENCIMENTOS';
const ABA_ACESSOS = 'APP_ACESSOS';
const PASTAS = {
  DESP: 'Apartamento - Anexos de despesas',
  VENC: 'Apartamento - Anexos de vencimentos',
  BACKUP: 'Apartamento - Backups'
};

// Colunas (0 = A) de cada aba
const CAB_DESP = ['CARIMBO', 'VENCIMENTO', 'CATEGORIA', 'VALOR', 'DESCRIÇÃO', 'ANEXO', 'DATA PAGAMENTO', 'LANÇADO POR'];
const COL_D = { CARIMBO: 0, VENC: 1, CAT: 2, VALOR: 3, DESC: 4, ANEXO: 5, PGTO: 6, USUARIO: 7 };
const CAB_FIXAS = ['ID', 'DESCRIÇÃO', 'CATEGORIA', 'DIA DO VENCIMENTO', 'VALOR PREVISTO', 'SITUAÇÃO', 'CRIADO POR'];
const COL_F = { ID: 0, DESC: 1, CAT: 2, DIA: 3, VALOR: 4, SITUACAO: 5, USUARIO: 6 };
const CAB_VENC = ['ID', 'CRIADO EM', 'TIPO', 'DESCRIÇÃO', 'FORNECEDOR / LOJA', 'INÍCIO', 'VENCE', 'REPETIR (MESES)',
  'ANEXO', 'SITUAÇÃO', 'CONCLUÍDO EM', 'OBSERVAÇÃO', 'ANTERIOR', 'CRIADO POR'];
const COL_V = { ID: 0, CRIADO: 1, TIPO: 2, DESC: 3, FORN: 4, INI: 5, VENCE: 6, REP: 7,
  ANEXO: 8, SITUACAO: 9, CONCLUIDO: 10, OBS: 11, ANTERIOR: 12, USUARIO: 13 };
// Tipos fixos de vencimento; o app também deixa criar outros ("+ novo")
const TIPOS_VENC = ['Garantia', 'Seguro'];

function doGet(e) {
  const chave = (e && e.parameter && e.parameter.k) || '';
  try {
    usuario_(chave);
  } catch (err) {
    return HtmlService.createHtmlOutput('<p style="font:16px system-ui;padding:24px">Acesso negado. Use o link pessoal que você recebeu.</p>')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  const t = HtmlService.createTemplateFromFile('Index');
  t.chave = chave;
  return t.evaluate()
    .setTitle('Apartamento')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
}

// =====================================================================
// Configuração e acessos
// =====================================================================

/**
 * Rode pelo editor na 1ª vez (e sempre que incluir alguém em APP_ACESSOS):
 * cria as abas que faltam e gera chave e link pessoal para quem ainda não tem. Não apaga nada.
 */
function configurar() {
  const ss = ss_();
  aba_(ABA_DESP, CAB_DESP, sh => {
    sh.getRange('A2:A').setNumberFormat('dd/MM/yyyy HH:mm:ss');
    sh.getRange('B2:B').setNumberFormat('dd/MM/yyyy');
    sh.getRange('D2:D').setNumberFormat('"R$" #,##0.00');
    sh.getRange('G2:G').setNumberFormat('dd/MM/yyyy');
  });
  aba_(ABA_FIXAS, CAB_FIXAS, sh => sh.getRange('E2:E').setNumberFormat('"R$" #,##0.00'));
  aba_(ABA_VENC, CAB_VENC, sh => {
    ['B', 'F', 'G', 'K'].forEach(c => sh.getRange(c + '2:' + c).setNumberFormat('dd/MM/yyyy'));
  });

  let sh = ss.getSheetByName(ABA_ACESSOS);
  if (!sh) {
    sh = ss.insertSheet(ABA_ACESSOS);
    sh.getRange('A1:C1').setValues([['NOME', 'CHAVE (não compartilhe)', 'LINK PESSOAL']]).setFontWeight('bold');
    sh.getRange('A2:A3').setValues([['BRUNO'], ['PESSOA 2']]);
    sh.getRange('E1').setValue('URL do app (/exec) — cole abaixo:').setFontWeight('bold');
    sh.setColumnWidth(2, 280); sh.setColumnWidth(3, 420); sh.setColumnWidth(5, 420);
  }
  for (let r = 2; r <= sh.getLastRow(); r++) {
    if (!String(sh.getRange(r, 1).getValue()).trim()) continue;
    if (!String(sh.getRange(r, 2).getValue()).trim()) {
      sh.getRange(r, 2).setValue(Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, ''));
    }
    // Fórmula sem separador de argumentos, para funcionar em qualquer idioma da planilha
    sh.getRange(r, 3).setFormula('=$E$2&"?k="&B' + r);
  }
  // A aba "Página1" que vem com a planilha nova sai, se estiver vazia
  const p1 = ss.getSheets().find(s => /^(P[áa]gina|Sheet)1$/.test(s.getName()));
  if (p1 && p1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(p1);
  SpreadsheetApp.flush();
  Logger.log('Planilha: ' + ss.getName() + ' — ' + ss.getUrl());
  Logger.log('Abas agora: ' + ss.getSheets().map(s => s.getName()).join(' | '));
}

/** Abre a aba (ou cria com cabeçalho e formatos, chamando formatar(sh) só na criação). */
function aba_(nome, cab, formatar) {
  const ss = ss_();
  let sh = ss.getSheetByName(nome);
  if (!sh) {
    sh = ss.insertSheet(nome);
    sh.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold').setBackground('#e5e7eb');
    sh.setFrozenRows(1);
    if (formatar) formatar(sh);
  }
  return sh;
}

function usuario_(chave) {
  chave = String(chave || '').trim();
  if (chave.length < 32) throw new Error('Acesso negado.');
  const sh = ss_().getSheetByName(ABA_ACESSOS);
  if (!sh || sh.getLastRow() < 2) throw new Error('Acesso negado.');
  const u = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues()
    .find(l => String(l[1]).trim() === chave && String(l[0]).trim());
  if (!u) throw new Error('Acesso negado.');
  return { nome: String(u[0]).trim().toUpperCase() };
}

// =====================================================================
// Leitura
// =====================================================================

function getTudo(chave) {
  const u = usuario_(chave);
  return { usuario: u, hoje: fmt_(new Date(), tz_()), desp: lerDesp_(), fixas: lerFixas_(), venc: lerVenc_() };
}

function lerDesp_() {
  const sh = ss_().getSheetByName(ABA_DESP);
  if (!sh || sh.getLastRow() < 2) return [];
  const tz = tz_();
  return sh.getRange(2, 1, sh.getLastRow() - 1, CAB_DESP.length).getValues().map((r, i) => {
    const venc = parseData_(r[COL_D.VENC]);
    const desc = String(r[COL_D.DESC] || '').trim();
    if (!venc && !desc) return null;
    const pgto = parseData_(r[COL_D.PGTO]);
    return {
      row: i + 2,
      venc: venc ? fmt_(venc, tz) : '',
      pgto: pgto ? fmt_(pgto, tz) : '',
      cat: String(r[COL_D.CAT] || '').trim(),
      valor: parseNum_(r[COL_D.VALOR]),
      desc: desc,
      anexo: String(r[COL_D.ANEXO] || '').trim(),
      usuario: String(r[COL_D.USUARIO] || '').trim()
    };
  }).filter(Boolean);
}

function lerFixas_() {
  const sh = ss_().getSheetByName(ABA_FIXAS);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, CAB_FIXAS.length).getValues().map((r, i) => ({
    row: i + 2,
    id: String(r[COL_F.ID] || '').trim(),
    desc: String(r[COL_F.DESC] || '').trim().toUpperCase(),
    cat: String(r[COL_F.CAT] || '').trim().toUpperCase(),
    dia: Math.min(31, Math.max(1, parseInt(r[COL_F.DIA], 10) || 1)),
    valor: parseNum_(r[COL_F.VALOR]),
    ativa: normTxt_(r[COL_F.SITUACAO]) !== 'INATIVA'
  })).filter(f => f.desc);
}

function lerVenc_() {
  const sh = ss_().getSheetByName(ABA_VENC);
  if (!sh || sh.getLastRow() < 2) return [];
  const tz = tz_();
  const d = v => { const x = parseData_(v); return x ? fmt_(x, tz) : ''; };
  const t = v => String(v == null ? '' : v).trim();
  return sh.getRange(2, 1, sh.getLastRow() - 1, CAB_VENC.length).getValues().map((l, i) => ({
    row: i + 2, id: t(l[COL_V.ID]), tipo: tipoVenc_(l[COL_V.TIPO]), desc: t(l[COL_V.DESC]), fornecedor: t(l[COL_V.FORN]),
    iniData: d(l[COL_V.INI]), venceData: d(l[COL_V.VENCE]), repMeses: parseNum_(l[COL_V.REP]) > 0 ? parseNum_(l[COL_V.REP]) : '',
    anexo: t(l[COL_V.ANEXO]), situacao: normTxt_(l[COL_V.SITUACAO]) || 'ATIVO', concluido: d(l[COL_V.CONCLUIDO]),
    obs: t(l[COL_V.OBS]), anterior: t(l[COL_V.ANTERIOR]), usuario: t(l[COL_V.USUARIO])
  })).filter(x => x.id && (x.tipo || x.desc));
}

// =====================================================================
// Despesas
// =====================================================================

/**
 * itens: [{ venc:'yyyy-MM-dd', pgto:'yyyy-MM-dd'|'', cat, valor:Number, desc,
 *           parcelas:Number (opcional), valorTotal:Boolean (opcional, valor = total a dividir),
 *           anexos:[{name,mime,data(base64)}] (opcional) }]
 */
function despSalvar(chave, itens) {
  const u = usuario_(chave);
  comLock_(() => {
    const sh = aba_(ABA_DESP, CAB_DESP);
    const tz = tz_();
    const agora = new Date();
    const linhas = [];

    itens.forEach(it => {
      if (!it.venc || !it.desc || !it.cat || !(Number(it.valor) > 0)) throw new Error('Preencha descrição, valor, categoria e vencimento.');
      const url = salvarAnexos_(PASTAS.DESP, it);
      const n = Math.max(1, Math.min(60, parseInt(it.parcelas, 10) || 1));
      // valorTotal: divide em centavos e joga a diferença do arredondamento na última parcela.
      const centavos = Math.round(Number(it.valor) * 100);
      const base = it.valorTotal ? Math.floor(centavos / n) : centavos;
      for (let p = 0; p < n; p++) {
        const desc = String(it.desc).trim().toUpperCase();
        const linha = new Array(CAB_DESP.length).fill('');
        linha[COL_D.CARIMBO] = agora;
        linha[COL_D.VENC] = Utilities.parseDate(addMeses_(it.venc, p), tz, 'yyyy-MM-dd');
        linha[COL_D.CAT] = String(it.cat).trim().toUpperCase();
        linha[COL_D.VALOR] = (it.valorTotal && p === n - 1 ? centavos - base * (n - 1) : base) / 100;
        linha[COL_D.DESC] = n > 1 ? desc + ' (' + (p + 1) + '/' + n + ')' : desc;
        linha[COL_D.ANEXO] = url;
        // Só a 1ª parcela herda a data de pagamento; as futuras ficam "a pagar".
        linha[COL_D.PGTO] = p === 0 && it.pgto ? Utilities.parseDate(it.pgto, tz, 'yyyy-MM-dd') : '';
        linha[COL_D.USUARIO] = u.nome;
        linhas.push(linha);
      }
    });

    const inicio = sh.getLastRow() + 1, n = linhas.length;
    sh.getRange(inicio, 1, n, CAB_DESP.length).setValues(linhas);
    sh.getRange(inicio, COL_D.CARIMBO + 1, n, 1).setNumberFormat('dd/MM/yyyy HH:mm:ss');
    sh.getRange(inicio, COL_D.VENC + 1, n, 1).setNumberFormat('dd/MM/yyyy');
    sh.getRange(inicio, COL_D.PGTO + 1, n, 1).setNumberFormat('dd/MM/yyyy');
    sh.getRange(inicio, COL_D.VALOR + 1, n, 1).setNumberFormat('"R$" #,##0.00');
  });
  return { desp: lerDesp_() };
}

function despMarcarPago(chave, row, desc, data) {
  usuario_(chave);
  comLock_(() => {
    const sh = ss_().getSheetByName(ABA_DESP);
    linhaConferida_(sh, row, CAB_DESP.length, r => String(r[COL_D.DESC] || '').trim() === String(desc || '').trim());
    sh.getRange(row, COL_D.PGTO + 1).setValue(Utilities.parseDate(data, tz_(), 'yyyy-MM-dd')).setNumberFormat('dd/MM/yyyy');
  });
  return { desp: lerDesp_() };
}

function despExcluir(chave, row, desc) {
  usuario_(chave);
  comLock_(() => {
    const sh = ss_().getSheetByName(ABA_DESP);
    linhaConferida_(sh, row, CAB_DESP.length, r => String(r[COL_D.DESC] || '').trim() === String(desc || '').trim());
    sh.deleteRow(row);
  });
  return { desp: lerDesp_() };
}

/**
 * Corrige um lançamento já gravado (uma linha; numa conta parcelada, só aquela parcela).
 * antes = { desc } como estava na tela (confere a linha).
 * novo = { venc:'yyyy-MM-dd', cat, valor, desc, pgto:'yyyy-MM-dd'|'' }.
 * Não mexe no carimbo, no anexo nem em LANÇADO POR.
 */
function despEditar(chave, row, antes, novo) {
  usuario_(chave);
  const cat = String(novo.cat || '').trim().toUpperCase();
  const desc = String(novo.desc || '').trim().toUpperCase();
  if (!novo.venc || !desc || !cat || !(Number(novo.valor) > 0)) throw new Error('Preencha descrição, valor, categoria e vencimento.');
  comLock_(() => {
    const sh = ss_().getSheetByName(ABA_DESP);
    const tz = tz_();
    linhaConferida_(sh, row, CAB_DESP.length, r => String(r[COL_D.DESC] || '').trim() === String(antes.desc || '').trim());
    const valor = Math.round(Number(novo.valor) * 100) / 100;
    // B:E = vencimento, categoria, valor, descrição
    const be = [Utilities.parseDate(novo.venc, tz, 'yyyy-MM-dd'), cat, valor, desc];
    semFiltro_(sh, () => {
      sh.getRange(row, COL_D.VENC + 1, 1, 4).setValues([be]);
      sh.getRange(row, COL_D.PGTO + 1).setValue(novo.pgto ? Utilities.parseDate(novo.pgto, tz, 'yyyy-MM-dd') : '');
    });
    sh.getRange(row, COL_D.VENC + 1).setNumberFormat('dd/MM/yyyy');
    sh.getRange(row, COL_D.PGTO + 1).setNumberFormat('dd/MM/yyyy');
    sh.getRange(row, COL_D.VALOR + 1).setNumberFormat('"R$" #,##0.00');
    conferirEdicao_(sh, row, CAB_DESP.length, l =>
      String(l[COL_D.DESC]).trim() === desc && String(l[COL_D.CAT]).trim().toUpperCase() === cat
      && Math.abs(parseNum_(l[COL_D.VALOR]) - valor) < 0.005);
  });
  return { desp: lerDesp_() };
}

function despAnexar(chave, row, desc, anexo) {
  usuario_(chave);
  acrescentarAnexo_(ABA_DESP, PASTAS.DESP, COL_D.ANEXO, anexo, sh => {
    linhaConferida_(sh, row, CAB_DESP.length, r => String(r[COL_D.DESC] || '').trim() === String(desc || '').trim());
    return row;
  });
  return { desp: lerDesp_() };
}

// =====================================================================
// Contas fixas (cadastro). O lançamento do mês usa despSalvar.
// =====================================================================

/** it: { desc, cat, dia, valor } */
function fixaSalvar(chave, it) {
  const u = usuario_(chave);
  const desc = String(it.desc || '').trim().toUpperCase(), cat = String(it.cat || '').trim().toUpperCase();
  const dia = parseInt(it.dia, 10);
  if (!desc || !cat || !(dia >= 1 && dia <= 31)) throw new Error('Preencha descrição, categoria e dia do vencimento (1 a 31).');
  comLock_(() => {
    const sh = aba_(ABA_FIXAS, CAB_FIXAS, s => s.getRange('E2:E').setNumberFormat('"R$" #,##0.00'));
    if (lerFixas_().some(f => f.ativa && f.desc === desc)) throw new Error('Já existe a conta fixa ' + desc + '.');
    const l = new Array(CAB_FIXAS.length).fill('');
    l[COL_F.ID] = proximoId_(sh, 'F');
    l[COL_F.DESC] = desc; l[COL_F.CAT] = cat; l[COL_F.DIA] = dia;
    l[COL_F.VALOR] = Number(it.valor) > 0 ? Math.round(Number(it.valor) * 100) / 100 : '';
    l[COL_F.SITUACAO] = 'ATIVA'; l[COL_F.USUARIO] = u.nome;
    sh.getRange(sh.getLastRow() + 1, 1, 1, l.length).setValues([l]);
  });
  return { fixas: lerFixas_() };
}

/** Tira a conta da lista (fica na planilha como INATIVA; para voltar, apague o INATIVA na aba). */
function fixaDesativar(chave, row, id) {
  usuario_(chave);
  comLock_(() => {
    const sh = ss_().getSheetByName(ABA_FIXAS);
    linhaConferida_(sh, row, CAB_FIXAS.length, r => String(r[COL_F.ID]).trim() === String(id || '').trim());
    sh.getRange(row, COL_F.SITUACAO + 1).setValue('INATIVA');
  });
  return { fixas: lerFixas_() };
}

// =====================================================================
// Vencimentos e garantias (status calculado na hora; a planilha guarda só SITUAÇÃO)
// =====================================================================

// Garantia/Seguro com qualquer grafia vira o nome fixo; outro tipo fica como foi escrito (1ª letra maiúscula)
function tipoVenc_(tipo) {
  const n = normTxt_(tipo);
  if (!n) return '';
  const fixo = TIPOS_VENC.find(t => normTxt_(t) === n || n.indexOf(normTxt_(t)) === 0);
  if (fixo) return fixo;
  const t = String(tipo).trim().replace(/\s+/g, ' ');
  return t.charAt(0).toUpperCase() + t.slice(1);
}

/**
 * it = { tipo, desc, fornecedor, iniData, venceData, repMeses, anexo, renovar: { row, id } }.
 * Com "renovar", o item antigo vira CONCLUÍDO (na data de início do novo) e o novo guarda o ID dele em ANTERIOR.
 */
function vencSalvar(chave, it) {
  const u = usuario_(chave);
  const tipo = tipoVenc_(it.tipo);
  if (!tipo) throw new Error('Escolha o tipo.');
  if (!it.venceData) throw new Error('Informe quando vence.');
  const anexoUrl = it.anexo && it.anexo.data ? salvarAnexo_(PASTAS.VENC, it.anexo) : '';
  comLock_(() => {
    const sh = aba_(ABA_VENC, CAB_VENC);
    const tz = tz_();
    const data = iso => iso ? Utilities.parseDate(iso, tz, 'yyyy-MM-dd') : '';
    const id = proximoId_(sh, 'V');

    let antigo = null;
    if (it.renovar && it.renovar.row) {
      const r = linhaConferida_(sh, it.renovar.row, CAB_VENC.length, r =>
        String(r[COL_V.ID]).trim() === String(it.renovar.id || '').trim() && normTxt_(r[COL_V.SITUACAO]) === 'ATIVO');
      antigo = { row: it.renovar.row, obs: String(r[COL_V.OBS] || '').trim(), id: String(r[COL_V.ID]).trim() };
    }

    const l = new Array(CAB_VENC.length).fill('');
    l[COL_V.ID] = id; l[COL_V.CRIADO] = new Date(); l[COL_V.TIPO] = tipo;
    l[COL_V.DESC] = String(it.desc || '').trim(); l[COL_V.FORN] = String(it.fornecedor || '').trim().toUpperCase();
    l[COL_V.INI] = data(it.iniData); l[COL_V.VENCE] = data(it.venceData);
    l[COL_V.REP] = parseInt(it.repMeses, 10) > 0 ? parseInt(it.repMeses, 10) : '';
    l[COL_V.ANEXO] = anexoUrl; l[COL_V.SITUACAO] = 'ATIVO';
    l[COL_V.ANTERIOR] = antigo ? antigo.id : ''; l[COL_V.USUARIO] = u.nome;
    const nova = sh.getLastRow() + 1;
    sh.getRange(nova, 1, 1, l.length).setValues([l]);
    [COL_V.CRIADO, COL_V.INI, COL_V.VENCE].forEach(c => sh.getRange(nova, c + 1).setNumberFormat('dd/MM/yyyy'));

    if (antigo) {
      sh.getRange(antigo.row, COL_V.SITUACAO + 1).setValue('CONCLUÍDO');
      sh.getRange(antigo.row, COL_V.CONCLUIDO + 1).setValue(data(it.iniData) || new Date()).setNumberFormat('dd/MM/yyyy');
      sh.getRange(antigo.row, COL_V.OBS + 1).setValue((antigo.obs ? antigo.obs + ' · ' : '') + 'Renovado: ' + id);
    }
  });
  return { venc: lerVenc_() };
}

/** Conclui (ou descarta, com cancelar = true) um item ativo, sem criar o próximo. */
function vencConcluir(chave, row, id, data, obs, cancelar) {
  usuario_(chave);
  comLock_(() => {
    const sh = ss_().getSheetByName(ABA_VENC);
    const r = linhaConferida_(sh, row, CAB_VENC.length, r => String(r[COL_V.ID]).trim() === String(id || '').trim());
    const antes = String(r[COL_V.OBS] || '').trim(), nova = String(obs || '').trim();
    sh.getRange(row, COL_V.SITUACAO + 1).setValue(cancelar ? 'CANCELADO' : 'CONCLUÍDO');
    sh.getRange(row, COL_V.CONCLUIDO + 1).setValue(data ? Utilities.parseDate(data, tz_(), 'yyyy-MM-dd') : new Date()).setNumberFormat('dd/MM/yyyy');
    if (nova) sh.getRange(row, COL_V.OBS + 1).setValue(antes ? antes + ' · ' + nova : nova);
  });
  return { venc: lerVenc_() };
}

/** Volta para ATIVO um item concluído/descartado por engano. */
function vencReabrir(chave, row, id) {
  usuario_(chave);
  comLock_(() => {
    const sh = ss_().getSheetByName(ABA_VENC);
    linhaConferida_(sh, row, CAB_VENC.length, r => String(r[COL_V.ID]).trim() === String(id || '').trim());
    sh.getRange(row, COL_V.SITUACAO + 1).setValue('ATIVO');
    sh.getRange(row, COL_V.CONCLUIDO + 1).clearContent();
  });
  return { venc: lerVenc_() };
}

function vencAnexar(chave, row, id, anexo) {
  usuario_(chave);
  acrescentarAnexo_(ABA_VENC, PASTAS.VENC, COL_V.ANEXO, anexo, sh => {
    linhaConferida_(sh, row, CAB_VENC.length, r => String(r[COL_V.ID]).trim() === String(id || '').trim());
    return row;
  });
  return { venc: lerVenc_() };
}

// =====================================================================
// Ver anexo dentro do app (foto ou PDF), sem depender de permissão no Drive de quem está vendo.
// O arquivo só é entregue se o link estiver num lançamento do app (despesas ou vencimentos).
// =====================================================================
const LIMITE_VER_ANEXO = 20 * 1024 * 1024;

/** mod = 'd' (despesas) | 'v' (vencimentos); url = link do anexo. */
function anexoVer(chave, mod, url) {
  usuario_(chave);
  const id = idDrive_(url);
  if (!id) throw new Error('Link de anexo não reconhecido.');
  const ler = { d: lerDesp_, v: lerVenc_ }[mod];
  if (!ler) throw new Error('Módulo inválido.');
  const visivel = ler().some(r => (String(r.anexo || '').match(/https?:\/\/[^\s,;]+/g) || []).some(x => idDrive_(x) === id));
  if (!visivel) throw new Error('Este anexo não está em nenhum lançamento do app.');
  const f = DriveApp.getFileById(id);
  const nome = f.getName(), link = f.getUrl();
  // Arquivo do Google (Docs/Planilhas) vai como PDF
  const blob = /^application\/vnd\.google-apps/.test(f.getMimeType()) ? f.getAs('application/pdf') : f.getBlob();
  const bytes = blob.getBytes();
  if (bytes.length > LIMITE_VER_ANEXO) return { nome: nome, url: link, grande: Math.round(bytes.length / 1048576) };
  return { nome: nome, url: link, mime: blob.getContentType() || f.getMimeType(), data: Utilities.base64Encode(bytes) };
}

/** ID do arquivo nos formatos de link do Drive: .../file/d/ID/..., ...open?id=ID, ...uc?id=ID */
function idDrive_(url) {
  const s = String(url || '');
  const m = s.match(/\/d\/([\w-]{20,})/) || s.match(/[?&]id=([\w-]{20,})/);
  return m ? m[1] : '';
}

// =====================================================================
// Alerta semanal por e-mail (vencidos, vencem em até 30 dias, garantias acabando)
// =====================================================================

const ALERTA_EMAILS = ''; // vazio = e-mail do dono do script; para incluir mais pessoas: 'a@x.com,b@y.com'
const ALERTA_DIAS = 30;

/** Rode pelo editor para receber o e-mail agora (mesmo que não haja nada vencido). */
function testarAlertas() { enviarAlertas_(true); }

/** Rotina agendada: só envia e-mail se houver algo vencido ou perto de vencer. */
function verificarAlertas() { enviarAlertas_(false); }

/** Rode uma vez pelo editor: agenda a verificação para toda segunda-feira, entre 7h e 8h. */
function ativarAlertas() {
  desativarAlertas();
  ScriptApp.newTrigger('verificarAlertas').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7).create();
}

function desativarAlertas() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'verificarAlertas').forEach(t => ScriptApp.deleteTrigger(t));
}

function enviarAlertas_(sempre) {
  const a = levantarAlertas_();
  if (!sempre && !a.vencidos.length && !a.proximos.length && !a.garantias.length) return;
  const assunto = '🏠 Apartamento: ' + a.vencidos.length + ' vencido(s), ' + a.proximos.length + ' a vencer'
    + (a.garantias.length ? ', ' + a.garantias.length + ' garantia(s) acabando' : '');
  const bloco = x => x.nome + '\n' + x.detalhe;
  const partes = [];
  if (a.vencidos.length) partes.push('⚠️ VENCIDOS\n\n' + a.vencidos.map(bloco).join('\n\n'));
  if (a.proximos.length) partes.push('⏳ VENCEM EM BREVE\n\n' + a.proximos.map(bloco).join('\n\n'));
  if (a.garantias.length) partes.push('🛡️ GARANTIAS ACABANDO (teste antes que acabe)\n\n' + a.garantias.map(bloco).join('\n\n'));
  const texto = partes.length ? partes.join('\n\n————————\n\n') : 'Nada vencido ou perto de vencer. 👍';
  MailApp.sendEmail(ALERTA_EMAILS || Session.getEffectiveUser().getEmail(), assunto, texto, { htmlBody: alertaHtml_(a) });
}

function alertaHtml_(a) {
  const esc = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const td = 'padding:10px 12px;border-bottom:1px solid #e5e7eb;vertical-align:top;font-size:14px;';
  const th = 'padding:8px 12px;text-align:left;font-size:12px;color:#6b7280;text-transform:uppercase;border-bottom:2px solid #e5e7eb;';
  const tabela = (titulo, cor, fundo, itens) => !itens.length ? '' :
    '<h2 style="margin:24px 0 8px;font-size:16px;color:' + cor + ';">' + titulo + ' (' + itens.length + ')</h2>'
    + '<table cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;background:#fff;border:1px solid #e5e7eb;border-left:4px solid ' + cor + ';">'
    + '<tr style="background:' + fundo + ';"><th style="' + th + '">Item</th><th style="' + th + '">Situação</th></tr>'
    + itens.map(x => '<tr><td style="' + td + 'font-weight:700;">' + esc(x.nome) + '</td><td style="' + td + 'color:#374151;">' + esc(x.detalhe) + '</td></tr>').join('')
    + '</table>';
  const vazio = !a.vencidos.length && !a.proximos.length && !a.garantias.length;
  return '<div style="font-family:Arial,Helvetica,sans-serif;background:#f3f4f6;padding:20px;"><div style="max-width:640px;margin:0 auto;">'
    + '<h1 style="margin:0 0 4px;font-size:20px;color:#111827;">🏠 Vencimentos do apartamento</h1>'
    + '<p style="margin:0;color:#6b7280;font-size:13px;">Vencidos e os que vencem em até ' + ALERTA_DIAS + ' dias.</p>'
    + (vazio ? '<p style="margin:24px 0;font-size:15px;">Nada vencido ou perto de vencer. 👍</p>' : '')
    + tabela('⚠️ Vencidos', '#b42318', '#fef3f2', a.vencidos)
    + tabela('⏳ Vencem em breve', '#b54708', '#fffaeb', a.proximos)
    + tabela('🛡️ Garantias acabando — teste antes que acabe', '#175cd3', '#eff8ff', a.garantias)
    + '<p style="margin:24px 0 0;color:#9ca3af;font-size:12px;">Para tirar um item da lista, conclua ou renove pelo app (aba Vencimentos).</p>'
    + '</div></div>';
}

function levantarAlertas_() {
  const tz = tz_(), hoje = fmt_(new Date(), tz);
  const out = { vencidos: [], proximos: [], garantias: [] };
  const dias = iso => Math.round((Utilities.parseDate(iso, tz, 'yyyy-MM-dd') - Utilities.parseDate(hoje, tz, 'yyyy-MM-dd')) / 864e5);
  lerVenc_().filter(x => x.situacao === 'ATIVO' && x.venceData).forEach(x => {
    const d = dias(x.venceData), br = x.venceData.split('-').reverse().join('/');
    const item = {
      nome: x.desc && normTxt_(x.desc) !== normTxt_(x.tipo) ? x.tipo + ' — ' + x.desc : (x.tipo || x.desc),
      detalhe: d < 0 ? 'venceu em ' + br : 'vence em ' + br + (d === 0 ? ' (hoje)' : ' (em ' + d + ' dia' + (d === 1 ? '' : 's') + ')')
    };
    // Garantia não pede ação ao vencer: só avisa enquanto ainda vale e está perto do fim
    if (x.tipo === 'Garantia') { if (d >= 0 && d <= ALERTA_DIAS) out.garantias.push(item); return; }
    if (d < 0) out.vencidos.push(item); else if (d <= ALERTA_DIAS) out.proximos.push(item);
  });
  return out;
}

// =====================================================================
// Backup semanal (cópia .xlsx numa pasta do Drive; guarda as 8 últimas)
// =====================================================================

const BACKUPS_MANTIDOS = 8;

function fazerBackup() {
  const pasta = pasta_(PASTAS.BACKUP);
  const data = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const resp = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + ss_().getId() + '/export?format=xlsx', {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) throw new Error('Falha no backup — HTTP ' + resp.getResponseCode());
  const prefixo = 'BACKUP Apartamento - ';
  pasta.createFile(resp.getBlob().setName(prefixo + data + '.xlsx'));
  const arquivos = [];
  const it = pasta.getFiles();
  while (it.hasNext()) { const f = it.next(); if (f.getName().indexOf(prefixo) === 0) arquivos.push(f); }
  arquivos.sort((a, b) => b.getDateCreated() - a.getDateCreated()).slice(BACKUPS_MANTIDOS).forEach(f => f.setTrashed(true));
}

/** Rode uma vez pelo editor: agenda o backup para todo domingo, entre 3h e 4h da manhã. */
function ativarBackupSemanal() {
  desativarBackup();
  ScriptApp.newTrigger('fazerBackup').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(3).create();
}

function desativarBackup() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'fazerBackup').forEach(t => ScriptApp.deleteTrigger(t));
}

// =====================================================================
// Auxiliares
// =====================================================================

let ss__ = null;
function ss_() {
  if (!ss__) ss__ = ID_PLANILHA ? SpreadsheetApp.openById(ID_PLANILHA) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss__) throw new Error('Planilha não encontrada: preencha ID_PLANILHA no Código.gs.');
  return ss__;
}

function tz_() {
  return ss_().getSpreadsheetTimeZone();
}

// Próximo ID da coluna A: maior número com o prefixo + 1 (F0001, V0001...)
function proximoId_(sh, prefixo) {
  const ids = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues() : [];
  const maior = ids.reduce((m, r) => Math.max(m, parseInt(String(r[0]).replace(/\D/g, ''), 10) || 0), 0);
  return prefixo + String(maior + 1).padStart(4, '0');
}

function pasta_(nome) {
  const it = DriveApp.getFoldersByName(nome);
  return it.hasNext() ? it.next() : DriveApp.createFolder(nome);
}

function comLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

// Lê a linha e garante que ela ainda é a mesma que estava na tela.
function linhaConferida_(sh, row, ncols, confere) {
  if (!sh) throw new Error('Aba não encontrada. Rode configurar() no editor.');
  const r = sh.getRange(row, 1, 1, ncols).getValues()[0];
  if (!confere(r)) throw new Error('A planilha mudou desde o último carregamento. Atualize a tela e tente de novo.');
  return r;
}

/**
 * Grava em linha já existente mesmo com filtro ligado na aba (linha escondida pelo filtro não recebe o valor,
 * sem erro nenhum — aconteceu no app dos pais em 24/11/2026). Tira os critérios, grava e devolve como estavam.
 */
function semFiltro_(sh, fn) {
  const f = sh.getFilter();
  if (!f) return fn();
  const rg = f.getRange(), c0 = rg.getColumn(), crit = [];
  for (let c = c0; c < c0 + rg.getNumColumns(); c++) {
    const k = f.getColumnFilterCriteria(c);
    if (k) crit.push([c, k]);
  }
  if (!crit.length) return fn();
  crit.forEach(([c]) => f.removeColumnFilterCriteria(c));
  try {
    const res = fn();
    SpreadsheetApp.flush();
    return res;
  } finally {
    crit.forEach(([c, k]) => { try { f.setColumnFilterCriteria(c, k); } catch (e) { console.log('Filtro col ' + c + ': ' + e.message); } });
  }
}

// Relê a linha depois de uma edição: se a planilha não aceitou (filtro, proteção), avisa em vez de dar "✓" falso.
function conferirEdicao_(sh, row, ncols, confere) {
  SpreadsheetApp.flush();
  if (!confere(sh.getRange(row, 1, 1, ncols).getValues()[0])) {
    throw new Error('A planilha não aceitou a alteração (linha ' + row + '). Confira se há filtro ou proteção na aba e avise o Bruno.');
  }
}

function salvarAnexo_(nomePasta, anexo) {
  const blob = Utilities.newBlob(Utilities.base64Decode(anexo.data), anexo.mime, anexo.name);
  return pasta_(nomePasta).createFile(blob).getUrl();
}

/** Um ou vários arquivos (it.anexos e/ou it.anexo): os links ficam na mesma célula, separados por ", ". */
function salvarAnexos_(nomePasta, it) {
  const lista = [].concat(it.anexos || [], it.anexo ? [it.anexo] : []);
  return lista.filter(a => a && a.data).map(a => salvarAnexo_(nomePasta, a)).join(', ');
}

/**
 * Anexo extra num lançamento já gravado: salva no Drive e ACRESCENTA o link na célula (", " entre links).
 * conferir(sh) confere a linha e devolve o número dela. Se falhar, o arquivo recém-salvo vai para a lixeira.
 */
function acrescentarAnexo_(aba, nomePasta, colAnexo, anexo, conferir) {
  if (!anexo || !anexo.data) throw new Error('Escolha um arquivo.');
  const url = salvarAnexo_(nomePasta, anexo);
  try {
    comLock_(() => {
      const sh = ss_().getSheetByName(aba);
      const cel = sh.getRange(conferir(sh), colAnexo + 1);
      const atual = String(cel.getValue() || '').trim();
      cel.setValue(atual ? atual + ', ' + url : url);
    });
  } catch (err) {
    const m = url.match(/\/d\/([\w-]+)/);
    if (m) try { DriveApp.getFileById(m[1]).setTrashed(true); } catch (e) {}
    throw err;
  }
}

function fmt_(d, tz) {
  return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
}

function parseData_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) return v;
  const m = String(v || '').match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  return m ? new Date(+m[3], +m[2] - 1, +m[1]) : null;
}

// Aceita número ou texto: "R$ 1.234,56", "76652,6", "223.297"
function parseNum_(v) {
  if (typeof v === 'number') return v;
  let s = String(v || '').replace(/[R$\s]/g, '');
  if (!s) return 0;
  if (s.indexOf(',') >= 0) s = s.replace(/\./g, '').replace(',', '.');
  else if (/^\d{1,3}(\.\d{3})+$/.test(s)) s = s.replace(/\./g, '');
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function normTxt_(s) {
  return String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').trim().replace(/\s+/g, ' ').toUpperCase();
}

// 'yyyy-MM-dd' + n meses, ajustando o dia ao fim do mês (ex.: 31/01 -> 28/02).
function addMeses_(iso, n) {
  const [y, m, d] = iso.split('-').map(Number);
  const alvo = new Date(Date.UTC(y, m - 1 + n, 1));
  const ultimo = new Date(Date.UTC(alvo.getUTCFullYear(), alvo.getUTCMonth() + 1, 0)).getUTCDate();
  return alvo.getUTCFullYear() + '-' + String(alvo.getUTCMonth() + 1).padStart(2, '0') + '-' + String(Math.min(d, ultimo)).padStart(2, '0');
}
