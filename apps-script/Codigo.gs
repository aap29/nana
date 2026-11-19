/**
 * NÃNA — Central de Cadastros (Marca & Estratégia)
 * Planilha: "Nãna — Cadastros (Marca & Estratégia)"
 *
 * O que faz:
 *  1. Recebe do app cada cadastro novo, cada acesso e cada plano ativado.
 *  2. Grava na aba "Cadastros" (uma linha por pessoa, sem senha).
 *  3. Manda e-mail para os criadores a cada cadastro novo.
 *  4. Todo dia às 8h: atualiza a situação de cada pessoa, monta a aba
 *     "Ações do dia" com mensagens prontas e links de WhatsApp, atualiza o
 *     "Painel" e manda o resumo por e-mail.
 *  5. Confirma quem é criador (CEO ou cofundador) com o código de criador.
 *
 * Instalação (uma vez só):
 *  - Extensões → Apps Script → cole este código → Salvar.
 *  - Rode a função "instalar" e autorize.
 *  - Implantar → Nova implantação → Tipo "App da Web" →
 *    Executar como: Eu · Quem pode acessar: Qualquer pessoa → Implantar.
 *  - Copie a URL que termina em /exec e cole no app (CFG.registro).
 */

/* CEO = a conta Google dona desta planilha (quem instalou o robô). */
const CEO_EMAIL = () => Session.getEffectiveUser().getEmail().toLowerCase();
const ASSINATURA = '\n\nEquipe Nãna · Marca & Estratégia\nWhatsApp (51) 98416-8848';
const PRECO = { mensal: 29.90, anual: 249.90 };
const APP_LINK = 'https://aap29.github.io/nana/instalar.html';

const COLS = ['ID', 'Cadastro em', 'Nome completo', 'Celular', 'E-mail', 'Cidade', 'Atividade',
  'Teste até', 'Plano', 'Plano até', 'Situação', 'Acessos', 'Último acesso', 'Aparelho',
  'Criador', 'Contato feito em', 'Observações'];
const C = Object.fromEntries(COLS.map((n, i) => [n, i + 1]));

const PLANILHA_ID = '1Y8vD7ojqufqmm3lfzY7tPo7wIxKnCieWtV9jObeXi9Y';
function ss_() { return SpreadsheetApp.openById(PLANILHA_ID); }
/* Código de criador: troque aqui e rode "instalar" de novo para mudar. */
const CODIGO_CRIADOR_INICIAL = 'TROQUE-ESTE-CODIGO';

/* ---------- instalação ---------- */
function instalar() {
  const ss = ss_();
  ss.setSpreadsheetTimeZone('America/Sao_Paulo');
  aba_('Cadastros', COLS);
  aba_('Acessos', ['Data/hora', 'E-mail', 'Evento', 'Aparelho']);
  const cr = aba_('Criadores', ['E-mail', 'Papel', 'Ativo', 'Recebe avisos']);
  if (cr.getLastRow() < 2) cr.appendRow([CEO_EMAIL(), 'CEO', 'Sim', 'Sim']);
  aba_('Ações do dia', ['Prioridade', 'Assistente', 'Nome', 'Celular', 'Motivo', 'Mensagem', 'Abrir no WhatsApp']);
  aba_('Painel', ['Indicador', 'Valor']);
  const p1 = ss.getSheetByName('Página1') || ss.getSheetByName('Sheet1');
  if (p1 && p1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(p1);
  const props = PropertiesService.getScriptProperties();
  if (CODIGO_CRIADOR_INICIAL !== 'TROQUE-ESTE-CODIGO') props.setProperty('CODIGO_CRIADOR', CODIGO_CRIADOR_INICIAL.toUpperCase());
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('resumoDiario').timeBased().atHour(8).everyDays(1).inTimezone('America/Sao_Paulo').create();
  resumoDiario();
}

function aba_(nome, cab) {
  const ss = ss_();
  let sh = ss.getSheetByName(nome);
  if (!sh) sh = ss.insertSheet(nome);
  if (sh.getLastRow() === 0) {
    sh.appendRow(cab);
    sh.getRange(1, 1, 1, cab.length).setFontWeight('bold').setBackground('#4b1d50').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

/* ---------- entrada do app ---------- */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse((e.postData && e.postData.contents) || '{}');
    if (d.acao === 'cadastro') return ok_(registrarCadastro_(d));
    if (d.acao === 'acesso') return ok_(registrarAcesso_(d));
    if (d.acao === 'plano') return ok_(registrarPlano_(d));
    return ok_({ ok: false, erro: 'ação desconhecida' });
  } catch (err) {
    return ok_({ ok: false, erro: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.acao === 'ehcriador') {
    const papel = papelCriador_(p.email);
    return ok_(papel ? { criador: true, papel } : { criador: false });
  }
  if (p.acao === 'criador') {
    const papel = papelCriador_(p.email);
    const certo = String(p.codigo || '').trim().toUpperCase() ===
      String(PropertiesService.getScriptProperties().getProperty('CODIGO_CRIADOR') || '').toUpperCase();
    return ok_(papel && certo ? { ok: true, papel } : { ok: false });
  }
  return ok_({ ok: true, app: 'Nãna' });
}

function ok_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

const limpa_ = v => String(v == null ? '' : v).slice(0, 200).replace(/^[=+\-@]/, "'$&");
const normEmail_ = e => String(e || '').trim().toLowerCase();

function linhaPorEmail_(sh, email) {
  const n = sh.getLastRow();
  if (n < 2) return 0;
  const em = sh.getRange(2, C['E-mail'], n - 1, 1).getValues();
  for (let i = 0; i < em.length; i++) if (normEmail_(em[i][0]) === email) return i + 2;
  return 0;
}

function registrarCadastro_(d) {
  const sh = aba_('Cadastros', COLS);
  const email = normEmail_(d.email);
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email)) return { ok: false, erro: 'e-mail inválido' };
  let r = linhaPorEmail_(sh, email);
  const agora = new Date();
  const papel = papelCriador_(email);
  if (!r) {
    sh.appendRow([limpa_(d.id), agora, limpa_(d.nome), limpa_(d.celular), email, limpa_(d.cidade), limpa_(d.atividade),
      d.trialAte ? new Date(d.trialAte + 'T12:00:00') : '', papel ? 'criador' : 'gratis', '', papel ? 'Criador' : 'Em teste', 1, agora,
      limpa_(d.aparelho), papel || '', '', '']);
    log_(email, 'cadastro', d.aparelho);
    avisarCadastro_(d, email, papel);
  } else {
    log_(email, 'cadastro repetido', d.aparelho);
  }
  return { ok: true };
}

function registrarAcesso_(d) {
  const sh = aba_('Cadastros', COLS);
  const email = normEmail_(d.email);
  const r = linhaPorEmail_(sh, email);
  if (r) {
    const n = Number(sh.getRange(r, C['Acessos']).getValue()) || 0;
    sh.getRange(r, C['Acessos']).setValue(n + 1);
    sh.getRange(r, C['Último acesso']).setValue(new Date());
    if (d.aparelho) sh.getRange(r, C['Aparelho']).setValue(limpa_(d.aparelho));
  } else if (d.nome) {
    return registrarCadastro_(d);
  }
  log_(email, 'acesso', d.aparelho);
  return { ok: true };
}

function registrarPlano_(d) {
  const sh = aba_('Cadastros', COLS);
  const email = normEmail_(d.email);
  const r = linhaPorEmail_(sh, email);
  if (!r) return { ok: false, erro: 'cadastro não encontrado' };
  sh.getRange(r, C['Plano']).setValue(limpa_(d.plano));
  sh.getRange(r, C['Plano até']).setValue(d.planoAte ? new Date(d.planoAte + 'T12:00:00') : '');
  sh.getRange(r, C['Situação']).setValue('Assinante');
  log_(email, 'plano ' + d.plano, '');
  return { ok: true };
}

function log_(email, ev, ap) {
  aba_('Acessos', ['Data/hora', 'E-mail', 'Evento', 'Aparelho']).appendRow([new Date(), email, ev, limpa_(ap)]);
}

/* ---------- criadores ---------- */
function criadores_() {
  const sh = aba_('Criadores', ['E-mail', 'Papel', 'Ativo', 'Recebe avisos']);
  const n = sh.getLastRow();
  const v = n > 1 ? sh.getRange(2, 1, n - 1, 4).getValues() : [];
  const lista = v.filter(r => r[0] && String(r[2]).toLowerCase() !== 'não' && String(r[2]).toLowerCase() !== 'nao')
    .map(r => ({ email: normEmail_(r[0]), papel: r[1] || 'Cofundador', avisos: String(r[3]).toLowerCase() !== 'não' && String(r[3]).toLowerCase() !== 'nao' }));
  if (!lista.some(c => c.email === CEO_EMAIL())) lista.unshift({ email: CEO_EMAIL(), papel: 'CEO', avisos: true });
  return lista;
}
function papelCriador_(email) {
  const c = criadores_().find(x => x.email === normEmail_(email));
  return c ? c.papel : '';
}
function destinatarios_() {
  return criadores_().filter(c => c.avisos).map(c => c.email).join(',');
}

/* ---------- e-mails ---------- */
function avisarCadastro_(d, email, papel) {
  const assunto = papel ? `Nãna: criador entrou (${papel})` : `Nãna: novo cadastro — ${d.nome}`;
  const wa = waLink_(d.celular, msg_('boasvindas', { nome: d.nome, trialAte: d.trialAte }));
  const html = `<div style="font-family:Arial,sans-serif;color:#2b1428">
    <h2 style="color:#4b1d50">Novo cadastro no Nãna</h2>
    <p><b>${esc_(d.nome)}</b><br>${esc_(d.atividade)} · ${esc_(d.cidade)}<br>Celular: ${esc_(d.celular)}<br>E-mail: ${esc_(email)}</p>
    <p>Teste grátis até <b>${d.trialAte ? fmt_(new Date(d.trialAte + 'T12:00:00')) : '—'}</b>.</p>
    ${papel ? '' : `<p><a href="${wa}" style="background:#a8431c;color:#fff;padding:12px 18px;border-radius:12px;text-decoration:none;font-weight:bold">Dar boas-vindas pelo WhatsApp</a></p>`}
    <p><a href="${ss_().getUrl()}">Abrir a planilha de cadastros</a></p></div>`;
  MailApp.sendEmail({ to: destinatarios_(), subject: assunto, htmlBody: html, name: 'Nãna' });
}

/* ---------- rotina diária ---------- */
function resumoDiario() {
  const sh = aba_('Cadastros', COLS);
  const hoje = new Date(); hoje.setHours(0, 0, 0, 0);
  const n = sh.getLastRow();
  const rows = n > 1 ? sh.getRange(2, 1, n - 1, COLS.length).getValues() : [];
  const dias = d => d instanceof Date ? Math.round((new Date(d).setHours(0, 0, 0, 0) - hoje) / 864e5) : null;
  const pessoas = rows.map((r, i) => {
    const p = { linha: i + 2, nome: r[C['Nome completo'] - 1], celular: r[C['Celular'] - 1], email: r[C['E-mail'] - 1],
      cidade: r[C['Cidade'] - 1], atividade: r[C['Atividade'] - 1], cadastro: r[C['Cadastro em'] - 1],
      trialAte: r[C['Teste até'] - 1], plano: String(r[C['Plano'] - 1] || 'gratis'), planoAte: r[C['Plano até'] - 1],
      ultimo: r[C['Último acesso'] - 1], criador: r[C['Criador'] - 1], contato: r[C['Contato feito em'] - 1] };
    p.sit = p.criador || p.plano === 'criador' ? 'Criador'
      : (p.plano !== 'gratis' && p.plano && (!p.planoAte || dias(p.planoAte) >= 0)) ? 'Assinante'
      : (p.trialAte && dias(p.trialAte) >= 0) ? 'Em teste' : 'Teste vencido';
    return p;
  });
  if (pessoas.length) sh.getRange(2, C['Situação'], pessoas.length, 1).setValues(pessoas.map(p => [p.sit]));

  const acoes = [];
  const add = (pri, k, p, motivo, m) => acoes.push([pri, k, p.nome, p.celular, motivo, msg_(m, p), waLink_(p.celular, msg_(m, p))]);
  const contatadoHoje = p => p.contato instanceof Date && dias(p.contato) === 0;
  pessoas.filter(p => p.sit !== 'Criador' && !contatadoHoje(p)).forEach(p => {
    const dt = dias(p.trialAte), dp = dias(p.planoAte), dc = dias(p.cadastro), du = dias(p.ultimo);
    if (p.sit === 'Em teste' && dt !== null && dt <= 2) add(1, 'Conversão', p, `Teste termina ${dt === 0 ? 'hoje' : 'em ' + dt + ' dia(s)'}`, 'conversao');
    else if (p.sit === 'Assinante' && dp !== null && dp <= 5) add(2, 'Renovação', p, `Plano vence em ${fmt_(p.planoAte)}`, 'renovacao');
    else if (dc !== null && dc >= -2) add(3, 'Recepção', p, 'Cadastrou nos últimos 2 dias', 'boasvindas');
    else if (p.sit !== 'Teste vencido' && du !== null && du <= -5) add(4, 'Retenção', p, `Sem entrar há ${-du} dias`, 'retencao');
    else if (p.sit === 'Teste vencido' && dt !== null && dt >= -30) add(5, 'Reconquista', p, `Teste venceu em ${fmt_(p.trialAte)}`, 'reconquista');
  });
  acoes.sort((a, b) => a[0] - b[0]);
  const ac = aba_('Ações do dia', ['Prioridade', 'Assistente', 'Nome', 'Celular', 'Motivo', 'Mensagem', 'Abrir no WhatsApp']);
  if (ac.getLastRow() > 1) ac.getRange(2, 1, ac.getLastRow() - 1, 7).clearContent();
  if (acoes.length) ac.getRange(2, 1, acoes.length, 7).setValues(acoes);

  const conta = s => pessoas.filter(p => p.sit === s).length;
  const passaram = pessoas.filter(p => p.sit === 'Assinante' || p.sit === 'Teste vencido').length;
  const receita = pessoas.filter(p => p.sit === 'Assinante')
    .reduce((t, p) => t + (p.plano === 'anual' ? PRECO.anual / 12 : p.plano === 'mensal' ? PRECO.mensal : 0), 0);
  const ind = [
    ['Atualizado em', new Date()],
    ['Cadastros (sem criadores)', pessoas.filter(p => p.sit !== 'Criador').length],
    ['Em teste', conta('Em teste')], ['Assinantes', conta('Assinante')], ['Teste vencido', conta('Teste vencido')],
    ['Conversão (assinantes ÷ quem já passou do teste)', passaram ? conta('Assinante') / passaram : 0],
    ['Receita mensal estimada (R$)', Math.round(receita * 100) / 100],
    ['Cadastros nos últimos 7 dias', pessoas.filter(p => p.sit !== 'Criador' && dias(p.cadastro) !== null && dias(p.cadastro) >= -7).length],
    ['Ações para hoje', acoes.length]];
  const pn = aba_('Painel', ['Indicador', 'Valor']);
  if (pn.getLastRow() > 1) pn.getRange(2, 1, pn.getLastRow() - 1, 2).clearContent();
  pn.getRange(2, 1, ind.length, 2).setValues(ind);
  pn.getRange(7, 2).setNumberFormat('0%');

  if (pessoas.length || acoes.length) {
    const lis = acoes.slice(0, 25).map(a => `<li><b>${esc_(a[1])}</b> · ${esc_(a[2])} — ${esc_(a[4])} · <a href="${a[6]}">WhatsApp</a></li>`).join('');
    MailApp.sendEmail({
      to: destinatarios_(), name: 'Nãna',
      subject: `Nãna hoje: ${conta('Em teste')} em teste · ${conta('Assinante')} assinantes · ${acoes.length} contatos`,
      htmlBody: `<div style="font-family:Arial,sans-serif;color:#2b1428"><h2 style="color:#4b1d50">Resumo do Nãna</h2>
        <p>Cadastros: <b>${pessoas.filter(p => p.sit !== 'Criador').length}</b> · Em teste: <b>${conta('Em teste')}</b> · Assinantes: <b>${conta('Assinante')}</b> · Teste vencido: <b>${conta('Teste vencido')}</b></p>
        ${acoes.length ? `<h3>Quem chamar hoje</h3><ol>${lis}</ol>` : '<p>Nenhum contato pendente hoje.</p>'}
        <p><a href="${ss_().getUrl()}">Abrir a planilha</a></p></div>`
    });
  }
}

/* ---------- mensagens ---------- */
function msg_(k, p) {
  const n = String(p.nome || '').split(' ')[0];
  const t = p.trialAte ? fmt_(p.trialAte instanceof Date ? p.trialAte : new Date(p.trialAte + 'T12:00:00')) : '';
  const brl = v => 'R$ ' + v.toFixed(2).replace('.', ',');
  const m = {
    boasvindas: `Oi, ${n}! Que bom ter você no Nãna. Seus 7 dias grátis vão até ${t}, com tudo liberado. Comece cadastrando seus produtos e sua primeira encomenda. Se quiser, te mostro em 10 minutos como deixar tudo pronto.`,
    conversao: `Oi, ${n}! Seu teste do Nãna termina em ${t}. Para continuar com pedidos, agenda e dinheiro separado sem parar, o plano mensal é ${brl(PRECO.mensal)} e o anual ${brl(PRECO.anual)} (sai ${brl(PRECO.anual / 12)} por mês). Quer que eu ative para você?`,
    reconquista: `Oi, ${n}! Seu teste do Nãna terminou, mas todos os seus dados continuam guardados. Quando quiser voltar a usar tudo, ativo seu plano na hora: mensal ${brl(PRECO.mensal)} ou anual ${brl(PRECO.anual)}.`,
    retencao: `Oi, ${n}! Sentimos sua falta no Nãna. Posso te ajudar a organizar as encomendas desta semana pelo app? Em 10 minutos deixamos tudo pronto.`,
    renovacao: `Oi, ${n}! Seu plano do Nãna vence em ${p.planoAte ? fmt_(p.planoAte) : ''}. Quer renovar para não perder o ritmo? Te mando o PIX.`
  };
  return m[k] + ASSINATURA;
}
function waLink_(cel, txt) {
  let d = String(cel || '').replace(/\D/g, '');
  if (d.length <= 11) d = '55' + d;
  return 'https://wa.me/' + d + '?text=' + encodeURIComponent(txt);
}
function fmt_(d) { return Utilities.formatDate(new Date(d), 'America/Sao_Paulo', 'dd/MM/yyyy'); }
function esc_(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }
