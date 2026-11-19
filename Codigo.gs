/**
 * Backend – Relatório da Realidade Escolar 2026 · SEMED Lagoa da Canoa (v4)
 *
 * Novidades da v4:
 * - Notificações automáticas por e-mail via MailApp:
 *   1) Para a escola quando a SEMED cadastra/reabre uma observação com link direto para a etapa.
 *   2) Para a SEMED (e-mail configurável no painel) quando a escola finaliza o relatório ou envia correção.
 * - Registro do e-mail do gestor/coordenador no vínculo da escola e gestão de e-mails no painel.
 * - Otimização de performance com cache de instâncias de abas e dados em memória por requisição.
 *
 * COMO ATUALIZAR:
 * 1) Na planilha: Extensões > Apps Script, substitua TODO o código por este arquivo.
 * 2) Troque SENHA_ADMIN abaixo pela senha do painel administrativo.
 * 3) Execute configurar() uma vez (autorize o acesso à planilha, ao Drive e ao envio de e-mails).
 * 4) Implantar > Gerenciar implantações > lápis (editar) > Versão: "Nova versão" > Implantar.
 *    Assim a URL do app da Web continua a mesma e o index.html não precisa mudar de endereço.
 *
 * ABAS (criadas automaticamente):
 *  Escolas      – Escola | CPF | Nome | Função | E-mail | Vinculado em
 *  Respostas    – CPF | Escola | Etapa | Título | Resposta | Atualizado em | Concluída
 *  Observacoes  – Escola | Etapa | Título | Observação | Criada em | Status | Atendida em
 *  Config       – Chave | Valor   (guarda prazo de edição, e-mail da SEMED e URL do app)
 */
const SENHA_ADMIN = 'semed2026';
const URL_PADRAO_SISTEMA = 'https://douglas14031999.github.io/relatorio-realidade-escolar-2026/';
const TZ = 'America/Maceio';
const PASTA_FOTOS = 'Evidências – Relatório da Realidade Escolar 2026';
const CAB = {
  Escolas: ['Escola', 'CPF', 'Nome', 'Função', 'E-mail', 'Vinculado em'],
  Respostas: ['CPF', 'Escola', 'Etapa', 'Título', 'Resposta', 'Atualizado em', 'Concluída'],
  Observacoes: ['Escola', 'Etapa', 'Título', 'Observação', 'Criada em', 'Status', 'Atendida em'],
  Config: ['Chave', 'Valor']
};
const TEXTO = { Escolas: ['B:B'], Respostas: ['A:A'], Observacoes: ['E:E', 'G:G'], Config: ['B:B'] };

let _SS = null;
const _SHEETS = {};
let _CACHE = {};

function getSS() {
  if (!_SS) _SS = SpreadsheetApp.getActiveSpreadsheet();
  return _SS;
}

function garantirColunasEscolas(s) {
  if (!s) return;
  const lastCol = s.getLastColumn();
  if (lastCol === 0) return;
  const header = s.getRange(1, 1, 1, Math.max(lastCol, 6)).getValues()[0].map(h => String(h || '').trim());
  if (header.indexOf('E-mail') === -1) {
    const vincIdx = header.indexOf('Vinculado em');
    if (vincIdx !== -1) {
      s.insertColumnBefore(vincIdx + 1);
      s.getRange(1, vincIdx + 1).setValue('E-mail').setFontWeight('bold').setBackground('#2f7a62').setFontColor('#fff');
    } else {
      s.getRange(1, 5).setValue('E-mail').setFontWeight('bold').setBackground('#2f7a62').setFontColor('#fff');
      s.getRange(1, 6).setValue('Vinculado em').setFontWeight('bold').setBackground('#2f7a62').setFontColor('#fff');
    }
  }
}

function configurar() {
  const ss = getSS();
  Object.keys(CAB).forEach(n => {
    let s = ss.getSheetByName(n);
    const h = CAB[n];
    if (!s) {
      s = ss.insertSheet(n);
      s.setFrozenRows(1);
      (TEXTO[n] || []).forEach(r => s.getRange(r).setNumberFormat('@'));
    }
    if (n === 'Escolas') garantirColunasEscolas(s);
    const lastCol = s.getLastColumn();
    const lastRow = s.getLastRow();
    if (lastCol < h.length || lastRow === 0) {
      const atual = lastCol > 0 && lastRow > 0 ? s.getRange(1, 1, 1, Math.max(h.length, lastCol)).getValues()[0] : [];
      h.forEach((t, k) => { if (!atual[k]) s.getRange(1, k + 1).setValue(t); });
      s.getRange(1, 1, 1, h.length).setFontWeight('bold').setBackground('#2f7a62').setFontColor('#fff');
    }
  });
  pasta();
}

function aba(nome) {
  if (_SHEETS[nome]) return _SHEETS[nome];
  const ss = getSS();
  let s = ss.getSheetByName(nome);
  const h = CAB[nome];
  if (!s) {
    s = ss.insertSheet(nome);
    s.setFrozenRows(1);
    (TEXTO[nome] || []).forEach(r => s.getRange(r).setNumberFormat('@'));
    if (h) {
      s.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#2f7a62').setFontColor('#fff');
    }
  }
  _SHEETS[nome] = s;
  return s;
}

const so = c => String(c || '').replace(/\D/g, '');
const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
const hoje = () => Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
const dia = v => v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : String(v || '').slice(0, 10);
const quando = v => v instanceof Date ? Utilities.formatDate(v, TZ, 'dd/MM/yyyy HH:mm') : String(v || '');
const sim = v => v === true || String(v).toUpperCase() === 'TRUE';
const br = iso => iso ? iso.split('-').reverse().join('/') : '';

function doGet() { return out({ ok: true, servico: 'Relatório da Realidade Escolar', versao: 4 }); }

function doPost(e) {
  _CACHE = {};
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    return out(roteia(JSON.parse(e.postData.contents)));
  } catch (err) {
    return out({ ok: false, erro: 'Erro no servidor: ' + err });
  } finally {
    try { lock.releaseLock(); } catch (x) { }
  }
}

/* ---------- configuração: prazo, e-mail admin e appUrl ---------- */
function prazo() {
  if (_CACHE.prazo !== undefined) return _CACHE.prazo;
  const l = aba('Config').getDataRange().getValues().slice(1).find(x => x[0] === 'prazo');
  _CACHE.prazo = l ? dia(l[1]) : '';
  return _CACHE.prazo;
}
function gravaPrazo(p) {
  _CACHE.prazo = p;
  const s = aba('Config'), d = s.getDataRange().getValues();
  const n = d.findIndex((x, k) => k > 0 && x[0] === 'prazo');
  if (n > 0) s.getRange(n + 1, 2).setValue(p); else s.appendRow(['prazo', p]);
}
function encerrado() { const p = prazo(); return !!p && hoje() > p; }

function adminEmail() {
  if (_CACHE.adminEmail !== undefined) return _CACHE.adminEmail;
  const l = aba('Config').getDataRange().getValues().slice(1).find(x => x[0] === 'admin_email');
  _CACHE.adminEmail = l ? String(l[1] || '').trim() : '';
  return _CACHE.adminEmail;
}
function gravaAdminEmail(email) {
  _CACHE.adminEmail = email;
  const s = aba('Config'), d = s.getDataRange().getValues();
  const n = d.findIndex((x, k) => k > 0 && x[0] === 'admin_email');
  if (n > 0) s.getRange(n + 1, 2).setValue(email); else s.appendRow(['admin_email', email]);
}

function appUrl() {
  if (_CACHE.appUrl !== undefined) return _CACHE.appUrl;
  const l = aba('Config').getDataRange().getValues().slice(1).find(x => x[0] === 'app_url');
  const saved = l ? String(l[1] || '').trim() : '';
  _CACHE.appUrl = (/^https?:\/\//i.test(saved) ? saved : (typeof URL_PADRAO_SISTEMA !== 'undefined' ? URL_PADRAO_SISTEMA : ''));
  return _CACHE.appUrl;
}
function gravaAppUrl(url) {
  if (!url || !/^https?:\/\//i.test(url)) return;
  _CACHE.appUrl = url;
  const s = aba('Config'), d = s.getDataRange().getValues();
  const n = d.findIndex((x, k) => k > 0 && x[0] === 'app_url');
  if (n > 0) s.getRange(n + 1, 2).setValue(url); else s.appendRow(['app_url', url]);
}

/* ---------- leitura de dados ---------- */
function escolas() {
  if (_CACHE.escolas) return _CACHE.escolas;
  const values = aba('Escolas').getDataRange().getValues();
  if (values.length <= 1) return [];
  const header = values[0].map(h => String(h || '').trim());
  let colEmail = header.indexOf('E-mail');
  if (colEmail === -1) colEmail = 4;

  const list = values.slice(1).map((l, k) => {
    const c = so(l[1]);
    let email = '';
    if (colEmail !== -1 && l[colEmail] && String(l[colEmail]).includes('@')) {
      email = String(l[colEmail]).trim();
    }
    return {
      row: k + 2,
      escola: String(l[0] || '').trim(),
      cpf: c ? c.padStart(11, '0') : '',
      nome: l[2] || '',
      funcao: l[3] || '',
      email: email
    };
  }).filter(x => x.escola);
  _CACHE.escolas = list;
  return list;
}

function resumo(cpf) {
  const m = escolas().filter(x => x.cpf === cpf);
  return {
    ok: true,
    nome: m[0] ? m[0].nome : '',
    email: m[0] ? m[0].email : '',
    escolas: m.map(x => x.escola),
    funcoes: Object.fromEntries(m.map(x => [x.escola, x.funcao])),
    emails: Object.fromEntries(m.map(x => [x.escola, x.email]))
  };
}

function obsEscola(escola) {
  if (!_CACHE.obs) {
    const all = {};
    aba('Observacoes').getDataRange().getValues().slice(1).forEach(l => {
      const esc = String(l[0]).trim();
      if (!all[esc]) all[esc] = {};
      all[esc][l[1]] = { texto: l[3], data: dia(l[4]), status: String(l[5] || 'pendente'), atendida: dia(l[6]) };
    });
    _CACHE.obs = all;
  }
  return _CACHE.obs[escola] || {};
}

function pendente(escola, etapa) { const o = obsEscola(escola)[etapa]; return !!o && o.status === 'pendente'; }

function respostasEscola(escola) {
  const respostas = {}, concluidas = [];
  aba('Respostas').getDataRange().getValues().slice(1).forEach(l => {
    if (String(l[1]).trim() === escola) { respostas[l[2]] = l[4]; if (sim(l[6])) concluidas.push(l[2]); }
  });
  return { respostas, concluidas };
}

function dono(cpf, escola) {
  if (cpf.length !== 11) return 'CPF inválido.';
  const x = escolas().find(y => y.escola === escola);
  return x && x.cpf === cpf ? '' : 'Escola não vinculada a este CPF.';
}

/* ---------- envio de notificações por e-mail ---------- */
function enviarEmailObservacao(escolaNome, etapaId, etapaTitulo, obsTexto, appUrlParam) {
  try {
    if (!escolaNome) {
      Logger.log('Aviso: Nenhum nome de escola informado. Para testar o envio de e-mails pelo editor do Apps Script, selecione a função "testarEnvioEmail" e clique em Executar.');
      return false;
    }
    const esc = escolas().find(x => x.escola === escolaNome);
    if (!esc || !esc.email) {
      Logger.log('Aviso: Escola sem e-mail cadastrado para envio de observação: ' + escolaNome);
      return false;
    }
    const rawBase = appUrlParam || appUrl() || '';
    const base = /^https?:\/\//i.test(rawBase) ? rawBase : '';
    const link = base ? (base + (base.includes('?') ? '&' : '?') + 'escola=' + encodeURIComponent(escolaNome) + '&etapa=' + encodeURIComponent(etapaId)) : '';

    const html = `
    <div style="font-family:Arial,Helvetica,sans-serif;line-height:1.6;color:#2c3e50;max-width:600px;margin:0 auto;border:1px solid #d0e6dc;border-radius:12px;overflow:hidden;background:#ffffff;">
      <div style="background-color:#205645;padding:24px 20px;text-align:center;color:#ffffff;">
        <h2 style="margin:0;font-size:18px;letter-spacing:0.5px;">PREFEITURA MUNICIPAL DE LAGOA DA CANOA</h2>
        <div style="font-size:13px;opacity:0.9;margin-top:4px;">SECRETARIA MUNICIPAL DE EDUCAÇÃO · SEMED</div>
      </div>
      <div style="padding:28px 24px;">
        <h3 style="margin-top:0;color:#205645;font-size:17px;">Observação da SEMED no Relatório da Realidade Escolar</h3>
        <p style="font-size:14px;margin-bottom:12px;">Olá, <b>${esc.nome || 'Gestor(a)'}</b>${esc.funcao ? ' (' + esc.funcao + ')' : ''},</p>
        <p style="font-size:14px;margin-bottom:16px;">A equipe da SEMED registrou uma observação referente ao relatório da sua unidade escolar:</p>
        
        <div style="background:#eef6f3;border-left:4px solid #205645;padding:12px 16px;border-radius:4px;margin-bottom:20px;">
          <strong style="color:#205645;font-size:15px;">${escolaNome}</strong><br>
          <span style="font-size:13px;color:#555;">Etapa: <b>${etapaTitulo || etapaId}</b></span>
        </div>

        <div style="background:#fff9f0;border:1px solid #fbd38d;border-left:4px solid #dd6b20;padding:16px;border-radius:6px;margin-bottom:24px;">
          <div style="font-weight:bold;color:#c05621;font-size:13px;margin-bottom:6px;text-transform:uppercase;">Instrução / Ajuste Solicitado:</div>
          <div style="font-size:14px;color:#2d3748;white-space:pre-wrap;">${obsTexto}</div>
        </div>

        ${link ? `
        <div style="text-align:center;margin:30px 0;">
          <a href="${link}" style="background-color:#205645;color:#ffffff;text-decoration:none;font-weight:bold;font-size:15px;padding:14px 28px;border-radius:8px;display:inline-block;box-shadow:0 3px 6px rgba(0,0,0,0.12);">Acessar Etapa no Relatório</a>
        </div>
        <p style="font-size:12px;color:#718096;text-align:center;">Se o botão não abrir diretamente, acesse pelo link:<br><a href="${link}" style="color:#205645;word-break:break-all;">${link}</a></p>
        ` : ''}

        <hr style="border:none;border-top:1px solid #e2e8f0;margin:24px 0;">
        <p style="font-size:12px;color:#a0aec0;margin-bottom:0;">Mensagem automática do Relatório da Realidade Escolar 2026 · SEMED Lagoa da Canoa - AL.</p>
      </div>
    </div>
    `;

    MailApp.sendEmail({
      to: esc.email,
      subject: '[SEMED] Ajuste solicitado no Relatório Escolar – ' + escolaNome,
      htmlBody: html,
      name: 'SEMED Lagoa da Canoa'
    });
    return true;
  } catch (err) {
    Logger.log('Erro ao enviar e-mail de observação: ' + err);
    return false;
  }
}

function enviarEmailFinalizacao(escolaNome, responsavelNome, responsavelFuncao, appUrlParam) {
  try {
    if (!escolaNome) {
      Logger.log('Aviso: Nenhum nome de escola informado. Para testar o envio de e-mails pelo editor do Apps Script, selecione a função "testarEnvioEmail" e clique em Executar.');
      return false;
    }
    const dest = adminEmail();
    if (!dest) {
      Logger.log('Aviso: E-mail da SEMED não configurado. Notificação de finalização não enviada.');
      return false;
    }
    const rawBase = appUrlParam || appUrl() || '';
    const base = /^https?:\/\//i.test(rawBase) ? rawBase : '';
    const link = base ? (base + (base.includes('?') ? '&' : '?') + 'admin=1') : '';

    const html = `
    <div style="font-family:Arial,Helvetica,sans-serif;line-height:1.6;color:#2c3e50;max-width:600px;margin:0 auto;border:1px solid #d0e6dc;border-radius:12px;overflow:hidden;background:#ffffff;">
      <div style="background-color:#205645;padding:24px 20px;text-align:center;color:#ffffff;">
        <h2 style="margin:0;font-size:18px;letter-spacing:0.5px;">PREFEITURA MUNICIPAL DE LAGOA DA CANOA</h2>
        <div style="font-size:13px;opacity:0.9;margin-top:4px;">SECRETARIA MUNICIPAL DE EDUCAÇÃO · SEMED</div>
      </div>
      <div style="padding:28px 24px;">
        <div style="background:#e6fffa;border:1px solid #b2f5ea;border-left:4px solid #319795;padding:12px 16px;border-radius:4px;margin-bottom:20px;">
          <h3 style="margin:0;color:#234e52;font-size:16px;">✓ Relatório da Realidade Escolar 2026 Concluído</h3>
        </div>
        
        <p style="font-size:14px;">Prezada equipe da SEMED,</p>
        <p style="font-size:14px;">A unidade escolar <b>${escolaNome}</b> gravou a etapa final de <b>Conclusão e Assinaturas</b> do Relatório da Realidade Escolar 2026.</p>

        <table style="width:100%;font-size:13px;border-collapse:collapse;margin:16px 0;background:#f7fafc;border-radius:6px;overflow:hidden;">
          <tr>
            <td style="padding:10px 14px;border-bottom:1px solid #edf2f7;font-weight:bold;color:#4a5568;width:35%;">Unidade Escolar:</td>
            <td style="padding:10px 14px;border-bottom:1px solid #edf2f7;color:#1a202c;"><b>${escolaNome}</b></td>
          </tr>
          <tr>
            <td style="padding:10px 14px;border-bottom:1px solid #edf2f7;font-weight:bold;color:#4a5568;">Responsável:</td>
            <td style="padding:10px 14px;border-bottom:1px solid #edf2f7;color:#1a202c;">${responsavelNome || 'Gestor'} (${responsavelFuncao || 'Equipe Diretiva'})</td>
          </tr>
          <tr>
            <td style="padding:10px 14px;font-weight:bold;color:#4a5568;">Data / Horário:</td>
            <td style="padding:10px 14px;color:#1a202c;">${quando(new Date())}</td>
          </tr>
        </table>

        ${link ? `
        <div style="text-align:center;margin:28px 0;">
          <a href="${link}" style="background-color:#205645;color:#ffffff;text-decoration:none;font-weight:bold;font-size:15px;padding:14px 28px;border-radius:8px;display:inline-block;box-shadow:0 3px 6px rgba(0,0,0,0.12);">Abrir Painel Administrativo</a>
        </div>
        ` : ''}

        <hr style="border:none;border-top:1px solid #e2e8f0;margin:24px 0;">
        <p style="font-size:12px;color:#a0aec0;margin-bottom:0;">Notificação gerada automaticamente pelo sistema Relatório da Realidade Escolar 2026.</p>
      </div>
    </div>
    `;

    MailApp.sendEmail({
      to: dest,
      subject: '[SEMED] Relatório Finalizado – ' + escolaNome,
      htmlBody: html,
      name: 'Relatório Escolar 2026'
    });
    return true;
  } catch (err) {
    Logger.log('Erro ao enviar e-mail de finalização: ' + err);
    return false;
  }
}

function enviarEmailCorrecao(escolaNome, responsavelNome, responsavelFuncao, etapaId, etapaTitulo, appUrlParam) {
  try {
    if (!escolaNome) {
      Logger.log('Aviso: Nenhum nome de escola informado. Para testar o envio de e-mails pelo editor do Apps Script, selecione a função "testarEnvioEmail" e clique em Executar.');
      return false;
    }
    const dest = adminEmail();
    if (!dest) return false;
    const rawBase = appUrlParam || appUrl() || '';
    const base = /^https?:\/\//i.test(rawBase) ? rawBase : '';
    const link = base ? (base + (base.includes('?') ? '&' : '?') + 'admin=1') : '';

    const html = `
    <div style="font-family:Arial,Helvetica,sans-serif;line-height:1.6;color:#2c3e50;max-width:600px;margin:0 auto;border:1px solid #d0e6dc;border-radius:12px;overflow:hidden;background:#ffffff;">
      <div style="background-color:#205645;padding:24px 20px;text-align:center;color:#ffffff;">
        <h2 style="margin:0;font-size:18px;letter-spacing:0.5px;">PREFEITURA MUNICIPAL DE LAGOA DA CANOA</h2>
        <div style="font-size:13px;opacity:0.9;margin-top:4px;">SECRETARIA MUNICIPAL DE EDUCAÇÃO · SEMED</div>
      </div>
      <div style="padding:28px 24px;">
        <h3 style="margin-top:0;color:#2b6cb0;font-size:16px;">Ajuste Enviado para Revisão</h3>
        <p style="font-size:14px;">A unidade escolar <b>${escolaNome}</b> enviou a correção solicitada para a etapa:</p>
        
        <div style="background:#ebf8ff;border-left:4px solid #3182ce;padding:12px 16px;border-radius:4px;margin-bottom:20px;">
          <strong style="color:#2b6cb0;font-size:14px;">${etapaTitulo || etapaId}</strong><br>
          <span style="font-size:13px;color:#4a5568;">Responsável: ${responsavelNome || ''} (${responsavelFuncao || ''})</span>
        </div>

        <p style="font-size:14px;">Acesse o painel administrativo da SEMED para analisar o ajuste enviado e aprovar a etapa.</p>

        ${link ? `
        <div style="text-align:center;margin:28px 0;">
          <a href="${link}" style="background-color:#205645;color:#ffffff;text-decoration:none;font-weight:bold;font-size:15px;padding:14px 28px;border-radius:8px;display:inline-block;box-shadow:0 3px 6px rgba(0,0,0,0.12);">Ver e Aprovar no Painel</a>
        </div>
        ` : ''}

        <hr style="border:none;border-top:1px solid #e2e8f0;margin:24px 0;">
        <p style="font-size:12px;color:#a0aec0;margin-bottom:0;">Notificação gerada automaticamente pelo sistema Relatório da Realidade Escolar 2026.</p>
      </div>
    </div>
    `;

    MailApp.sendEmail({
      to: dest,
      subject: '[SEMED] Correção Enviada – ' + escolaNome + ' (' + (etapaTitulo || etapaId) + ')',
      htmlBody: html,
      name: 'Relatório Escolar 2026'
    });
    return true;
  } catch (err) {
    Logger.log('Erro ao enviar e-mail de correção: ' + err);
    return false;
  }
}

/**
 * Função para testar o envio de e-mails e autorizar as permissões do Google diretamente pelo editor do Apps Script.
 * Selecione esta função no menu superior do Apps Script e clique em "Executar".
 */
function testarEnvioEmail() {
  const emailAdmin = adminEmail() || (typeof Session !== 'undefined' && Session.getActiveUser ? Session.getActiveUser().getEmail() : '');
  const escComEmail = escolas().find(x => x.email);
  const dest = emailAdmin || (escComEmail ? escComEmail.email : '');

  if (!dest) {
    Logger.log('Atenção: Nenhum e-mail de destino encontrado para o teste. Configure o e-mail da SEMED no painel de administração ou cadastre um e-mail para alguma escola.');
    return false;
  }

  const html = `
  <div style="font-family:Arial,Helvetica,sans-serif;line-height:1.6;color:#2c3e50;max-width:600px;margin:0 auto;border:1px solid #d0e6dc;border-radius:12px;overflow:hidden;background:#ffffff;">
    <div style="background-color:#205645;padding:20px;text-align:center;color:#ffffff;">
      <h2 style="margin:0;font-size:18px;">PREFEITURA MUNICIPAL DE LAGOA DA CANOA</h2>
      <div style="font-size:13px;opacity:0.9;margin-top:4px;">SECRETARIA MUNICIPAL DE EDUCAÇÃO · SEMED</div>
    </div>
    <div style="padding:24px;">
      <h3 style="margin-top:0;color:#205645;">Teste de Notificação por E-mail</h3>
      <p style="font-size:14px;">Este e-mail confirma que o serviço de envio do <b>Relatório da Realidade Escolar 2026</b> foi autorizado e está funcionando perfeitamente!</p>
      <div style="background:#eef6f3;border-left:4px solid #205645;padding:12px 16px;border-radius:4px;margin:16px 0;">
        <b>Status:</b> Permissões e serviço MailApp autorizados com sucesso.<br>
        <b>Data/Hora do Teste:</b> ${new Date().toLocaleString('pt-BR')}
      </div>
      <p style="font-size:12px;color:#a0aec0;margin-bottom:0;">Notificação de teste gerada pelo Apps Script.</p>
    </div>
  </div>
  `;

  MailApp.sendEmail({
    to: dest,
    subject: '[SEMED] Teste de Notificação – Relatório Escolar 2026',
    htmlBody: html,
    name: 'Relatório Escolar 2026'
  });

  Logger.log('Sucesso! E-mail de teste enviado para: ' + dest);
  return true;
}

/* ---------- roteador ---------- */
function roteia(r) {
  const a = String(r.action || ''), cpf = so(r.cpf);
  if (r.appUrl) gravaAppUrl(r.appUrl);

  if (a.indexOf('admin_') === 0) {
    const s = String(r.senha || '').trim().toLowerCase();
    const esperada = String(SENHA_ADMIN || '').trim().toLowerCase();
    if (s !== esperada) return { ok: false, erro: 'Senha incorreta.' };
    return admin(r);
  }
  if (a === 'foto') {
    const s = String(r.senha || '').trim().toLowerCase();
    const esperada = String(SENHA_ADMIN || '').trim().toLowerCase();
    if (s !== esperada) { const e = dono(cpf, r.escola); if (e) return { ok: false, erro: e }; }
    return foto(r.id);
  }

  if (a === 'login') return resumo(cpf);
  if (a === 'disponiveis') return { ok: true, escolas: escolas().filter(x => !x.cpf).map(x => x.escola) };
  
  if (a === 'vincular') {
    const nome = String(r.nome || '').trim();
    const email = String(r.email || '').trim().toLowerCase();
    if (cpf.length !== 11) return { ok: false, erro: 'CPF inválido.' };
    if (nome.split(/\s+/).length < 2) return { ok: false, erro: 'Informe seu nome completo.' };
    if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      return { ok: false, erro: 'Informe um e-mail válido para receber notificações da SEMED.' };
    }
    const x = escolas().find(y => y.escola === r.escola);
    if (!x) return { ok: false, erro: 'Escola não encontrada.' };
    if (x.cpf && x.cpf !== cpf) return { ok: false, ocupada: true, erro: 'Esta escola acabou de ser vinculada por outra pessoa. Escolha outra.' };
    
    garantirColunasEscolas(aba('Escolas'));
    aba('Escolas').getRange(x.row, 2, 1, 5).setValues([[cpf, nome, r.funcao || '', email, new Date()]]);
    _CACHE.escolas = null;
    return resumo(cpf);
  }

  if (a === 'atualizar_email') {
    const email = String(r.email || '').trim().toLowerCase();
    if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return { ok: false, erro: 'Informe um e-mail válido.' };
    const x = escolas().find(y => y.escola === r.escola && y.cpf === cpf);
    if (!x) return { ok: false, erro: 'Escola não vinculada a este CPF.' };
    garantirColunasEscolas(aba('Escolas'));
    aba('Escolas').getRange(x.row, 5).setValue(email);
    _CACHE.escolas = null;
    return { ok: true, email };
  }

  const erro = dono(cpf, r.escola);
  if (erro) return { ok: false, erro };

  if (a === 'load') {
    const d = respostasEscola(r.escola);
    const esc = escolas().find(x => x.escola === r.escola);
    return {
      ok: true,
      respostas: d.respostas,
      concluidas: d.concluidas,
      prazo: prazo(),
      encerrado: encerrado(),
      obs: obsEscola(r.escola),
      email: esc ? (esc.email || '') : ''
    };
  }
  if (a === 'save') {
    if (!r.etapa) return { ok: false, erro: 'Etapa inválida.' };
    if (encerrado() && !pendente(r.escola, r.etapa))
      return { ok: false, erro: 'O prazo para edição terminou em ' + br(prazo()) + '. O relatório está disponível apenas para visualização e impressão.' };
    salva(cpf, r);

    const esc = escolas().find(x => x.escola === r.escola);
    const respNome = esc ? esc.nome : (r.nome || '');
    const respFunc = esc ? esc.funcao : '';

    if (r.concluida && pendente(r.escola, r.etapa)) {
      atende(r.escola, r.etapa);
      enviarEmailCorrecao(r.escola, respNome, respFunc, r.etapa, r.titulo || '', r.appUrl);
    }

    if (r.etapa === 'fim' && r.concluida) {
      enviarEmailFinalizacao(r.escola, respNome, respFunc, r.appUrl);
    }

    return { ok: true };
  }
  if (a === 'upload') {
    if (encerrado() && !pendente(r.escola, 'anexo')) return { ok: false, erro: 'O prazo para edição terminou.' };
    const blob = Utilities.newBlob(Utilities.base64Decode(r.data), 'image/jpeg', String(r.escola).slice(0, 40) + ' – ' + (r.nome || 'evidencia.jpg'));
    return { ok: true, id: pasta().createFile(blob).getId() };
  }
  return { ok: false, erro: 'Ação inválida.' };
}

function salva(cpf, r) {
  const s = aba('Respostas'), d = s.getDataRange().getValues();
  const n = d.findIndex((l, k) => k > 0 && String(l[1]).trim() === r.escola && l[2] === r.etapa);
  const valor = String(r.valor == null ? '' : r.valor).slice(0, 49000);
  const agora = new Date();
  if (n > 0) {
    const concl = !!r.concluida || sim(d[n][6]);
    s.getRange(n + 1, 1, 1, 7).setValues([[cpf, r.escola, r.etapa, r.titulo || '', valor, agora, concl]]);
  } else {
    s.appendRow([cpf, r.escola, r.etapa, r.titulo || '', valor, agora, !!r.concluida]);
  }
}

function atende(escola, etapa) {
  const s = aba('Observacoes'), d = s.getDataRange().getValues();
  const n = d.findIndex((l, k) => k > 0 && String(l[0]).trim() === escola && l[1] === etapa);
  if (n > 0) s.getRange(n + 1, 6, 1, 2).setValues([['atendida', hoje()]]);
  _CACHE.obs = null;
}

/* ---------- painel administrativo ---------- */
function admin(r) {
  const a = r.action;
  if (a === 'admin_resumo') {
    const agg = {}, pend = {}, rev = {};
    aba('Respostas').getDataRange().getValues().slice(1).forEach(l => {
      const k = String(l[1]).trim(); if (!k) return;
      const g = agg[k] = agg[k] || { c: {}, n: 0, u: null };
      g.n++; if (sim(l[6])) g.c[l[2]] = 1;
      if (l[5] instanceof Date && (!g.u || l[5] > g.u)) g.u = l[5];
    });
    aba('Observacoes').getDataRange().getValues().slice(1).forEach(l => {
      const st = String(l[5] || '');
      const k = String(l[0]).trim();
      if (st === 'pendente') { pend[k] = (pend[k] || 0) + 1; }
      else if (st === 'revisao') { rev[k] = (rev[k] || 0) + 1; }
    });
    return {
      ok: true,
      prazo: prazo(),
      emailAdmin: adminEmail(),
      escolas: escolas().map(x => {
        const g = agg[x.escola] || { c: {}, n: 0, u: null };
        return {
          escola: x.escola,
          responsavel: x.nome,
          funcao: x.funcao,
          email: x.email || '',
          cpf: x.cpf ? '***.' + x.cpf.slice(3, 6) + '.' + x.cpf.slice(6, 9) + '-**' : '',
          feitas: Object.keys(g.c).length,
          total: r.total,
          fim: !!g.c.fim,
          iniciada: g.n > 0,
          pend: pend[x.escola] || 0,
          rev: rev[x.escola] || 0,
          atualizado: g.u ? quando(g.u) : ''
        };
      })
    };
  }
  if (a === 'admin_relatorio') {
    const d = respostasEscola(r.escola);
    const escObj = escolas().find(x => x.escola === r.escola);
    return {
      ok: true,
      respostas: d.respostas,
      concluidas: d.concluidas,
      obs: obsEscola(r.escola),
      email: escObj ? (escObj.email || '') : '',
      responsavel: escObj ? (escObj.nome || '') : ''
    };
  }
  if (a === 'admin_add') {
    const n = String(r.escola || '').trim();
    if (n.length < 3) return { ok: false, erro: 'Informe o nome da escola.' };
    if (escolas().some(x => x.escola.toLowerCase() === n.toLowerCase())) return { ok: false, erro: 'Esta escola já está cadastrada.' };
    garantirColunasEscolas(aba('Escolas'));
    aba('Escolas').appendRow([n, '', '', '', '', '']);
    _CACHE.escolas = null;
    return { ok: true };
  }
  if (a === 'admin_liberar') {
    const x = escolas().find(y => y.escola === r.escola);
    if (x) aba('Escolas').getRange(x.row, 2, 1, 5).clearContent();
    _CACHE.escolas = null;
    return { ok: true };
  }
  if (a === 'admin_excluir') {
    const x = escolas().find(y => y.escola === r.escola);
    if (x) aba('Escolas').deleteRow(x.row);
    _CACHE.escolas = null;
    return { ok: true };
  }
  if (a === 'admin_prazo') {
    const p = String(r.prazo || '');
    if (p && !/^\d{4}-\d{2}-\d{2}$/.test(p)) return { ok: false, erro: 'Data inválida.' };
    gravaPrazo(p);
    return { ok: true, prazo: p };
  }
  if (a === 'admin_email') {
    const em = String(r.email || '').trim().toLowerCase();
    if (em && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(em)) return { ok: false, erro: 'Informe um e-mail válido.' };
    gravaAdminEmail(em);
    return { ok: true, email: em };
  }
  if (a === 'admin_set_email') {
    const em = String(r.email || '').trim().toLowerCase();
    if (em && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(em)) return { ok: false, erro: 'Informe um e-mail válido.' };
    const x = escolas().find(y => y.escola === r.escola);
    if (!x) return { ok: false, erro: 'Escola não encontrada.' };
    garantirColunasEscolas(aba('Escolas'));
    aba('Escolas').getRange(x.row, 5).setValue(em);
    _CACHE.escolas = null;
    return { ok: true, email: em };
  }
  if (a === 'admin_obs') {
    if (!escolas().some(x => x.escola === r.escola)) return { ok: false, erro: 'Escola não encontrada.' };
    const texto = String(r.texto || '').trim().slice(0, 4000);
    const s = aba('Observacoes'), d = s.getDataRange().getValues();
    const n = d.findIndex((l, k) => k > 0 && String(l[0]).trim() === r.escola && l[1] === r.etapa);
    _CACHE.obs = null;
    if (!texto) { if (n > 0) s.deleteRow(n + 1); return { ok: true }; }
    const linha = [r.escola, r.etapa, r.titulo || '', texto, hoje(), 'pendente', ''];
    if (n > 0) s.getRange(n + 1, 1, 1, 7).setValues([linha]); else s.appendRow(linha);
    
    // Dispara e-mail de notificação para a escola
    const enviado = enviarEmailObservacao(r.escola, r.etapa, r.titulo || '', texto, r.appUrl);
    return { ok: true, emailEnviado: enviado };
  }
  if (a === 'admin_aprovar') {
    const s = aba('Observacoes'), d = s.getDataRange().getValues();
    const n = d.findIndex((l, k) => k > 0 && String(l[0]).trim() === r.escola && l[1] === r.etapa);
    _CACHE.obs = null;
    if (n > 0) s.getRange(n + 1, 6, 1, 2).setValues([['aprovada', hoje()]]);
    return { ok: true };
  }
  return { ok: false, erro: 'Ação inválida.' };
}

/* ---------- fotos no Drive ---------- */
function pasta() {
  const it = DriveApp.getFoldersByName(PASTA_FOTOS);
  return it.hasNext() ? it.next() : DriveApp.createFolder(PASTA_FOTOS);
}
function foto(id) {
  const f = DriveApp.getFileById(id), pid = pasta().getId(), p = f.getParents();
  let dentro = false;
  while (p.hasNext()) if (p.next().getId() === pid) dentro = true;
  if (!dentro) return { ok: false, erro: 'Arquivo não permitido.' };
  return { ok: true, data: 'data:' + (f.getMimeType() || 'image/jpeg') + ';base64,' + Utilities.base64Encode(f.getBlob().getBytes()) };
}
