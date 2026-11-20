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

const LOGO_EMAIL_B64 = "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";

function getLogoBlob() {
  try {
    return Utilities.newBlob(Utilities.base64Decode(LOGO_EMAIL_B64), 'image/png', 'logo-canoa.png');
  } catch (e) {
    Logger.log('Erro ao criar blob da logo: ' + e);
    return null;
  }
}

function escapeHtml(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function emailLayoutBase(title, greetingHtml, cardsHtml, btnText, btnLink, disclaimerText) {
  const cleanLink = /^https?:\/\//i.test(btnLink || '') ? btnLink : '';
  const disc = disclaimerText || 'Mensagem automática do Relatório da Realidade Escolar 2026. Em caso de dúvida, responda a este e-mail.';

  return `<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>${escapeHtml(title)}</title>
</head>
<body style="margin:0;padding:0;background-color:#EEF3F0;font-family:Arial,Helvetica,sans-serif;-webkit-font-smoothing:antialiased;">

<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color:#EEF3F0;margin:0;padding:24px 0;">
<tr>
  <td align="center" style="padding:0 12px;">

    <!-- Container Principal -->
    <table role="presentation" width="600" cellpadding="0" cellspacing="0" border="0" style="width:100%;max-width:600px;background-color:#FFFFFF;border-radius:12px;overflow:hidden;border:1px solid #DCE6E1;">

      <!-- Faixa tricolor (identidade Lagoa da Canoa) -->
      <tr>
        <td height="6" style="font-size:0;line-height:0;padding:0;">
          <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">
            <tr>
              <td width="60%" height="6" bgcolor="#3A8266" style="font-size:0;line-height:0;background-color:#3A8266;">&nbsp;</td>
              <td width="22%" height="6" bgcolor="#FF7A2B" style="font-size:0;line-height:0;background-color:#FF7A2B;">&nbsp;</td>
              <td width="18%" height="6" bgcolor="#E8005F" style="font-size:0;line-height:0;background-color:#E8005F;">&nbsp;</td>
            </tr>
          </table>
        </td>
      </tr>

      <!-- Cabeçalho com Logo Institucional Embutida (CID) -->
      <tr>
        <td align="center" style="padding:28px 32px 18px 32px;background-color:#FFFFFF;">
          <img src="cid:logoImg" width="128" height="128" alt="Prefeitura de Lagoa da Canoa – Nossa gente, nosso orgulho" style="width:128px;height:auto;max-width:128px;display:block;margin:0 auto;border:0;">
          <div style="margin-top:14px;font-family:Arial,Helvetica,sans-serif;font-size:13px;line-height:18px;color:#5B6B64;letter-spacing:0.3px;">
            Secretaria Municipal de Educação
          </div>
        </td>
      </tr>

      <!-- Título Principal -->
      <tr>
        <td style="padding:16px 36px 0 36px;border-top:1px solid #E3ECE7;">
          <h1 style="margin:12px 0 0 0;font-family:Georgia,'Times New Roman',serif;font-size:23px;line-height:31px;font-weight:bold;color:#1F5A45;">
            ${escapeHtml(title)}
          </h1>
        </td>
      </tr>

      <!-- Mensagem / Saudação -->
      <tr>
        <td style="padding:16px 36px 0 36px;font-family:Arial,Helvetica,sans-serif;font-size:15px;line-height:24px;color:#33413B;">
          ${greetingHtml}
        </td>
      </tr>

      <!-- Cards de Informação -->
      ${cardsHtml}

      <!-- Botão de Ação (compatível com todos os clientes sem VML) -->
      ${cleanLink ? `
      <tr>
        <td align="center" style="padding:28px 36px 0 36px;">
          <table role="presentation" border="0" cellpadding="0" cellspacing="0" style="margin:0 auto;">
            <tr>
              <td align="center" bgcolor="#1F5A45" style="border-radius:8px;background-color:#1F5A45;">
                <a href="${cleanLink}" target="_blank" style="font-family:Arial,Helvetica,sans-serif;font-size:16px;font-weight:bold;color:#FFFFFF;text-decoration:none;padding:14px 34px;display:inline-block;border-radius:8px;background-color:#1F5A45;border-bottom:3px solid #14402F;">
                  ${escapeHtml(btnText)}
                </a>
              </td>
            </tr>
          </table>
        </td>
      </tr>

      <!-- Link Alternativo para Copiar -->
      <tr>
        <td align="center" style="padding:18px 36px 28px 36px;font-family:Arial,Helvetica,sans-serif;font-size:12px;line-height:18px;color:#7A8882;">
          O botão não abriu? Acesse pelo endereço abaixo:<br>
          <a href="${cleanLink}" target="_blank" style="color:#1F5A45;word-break:break-all;">${cleanLink}</a>
        </td>
      </tr>
      ` : ''}

      <!-- Rodapé Institucional Verde -->
      <tr>
        <td align="center" bgcolor="#1F5A45" style="padding:20px 32px;background-color:#1F5A45;">
          <div style="font-family:Georgia,'Times New Roman',serif;font-size:16px;line-height:22px;font-weight:bold;color:#FFFFFF;">
            Nossa gente, nosso orgulho.
          </div>
          <div style="margin-top:6px;font-family:Arial,Helvetica,sans-serif;font-size:12px;line-height:18px;color:#B9D4C8;">
            Prefeitura Municipal de Lagoa da Canoa · Secretaria Municipal de Educação
          </div>
        </td>
      </tr>

    </table>

    <!-- Aviso inferior -->
    <table role="presentation" width="600" cellpadding="0" cellspacing="0" border="0" style="width:100%;max-width:600px;">
      <tr>
        <td align="center" style="padding:16px 20px 0 20px;font-family:Arial,Helvetica,sans-serif;font-size:11px;line-height:16px;color:#7A8882;">
          ${escapeHtml(disc)}
        </td>
      </tr>
    </table>

  </td>
</tr>
</table>

</body>
</html>`;
}

function enviarEmailObservacao(escolaNome, etapaId, etapaTitulo, obsTexto, appUrlParam) {
  try {
    if (!escolaNome) {
      Logger.log('Aviso: Nenhum nome de escola informado.');
      return false;
    }
    const esc = escolas().find(x => x.escola === escolaNome);
    if (!esc || !esc.email) {
      Logger.log('Aviso: Escola sem e-mail cadastrado para envio de observação: ' + escolaNome);
      return false;
    }
    const rawBase = appUrlParam || appUrl() || URL_PADRAO_SISTEMA || '';
    const base = /^https?:\/\//i.test(rawBase) ? rawBase : '';
    const link = base ? (base + (base.includes('?') ? '&' : '?') + 'escola=' + encodeURIComponent(escolaNome) + '&etapa=' + encodeURIComponent(etapaId)) : '';

    const title = 'A SEMED deixou uma observação no relatório da sua escola';
    const greeting = `Olá, <b>${escapeHtml(esc.nome || 'Gestor(a)')}</b> (${escapeHtml(esc.funcao || 'Equipe Gestora')}).<br>A equipe da SEMED analisou o relatório da sua unidade escolar e registrou o ajuste abaixo.`;

    const cards = `
    <!-- Escola e etapa -->
    <tr>
      <td style="padding:20px 36px 0 36px;">
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">
          <tr>
            <td width="5" bgcolor="#3A8266" style="width:5px;font-size:0;line-height:0;border-radius:3px 0 0 3px;">&nbsp;</td>
            <td bgcolor="#EEF5F1" style="padding:14px 18px;font-family:Arial,Helvetica,sans-serif;border-radius:0 8px 8px 0;">
              <div style="font-size:12px;line-height:16px;color:#5B6B64;">Escola</div>
              <div style="margin-top:2px;font-size:16px;line-height:22px;font-weight:bold;color:#1F5A45;">${escapeHtml(escolaNome)}</div>
              <div style="margin-top:10px;font-size:12px;line-height:16px;color:#5B6B64;">Etapa do relatório</div>
              <div style="margin-top:2px;font-size:15px;line-height:21px;font-weight:bold;color:#1B2621;">${escapeHtml(etapaTitulo || etapaId)}</div>
            </td>
          </tr>
        </table>
      </td>
    </tr>

    <!-- Ajuste solicitado -->
    <tr>
      <td style="padding:14px 36px 0 36px;">
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">
          <tr>
            <td width="5" bgcolor="#FF7A2B" style="width:5px;font-size:0;line-height:0;border-radius:3px 0 0 3px;">&nbsp;</td>
            <td bgcolor="#FFF4EB" style="padding:16px 18px;font-family:Arial,Helvetica,sans-serif;border-radius:0 8px 8px 0;">
              <div style="font-size:13px;line-height:18px;font-weight:bold;color:#B8470A;">Ajuste solicitado</div>
              <div style="margin-top:6px;font-size:16px;line-height:25px;color:#2A2018;">${escapeHtml(obsTexto).replace(/\n/g, '<br>')}</div>
            </td>
          </tr>
        </table>
      </td>
    </tr>`;

    const html = emailLayoutBase(title, greeting, cards, 'Abrir etapa no relatório', link);

    const plainText = `A SEMED deixou uma observação no relatório da sua escola

Olá, ${esc.nome || 'Gestor(a)'} (${esc.funcao || 'Equipe Gestora'}),

A equipe da SEMED analisou o relatório da sua unidade escolar e registrou o seguinte ajuste:

Escola: ${escolaNome}
Etapa: ${etapaTitulo || etapaId}

Ajuste Solicitado:
${obsTexto}

Para acessar e ajustar diretamente esta etapa no sistema, acesse:
${link || '(Acesse o sistema e abra a etapa indicada)'}

---
Prefeitura Municipal de Lagoa da Canoa · Secretaria Municipal de Educação
Mensagem automática do Relatório da Realidade Escolar 2026.`;

    const inlineImgs = {};
    const logoBlob = getLogoBlob();
    if (logoBlob) inlineImgs.logoImg = logoBlob;

    MailApp.sendEmail({
      to: esc.email,
      subject: '[SEMED] Ajuste solicitado no Relatório Escolar – ' + escolaNome,
      body: plainText,
      htmlBody: html,
      inlineImages: inlineImgs,
      name: 'Relatório Escolar 2026'
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
      Logger.log('Aviso: Nenhum nome de escola informado.');
      return false;
    }
    const dest = adminEmail();
    if (!dest) {
      Logger.log('Aviso: E-mail da SEMED não configurado. Notificação de finalização não enviada.');
      return false;
    }
    const rawBase = appUrlParam || appUrl() || URL_PADRAO_SISTEMA || '';
    const base = /^https?:\/\//i.test(rawBase) ? rawBase : '';
    const link = base ? (base + (base.includes('?') ? '&' : '?') + 'admin=1') : '';

    const title = 'Relatório da Realidade Escolar 2026 Finalizado';
    const greeting = `Olá, equipe da SEMED.<br>A unidade escolar abaixo concluiu o preenchimento e assinou todas as etapas do Relatório da Realidade Escolar 2026:`;

    const cards = `
    <!-- Resumo de Conclusão -->
    <tr>
      <td style="padding:20px 36px 0 36px;">
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">
          <tr>
            <td width="5" bgcolor="#3A8266" style="width:5px;font-size:0;line-height:0;border-radius:3px 0 0 3px;">&nbsp;</td>
            <td bgcolor="#EEF5F1" style="padding:16px 20px;font-family:Arial,Helvetica,sans-serif;border-radius:0 8px 8px 0;">
              <div style="font-size:12px;line-height:16px;color:#5B6B64;">Unidade escolar</div>
              <div style="margin-top:2px;font-size:17px;line-height:23px;font-weight:bold;color:#1F5A45;">${escapeHtml(escolaNome)}</div>
              
              <div style="margin-top:12px;font-size:12px;line-height:16px;color:#5B6B64;">Responsável pelo envio</div>
              <div style="margin-top:2px;font-size:15px;line-height:21px;font-weight:bold;color:#1B2621;">${escapeHtml(responsavelNome || 'Gestor(a)')} (${escapeHtml(responsavelFuncao || 'Equipe Diretiva')})</div>
              
              <div style="margin-top:12px;font-size:12px;line-height:16px;color:#5B6B64;">Data de finalização</div>
              <div style="margin-top:2px;font-size:14px;line-height:20px;color:#33413B;">${quando(new Date())}</div>

              <div style="margin-top:12px;font-size:12px;line-height:16px;color:#5B6B64;">Situação</div>
              <div style="margin-top:2px;font-size:14px;line-height:20px;font-weight:bold;color:#2F7A62;">✓ Todas as 30 etapas concluídas e assinadas</div>
            </td>
          </tr>
        </table>
      </td>
    </tr>`;

    const html = emailLayoutBase(title, greeting, cards, 'Abrir painel administrativo', link);

    const plainText = `Relatório da Realidade Escolar 2026 Finalizado

Unidade Escolar: ${escolaNome}
Responsável: ${responsavelNome || 'Gestor(a)'} (${responsavelFuncao || 'Equipe Diretiva'})
Data: ${quando(new Date())}
Status: Todas as 30 etapas concluídas e assinadas

Acesse o painel administrativo da SEMED para consultar ou imprimir:
${link || '(Acesse o painel administrativo)'}

---
Prefeitura Municipal de Lagoa da Canoa · Secretaria Municipal de Educação`;

    const inlineImgs = {};
    const logoBlob = getLogoBlob();
    if (logoBlob) inlineImgs.logoImg = logoBlob;

    MailApp.sendEmail({
      to: dest,
      subject: '[SEMED] Relatório Finalizado – ' + escolaNome,
      body: plainText,
      htmlBody: html,
      inlineImages: inlineImgs,
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
      Logger.log('Aviso: Nenhum nome de escola informado.');
      return false;
    }
    const dest = adminEmail();
    if (!dest) {
      Logger.log('Aviso: E-mail da SEMED não configurado. Notificação de correção não enviada.');
      return false;
    }
    const rawBase = appUrlParam || appUrl() || URL_PADRAO_SISTEMA || '';
    const base = /^https?:\/\//i.test(rawBase) ? rawBase : '';
    const link = base ? (base + (base.includes('?') ? '&' : '?') + 'admin=1') : '';

    const title = 'Ajuste enviado para revisão da SEMED';
    const greeting = `Olá, equipe da SEMED.<br>A unidade escolar abaixo enviou o ajuste solicitado no Relatório da Realidade Escolar 2026. Acesse o painel para analisar e aprovar:`;

    const cards = `
    <!-- Escola e etapa -->
    <tr>
      <td style="padding:20px 36px 0 36px;">
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">
          <tr>
            <td width="5" bgcolor="#3A8266" style="width:5px;font-size:0;line-height:0;border-radius:3px 0 0 3px;">&nbsp;</td>
            <td bgcolor="#EEF5F1" style="padding:14px 18px;font-family:Arial,Helvetica,sans-serif;border-radius:0 8px 8px 0;">
              <div style="font-size:12px;line-height:16px;color:#5B6B64;">Escola</div>
              <div style="margin-top:2px;font-size:16px;line-height:22px;font-weight:bold;color:#1F5A45;">${escapeHtml(escolaNome)}</div>
              <div style="margin-top:10px;font-size:12px;line-height:16px;color:#5B6B64;">Etapa corrigida</div>
              <div style="margin-top:2px;font-size:15px;line-height:21px;font-weight:bold;color:#1B2621;">${escapeHtml(etapaTitulo || etapaId)}</div>
              <div style="margin-top:10px;font-size:12px;line-height:16px;color:#5B6B64;">Responsável pelo envio</div>
              <div style="margin-top:2px;font-size:14px;line-height:20px;color:#33413B;">${escapeHtml(responsavelNome || 'Gestor(a)')} (${escapeHtml(responsavelFuncao || 'Equipe Diretiva')})</div>
            </td>
          </tr>
        </table>
      </td>
    </tr>

    <!-- Status da revisão -->
    <tr>
      <td style="padding:14px 36px 0 36px;">
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">
          <tr>
            <td width="5" bgcolor="#2B6CB0" style="width:5px;font-size:0;line-height:0;border-radius:3px 0 0 3px;">&nbsp;</td>
            <td bgcolor="#EBF8FF" style="padding:16px 18px;font-family:Arial,Helvetica,sans-serif;border-radius:0 8px 8px 0;">
              <div style="font-size:13px;line-height:18px;font-weight:bold;color:#2B6CB0;">Situação da etapa</div>
              <div style="margin-top:6px;font-size:16px;line-height:24px;color:#1A365D;">Aguardando análise e aprovação da SEMED</div>
              <div style="margin-top:6px;font-size:12px;color:#4A5568;">Enviado em: ${quando(new Date())}</div>
            </td>
          </tr>
        </table>
      </td>
    </tr>`;

    const html = emailLayoutBase(title, greeting, cards, 'Ver e aprovar no painel', link);

    const plainText = `Ajuste enviado para revisão da SEMED

Unidade Escolar: ${escolaNome}
Etapa Corrigida: ${etapaTitulo || etapaId}
Responsável: ${responsavelNome || 'Gestor(a)'} (${responsavelFuncao || 'Equipe Diretiva'})
Data: ${quando(new Date())}

Acesse o painel administrativo da SEMED para analisar e aprovar:
${link || '(Acesse o painel administrativo)'}

---
Prefeitura Municipal de Lagoa da Canoa · Secretaria Municipal de Educação`;

    const inlineImgs = {};
    const logoBlob = getLogoBlob();
    if (logoBlob) inlineImgs.logoImg = logoBlob;

    MailApp.sendEmail({
      to: dest,
      subject: '[SEMED] Correção Enviada – ' + escolaNome + ' (' + (etapaTitulo || etapaId) + ')',
      body: plainText,
      htmlBody: html,
      inlineImages: inlineImgs,
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

  const link = appUrl() || URL_PADRAO_SISTEMA || '';
  const title = 'Teste de Notificação por E-mail';
  const greeting = 'Olá!<br>Este e-mail confirma que o serviço de envio do <strong>Relatório da Realidade Escolar 2026</strong> está ativo, com a logomarca oficial embutida diretamente e alta entregabilidade na Caixa de Entrada.';

  const cards = `
  <tr>
    <td style="padding:20px 36px 0 36px;">
      <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">
        <tr>
          <td width="5" bgcolor="#3A8266" style="width:5px;font-size:0;line-height:0;border-radius:3px 0 0 3px;">&nbsp;</td>
          <td bgcolor="#EEF5F1" style="padding:16px 20px;font-family:Arial,Helvetica,sans-serif;border-radius:0 8px 8px 0;">
            <div style="font-size:12px;line-height:16px;color:#5B6B64;">Status do serviço</div>
            <div style="margin-top:2px;font-size:16px;line-height:22px;font-weight:bold;color:#1F5A45;">✓ MailApp operacional (Alta Entregabilidade)</div>
            
            <div style="margin-top:12px;font-size:12px;line-height:16px;color:#5B6B64;">Data e hora do teste</div>
            <div style="margin-top:2px;font-size:14px;line-height:20px;color:#33413B;">${new Date().toLocaleString('pt-BR')}</div>
            
            <div style="margin-top:12px;font-size:12px;line-height:16px;color:#5B6B64;">Endereço de destino</div>
            <div style="margin-top:2px;font-size:14px;line-height:20px;color:#33413B;">${escapeHtml(dest)}</div>

            <div style="margin-top:12px;font-size:12px;line-height:16px;color:#5B6B64;">Endereço do sistema (GitHub Pages)</div>
            <div style="margin-top:2px;font-size:13px;line-height:19px;color:#1F5A45;word-break:break-all;"><b>${escapeHtml(link)}</b></div>
          </td>
        </tr>
      </table>
    </td>
  </tr>`;

  const html = emailLayoutBase(title, greeting, cards, 'Acessar Relatório Escolar 2026', link, 'Mensagem de teste gerada pelo editor do Apps Script.');

  const plainText = `Teste de Notificação por E-mail – Relatório da Realidade Escolar 2026

Olá!
Este e-mail confirma que o serviço de envio do Relatório da Realidade Escolar 2026 está ativo e configurado com alta entregabilidade.

Destino: ${dest}
Data: ${new Date().toLocaleString('pt-BR')}
Link: ${link}

---
Prefeitura Municipal de Lagoa da Canoa · Secretaria Municipal de Educação`;

  const inlineImgs = {};
  const logoBlob = getLogoBlob();
  if (logoBlob) inlineImgs.logoImg = logoBlob;

  MailApp.sendEmail({
    to: dest,
    subject: '[SEMED] Teste de Notificação – Relatório Escolar 2026',
    body: plainText,
    htmlBody: html,
    inlineImages: inlineImgs,
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
