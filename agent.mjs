#!/usr/bin/env node
// Еженедельный newsletter об акциях.
//   node agent.mjs draft   [--date YYYY-MM-DD] [--no-llm] [--no-mail]
//   node agent.mjs approve [YYYY-MM-DD]
//   node agent.mjs send    [--now] [--dry-run]
//   node agent.mjs status
//
// Принцип: цены, скидки и сроки попадают в письмо только из Excel через шаблон.
// Claude пишет лишь тексты (тема, вступление, описания), и это проверяется кодом.

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import ExcelJS from 'exceljs';
import Handlebars from 'handlebars';

const ROOT = path.dirname(fileURLToPath(import.meta.url));
const STATE_DIR = path.join(ROOT, 'state');
const rel = (p) => path.join(ROOT, p);

if (fs.existsSync(rel('.env'))) process.loadEnvFile(rel('.env'));
const config = JSON.parse(fs.readFileSync(rel('config.json'), 'utf8'));
const TZ = config.schedule.timezone;

// ---------- Время по Берлину ----------

function berlinNow() {
  const parts = Object.fromEntries(
    new Intl.DateTimeFormat('en-CA', {
      timeZone: TZ, year: 'numeric', month: '2-digit', day: '2-digit',
      hour: '2-digit', minute: '2-digit', hour12: false, weekday: 'short',
    }).formatToParts(new Date()).map((p) => [p.type, p.value]),
  );
  return { date: `${parts.year}-${parts.month}-${parts.day}`, hour: Number(parts.hour) % 24 };
}

const addDays = (iso, n) => {
  const d = new Date(`${iso}T00:00:00Z`);
  d.setUTCDate(d.getUTCDate() + n);
  return d.toISOString().slice(0, 10);
};
const weekday = (iso) => new Date(`${iso}T00:00:00Z`).getUTCDay();
const fmtDate = (iso) => iso ? iso.split('-').reverse().join('.') : '';
const fmtEur = (n) => new Intl.NumberFormat('de-DE', { style: 'currency', currency: 'EUR' }).format(n);

function nextSendDate() {
  const { date, hour } = berlinNow();
  const target = config.schedule.weekday;
  let diff = (target - weekday(date) + 7) % 7;
  if (diff === 0 && hour >= config.schedule.hour) diff = 7;
  return addDays(date, diff);
}

function isoWeek(iso) {
  const d = new Date(`${iso}T00:00:00Z`);
  d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay() || 7));
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil(((d - yearStart) / 86400000 + 1) / 7);
}

// ---------- Excel с акциями ----------

const COLUMN_ALIASES = {
  produkt: ['produkt', 'товар', 'artikel', 'product'],
  marke: ['marke', 'бренд', 'brand'],
  alter_preis: ['alter preis', 'старая цена', 'streichpreis', 'old price'],
  neuer_preis: ['neuer preis', 'новая цена', 'aktionspreis', 'new price'],
  gueltig_von: ['gültig von', 'gueltig von', 'начало', 'start'],
  gueltig_bis: ['gültig bis', 'gueltig bis', 'срок', 'ende', 'end'],
  shop: ['shop', 'магазин', 'filiale'],
  bedingungen: ['bedingungen', 'условия', 'conditions'],
  link: ['link', 'url', 'ссылка'],
  bild: ['bild', 'bild-url', 'image', 'картинка'],
};

function cellValue(v) {
  if (v == null) return '';
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  if (typeof v === 'object') {
    if ('result' in v) return cellValue(v.result);           // формула
    if ('hyperlink' in v) return v.hyperlink || v.text || '';
    if ('richText' in v) return v.richText.map((t) => t.text).join('');
    if ('text' in v) return String(v.text);
  }
  return typeof v === 'string' ? v.trim() : v;
}

function parseNumber(v) {
  if (typeof v === 'number') return v;
  const n = Number(String(v).replace(/[€\s]/g, '').replace(/\.(?=\d{3}\b)/g, '').replace(',', '.'));
  return Number.isFinite(n) ? n : NaN;
}

function parseDate(v) {
  if (!v) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(v)) return v;
  const m = String(v).match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
  return m ? `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}` : '';
}

async function loadPromotions(sendDate) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(rel(config.files.promotions));
  const ws = wb.worksheets[0];

  const colIndex = {};
  ws.getRow(1).eachCell((cell, i) => {
    const h = String(cellValue(cell.value)).toLowerCase();
    for (const [key, aliases] of Object.entries(COLUMN_ALIASES)) {
      if (aliases.includes(h)) colIndex[key] = i;
    }
  });
  for (const req of ['produkt', 'alter_preis', 'neuer_preis', 'gueltig_bis']) {
    if (!colIndex[req]) throw new Error(`В Excel нет обязательной колонки «${COLUMN_ALIASES[req][0]}»`);
  }

  const weekEnd = addDays(sendDate, 6);
  const active = [];
  const warnings = [];
  ws.eachRow((row, r) => {
    if (r === 1) return;
    const get = (k) => (colIndex[k] ? cellValue(row.getCell(colIndex[k]).value) : '');
    const p = {
      id: `R${r}`,
      produkt: String(get('produkt')),
      marke: String(get('marke')),
      alter_preis: parseNumber(get('alter_preis')),
      neuer_preis: parseNumber(get('neuer_preis')),
      gueltig_von: parseDate(get('gueltig_von')),
      gueltig_bis: parseDate(get('gueltig_bis')),
      shop: String(get('shop')),
      bedingungen: String(get('bedingungen')),
      link: String(get('link')),
      bild: String(get('bild')),
    };
    if (!p.produkt) return; // пустая строка

    const problems = [];
    if (!Number.isFinite(p.alter_preis) || !Number.isFinite(p.neuer_preis)) problems.push('цена не число');
    else if (p.neuer_preis >= p.alter_preis) problems.push('новая цена не ниже старой');
    if (!p.gueltig_bis) problems.push('нет даты окончания');
    if (problems.length) { warnings.push(`Строка ${r} (${p.produkt}): ${problems.join(', ')}`); return; }

    if (p.gueltig_bis < sendDate) return;                // уже закончилась
    if (p.gueltig_von && p.gueltig_von > weekEnd) return; // начнётся позже этой недели

    p.rabatt_prozent = Math.round((1 - p.neuer_preis / p.alter_preis) * 100);
    active.push(p);
  });

  active.sort((a, b) => b.rabatt_prozent - a.rabatt_prozent || a.gueltig_bis.localeCompare(b.gueltig_bis));
  if (active.length > config.limits.max_items) {
    warnings.push(`Активных акций ${active.length}, в письмо попадут первые ${config.limits.max_items}`);
  }
  return { promotions: active.slice(0, config.limits.max_items), warnings };
}

// ---------- Подписчики (CSV, в т. ч. экспорт Google Контактов) ----------

function parseCsv(text) {
  const rows = [];
  let row = [], field = '', quoted = false;
  const sep = text.split('\n')[0].includes(';') && !text.split('\n')[0].includes(',') ? ';' : ',';
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (quoted) {
      if (c === '"' && text[i + 1] === '"') { field += '"'; i++; }
      else if (c === '"') quoted = false;
      else field += c;
    } else if (c === '"') quoted = true;
    else if (c === sep) { row.push(field); field = ''; }
    else if (c === '\n' || c === '\r') {
      if (c === '\r' && text[i + 1] === '\n') i++;
      row.push(field); rows.push(row); row = []; field = '';
    } else field += c;
  }
  if (field || row.length) { row.push(field); rows.push(row); }
  const header = (rows.shift() || []).map((h) => h.replace(/^﻿/, '').trim().toLowerCase());
  return rows.filter((r) => r.some(Boolean)).map((r) => Object.fromEntries(header.map((h, i) => [h, (r[i] || '').trim()])));
}

function loadSubscribers() {
  const file = rel(config.files.subscribers);
  const unsubFile = rel(config.files.unsubscribed);
  const unsub = new Set(
    fs.existsSync(unsubFile)
      ? fs.readFileSync(unsubFile, 'utf8').split(/[\r\n,;]+/).map((s) => s.trim().toLowerCase()).filter((s) => s.includes('@'))
      : [],
  );
  const { require_label, b2b_label } = config.subscribers;
  const seen = new Set();
  const list = [];
  for (const r of parseCsv(fs.readFileSync(file, 'utf8'))) {
    // Простой формат: email,vorname,nachname,segment
    // Google Контакты: "First Name","Last Name","Labels","E-mail 1 - Value"
    const email = (r.email || r['e-mail 1 - value'] || '').split(':::')[0].trim().toLowerCase();
    if (!email.includes('@') || seen.has(email) || unsub.has(email)) continue;
    const labels = r.labels || r.segment || '';
    if (require_label && !labels.toLowerCase().includes(require_label.toLowerCase())) continue;
    seen.add(email);
    list.push({
      email,
      vorname: r.vorname || r['first name'] || '',
      nachname: r.nachname || r['last name'] || '',
      segment: labels.toUpperCase().includes(b2b_label.toUpperCase()) ? 'B2B' : 'B2C',
    });
  }
  return list;
}

// ---------- Claude ----------

const FORBIDDEN_IN_TEXT = /€|%|\beuro\b|\bprozent\b|\d+[.,]\d{2}\b/i;

function validateCopy(copy, promotions) {
  const errs = [];
  if (copy.status !== 'ok') return [`Модель вернула needs_review: ${(copy.issues || []).join('; ')}`];
  if (!copy.subject || copy.subject.length > 60) errs.push('тема пустая или длиннее 60 знаков');
  if (!copy.preheader || copy.preheader.length > 90) errs.push('прехедер пустой или длиннее 90 знаков');
  const ids = new Set(promotions.map((p) => p.id));
  const got = new Set((copy.items || []).map((i) => i.id));
  for (const id of ids) if (!got.has(id)) errs.push(`нет описания для ${id}`);
  for (const id of got) if (!ids.has(id)) errs.push(`лишний id ${id}`);
  const texts = [copy.subject, copy.preheader, copy.intro, copy.closing, ...(copy.items || []).map((i) => i.teaser)];
  for (const t of texts) {
    if (FORBIDDEN_IN_TEXT.test(t || '')) errs.push(`в тексте есть цена или процент: «${t}»`);
    if (/!{2,}/.test(t || '')) errs.push(`несколько «!» подряд: «${t}»`);
  }
  for (const i of copy.items || []) if ((i.teaser || '').length > 120) errs.push(`описание ${i.id} длиннее 120 знаков`);
  return errs;
}

async function writeCopyWithClaude(sendDate, segment, promotions) {
  const { default: Anthropic } = await import('@anthropic-ai/sdk');
  const client = new Anthropic();
  const system = fs.readFileSync(rel('prompts/system_prompt.md'), 'utf8');
  const input = {
    send_date: sendDate,
    segment,
    promotions: promotions.map(({ link, bild, gueltig_von, ...p }) => p),
  };
  const messages = [{ role: 'user', content: JSON.stringify(input, null, 2) }];

  for (let attempt = 1; attempt <= 2; attempt++) {
    const res = await client.messages.create({
      model: process.env.ANTHROPIC_MODEL || 'claude-sonnet-5-5',
      max_tokens: 4000,
      system,
      messages,
    });
    const text = res.content.filter((b) => b.type === 'text').map((b) => b.text).join('');
    let copy;
    try {
      copy = JSON.parse(text.slice(text.indexOf('{'), text.lastIndexOf('}') + 1));
    } catch {
      copy = null;
    }
    const errs = copy ? validateCopy(copy, promotions) : ['ответ не является JSON'];
    if (!errs.length) return { copy, errors: [] };
    if (copy?.status === 'needs_review' || attempt === 2) return { copy, errors: errs };
    // Вторая попытка: показываем модели её ошибки
    messages.push({ role: 'assistant', content: text });
    messages.push({ role: 'user', content: `Проверка не пройдена:\n- ${errs.join('\n- ')}\nИсправь и верни только JSON.` });
  }
}

// Шаблонные тексты для теста без API-ключа (--no-llm)
function fallbackCopy(sendDate, segment, promotions) {
  const b2b = segment === 'B2B';
  return {
    status: 'ok',
    issues: [],
    subject: `Ihre Wochenangebote ab ${fmtDate(sendDate)}`.slice(0, 60),
    preheader: `${promotions.length} aktuelle Aktionen${b2b ? ' für Ihr Sortiment' : ''} – jetzt im Überblick`,
    intro: b2b
      ? 'nachfolgend finden Sie die aktuellen Aktionen dieser Woche für Ihr Sortiment.'
      : 'nachfolgend finden Sie die aktuellen Angebote dieser Woche.',
    highlight_ids: promotions.slice(0, config.limits.highlights).map((p) => p.id),
    items: promotions.map((p) => ({
      id: p.id,
      teaser: `${p.produkt}${p.marke ? ` von ${p.marke}` : ''}${p.bedingungen ? ` – ${p.bedingungen}` : ''}.`.slice(0, 120),
    })),
    closing: 'Bei Fragen zu den Angeboten steht Ihnen unser Team gern zur Verfügung.',
  };
}

// ---------- Рендер HTML ----------

const template = Handlebars.compile(fs.readFileSync(rel('templates/newsletter.html'), 'utf8'));

function salutation(sub) {
  const name = [sub?.vorname, sub?.nachname].filter(Boolean).join(' ');
  return name ? `Guten Tag ${name},` : 'Guten Tag,';
}

function unsubscribeUrl(email) {
  const to = process.env.UNSUBSCRIBE_EMAIL || 'abmelden@example.com';
  return `mailto:${to}?subject=${encodeURIComponent('Newsletter abbestellen')}&body=${encodeURIComponent(email || '')}`;
}

function render(draft, sub) {
  const byId = Object.fromEntries(draft.promotions.map((p) => [p.id, {
    ...p,
    neuer_preis_fmt: fmtEur(p.neuer_preis),
    alter_preis_fmt: fmtEur(p.alter_preis),
    gueltig_bis_fmt: fmtDate(p.gueltig_bis),
  }]));
  const teaser = Object.fromEntries(draft.copy.items.map((i) => [i.id, i.teaser]));
  return template({
    ...draft.copy,
    brand: config.brand,
    legal: config.legal,
    price_reference_label: config.price_reference_label,
    week_label: `KW ${isoWeek(draft.send_date)} · ab ${fmtDate(draft.send_date)}`,
    is_b2b: draft.segment === 'B2B',
    salutation: salutation(sub),
    highlights: (draft.copy.highlight_ids || []).slice(0, config.limits.highlights).map((id) => byId[id]).filter(Boolean),
    items: draft.promotions.map((p) => ({ ...byId[p.id], teaser: teaser[p.id] })),
    unsubscribe_url: unsubscribeUrl(sub?.email),
  });
}

// ---------- Почта ----------

async function transport() {
  const { default: nodemailer } = await import('nodemailer');
  for (const k of ['SMTP_HOST', 'SMTP_USER', 'SMTP_PASS', 'MAIL_FROM']) {
    if (!process.env[k]) throw new Error(`Не задан ${k} в .env`);
  }
  return nodemailer.createTransport({
    host: process.env.SMTP_HOST,
    port: Number(process.env.SMTP_PORT || 587),
    auth: { user: process.env.SMTP_USER, pass: process.env.SMTP_PASS },
  });
}

const htmlToText = (html) => html
  .replace(/<style[\s\S]*?<\/style>/gi, '')
  .replace(/<br\s*\/?>|<\/(p|tr|h2|div)>/gi, '\n')
  .replace(/<[^>]+>/g, '')
  .replace(/&nbsp;/g, ' ').replace(/&middot;/g, '·').replace(/&minus;/g, '−').replace(/&rsaquo;/g, '›')
  .replace(/&#\d+;|&[a-z]+;/gi, '')
  .replace(/[ \t]+/g, ' ').replace(/\n\s*\n+/g, '\n\n').trim();

// ---------- Состояние ----------

const draftDir = (date) => path.join(STATE_DIR, date);
const readJson = (f, d) => (fs.existsSync(f) ? JSON.parse(fs.readFileSync(f, 'utf8')) : d);
const writeJson = (f, v) => { fs.mkdirSync(path.dirname(f), { recursive: true }); fs.writeFileSync(f, JSON.stringify(v, null, 2)); };

function allDraftDates() {
  return fs.existsSync(STATE_DIR) ? fs.readdirSync(STATE_DIR).filter((d) => /^\d{4}-\d{2}-\d{2}$/.test(d)).sort() : [];
}

// ---------- Команды ----------

async function cmdDraft(args) {
  const sendDate = args.date || nextSendDate();
  // Без ключа Claude агент работает бесплатно: тексты берутся из шаблона
  const useLlm = !args['no-llm'] && Boolean(process.env.ANTHROPIC_API_KEY);
  if (!useLlm) console.log('Режим без ИИ: шаблонные тексты (ANTHROPIC_API_KEY не задан или указан --no-llm)');
  const dir = draftDir(sendDate);
  const meta = readJson(path.join(dir, 'meta.json'), {});
  if (meta.status === 'sent') throw new Error(`Рассылка на ${sendDate} уже отправлена`);

  const { promotions, warnings } = await loadPromotions(sendDate);
  const subscribers = loadSubscribers();
  const segments = [...new Set(subscribers.map((s) => s.segment))].sort();
  console.log(`Дата отправки: ${sendDate} · активных акций: ${promotions.length} · подписчиков: ${subscribers.length} (${segments.join(', ') || '—'})`);
  warnings.forEach((w) => console.log(`  ⚠ ${w}`));

  const result = { send_date: sendDate, status: 'draft', created_at: new Date().toISOString(), warnings, segments: {} };
  if (!promotions.length) {
    result.status = 'needs_review';
    warnings.push('Нет активных акций — письмо не будет отправлено');
  }

  for (const segment of promotions.length ? segments : []) {
    const { copy, errors } = useLlm
      ? await writeCopyWithClaude(sendDate, segment, promotions)
      : { copy: fallbackCopy(sendDate, segment, promotions), errors: [] };
    const draft = { send_date: sendDate, segment, promotions, copy };
    writeJson(path.join(dir, `${segment}.json`), draft);
    if (errors.length) {
      result.status = 'needs_review';
      warnings.push(...errors.map((e) => `[${segment}] ${e}`));
      continue;
    }
    const preview = render(draft, { vorname: 'Max', nachname: 'Mustermann', email: 'vorschau@example.com' });
    fs.writeFileSync(path.join(dir, `${segment}.preview.html`), preview);
    result.segments[segment] = { subject: copy.subject, recipients: subscribers.filter((s) => s.segment === segment).length };
    console.log(`  ✓ ${segment}: «${copy.subject}» → state/${sendDate}/${segment}.preview.html`);
  }
  writeJson(path.join(dir, 'meta.json'), result);

  if (!args['no-mail'] && process.env.ADMIN_EMAIL) await mailPreviewToAdmin(result);
  console.log(result.status === 'draft'
    ? `\nЧерновик готов. Проверьте превью и выполните:  node agent.mjs approve ${sendDate}`
    : `\nНУЖНА ПРОВЕРКА — исправьте Excel и снова запустите draft. Проблемы:\n- ${warnings.join('\n- ')}`);
}

async function mailPreviewToAdmin(meta) {
  const t = await transport();
  const notice = (segment) => `<div style="background:#fff3cd;border-bottom:2px solid #e0b000;padding:12px 16px;font:14px Arial,sans-serif;color:#5c4400;">
    <b>ENTWURF ${segment}</b> – Versand am ${fmtDate(meta.send_date)}, ${config.schedule.hour}:00 Uhr (Berlin).<br>
    Freigabe: <code>node agent.mjs approve ${meta.send_date}</code>
    ${meta.warnings.length ? `<br><br><b>Hinweise:</b><br>${meta.warnings.map((w) => `• ${Handlebars.escapeExpression(w)}`).join('<br>')}` : ''}
  </div>`;
  const segments = Object.keys(meta.segments);
  if (!segments.length) {
    await t.sendMail({
      from: process.env.MAIL_FROM, to: process.env.ADMIN_EMAIL,
      subject: `[NEWSLETTER PRÜFEN] ${fmtDate(meta.send_date)}`,
      text: `Der Entwurf benötigt eine Prüfung:\n- ${meta.warnings.join('\n- ')}`,
    });
  }
  for (const segment of segments) {
    const html = fs.readFileSync(path.join(draftDir(meta.send_date), `${segment}.preview.html`), 'utf8')
      .replace(/<body[^>]*>/, (m) => m + notice(segment));
    await t.sendMail({
      from: process.env.MAIL_FROM, to: process.env.ADMIN_EMAIL,
      subject: `[ENTWURF ${segment}] ${meta.segments[segment].subject}`,
      html, text: htmlToText(html),
    });
  }
  console.log(`  ✉ Превью отправлено на ${process.env.ADMIN_EMAIL}`);
}

function cmdApprove(args) {
  const date = args._[0] || allDraftDates().filter((d) => readJson(path.join(draftDir(d), 'meta.json'), {}).status === 'draft').pop();
  if (!date) throw new Error('Нет черновика для согласования');
  const f = path.join(draftDir(date), 'meta.json');
  const meta = readJson(f, null);
  if (!meta) throw new Error(`Черновик на ${date} не найден`);
  if (meta.status !== 'draft') throw new Error(`Черновик на ${date} в статусе «${meta.status}», согласовать нельзя`);
  meta.status = 'approved';
  meta.approved_at = new Date().toISOString();
  writeJson(f, meta);
  console.log(`✓ Рассылка на ${date} согласована. Отправка: ${fmtDate(date)} в ${config.schedule.hour}:00 (Берлин).`);
}

async function cmdSend(args) {
  const now = berlinNow();
  const dates = allDraftDates().filter((d) => readJson(path.join(draftDir(d), 'meta.json'), {}).status === 'approved');
  const date = dates.find((d) => args.now || d < now.date || (d === now.date && now.hour >= config.schedule.hour));
  if (!date) {
    console.log(dates.length ? `Согласовано на ${dates.join(', ')}, время отправки ещё не наступило.` : 'Нет согласованных рассылок.');
    return;
  }
  const dir = draftDir(date);
  const meta = readJson(path.join(dir, 'meta.json'));
  const sentFile = path.join(dir, 'sent.json');
  const sent = new Set(readJson(sentFile, []));     // защита от повторной отправки при перезапуске
  const subscribers = loadSubscribers();
  const t = args['dry-run'] ? null : await transport();
  const tag = `newsletter-${date}`;
  let ok = 0, failed = 0;

  for (const segment of Object.keys(meta.segments)) {
    const draft = readJson(path.join(dir, `${segment}.json`));
    for (const sub of subscribers.filter((s) => s.segment === segment && !sent.has(s.email))) {
      const html = render(draft, sub);
      if (args['dry-run']) { console.log(`  [dry-run] ${segment} → ${sub.email}`); ok++; continue; }
      try {
        await t.sendMail({
          from: process.env.MAIL_FROM,
          to: sub.email,
          subject: draft.copy.subject,
          html,
          text: htmlToText(html),
          headers: {
            'List-Unsubscribe': `<${unsubscribeUrl(sub.email)}>`,
            'X-Mailin-Tag': tag,  // группировка статистики открытий в Brevo
          },
        });
        sent.add(sub.email);
        writeJson(sentFile, [...sent]);
        ok++;
      } catch (e) {
        failed++;
        console.error(`  ✗ ${sub.email}: ${e.message}`);
      }
      await new Promise((r) => setTimeout(r, 150));
    }
  }
  if (!args['dry-run'] && !failed) {
    meta.status = 'sent';
    meta.sent_at = new Date().toISOString();
    writeJson(path.join(dir, 'meta.json'), meta);
  }
  console.log(`Готово: отправлено ${ok}, ошибок ${failed}.${failed ? ' Запустите send ещё раз — уже отправленные адреса будут пропущены.' : ''}`);
  if (!args['dry-run']) console.log(`Открытия смотрите в Brevo → Transactional → Statistics, тег «${tag}».`);
}

function cmdStatus() {
  const dates = allDraftDates();
  if (!dates.length) return console.log('Черновиков пока нет.');
  for (const d of dates) {
    const m = readJson(path.join(draftDir(d), 'meta.json'), {});
    const sent = readJson(path.join(draftDir(d), 'sent.json'), []).length;
    console.log(`${d}  ${m.status?.padEnd(12)}  сегменты: ${Object.keys(m.segments || {}).join(', ') || '—'}  отправлено: ${sent}`);
  }
}

// ---------- CLI ----------

function parseArgs(argv) {
  const args = { _: [] };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (!a.startsWith('--')) args._.push(a);
    else if (a === '--date') args.date = argv[++i];
    else args[a.slice(2)] = true;
  }
  return args;
}

const [cmd, ...rest] = process.argv.slice(2);
const args = parseArgs(rest);
const commands = { draft: cmdDraft, approve: cmdApprove, send: cmdSend, status: cmdStatus };
if (!commands[cmd]) {
  console.log('Использование: node agent.mjs <draft|approve|send|status> [опции]');
  process.exit(1);
}
try {
  await commands[cmd](args);
} catch (e) {
  console.error(`Ошибка: ${e.message}`);
  process.exit(1);
}
