'use strict';

const fs = require('fs');
const fsp = require('fs/promises');
const path = require('path');
const nodemailer = require('nodemailer');

const { renderTemplate, templateValues } = require('./render');
const { isValidEmail, singleLine } = require('./validate');

const TEMPLATE_DIR = process.env.TEMPLATE_DIR
  ? path.resolve(process.env.TEMPLATE_DIR)
  : path.join(__dirname, '..', 'templates');
const OUTBOX_DIR = path.join(__dirname, '..', 'data', 'outbox');

function readTemplate(name) {
  return fs.readFileSync(path.join(TEMPLATE_DIR, name), 'utf8');
}

function truthy(value) {
  return /^(1|true|yes|ja|on)$/i.test(String(value ?? '').trim());
}

function mailConfig() {
  const host = (process.env.SMTP_HOST || '').trim();
  const mode = (process.env.MAIL_TRANSPORT || (host ? 'smtp' : 'console')).trim().toLowerCase();
  const port = Number(process.env.SMTP_PORT || 587);
  return {
    mode,
    host,
    port,
    secure: process.env.SMTP_SECURE == null ? port === 465 : truthy(process.env.SMTP_SECURE),
    user: (process.env.SMTP_USER || '').trim(),
    pass: process.env.SMTP_PASS || '',
    fromAddress: (process.env.MAIL_FROM || process.env.SMTP_USER || 'ingen-svar@eksempel.dk').trim(),
    fromName: singleLine(process.env.MAIL_FROM_NAME || 'Spørgeskema'),
    replyTo: (process.env.MAIL_REPLY_TO || '').trim(),
    copyTo: (process.env.MAIL_COPY_TO || '').trim(),
  };
}

function createTransport(config = mailConfig()) {
  if (config.mode === 'smtp') {
    if (!config.host) throw new Error('SMTP_HOST mangler. Sæt den i .env, eller vælg MAIL_TRANSPORT=console.');
    return nodemailer.createTransport({
      host: config.host,
      port: config.port,
      secure: config.secure,
      auth: config.user ? { user: config.user, pass: config.pass } : undefined,
      connectionTimeout: 15000,
      greetingTimeout: 15000,
      socketTimeout: 30000,
    });
  }
  // console: intet sendes ud af huset – mailen skrives til data/outbox/ og logges.
  return nodemailer.createTransport({ streamTransport: true, buffer: true, newline: 'unix' });
}

async function saveToOutbox(info, to) {
  await fsp.mkdir(OUTBOX_DIR, { recursive: true });
  const stamp = new Date().toISOString().replace(/[:.]/g, '-');
  const safeTo = to.replace(/[^a-zA-Z0-9@._-]/g, '_');
  const file = path.join(OUTBOX_DIR, `${stamp}_${safeTo}.eml`);
  await fsp.writeFile(file, info.message ?? '');
  return file;
}

function buildMessage(survey, answers, config = mailConfig()) {
  const to = String(answers[survey.emailField] ?? '').trim();
  if (!isValidEmail(to)) throw new Error('Modtagerens e-mailadresse er ikke gyldig.');

  const values = templateValues(survey, answers, { senderName: config.fromName });
  const subject = singleLine(renderTemplate(survey.email.subject, values));
  const replyTo = config.replyTo || survey.email.replyTo || '';

  const message = {
    from: { name: config.fromName, address: config.fromAddress },
    to,
    subject,
    text: renderTemplate(readTemplate('confirmation.txt'), values),
    html: renderTemplate(readTemplate('confirmation.html'), values),
  };
  if (replyTo && isValidEmail(replyTo)) message.replyTo = replyTo;
  if (config.copyTo) {
    const copies = config.copyTo.split(',').map((item) => item.trim()).filter(isValidEmail);
    if (copies.length > 0) message.bcc = copies;
  }
  return message;
}

async function sendConfirmation(survey, answers, options = {}) {
  const config = options.config || mailConfig();
  const transport = options.transport || createTransport(config);
  const message = buildMessage(survey, answers, config);
  const info = await transport.sendMail(message);

  const result = { to: message.to, subject: message.subject, mode: config.mode, messageId: info.messageId };
  if (config.mode !== 'smtp') {
    result.file = await saveToOutbox(info, message.to);
    console.log(`[mail] (console-tilstand) "${message.subject}" til ${message.to} → ${path.relative(process.cwd(), result.file)}`);
  } else {
    console.log(`[mail] sendt til ${message.to} (${info.messageId || 'uden id'})`);
  }
  return result;
}

async function verifyTransport(config = mailConfig()) {
  if (config.mode !== 'smtp') {
    return { ok: true, mode: config.mode, detail: 'console-tilstand: mails gemmes i data/outbox/ og sendes ikke.' };
  }
  const transport = createTransport(config);
  try {
    await transport.verify();
    return { ok: true, mode: 'smtp', detail: `forbundet til ${config.host}:${config.port}` };
  } catch (error) {
    return { ok: false, mode: 'smtp', detail: error.message };
  } finally {
    transport.close?.();
  }
}

module.exports = { mailConfig, createTransport, buildMessage, sendConfirmation, verifyTransport };
