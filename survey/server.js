'use strict';

require('dotenv').config();

const crypto = require('crypto');
const path = require('path');
const express = require('express');

const { loadSurvey, CONFIG_PATH } = require('./lib/config');
const { validateSubmission, isValidEmail } = require('./lib/validate');
const { renderTemplate, templateValues } = require('./lib/render');
const { sendConfirmation, verifyTransport, mailConfig, createTransport } = require('./lib/mailer');
const storage = require('./lib/storage');

const PORT = Number(process.env.PORT || 3000);
const HOST = process.env.HOST || '0.0.0.0';
const MAIL_TIMEOUT_MS = Number(process.env.MAIL_TIMEOUT_MS || 20000);
const RATE_LIMIT_MAX = Number(process.env.RATE_LIMIT_MAX || 5);
const RATE_LIMIT_WINDOW_MS = Number(process.env.RATE_LIMIT_WINDOW_MS || 10 * 60 * 1000);
const ADMIN_TOKEN = process.env.ADMIN_TOKEN || '';

let survey;
try {
  survey = loadSurvey();
} catch (error) {
  console.error(`\nSpørgeskemaet kunne ikke indlæses (${CONFIG_PATH}):\n  ${error.message}\n`);
  process.exit(1);
}

const app = express();
app.disable('x-powered-by');
if (process.env.TRUST_PROXY) app.set('trust proxy', process.env.TRUST_PROXY === 'true' ? 1 : process.env.TRUST_PROXY);
app.use(express.json({ limit: '64kb' }));
app.use(express.urlencoded({ extended: false, limit: '64kb' }));

// --- enkel rate limit pr. IP, nok til at bremse gentagne indsendelser ---
const hits = new Map();
setInterval(() => {
  const cutoff = Date.now() - RATE_LIMIT_WINDOW_MS;
  for (const [key, stamps] of hits) {
    const kept = stamps.filter((stamp) => stamp > cutoff);
    if (kept.length === 0) hits.delete(key);
    else hits.set(key, kept);
  }
}, RATE_LIMIT_WINDOW_MS).unref();

function rateLimited(req) {
  const key = req.ip || 'ukendt';
  const now = Date.now();
  const stamps = (hits.get(key) || []).filter((stamp) => stamp > now - RATE_LIMIT_WINDOW_MS);
  stamps.push(now);
  hits.set(key, stamps);
  return stamps.length > RATE_LIMIT_MAX;
}

function timeLimited(promise, ms, message) {
  let timer;
  return Promise.race([
    promise.finally(() => clearTimeout(timer)),
    new Promise((_, reject) => {
      timer = setTimeout(() => reject(new Error(message)), ms);
    }),
  ]);
}

function thankYouFor(answers) {
  const values = templateValues(survey, answers, { senderName: mailConfig().fromName });
  return {
    heading: renderTemplate(survey.thankYou.heading, values),
    body: renderTemplate(survey.thankYou.body, values),
  };
}

function publicSurvey() {
  return {
    title: survey.title,
    intro: survey.intro,
    submitLabel: survey.submitLabel,
    privacyNote: survey.privacyNote,
    emailField: survey.emailField,
    thankYou: survey.thankYou,
    questions: survey.questions,
  };
}

app.get('/api/survey', (req, res) => {
  res.json(publicSurvey());
});

app.post('/api/submit', async (req, res) => {
  // Skjult felt som kun robotter udfylder.
  if (typeof req.body?._hp === 'string' && req.body._hp.trim() !== '') {
    // Robotten faar samme kvitteringsside som et menneske, men intet gemmes eller sendes.
    const claimed = String(req.body?.[survey.emailField] ?? '');
    return res.status(202).json({
      ok: true,
      mailSent: true,
      thankYou: thankYouFor({ [survey.emailField]: isValidEmail(claimed) ? claimed : 'din e-mailadresse' }),
    });
  }
  if (rateLimited(req)) {
    return res.status(429).json({ ok: false, message: 'Du har sendt skemaet flere gange. Prøv igen om lidt.' });
  }

  const { ok, errors, answers } = validateSubmission(survey, req.body);
  if (!ok) {
    return res.status(400).json({ ok: false, message: 'Skemaet mangler noget – se felterne nedenfor.', errors });
  }

  const record = storage.newResponse(answers, {
    mailSent: null,
    userAgent: String(req.get('user-agent') || '').slice(0, 250),
  });

  try {
    await storage.appendResponse(record);
  } catch (error) {
    console.error('[gem] kunne ikke gemme besvarelsen:', error);
    return res.status(500).json({ ok: false, message: 'Dine svar kunne ikke gemmes. Prøv igen om lidt.' });
  }

  let mailSent = false;
  let mailError = '';
  try {
    await timeLimited(
      sendConfirmation(survey, answers),
      MAIL_TIMEOUT_MS,
      `Mailserveren svarede ikke inden for ${Math.round(MAIL_TIMEOUT_MS / 1000)} sekunder.`,
    );
    mailSent = true;
  } catch (error) {
    mailError = error.message;
    console.error(`[mail] kunne ikke sende kvittering til ${answers[survey.emailField]}:`, error.message);
  }

  storage
    .updateResponseMeta(record.id, mailSent ? { mailSent: true } : { mailSent: false, mailError })
    .catch((error) => console.error('[gem] kunne ikke opdatere mailstatus:', error.message));

  res.status(201).json({ ok: true, id: record.id, mailSent, thankYou: thankYouFor(answers) });
});

// --- adgang til besvarelserne (kun med ADMIN_TOKEN) ---
function authorised(req) {
  if (!ADMIN_TOKEN) return false;
  const supplied = String(req.query.token || (req.get('authorization') || '').replace(/^Bearer\s+/i, ''));
  const a = Buffer.from(supplied);
  const b = Buffer.from(ADMIN_TOKEN);
  return a.length === b.length && crypto.timingSafeEqual(a, b);
}

function guard(req, res, next) {
  if (!ADMIN_TOKEN) {
    return res.status(404).json({ ok: false, message: 'Eksport er slået fra. Sæt ADMIN_TOKEN i .env for at bruge den.' });
  }
  if (!authorised(req)) return res.status(401).json({ ok: false, message: 'Forkert eller manglende token.' });
  return next();
}

app.get('/admin/responses.csv', guard, async (req, res, next) => {
  try {
    const responses = await storage.readResponses();
    res.type('text/csv; charset=utf-8');
    res.set('content-disposition', `attachment; filename="besvarelser-${new Date().toISOString().slice(0, 10)}.csv"`);
    res.send(storage.toCsv(survey, responses));
  } catch (error) {
    next(error);
  }
});

app.get('/admin/responses.json', guard, async (req, res, next) => {
  try {
    res.json({ ok: true, responses: await storage.readResponses() });
  } catch (error) {
    next(error);
  }
});

app.get('/api/health', (req, res) => {
  res.json({ ok: true, survey: survey.title, questions: survey.questions.length, responses: storage.responseCount(), mail: mailConfig().mode });
});

app.use(express.static(path.join(__dirname, 'public'), { extensions: ['html'], maxAge: '1h' }));

app.use((req, res) => res.status(404).json({ ok: false, message: 'Siden findes ikke.' }));
app.use((error, req, res, next) => {
  console.error('[fejl]', error);
  if (res.headersSent) return next(error);
  res.status(500).json({ ok: false, message: 'Der skete en uventet fejl.' });
});

if (require.main === module) {
  const config = mailConfig();
  const server = app.listen(PORT, HOST, async () => {
    console.log(`\n  ${survey.title}`);
    console.log(`  Kører på http://localhost:${PORT}`);
    console.log(`  Besvarelser: ${storage.RESPONSES_FILE}`);
    if (ADMIN_TOKEN) console.log(`  Eksport:     http://localhost:${PORT}/admin/responses.csv?token=***`);
    const status = await verifyTransport(config);
    if (config.mode !== 'smtp') {
      console.log(`  Mail:        ${status.detail}`);
      console.log('               Udfyld SMTP_* i .env for at sende rigtige mails.\n');
    } else if (status.ok) {
      console.log(`  Mail:        SMTP klar (${status.detail}), afsender ${config.fromAddress}\n`);
    } else {
      console.log(`  Mail:        SMTP svarer ikke – ${status.detail}`);
      console.log('               Besvarelser gemmes stadig, men kvitteringen fejler.\n');
    }
  });
  const shutdown = () => server.close(() => process.exit(0));
  process.on('SIGINT', shutdown);
  process.on('SIGTERM', shutdown);
}

module.exports = app;
