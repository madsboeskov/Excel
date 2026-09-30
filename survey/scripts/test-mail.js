'use strict';

// Sender én testkvittering, så du kan tjekke SMTP-opsætningen uden at
// udfylde skemaet:  npm run test:mail -- din@mail.dk

require('dotenv').config();

const { loadSurvey } = require('../lib/config');
const { verifyTransport, sendConfirmation, mailConfig } = require('../lib/mailer');
const { isValidEmail } = require('../lib/validate');

async function main() {
  const to = process.argv[2];
  if (!to || !isValidEmail(to)) {
    console.error('Brug: npm run test:mail -- din@mail.dk');
    process.exit(1);
  }

  const survey = loadSurvey();
  const config = mailConfig();
  console.log(`Tilstand: ${config.mode}${config.mode === 'smtp' ? ` (${config.host}:${config.port})` : ''}`);

  const status = await verifyTransport(config);
  console.log(`Forbindelse: ${status.ok ? 'OK' : 'FEJL'} – ${status.detail}`);
  if (!status.ok) process.exit(1);

  // Et gyldigt prøvesvar ud fra spørgeskemaets egne felter.
  const answers = {};
  for (const question of survey.questions) {
    if (question.id === survey.emailField) answers[question.id] = to;
    else if (question.id === survey.nameField) answers[question.id] = 'Testperson';
    else if (question.type === 'checkbox') answers[question.id] = question.options.slice(0, 1).map((o) => o.value);
    else if (question.type === 'radio' || question.type === 'select') answers[question.id] = question.options[0].value;
    else if (question.type === 'rating') answers[question.id] = question.max;
    else if (question.type === 'consent') answers[question.id] = true;
    else if (question.type === 'number') answers[question.id] = 1;
    else answers[question.id] = 'Testsvar';
  }

  const result = await sendConfirmation(survey, answers, { config });
  console.log(`Sendt til ${result.to}${result.file ? ` (gemt som ${result.file})` : ''}`);
}

main().catch((error) => {
  console.error('Fejl:', error.message);
  process.exit(1);
});
