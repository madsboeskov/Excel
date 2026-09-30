'use strict';

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function labelFor(question, value) {
  const option = question.options.find((candidate) => candidate.value === value);
  return option ? option.label : String(value);
}

// Svaret som det skal vises for et menneske (mail, CSV, kvitteringsside).
function formatAnswer(question, value) {
  if (value == null || value === '') return '–';
  switch (question.type) {
    case 'consent':
      return value ? 'Ja' : 'Nej';
    case 'checkbox':
      return Array.isArray(value) && value.length > 0
        ? value.map((item) => labelFor(question, item)).join(', ')
        : '–';
    case 'rating':
      return `${value} af ${question.max}`;
    case 'select':
    case 'radio':
      return labelFor(question, value);
    default:
      return String(value);
  }
}

function answersHtml(survey, answers) {
  const rows = survey.questions
    .map((question) => {
      const value = formatAnswer(question, answers[question.id]);
      return `<tr>
        <td style="padding:8px 12px;border-bottom:1px solid #eceef2;color:#6b7280;font-size:13px;vertical-align:top;width:45%;">${escapeHtml(question.label)}</td>
        <td style="padding:8px 12px;border-bottom:1px solid #eceef2;font-size:14px;vertical-align:top;white-space:pre-wrap;">${escapeHtml(value)}</td>
      </tr>`;
    })
    .join('\n');
  return `<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;border:1px solid #eceef2;border-radius:8px;">
    <tr><td colspan="2" style="padding:10px 12px;background:#fafbfc;border-bottom:1px solid #eceef2;font-size:13px;font-weight:600;">Dine svar</td></tr>
    ${rows}
  </table>`;
}

function answersText(survey, answers) {
  const lines = survey.questions.map((question) => {
    const value = formatAnswer(question, answers[question.id]).replace(/\n/g, '\n    ');
    const label = /[?:!.]$/.test(question.label) ? question.label : `${question.label}:`;
    return `${label}\n    ${value}`;
  });
  return ['Dine svar', '---------', ...lines].join('\n');
}

// Erstatter {{noegle}} i skabeloner. Ukendte nøgler bliver stående, så
// slåfejl i en skabelon er synlige i stedet for lydløst at blive til tom tekst.
function renderTemplate(template, values) {
  return String(template).replace(/\{\{\s*([a-zA-Z0-9_.]+)\s*\}\}/g, (match, key) =>
    Object.prototype.hasOwnProperty.call(values, key) ? String(values[key]) : match,
  );
}

function templateValues(survey, answers, options = {}) {
  const now = options.date instanceof Date ? options.date : new Date();
  const email = String(answers[survey.emailField] ?? '');
  const name = survey.nameField ? String(answers[survey.nameField] ?? '') : '';
  const values = {
    surveyTitle: survey.title,
    name: name || 'og velkommen',
    email,
    senderName: options.senderName || survey.title,
    date: new Intl.DateTimeFormat('da-DK', {
      dateStyle: 'long',
      timeStyle: 'short',
      timeZone: options.timeZone || 'Europe/Copenhagen',
    }).format(now),
    answersHtml: survey.email.includeAnswers ? answersHtml(survey, answers) : '',
    answersText: survey.email.includeAnswers ? answersText(survey, answers) : '',
  };
  for (const question of survey.questions) {
    values[`answer.${question.id}`] = formatAnswer(question, answers[question.id]);
  }
  return values;
}

module.exports = { escapeHtml, formatAnswer, answersHtml, answersText, renderTemplate, templateValues };
