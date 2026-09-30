'use strict';

const fs = require('fs');
const path = require('path');

const CONFIG_PATH = process.env.SURVEY_CONFIG
  ? path.resolve(process.env.SURVEY_CONFIG)
  : path.join(__dirname, '..', 'config', 'survey.json');

const SINGLE_CHOICE = new Set(['select', 'radio']);
const TEXTUAL = new Set(['text', 'email', 'textarea', 'tel', 'date', 'number']);
const KNOWN_TYPES = new Set([...SINGLE_CHOICE, ...TEXTUAL, 'checkbox', 'rating', 'consent']);

function normaliseOptions(options, questionId) {
  if (!Array.isArray(options)) return [];
  return options.map((option, index) => {
    if (typeof option === 'string') return { value: option, label: option };
    if (option && typeof option === 'object' && option.value != null) {
      return { value: String(option.value), label: String(option.label ?? option.value) };
    }
    throw new Error(`Spørgsmål "${questionId}": svarmulighed nr. ${index + 1} skal være en tekst eller {value,label}.`);
  });
}

function normaliseQuestion(raw, index) {
  const id = String(raw.id ?? '').trim();
  if (!/^[a-zA-Z0-9_-]+$/.test(id)) {
    throw new Error(`Spørgsmål nr. ${index + 1}: "id" mangler eller indeholder ugyldige tegn (brug a-z, 0-9, _ og -).`);
  }
  const type = String(raw.type ?? 'text').trim();
  if (!KNOWN_TYPES.has(type)) {
    throw new Error(`Spørgsmål "${id}": ukendt type "${type}". Gyldige typer: ${[...KNOWN_TYPES].join(', ')}.`);
  }

  const question = {
    id,
    type,
    label: String(raw.label ?? id),
    help: raw.help ? String(raw.help) : '',
    placeholder: raw.placeholder ? String(raw.placeholder) : '',
    required: raw.required === true,
    options: normaliseOptions(raw.options, id),
  };

  if (SINGLE_CHOICE.has(type) || type === 'checkbox') {
    if (question.options.length === 0) {
      throw new Error(`Spørgsmål "${id}": typen "${type}" kræver mindst én svarmulighed i "options".`);
    }
  }
  if (type === 'rating') {
    question.max = Number.isInteger(raw.max) && raw.max >= 2 && raw.max <= 10 ? raw.max : 5;
    question.minLabel = raw.minLabel ? String(raw.minLabel) : '';
    question.maxLabel = raw.maxLabel ? String(raw.maxLabel) : '';
  }
  if (type === 'number') {
    if (raw.min != null) question.min = Number(raw.min);
    if (raw.max != null) question.max = Number(raw.max);
  }
  if (TEXTUAL.has(type)) {
    const fallback = type === 'textarea' ? 2000 : 250;
    question.maxLength = Number.isInteger(raw.maxLength) && raw.maxLength > 0 ? Math.min(raw.maxLength, 10000) : fallback;
  }
  if (type === 'checkbox') {
    if (Number.isInteger(raw.maxSelected) && raw.maxSelected > 0) question.maxSelected = raw.maxSelected;
    if (Number.isInteger(raw.minSelected) && raw.minSelected > 0) question.minSelected = raw.minSelected;
  }
  return question;
}

function loadSurvey() {
  let raw;
  try {
    raw = JSON.parse(fs.readFileSync(CONFIG_PATH, 'utf8'));
  } catch (error) {
    throw new Error(`Kunne ikke læse spørgeskemaet fra ${CONFIG_PATH}: ${error.message}`);
  }

  const questions = (Array.isArray(raw.questions) ? raw.questions : []).map(normaliseQuestion);
  if (questions.length === 0) throw new Error('Spørgeskemaet indeholder ingen spørgsmål.');

  const seen = new Set();
  for (const question of questions) {
    if (seen.has(question.id)) throw new Error(`Spørgsmålet "${question.id}" optræder mere end én gang.`);
    seen.add(question.id);
  }

  const emailField = String(raw.emailField ?? 'email');
  const emailQuestion = questions.find((question) => question.id === emailField);
  if (!emailQuestion) {
    throw new Error(`"emailField" peger på "${emailField}", men der findes ikke et spørgsmål med det id.`);
  }
  if (emailQuestion.type !== 'email') {
    throw new Error(`Spørgsmålet "${emailField}" skal have typen "email", så kvitteringen kan sendes.`);
  }
  if (!emailQuestion.required) {
    throw new Error(`Spørgsmålet "${emailField}" skal være påkrævet ("required": true), ellers kan kvitteringen ikke sendes.`);
  }

  const nameField = raw.nameField ? String(raw.nameField) : '';
  if (nameField && !questions.some((question) => question.id === nameField)) {
    throw new Error(`"nameField" peger på "${nameField}", men der findes ikke et spørgsmål med det id.`);
  }

  return {
    title: String(raw.title ?? 'Spørgeskema'),
    intro: String(raw.intro ?? ''),
    submitLabel: String(raw.submitLabel ?? 'Send'),
    privacyNote: String(raw.privacyNote ?? ''),
    emailField,
    nameField,
    thankYou: {
      heading: String(raw.thankYou?.heading ?? 'Tak for dit svar!'),
      body: String(raw.thankYou?.body ?? 'Vi har sendt en bekræftelse til {{email}}.'),
    },
    email: {
      subject: String(raw.email?.subject ?? 'Tak for dit svar – {{surveyTitle}}'),
      replyTo: String(raw.email?.replyTo ?? ''),
      includeAnswers: raw.email?.includeAnswers !== false,
    },
    questions,
  };
}

module.exports = { loadSurvey, CONFIG_PATH };
