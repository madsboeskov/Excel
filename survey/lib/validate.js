'use strict';

// Bevidst konservativ: ingen kontroltegn, præcis ét @, punktum i domænet.
const EMAIL_RE = /^[^\s@<>,;:"'\\()[\]]{1,64}@[a-zA-Z0-9]([a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(\.[a-zA-Z0-9]([a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)+$/;
const DATE_RE = /^\d{4}-\d{2}-\d{2}$/;
const PHONE_RE = /^[0-9+()\/.\s-]{4,30}$/;
const CONTROL_CHARS = /[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/g;

function isValidEmail(value) {
  return typeof value === 'string' && value.length <= 254 && EMAIL_RE.test(value);
}

function cleanText(value) {
  return String(value).replace(CONTROL_CHARS, '').replace(/\r\n?/g, '\n').trim();
}

// Værdier der ender i mailheadere (navn, emne) må aldrig indeholde linjeskift.
function singleLine(value) {
  return cleanText(value).replace(/\s*\n+\s*/g, ' ').trim();
}

function validateSubmission(survey, body) {
  const errors = {};
  const answers = {};
  const payload = body && typeof body === 'object' && !Array.isArray(body) ? body : {};

  for (const question of survey.questions) {
    const raw = payload[question.id];
    const fail = (message) => { errors[question.id] = message; };

    switch (question.type) {
      case 'text':
      case 'textarea':
      case 'email':
      case 'tel':
      case 'date': {
        const value = raw == null ? '' : cleanText(raw);
        if (!value) {
          if (question.required) fail('Feltet skal udfyldes.');
          answers[question.id] = '';
          break;
        }
        if (value.length > question.maxLength) {
          fail(`Må højst være ${question.maxLength} tegn (du har skrevet ${value.length}).`);
          break;
        }
        if (question.type === 'email' && !isValidEmail(value)) {
          fail('Skriv en gyldig e-mailadresse, fx navn@eksempel.dk.');
          break;
        }
        if (question.type === 'tel' && !PHONE_RE.test(value)) {
          fail('Skriv et gyldigt telefonnummer.');
          break;
        }
        if (question.type === 'date' && !DATE_RE.test(value)) {
          fail('Brug formatet ÅÅÅÅ-MM-DD.');
          break;
        }
        answers[question.id] = question.type === 'textarea' ? value : singleLine(value);
        break;
      }

      case 'number': {
        const value = raw == null ? '' : cleanText(raw);
        if (!value) {
          if (question.required) fail('Feltet skal udfyldes.');
          answers[question.id] = '';
          break;
        }
        const number = Number(value.replace(',', '.'));
        if (!Number.isFinite(number)) {
          fail('Skriv et tal.');
          break;
        }
        if (question.min != null && number < question.min) {
          fail(`Skal være mindst ${question.min}.`);
          break;
        }
        if (question.max != null && number > question.max) {
          fail(`Må højst være ${question.max}.`);
          break;
        }
        answers[question.id] = number;
        break;
      }

      case 'select':
      case 'radio': {
        const value = raw == null ? '' : singleLine(raw);
        if (!value) {
          if (question.required) fail('Vælg en mulighed.');
          answers[question.id] = '';
          break;
        }
        if (!question.options.some((option) => option.value === value)) {
          fail('Vælg en af de viste muligheder.');
          break;
        }
        answers[question.id] = value;
        break;
      }

      case 'checkbox': {
        const list = raw == null || raw === '' ? [] : Array.isArray(raw) ? raw : [raw];
        const values = [...new Set(list.map((item) => singleLine(item)).filter(Boolean))];
        const unknown = values.filter((value) => !question.options.some((option) => option.value === value));
        if (unknown.length > 0) {
          fail('Vælg kun blandt de viste muligheder.');
          break;
        }
        const minSelected = question.minSelected ?? (question.required ? 1 : 0);
        if (values.length < minSelected) {
          fail(minSelected === 1 ? 'Vælg mindst én mulighed.' : `Vælg mindst ${minSelected} muligheder.`);
          break;
        }
        if (question.maxSelected && values.length > question.maxSelected) {
          fail(`Vælg højst ${question.maxSelected} muligheder.`);
          break;
        }
        answers[question.id] = values;
        break;
      }

      case 'rating': {
        if (raw == null || raw === '') {
          if (question.required) fail('Vælg en vurdering.');
          answers[question.id] = '';
          break;
        }
        const number = Number(raw);
        if (!Number.isInteger(number) || number < 1 || number > question.max) {
          fail(`Vælg et tal mellem 1 og ${question.max}.`);
          break;
        }
        answers[question.id] = number;
        break;
      }

      case 'consent': {
        const accepted = raw === true || raw === 'true' || raw === 'on' || raw === 1 || raw === '1';
        if (question.required && !accepted) {
          fail('Feltet skal være markeret.');
          break;
        }
        answers[question.id] = accepted;
        break;
      }

      default:
        fail('Ukendt spørgsmålstype.');
    }
  }

  return { ok: Object.keys(errors).length === 0, errors, answers };
}

module.exports = { validateSubmission, isValidEmail, cleanText, singleLine };
