'use strict';

const fs = require('fs');
const fsp = require('fs/promises');
const path = require('path');
const crypto = require('crypto');

const { formatAnswer } = require('./render');

const DATA_DIR = process.env.DATA_DIR ? path.resolve(process.env.DATA_DIR) : path.join(__dirname, '..', 'data');
const RESPONSES_FILE = path.join(DATA_DIR, 'responses.jsonl');

// Én skrivning ad gangen, så to samtidige besvarelser ikke flettes sammen i filen.
let writeQueue = Promise.resolve();

async function appendResponse(record) {
  const line = `${JSON.stringify(record)}\n`;
  writeQueue = writeQueue.then(async () => {
    await fsp.mkdir(DATA_DIR, { recursive: true });
    await fsp.appendFile(RESPONSES_FILE, line, 'utf8');
  });
  await writeQueue;
  return record;
}

function newResponse(answers, meta = {}) {
  return {
    id: crypto.randomUUID(),
    receivedAt: new Date().toISOString(),
    answers,
    meta,
  };
}

// Besvarelsen gemmes, foer mailen forsoeges sendt, saa intet gaar tabt hvis
// mailserveren er nede. Naar forsoeget er slut, opdateres statuslinjen her.
// Filen skrives om i sin helhed - fint ved de maengder et spoergeskema giver,
// og skrivningerne er serialiseret gennem den samme koe som appendResponse.
async function updateResponseMeta(id, patch) {
  writeQueue = writeQueue.then(async () => {
    let raw;
    try {
      raw = await fsp.readFile(RESPONSES_FILE, 'utf8');
    } catch (error) {
      if (error.code === 'ENOENT') return;
      throw error;
    }
    let changed = false;
    const lines = raw.split('\n').map((line) => {
      if (line.trim() === '') return line;
      let record;
      try {
        record = JSON.parse(line);
      } catch {
        return line;
      }
      if (record.id !== id) return line;
      changed = true;
      record.meta = { ...record.meta, ...patch };
      return JSON.stringify(record);
    });
    if (changed) await fsp.writeFile(RESPONSES_FILE, lines.join('\n'), 'utf8');
  });
  await writeQueue;
}

async function readResponses() {
  let raw;
  try {
    raw = await fsp.readFile(RESPONSES_FILE, 'utf8');
  } catch (error) {
    if (error.code === 'ENOENT') return [];
    throw error;
  }
  return raw
    .split('\n')
    .filter((line) => line.trim() !== '')
    .map((line) => {
      try {
        return JSON.parse(line);
      } catch {
        return null;
      }
    })
    .filter(Boolean);
}

function csvCell(value) {
  const text = String(value ?? '');
  return /[";\n\r]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

// Semikolon + BOM, fordi dansk Excel ellers lægger hele rækken i én kolonne.
function toCsv(survey, responses) {
  const header = ['Modtaget', 'Id', ...survey.questions.map((question) => question.label), 'Kvittering sendt'];
  const rows = responses.map((response) => {
    const received = new Date(response.receivedAt);
    const stamp = Number.isNaN(received.getTime())
      ? response.receivedAt
      : new Intl.DateTimeFormat('da-DK', { dateStyle: 'short', timeStyle: 'medium', timeZone: 'Europe/Copenhagen' }).format(received);
    return [
      stamp,
      response.id,
      ...survey.questions.map((question) => formatAnswer(question, response.answers?.[question.id])),
      response.meta?.mailSent ? 'Ja' : 'Nej',
    ];
  });
  const body = [header, ...rows].map((row) => row.map(csvCell).join(';')).join('\r\n');
  return `﻿${body}\r\n`;
}

function responseCount() {
  try {
    return fs.readFileSync(RESPONSES_FILE, 'utf8').split('\n').filter((line) => line.trim() !== '').length;
  } catch {
    return 0;
  }
}

module.exports = { appendResponse, updateResponseMeta, newResponse, readResponses, toCsv, responseCount, RESPONSES_FILE, DATA_DIR };
