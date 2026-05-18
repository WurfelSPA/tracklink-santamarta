#!/usr/bin/env node
/**
 * TrackGTS Report Downloader v2
 * - Login con Puppeteer
 * - Fetch de la API desde DENTRO del browser (usa cookies de sesión automáticamente)
 * - Extrae .xlsx del zip base64 y lo guarda en la raíz del repo
 */
'use strict';

const puppeteer = require('puppeteer');
const AdmZip    = require('adm-zip');
const fs        = require('fs');
const path      = require('path');

function generateRequestId() {
  const uuid = 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, c => {
    const r = Math.random() * 16 | 0;
    return (c === 'x' ? r : (r & 0x3 | 0x8)).toString(16);
  });
  return `${uuid}__t_${String(Date.now()).slice(-4)}`;
}

function getWeekDateRange() {
  const now = new Date();
  const pad = n => String(n).padStart(2, '0');
  const fmt = d =>
    `${d.getFullYear()}/${pad(d.getMonth()+1)}/${pad(d.getDate())} ` +
    `${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;

  const start = new Date(now);
  start.setDate(start.getDate() - 8);
  start.setHours(4, 0, 0, 0);

  const end = new Date(now);
  end.setDate(end.getDate() - 1);
  end.setHours(3, 59, 59, 0);

  return { startDate: fmt(start), endDate: fmt(end) };
}

async function main() {
  const { TL_USER, TL_PASSWORD, TL_DOMAIN } = process.env;
  if (!TL_USER || !TL_PASSWORD || !TL_DOMAIN) {
    throw new Error('Faltan variables de entorno: TL_USER, TL_PASSWORD, TL_DOMAIN');
  }

  console.log('Iniciando descarga de reporte TrackGTS...');

  const browser = await puppeteer.launch({
    headless: true,
    args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage'],
  });

  try {
    const page = await browser.newPage();
    page.setDefaultTimeout(60_000);

    // ── 1. Login ──────────────────────────────────────────────────────────
    console.log('Abriendo pagina de login...');
    await page.goto(`https://${TL_DOMAIN}.trackgts.com/admin/login.html`, {
      waitUntil: 'networkidle2',
      timeout: 30_000,
    });

    console.log('Llenando credenciales...');
    await page.evaluate((u, p, d) => {
      const set = (sel, val) => { const el = document.querySelector(sel); if (el) el.value = val; };
      set('#user',     u);
      set('#password', p);
      set('#domain',   d);
    }, TL_USER, TL_PASSWORD, TL_DOMAIN);

    console.log('Llamando onLoginOn()...');
    await page.evaluate(() => {
      if (typeof onLoginOn === 'function') return onLoginOn();
      if (typeof loginOn   === 'function') return loginOn();
      document.querySelector('button[type=submit], input[type=submit]')?.click();
    });

    console.log('Esperando 25s para que la sesion se establezca...');
    await new Promise(r => setTimeout(r, 25_000));

    const urlDespuesLogin = page.url();
    console.log(`URL actual: ${urlDespuesLogin}`);

    // ── 2. Fetch de la API DESDE el browser (cookies automaticas) ─────────
    const requestId = generateRequestId();
    const { startDate, endDate } = getWeekDateRange();
    const apiUrl = `https://www.trackgts.com:82/api/reportTravel/GetSpeedingReportByUnitsPagesZip/25/${requestId}`;

    // Nota: "speed":NaN es requerido por la API (no es JSON estandar)
    const body =
      `{"startDate":"${startDate}","endDate":"${endDate}",` +
      `"unitIds":"4349,6399,4436","reportName":"INFORME EXCESOS DE VELOCIDAD",` +
      `"parameters":"undefined","userTimeZone":-4,"userfuelMeasure":0,` +
      `"userMeasureDistance":0,"speed":NaN,"language":0}`;

    console.log(`POST -> ${apiUrl}`);
    console.log(`Rango: ${startDate} -> ${endDate}`);

    // El fetch corre en el contexto del browser: las cookies de sesion se envian automaticamente
    const result = await page.evaluate(async (url, reqBody) => {
      try {
        const res = await fetch(url, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json;charset=UTF-8' },
          body: reqBody,
        });

        let json;
        const text = await res.text();
        try { json = JSON.parse(text); }
        catch { return { error: `Respuesta no-JSON (HTTP ${res.status}): ${text.slice(0, 400)}` }; }

        if (!json.FileContents) {
          return { error: `Sin FileContents: ${JSON.stringify(json).slice(0, 500)}` };
        }

        return { fileContents: json.FileContents };
      } catch (e) {
        return { error: e.message };
      }
    }, apiUrl, body);

    if (result.error) throw new Error(result.error);

    // ── 3. Extraer .xlsx del zip base64 ───────────────────────────────────
    console.log('Extrayendo .xlsx del zip...');
    const zip   = new AdmZip(Buffer.from(result.fileContents, 'base64'));
    const xlsx  = zip.getEntries().find(e => e.entryName.toLowerCase().endsWith('.xlsx'));

    if (!xlsx) {
      const names = zip.getEntries().map(e => e.entryName).join(', ');
      throw new Error(`No hay .xlsx en el zip. Encontrado: ${names}`);
    }

    const dest = path.join(process.cwd(), 'INFORME EXCESOS DE VELOCIDAD.xlsx');
    fs.writeFileSync(dest, xlsx.getData());
    console.log(`OK - Guardado: ${dest}`);

  } finally {
    await browser.close();
  }
}

main().catch(err => {
  console.error('ERROR:', err.message);
  process.exit(1);
});
