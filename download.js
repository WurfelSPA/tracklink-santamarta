#!/usr/bin/env node
/**
 * TrackGTS Report Downloader v5
 * Usa page.evaluate() + onLoginOn() — mismo enfoque que el flujo n8n que funciona.
 * Sin detección de iframes, sin typing manual de campos.
 */
'use strict';

const puppeteer = require('puppeteer');
const AdmZip    = require('adm-zip');
const fs        = require('fs');
const path      = require('path');

async function main() {
  const { TL_USER, TL_PASSWORD, TL_DOMAIN } = process.env;
  if (!TL_USER || !TL_PASSWORD || !TL_DOMAIN) {
    throw new Error('Faltan variables de entorno: TL_USER, TL_PASSWORD, TL_DOMAIN');
  }

  console.log('=== TrackGTS Report Downloader v5 ===');

  const browser = await puppeteer.launch({
    headless: true,
    args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage'],
  });

  try {
    const page = await browser.newPage();
    page.setDefaultTimeout(60_000);

    // ── 1. Abrir página de login ──────────────────────────────────────────
    const loginUrl = `https://${TL_DOMAIN}.trackgts.com/admin/login.html`;
    console.log(`[1] Abriendo: ${loginUrl}`);
    await page.goto(loginUrl, { waitUntil: 'networkidle0', timeout: 60_000 });
    await page.waitForSelector('#username', { timeout: 30_000 });

    // Forzar idioma y recargar (igual que n8n)
    await page.evaluate(() => localStorage.setItem('sltLanguage', '0'));
    await page.reload({ waitUntil: 'networkidle0' });
    await page.waitForSelector('#username', { timeout: 30_000 });
    console.log('[1] Página de login lista');

    // ── 2. Login directo vía onLoginOn() (igual que n8n) ──────────────────
    console.log('[2] Ejecutando login...');
    await page.evaluate((user, password, domain) => {
      // Encriptación CryptoJS que usa TrackGTS
      const K  = 'd5fg4df5sg4ds5fg';
      const S  = { a:'1', b:'2', c:'3', d:'4', e:'5', f:'6', g:'7', h:'8', i:'9' };
      const k  = CryptoJS.enc.Utf8.parse(K);
      const iv = CryptoJS.enc.Utf8.parse(K);
      const a  = [];
      for (const c of password) {
        a.push(
          CryptoJS.AES.encrypt(
            CryptoJS.enc.Utf8.parse(S[c] || c), k,
            { iv, mode: CryptoJS.mode.CBC, padding: CryptoJS.pad.Pkcs7 }
          ).toString()
        );
      }
      ARRAYPSWD = a;
      document.getElementById('username').value = user;
      document.getElementById('domain').value   = domain;
      document.getElementById('password').value = '********'; // solo visual
      LOGININPROCESS = false;
      onLoginOn();
    }, TL_USER, TL_PASSWORD, TL_DOMAIN);

    // ── 3. Esperar que la sesión se establezca ────────────────────────────
    console.log('[3] Esperando 15s para que la sesión se establezca...');
    await new Promise(r => setTimeout(r, 15_000));
    console.log(`[3] URL actual: ${page.url()}`);

    // ── 4. Fetch del reporte desde el contexto del browser ────────────────
    const now   = new Date();
    const start = new Date(now);
    start.setDate(start.getDate() - 7);
    const pad = n => String(n).padStart(2, '0');
    const fmt = d =>
      `${d.getFullYear()}/${pad(d.getMonth()+1)}/${pad(d.getDate())}`;

    const startDate = `${fmt(start)} 04:00:00`;
    const endDate   = `${fmt(now)} 03:59:59`;
    console.log(`[4] Rango: ${startDate} → ${endDate}`);

    const result = await page.evaluate(async (startDate, endDate) => {
      const h    = JSONUSER.hash;
      const body = JSON.stringify({
        startDate,
        endDate,
        unitIds:             '4349,6399,4436',
        reportName:          'INFORME EXCESOS DE VELOCIDAD',
        parameters:          'undefined',
        userTimeZone:        -4,
        userfuelMeasure:     0,
        userMeasureDistance: 0,
        speed:               NaN,
        language:            0,
      });

      const res = await fetch(
        `https://www.trackgts.com:82/api/reportTravel/GetSpeedingReportByUnitsPagesZip/25/${h}`,
        { method: 'POST', headers: { 'Content-Type': 'application/json;charset=UTF-8' }, body }
      );
      const json = await res.json();
      if (!json.FileContents)
        return { error: `Sin FileContents: ${JSON.stringify(json).slice(0, 300)}` };
      return { fileContents: json.FileContents };
    }, startDate, endDate);

    if (result.error) throw new Error(result.error);
    console.log('[4] Reporte descargado OK');

    // ── 5. Extraer .xlsx del ZIP ──────────────────────────────────────────
    console.log('[5] Extrayendo .xlsx...');
    const zip  = new AdmZip(Buffer.from(result.fileContents, 'base64'));
    const xlsx = zip.getEntries().find(e => e.entryName.toLowerCase().endsWith('.xlsx'));
    if (!xlsx)
      throw new Error(`Sin .xlsx en el ZIP. Entradas: ${zip.getEntries().map(e => e.entryName).join(', ')}`);

    const dest = path.join(process.cwd(), 'INFORME EXCESOS DE VELOCIDAD.xlsx');
    fs.writeFileSync(dest, xlsx.getData());
    console.log(`[5] Guardado: ${dest}`);
    console.log('=== COMPLETADO ===');

  } finally {
    await browser.close();
  }
}

main().catch(err => {
  console.error('ERROR FATAL:', err.message);
  process.exit(1);
});
