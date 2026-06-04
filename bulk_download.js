#!/usr/bin/env node
/**
 * bulk_download.js — Descarga los últimos 63 días (9 ventanas de 7 días)
 * en una sola sesión de login. Guarda cada ventana como bulk_0.xlsx … bulk_8.xlsx
 * para que merge_history.py las fusione después.
 */
'use strict';

const puppeteer = require('puppeteer');
const AdmZip    = require('adm-zip');
const fs        = require('fs');
const path      = require('path');

const WINDOWS = 9;   // 9 × 7 días = 63 días de cobertura

const pad = n => String(n).padStart(2, '0');
const fmt = d => `${d.getFullYear()}/${pad(d.getMonth()+1)}/${pad(d.getDate())}`;

async function fetchWindow(page, startDate, endDate, label) {
  console.log(`  [fetch] ${startDate} → ${endDate}`);
  const result = await page.evaluate(async (startDate, endDate) => {
    const h    = JSONUSER.hash;
    const body = JSON.stringify({
      startDate, endDate,
      unitIds:             '4349,6399,4436',
      reportName:          'INFORME EXCESOS DE VELOCIDAD',
      parameters:          'undefined',
      userTimeZone:        -4,
      userfuelMeasure:     0,
      userMeasureDistance: 0,
      speed:               NaN,
      language:            0,
    });
    const res  = await fetch(
      `https://www.trackgts.com:82/api/reportTravel/GetSpeedingReportByUnitsPagesZip/25/${h}`,
      { method:'POST', headers:{'Content-Type':'application/json;charset=UTF-8'}, body }
    );
    const json = await res.json();
    if (!json.FileContents) return { error: JSON.stringify(json).slice(0,200) };
    return { fileContents: json.FileContents };
  }, startDate, endDate);

  if (result.error) {
    console.log(`  [WARN] ventana ${label}: ${result.error}`);
    return false;
  }

  const zip  = new AdmZip(Buffer.from(result.fileContents, 'base64'));
  const xlsx = zip.getEntries().find(e => e.entryName.toLowerCase().endsWith('.xlsx'));
  if (!xlsx) { console.log(`  [WARN] ventana ${label}: sin .xlsx en ZIP`); return false; }

  const dest = path.join(process.cwd(), `bulk_${label}.xlsx`);
  fs.writeFileSync(dest, xlsx.getData());
  console.log(`  [OK] bulk_${label}.xlsx guardado`);
  return true;
}

async function main() {
  const { TL_USER, TL_PASSWORD, TL_DOMAIN } = process.env;
  if (!TL_USER || !TL_PASSWORD || !TL_DOMAIN)
    throw new Error('Faltan variables: TL_USER, TL_PASSWORD, TL_DOMAIN');

  console.log(`=== Bulk Download (${WINDOWS} ventanas de 7 días) ===`);

  const browser = await puppeteer.launch({
    headless: true,
    args: ['--no-sandbox','--disable-setuid-sandbox','--disable-dev-shm-usage'],
  });

  try {
    const page = await browser.newPage();
    page.setDefaultTimeout(60_000);

    // ── Login (una sola vez) ──────────────────────────────────────────────────
    const loginUrl = `https://${TL_DOMAIN}.trackgts.com/admin/login.html`;
    console.log(`[1] Login en ${loginUrl}`);
    await page.goto(loginUrl, { waitUntil:'networkidle0', timeout:60_000 });
    await page.waitForSelector('#username', { timeout:30_000 });
    await page.evaluate(() => localStorage.setItem('sltLanguage','0'));
    await page.reload({ waitUntil:'networkidle0' });
    await page.waitForSelector('#username', { timeout:30_000 });

    await page.evaluate((user, password, domain) => {
      const K='d5fg4df5sg4ds5fg', S={a:'1',b:'2',c:'3',d:'4',e:'5',f:'6',g:'7',h:'8',i:'9'};
      const k=CryptoJS.enc.Utf8.parse(K), iv=CryptoJS.enc.Utf8.parse(K), a=[];
      for (const c of password) {
        a.push(CryptoJS.AES.encrypt(CryptoJS.enc.Utf8.parse(S[c]||c),k,
          {iv,mode:CryptoJS.mode.CBC,padding:CryptoJS.pad.Pkcs7}).toString());
      }
      ARRAYPSWD=a;
      document.getElementById('username').value=user;
      document.getElementById('domain').value=domain;
      document.getElementById('password').value='********';
      LOGININPROCESS=false;
      onLoginOn();
    }, TL_USER, TL_PASSWORD, TL_DOMAIN);

    console.log('[2] Esperando sesión (15s)...');
    await new Promise(r => setTimeout(r, 15_000));
    console.log(`[2] URL: ${page.url()}`);

    // ── Descargar ventanas ────────────────────────────────────────────────────
    const now = new Date();
    let ok = 0;

    for (let i = 0; i < WINDOWS; i++) {
      const winEnd   = new Date(now); winEnd.setDate(winEnd.getDate() - i * 7);
      const winStart = new Date(winEnd); winStart.setDate(winStart.getDate() - 7);
      const startDate = `${fmt(winStart)} 04:00:00`;
      const endDate   = `${fmt(winEnd)} 03:59:59`;

      const success = await fetchWindow(page, startDate, endDate, i);
      if (success) ok++;

      // Pausa breve entre llamadas para no saturar la API
      if (i < WINDOWS - 1)
        await new Promise(r => setTimeout(r, 2_000));
    }

    console.log(`\n=== Completado: ${ok}/${WINDOWS} ventanas descargadas ===`);

  } finally {
    await browser.close();
  }
}

main().catch(err => {
  console.error('ERROR FATAL:', err.message);
  process.exit(1);
});
