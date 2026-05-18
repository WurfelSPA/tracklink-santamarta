#!/usr/bin/env node
/**
 * TrackGTS Report Downloader v3
 * - Usa page.type() para simular escritura real (dispara eventos del formulario)
 * - Verifica login exitoso antes de llamar a la API
 * - Fetch desde el contexto del browser (cookies de sesion automaticas)
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

async function typeInField(page, selector, value) {
  try {
    await page.click(selector, { clickCount: 3 }); // triple click = select all
    await page.type(selector, value, { delay: 30 });
    return true;
  } catch {
    return false;
  }
}

async function main() {
  const { TL_USER, TL_PASSWORD, TL_DOMAIN } = process.env;
  if (!TL_USER || !TL_PASSWORD || !TL_DOMAIN) {
    throw new Error('Faltan variables: TL_USER, TL_PASSWORD, TL_DOMAIN');
  }

  console.log('=== TrackGTS Report Downloader v3 ===');

  const browser = await puppeteer.launch({
    headless: true,
    args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage'],
  });

  try {
    const page = await browser.newPage();
    page.setDefaultTimeout(60_000);

    // ── 1. Abrir login ────────────────────────────────────────────────────
    const loginUrl = `https://${TL_DOMAIN}.trackgts.com/admin/login.html`;
    console.log(`[1] Abriendo: ${loginUrl}`);
    await page.goto(loginUrl, { waitUntil: 'networkidle2', timeout: 30_000 });

    // Diagnostico: que inputs existen en la pagina?
    const formInfo = await page.evaluate(() => ({
      inputs: Array.from(document.querySelectorAll('input'))
        .map(el => ({ id: el.id, name: el.name, type: el.type, placeholder: el.placeholder })),
      hasOnLoginOn: typeof onLoginOn !== 'undefined',
      buttons: Array.from(document.querySelectorAll('button, input[type=submit]'))
        .map(el => ({ id: el.id, text: el.innerText || el.value, onclick: el.getAttribute('onclick') })),
    }));
    console.log('[1] Formulario:', JSON.stringify(formInfo));

    // ── 2. Llenar campos con page.type() (simula teclado real) ────────────
    console.log('[2] Llenando usuario...');
    const filledUser = await typeInField(page, '#user', TL_USER)
      || await typeInField(page, 'input[name="user"]', TL_USER)
      || await typeInField(page, 'input[type="text"]:first-of-type', TL_USER);
    console.log(`    usuario: ${filledUser ? 'OK' : 'FALLO'}`);

    console.log('[2] Llenando password...');
    const filledPass = await typeInField(page, '#password', TL_PASSWORD)
      || await typeInField(page, 'input[name="password"]', TL_PASSWORD)
      || await typeInField(page, 'input[type="password"]', TL_PASSWORD);
    console.log(`    password: ${filledPass ? 'OK' : 'FALLO'}`);

    console.log('[2] Llenando dominio...');
    const filledDomain = await typeInField(page, '#domain', TL_DOMAIN)
      || await typeInField(page, 'input[name="domain"]', TL_DOMAIN);
    console.log(`    dominio: ${filledDomain ? 'OK' : 'no existe campo (puede ser normal)'}`);

    // ── 3. Disparar login ─────────────────────────────────────────────────
    console.log('[3] Disparando login...');
    const trigger = await page.evaluate(() => {
      if (typeof onLoginOn === 'function') { onLoginOn(); return 'onLoginOn()'; }
      if (typeof loginOn   === 'function') { loginOn();   return 'loginOn()'; }
      const btn = document.querySelector('button[onclick], button[type=submit], input[type=submit]');
      if (btn) { btn.click(); return `click en: ${btn.id || btn.className || btn.type}`; }
      return 'ninguno';
    });
    console.log(`[3] Trigger: ${trigger}`);

    // ── 4. Esperar redireccion post-login ─────────────────────────────────
    console.log('[4] Esperando redireccion del login...');
    try {
      await page.waitForNavigation({ waitUntil: 'networkidle2', timeout: 30_000 });
      console.log('[4] Navegacion detectada!');
    } catch {
      console.log('[4] Sin navegacion en 30s, continuando de todos modos...');
    }

    const finalUrl = page.url();
    console.log(`[4] URL final: ${finalUrl}`);

    if (finalUrl.toLowerCase().includes('login')) {
      // Si sigue en login, intentar una vez mas con un boton alternativo
      console.log('[4] Todavia en login, intentando click en boton...');
      await page.evaluate(() => {
        const btn = document.querySelector('button, input[type=submit]');
        if (btn) btn.click();
      });
      await new Promise(r => setTimeout(r, 15_000));
      const url2 = page.url();
      console.log(`[4] URL tras segundo intento: ${url2}`);
      if (url2.toLowerCase().includes('login')) {
        throw new Error(`Login fallido - sigue en pagina de login. Verifica TL_USER, TL_PASSWORD, TL_DOMAIN en los Secrets de GitHub.`);
      }
    }

    console.log('[4] Login exitoso!');

    // ── 5. Llamar API desde el browser (cookies automaticas) ──────────────
    const requestId = generateRequestId();
    const { startDate, endDate } = getWeekDateRange();
    const apiUrl = `https://www.trackgts.com:82/api/reportTravel/GetSpeedingReportByUnitsPagesZip/25/${requestId}`;
    const body =
      `{"startDate":"${startDate}","endDate":"${endDate}",` +
      `"unitIds":"4349,6399,4436","reportName":"INFORME EXCESOS DE VELOCIDAD",` +
      `"parameters":"undefined","userTimeZone":-4,"userfuelMeasure":0,` +
      `"userMeasureDistance":0,"speed":NaN,"language":0}`;

    console.log(`[5] POST -> ${apiUrl}`);
    console.log(`[5] Rango: ${startDate} -> ${endDate}`);

    const result = await page.evaluate(async (url, reqBody) => {
      try {
        const res  = await fetch(url, {
          method : 'POST',
          headers: { 'Content-Type': 'application/json;charset=UTF-8' },
          body   : reqBody,
        });
        const text = await res.text();
        let json;
        try { json = JSON.parse(text); }
        catch { return { error: `No-JSON (HTTP ${res.status}): ${text.slice(0, 400)}` }; }
        if (!json.FileContents) return { error: `Sin FileContents: ${JSON.stringify(json).slice(0, 500)}` };
        return { fileContents: json.FileContents };
      } catch (e) {
        return { error: e.message };
      }
    }, apiUrl, body);

    if (result.error) throw new Error(result.error);
    console.log('[5] Reporte descargado OK');

    // ── 6. Extraer .xlsx del zip ───────────────────────────────────────────
    console.log('[6] Extrayendo .xlsx...');
    const zip  = new AdmZip(Buffer.from(result.fileContents, 'base64'));
    const xlsx = zip.getEntries().find(e => e.entryName.toLowerCase().endsWith('.xlsx'));
    if (!xlsx) throw new Error(`Sin .xlsx en zip. Entradas: ${zip.getEntries().map(e=>e.entryName).join(', ')}`);

    const dest = path.join(process.cwd(), 'INFORME EXCESOS DE VELOCIDAD.xlsx');
    fs.writeFileSync(dest, xlsx.getData());
    console.log(`[6] Guardado: ${dest}`);
    console.log('=== COMPLETADO ===');

  } finally {
    await browser.close();
  }
}

main().catch(err => {
  console.error('ERROR FATAL:', err.message);
  process.exit(1);
});
