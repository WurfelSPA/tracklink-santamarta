#!/usr/bin/env node
/**
 * TrackGTS Report Downloader v4
 * - Detecta iframe y trabaja dentro de él
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

async function findLoginFrame(page) {
  // Espera hasta 15s a que aparezca algun input en cualquier frame
  for (let i = 0; i < 30; i++) {
    const frames = page.frames();
    for (const frame of frames) {
      try {
        const info = await frame.evaluate(() => ({
          url    : window.location.href,
          inputs : Array.from(document.querySelectorAll('input'))
                     .map(el => ({ id: el.id, name: el.name, type: el.type })),
          hasOnLoginOn: typeof onLoginOn === 'function',
        }));
        if (info.inputs.length > 0) {
          console.log(`[1] Formulario en frame: ${info.url}`);
          console.log(`[1] Inputs: ${JSON.stringify(info.inputs)}`);
          console.log(`[1] hasOnLoginOn: ${info.hasOnLoginOn}`);
          return { frame, info };
        }
      } catch { /* frame puede no estar listo aun */ }
    }
    await new Promise(r => setTimeout(r, 500));
  }
  return null;
}

async function typeInFrame(frame, selector, value) {
  try {
    await frame.click(selector, { clickCount: 3 });
    await frame.type(selector, value, { delay: 30 });
    return true;
  } catch { return false; }
}

async function main() {
  const { TL_USER, TL_PASSWORD, TL_DOMAIN } = process.env;
  if (!TL_USER || !TL_PASSWORD || !TL_DOMAIN) {
    throw new Error('Faltan variables: TL_USER, TL_PASSWORD, TL_DOMAIN');
  }

  console.log('=== TrackGTS Report Downloader v4 ===');

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

    // Log todos los frames
    const allFrames = page.frames();
    console.log(`[1] Frames detectados: ${allFrames.length}`);
    allFrames.forEach((f, i) => {
      try { console.log(`    Frame ${i}: ${f.url()}`); } catch {}
    });

    // Buscar el frame con el formulario de login
    const found = await findLoginFrame(page);
    if (!found) throw new Error('No se encontro formulario de login en ningun frame despues de 15s');

    const { frame, info } = found;

    // ── 2. Llenar campos ──────────────────────────────────────────────────
    console.log('[2] Llenando campos...');

    const userSel = info.inputs.find(f => f.id === 'user' || f.name === 'user')
      ? '#user' : 'input[type="text"]:first-of-type';
    const passSel = info.inputs.find(f => f.id === 'password' || f.type === 'password')
      ? (info.inputs.find(f => f.id === 'password') ? '#password' : 'input[type="password"]')
      : 'input[type="password"]';

    const okUser = await typeInFrame(frame, '#user', TL_USER)
      || await typeInFrame(frame, 'input[name="user"]', TL_USER)
      || await typeInFrame(frame, 'input[type="text"]:first-of-type', TL_USER);
    console.log(`    usuario (${userSel}): ${okUser ? 'OK' : 'FALLO'}`);

    const okPass = await typeInFrame(frame, '#password', TL_PASSWORD)
      || await typeInFrame(frame, 'input[name="password"]', TL_PASSWORD)
      || await typeInFrame(frame, 'input[type="password"]', TL_PASSWORD);
    console.log(`    password: ${okPass ? 'OK' : 'FALLO'}`);

    const hasDomain = info.inputs.some(f => f.id === 'domain' || f.name === 'domain');
    if (hasDomain) {
      const okDom = await typeInFrame(frame, '#domain', TL_DOMAIN)
        || await typeInFrame(frame, 'input[name="domain"]', TL_DOMAIN);
      console.log(`    dominio: ${okDom ? 'OK' : 'FALLO'}`);
    } else {
      console.log('    dominio: no hay campo (OK, viene del subdominio)');
    }

    // ── 3. Disparar login ─────────────────────────────────────────────────
    console.log('[3] Disparando login...');
    const trigger = await frame.evaluate(() => {
      if (typeof onLoginOn === 'function') { onLoginOn(); return 'onLoginOn()'; }
      if (typeof loginOn   === 'function') { loginOn();   return 'loginOn()'; }
      const btn = document.querySelector('button[onclick], button[type=submit], input[type=submit], button');
      if (btn) { btn.click(); return `click: ${btn.id || btn.innerText || btn.type}`; }
      return 'ninguno';
    });
    console.log(`[3] Trigger: ${trigger}`);

    // ── 4. Esperar sesion ─────────────────────────────────────────────────
    console.log('[4] Esperando 25s para que la sesion se establezca...');
    await new Promise(r => setTimeout(r, 25_000));
    console.log(`[4] URL main: ${page.url()}`);

    // ── 5. Fetch de la API desde el browser ───────────────────────────────
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

    // Intentar fetch desde page y desde frame
    let result = null;
    for (const ctx of [page, frame]) {
      try {
        result = await ctx.evaluate(async (url, reqBody) => {
          const res  = await fetch(url, {
            method : 'POST',
            headers: { 'Content-Type': 'application/json;charset=UTF-8' },
            body   : reqBody,
          });
          const text = await res.text();
          let json;
          try { json = JSON.parse(text); } catch {
            return { error: `No-JSON (${res.status}): ${text.slice(0, 400)}` };
          }
          if (!json.FileContents) return { error: `Sin FileContents: ${JSON.stringify(json).slice(0, 500)}` };
          return { fileContents: json.FileContents };
        }, apiUrl, body);
        if (result && result.fileContents) break;
        console.log(`[5] Contexto ${ctx === page ? 'page' : 'frame'}: ${result?.error}`);
      } catch (e) {
        console.log(`[5] Error en contexto: ${e.message}`);
      }
    }

    if (!result || result.error) throw new Error(result?.error || 'Sin respuesta de la API');
    console.log('[5] Reporte descargado OK');

    // ── 6. Extraer xlsx ───────────────────────────────────────────────────
    console.log('[6] Extrayendo .xlsx...');
    const zip  = new AdmZip(Buffer.from(result.fileContents, 'base64'));
    const xlsx = zip.getEntries().find(e => e.entryName.toLowerCase().endsWith('.xlsx'));
    if (!xlsx) throw new Error(`Sin .xlsx. Entradas: ${zip.getEntries().map(e=>e.entryName).join(', ')}`);

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
