#!/usr/bin/env node
/**
 * TrackGTS Report Downloader
 * Login con Puppeteer → llama API → extrae .xlsx del zip → guarda en raíz del repo
 *
 * Variables de entorno requeridas:
 *   TL_USER     – usuario TrackGTS
 *   TL_PASSWORD – contraseña (plain text; la página la encripta internamente)
 *   TL_DOMAIN   – dominio TrackGTS (ej: "tlchile")
 */

'use strict';

const puppeteer = require('puppeteer');
const AdmZip    = require('adm-zip');
const https     = require('https');
const fs        = require('fs');
const path      = require('path');

// ── Helpers ──────────────────────────────────────────────────────────────────

function generateRequestId() {
  const uuid = 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, c => {
    const r = Math.random() * 16 | 0;
    return (c === 'x' ? r : (r & 0x3 | 0x8)).toString(16);
  });
  const suffix = String(Date.now()).slice(-4);
  return `${uuid}__t_${suffix}`;
}

function getWeekDateRange() {
  const now = new Date();
  const pad = n => String(n).padStart(2, '0');
  const fmt = d =>
    `${d.getFullYear()}/${pad(d.getMonth()+1)}/${pad(d.getDate())} ` +
    `${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;

  // Inicio: hace 8 días a las 04:00 (cubre 7 días completos)
  const start = new Date(now);
  start.setDate(start.getDate() - 8);
  start.setHours(4, 0, 0, 0);

  // Fin: ayer a las 03:59:59
  const end = new Date(now);
  end.setDate(end.getDate() - 1);
  end.setHours(3, 59, 59, 0);

  return { startDate: fmt(start), endDate: fmt(end) };
}

function httpsPost(url, body, cookieStr) {
  return new Promise((resolve, reject) => {
    const u   = new URL(url);
    const buf = Buffer.from(body, 'utf8');

    const req = https.request({
      hostname           : u.hostname,
      port               : u.port || 443,
      path               : u.pathname + u.search,
      method             : 'POST',
      rejectUnauthorized : false,   // API en puerto 82 puede usar cert autofirmado
      headers: {
        'Content-Type'   : 'application/json;charset=UTF-8',
        'Content-Length' : buf.length,
        'Cookie'         : cookieStr,
      },
    }, res => {
      const chunks = [];
      res.on('data', c => chunks.push(c));
      res.on('end', () => {
        const text = Buffer.concat(chunks).toString('utf8');
        try   { resolve(JSON.parse(text)); }
        catch { reject(new Error(`Respuesta no-JSON (HTTP ${res.statusCode}): ${text.slice(0, 400)}`)); }
      });
    });

    req.on('error', reject);
    req.write(buf);
    req.end();
  });
}

// ── Main ─────────────────────────────────────────────────────────────────────

async function main() {
  const { TL_USER, TL_PASSWORD, TL_DOMAIN } = process.env;
  if (!TL_USER || !TL_PASSWORD || !TL_DOMAIN) {
    throw new Error('Faltan variables de entorno: TL_USER, TL_PASSWORD, TL_DOMAIN');
  }

  console.log('🚀 Iniciando descarga de reporte TrackGTS...');

  const browser = await puppeteer.launch({
    headless: true,
    args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage'],
  });

  try {
    const page = await browser.newPage();
    page.setDefaultTimeout(30_000);

    // ── 1. Login ──────────────────────────────────────────────────────────
    console.log('📍 Abriendo página de login...');
    await page.goto(`https://${TL_DOMAIN}.trackgts.com/admin/login.html`, {
      waitUntil: 'networkidle2',
    });

    console.log('📝 Llenando credenciales...');
    await page.evaluate((u, p, d) => {
      const set = (sel, val) => { const el = document.querySelector(sel); if (el) el.value = val; };
      set('#user',     u);
      set('#password', p);
      set('#domain',   d);
    }, TL_USER, TL_PASSWORD, TL_DOMAIN);

    console.log('🔑 Llamando onLoginOn()...');
    await page.evaluate(() => {
      if (typeof onLoginOn === 'function') return onLoginOn();
      if (typeof loginOn   === 'function') return loginOn();
      document.querySelector('button[type=submit], input[type=submit]')?.click();
    });

    console.log('⏳ Esperando 20 s para establecer sesión...');
    await new Promise(r => setTimeout(r, 20_000));

    // ── 2. Recolectar cookies ─────────────────────────────────────────────
    const cdp = await page.target().createCDPSession();
    const { cookies } = await cdp.send('Network.getAllCookies');
    const cookieStr = cookies.map(c => `${c.name}=${c.value}`).join('; ');
    console.log(`🍪 ${cookies.length} cookies recolectadas`);

    // ── 3. Llamar API del reporte ─────────────────────────────────────────
    const requestId = generateRequestId();
    const { startDate, endDate } = getWeekDateRange();

    // Nota: "speed":NaN no es JSON estándar pero la API lo requiere así
    const body =
      `{"startDate":"${startDate}","endDate":"${endDate}",` +
      `"unitIds":"4349,6399,4436","reportName":"INFORME EXCESOS DE VELOCIDAD",` +
      `"parameters":"undefined","userTimeZone":-4,"userfuelMeasure":0,` +
      `"userMeasureDistance":0,"speed":NaN,"language":0}`;

    const apiUrl = `https://www.trackgts.com:82/api/reportTravel/GetSpeedingReportByUnitsPagesZip/25/${requestId}`;
    console.log(`📡 POST → ${apiUrl}`);
    console.log(`📅 Rango: ${startDate}  →  ${endDate}`);

    const json = await httpsPost(apiUrl, body, cookieStr);

    if (!json?.FileContents) {
      throw new Error(`Sin FileContents en la respuesta: ${JSON.stringify(json).slice(0, 500)}`);
    }

    // ── 4. Extraer .xlsx del zip en base64 ────────────────────────────────
    console.log('📦 Extrayendo .xlsx del zip...');
    const zipBuf  = Buffer.from(json.FileContents, 'base64');
    const zip     = new AdmZip(zipBuf);
    const entries = zip.getEntries();
    const xlsx    = entries.find(e => e.entryName.toLowerCase().endsWith('.xlsx'));

    if (!xlsx) {
      throw new Error(`No hay .xlsx en el zip. Encontrado: ${entries.map(e => e.entryName).join(', ')}`);
    }

    // Guardar en la raíz del repo (mismo directorio que este script en el workflow)
    const dest = path.join(process.cwd(), 'INFORME EXCESOS DE VELOCIDAD.xlsx');
    fs.writeFileSync(dest, xlsx.getData());
    console.log(`✅ Guardado → ${dest}`);

  } finally {
    await browser.close();
  }
}

main().catch(err => {
  console.error('❌', err.message);
  process.exit(1);
});
