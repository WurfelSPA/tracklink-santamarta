/**
 * download.js — Descarga el reporte semanal "INFORME EXCESOS DE VELOCIDAD"
 * de TrackGTS y lo guarda en la raíz del repo como .xlsx.
 *
 * Variables de entorno requeridas (configurar en GitHub Secrets):
 *   TL_USER     — usuario (ej: amelendez)
 *   TL_PASSWORD — contraseña
 *   TL_DOMAIN   — dominio de TrackGTS (ej: tlchile)
 */
const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path');
const AdmZip = require('adm-zip');

async function main() {
  const { TL_USER, TL_PASSWORD, TL_DOMAIN } = process.env;
  if (!TL_USER || !TL_PASSWORD || !TL_DOMAIN) {
    throw new Error('Faltan env vars: TL_USER, TL_PASSWORD, TL_DOMAIN');
  }

  console.log('▶ Lanzando navegador headless...');
  const browser = await puppeteer.launch({
    headless: 'new',
    args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage'],
  });

  try {
    const page = await browser.newPage();

    console.log('▶ Cargando página de login...');
    await page.goto('https://tlchile.trackgts.com/admin/login.html', {
      waitUntil: 'networkidle0',
      timeout: 60000,
    });
    await page.waitForSelector('#username', { timeout: 30000 });

    console.log('▶ Forzando idioma ES y recargando...');
    await page.evaluate(() => localStorage.setItem('sltLanguage', '0'));
    await page.reload({ waitUntil: 'networkidle0' });
    await page.waitForSelector('#username', { timeout: 30000 });

    console.log('▶ Enviando credenciales...');
    await page.evaluate((user, pass, domain) => {
      const K = 'd5fg4df5sg4ds5fg';
      const S = { a:'1', b:'2', c:'3', d:'4', e:'5', f:'6', g:'7', h:'8', i:'9' };
      const k = CryptoJS.enc.Utf8.parse(K);
      const iv = CryptoJS.enc.Utf8.parse(K);
      const a = [];
      for (const c of pass) {
        a.push(
          CryptoJS.AES.encrypt(
            CryptoJS.enc.Utf8.parse(S[c] || c),
            k,
            { iv, mode: CryptoJS.mode.CBC, padding: CryptoJS.pad.Pkcs7 }
          ).toString()
        );
      }
      ARRAYPSWD = a;
      document.getElementById('username').value = user;
      document.getElementById('domain').value = domain;
      document.getElementById('password').value = pass;
      LOGININPROCESS = false;
      onLoginOn();
    }, TL_USER, TL_PASSWORD, TL_DOMAIN);

    console.log('▶ Esperando que se complete el login (15s)...');
    await new Promise(r => setTimeout(r, 15000));

    console.log('▶ Llamando API de reportes /25/...');
    const result = await page.evaluate(async () => {
      function uuid() {
        return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, c => {
          const r = (Math.random() * 16) | 0;
          const v = c === 'x' ? r : (r & 0x3) | 0x8;
          return v.toString(16);
        });
      }
      const tempId = uuid() + '__t_' + String(Date.now()).slice(-4);
      const e = new Date();
      const s = new Date(e);
      s.setDate(s.getDate() - 7);
      const fmt = d => {
        const p = n => String(n).padStart(2, '0');
        return `${d.getFullYear()}/${p(d.getMonth() + 1)}/${p(d.getDate())}`;
      };
      // El body REPLICA el que envía el botón "Descargar" del reporte guardado
      const body =
        '{"startDate":"' + fmt(s) + ' 04:00:00",' +
        '"endDate":"' + fmt(e) + ' 03:59:59",' +
        '"unitIds":"4349,6399,4436",' +
        '"reportName":"INFORME EXCESOS DE VELOCIDAD",' +
        '"parameters":"undefined",' +
        '"userTimeZone":-4,"userfuelMeasure":0,"userMeasureDistance":0,' +
        '"speed":NaN,"language":0}';
      const url = `https://www.trackgts.com:82/api/reportTravel/GetSpeedingReportByUnitsPagesZip/25/${tempId}`;
      const res = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json;charset=UTF-8' },
        body,
      });
      const text = await res.text();
      let parsed = null;
      try { parsed = JSON.parse(text); } catch (e) {}
      return {
        status: res.status,
        contentType: res.headers.get('content-type') || '',
        rawSnippet: text.substring(0, 500),
        parsed,
      };
    });

    console.log(`▶ API status: ${result.status}, content-type: ${result.contentType}`);

    if (!result.parsed || !result.parsed.FileContents) {
      throw new Error(
        `La API no devolvió FileContents. Status: ${result.status}. ` +
        `Respuesta cruda: ${result.rawSnippet}`
      );
    }

    console.log('▶ Decodificando ZIP base64...');
    const zipBuffer = Buffer.from(result.parsed.FileContents, 'base64');
    const zip = new AdmZip(zipBuffer);
    const entries = zip.getEntries();
    console.log(`▶ ZIP contiene ${entries.length} archivo(s):`);
    entries.forEach(e => console.log(`   - ${e.entryName} (${e.header.size} bytes)`));

    const xlsxEntry = entries.find(e => e.entryName.toLowerCase().endsWith('.xlsx'));
    if (!xlsxEntry) {
      throw new Error('No se encontró archivo .xlsx dentro del ZIP');
    }

    const outputPath = path.join(process.cwd(), 'INFORME EXCESOS DE VELOCIDAD.xlsx');
    fs.writeFileSync(outputPath, xlsxEntry.getData());
    console.log(`✓ Guardado: ${outputPath} (${xlsxEntry.getData().length} bytes)`);
  } finally {
    await browser.close();
  }
}

main().catch(err => {
  console.error('✗ Error:', err.message);
  console.error(err.stack);
  process.exit(1);
});
