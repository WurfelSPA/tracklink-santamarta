#!/usr/bin/env node
/**
 * generate-pdf.js
 * Genera el Reporte Semanal de Excesos de Velocidad en PDF.
 *
 * Se ejecuta dentro del GitHub Action DESPUÉS de merge_history.py,
 * por lo que el archivo "INFORME EXCESOS DE VELOCIDAD.xlsx" ya está
 * actualizado con los datos de la semana.
 *
 * Variables de entorno esperadas (seteadas por el step anterior):
 *   REPORT_START  — fecha inicio "YYYY-MM-DD"
 *   REPORT_END    — fecha fin   "YYYY-MM-DD"
 *
 * Salida: reporte-semanal.pdf (en el directorio de trabajo)
 */
'use strict';

const puppeteer = require('puppeteer');
const XLSX      = require('xlsx');
const fs        = require('fs');
const path      = require('path');

const EXCEL_FILE = path.join(process.cwd(), 'INFORME EXCESOS DE VELOCIDAD.xlsx');
const OUTPUT_PDF = path.join(process.cwd(), 'reporte-semanal.pdf');
const DIAS_ES    = ['Dom', 'Lun', 'Mar', 'Mié', 'Jue', 'Vie', 'Sáb'];

// ─────────────────────────────────────────────────────────────────────────────
async function main() {
  const startDate = process.env.REPORT_START;
  const endDate   = process.env.REPORT_END;

  if (!startDate || !endDate) {
    throw new Error('Faltan variables REPORT_START y REPORT_END.');
  }

  console.log(`[pdf] Generando reporte: ${startDate} → ${endDate}`);

  if (!fs.existsSync(EXCEL_FILE)) {
    throw new Error(`No se encontró: ${EXCEL_FILE}`);
  }

  // 1. Parsear y filtrar
  const { rows, columns } = parseAndFilter(fs.readFileSync(EXCEL_FILE), startDate, endDate);
  console.log(`[pdf] Filas en el período: ${rows.length}`);

  if (rows.length === 0) {
    console.warn('[pdf] Sin datos para el período — se genera PDF con aviso.');
  }

  // 2. Calcular estadísticas
  const stats = computeStats(rows, columns, startDate, endDate);

  // 3. Generar HTML
  const html = generateHTML(stats);

  // 4. Puppeteer → PDF
  console.log('[pdf] Iniciando Puppeteer...');
  const browser = await puppeteer.launch({
    headless: 'new',
    args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage'],
  });

  try {
    const page = await browser.newPage();
    await page.setViewport({ width: 1122, height: 794 }); // A4 landscape en 96dpi
    await page.setContent(html, { waitUntil: 'networkidle0', timeout: 40000 });

    // Esperar a que Chart.js termine de renderizar
    await page
      .waitForFunction('window.__chartsReady === true', { timeout: 15000 })
      .catch(() => console.warn('[pdf] Chart timeout — continuando'));

    const pdf = await page.pdf({
      width:           '297mm',
      height:          '210mm',
      printBackground: true,
      margin:          { top: 0, right: 0, bottom: 0, left: 0 },
    });

    fs.writeFileSync(OUTPUT_PDF, pdf);
    console.log(`[pdf] ✅ PDF guardado: ${OUTPUT_PDF} (${pdf.length} bytes)`);
  } finally {
    await browser.close();
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// PARSEO
// ─────────────────────────────────────────────────────────────────────────────
function parseAndFilter(buffer, startDate, endDate) {
  const wb = XLSX.read(buffer, { type: 'buffer', cellDates: true, dateNF: 'yyyy/mm/dd hh:mm:ss' });

  const sheetName =
    wb.SheetNames.find((n) => n.toLowerCase().includes('detalle')) || wb.SheetNames[0];
  const ws  = wb.Sheets[sheetName];
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false });

  let headerIdx = -1, headers = [];
  for (let i = 0; i < Math.min(15, aoa.length); i++) {
    if (aoa[i] && aoa[i].some((v) => String(v || '').trim() === 'Alias')) {
      headerIdx = i;
      headers   = aoa[i].map((v) => String(v || '').trim());
      break;
    }
  }
  if (headerIdx === -1) throw new Error('No se encontró la columna "Alias".');

  const aliasIdx  = headers.indexOf('Alias');
  const fechaIdx  = headers.findIndex(
    (h) => h.toLowerCase().includes('fecha') && h.toLowerCase().includes('inicio')
  );
  const velIdx    = headers.findIndex(
    (h) =>
      (h.toLowerCase().includes('velocidad') || h.toLowerCase().startsWith('vel')) &&
      !h.toLowerCase().includes('permit') &&
      !h.toLowerCase().includes('limit')
  );
  const patenteIdx = headers.findIndex(
    (h) =>
      h.toLowerCase().includes('patente') ||
      h.toLowerCase().includes('vehículo') ||
      h.toLowerCase().includes('vehiculo') ||
      h.toLowerCase() === 'unidad' ||
      h.toLowerCase().includes('descripcion') ||
      h.toLowerCase().includes('descripción')
  );

  const columns = {
    alias:    headers[aliasIdx]   || 'Alias',
    fecha:    headers[fechaIdx]   || 'Fecha de Inicio',
    velocidad: headers[velIdx]   || null,
    patente:  patenteIdx >= 0 ? headers[patenteIdx] : null,
  };

  const [sy, sm, sd] = startDate.split('-').map(Number);
  const [ey, em, ed] = endDate.split('-').map(Number);
  const startTs = new Date(sy, sm - 1, sd, 0, 0, 0).getTime();
  const endTs   = new Date(ey, em - 1, ed, 23, 59, 59).getTime();

  const rows = [];
  for (let i = headerIdx + 1; i < aoa.length; i++) {
    const row = aoa[i];
    if (!row || !row[aliasIdx]) continue;

    const obj = {};
    headers.forEach((h, idx) => { obj[h] = row[idx]; });

    const fecha = parseFecha(obj[columns.fecha]);
    if (!fecha) continue;
    const ts = fecha.getTime();
    if (ts < startTs || ts > endTs) continue;

    obj['__fecha'] = fecha;
    obj['__vel']   = parseFloat(String(obj[columns.velocidad] || '0').replace(',', '.')) || 0;
    rows.push(obj);
  }

  return { rows, columns };
}

function parseFecha(val) {
  if (!val)              return null;
  if (val instanceof Date) return isNaN(val) ? null : val;
  const s = String(val).trim();
  const p1 = s.match(/^(\d{4})[\/\-](\d{2})[\/\-](\d{2})\s+(\d{2}):(\d{2}):(\d{2})/);
  if (p1) return new Date(+p1[1], +p1[2]-1, +p1[3], +p1[4], +p1[5], +p1[6]);
  const p2 = s.match(/^(\d{2})[\/\-](\d{2})[\/\-](\d{4})\s+(\d{2}):(\d{2}):(\d{2})/);
  if (p2) return new Date(+p2[3], +p2[2]-1, +p2[1], +p2[4], +p2[5], +p2[6]);
  const d = new Date(val);
  return isNaN(d) ? null : d;
}

// ─────────────────────────────────────────────────────────────────────────────
// ESTADÍSTICAS
// ─────────────────────────────────────────────────────────────────────────────
function computeStats(rows, columns, startDate, endDate) {
  const conductores = {}, byDayMap = {}, byHourMap = {};

  rows.forEach((row) => {
    const alias   = String(row[columns.alias] || '').trim();
    const d       = row['__fecha'];
    const speed   = row['__vel'];
    const vehicle = columns.patente ? String(row[columns.patente] || '').trim() : '';

    if (!conductores[alias]) {
      conductores[alias] = { count: 0, maxSpeed: 0, maxSpeedTime: null, vehicle };
    }
    conductores[alias].count++;
    if (speed > conductores[alias].maxSpeed) {
      conductores[alias].maxSpeed    = speed;
      conductores[alias].maxSpeedTime = d;
    }

    const dayKey = `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`;
    byDayMap[dayKey]   = (byDayMap[dayKey]   || 0) + 1;
    byHourMap[d.getHours()] = (byHourMap[d.getHours()] || 0) + 1;
  });

  const totalIncidencias = rows.length;

  const conductoresArr = Object.entries(conductores)
    .sort((a, b) => b[1].count - a[1].count)
    .map(([name, data]) => ({
      name,
      vehicle:      data.vehicle,
      count:        data.count,
      pct:          ((data.count / (totalIncidencias || 1)) * 100).toFixed(1),
      maxSpeed:     data.maxSpeed.toFixed(1),
      maxSpeedTime: data.maxSpeedTime ? formatDateTime(data.maxSpeedTime) : '—',
    }));

  const [sy, sm, sd] = startDate.split('-').map(Number);
  const weekDays = Array.from({ length: 7 }, (_, i) => {
    const d   = new Date(sy, sm-1, sd+i);
    const key = `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`;
    return { key, label: `${DIAS_ES[d.getDay()]} ${pad(d.getDate())}/${pad(d.getMonth()+1)}`, count: byDayMap[key] || 0 };
  });

  const topHours = Array.from({ length: 24 }, (_, i) => ({
    hour: i, label: `${pad(i)}:00–${pad(i)}:59`, count: byHourMap[i] || 0,
  })).filter(h => h.count > 0).sort((a,b) => b.count-a.count).slice(0, 8);

  const peakDay     = weekDays.reduce((m,d) => d.count>m.count ? d : m, weekDays[0] || {label:'—',count:0,key:''});
  const mananaTotal = [8,9,10,11,12].reduce((s,h) => s+(byHourMap[h]||0), 0);
  const mananaPct   = totalIncidencias > 0 ? Math.round((mananaTotal/totalIncidencias)*100) : 0;
  const sortedDays  = weekDays.slice().sort((a,b) => b.count-a.count);

  const [ey, em, ed] = endDate.split('-').map(Number);
  return {
    startDate, endDate,
    startDisplay: `${pad(sd)}/${pad(sm)}/${sy}`,
    endDisplay:   `${pad(ed)}/${pad(em)}/${ey}`,
    totalIncidencias, conductoresArr, weekDays, topHours,
    peakDay, mananaTotal, mananaPct, sortedDays,
    globalMaxSpeed:    (conductoresArr[0]?.maxSpeed)    || '0.0',
    globalMaxConductor:(conductoresArr[0]?.name)        || '—',
    globalMaxTime:     (conductoresArr[0]?.maxSpeedTime)|| '—',
  };
}

function pad(n)           { return String(n).padStart(2,'0'); }
function formatDateTime(d) {
  return `${DIAS_ES[d.getDay()]} ${pad(d.getDate())}/${pad(d.getMonth()+1)} · ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
}

// ─────────────────────────────────────────────────────────────────────────────
// HTML
// ─────────────────────────────────────────────────────────────────────────────
function generateHTML(s) {
  const dayLabels  = JSON.stringify(s.weekDays.map(d => d.label));
  const dayCounts  = JSON.stringify(s.weekDays.map(d => d.count));
  const dayColors  = JSON.stringify(s.weekDays.map(d => d.key === s.peakDay.key ? '#2d3748' : '#94a3b8'));
  const hourLabels = JSON.stringify(s.topHours.map(h => h.label));
  const hourCounts = JSON.stringify(s.topHours.map(h => h.count));
  const hourColors = JSON.stringify(s.topHours.map((_,i) => i===0?'#2d3748':i<3?'#4a5568':'#94a3b8'));

  const footerText   = `Tracklink Chile Fleet Dashboard By Würfel SpA · Reporte · ${s.totalIncidencias} eventos · generado automáticamente por IA`;
  const footer       = `<div class="tl-footer">${footerText}</div>`;

  const peakHourLabel = s.topHours[0]?.label || '—';
  const peakHourCount = s.topHours[0]?.count || 0;

  const conductorCards = s.conductoresArr.map((c, idx) => `
    <div class="conductor-card ${idx===0?'c-dark':'c-light'}">
      <div class="c-name">${c.name}</div>
      <div class="c-vehicle">${c.vehicle ? `Unidad: <strong>${c.vehicle}</strong>` : `Conductor ${idx+1}`}</div>
      <div class="c-stats">
        <div class="c-stat">
          <div class="c-stat-val">${c.count} excesos</div>
          <div class="c-stat-lbl">INCIDENCIAS</div>
          <div class="c-stat-sub">${c.pct}% del total</div>
        </div>
        <div class="c-stat">
          <div class="c-stat-val">${c.maxSpeed} km/h</div>
          <div class="c-stat-lbl">VEL. MÁXIMA</div>
          <div class="c-stat-sub">${c.maxSpeedTime}</div>
        </div>
      </div>
    </div>`).join('');

  const conducInsight = s.conductoresArr.length >= 2 ? `
    <div class="alert alert-red">
      <span>⊗</span>
      <div>${s.conductoresArr[0].name} lidera con ${s.conductoresArr[0].count} excesos (${s.conductoresArr[0].pct}%).
      ${Math.abs(s.conductoresArr[0].count - s.conductoresArr[1].count) < s.conductoresArr[0].count*0.12
        ? ' Ambos conductores presentan una distribución casi equitativa de infracciones, lo que indica un patrón sistemático en ambas unidades.'
        : ` La diferencia entre conductores es de ${s.conductoresArr[0].count - s.conductoresArr[1].count} incidencias.`}
      </div>
    </div>` : '';

  const conclusionConductores = s.conductoresArr.map(c => c.name).join(' y ');
  const diasCriticos          = s.sortedDays.slice(0,3).map(d => d.label).join(', ');

  return `<!DOCTYPE html>
<html lang="es"><head><meta charset="UTF-8">
<style>
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  @page { size: 297mm 210mm; margin: 0; }
  body { font-family: 'Segoe UI', system-ui, Arial, sans-serif; background: #fff; }
  .page { width:297mm; height:210mm; position:relative; overflow:hidden; page-break-after:always; background:#fff; }
  .page:last-child { page-break-after:avoid; }

  /* Portada */
  .cover { display:flex; }
  .cv-left { width:54%; padding:44px 48px; display:flex; flex-direction:column; justify-content:space-between; }
  .cv-right { width:46%; background:linear-gradient(145deg,#1a3352 0%,#1e4080 40%,#0f2340 100%); position:relative; overflow:hidden; display:flex; align-items:center; justify-content:center; }
  .cv-right::after { content:''; position:absolute; inset:0; background:repeating-linear-gradient(-45deg,transparent,transparent 28px,rgba(255,255,255,.025) 28px,rgba(255,255,255,.025) 56px); }
  .cv-title { font-size:36px; font-weight:800; color:#1a202c; line-height:1.1; margin-bottom:14px; }
  .cv-sub   { font-size:13.5px; font-weight:700; color:#4a5568; margin-bottom:18px; }
  .cv-desc  { font-size:12.5px; color:#718096; line-height:1.65; max-width:370px; }
  .cv-badges { display:flex; gap:8px; margin-top:20px; flex-wrap:wrap; }
  .badge { background:#f7fafc; border:1px solid #e2e8f0; padding:5px 13px; font-size:10.5px; font-weight:700; color:#4a5568; letter-spacing:.06em; }

  /* Interior */
  .pi { padding:36px 48px 28px; height:100%; display:flex; flex-direction:column; }
  .pg-title { font-size:26px; font-weight:800; color:#1a202c; margin-bottom:6px; }
  .pg-sub   { font-size:12px; color:#718096; margin-bottom:24px; }
  .pf { position:absolute; bottom:16px; left:0; right:0; display:flex; justify-content:center; }
  .tl-footer { font-size:9.5px; color:#94a3b8; letter-spacing:.03em; text-align:center; }

  /* KPIs */
  .kpi-grid { display:grid; grid-template-columns:1fr 1fr; gap:14px; flex:1; margin-bottom:14px; }
  .kpi { border:1px solid #e2e8f0; border-radius:8px; padding:22px 26px; display:flex; flex-direction:column; justify-content:center; }
  .kpi-val  { font-size:50px; font-weight:800; color:#2d3748; line-height:1; margin-bottom:5px; }
  .kpi-lbl  { font-size:14px; font-weight:700; color:#4a5568; margin-bottom:5px; }
  .kpi-desc { font-size:11.5px; color:#718096; line-height:1.45; }

  /* Alertas */
  .alert { padding:12px 16px; display:flex; gap:10px; align-items:flex-start; font-size:12px; line-height:1.5; border-radius:6px; }
  .alert span { flex-shrink:0; font-size:14px; margin-top:1px; }
  .alert-yellow { background:#fefce8; border-left:4px solid #eab308; }
  .alert-red    { background:#fff5f5; border-left:4px solid #fc8181; }
  .alert-blue   { background:#eff6ff; border-left:4px solid #93c5fd; }

  /* Conductores */
  .conductores-row { display:flex; gap:20px; flex:1; }
  .conductor-card { flex:1; border-radius:8px; padding:22px 24px; }
  .c-dark  { background:#374151; color:#fff; }
  .c-light { background:#fff; border:1px solid #e2e8f0; color:#1a202c; }
  .c-name { font-size:19px; font-weight:700; margin-bottom:4px; }
  .c-vehicle { font-size:11.5px; opacity:.7; margin-bottom:18px; }
  .c-stats { display:grid; grid-template-columns:1fr 1fr; gap:10px; }
  .c-stat { padding:13px; border-radius:6px; }
  .c-dark  .c-stat { background:rgba(255,255,255,.1); }
  .c-light .c-stat { background:#f8fafc; }
  .c-stat-val { font-size:20px; font-weight:700; margin-bottom:3px; }
  .c-stat-lbl { font-size:9.5px; font-weight:700; text-transform:uppercase; letter-spacing:.06em; opacity:.7; margin-bottom:2px; }
  .c-stat-sub { font-size:10.5px; opacity:.65; line-height:1.3; }

  /* Charts */
  .charts-row { display:grid; grid-template-columns:1fr 1fr; gap:28px; flex:1; }
  .chart-sec h3 { font-size:14px; font-weight:700; color:#2d3748; margin-bottom:10px; }
  .chart-wrap { position:relative; height:135px; }
  .chart-note { font-size:11.5px; color:#718096; line-height:1.5; margin-top:8px; }

  /* Conclusiones */
  .concl-grid { display:grid; grid-template-columns:1fr 1fr; gap:14px; flex:1; margin-bottom:14px; }
  .concl-card { border:1px solid #e2e8f0; border-left:4px solid #2d3748; border-radius:0 8px 8px 0; padding:18px 18px 18px 20px; }
  .concl-card h4 { font-size:13px; font-weight:700; color:#2d3748; margin-bottom:6px; }
  .concl-card p  { font-size:11.5px; color:#718096; line-height:1.6; }
</style>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.4/dist/chart.umd.min.js"></script>
</head><body>

<!-- PORTADA -->
<div class="page cover">
  <div class="cv-left">
    <div style="font-size:11px;font-weight:700;letter-spacing:.06em;color:#1a202c">TRACKLINK<span style="color:#d97706">⋘</span></div>
    <div>
      <h1 class="cv-title">Reporte de Excesos de<br>Velocidad</h1>
      <p class="cv-sub">Flota Santa Marta · Período analizado: ${s.startDisplay} al ${s.endDisplay}</p>
      <p class="cv-desc">Durante la semana analizada se registraron un total de <strong>${s.totalIncidencias} incidencias</strong>
        de exceso de velocidad distribuidas entre ${s.conductoresArr.length} conductor${s.conductoresArr.length!==1?'es':''} identificado${s.conductoresArr.length!==1?'s':''}.
        Este informe detalla las incidencias por conductor, velocidades máximas registradas,
        distribución horaria y concentración diaria.</p>
      <div class="cv-badges">
        <div class="badge">SEMANA ANALIZADA</div>
        <div class="badge">${s.conductoresArr.length} CONDUCTOR${s.conductoresArr.length!==1?'ES':''}</div>
        <div class="badge">${s.totalIncidencias} INCIDENCIAS</div>
      </div>
    </div>
    <div></div>
  </div>
  <div class="cv-right">
    <div style="position:relative;z-index:1">
      <svg width="300" height="170" viewBox="0 0 300 170" fill="none" opacity=".35">
        <rect x="0" y="130" width="300" height="3" fill="#fff" opacity=".4"/>
        <rect x="55" y="65" width="170" height="68" rx="4" fill="#fff" opacity=".55"/>
        <rect x="188" y="50" width="46" height="83" rx="3" fill="#fff" opacity=".75"/>
        <circle cx="88" cy="136" r="14" fill="#1e3a5f"/><circle cx="88" cy="136" r="7" fill="#fff" opacity=".5"/>
        <circle cx="200" cy="136" r="14" fill="#1e3a5f"/><circle cx="200" cy="136" r="7" fill="#fff" opacity=".5"/>
        <rect x="192" y="55" width="38" height="32" rx="2" fill="#93c5fd" opacity=".5"/>
        <rect x="228" y="70" width="7" height="11" rx="1" fill="#fbbf24" opacity=".8"/>
        <rect x="0" y="82" width="52" height="50" rx="3" fill="#fff" opacity=".25"/>
        <circle cx="13" cy="134" r="10" fill="#1e3a5f" opacity=".5"/><circle cx="50" cy="134" r="10" fill="#1e3a5f" opacity=".5"/>
      </svg>
      <div style="position:absolute;bottom:-8px;left:0;right:0;text-align:center;color:#fff;font-size:9px;opacity:.5;letter-spacing:.02em">
        Tracklink Chile Fleet Dashboard By Würfel SpA · generado automáticamente por IA
      </div>
    </div>
  </div>
</div>

<!-- RESUMEN EJECUTIVO -->
<div class="page"><div class="pi">
  <h2 class="pg-title">Resumen Ejecutivo</h2>
  <p class="pg-sub">Los indicadores clave del período reflejan la concentración de incidencias en las unidades operativas.</p>
  <div class="kpi-grid">
    <div class="kpi">
      <div class="kpi-val">${s.totalIncidencias}</div>
      <div class="kpi-lbl">Total de excesos</div>
      <div class="kpi-desc">Incidencias registradas del ${s.startDisplay} al ${s.endDisplay}</div>
    </div>
    <div class="kpi">
      <div class="kpi-val">${s.conductoresArr.length}</div>
      <div class="kpi-lbl">Conductores</div>
      <div class="kpi-desc">Conductores con excesos de velocidad registrados en el período</div>
    </div>
    <div class="kpi">
      <div class="kpi-val">${s.globalMaxSpeed}</div>
      <div class="kpi-lbl">Vel. máx. (km/h)</div>
      <div class="kpi-desc">Registrada por ${s.globalMaxConductor}${s.globalMaxTime!=='—' ? ' el '+s.globalMaxTime : ''}</div>
    </div>
    <div class="kpi">
      <div class="kpi-val">${s.peakDay.count}</div>
      <div class="kpi-lbl">Pico diario</div>
      <div class="kpi-desc">Excesos el ${s.peakDay.label} — el día de mayor concentración</div>
    </div>
  </div>
  <div class="alert alert-yellow">
    <span>⚠</span>
    <div>La franja horaria más crítica fue <strong>${peakHourLabel}</strong>, con <strong>${peakHourCount} incidencias</strong>.${s.mananaPct>0 ? ` El bloque 08:00–13:00 concentra el <strong>${s.mananaPct}%</strong> del total semanal.` : ''}</div>
  </div>
  <div class="pf">${footer}</div>
</div></div>

<!-- CONDUCTORES -->
<div class="page"><div class="pi">
  <h2 class="pg-title">Incidencias y Velocidades por Conductor</h2>
  <p class="pg-sub">El análisis identifica a los conductores responsables del total de excesos registrados durante el período.</p>
  <div class="conductores-row">${conductorCards}</div>
  ${conducInsight}
  <div class="pf">${footer}</div>
</div></div>

<!-- DISTRIBUCIÓN -->
<div class="page"><div class="pi">
  <h2 class="pg-title">Distribución Diaria y Horaria de Excesos</h2>
  <div class="charts-row">
    <div class="chart-sec">
      <h3>Concentración por día</h3>
      <div class="chart-wrap"><canvas id="chartDays"></canvas></div>
      <p class="chart-note">El ${s.peakDay.label} concentró el mayor pico con <strong>${s.peakDay.count} excesos</strong>.
        Días de menor actividad: ${s.sortedDays.slice(-2).reverse().map(d=>d.label+' ('+d.count+')').join(' y ')}.</p>
    </div>
    <div class="chart-sec">
      <h3>Franjas horarias críticas</h3>
      <div class="chart-wrap"><canvas id="chartHours"></canvas></div>
      <p class="chart-note">La franja <strong>${peakHourLabel}</strong> es la más crítica con <strong>${peakHourCount} incidencias</strong>.${s.mananaPct>0 ? ` El bloque 08:00–13:00 concentra el ${s.mananaPct}% del total.` : ''}</p>
    </div>
  </div>
  <div class="pf">${footer}</div>
</div></div>

<!-- CONCLUSIONES -->
<div class="page"><div class="pi">
  <h2 class="pg-title">Conclusiones y Recomendaciones</h2>
  <div class="concl-grid">
    <div class="concl-card">
      <h4>Intervención inmediata con conductores</h4>
      <p>${conclusionConductores} deben ser convocados a sesión de retroalimentación individual para reforzar los protocolos de velocidad permitida en ruta.</p>
    </div>
    <div class="concl-card">
      <h4>Refuerzo en franja horaria matutina</h4>
      <p>${s.mananaPct>0 ? `El bloque 08:00–13:00 concentra el ${s.mananaPct}% de los excesos.` : 'La franja matutina concentra la mayor parte de los excesos.'} Se recomienda reforzar recordatorios de conducción segura al inicio del turno.</p>
    </div>
    <div class="concl-card">
      <h4>Monitoreo especial días críticos</h4>
      <p>${diasCriticos} son los días de mayor concentración. Se sugiere implementar alertas automáticas o supervisión activa en esas jornadas para prevenir reincidencias.</p>
    </div>
    <div class="concl-card">
      <h4>Seguimiento del próximo período</h4>
      <p>Se recomienda establecer un umbral de tolerancia semanal y comparar los resultados del siguiente informe para medir el impacto de las acciones correctivas.</p>
    </div>
  </div>
  <div class="alert alert-blue">
    <span>ⓘ</span>
    <div>Período: <strong>${s.startDisplay} al ${s.endDisplay}</strong> · Total: <strong>${s.totalIncidencias}</strong> · ${s.conductoresArr.map(c=>`<strong>${c.name}</strong>${c.vehicle?' ('+c.vehicle+')':''}`).join(' y ')}</div>
  </div>
  <div class="pf">${footer}</div>
</div></div>

<script>
window.__chartsReady = false;
(function() {
  Chart.defaults.font.family = "'Segoe UI', system-ui, Arial, sans-serif";
  Chart.defaults.font.size = 11;
  new Chart(document.getElementById('chartDays').getContext('2d'), {
    type:'bar', data:{ labels:${dayLabels}, datasets:[{ data:${dayCounts}, backgroundColor:${dayColors}, borderRadius:3, barThickness:26 }] },
    options:{ responsive:true, maintainAspectRatio:false, plugins:{legend:{display:false}},
      scales:{ x:{grid:{display:false},ticks:{font:{size:10},maxRotation:0}}, y:{grid:{color:'#f1f5f9'},ticks:{font:{size:10}},beginAtZero:true} } }
  });
  new Chart(document.getElementById('chartHours').getContext('2d'), {
    type:'bar', data:{ labels:${hourLabels}, datasets:[{ data:${hourCounts}, backgroundColor:${hourColors}, borderRadius:3, barThickness:14 }] },
    options:{ indexAxis:'y', responsive:true, maintainAspectRatio:false, plugins:{legend:{display:false}},
      scales:{ x:{grid:{color:'#f1f5f9'},ticks:{font:{size:10}},beginAtZero:true}, y:{grid:{display:false},ticks:{font:{size:10}}} } }
  });
  window.__chartsReady = true;
})();
</script>
</body></html>`;
}

// ─────────────────────────────────────────────────────────────────────────────
main().catch(err => { console.error('[pdf] ERROR FATAL:', err.message); process.exit(1); });
