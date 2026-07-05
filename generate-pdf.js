#!/usr/bin/env node
'use strict';
/**
 * generate-pdf.js
 * Genera el Reporte Semanal de Excesos de Velocidad en PDF (5 páginas,
 * formato tipo "documento Gamma" — una sección por página).
 *
 * Variables de entorno esperadas:
 *   REPORT_START — "YYYY-MM-DD"
 *   REPORT_END   — "YYYY-MM-DD"
 *
 * Salida: reporte-semanal.pdf
 */

const XLSX = require('xlsx');
const fs   = require('fs');
const path = require('path');

const EXCEL_FILE = path.join(process.cwd(), 'INFORME EXCESOS DE VELOCIDAD.xlsx');
const OUTPUT_PDF = path.join(process.cwd(), 'reporte-semanal.pdf');

const SITE_NAME = 'Relleno Sanitario Santa Marta';
const DIAS_ES      = ['Dom', 'Lun', 'Mar', 'Mié', 'Jue', 'Vie', 'Sáb'];
const DIAS_ES_FULL = ['domingo', 'lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado'];
const MESES_ES     = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];

// Corrección cosmética de tildes para la nómina conocida de conductores.
// (Los datos crudos de TrackGTS vienen sin tildes / en mayúsculas.)
// Agregar nuevos conductores acá a medida que se sumen a la flota.
const NAME_ACCENT_FIX = {
  'NICOLAS VARGAS':       'Nicolás Vargas',
  'JAVIER FARIAS':        'Javier Farías',
  'JUAN PABLO CABRERA':   'Juan Pablo Cabrera',
  'DANIEL HERMOSILLA':    'Daniel Hermosilla',
  'DANILO LAGOS':         'Danilo Lagos',
  'SERGIO NIEDBALSKI':    'Sergio Niedbalski',
  'ALONSO RODRIGUEZ':     'Alonso Rodríguez',
  'HENRRY RIVERO':        'Henrry Rivero',
  'DIEGO HERMOSILLA':     'Diego Hermosilla',
  'CLAUDIO GUTIERREZ':    'Claudio Gutiérrez',
};

const PAGE_W = 1280;
const PAGE_H = 720;

// ─────────────────────────────────────────────────────────────────────────────
async function main() {
  const startDate = process.env.REPORT_START;
  const endDate   = process.env.REPORT_END;
  if (!startDate || !endDate) throw new Error('Faltan variables REPORT_START y REPORT_END.');

  console.log(`[pdf] Generando reporte: ${startDate} → ${endDate}`);
  if (!fs.existsSync(EXCEL_FILE)) throw new Error(`No se encontró: ${EXCEL_FILE}`);

  const { rows, columns } = parseAndFilter(fs.readFileSync(EXCEL_FILE), startDate, endDate);
  console.log(`[pdf] Filas en el período: ${rows.length}`);
  if (rows.length === 0) console.warn('[pdf] Sin datos para el período — se genera PDF con aviso.');

  const stats = computeStats(rows, columns, startDate, endDate);
  const html  = generateHTML(stats);

  fs.writeFileSync(path.join(process.cwd(), 'reporte-preview.html'), html);

  const puppeteer = require('puppeteer');
  console.log('[pdf] Iniciando Puppeteer...');
  const browser = await puppeteer.launch({
    headless: 'new',
    args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage'],
  });
  try {
    const page = await browser.newPage();
    await page.setViewport({ width: PAGE_W, height: PAGE_H });
    await page.setContent(html, { waitUntil: 'networkidle0', timeout: 40000 });
    const pdf = await page.pdf({
      width: `${PAGE_W}px`, height: `${PAGE_H}px`,
      printBackground: true,
      margin: { top: 0, right: 0, bottom: 0, left: 0 },
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
  const sheetName = wb.SheetNames.find((n) => n.toLowerCase().includes('detalle')) || wb.SheetNames[0];
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

  const aliasIdx     = headers.indexOf('Alias');
  const conductorIdx = headers.findIndex((h) => h.toLowerCase() === 'conductor');
  const fechaIdx     = headers.findIndex((h) => h.toLowerCase().includes('fecha') && h.toLowerCase().includes('inicio'));
  const velIdx       = headers.findIndex((h) => /veloc.*m[aá]x/i.test(h));

  if (velIdx === -1) throw new Error('No se encontró la columna de "Velocidad Máxima".');

  const columns = {
    alias:     headers[aliasIdx]     || 'Alias',
    conductor: conductorIdx >= 0 ? headers[conductorIdx] : null,
    fecha:     headers[fechaIdx]     || 'Fecha de Inicio',
    velocidad: headers[velIdx],
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

    const aliasVal     = String(obj[columns.alias] || '').trim();
    const conductorRaw = columns.conductor ? String(obj[columns.conductor] || '').trim() : '';

    obj['__fecha']     = fecha;
    obj['__vel']       = parseFloat(String(obj[columns.velocidad] || '0').replace(',', '.')) || 0;
    obj['__alias']     = aliasVal;
    // Si no hay nombre de conductor, se usa el alias del vehículo (instrucción del prompt original)
    // y se marca para no formatearlo como si fuera un nombre de persona.
    obj['__conductor']     = conductorRaw || aliasVal;
    obj['__vehicleOnly']   = !conductorRaw;
    rows.push(obj);
  }

  return { rows, columns };
}

function parseFecha(val) {
  if (!val) return null;
  if (val instanceof Date) return isNaN(val) ? null : val;
  const s = String(val).trim();
  const p1 = s.match(/^(\d{4})[\/\-](\d{2})[\/\-](\d{2})\s+(\d{2}):(\d{2}):(\d{2})/);
  if (p1) return new Date(+p1[1], +p1[2] - 1, +p1[3], +p1[4], +p1[5], +p1[6]);
  const p2 = s.match(/^(\d{2})[\/\-](\d{2})[\/\-](\d{4})\s+(\d{2}):(\d{2}):(\d{2})/);
  if (p2) return new Date(+p2[3], +p2[2] - 1, +p2[1], +p2[4], +p2[5], +p2[6]);
  const d = new Date(val);
  return isNaN(d) ? null : d;
}

// ─────────────────────────────────────────────────────────────────────────────
// HELPERS DE TEXTO
// ─────────────────────────────────────────────────────────────────────────────
function pad(n) { return String(n).padStart(2, '0'); }

function splitAlias(alias) {
  const parts = String(alias || '').trim().split(/\s+/).filter(Boolean);
  if (parts.length === 0) return { model: '', code: '' };
  const code  = parts[parts.length - 1];
  const model = parts.slice(0, -1).join(' ');
  return { model, code };
}

function titleCase(name) {
  return String(name || '').trim().split(/\s+/).filter(Boolean)
    .map((w) => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase()).join(' ');
}

function displayName(name) {
  const key = String(name || '').trim().toUpperCase().replace(/\s+/g, ' ');
  return NAME_ACCENT_FIX[key] || titleCase(name);
}

function abbreviateName(name) {
  const full = displayName(name); // nombre ya con Mayúscula/tilde aplicada
  const words = full.trim().split(/\s+/).filter(Boolean);
  if (words.length === 0) return full;
  const surname    = words[words.length - 1];
  const firstWords = words.slice(0, -1);
  const initials   = firstWords.map((w) => w[0].toUpperCase()).join('');
  return initials ? `${initials}. ${surname}` : surname;
}

function ordinal(n) {
  const words = ['primer', 'segundo', 'tercer', 'cuarto', 'quinto', 'sexto', 'séptimo', 'octavo', 'noveno', 'décimo'];
  return words[n - 1] || `${n}°`;
}

function fmtSpeed(v) { return Number(v).toFixed(1).replace('.', ','); }
function fmtPct(v)   { return Number(v).toFixed(1).replace('.', ','); }

function formatVerboseRange(startDate, endDate) {
  const [sy, sm, sd] = startDate.split('-').map(Number);
  const [ey, em, ed] = endDate.split('-').map(Number);
  if (sy === ey && sm === em) {
    return `${sd} al ${ed} de ${MESES_ES[em - 1]} de ${ey}`;
  }
  if (sy === ey) {
    return `${sd} de ${MESES_ES[sm - 1]} – ${ed} de ${MESES_ES[em - 1]} de ${ey}`;
  }
  return `${sd} de ${MESES_ES[sm - 1]} de ${sy} – ${ed} de ${MESES_ES[em - 1]} de ${ey}`;
}

function escapeHtml(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

// ─────────────────────────────────────────────────────────────────────────────
// ESTADÍSTICAS
// ─────────────────────────────────────────────────────────────────────────────
function computeStats(rows, columns, startDate, endDate) {
  const conductores = {};
  const byDayMap  = {};
  const byHourArr = new Array(24).fill(0);

  rows.forEach((row) => {
    const key   = row['__conductor'];
    const d     = row['__fecha'];
    const speed = row['__vel'];
    const alias = row['__alias'];

    if (!conductores[key]) {
      conductores[key] = { count: 0, maxSpeed: 0, maxSpeedTime: null, aliasCounts: {}, vehicleOnly: false };
    }
    const c = conductores[key];
    c.count++;
    c.aliasCounts[alias] = (c.aliasCounts[alias] || 0) + 1;
    if (row['__vehicleOnly']) c.vehicleOnly = true;
    if (speed > c.maxSpeed) { c.maxSpeed = speed; c.maxSpeedTime = d; }

    const dayKey = `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
    byDayMap[dayKey] = (byDayMap[dayKey] || 0) + 1;
    byHourArr[d.getHours()]++;
  });

  const totalIncidencias = rows.length;

  const conductoresArr = Object.entries(conductores)
    .map(([name, data]) => {
      const topAlias = Object.entries(data.aliasCounts).sort((a, b) => b[1] - a[1])[0];
      const { model, code } = splitAlias(topAlias ? topAlias[0] : '');
      const label = data.vehicleOnly ? `Unidad ${code || name}` : displayName(name);
      return {
        rawName: name,
        name: label,
        nameAbbr: data.vehicleOnly ? (code || name) : abbreviateName(name),
        nameUpper: data.vehicleOnly ? (code || String(name).toUpperCase()) : String(name).toUpperCase(),
        unitCode: code,
        unitModel: model,
        vehicleOnly: data.vehicleOnly,
        count: data.count,
        pct: totalIncidencias ? (data.count / totalIncidencias) * 100 : 0,
        maxSpeed: data.maxSpeed,
        maxSpeedTime: data.maxSpeedTime,
      };
    })
    .sort((a, b) => b.count - a.count);

  // Días del período (largo variable, según rango real)
  const [sy, sm, sd] = startDate.split('-').map(Number);
  const dayCount = Math.round((new Date(endDate).setHours(12) - new Date(startDate).setHours(12)) / 86400000) + 1;
  const weekDays = Array.from({ length: Math.max(dayCount, 1) }, (_, i) => {
    const d   = new Date(sy, sm - 1, sd + i);
    const key = `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
    return { key, date: d, label: `${DIAS_ES[d.getDay()]} ${pad(d.getDate())}/${pad(d.getMonth() + 1)}`, count: byDayMap[key] || 0 };
  });
  const sortedDays = weekDays.slice().sort((a, b) => b.count - a.count);
  const peakDay = sortedDays[0] || { label: '—', count: 0, key: '' };

  // Horas (24 franjas completas)
  const hours = Array.from({ length: 24 }, (_, h) => ({ hour: h, label: pad(h), fullLabel: `${pad(h)}:00–${pad(h)}:59`, count: byHourArr[h] }));
  const topHours = hours.slice().sort((a, b) => b.count - a.count);
  const peakHour = topHours[0] || { hour: 0, fullLabel: '—', count: 0 };

  // Ventana operativa "núcleo" — la franja contigua más corta que cubre ≥75% de las incidencias.
  const coreWindow = findCoreWindow(byHourArr, totalIncidencias);

  // Máxima velocidad global (puede no ser el conductor con más excesos).
  const globalMaxEntry = conductoresArr.slice().sort((a, b) => b.maxSpeed - a.maxSpeed)[0] || null;

  const top3Count = conductoresArr.slice(0, 3).reduce((s, c) => s + c.count, 0);
  const top3Pct   = totalIncidencias ? (top3Count / totalIncidencias) * 100 : 0;

  const [ey, em, ed] = endDate.split('-').map(Number);
  return {
    startDate, endDate,
    startDisplay: `${pad(sd)}/${pad(sm)}/${sy}`,
    endDisplay:   `${pad(ed)}/${pad(em)}/${ey}`,
    rangeVerbose: formatVerboseRange(startDate, endDate),
    totalIncidencias,
    conductoresArr,
    weekDays, sortedDays, peakDay,
    hours, topHours, peakHour,
    coreWindow,
    top3Count, top3Pct,
    globalMaxSpeed:     globalMaxEntry ? globalMaxEntry.maxSpeed : 0,
    globalMaxConductor: globalMaxEntry ? globalMaxEntry.name     : '—',
    globalMaxUnit:      globalMaxEntry ? globalMaxEntry.unitCode : '—',
    globalMaxModel:     globalMaxEntry ? globalMaxEntry.unitModel: '',
    globalMaxRank:      globalMaxEntry ? conductoresArr.findIndex((c) => c.rawName === globalMaxEntry.rawName) + 1 : 0,
  };
}

function findCoreWindow(hourCounts, total) {
  let best = null;
  for (let len = 4; len <= 18; len++) {
    for (let start = 0; start + len - 1 <= 23; start++) {
      const end = start + len - 1;
      let sum = 0;
      for (let h = start; h <= end; h++) sum += hourCounts[h];
      const coverage = total ? sum / total : 0;
      if (coverage >= 0.75) { best = { start, end, sum, coverage }; break; }
    }
    if (best) break;
  }
  if (!best) {
    let bestSum = -1, bestStart = 0;
    for (let start = 0; start <= 16; start++) {
      let sum = 0;
      for (let h = start; h < start + 8; h++) sum += hourCounts[h];
      if (sum > bestSum) { bestSum = sum; bestStart = start; }
    }
    best = { start: bestStart, end: bestStart + 7, sum: bestSum, coverage: total ? bestSum / total : 0 };
  }
  return best;
}

// ─────────────────────────────────────────────────────────────────────────────
// GRÁFICOS SVG (sin dependencias externas — se renderizan igual en cualquier motor)
// ─────────────────────────────────────────────────────────────────────────────
function hexToRgb(hex) { const n = parseInt(hex.replace('#', ''), 16); return { r: (n >> 16) & 255, g: (n >> 8) & 255, b: n & 255 }; }
function lerpColor(c1, c2, t) {
  const p1 = hexToRgb(c1), p2 = hexToRgb(c2);
  const r = Math.round(p1.r + (p2.r - p1.r) * t);
  const g = Math.round(p1.g + (p2.g - p1.g) * t);
  const b = Math.round(p1.b + (p2.b - p1.b) * t);
  return `rgb(${r},${g},${b})`;
}
function rankColor(rank, n) {
  if (n <= 1) return '#2d3748';
  return lerpColor('#2d3748', '#cbd5e0', rank / (n - 1));
}
function valueRankColors(values) {
  const idx = values.map((v, i) => i).sort((a, b) => values[b] - values[a]);
  const colors = new Array(values.length);
  idx.forEach((origIdx, rank) => { colors[origIdx] = rankColor(rank, values.length); });
  return colors;
}
function niceTicks(maxVal, targetCount) {
  maxVal = Math.max(maxVal, 1);
  const rawStep = maxVal / targetCount;
  const mag  = Math.pow(10, Math.floor(Math.log10(rawStep)));
  const norm = rawStep / mag;
  let step;
  if (norm < 1.5) step = 1 * mag; else if (norm < 3) step = 2 * mag; else if (norm < 7) step = 5 * mag; else step = 10 * mag;
  const max = Math.ceil(maxVal / step) * step;
  const ticks = [];
  for (let v = 0; v <= max + 1e-9; v += step) ticks.push(Math.round(v));
  return { ticks, max };
}

function horizontalBarChart(items, { width, height, labelWidth = 190 }) {
  const values = items.map((i) => i.value);
  const colors = valueRankColors(values);
  const { ticks, max } = niceTicks(Math.max(...values, 1), 6);
  const plotX = labelWidth, plotW = width - labelWidth - 46;
  const topPad = 6, bottomPad = 26, plotH = height - topPad - bottomPad;
  const rowH = plotH / items.length;
  const barH = Math.min(24, rowH * 0.58);

  let grid = '', axis = '', bars = '';
  ticks.forEach((t) => {
    const x = plotX + (t / max) * plotW;
    grid += `<line x1="${x.toFixed(1)}" y1="${topPad}" x2="${x.toFixed(1)}" y2="${topPad + plotH}" stroke="#e2e8f0" stroke-width="1" stroke-dasharray="3,3"/>`;
    axis += `<text x="${x.toFixed(1)}" y="${topPad + plotH + 19}" font-size="10.5" fill="#94a3b8" text-anchor="middle" font-family="Inter,sans-serif">${t}</text>`;
  });
  items.forEach((it, i) => {
    const y = topPad + i * rowH + (rowH - barH) / 2;
    const w = Math.max((it.value / max) * plotW, 2);
    bars += `<text x="${plotX - 12}" y="${(y + barH / 2 + 4).toFixed(1)}" font-size="11.5" font-weight="700" fill="#334155" text-anchor="end" font-family="Inter,sans-serif">${escapeHtml(it.label)}</text>`;
    bars += `<rect x="${plotX}" y="${y.toFixed(1)}" width="${w.toFixed(1)}" height="${barH}" rx="4" fill="${colors[i]}"/>`;
    bars += `<text x="${(plotX + w + 9).toFixed(1)}" y="${(y + barH / 2 + 4).toFixed(1)}" font-size="11.5" font-weight="700" fill="#1a202c" font-family="Inter,sans-serif">${it.value}</text>`;
  });
  return `<svg viewBox="0 0 ${width} ${height}" width="${width}" height="${height}" xmlns="http://www.w3.org/2000/svg">${grid}${bars}${axis}</svg>`;
}

function verticalBarChart(items, { width, height, showValueLabels = false, peakBadge = null }) {
  const values = items.map((i) => i.value);
  const colors = valueRankColors(values);
  const { ticks, max } = niceTicks(Math.max(...values, 1), 5);
  const leftPad = 32, rightPad = 8, topPad = peakBadge ? 40 : (showValueLabels ? 30 : 12), bottomPad = 26;
  const plotW = width - leftPad - rightPad, plotH = height - topPad - bottomPad;
  const n = items.length, slot = plotW / n;
  const barW = Math.min(slot * 0.62, 46);

  let grid = '', axisY = '', axisX = '', bars = '', peakLine = '';
  ticks.forEach((t) => {
    const y = topPad + plotH - (t / max) * plotH;
    grid  += `<line x1="${leftPad}" y1="${y.toFixed(1)}" x2="${leftPad + plotW}" y2="${y.toFixed(1)}" stroke="#e2e8f0" stroke-width="1" stroke-dasharray="3,3"/>`;
    axisY += `<text x="${leftPad - 8}" y="${(y + 3).toFixed(1)}" font-size="10" fill="#94a3b8" text-anchor="end" font-family="Inter,sans-serif">${t}</text>`;
  });

  let peakIdx = 0;
  items.forEach((it, i) => { if (it.value > items[peakIdx].value) peakIdx = i; });

  items.forEach((it, i) => {
    const x = leftPad + i * slot + (slot - barW) / 2;
    const h = (it.value / max) * plotH;
    const y = topPad + plotH - h;
    bars += `<rect x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${barW.toFixed(1)}" height="${Math.max(h,1).toFixed(1)}" rx="3" fill="${colors[i]}"/>`;
    if (showValueLabels) {
      bars += `<text x="${(x + barW / 2).toFixed(1)}" y="${(y - 7).toFixed(1)}" font-size="11.5" font-weight="700" fill="#1a202c" text-anchor="middle" font-family="Inter,sans-serif">${it.value}</text>`;
    }
    axisX += `<text x="${(x + barW / 2).toFixed(1)}" y="${(topPad + plotH + 17).toFixed(1)}" font-size="9.5" fill="#94a3b8" text-anchor="middle" font-family="Inter,sans-serif">${escapeHtml(it.label)}</text>`;
  });

  if (peakBadge) {
    const px = leftPad + peakIdx * slot + slot / 2;
    const py = topPad + plotH - (items[peakIdx].value / max) * plotH;
    peakLine += `<line x1="${leftPad}" y1="${py.toFixed(1)}" x2="${leftPad + plotW}" y2="${py.toFixed(1)}" stroke="#1a202c" stroke-width="1" stroke-dasharray="4,3" opacity="0.55"/>`;
    const bw = 118, bh = 24, bx = leftPad, by = Math.max(py - bh - 6, 2);
    peakLine += `<rect x="${bx}" y="${by.toFixed(1)}" width="${bw}" height="${bh}" rx="5" fill="#1a202c"/>`;
    peakLine += `<text x="${bx + bw / 2}" y="${(by + bh / 2 + 4).toFixed(1)}" font-size="11" font-weight="700" fill="#fff" text-anchor="middle" font-family="Inter,sans-serif">${escapeHtml(peakBadge)}</text>`;
  }

  return `<svg viewBox="0 0 ${width} ${height}" width="${width}" height="${height}" xmlns="http://www.w3.org/2000/svg">${grid}${bars}${axisX}${axisY}${peakLine}</svg>`;
}

// ─────────────────────────────────────────────────────────────────────────────
// HTML — 5 páginas 1280×720 (16:9), una sección por página
// ─────────────────────────────────────────────────────────────────────────────
function generateHTML(s) {
  const footer = `<div class="tl-footer">Tracklink Chile Fleet Dashboard · Würfel SPA · ${SITE_NAME} · ${s.startDisplay} — ${s.endDisplay}</div>`;

  // ── Página 3: gráfico + tabla de conductores ──────────────────────────
  const conductorChartItems = s.conductoresArr.map((c) => ({ label: c.nameUpper, value: c.count }));
  const conductorChartSvg   = horizontalBarChart(conductorChartItems, { width: 620, height: 400 });

  const tableRows = s.conductoresArr.map((c) => {
    const isMax = c.rawName && s.globalMaxConductor !== '—' && c.name === s.globalMaxConductor;
    return `<tr class="${isMax ? 'row-max' : ''}">
      <td>${escapeHtml(c.nameAbbr)}</td>
      <td>${escapeHtml(c.unitCode || '—')}</td>
      <td class="num">${fmtSpeed(c.maxSpeed)} km/h</td>
    </tr>`;
  }).join('');

  const top3Names = s.conductoresArr.slice(0, 3).map((c) => c.name);
  const top3Text  = top3Names.length >= 2
    ? `${top3Names.slice(0, -1).join(', ')} y ${top3Names[top3Names.length - 1]}`
    : (top3Names[0] || '—');

  // ── Página 4: distribución horaria ─────────────────────────────────────
  const hourChartItems = s.hours.map((h) => ({ label: h.label, value: h.count }));
  const hourChartSvg   = verticalBarChart(hourChartItems, { width: 1168, height: 380, peakBadge: `Pico máximo: ${s.peakHour.count}` });
  const top3Hours = s.topHours.slice(0, 3);

  // ── Página 5: días + acciones ───────────────────────────────────────────
  const dayChartItems = s.sortedDays.map((d) => ({ label: d.label, value: d.count }));
  const dayChartSvg    = verticalBarChart(dayChartItems, { width: 600, height: 360, showValueLabels: true });

  const lowestDays = s.sortedDays.slice(-2).reverse().map((d) => `${d.label} (${d.count})`).join(' y ');

  const criticalDayNames = s.sortedDays.slice(0, 2).map((d) => DIAS_ES_FULL[d.date.getDay()]);
  const criticalDayText  = criticalDayNames.length >= 2 ? `${criticalDayNames[0]} y ${criticalDayNames[1]}` : (criticalDayNames[0] || '—');

  const coreLabel = `${pad(s.coreWindow.start)}:00–${pad(s.coreWindow.end)}:59`;
  const corePct   = fmtPct(s.coreWindow.coverage * 100);

  return `<!DOCTYPE html>
<html lang="es"><head><meta charset="UTF-8">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Poppins:wght@600;700;800&display=swap" rel="stylesheet">
<style>
  *,*::before,*::after{box-sizing:border-box;margin:0;padding:0;}
  @page{size:${PAGE_W}px ${PAGE_H}px;margin:0;}
  html,body{height:100%;}
  body{font-family:'Inter',system-ui,Arial,sans-serif;background:#fff;color:#1a202c;}
  .page{width:${PAGE_W}px;height:${PAGE_H}px;position:relative;overflow:hidden;page-break-after:always;background:#fff;}
  .page:last-child{page-break-after:avoid;}
  h1,h2,h3,.num-font{font-family:'Poppins',sans-serif;}

  /* Portada */
  .cover{display:flex;width:100%;height:100%;}
  .cv-left{width:46%;height:100%;background:linear-gradient(150deg,#1a2b45 0%,#22406f 45%,#0f1e34 100%);position:relative;overflow:hidden;display:flex;align-items:center;justify-content:center;}
  .cv-right{width:54%;height:100%;padding:64px 68px;display:flex;flex-direction:column;justify-content:center;gap:20px;}
  .cv-eyebrow{display:flex;align-items:center;gap:16px;margin-bottom:6px;}
  .cv-title{font-size:38px;font-weight:800;color:#1a202c;line-height:1.15;}
  .cv-sub{font-size:15px;font-weight:600;color:#4a5568;}
  .cv-desc{font-size:13.5px;color:#718096;line-height:1.7;max-width:520px;}

  /* Páginas interiores */
  .pi{padding:46px 64px 40px;height:100%;display:flex;flex-direction:column;}
  .pg-title{font-size:28px;font-weight:800;color:#1a202c;margin-bottom:14px;}
  .pg-intro{font-size:13.5px;color:#4a5568;line-height:1.65;margin-bottom:22px;max-width:1130px;}
  .pg-intro strong{color:#1a202c;}
  .pf{position:absolute;bottom:16px;left:0;right:0;display:flex;justify-content:center;}
  .tl-footer{font-size:9.5px;color:#cbd5e0;letter-spacing:.03em;text-align:center;}

  .alert{padding:14px 18px;display:flex;gap:12px;align-items:flex-start;font-size:13px;line-height:1.5;border-radius:8px;}
  .alert span{flex-shrink:0;font-size:16px;margin-top:1px;}
  .alert-yellow{background:#fefce8;border-left:4px solid #eab308;color:#4a5568;}
  .alert-yellow strong{color:#1a202c;}
  .note-box{padding:14px 18px;display:flex;gap:12px;align-items:flex-start;font-size:12.5px;line-height:1.5;border-radius:8px;background:#eef2f7;color:#4a5568;}
  .note-box strong{color:#1a202c;}

  /* Página 2 — KPIs */
  .kpi-grid{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:18px;}
  .kpi{border:1px solid #e2e8f0;border-radius:10px;padding:20px 26px;text-align:center;}
  .kpi-val{font-size:44px;font-weight:800;color:#2d3748;line-height:1;margin-bottom:6px;}
  .kpi-lbl{font-size:14px;font-weight:700;color:#374151;margin-bottom:5px;}
  .kpi-desc{font-size:11.5px;color:#94a3b8;}

  /* Página 3 */
  .p3-row{display:flex;gap:32px;flex:1;min-height:0;align-items:stretch;}
  .p3-chart{flex:1 1 56%;}
  .p3-table-wrap{flex:1 1 44%;display:flex;flex-direction:column;}
  table.cond-table{width:100%;border-collapse:collapse;font-size:12px;}
  .cond-table thead th{font-size:11px;font-weight:700;color:#fff;background:#374151;text-align:left;padding:9px 12px;}
  .cond-table td{padding:8px 12px;border-bottom:1px solid #edf2f7;color:#334155;}
  .cond-table td.num{font-weight:600;}
  .cond-table tr.row-max td{font-weight:800;color:#1a202c;}
  .p3-note{font-size:12px;color:#718096;line-height:1.6;margin-top:16px;}

  /* Página 4 */
  .p4-chart-wrap{flex:1;display:flex;align-items:center;}

  /* Página 5 */
  .p5-row{display:flex;gap:32px;flex:1;min-height:0;}
  .p5-chart-col{flex:1 1 48%;display:flex;flex-direction:column;}
  .p5-chart-title{font-size:13px;font-weight:700;color:#374151;margin-bottom:8px;}
  .p5-actions-col{flex:1 1 52%;}
  .concl-grid{display:grid;grid-template-columns:1fr 1fr;gap:14px;}
  .concl-card{border-left:4px solid #2d3748;background:#f8fafc;border-radius:0 8px 8px 0;padding:16px 18px;}
  .concl-card h4{font-size:13px;font-weight:700;color:#2d3748;margin-bottom:6px;}
  .concl-card p{font-size:11.5px;color:#718096;line-height:1.55;}
</style>
</head><body>

<!-- PÁGINA 1 — PORTADA -->
<div class="page cover">
  <div class="cv-left">
    <svg width="360" height="230" viewBox="0 0 360 230" fill="none">
      <rect x="0" y="182" width="360" height="4" fill="#fff" opacity=".35"/>
      <path d="M40 182 L80 110 L140 110 L165 60 L230 60 L250 110 L320 110 L320 182 Z" fill="#fff" opacity=".14"/>
      <rect x="60" y="90" width="210" height="92" rx="4" fill="#fff" opacity=".5"/>
      <rect x="215" y="70" width="58" height="102" rx="3" fill="#fff" opacity=".72"/>
      <rect x="222" y="78" width="44" height="36" rx="2" fill="#93c5fd" opacity=".55"/>
      <circle cx="105" cy="188" r="17" fill="#0f1e34"/><circle cx="105" cy="188" r="8" fill="#fff" opacity=".55"/>
      <circle cx="240" cy="188" r="17" fill="#0f1e34"/><circle cx="240" cy="188" r="8" fill="#fff" opacity=".55"/>
      <rect x="0" y="140" width="55" height="42" rx="3" fill="#fff" opacity=".22"/>
      <rect x="270" y="130" width="10" height="14" rx="1" fill="#fbbf24" opacity=".85"/>
    </svg>
  </div>
  <div class="cv-right">
    <div class="cv-eyebrow">
      <img src="https://raw.githubusercontent.com/WurfelSPA/tracklink-santamarta/main/logo.png" style="height:44px;width:auto;object-fit:contain;" alt="">
      <svg width="98" height="24" viewBox="0 0 110 28"><text x="0" y="22" font-family="Arial Black,Arial,sans-serif" font-size="20" font-weight="900" fill="#1a202c">TRACK</text><text x="62" y="22" font-family="Arial Black,Arial,sans-serif" font-size="20" font-weight="900" fill="#2d7be5">LINK</text></svg>
    </div>
    <h1 class="cv-title">Reporte de Excesos de<br>Velocidad</h1>
    <p class="cv-sub">${SITE_NAME} · Período: ${s.rangeVerbose}</p>
    <p class="cv-desc">Durante la semana analizada se registraron un total de <strong>${s.totalIncidencias} excesos de velocidad</strong> en la flota vehicular de ${SITE_NAME.toLowerCase()}. Este reporte presenta un análisis detallado por conductor, franja horaria y día de la semana, con el objetivo de identificar patrones de riesgo y apoyar la toma de decisiones en materia de seguridad vial operacional.</p>
  </div>
</div>

<!-- PÁGINA 2 — RESUMEN EJECUTIVO -->
<div class="page"><div class="pi">
  <h2 class="pg-title">Resumen Ejecutivo</h2>
  <p class="pg-intro">Los cuatro indicadores clave del período revelan concentraciones críticas de riesgo que requieren atención inmediata por parte de la supervisión de flota.</p>
  <div class="kpi-grid">
    <div class="kpi"><div class="kpi-val">${s.totalIncidencias}</div><div class="kpi-lbl">Total de Excesos</div><div class="kpi-desc">Incidencias de velocidad registradas en el período completo</div></div>
    <div class="kpi"><div class="kpi-val">${fmtPct(s.conductoresArr[0]?.pct || 0)}%</div><div class="kpi-lbl">Conductor Crítico</div><div class="kpi-desc">${s.conductoresArr[0]?.name || '—'} concentra ${s.conductoresArr[0]?.count || 0} excesos del total semanal</div></div>
    <div class="kpi"><div class="kpi-val">${fmtSpeed(s.globalMaxSpeed)}</div><div class="kpi-lbl">Vel. Máx. (km/h)</div><div class="kpi-desc">Registrada por ${s.globalMaxConductor} en unidad ${s.globalMaxUnit}</div></div>
    <div class="kpi"><div class="kpi-val">${s.peakHour.count}</div><div class="kpi-lbl">Hora Pico (${pad(s.peakHour.hour)}:00h)</div><div class="kpi-desc">Franja horaria con mayor concentración de incidencias</div></div>
  </div>
  <div class="alert alert-yellow"><span class="icon-warn"><svg width="16" height="16" viewBox="0 0 24 24" fill="none"><path d="M12 3L22 20H2L12 3Z" stroke="#eab308" stroke-width="2" stroke-linejoin="round" fill="#fef9c3"/><path d="M12 10v4M12 17h.01" stroke="#a16207" stroke-width="2" stroke-linecap="round"/></svg></span><div>El ${DIAS_ES_FULL[s.peakDay.date ? s.peakDay.date.getDay() : 0]} ${s.peakDay.label.split(' ')[1] || ''} fue el día más crítico del período, concentrando <strong>${s.peakDay.count} excesos</strong> (${fmtPct(s.totalIncidencias ? s.peakDay.count / s.totalIncidencias * 100 : 0)}% del total semanal).</div></div>
  <div class="pf">${footer}</div>
</div></div>

<!-- PÁGINA 3 — CONDUCTORES -->
<div class="page"><div class="pi">
  <h2 class="pg-title">Excesos por Conductor y Velocidad Máxima</h2>
  <p class="pg-intro">La distribución de incidencias es marcadamente desigual: los tres conductores con mayor cantidad de excesos acumulan el <strong>${fmtPct(s.top3Pct)}% del total</strong>, lo que indica la necesidad de intervención focalizada. ${top3Text ? `${top3Names[0] || ''} lidera con ${s.conductoresArr[0]?.count || 0} excesos${top3Names[1] ? `, seguido por ${top3Names[1]} (${s.conductoresArr[1]?.count || 0})` : ''}${top3Names[2] ? ` y ${top3Names[2]} (${s.conductoresArr[2]?.count || 0})` : ''}.` : ''}</p>
  <div class="p3-row">
    <div class="p3-chart">${conductorChartSvg}</div>
    <div class="p3-table-wrap">
      <table class="cond-table"><thead><tr><th>Conductor</th><th>Unidad</th><th>Vel. Máx.</th></tr></thead><tbody>${tableRows}</tbody></table>
      <p class="p3-note">${s.globalMaxConductor} registró la velocidad puntual más alta del período (${fmtSpeed(s.globalMaxSpeed)} km/h), ${s.globalMaxRank > 0 ? `${s.globalMaxRank <= 3 ? 'ubicándose' : 'a pesar de ubicarse'} en el ${ordinal(s.globalMaxRank)} lugar por volumen de excesos.` : ''}</p>
    </div>
  </div>
  <div class="pf">${footer}</div>
</div></div>

<!-- PÁGINA 4 — DISTRIBUCIÓN HORARIA -->
<div class="page"><div class="pi">
  <h2 class="pg-title">Distribución Horaria de Excesos</h2>
  <p class="pg-intro">El <strong>${corePct}% de los excesos</strong> (${s.coreWindow.sum} incidencias) ocurre en la franja comprendida entre las ${coreLabel.replace('–',' y las ')} horas, coincidiendo con la jornada operativa principal. El pico máximo se registra a las ${s.peakHour.fullLabel} con ${s.peakHour.count} excesos${top3Hours[1] ? `, seguido de las ${top3Hours[1].fullLabel} con ${top3Hours[1].count}` : ''}${top3Hours[2] ? ` y las ${top3Hours[2].fullLabel} con ${top3Hours[2].count}` : ''}.</p>
  <div class="p4-chart-wrap">${hourChartSvg}</div>
  <div class="note-box"><span class="icon-info"><svg width="16" height="16" viewBox="0 0 24 24" fill="none"><circle cx="12" cy="12" r="9" stroke="#64748b" stroke-width="2" fill="#e2e8f0"/><path d="M12 11v5M12 8h.01" stroke="#475569" stroke-width="2" stroke-linecap="round"/></svg></span><div>La franja de mayor riesgo (${coreLabel}) coincide con las horas de mayor actividad de ingreso y egreso de vehículos. Se recomienda reforzar controles y señalización en este horario.</div></div>
  <div class="pf">${footer}</div>
</div></div>

<!-- PÁGINA 5 — DÍAS Y CONCLUSIONES -->
<div class="page"><div class="pi">
  <h2 class="pg-title">Concentración por Día y Conclusiones</h2>
  <div class="p5-row">
    <div class="p5-chart-col">
      <div class="p5-chart-title">Excesos por Día de la Semana</div>
      ${dayChartSvg}
    </div>
    <div class="p5-actions-col">
      <div class="concl-grid">
        <div class="concl-card"><h4>Capacitación Focalizada</h4><p>Priorizar a ${top3Text} en programas de conducción segura y manejo defensivo.</p></div>
        <div class="concl-card"><h4>Control en Horario Crítico</h4><p>Reforzar la supervisión entre las ${coreLabel}, especialmente los días ${criticalDayText}.</p></div>
        <div class="concl-card"><h4>Revisión de Unidad ${s.globalMaxUnit}</h4><p>Verificar el estado mecánico ${s.globalMaxModel ? `del ${titleCase(s.globalMaxModel)}` : 'de la unidad'} de ${s.globalMaxConductor.split(' ').pop()}, dado el registro de velocidad máxima de ${fmtSpeed(s.globalMaxSpeed)} km/h.</p></div>
        <div class="concl-card"><h4>Seguimiento Continuo</h4><p>Establecer alertas automáticas para conductores que superen umbrales de excesos semanales definidos por la supervisión.</p></div>
      </div>
    </div>
  </div>
  <div class="pf">${footer}</div>
</div></div>

</body></html>`;
}

module.exports = { parseAndFilter, computeStats, generateHTML, PAGE_W, PAGE_H };
if (require.main === module) {
  main().catch((err) => { console.error('[pdf] ERROR FATAL:', err.stack || err.message); process.exit(1); });
}
