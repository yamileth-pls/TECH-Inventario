/**
 * ==========================================================================
 *  TIEMPOS DE RESPUESTA SURTIDO  ·  Technochairs
 *  Backend (Google Apps Script)
 * ==========================================================================
 *  Hoja de datos:  "Inventario"  (columnas por posición)
 *    A  Prefactura
 *    C  Cliente
 *    D  Fecha
 *    E  Hora recibo
 *    N  Nave
 *    P  Hora de surtido
 *    R  Estatus almacén
 *    Y  Estatus de entrega
 *    AJ Semana
 *
 *  Reglas:
 *    - Se considera ENTREGADA una solicitud con estatus VALIDACION o ENTREGADO.
 *    - El tiempo de surtido solo cuenta horario laboral:
 *      Lunes a Viernes, 8:30 - 18:00, sin contar la hora de comida (13:00 - 14:00).
 *    - Meta: surtir en 5 horas hábiles (300 min).
 * ==========================================================================
 */

// ------------------------- CONFIGURACIÓN -------------------------
var SPREADSHEET_ID = '1we4yhPySbL-wODeZcrNMYvMZ0HmrHmS2sM1oLZsrQoM';
var SHEET_NAME     = 'Inventario';
var META_TITLE     = 'Tiempos de respuesta Surtido';
var META_SUBTITLE  = 'Consulta de Inventarios · Área de Almacenes';
var SLA_HOURS      = 5;                 // meta: horas hábiles para surtir
var SLA_MINUTES    = SLA_HOURS * 60;    // 300
var MAX_SPAN_DAYS  = 60;                // más allá de esto se considera dato corrupto, no un pedido real

// Estatus (normalizados, sin acentos ni mayúsculas) que cuentan como "entregado"
var ENTREGADO_ESTATUS = ['validacion', 'entregado'];

// Ventana laboral del día (horas locales)
var WORKDAY_SEGMENTS = [
  { startH: 8, startM: 30, endH: 13, endM: 0 },  // mañana
  { startH: 14, startM: 0, endH: 18, endM: 0 }   // tarde (después de comida)
];

// ------------------------- WEB APP -------------------------
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Tiempos de respuesta Surtido')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ------------------------- HELPERS -------------------------
function openSpreadsheet_() {
  var ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { ss = null; }
  if (!ss) ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  if (!ss) throw new Error('No se pudo abrir el spreadsheet.');
  return ss;
}

function getSheet_() {
  var ss = openSpreadsheet_();
  return ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
}

// Normaliza texto: minúsculas, sin acentos, sin espacios extra.
function normalize_(s) {
  return String(s == null ? '' : s)
    .toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function fmt_(d) {
  return d ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy') : '';
}

// ------------------------- CÁLCULO DE TIEMPO (sin cambios) -------------------------
/**
 * Minutos hábiles transcurridos entre dos fechas/hora, contando solo
 * Lunes-Viernes, 8:30-18:00, excluyendo la hora de comida 13:00-14:00.
 */
function businessMinutesBetween_(start, end) {
  if (!(start instanceof Date) || !(end instanceof Date)) return null;
  if (isNaN(start.getTime()) || isNaN(end.getTime())) return null;
  if (end <= start) return 0;

  var spanDays = (end.getTime() - start.getTime()) / 86400000;
  if (spanDays > MAX_SPAN_DAYS) return null; // evita bucles larguísimos por fechas corruptas

  var totalMin = 0;
  var cursor = new Date(start.getFullYear(), start.getMonth(), start.getDate());
  var lastDay = new Date(end.getFullYear(), end.getMonth(), end.getDate());
  var guard = 0;

  while (cursor <= lastDay) {
    guard++;
    if (guard > MAX_SPAN_DAYS + 5) return null; // cinturón de seguridad extra
    var dow = cursor.getDay(); // 0 domingo ... 6 sábado
    if (dow >= 1 && dow <= 5) {
      for (var s = 0; s < WORKDAY_SEGMENTS.length; s++) {
        var seg = WORKDAY_SEGMENTS[s];
        var segStart = new Date(cursor.getFullYear(), cursor.getMonth(), cursor.getDate(), seg.startH, seg.startM, 0);
        var segEnd = new Date(cursor.getFullYear(), cursor.getMonth(), cursor.getDate(), seg.endH, seg.endM, 0);
        var overlapStart = start > segStart ? start : segStart;
        var overlapEnd = end < segEnd ? end : segEnd;
        if (overlapEnd > overlapStart) {
          totalMin += (overlapEnd.getTime() - overlapStart.getTime()) / 60000;
        }
      }
    }
    cursor.setDate(cursor.getDate() + 1);
  }

  return Math.round(totalMin);
}

/**
 * Combina la fecha real (columna D) con una hora ya parseada ({h,m,s}).
 */
function combineDateTime_(fecha, timeParts) {
  return new Date(
    fecha.getFullYear(), fecha.getMonth(), fecha.getDate(),
    timeParts.h, timeParts.m, timeParts.s
  );
}

/**
 * Acepta una celda de hora tanto si Sheets la guardó como valor de hora real
 * (Date con fecha base 1899-12-30) como si quedó guardada en texto,
 * p.ej. "2:10:00 p.m." o "14:10".
 */
function parseTimeCell_(cell) {
  if (cell instanceof Date && !isNaN(cell.getTime())) {
    return { h: cell.getHours(), m: cell.getMinutes(), s: cell.getSeconds() };
  }
  if (typeof cell === 'string' && cell.trim() !== '') {
    var s = cell.trim().toLowerCase().replace(/\./g, '').replace(/\s+/g, ' ');
    var m = s.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(am|pm)?$/);
    if (m) {
      var h = parseInt(m[1], 10);
      var min = parseInt(m[2], 10);
      var sec = m[3] ? parseInt(m[3], 10) : 0;
      var ampm = m[4];
      if (ampm === 'pm' && h < 12) h += 12;
      if (ampm === 'am' && h === 12) h = 0;
      if (h >= 0 && h < 24 && min >= 0 && min < 60) {
        return { h: h, m: min, s: sec };
      }
    }
  }
  return null;
}

/**
 * Acepta la fecha tanto si Sheets la guardó como valor de fecha real
 * como si quedó en texto, p.ej. "23/3/2026".
 */
function parseFecha_(cell) {
  if (cell instanceof Date && !isNaN(cell.getTime())) return cell;
  if (typeof cell === 'string' && cell.trim() !== '') {
    var m = cell.trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (m) {
      var d = new Date(parseInt(m[3], 10), parseInt(m[2], 10) - 1, parseInt(m[1], 10));
      if (!isNaN(d.getTime())) return d;
    }
  }
  return null;
}

// ------------------------- LECTURA DE REGISTROS -------------------------
function buildRecords_() {
  var sheet = getSheet_();
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2) return { error: 'La hoja "' + SHEET_NAME + '" no tiene datos.' };

  var values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var recs = [];

  values.forEach(function (row) {
    try {
      var prefactura = row[0];  // A
      var cliente = row[2];     // C
      var fecha = row[3];       // D
      var horaRecibo = row[4];  // E
      var nave = row[13];       // N
      var horaSurtido = row[15];// P
      var estatusAlmacen = row[17]; // R
      var estatusEntrega = row[24]; // Y
      var semana = row[35];     // AJ

      // fila vacía o sin folio: se ignora
      if (!prefactura) return;

      var fechaDt = parseFecha_(fecha);
      if (!fechaDt) return; // fecha ausente o irreconocible
      var anio = fechaDt.getFullYear();
      if (anio < 2000 || anio > 2100) return; // fecha claramente corrupta

      var tiempoMin = null;
      var cumple = null;

      var tRecibo = parseTimeCell_(horaRecibo);
      var tSurtido = parseTimeCell_(horaSurtido);

      if (tRecibo && tSurtido) {
        var startDT = combineDateTime_(fechaDt, tRecibo);
        var endDT = combineDateTime_(fechaDt, tSurtido);
        // si la hora de surtido es "menor" que la de recibo, cruzó medianoche
        if (endDT <= startDT) endDT.setDate(endDT.getDate() + 1);

        tiempoMin = businessMinutesBetween_(startDT, endDT);
        cumple = tiempoMin !== null ? tiempoMin <= SLA_MINUTES : null;
      }

      var estatus = (estatusAlmacen || estatusEntrega) ? String(estatusAlmacen || estatusEntrega).trim() : '(Sin estatus)';

      recs.push({
        prefactura: String(prefactura),
        cliente: cliente ? String(cliente).trim() : '',
        fecha: fechaDt,
        anio: anio,
        mes: fechaDt.getMonth() + 1,
        nave: nave ? String(nave).trim() : '(Sin nave)',
        estatus: estatus,
        entregado: ENTREGADO_ESTATUS.indexOf(normalize_(estatus)) >= 0,
        semana: semana !== '' && semana !== null ? String(semana).trim() : '',
        tiempoMin: tiempoMin,
        cumple: cumple
      });
    } catch (rowErr) {
      // fila con datos inesperados: se omite y se sigue con las demás
    }
  });

  return { recs: recs };
}

// Ordena etiquetas de semana: numéricas primero (ascendente), luego texto.
function cmpSemana_(a, b) {
  var na = parseFloat(a), nb = parseFloat(b);
  var aNum = !isNaN(na), bNum = !isNaN(nb);
  if (aNum && bNum) return na - nb;
  if (aNum) return -1;
  if (bNum) return 1;
  return String(a).localeCompare(String(b));
}

// ------------------------- API PRINCIPAL -------------------------
function getDashboard(filters) {
  filters = filters || {};
  var fAnio    = String(filters.anio    || 'todos');
  var fMes     = String(filters.mes     || 'todos');
  var fSemana  = String(filters.semana  || 'todas');
  var fNave    = String(filters.nave    || 'todas');
  var fEstatus = String(filters.estatus || 'todos');
  var fBuscar  = normalize_(filters.buscar || '');

  var built = buildRecords_();
  if (built.error) return { error: built.error };
  var recs = built.recs;

  function passBase(r) {
    if (fAnio !== 'todos' && String(r.anio) !== fAnio) return false;
    if (fMes !== 'todos' && String(r.mes) !== fMes) return false;
    if (fNave !== 'todas' && r.nave !== fNave) return false;
    if (fEstatus !== 'todos' && r.estatus !== fEstatus) return false;
    if (fBuscar && normalize_(r.prefactura + ' ' + r.cliente).indexOf(fBuscar) === -1) return false;
    return true;
  }
  function passSemana(r) {
    return fSemana === 'todas' || r.semana === fSemana;
  }

  // ---- Opciones para los filtros (sobre todo el universo) ----
  var aniosSet = {}, navesSet = {}, estatusSet = {};
  recs.forEach(function (r) {
    aniosSet[r.anio] = true; navesSet[r.nave] = true; estatusSet[r.estatus] = true;
  });
  var anios = Object.keys(aniosSet).map(Number).sort(function (a, b) { return a - b; });
  var naves = Object.keys(navesSet).sort();
  var estatusList = Object.keys(estatusSet).sort();

  // ---- Registros que pasan Año/Mes/Nave/Estatus/Buscar (para el gráfico semanal) ----
  var forWeekly = recs.filter(passBase);

  // ---- Registros que pasan TODOS los filtros (para KPIs y tabla) ----
  var filtered = forWeekly.filter(passSemana);

  // ---- Semanas disponibles (según los demás filtros) ----
  var semSet = {};
  forWeekly.forEach(function (r) { if (r.semana !== '') semSet[r.semana] = true; });
  var semanas = Object.keys(semSet).sort(cmpSemana_);

  // ---- KPIs ----
  var k = { total: 0, entregadas: 0, enProceso: 0, conCalculo: 0, sinCalculo: 0, dentro: 0, fuera: 0 };
  var tiempos = [];
  filtered.forEach(function (r) {
    k.total++;
    if (!r.entregado) { k.enProceso++; return; }
    k.entregadas++;
    if (typeof r.tiempoMin !== 'number') { k.sinCalculo++; return; }
    k.conCalculo++;
    tiempos.push(r.tiempoMin);
    if (r.cumple === true) k.dentro++; else k.fuera++;
  });

  var prom = null, mediana = null, min = null, max = null;
  if (tiempos.length) {
    prom = Math.round(tiempos.reduce(function (a, b) { return a + b; }, 0) / tiempos.length);
    var sorted = tiempos.slice().sort(function (a, b) { return a - b; });
    mediana = sorted.length % 2
      ? sorted[(sorted.length - 1) / 2]
      : Math.round((sorted[sorted.length / 2 - 1] + sorted[sorted.length / 2]) / 2);
    min = sorted[0];
    max = sorted[sorted.length - 1];
  }

  var kpis = {
    total: k.total,
    entregadas: k.entregadas,
    enProceso: k.enProceso,
    conCalculo: k.conCalculo,
    sinCalculo: k.sinCalculo,
    dentro: k.dentro,
    fuera: k.fuera,
    prom: prom === null ? -1 : prom,
    mediana: mediana === null ? -1 : mediana,
    min: min === null ? -1 : min,
    max: max === null ? -1 : max,
    pctCumpl:    k.conCalculo ? Math.round(k.dentro / k.conCalculo * 100) : 0,
    pctFuera:    k.conCalculo ? Math.round(k.fuera / k.conCalculo * 100) : 0,
    pctSinCalc:  k.total ? Math.round(k.sinCalculo / k.total * 100) : 0,
    pctAvance:   k.total ? Math.round(k.entregadas / k.total * 100) : 0
  };

  // ---- Serie semanal (solo entregadas con cálculo) ----
  var groups = {};
  forWeekly.forEach(function (r) {
    if (!r.entregado || typeof r.tiempoMin !== 'number') return;
    var key = r.semana !== '' ? ('S' + r.semana) : fmt_(r.fecha);
    if (!groups[key]) groups[key] = { sem: r.semana, count: 0, sum: 0, dentro: 0, fuera: 0, fecha: r.fecha };
    var g = groups[key];
    g.count++;
    g.sum += r.tiempoMin;
    if (r.cumple === true) g.dentro++; else g.fuera++;
  });
  var weekly = Object.keys(groups).sort(function (a, b) {
    var ga = groups[a], gb = groups[b];
    if (ga.sem !== '' && gb.sem !== '') return cmpSemana_(ga.sem, gb.sem);
    if (ga.sem !== '') return -1;
    if (gb.sem !== '') return 1;
    return ga.fecha - gb.fecha;
  }).map(function (key) {
    var g = groups[key];
    return {
      label: key,
      cant: g.count,
      dentro: g.dentro,
      fuera: g.fuera,
      prom: Math.round(g.sum / g.count),
      pctCumpl: g.count ? Math.round(g.dentro / g.count * 100) : 0
    };
  });

  // ---- Tabla: solicitudes entregadas fuera de meta ----
  var fueraRows = filtered.filter(function (r) { return r.entregado && r.cumple === false; })
    .sort(function (a, b) { return b.tiempoMin - a.tiempoMin; });
  var atrasos = fueraRows.slice(0, 300).map(function (r) {
    return {
      no: r.prefactura, cte: r.cliente, fecha: fmt_(r.fecha), nave: r.nave,
      estatus: r.estatus, semana: r.semana, min: r.tiempoMin,
      exceso: r.tiempoMin - SLA_MINUTES
    };
  });

  // ---- Detalle para exportar CSV ----
  var detalle = filtered.map(function (r) {
    var res = !r.entregado ? 'En proceso' :
              typeof r.tiempoMin !== 'number' ? 'Sin cálculo' :
              r.cumple === true ? 'Dentro de meta' : 'Fuera de meta';
    return {
      no: r.prefactura, cte: r.cliente, fecha: fmt_(r.fecha), nave: r.nave,
      estatus: r.estatus, semana: r.semana,
      min: typeof r.tiempoMin === 'number' ? r.tiempoMin : '',
      resultado: res
    };
  });

  return {
    meta: {
      title: META_TITLE,
      subtitle: META_SUBTITLE,
      slaHours: SLA_HOURS,
      slaMinutes: SLA_MINUTES,
      updated: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yy, HH:mm')
    },
    options: { anios: anios, semanas: semanas, naves: naves, estatus: estatusList },
    kpis: kpis,
    weekly: weekly,
    atrasos: atrasos,
    atrasosTotal: fueraRows.length,
    detalle: detalle
  };
}
