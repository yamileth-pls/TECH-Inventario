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
 *    - Meta: surtir en 6 horas laborales (360 min).
 * ==========================================================================
 */

// ------------------------- CONFIGURACIÓN -------------------------
var SPREADSHEET_ID = '1we4yhPySbL-wODeZcrNMYvMZ0HmrHmS2sM1oLZsrQoM';
var SHEET_NAME     = 'Inventario';
var META_TITLE     = 'Tiempos de respuesta Surtido';
var META_SUBTITLE  = 'Consulta de Inventarios · Área de Almacenes';
var SLA_HOURS      = 6;                 // meta: horas laborales para surtir
var SLA_MINUTES    = SLA_HOURS * 60;    // 360
var MAX_SPAN_DAYS  = 60;                // más allá de esto se considera dato corrupto, no un pedido real

// Estatus (normalizados, sin acentos ni mayúsculas) que cuentan como "entregado"
var ENTREGADO_ESTATUS = ['validacion', 'entregado'];

// Estatus (normalizados) que se omiten por completo: no cuentan en ningún indicador ni filtro
var EXCLUIR_ESTATUS = ['cancelado', 'cancelada', 'pausado', 'pausada'];

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

// Piezas (columna PZS) como entero >= 0. Vacío o ilegible = 0 (no se pondera).
function parsePzs_(v) {
  var n = typeof v === 'number' ? v : parseFloat(String(v == null ? '' : v).replace(/,/g, ''));
  if (isNaN(n) || n < 0) return 0;
  return Math.round(n);
}

// dd/MM/yyyy sin llamar a servicios de Apps Script (se usa miles de veces por carga).
function fmt_(d) {
  if (!d) return '';
  return ('0' + d.getDate()).slice(-2) + '/' + ('0' + (d.getMonth() + 1)).slice(-2) + '/' + d.getFullYear();
}

// dd/MM/yyyy HH:mm (24 h)
function fmtDT_(d) {
  return fmt_(d) + ' ' + ('0' + d.getHours()).slice(-2) + ':' + ('0' + d.getMinutes()).slice(-2);
}

// Número de semana ISO de una fecha (Lunes = inicio de semana).
function isoWeek_(d) {
  var date = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  var dayNum = (date.getUTCDay() + 6) % 7;            // Lunes = 0
  date.setUTCDate(date.getUTCDate() - dayNum + 3);    // jueves de esa semana
  var firstThu = new Date(Date.UTC(date.getUTCFullYear(), 0, 4));
  var firstDayNum = (firstThu.getUTCDay() + 6) % 7;
  firstThu.setUTCDate(firstThu.getUTCDate() - firstDayNum + 3);
  return 1 + Math.round((date - firstThu) / (7 * 24 * 3600 * 1000));
}

// Índice de columna por nombre de encabezado exacto (sin acentos/mayúsculas). -1 si no existe.
function findCol_(normHeaders, aliases) {
  for (var a = 0; a < aliases.length; a++) {
    var idx = normHeaders.indexOf(normalize_(aliases[a]));
    if (idx >= 0) return idx;
  }
  return -1;
}

// ------------------------- CÁLCULO DE TIEMPO (sin cambios) -------------------------
/**
 * Minutos laborales transcurridos entre dos fechas/hora, contando solo
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

// ------------------------- TIEMPO SIN SUPONER FECHAS -------------------------
// Texto de una celda para mostrarlo en la alerta de datos.
function valTxt_(v) {
  if (v === '' || v == null) return '(vacía)';
  if (v instanceof Date) return isNaN(v.getTime()) ? '(fecha inválida)' : fmt_(v);
  return String(v).trim();
}

// Calcula el tiempo laboral de una línea SIN suponer ninguna fecha ni hora.
// Solo se calcula si FECHA, HORA RECIBO, FECHA DE SURTIDO y HORA DE SURTIDO están completas y son
// coherentes. Si algo falta o está mal, no se calcula y se devuelve el "problema" con el valor
// capturado, para avisar qué se debe corregir en la hoja.
function evalTiempo_(fechaDt, rawHRec, rawHSur, rawFSur) {
  var r = { tiempoMin: null, cumple: null, startDT: null, endDT: null, surtidoTxt: '', problema: null };
  var tR = parseTimeCell_(rawHRec), tS = parseTimeCell_(rawHSur);
  if (tR) r.startDT = combineDateTime_(fechaDt, tR);
  var horaSur = tS ? (' ' + ('0' + tS.h).slice(-2) + ':' + ('0' + tS.m).slice(-2)) : '';
  function prob(codigo, campo, valor, texto) { r.problema = { codigo: codigo, campo: campo, valor: valor, texto: texto }; return r; }

  // 1) horas
  if (!tR || !tS) {
    var campos = [], vals = [];
    if (!tR) { campos.push('HORA RECIBO'); vals.push(valTxt_(rawHRec)); }
    if (!tS) { campos.push('HORA DE SURTIDO'); vals.push(valTxt_(rawHSur)); }
    r.surtidoTxt = tS ? valTxt_(rawFSur) + horaSur : '';
    return prob('hora', campos.join(' y '), vals.join(' / '), 'Falta capturar la hora (o no es una hora válida)');
  }

  // 2) FECHA DE SURTIDO
  var fSur = null;
  if (rawFSur === '' || rawFSur == null) {
    r.surtidoTxt = '(vacía)' + horaSur;
    return prob('fecha_vacia', 'FECHA DE SURTIDO', '(vacía)', 'Falta capturar la FECHA DE SURTIDO');
  }
  if (rawFSur instanceof Date) { if (!isNaN(rawFSur.getTime())) fSur = rawFSur; }
  else fSur = parseFecha_(rawFSur);                 // texto con formato d/m/aaaa
  if (!fSur) {
    r.surtidoTxt = valTxt_(rawFSur) + horaSur;
    return prob('fecha_invalida', 'FECHA DE SURTIDO', valTxt_(rawFSur), 'La FECHA DE SURTIDO no es una fecha válida');
  }
  var y = fSur.getFullYear();
  r.endDT = combineDateTime_(fSur, tS);
  r.surtidoTxt = fmtDT_(r.endDT);
  if (y < 2000 || y > 2100) {
    return prob('fecha_anio', 'FECHA DE SURTIDO', valTxt_(fSur), 'La FECHA DE SURTIDO tiene un año fuera de rango (' + y + ')');
  }

  // 3) coherencia entre recibo y surtido
  if (r.endDT < r.startDT) {
    return prob('surtido_antes', 'FECHA / HORA DE SURTIDO', 'surtido ' + fmtDT_(r.endDT) + ' · recibo ' + fmtDT_(r.startDT),
                'El surtido es anterior al recibo');
  }
  if ((r.endDT.getTime() - r.startDT.getTime()) / 86400000 > MAX_SPAN_DAYS) {
    return prob('rango', 'FECHA DE SURTIDO', valTxt_(fSur), 'Pasan más de ' + MAX_SPAN_DAYS + ' días entre el recibo y el surtido');
  }

  r.tiempoMin = businessMinutesBetween_(r.startDT, r.endDT);
  r.cumple = r.tiempoMin !== null ? r.tiempoMin <= SLA_MINUTES : null;
  return r;
}

// ------------------------- LECTURA DE REGISTROS -------------------------
// Las columnas se ubican por NOMBRE de encabezado (no por posición), así que
// insertar/mover columnas en la hoja no rompe el tablero.
function buildRecords_() {
  var t0 = Date.now();
  var sheet = getSheet_();
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2) return { error: 'La hoja "' + SHEET_NAME + '" no tiene datos.' };

  var values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var readMs = Date.now() - t0;

  var headers = values[0];
  var nh = headers.map(normalize_);
  var cPref  = findCol_(nh, ['prefactura', 'columna 1']);
  if (cPref < 0) cPref = 0;                                  // la prefactura siempre es la columna A
  var cCli   = findCol_(nh, ['cliente']);
  var cFec   = findCol_(nh, ['fecha']);
  var cHRec  = findCol_(nh, ['hora recibo']);
  var cNave  = findCol_(nh, ['nave']);
  var cFSur  = findCol_(nh, ['fecha de surtido']);           // opcional
  var cHSur  = findCol_(nh, ['hora de surtido']);
  var cEstA  = findCol_(nh, ['estatus almacen']);
  var cEstE  = findCol_(nh, ['estatus de entrega']);
  var cEstR  = findCol_(nh, ['estatus real']);               // opcional: solo para omitir cancelados/pausados
  var cSem   = findCol_(nh, ['semana']);                     // opcional: si no existe se calcula de la fecha
  var cDesc  = findCol_(nh, ['descripcion']);                // opcional: solo para el desglose por prefactura
  var cPzs   = findCol_(nh, ['pzs']);                        // opcional

  var missing = [];
  if (cFec < 0)  missing.push('Fecha');
  if (cHRec < 0) missing.push('HORA RECIBO');
  if (cHSur < 0) missing.push('HORA DE SURTIDO');
  if (cEstA < 0 && cEstE < 0) missing.push('Estatus Almacen');
  if (missing.length) {
    return { error: 'No se encontraron las columnas: ' + missing.join(', ') +
             '.\nEncabezados detectados: ' + headers.join(' | ') };
  }

  var recs = [];
  var alertas = [];      // líneas ENTREGADAS cuya FECHA (recibo) es inválida: no entran a recs
  for (var i = 1; i < values.length; i++) {
    try {
      var row = values[i];
      var prefactura = row[cPref];

      // fila vacía o sin folio: se ignora
      if (!prefactura) continue;

      var estA = cEstA >= 0 ? row[cEstA] : '';
      var estE = cEstE >= 0 ? row[cEstE] : '';
      var estatus = (estA || estE) ? String(estA || estE).trim() : '(Sin estatus)';

      // pausadas / canceladas: se omiten (no afectan ningún indicador)
      if (EXCLUIR_ESTATUS.indexOf(normalize_(estatus)) >= 0) continue;
      if (cEstR >= 0 && EXCLUIR_ESTATUS.indexOf(normalize_(row[cEstR])) >= 0) continue;

      var entregado = ENTREGADO_ESTATUS.indexOf(normalize_(estatus)) >= 0;

      // FECHA (recibo): sin ella la línea no se puede ubicar en ningún año / mes / semana
      var fechaDt = parseFecha_(row[cFec]);
      var anio = fechaDt ? fechaDt.getFullYear() : 0;
      if (!fechaDt || anio < 2000 || anio > 2100) {
        if (entregado) {
          alertas.push({ no: String(prefactura), fila: i + 1, cte: (cCli >= 0 && row[cCli]) ? String(row[cCli]).trim() : '',
            desc: (cDesc >= 0 && row[cDesc]) ? String(row[cDesc]).trim() : '', campo: 'FECHA', valor: valTxt_(row[cFec]),
            texto: 'La FECHA (de recibo) no es válida: la línea no se puede ubicar y queda fuera del tablero', recibo: '' });
        }
        continue;
      }

      var ev = evalTiempo_(fechaDt, row[cHRec], row[cHSur], cFSur >= 0 ? row[cFSur] : '');
      var tiempoMin = ev.tiempoMin, cumple = ev.cumple;

      var nave = cNave >= 0 ? row[cNave] : '';
      var semCell = cSem >= 0 ? row[cSem] : '';
      var semana = (semCell !== '' && semCell != null && !(semCell instanceof Date))
        ? String(semCell).trim()
        : String(isoWeek_(fechaDt));

      recs.push({
        prefactura: String(prefactura),
        cliente: (cCli >= 0 && row[cCli]) ? String(row[cCli]).trim() : '',
        descripcion: (cDesc >= 0 && row[cDesc]) ? String(row[cDesc]).trim() : '',
        pzs: (cPzs >= 0 && row[cPzs] !== '' && row[cPzs] != null) ? String(row[cPzs]).trim() : '',
        pzsN: cPzs >= 0 ? parsePzs_(row[cPzs]) : 0,
        fecha: fechaDt,
        anio: anio,
        mes: fechaDt.getMonth() + 1,
        nave: nave ? String(nave).trim() : '(Sin nave)',
        estatus: estatus,
        entregado: entregado,
        semana: semana,
        tiempoMin: tiempoMin,
        cumple: cumple,
        fila: i + 1,
        problema: ev.problema,
        cat: lineCat_(entregado, tiempoMin, cumple),
        iniDt: ev.startDT,
        finDt: ev.problema ? null : ev.endDT,         // una fecha con problema no cuenta como "último surtido" de la prefactura
        recibo: ev.startDT ? fmtDT_(ev.startDT) : '',
        surtido: ev.surtidoTxt
      });
    } catch (rowErr) {
      // fila con datos inesperados: se omite y se sigue con las demás
    }
  }

  DIAG_.readMs = readMs;
  DIAG_.rows = recs.length;
  return { recs: recs, alertas: alertas };
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

// ------------------------- CLASIFICACIÓN -------------------------
// Categoría de una línea (renglón de la hoja):
//   enProceso  = aún sin VALIDACION / ENTREGADO
//   sinCalculo = entregada pero sin hora de recibo o de surtido
//   dentro / fuera = entregada con tiempo laboral calculado (<= / > meta)
function lineCat_(entregado, tiempoMin, cumple) {
  if (!entregado) return 'enProceso';
  if (typeof tiempoMin !== 'number') return 'sinCalculo';
  return cumple === true ? 'dentro' : 'fuera';
}

// Agrupa las líneas por prefactura (módulo "Nivel de Servicio").
// Una prefactura CUMPLE solo si TODAS sus líneas se surtieron en <= meta. Reglas, en orden:
//   1) alguna línea fuera de meta            -> fuera   (ya no puede cumplir)
//   2) alguna línea aún sin entregar         -> enProceso
//   3) alguna línea entregada sin cálculo    -> sinCalculo
//   4) todas entregadas y dentro de meta     -> dentro
// Su tiempo laboral es el de su línea más tardada.
function buildPrefacturas_(recs) {
  var map = {}, order = [];
  recs.forEach(function (r) {
    if (!map[r.prefactura]) { map[r.prefactura] = []; order.push(r.prefactura); }
    map[r.prefactura].push(r);
  });

  return order.map(function (no) {
    var ls = map[no];
    var first = ls[0], cte = '', maxMin = null, ini = null, fin = null, allDel = true;
    var cnt = { dentro: 0, fuera: 0, enProceso: 0, sinCalculo: 0 };
    var naves = {}, estatuses = {};

    ls.forEach(function (r) {
      cnt[r.cat]++;
      if (!r.entregado) allDel = false;
      if (typeof r.tiempoMin === 'number' && (maxMin === null || r.tiempoMin > maxMin)) maxMin = r.tiempoMin;
      if (r.iniDt && (!ini || r.iniDt < ini)) ini = r.iniDt;
      if (r.finDt && (!fin || r.finDt > fin)) fin = r.finDt;
      if (!cte && r.cliente) cte = r.cliente;
      naves[r.nave] = true;
      estatuses[r.estatus] = true;
      if (r.fecha < first.fecha) first = r;
    });

    var cat = cnt.fuera ? 'fuera' : cnt.enProceso ? 'enProceso' : cnt.sinCalculo ? 'sinCalculo' : 'dentro';
    return {
      prefactura: no, cliente: cte,
      fecha: first.fecha, anio: first.anio, mes: first.mes, semana: first.semana,
      nave: Object.keys(naves), estatus: Object.keys(estatuses),
      entregado: allDel, cat: cat, tiempoMin: maxMin,
      cumple: cat === 'dentro' ? true : (cat === 'fuera' ? false : null),
      recibo: ini ? fmtDT_(ini) : '', surtido: fin ? fmtDT_(fin) : '',
      nLineas: ls.length, nDentro: cnt.dentro, nFuera: cnt.fuera, nProceso: cnt.enProceso, nSinCalc: cnt.sinCalculo,
      lines: ls
    };
  });
}

// ------------------------- FILTROS -------------------------
// Aplica los filtros del tablero a "unidades" (líneas o prefacturas).
// forWeekly = todo menos Semana (para la gráfica y las opciones de semana);
// filtered = todos los filtros (para KPIs, tabla y CSV).
function filterUnits_(units, filters) {
  filters = filters || {};
  var fAnio    = String(filters.anio    || 'todos');
  var fMes     = String(filters.mes     || 'todos');
  var fSemana  = String(filters.semana  || 'todas');
  var fNave    = String(filters.nave    || 'todas');
  var fEstatus = String(filters.estatus || 'todos');
  var fBuscar  = normalize_(filters.buscar || '');

  // nave / estatus: texto en una línea, lista en una prefactura ("alguna de sus líneas")
  function has(v, x) { return [].concat(v).indexOf(x) >= 0; }

  var forWeekly = units.filter(function (r) {
    if (fAnio !== 'todos' && String(r.anio) !== fAnio) return false;
    if (fMes !== 'todos' && String(r.mes) !== fMes) return false;
    if (fNave !== 'todas' && !has(r.nave, fNave)) return false;
    if (fEstatus !== 'todos' && !has(r.estatus, fEstatus)) return false;
    if (fBuscar && normalize_(r.prefactura + ' ' + r.cliente).indexOf(fBuscar) === -1) return false;
    return true;
  });
  var filtered = fSemana === 'todas' ? forWeekly : forWeekly.filter(function (r) { return r.semana === fSemana; });
  return { forWeekly: forWeekly, filtered: filtered };
}

// Módulos: 'surtido' = por línea · 'servicio' = por prefactura · 'piezas' = por pieza (cada línea pesa su PZS)
function normModulo_(m) {
  return (m === 'servicio' || m === 'piezas') ? m : 'surtido';
}

// Unidades del módulo. En 'piezas' se descartan las líneas sin PZS (no tienen peso).
function unitsFor_(recs, modulo) {
  if (modulo === 'servicio') return buildPrefacturas_(recs);
  if (modulo === 'piezas') return recs.filter(function (r) { return r.pzsN > 0; });
  return recs;
}

// Peso de una unidad en los conteos: 1 (línea / prefactura) o sus piezas.
function weightOf_(u, modulo) {
  return modulo === 'piezas' ? u.pzsN : 1;
}

// Mediana ponderada de [{v, w}] (w entero). Con todos los pesos = 1 equivale a la mediana normal.
function weightedMedian_(items) {
  var arr = items.slice().sort(function (a, b) { return a.v - b.v; });
  var n = arr.reduce(function (t, i) { return t + i.w; }, 0);
  if (!n) return null;
  function valueAtRank(rank) {            // rank 1-based
    var cum = 0;
    for (var i = 0; i < arr.length; i++) { cum += arr[i].w; if (cum >= rank) return arr[i].v; }
    return arr[arr.length - 1].v;
  }
  return n % 2 ? valueAtRank((n + 1) / 2) : Math.round((valueAtRank(n / 2) + valueAtRank(n / 2 + 1)) / 2);
}

// ------------------------- API PRINCIPAL -------------------------
// Diagnóstico de la última llamada (se muestra en el encabezado para ver dónde se va el tiempo).
var DIAG_ = {};

function getDashboard(filters, modulo) {
  var tStart = Date.now();
  modulo = normModulo_(modulo);

  var built = buildRecords_();
  if (built.error) return { error: built.error };
  var recs = built.recs;
  var units = unitsFor_(recs, modulo);
  var sinPzs = modulo === 'piezas' ? recs.length - units.length : 0;

  // ---- Opciones para los filtros (sobre todas las líneas) ----
  var aniosSet = {}, navesSet = {}, estatusSet = {};
  recs.forEach(function (r) {
    aniosSet[r.anio] = true; navesSet[r.nave] = true; estatusSet[r.estatus] = true;
  });
  var anios = Object.keys(aniosSet).map(Number).sort(function (a, b) { return a - b; });
  var naves = Object.keys(navesSet).sort();
  var estatusList = Object.keys(estatusSet).sort();

  // ---- Unidades filtradas ----
  var f = filterUnits_(units, filters);
  var forWeekly = f.forWeekly;
  var filtered = f.filtered;

  // ---- Semanas disponibles (según los demás filtros) ----
  var semSet = {};
  forWeekly.forEach(function (r) { if (r.semana !== '') semSet[r.semana] = true; });
  var semanas = Object.keys(semSet).sort(cmpSemana_);

  // ---- KPIs ----
  // Los conteos se ponderan por el peso de la unidad (1, o las piezas en el módulo 'piezas').
  var k = { total: 0, entregadas: 0, enProceso: 0, sinCalculo: 0, dentro: 0, fuera: 0 };
  var tiempos = [];
  filtered.forEach(function (r) {
    var w = weightOf_(r, modulo);
    k.total += w;
    k[r.cat] += w;
    if (r.entregado) k.entregadas += w;
    if (r.cat === 'dentro' || r.cat === 'fuera') tiempos.push({ v: r.tiempoMin, w: w });
  });
  k.conCalculo = k.dentro + k.fuera;

  var prom = null, mediana = null, min = null, max = null;
  if (tiempos.length) {
    var sumW = tiempos.reduce(function (t, i) { return t + i.w; }, 0);
    prom = Math.round(tiempos.reduce(function (t, i) { return t + i.v * i.w; }, 0) / sumW);
    mediana = weightedMedian_(tiempos);
    var vals = tiempos.map(function (i) { return i.v; });
    min = Math.min.apply(null, vals);
    max = Math.max.apply(null, vals);
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

  // ---- Serie semanal (solo unidades con cálculo) ----
  var groups = {};
  forWeekly.forEach(function (r) {
    if (r.cat !== 'dentro' && r.cat !== 'fuera') return;
    var key = r.semana !== '' ? ('S' + r.semana) : fmt_(r.fecha);
    if (!groups[key]) groups[key] = { sem: r.semana, count: 0, sum: 0, dentro: 0, fuera: 0, fecha: r.fecha };
    var g = groups[key], w = weightOf_(r, modulo);
    g.count += w;
    g.sum += r.tiempoMin * w;
    if (r.cat === 'dentro') g.dentro += w; else g.fuera += w;
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

  // ---- Tabla: unidades fuera de meta ----
  var fueraRows = filtered.filter(function (r) { return r.cat === 'fuera'; })
    .sort(function (a, b) { return b.tiempoMin - a.tiempoMin; });
  var atrasos = fueraRows.slice(0, 300).map(function (r) {
    if (modulo === 'servicio') {
      return {
        no: r.prefactura, cte: r.cliente, recibo: r.recibo, surtido: r.surtido,
        semana: r.semana, lineas: r.nLineas, lineasFuera: r.nFuera,
        min: r.tiempoMin, exceso: r.tiempoMin - SLA_MINUTES
      };
    }
    return {
      no: r.prefactura, cte: r.cliente, fecha: fmt_(r.fecha), nave: r.nave,
      estatus: r.estatus, semana: r.semana, min: r.tiempoMin,
      desc: r.descripcion, pzs: r.pzsN,
      recibo: r.recibo, surtido: r.surtido,
      exceso: r.tiempoMin - SLA_MINUTES
    };
  });

  // Alerta de fechas / horas por corregir: líneas entregadas que NO se pudieron calcular por un dato faltante o inválido.
  // Respeta los filtros; las de FECHA (recibo) inválida no se pueden ubicar, así que siempre se listan.
  var alertItems = filterUnits_(recs, filters).filtered
    .filter(function (r) { return r.entregado && r.problema; })
    .map(function (r) {
      return { no: r.prefactura, fila: r.fila, cte: r.cliente, desc: r.descripcion,
               campo: r.problema.campo, valor: r.problema.valor, texto: r.problema.texto, recibo: r.recibo };
    })
    .concat(built.alertas);
  alertItems.sort(function (a, b) {
    return String(a.no).localeCompare(String(b.no), 'es', { numeric: true }) || a.fila - b.fila;
  });
  var alertPf = {};
  alertItems.forEach(function (a) { alertPf[a.no] = true; });

  // Si hay texto en "Buscar", se devuelve la lista de coincidencias (todas, sin importar su categoría)
  var buscarTxt = String((filters && filters.buscar) || '').trim();
  var resultados = buscarTxt ? detalleRows_(filtered, 'total', modulo) : [];
  var resultadosTotal = resultados.length;

  return {
    meta: {
      buscar: buscarTxt,
      alertas: { n: alertItems.length, prefacturas: Object.keys(alertPf).length, items: alertItems.slice(0, 1000) },
      title: META_TITLE,
      subtitle: META_SUBTITLE,
      modulo: modulo,
      sinPzs: sinPzs,
      slaHours: SLA_HOURS,
      slaMinutes: SLA_MINUTES,
      updated: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yy, HH:mm'),
      diag: { rows: DIAG_.rows, readMs: DIAG_.readMs, totalMs: Date.now() - tStart }
    },
    options: { anios: anios, semanas: semanas, naves: naves, estatus: estatusList },
    kpis: kpis,
    weekly: weekly,
    atrasos: atrasos,
    resultados: resultados.slice(0, 500),
    resultadosTotal: resultadosTotal,
    atrasosTotal: fueraRows.length,
    atrasosPzs: modulo === 'piezas' ? fueraRows.reduce(function (t, r) { return t + r.pzsN; }, 0) : 0
  };
}

// ¿La unidad pertenece a la tarjeta (categoría) seleccionada?
function inCategory_(u, cat) {
  switch (cat) {
    case 'entregadas': return u.entregado;
    case 'dentro': case 'fuera': case 'enProceso': case 'sinCalculo': return u.cat === cat;
    default: return true; // 'total'
  }
}

var CAT_TEXT_ = { dentro: 'Dentro de meta', fuera: 'Fuera de meta', enProceso: 'En proceso', sinCalculo: 'Sin cálculo' };

// Filas de desglose (para una tarjeta, el CSV o los resultados de la búsqueda).
//  - 'surtido' / 'piezas': una fila por línea
//  - 'servicio': una fila por prefactura, con sus líneas anidadas en "lineas"
function detalleRows_(units, cat, modulo) {
  var rows = units.filter(function (u) { return inCategory_(u, cat); });
  if (cat === 'fuera') rows = rows.slice().sort(function (a, b) { return b.tiempoMin - a.tiempoMin; });

  function lineRow(r) {
    return {
      no: r.prefactura, cte: r.cliente, fecha: fmt_(r.fecha), nave: r.nave,
      estatus: r.estatus, semana: r.semana,
      desc: r.descripcion, pzs: r.pzs,
      recibo: r.recibo, surtido: r.surtido,
      min: typeof r.tiempoMin === 'number' ? r.tiempoMin : '',
      resultado: CAT_TEXT_[r.cat],
      motivo: (r.problema && r.entregado) ? (r.problema.campo + ': ' + r.problema.valor) : ''
    };
  }

  if (modulo === 'servicio') {
    return rows.map(function (p) {
      return {
        no: p.prefactura, cte: p.cliente, recibo: p.recibo, surtido: p.surtido, semana: p.semana,
        nLineas: p.nLineas, nDentro: p.nDentro, nFuera: p.nFuera, nProceso: p.nProceso, nSinCalc: p.nSinCalc,
        min: typeof p.tiempoMin === 'number' ? p.tiempoMin : '',
        resultado: CAT_TEXT_[p.cat],
        lineas: p.lines.map(lineRow)
      };
    });
  }
  return rows.map(lineRow);
}

// Desglose de las unidades filtradas que pertenecen a una tarjeta.
// Se usa al hacer clic en una tarjeta y al exportar CSV.
function getDetalle(filters, cat, modulo) {
  modulo = normModulo_(modulo);
  var built = buildRecords_();
  if (built.error) throw new Error(built.error);
  return detalleRows_(filterUnits_(unitsFor_(built.recs, modulo), filters).filtered, cat, modulo);
}
