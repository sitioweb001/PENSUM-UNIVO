// ============================================================
// UGB Pénsum — Apps Script v8 DEFINITIVO
// JSONP idéntico al sistema INMU que sí funciona
// ============================================================

// ── JSONP helper — igual que en el sistema que funciona ──
// Si viene ?callback=nombre → responde nombre(json);
// Si no → responde JSON normal
function _jsonResp(obj, e) {
  var callback = (e && e.parameter && e.parameter.callback) ? e.parameter.callback : null;
  var json = JSON.stringify(obj);
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// doGet
// ============================================================
function doGet(e) {
  try {
    var p       = (e && e.parameter) ? e.parameter : {};
    var action  = p.action  || '';
    var student = p.student || '';

    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // ── Lista de estudiantes ──
    if (action === 'list_students') {
      var career = p.career || '';
      var students = _getStudentList(ss, career);
      return _jsonResp({ status: 'success', students: students, count: students.length }, e);
    }

    // ── Hash de contraseña (para verificar login) ──
    if (action === 'get_hash') {
      var sName = p.student || '';
      if (!sName) return _jsonResp({ status: 'error', message: 'Falta nombre' }, e);
      var hash = _getHash(ss, sName);
      return _jsonResp({ status: 'success', hash: hash }, e);
    }

    // ── Datos completos de un estudiante ──
    if (student) {
      var result = _loadStudent(ss, student);
      return _jsonResp({ status: 'success', payload: result }, e);
    }

    // ── Sin parámetros → HTML ──
    try {
      return HtmlService.createHtmlOutputFromFile('INDEX_FINAL')
        .setTitle('UNIVO — Pénsum Ing. Civil')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    } catch(err2) {
      return _jsonResp({ status: 'error', message: 'INDEX_FINAL.html no encontrado' }, e);
    }

  } catch(err) {
    return _jsonResp({ status: 'error', message: err.toString() }, e);
  }
}

// ============================================================
// doPost
// ============================================================
function doPost(e) {
  try {
    var payload = {};
    try { payload = JSON.parse(e.postData.contents); } catch(_) {}

    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // ── Guardar hash de contraseña ──
    if (payload.action === 'save_hash') {
      var sName = payload.student || '';
      var hash  = payload.hash    || '';
      if (!sName || !hash) return _jsonResp({ status: 'error', message: 'Datos incompletos' }, null);
      _saveHashInSheet(ss, sName, hash);
      return _jsonResp({ status: 'success' }, null);
    }

    // ── Eliminar estudiante ──
    if (payload.action === 'delete_student') {
      var name = payload.student || '';
      if (!name) return _jsonResp({ status: 'error', message: 'Nombre vacío' }, null);
      _deleteStudent(ss, name);
      return _jsonResp({ status: 'success', message: name + ' eliminado' }, null);
    }

    // ── Guardar datos del estudiante ──
    var student     = payload.student     || 'Sin nombre';
    var data        = payload.data        || [];
    var events      = payload.events      || [];
    var cyclesDone  = payload.cyclesDone  || {};
    var asistencias = payload.asistencias || {};

    _registrarEstudiante(ss, student);
    _saveNotas(ss, student, data);
    _saveCalendario(ss, student, events);
    _saveCiclosDone(ss, student, cyclesDone);
    _saveAsistencias(ss, student, asistencias);
    _resumen(ss, student, data, cyclesDone, asistencias);

    return _jsonResp({ status: 'success', ts: new Date().toISOString() }, null);

  } catch(err) {
    return _jsonResp({ status: 'error', message: err.toString() }, null);
  }
}

// ============================================================
// CONTRASEÑAS — solo hashes SHA-256, nunca texto plano
// ============================================================
function _getHash(ss, studentName) {
  var sh = ss.getSheetByName('Passwords');
  if (!sh || sh.getLastRow() < 2) return null;
  var rows = sh.getRange(2, 1, sh.getLastRow()-1, 2).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (rows[i][0] === studentName) return rows[i][1] || null;
  }
  return null;
}

function _saveHashInSheet(ss, studentName, hash) {
  var sh = _getOrCreate(ss, 'Passwords');
  if (sh.getLastRow() === 0) {
    sh.appendRow(['Estudiante', 'Hash SHA-256 (no es la contraseña)']);
    _hdr(sh, 2, '#374151');
    try { sh.protect().setWarningOnly(true); } catch(_) {}
  }
  if (sh.getLastRow() > 1) {
    var rows = sh.getRange(2, 1, sh.getLastRow()-1, 1).getValues();
    for (var i = 0; i < rows.length; i++) {
      if (rows[i][0] === studentName) {
        sh.getRange(i+2, 2).setValue(hash);
        return;
      }
    }
  }
  sh.appendRow([studentName, hash]);
}

function _deleteHash(ss, studentName) {
  var sh = ss.getSheetByName('Passwords');
  if (!sh || sh.getLastRow() < 2) return;
  var rows = sh.getRange(2, 1, sh.getLastRow()-1, 1).getValues();
  for (var i = rows.length-1; i >= 0; i--) {
    if (rows[i][0] === studentName) sh.deleteRow(i+2);
  }
}

// ============================================================
// LECTURA
// ============================================================
function _getStudentList(ss, career) {
  var students = [];
  var suffix = career ? ('__' + career) : '';

  var shE = ss.getSheetByName('Estudiantes');
  if (shE && shE.getLastRow() > 1) {
    shE.getRange(2, 1, shE.getLastRow()-1, 1).getValues().forEach(function(r) {
      var raw = r[0];
      if (!raw) return;
      // Si se pidió una carrera específica, solo incluir nombres con ese sufijo
      if (suffix) {
        if (raw.indexOf(suffix) !== raw.length - suffix.length) return;
        raw = raw.slice(0, raw.length - suffix.length);
      }
      if (raw && students.indexOf(raw) < 0) students.push(raw);
    });
  }
  // Fallback: hojas Notas_*
  ss.getSheets().forEach(function(sh) {
    var n = sh.getName();
    if (n.indexOf('Notas_') === 0) {
      var sName = n.replace('Notas_', '');
      if (suffix) {
        if (sName.indexOf(suffix) !== sName.length - suffix.length) return;
        sName = sName.slice(0, sName.length - suffix.length);
      }
      if (sName && students.indexOf(sName) < 0) {
        students.push(sName);
        _registrarEstudiante(ss, sName + suffix);
      }
    }
  });
  return students;
}

function _loadStudent(ss, student) {
  var r = { data: [], events: [], cyclesDone: {}, asistencias: {} };

  // Buscar hoja con el nombre exacto; si no existe, intentar sin sufijo de carrera
  // (compatibilidad con datos creados antes de implementar el selector de carreras)
  function _findSheet(prefix, name) {
    var sh = ss.getSheetByName(prefix + name);
    if (sh) return sh;
    // Intentar sin sufijo __career
    var parts = name.split('__');
    if (parts.length > 1) {
      var noSuffix = parts.slice(0, parts.length-1).join('__');
      return ss.getSheetByName(prefix + noSuffix);
    }
    return null;
  }

  var shN = _findSheet('Notas_', student);
  if (shN && shN.getLastRow() > 1) {
    r.data = shN.getRange(2, 1, shN.getLastRow()-1, 20).getValues()
      .filter(function(row) { return row[0]; })
      .map(function(row) {
        return {
          num: row[0], code: row[1], name: row[2], cycle: row[3], year: row[4], uv: row[5],
          lab1_c1: row[6],  lab2_c1: row[7],  parcial_c1: row[8],
          lab1_c2: row[10], lab2_c2: row[11], parcial_c2: row[12],
          lab1_c3: row[14], lab2_c3: row[15], parcial_c3: row[16],
          finalGrade: (row[18] !== '' && row[18] !== null) ? row[18] : '',
          status: row[19] || 'pending'
        };
      });
  }

  var shC = _findSheet('Calendario_', student);
  if (shC && shC.getLastRow() > 1) {
    r.events = shC.getRange(2, 1, shC.getLastRow()-1, 7).getValues()
      .filter(function(row) { return row[0]; })
      .map(function(row, i) {
        return {
          id: row[6] ? Number(row[6]) : (new Date().getTime()+i),
          date: row[0], type: row[1], subject: row[2]||'',
          comment: row[3]||'', done: row[4]==='SÍ',
          nota: (row[5]!=='' && row[5]!==null) ? row[5] : undefined
        };
      });
  }

  var shD = _findSheet('CiclosDone_', student);
  if (shD && shD.getLastRow() > 1) {
    shD.getRange(2, 1, shD.getLastRow()-1, 2).getValues().forEach(function(row) {
      if (row[0] !== '') r.cyclesDone[row[0]] = (row[1]==='SÍ'||row[1]===true);
    });
  }

  var shA = _findSheet('Asistencias_', student);
  if (shA && shA.getLastRow() > 1) {
    shA.getRange(2, 1, shA.getLastRow()-1, 3).getValues().forEach(function(row) {
      if (row[0]) r.asistencias[row[0]] = { fecha: row[0], hora: row[1], ts: row[2]||'' };
    });
  }

  return r;
}

// ============================================================
// ESCRITURA
// ============================================================
function _saveNotas(ss, student, data) {
  var sh = _getOrCreate(ss, 'Notas_' + student);
  sh.clearContents();
  var h = ['#','Código','Materia','Ciclo','Año','UV',
    'L1-C1','L2-C1','P-C1','Cómputo1',
    'L1-C2','L2-C2','P-C2','Cómputo2',
    'L1-C3','L2-C3','P-C3','Cómputo3',
    'Nota Final','Estado'];
  sh.appendRow(h);
  _hdr(sh, h.length, '#1e40af');
  if (!data.length) return;
  var rows = data.map(function(r) {
    return [
      r.num, r.code, r.name, r.cycle, r.year, r.uv,
      _n(r.lab1_c1), _n(r.lab2_c1), _n(r.parcial_c1), _n(r.computo1),
      _n(r.lab1_c2), _n(r.lab2_c2), _n(r.parcial_c2), _n(r.computo2),
      _n(r.lab1_c3), _n(r.lab2_c3), _n(r.parcial_c3), _n(r.computo3),
      _n(r.finalGrade), r.status
    ];
  });
  sh.getRange(2, 1, rows.length, h.length).setValues(rows);
  rows.forEach(function(row, i) {
    var bg = { pass:'#d1fae5', fail:'#fee2e2', inprogress:'#fef3c7' }[row[19]] || '#fff';
    sh.getRange(i+2, 1, 1, h.length).setBackground(bg);
  });
  sh.autoResizeColumns(1, h.length);
}

function _saveCalendario(ss, student, events) {
  var sh = _getOrCreate(ss, 'Calendario_' + student);
  sh.clearContents();
  var h = ['Fecha','Tipo','Materia','Comentario','Hecho','Nota','ID'];
  sh.appendRow(h);
  _hdr(sh, h.length, '#b45309');
  if (!events.length) return;
  sh.getRange(2, 1, events.length, h.length).setValues(
    events.map(function(ev) {
      return [ev.date, ev.type, ev.subject||'', ev.comment||'',
              ev.done?'SÍ':'NO', ev.nota!==undefined?ev.nota:'', ev.id||''];
    })
  );
  sh.autoResizeColumns(1, h.length);
}

function _saveCiclosDone(ss, student, cyclesDone) {
  var sh   = _getOrCreate(ss, 'CiclosDone_' + student);
  var keys = Object.keys(cyclesDone);
  sh.clearContents();
  sh.appendRow(['CicloId','Finalizado','Actualizado']);
  _hdr(sh, 3, '#047857');
  if (!keys.length) return;
  var now = new Date().toISOString();
  sh.getRange(2, 1, keys.length, 3).setValues(
    keys.map(function(k) { return [k, cyclesDone[k]?'SÍ':'NO', now]; })
  );
}

function _saveAsistencias(ss, student, asistencias) {
  var sh  = _getOrCreate(ss, 'Asistencias_' + student);
  var arr = Object.values(asistencias).sort(function(a,b) {
    return b.fecha.localeCompare(a.fecha);
  });
  sh.clearContents();
  sh.appendRow(['Fecha','Hora','Timestamp']);
  _hdr(sh, 3, '#1d4ed8');
  if (!arr.length) return;
  sh.getRange(2, 1, arr.length, 3).setValues(
    arr.map(function(a) { return [a.fecha, a.hora, a.ts||'']; })
  );
}

function _deleteStudent(ss, name) {
  _deleteHash(ss, name);
  // Lista maestra
  var shE = ss.getSheetByName('Estudiantes');
  if (shE && shE.getLastRow() > 1) {
    var vals = shE.getRange(2, 1, shE.getLastRow()-1, 1).getValues();
    for (var i = vals.length-1; i >= 0; i--) {
      if (vals[i][0] === name) shE.deleteRow(i+2);
    }
  }
  // Hojas del estudiante
  ['Notas_','Calendario_','CiclosDone_','Asistencias_'].forEach(function(prefix) {
    var sh = ss.getSheetByName(prefix + name);
    if (sh) ss.deleteSheet(sh);
  });
  // Resumen
  var shR = ss.getSheetByName('Resumen');
  if (shR && shR.getLastRow() > 1) {
    var valsR = shR.getRange(2, 1, shR.getLastRow()-1, 1).getValues();
    for (var j = valsR.length-1; j >= 0; j--) {
      if (valsR[j][0] === name) shR.deleteRow(j+2);
    }
  }
}

function _resumen(ss, student, data, cyclesDone, asistencias) {
  var sh = _getOrCreate(ss, 'Resumen');
  if (sh.getLastRow() === 0) {
    sh.appendRow(['Estudiante','Actualizado','Total','Aprobadas','Reprobadas','En curso','Pendientes','%','Ciclos','Asistencias']);
    _hdr(sh, 10, '#1e40af');
  }
  var ap  = data.filter(function(r){return r.status==='pass';}).length;
  var row = [
    student, new Date().toISOString(), data.length, ap,
    data.filter(function(r){return r.status==='fail';}).length,
    data.filter(function(r){return r.status==='inprogress';}).length,
    data.filter(function(r){return r.status==='pending';}).length,
    data.length ? Math.round(ap/data.length*100)+'%' : '0%',
    Object.values(cyclesDone).filter(Boolean).length+'/10',
    Object.keys(asistencias).length
  ];
  var vals = sh.getLastRow()>1 ? sh.getRange(2,1,sh.getLastRow()-1,1).getValues().flat() : [];
  var idx  = vals.indexOf(student);
  if (idx>=0) sh.getRange(idx+2,1,1,row.length).setValues([row]);
  else        sh.appendRow(row);
  sh.autoResizeColumns(1,10);
}

function _registrarEstudiante(ss, name) {
  var sh = _getOrCreate(ss, 'Estudiantes');
  if (sh.getLastRow()===0) { sh.appendRow(['Nombre','Registrado']); _hdr(sh,2,'#0f172a'); }
  if (sh.getLastRow()>1) {
    var ex = sh.getRange(2,1,sh.getLastRow()-1,1).getValues().map(function(r){return r[0];});
    if (ex.indexOf(name) >= 0) return;
  }
  sh.appendRow([name, new Date().toISOString()]);
}

function _getOrCreate(ss, name) { return ss.getSheetByName(name) || ss.insertSheet(name); }
function _hdr(sh, cols, color) {
  sh.getRange(1,1,1,cols).setFontWeight('bold').setBackground(color).setFontColor('#fff').setFontSize(10);
}
function _n(v) {
  if (v===null||v===undefined||v==='') return '';
  var n = parseFloat(v); return isNaN(n)?'':n;
}
