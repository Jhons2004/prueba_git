/**
 * Asistencia.gs - Núcleo del módulo de asistencia
 * Lee alumnos desde libros de notas, pero SOLO escribe en hoja maestra de asistencia.
 */

/* =========================================================
 * CREAR / ASEGURAR ESTRUCTURA
 * ========================================================= */

function mdjAsisCrearEstructuraBaseAsistencia() {
  const ss = mdjAsisGetAsistenciaSpreadsheet_();

  let index = ss.getSheetByName(MDJ_ASIS_INDEX.sheetName);
  if (!index) index = ss.insertSheet(MDJ_ASIS_INDEX.sheetName, 0);

  mdjAsisPrepararIndex_(index);

  const rowsIndex = [];

  mdjAsisListarGradosConfigurados_().forEach(function(grado) {
    let sh = ss.getSheetByName(grado.key);
    if (!sh) sh = ss.insertSheet(grado.key);

    mdjAsisInicializarHojaGrado_(sh, grado, false);

    rowsIndex.push([
      grado.nombre,
      grado.key,
      mdjAsisContarAlumnosHoja_(sh),
      '',
      ''
    ]);
  });

  if (rowsIndex.length) {
    index.getRange(MDJ_ASIS_INDEX.rowFirstData, 1, rowsIndex.length, 5)
      .setValues(rowsIndex);
  }

  SpreadsheetApp.flush();

  return {
    ok: true,
    message: 'Estructura base creada correctamente.',
    spreadsheetId: MDJ_ASIS_APP.asistenciaSpreadsheetId,
    totalGrados: mdjAsisListarGradosConfigurados_().length
  };
}


function mdjAsisPrepararIndex_(sh) {
  sh.clear();
  sh.getRange('A1:E1').breakApart().merge()
    .setValue('MDJ - CONTROL DE ASISTENCIA')
    .setBackground('#0f4c81')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(16)
    .setHorizontalAlignment('center');

  sh.getRange('A2:E2').setValues([[
    'Grado',
    'Key',
    'Total Alumnos',
    'Última Sincronización',
    'Última Asistencia'
  ]])
    .setBackground('#eef5fb')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sh.setFrozenRows(2);
  sh.setColumnWidth(1, 280);
  sh.setColumnWidth(2, 230);
  sh.setColumnWidth(3, 140);
  sh.setColumnWidth(4, 190);
  sh.setColumnWidth(5, 190);
}


function mdjAsisInicializarHojaGrado_(sh, grado, limpiar) {
  if (limpiar) sh.clear();

  try {
    sh.getRange('A1:Z2000').breakApart();
  } catch (e) {}

  sh.getRange('A1:E1')
    .setValue(grado.nombre)
    .setBackground('#0f4c81')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(14)
    .setHorizontalAlignment('center');

  sh.getRange('A2:E2')
    .setValue('Ciclo Escolar ' + new Date().getFullYear())
    .setBackground('#eef5fb')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sh.getRange('A4:E4').setValues([[
    'No.',
    'Nombre',
    'Total Días',
    'Días Asistidos',
    '% Asistencia'
  ]])
    .setBackground('#0f4c81')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sh.setFrozenRows(MDJ_ASIS_LAYOUT.rowHeaders);
  sh.setFrozenColumns(2);

  sh.setColumnWidth(1, 70);
  sh.setColumnWidth(2, 340);
  sh.setColumnWidth(3, 110);
  sh.setColumnWidth(4, 130);
  sh.setColumnWidth(5, 130);

  sh.getRange('A:E').setVerticalAlignment('middle');
}


/* =========================================================
 * SINCRONIZACIÓN DESDE LIBROS DE NOTAS
 * ========================================================= */

function mdjAsisPreviewAlumnosDesdeNotas(token, libroKey) {
  try {
    mdjAsisRequireRole_(token, ['admin']);

    const grado = mdjAsisGetLibroConfig_(libroKey); // <- Muro de fuego activado
    const alumnos = mdjAsisLeerAlumnosDesdeNotas_(grado);

    return {
      ok: true,
      grado: grado.nombre,
      total: alumnos.length,
      alumnos: alumnos.slice(0, 80)
    };

  } catch (e) {
    return { ok: false, message: e.message };
  }
}


function mdjAsisSincronizarTodos(token) {
  try {
    mdjAsisRequireRole_(token, ['admin']);

    const resultados = mdjAsisListarGradosConfigurados_().map(function(g) {
      const r = mdjAsisSincronizarGrado(token, g.key, false);
      return {
        grado: g.nombre,
        key: g.key,
        ok: r.ok,
        message: r.message,
        total: r.total || 0,
        agregados: r.agregados || 0
      };
    });

    const errores = resultados.filter(function(r) { return !r.ok; }).length;

    return {
      ok: errores === 0,
      message: errores === 0
        ? 'Todos los grados fueron sincronizados correctamente.'
        : 'Sincronización terminada con ' + errores + ' error(es).',
      resultados: resultados
    };

  } catch (e) {
    return { ok: false, message: e.message };
  }
}


function mdjAsisSincronizarGrado(token, libroKey, forceReset) {
  try {
    mdjAsisRequireRole_(token, ['admin']);

    forceReset = !!forceReset;

    const grado = mdjAsisGetLibroConfig_(libroKey); // <- Muro de fuego activado

    const ss = mdjAsisGetAsistenciaSpreadsheet_();
    let sh = ss.getSheetByName(grado.key);
    if (!sh) sh = ss.insertSheet(grado.key);

    if (forceReset) {
      mdjAsisInicializarHojaGrado_(sh, grado, true);
    } else {
      mdjAsisInicializarHojaGrado_(sh, grado, false);
    }

    const alumnosFuente = mdjAsisLeerAlumnosDesdeNotas_(grado);
    if (!alumnosFuente.length) {
      throw new Error('No se encontraron alumnos en la hoja "' + MDJ_ASIS_APP.hojaFuenteAlumnos + '".');
    }

    const existentes = mdjAsisLeerAlumnosHoja_(sh);
    const mapa = {};

    existentes.forEach(function(a) {
      mapa[mdjAsisNormalizeStudentName_(a.nombre)] = true;
    });

    let agregados = 0;

    if (forceReset || existentes.length === 0) {
      const rows = alumnosFuente.map(function(a, idx) {
        return [
          a.numero || (idx + 1),
          a.nombre,
          0,
          0,
          0
        ];
      });

      if (rows.length) {
        sh.getRange(MDJ_ASIS_LAYOUT.rowFirstStudent, 1, rows.length, 5).setValues(rows);
        sh.getRange(MDJ_ASIS_LAYOUT.rowFirstStudent, 5, rows.length, 1).setNumberFormat('0.00%');
      }

      agregados = rows.length;

    } else {
      const nuevos = alumnosFuente.filter(function(a) {
        return !mapa[mdjAsisNormalizeStudentName_(a.nombre)];
      });

      if (nuevos.length) {
        const startRow = Math.max(sh.getLastRow() + 1, MDJ_ASIS_LAYOUT.rowFirstStudent);

        const rowsNuevos = nuevos.map(function(a, idx) {
          return [
            a.numero || (existentes.length + idx + 1),
            a.nombre,
            0,
            0,
            0
          ];
        });

        sh.getRange(startRow, 1, rowsNuevos.length, 5).setValues(rowsNuevos);
        sh.getRange(startRow, 5, rowsNuevos.length, 1).setNumberFormat('0.00%');

        agregados = rowsNuevos.length;
      }
    }

    mdjAsisRecalcularTotalesHoja_(sh);
    mdjAsisAplicarEstiloHojaGrado_(sh);
    mdjAsisActualizarIndexGrado_(grado.key, {
      totalAlumnos: mdjAsisContarAlumnosHoja_(sh),
      ultimaSincronizacion: mdjAsisFechaHoraAhora_()
    });

    SpreadsheetApp.flush();

    return {
      ok: true,
      message: 'Grado sincronizado correctamente. Alumnos agregados: ' + agregados,
      grado: grado.nombre,
      key: grado.key,
      agregados: agregados,
      total: mdjAsisContarAlumnosHoja_(sh)
    };

  } catch (e) {
    return { ok: false, message: e.message };
  }
}


/* =========================================================
 * FUNCIONES DE LECTURA DE ALUMNOS CON FILTRO ESTRICTO
 * ========================================================= */

function mdjAsisLeerAlumnosDesdeNotas_(grado) {
  const ss = mdjGetSpreadsheetByKey_(grado.key);
  
  // Buscar hoja (Padrón o I BIMESTRE)
  const sh = mdjGetPadronSheet_(ss);

  if (!sh) {
    throw new Error('No se encontró la hoja de alumnos para "' + grado.nombre + '".');
  }

  // Detectar configuración según la hoja que encontramos
  const isBimestre = sh.getName() === MDJ_ASIS_APP.hojaFuenteAlumnos;
  const startRow = isBimestre ? MDJ_ASIS_APP.layoutFuenteAlumnos.rowFirstStudent : 2;
  const colNum = isBimestre ? MDJ_ASIS_APP.layoutFuenteAlumnos.colNumero : 1;
  const colNom = isBimestre ? MDJ_ASIS_APP.layoutFuenteAlumnos.colNombre : 2;

  const lastRow = sh.getLastRow();
  if (lastRow < startRow) return [];

  const numRows = lastRow - startRow + 1;
  const maxCol = Math.max(colNum, colNom);
  const values = sh.getRange(startRow, 1, numRows, maxCol).getDisplayValues();

  const alumnos = [];
  const vistos = {};
  let vaciosSeguidos = 0;

  values.forEach(function(row) {
    const numero = String(row[colNum - 1] || '').trim();
    let nombre = String(row[colNom - 1] || '').trim();

    // 1. FILTRO ANTI-BASURA: Ignorar celdas vacías, guiones, ceros o errores de fórmula
    if (!nombre || nombre === '-' || nombre === '0' || nombre === '#N/A' || nombre === '#REF!') {
      vaciosSeguidos++;
      return;
    }

    // 2. FILTRO DE PROMEDIOS: Ignorar filas de totales que están al final de las notas
    const lowerName = nombre.toLowerCase();
    if (lowerName.indexOf('promedio') !== -1 || lowerName.indexOf('total') !== -1 || lowerName.indexOf('estadistica') !== -1) {
      return;
    }

    if (vaciosSeguidos > 20) return; // Si hay demasiados vacíos seguidos, terminamos.

    // Limpieza de nombre si tienes la función declarada
    if (typeof mdjAsisCleanStudentName_ === 'function') {
       nombre = mdjAsisCleanStudentName_(nombre);
    }

    let key = nombre.toLowerCase().replace(/\s+/g, '');
    if (typeof mdjAsisNormalizeStudentName_ === 'function') {
       key = mdjAsisNormalizeStudentName_(nombre) || key;
    }

    // Evitar duplicados
    if (!key || vistos[key]) return;
    vistos[key] = true;

    alumnos.push({
      numero: numero,
      nombre: nombre
    });
  });

  return alumnos;
}


function mdjAsisResolverSpreadsheetFuente_(grado) {
  return mdjGetSpreadsheetByKey_(grado.key).getId();
}


/* =========================================================
 * TOMAR ASISTENCIA
 * ========================================================= */

function mdjAsisGetListaAlumnos(token, libroKey, fechaISO) {
  try {
    const sess = mdjAsisRequireSession_(token);
    const role = mdjAsisNormalizeRole_(sess.user.rol);

    if (role !== 'admin' && role !== 'maestro') {
      throw new Error('No tienes permisos para tomar asistencia.');
    }

    if (!mdjAsisIsFechaISO_(fechaISO)) {
      throw new Error('Fecha inválida. Usa formato yyyy-MM-dd.');
    }

    const grado = mdjAsisGetLibroConfig_(libroKey);
    mdjAsisAssertGradoAccess_(sess.user, libroKey);

    const sh = mdjAsisGetSheetGrado_(libroKey);
    // VALIDACIÓN CRÍTICA PARA EVITAR EL COLAPSO:
    if (!sh) {
      throw new Error('La pestaña para este grado no existe en la hoja maestra. Pídele al administrador que sincronice el grado primero.');
    }

    const alumnos = mdjAsisLeerAlumnosHoja_(sh);

    if (!alumnos.length) {
      return {
        ok: true,
        grado: grado,
        fechaISO: fechaISO,
        fechaHeader: mdjAsisFechaToHeader_(fechaISO),
        fechaExiste: false,
        alumnos: []
      };
    }

    const fechaHeader = mdjAsisFechaToHeader_(fechaISO);
    const colFecha = mdjAsisBuscarColFecha_(sh, fechaHeader);
    const fechaExiste = colFecha > 0;

    let estados = [];
    let notas = [];

    if (fechaExiste) {
      estados = sh.getRange(MDJ_ASIS_LAYOUT.rowFirstStudent, colFecha, alumnos.length, 1).getDisplayValues();
      notas = sh.getRange(MDJ_ASIS_LAYOUT.rowFirstStudent, colFecha, alumnos.length, 1).getNotes();
    }

    const out = alumnos.map(function(a, idx) {
      return {
        rowNumber: a.rowNumber,
        numero: a.numero,
        nombre: a.nombre,
        estado: fechaExiste ? mdjAsisNormalizeEstado_(estados[idx][0]) : '',
        justificacion: fechaExiste ? String(notas[idx][0] || '') : ''
      };
    });

    return {
      ok: true,
      grado: grado,
      fechaISO: fechaISO,
      fechaHeader: fechaHeader,
      fechaExiste: fechaExiste,
      alumnos: out
    };

  } catch (e) {
    return { ok: false, message: e.message };
  }
}


function mdjAsisGuardarAsistencia(token, payload) {
  try {
    const sess = mdjAsisRequireSession_(token);
    const role = mdjAsisNormalizeRole_(sess.user.rol);

    if (role !== 'admin' && role !== 'maestro') {
      throw new Error('No tienes permisos para guardar asistencia.');
    }

    payload = payload || {};

    const libroKey = String(payload.libroKey || '').trim();
    const fechaISO = String(payload.fechaISO || '').trim();
    const registros = payload.registros || [];

    if (!libroKey) throw new Error('Grado no enviado.');
    if (!mdjAsisIsFechaISO_(fechaISO)) throw new Error('Fecha inválida.');
    if (!registros.length) throw new Error('No hay registros para guardar.');

    const grado = mdjAsisGetLibroConfig_(libroKey);
    mdjAsisAssertGradoAccess_(sess.user, libroKey);

    const sh = mdjAsisGetSheetGrado_(libroKey);
    // VALIDACIÓN CRÍTICA:
    if (!sh) throw new Error('La pestaña para este grado no existe.');

    const alumnos = mdjAsisLeerAlumnosHoja_(sh);

    if (!alumnos.length) {
      throw new Error('La hoja del grado no tiene alumnos. Sincroniza primero desde el panel administrativo.');
    }

    if (!Array.isArray(registros)) throw new Error('Registros de asistencia inválidos.');
    const rowsAllowed = {};
    alumnos.forEach(function(a) { rowsAllowed[a.rowNumber] = true; });
    const rowSeen = {};
    registros.forEach(function(r) {
      const row = Number(r.rowNumber);
      if (!Number.isInteger(row) || !rowsAllowed[row] || rowSeen[row]) {
        throw new Error('El envío contiene alumnos duplicados o filas ajenas al grado.');
      }
      rowSeen[row] = true;
      const original = String(r.estado || '').trim();
      if (original && !mdjAsisNormalizeEstado_(original)) throw new Error('Estado de asistencia no permitido.');
    });
    if (registros.length !== alumnos.length) {
      throw new Error('La lista de alumnos cambió o está incompleta. Vuelve a cargar antes de guardar.');
    }
    const fechaHeader = mdjAsisFechaToHeader_(fechaISO);
    const colFecha = mdjAsisGetOrCreateColFecha_(sh, fechaHeader);

    const mapaRegistros = {};
    registros.forEach(function(r) {
      const row = Number(r.rowNumber || 0);
      if (row > 0) {
        mapaRegistros[row] = {
          estado: mdjAsisNormalizeEstado_(r.estado),
          justificacion: String(r.justificacion || '').trim()
        };
      }
    });

    const values = [];
    const notes = [];

    alumnos.forEach(function(a) {
      const reg = mapaRegistros[a.rowNumber] || {};
      const estado = mdjAsisNormalizeEstado_(reg.estado);
      const justificacion = String(reg.justificacion || '').trim();

      values.push([estado]);
      notes.push([justificacion]);
    });

    const lock = LockService.getScriptLock();
    if (!lock.tryLock(20000)) throw new Error('Hay otra operación guardando asistencia. Reintenta.');
    try {
    sh.getRange(MDJ_ASIS_LAYOUT.rowFirstStudent, colFecha, values.length, 1).setValues(values);
    sh.getRange(MDJ_ASIS_LAYOUT.rowFirstStudent, colFecha, notes.length, 1).setNotes(notes);

    mdjAsisColorearColumnaFecha_(sh, colFecha, values.length);
    mdjAsisRecalcularTotalesHoja_(sh);
    mdjAsisAplicarEstiloHojaGrado_(sh);

    mdjAsisActualizarIndexGrado_(grado.key, {
      totalAlumnos: alumnos.length,
      ultimaAsistencia: mdjAsisFechaHoraAhora_()
    });

    SpreadsheetApp.flush();
    } finally { lock.releaseLock(); }

    return {
      ok: true,
      message: 'Asistencia guardada correctamente.',
      grado: grado.nombre,
      fechaISO: fechaISO,
      fechaHeader: fechaHeader,
      total: values.length
    };

  } catch (e) {
    return { ok: false, message: e.message };
  }
}


/* =========================================================
 * REPORTES ADMIN
 * ========================================================= */

function mdjAsisGetReporteGrado(token, libroKey) {
  try {
    mdjAsisRequireRole_(token, ['admin']);

    const grado = mdjAsisGetLibroConfig_(libroKey);
    mdjAsisAssertGradoAccess_(sess.user, libroKey);

    const sh = mdjAsisGetSheetGrado_(libroKey);
    if (!sh) throw new Error('La pestaña para este grado no existe.');

    const alumnos = mdjAsisLeerAlumnosHoja_(sh);
    const fechas = mdjAsisLeerFechasHoja_(sh);

    if (!alumnos.length) {
      return {
        ok: true,
        grado: grado,
        totalAlumnos: 0,
        totalDias: fechas.length,
        fechas: fechas,
        alumnos: []
      };
    }

    let matriz = [];
    let notas = [];

    if (fechas.length) {
      matriz = sh.getRange(
        MDJ_ASIS_LAYOUT.rowFirstStudent,
        MDJ_ASIS_LAYOUT.colFirstDate,
        alumnos.length,
        fechas.length
      ).getDisplayValues();

      notas = sh.getRange(
        MDJ_ASIS_LAYOUT.rowFirstStudent,
        MDJ_ASIS_LAYOUT.colFirstDate,
        alumnos.length,
        fechas.length
      ).getNotes();
    }

    const out = alumnos.map(function(a, idx) {
      const dias = fechas.map(function(f, dIdx) {
        const estado = fechas.length ? mdjAsisNormalizeEstado_(matriz[idx][dIdx]) : '';
        const justificacion = fechas.length ? String(notas[idx][dIdx] || '') : '';

        return {
          fecha: f.fechaISO,
          header: f.header,
          col: f.col,
          estado: estado,
          justificacion: justificacion
        };
      });

      const asistidos = dias.filter(function(d) {
        return mdjAsisEstadoCuentaAsistencia_(d.estado);
      }).length;

      const totalDias = fechas.length;
      const porcentajeNum = totalDias ? asistidos / totalDias : 0;

      return {
        rowNumber: a.rowNumber,
        numero: a.numero,
        nombre: a.nombre,
        totalDias: totalDias,
        diasAsistidos: asistidos,
        porcentaje: Utilities.formatString('%.2f%%', porcentajeNum * 100),
        dias: dias
      };
    });

    return {
      ok: true,
      grado: grado,
      totalAlumnos: alumnos.length,
      totalDias: fechas.length,
      fechas: fechas,
      alumnos: out
    };

  } catch (e) {
    return { ok: false, message: e.message };
  }
}


function mdjAsisEditarCelda(token, payload) {
  try {
    mdjAsisRequireRole_(token, ['admin']);

    payload = payload || {};

    const libroKey = String(payload.libroKey || '').trim();
    const rowNumber = Number(payload.rowNumber || 0);
    const col = Number(payload.col || 0);
    const estado = mdjAsisNormalizeEstado_(payload.estado);
    const justificacion = String(payload.justificacion || '').trim();

    if (!libroKey) throw new Error('Grado no enviado.');
    if (rowNumber < MDJ_ASIS_LAYOUT.rowFirstStudent) throw new Error('Fila inválida.');
    if (col < MDJ_ASIS_LAYOUT.colFirstDate) throw new Error('Columna inválida.');

    const sh = mdjAsisGetSheetGrado_(libroKey);
    if (!sh) throw new Error('La pestaña para este grado no existe.');

    sh.getRange(rowNumber, col).setValue(estado);
    sh.getRange(rowNumber, col).setNote(justificacion);

    mdjAsisColorearCeldaEstado_(sh.getRange(rowNumber, col), estado);
    mdjAsisRecalcularTotalesHoja_(sh);

    SpreadsheetApp.flush();

    return {
      ok: true,
      message: 'Celda actualizada correctamente.'
    };

  } catch (e) {
    return { ok: false, message: e.message };
  }
}


function mdjAsisObtenerEnlaceHoja(token) {
  try {
    mdjAsisRequireRole_(token, ['admin']);

    return {
      ok: true,
      spreadsheetId: MDJ_ASIS_APP.asistenciaSpreadsheetId,
      url: mdjAsisBuildAsistenciaUrl_()
    };

  } catch (e) {
    return { ok: false, message: e.message };
  }
}


/* =========================================================
 * USUARIOS ADMIN
 * ========================================================= */

function mdjAsisListarUsuarios(token) {
  try {
    mdjAsisRequireRole_(token, ['admin']);

    const sh = mdjAsisGetControlSheet_(MDJ_ASIS_APP.hojaUsuarios);
    const values = sh.getDataRange().getDisplayValues();

    if (values.length < 2) {
      return { ok: true, usuarios: [] };
    }

    const header = mdjAsisBuildUserHeaderMap_(values[0]);
    const usuarios = [];

    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const usuario = String(row[header.usuario] || '').trim();
      if (!usuario) continue;

      usuarios.push({
        rowNumber: i + 1,
        usuario: usuario,
        nombre: String(row[header.nombre] || usuario),
        rol: mdjAsisNormalizeRole_(row[header.rol]),
        activo: mdjAsisIsTruthy_(row[header.activo])
      });
    }

    return { ok: true, usuarios: usuarios };

  } catch (e) {
    return { ok: false, message: e.message };
  }
}


function mdjAsisCrearUsuario(token, payload) {
  try {
    mdjAsisRequireRole_(token, ['admin']);

    payload = payload || {};

    const usuario = String(payload.usuario || '').trim();
    const nombre = String(payload.nombre || '').trim();
    const password = String(payload.password || '').trim();
    const rol = mdjAsisNormalizeRole_(payload.rol);
    if (!['admin','maestro'].includes(String(payload.rol || '').trim().toLowerCase())) throw new Error('Rol de usuario no permitido.');
    const activo = !!payload.activo;

    if (!usuario || !nombre || !password) {
      throw new Error('Usuario, nombre y contraseña son requeridos.');
    }

    if (mdjAsisFindUser_(usuario)) {
      throw new Error('Ya existe un usuario con ese nombre.');
    }

    const sh = mdjAsisGetControlSheet_(MDJ_ASIS_APP.hojaUsuarios);
    const values = sh.getDataRange().getDisplayValues();
    const header = mdjAsisBuildUserHeaderMap_(values[0]);

    const row = new Array(values[0].length).fill('');
    row[header.usuario] = usuario;
    row[header.password] = mdjHashTextoSha256(password);
    row[header.nombre] = nombre;
    row[header.rol] = rol;
    row[header.activo] = activo ? 'SI' : 'NO';

    sh.appendRow(row);

    return { ok: true, message: 'Usuario creado correctamente.' };

  } catch (e) {
    return { ok: false, message: e.message };
  }
}


function mdjAsisActualizarUsuario(token, payload) {
  try {
    mdjAsisRequireRole_(token, ['admin']);

    payload = payload || {};

    const rowNumber = Number(payload.rowNumber || 0);
    const nombre = String(payload.nombre || '').trim();
    const password = String(payload.password || '').trim();
    const rol = mdjAsisNormalizeRole_(payload.rol);
    if (!['admin','maestro'].includes(String(payload.rol || '').trim().toLowerCase())) throw new Error('Rol de usuario no permitido.');
    const activo = !!payload.activo;

    if (!Number.isInteger(rowNumber) || rowNumber < 2) throw new Error('Fila de usuario inválida.');
    if (rowNumber > mdjAsisGetControlSheet_(MDJ_ASIS_APP.hojaUsuarios).getLastRow()) throw new Error('Fila inexistente.');
    if (!nombre) throw new Error('El nombre es requerido.');

    const sh = mdjAsisGetControlSheet_(MDJ_ASIS_APP.hojaUsuarios);
    const values = sh.getDataRange().getDisplayValues();
    const header = mdjAsisBuildUserHeaderMap_(values[0]);

    sh.getRange(rowNumber, header.nombre + 1).setValue(nombre);
    sh.getRange(rowNumber, header.rol + 1).setValue(rol);
    sh.getRange(rowNumber, header.activo + 1).setValue(activo ? 'SI' : 'NO');

    if (password) {
      sh.getRange(rowNumber, header.password + 1).setValue(mdjHashTextoSha256(password));
    }

    return { ok: true, message: 'Usuario actualizado correctamente.' };

  } catch (e) {
    return { ok: false, message: e.message };
  }
}


function mdjAsisEliminarUsuario(token, rowNumber) {
  try {
    mdjAsisRequireRole_(token, ['admin']);

    rowNumber = Number(rowNumber || 0);
    if (!Number.isInteger(rowNumber) || rowNumber < 2) throw new Error('Fila de usuario inválida.');
    if (rowNumber > mdjAsisGetControlSheet_(MDJ_ASIS_APP.hojaUsuarios).getLastRow()) throw new Error('Fila inexistente.');

    const sh = mdjAsisGetControlSheet_(MDJ_ASIS_APP.hojaUsuarios);
    const values = sh.getDataRange().getDisplayValues();
    const header = mdjAsisBuildUserHeaderMap_(values[0]);
    const session = mdjAsisRequireRole_(token, ['admin']);
    const deletedUsername = String(values[rowNumber - 1][header.usuario] || '');
    if (mdjAsisNormalizeText_(deletedUsername) === mdjAsisNormalizeText_(session.user.usuario)) {
      throw new Error('No puedes eliminar tu propia cuenta.');
    }
    const targetRole = mdjAsisNormalizeRole_(values[rowNumber - 1][header.rol]);
    if (targetRole === 'admin') {
      const totalAdmins = values.slice(1).filter(function(r) {
        return mdjAsisNormalizeRole_(r[header.rol]) === 'admin' && mdjAsisIsTruthy_(r[header.activo]);
      }).length;
      if (totalAdmins <= 1) throw new Error('No se puede eliminar el último administrador activo.');
    }
    sh.deleteRow(rowNumber);

    return { ok: true, message: 'Usuario eliminado correctamente.' };

  } catch (e) {
    return { ok: false, message: e.message };
  }
}


/* =========================================================
 * FUNCIONES INTERNAS DE HOJAS
 * ========================================================= */

function mdjAsisLeerAlumnosHoja_(sh) {
  const lastRow = sh.getLastRow();
  if (lastRow < MDJ_ASIS_LAYOUT.rowFirstStudent) return [];

  const numRows = lastRow - MDJ_ASIS_LAYOUT.rowFirstStudent + 1;
  const values = sh.getRange(MDJ_ASIS_LAYOUT.rowFirstStudent, 1, numRows, 2).getDisplayValues();

  const alumnos = [];

  values.forEach(function(row, idx) {
    const numero = String(row[0] || '').trim();
    const nombre = mdjAsisCleanStudentName_(row[1]);

    if (!nombre) return;

    alumnos.push({
      rowNumber: MDJ_ASIS_LAYOUT.rowFirstStudent + idx,
      numero: numero,
      nombre: nombre
    });
  });

  return alumnos;
}


function mdjAsisContarAlumnosHoja_(sh) {
  return mdjAsisLeerAlumnosHoja_(sh).length;
}


function mdjAsisLeerFechasHoja_(sh) {
  const lastCol = sh.getLastColumn();

  if (lastCol < MDJ_ASIS_LAYOUT.colFirstDate) return [];

  const width = lastCol - MDJ_ASIS_LAYOUT.colFirstDate + 1;
  const headers = sh.getRange(MDJ_ASIS_LAYOUT.rowHeaders, MDJ_ASIS_LAYOUT.colFirstDate, 1, width)
    .getDisplayValues()[0];

  const fechas = [];

  headers.forEach(function(h, idx) {
    h = String(h || '').trim();
    if (!h) return;

    fechas.push({
      header: h,
      fechaISO: mdjAsisHeaderToFecha_(h),
      col: MDJ_ASIS_LAYOUT.colFirstDate + idx
    });
  });

  return fechas;
}


function mdjAsisBuscarColFecha_(sh, fechaHeader) {
  const fechas = mdjAsisLeerFechasHoja_(sh);

  for (let i = 0; i < fechas.length; i++) {
    if (String(fechas[i].header) === String(fechaHeader)) {
      return fechas[i].col;
    }
  }

  return -1;
}


function mdjAsisGetOrCreateColFecha_(sh, fechaHeader) {
  let col = mdjAsisBuscarColFecha_(sh, fechaHeader);

  if (col > 0) return col;

  const lastCol = Math.max(sh.getLastColumn(), MDJ_ASIS_LAYOUT.colPorcentaje);
  col = lastCol + 1;

  if (col < MDJ_ASIS_LAYOUT.colFirstDate) {
    col = MDJ_ASIS_LAYOUT.colFirstDate;
  }

  sh.getRange(MDJ_ASIS_LAYOUT.rowHeaders, col)
    .setValue(fechaHeader)
    .setBackground('#1f6fb2')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sh.setColumnWidth(col, 105);

  return col;
}


function mdjAsisRecalcularTotalesHoja_(sh) {
  const alumnos = mdjAsisLeerAlumnosHoja_(sh);
  const fechas = mdjAsisLeerFechasHoja_(sh);

  if (!alumnos.length) return;

  const totalDias = fechas.length;

  if (!totalDias) {
    const zeros = alumnos.map(function() {
      return [0, 0, 0];
    });

    sh.getRange(MDJ_ASIS_LAYOUT.rowFirstStudent, MDJ_ASIS_LAYOUT.colTotalDias, zeros.length, 3)
      .setValues(zeros);

    sh.getRange(MDJ_ASIS_LAYOUT.rowFirstStudent, MDJ_ASIS_LAYOUT.colPorcentaje, zeros.length, 1)
      .setNumberFormat('0.00%');

    return;
  }

  const matriz = sh.getRange(
    MDJ_ASIS_LAYOUT.rowFirstStudent,
    MDJ_ASIS_LAYOUT.colFirstDate,
    alumnos.length,
    totalDias
  ).getDisplayValues();

  const resumen = matriz.map(function(row) {
    let asistidos = 0;

    row.forEach(function(v) {
      if (mdjAsisEstadoCuentaAsistencia_(v)) asistidos++;
    });

    const porcentaje = totalDias ? asistidos / totalDias : 0;

    return [
      totalDias,
      asistidos,
      porcentaje
    ];
  });

  sh.getRange(MDJ_ASIS_LAYOUT.rowFirstStudent, MDJ_ASIS_LAYOUT.colTotalDias, resumen.length, 3)
    .setValues(resumen);

  sh.getRange(MDJ_ASIS_LAYOUT.rowFirstStudent, MDJ_ASIS_LAYOUT.colPorcentaje, resumen.length, 1)
    .setNumberFormat('0.00%');
}


function mdjAsisAplicarEstiloHojaGrado_(sh) {
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();

  if (lastRow >= MDJ_ASIS_LAYOUT.rowFirstStudent) {
    sh.getRange(
      MDJ_ASIS_LAYOUT.rowFirstStudent,
      1,
      lastRow - MDJ_ASIS_LAYOUT.rowFirstStudent + 1,
      Math.max(lastCol, 5)
    )
      .setBorder(true, true, true, true, true, true, '#dbe7f3', SpreadsheetApp.BorderStyle.SOLID);
  }

  if (lastCol >= MDJ_ASIS_LAYOUT.colFirstDate) {
    sh.getRange(
      MDJ_ASIS_LAYOUT.rowHeaders,
      MDJ_ASIS_LAYOUT.colFirstDate,
      1,
      lastCol - MDJ_ASIS_LAYOUT.colFirstDate + 1
    )
      .setBackground('#1f6fb2')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
  }
}


function mdjAsisColorearColumnaFecha_(sh, colFecha, totalRows) {
  if (!totalRows) return;

  const range = sh.getRange(MDJ_ASIS_LAYOUT.rowFirstStudent, colFecha, totalRows, 1);
  const values = range.getDisplayValues();

  for (let i = 0; i < values.length; i++) {
    mdjAsisColorearCeldaEstado_(
      sh.getRange(MDJ_ASIS_LAYOUT.rowFirstStudent + i, colFecha),
      values[i][0]
    );
  }
}


function mdjAsisColorearCeldaEstado_(range, estado) {
  const e = mdjAsisNormalizeEstado_(estado);

  if (e === 'SI') {
    range.setBackground('#d4f5dd').setFontColor('#0d5e3a').setFontWeight('bold').setHorizontalAlignment('center');
  } else if (e === 'NO') {
    range.setBackground('#fddddd').setFontColor('#8a1f1f').setFontWeight('bold').setHorizontalAlignment('center');
  } else if (e === 'JUSTIFICADO') {
    range.setBackground('#fff3cd').setFontColor('#7a5a00').setFontWeight('bold').setHorizontalAlignment('center');
  } else {
    range.setBackground('#ffffff').setFontColor('#000000').setFontWeight('normal').setHorizontalAlignment('center');
  }
}


function mdjAsisActualizarIndexGrado_(gradoKey, data) {
  data = data || {};

  const ss = mdjAsisGetAsistenciaSpreadsheet_();
  let sh = ss.getSheetByName(MDJ_ASIS_INDEX.sheetName);

  if (!sh) {
    sh = ss.insertSheet(MDJ_ASIS_INDEX.sheetName, 0);
    mdjAsisPrepararIndex_(sh);
  }

  let grado;
  try {
    grado = mdjAsisGetLibroConfig_(gradoKey);
  } catch (e) {
    return;
  }

  const lastRow = Math.max(sh.getLastRow(), MDJ_ASIS_INDEX.rowFirstData);
  const values = sh.getRange(MDJ_ASIS_INDEX.rowFirstData, 1, Math.max(lastRow - 2, 1), 5).getDisplayValues();

  let targetRow = -1;

  for (let i = 0; i < values.length; i++) {
    if (String(values[i][1] || '') === gradoKey) {
      targetRow = MDJ_ASIS_INDEX.rowFirstData + i;
      break;
    }
  }

  if (targetRow === -1) {
    targetRow = sh.getLastRow() + 1;
    sh.getRange(targetRow, 1, 1, 5).setValues([[
      grado.nombre,
      grado.key,
      data.totalAlumnos || 0,
      data.ultimaSincronizacion || '',
      data.ultimaAsistencia || ''
    ]]);
    return;
  }

  if (data.totalAlumnos !== undefined) {
    sh.getRange(targetRow, MDJ_ASIS_INDEX.colTotalAlumnos).setValue(data.totalAlumnos);
  }

  if (data.ultimaSincronizacion !== undefined) {
    sh.getRange(targetRow, MDJ_ASIS_INDEX.colUltimaSincronizacion).setValue(data.ultimaSincronizacion);
  }

  if (data.ultimaAsistencia !== undefined) {
    sh.getRange(targetRow, MDJ_ASIS_INDEX.colUltimaAsistencia).setValue(data.ultimaAsistencia);
  }
}


/* =========================================================
 * PRUEBAS RÁPIDAS
 * ========================================================= */

function mdjAsisTestAbrirAsistencia() {
  const ss = mdjAsisGetAsistenciaSpreadsheet_();

  return {
    ok: true,
    id: ss.getId(),
    nombre: ss.getName(),
    url: mdjAsisBuildAsistenciaUrl_()
  };
}


function mdjAsisTestLeerAlumnosPrimero() {
  const grado = mdjAsisGetLibroConfig_('primero_basico');
  const alumnos = mdjAsisLeerAlumnosDesdeNotas_(grado);

  return {
    ok: true,
    grado: grado.nombre,
    total: alumnos.length,
    primeros: alumnos.slice(0, 10)
  };
}