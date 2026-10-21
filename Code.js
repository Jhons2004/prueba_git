/**
 * Codigo.gs
 * Colegio Mundo de Juguetes
 * Núcleo del sistema + conversión masiva a Google Sheets
 *
 * REQUIERE:
 * - Auth.gs existente (mdjGetSession, mdjLogin, mdjLogout, etc.)
 * - Boletas.gs existente
 * - Drive API avanzada activada
 */

const MDJ_CONVERTED_PREFIX = 'MDJ_CONVERTED_';
const MDJ_MIME_GOOGLE_SHEETS = 'application/vnd.google-apps.spreadsheet';
const MDJ_MIME_XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const MDJ_MIME_XLS = 'application/vnd.ms-excel';

/* ============================================================
 * WEB APP
 * ============================================================ */

function doGet(e) {
  const page = String((e && e.parameter && e.parameter.page) || 'login').trim().toLowerCase();

  const fileMap = {
    login: MDJ_APP.loginPage,
    admin: MDJ_APP.adminPage,
    administrativo: MDJ_APP.adminPage,
    maestro: MDJ_APP.teacherPage,
    teacher: MDJ_APP.teacherPage,
    tomar: MDJ_ASIS_APP.tomarPage,
    asistencia: MDJ_ASIS_APP.tomarPage,
    'asistencia-admin': MDJ_ASIS_APP.adminPage,
    'admin-asistencia': MDJ_ASIS_APP.adminPage
  };

  const fileName = fileMap[page] || MDJ_APP.loginPage;
  const tpl = HtmlService.createTemplateFromFile(fileName);

  tpl.app = MDJ_APP;
  tpl.logoDataUri = (typeof MDJ_LOGO_DATA_URI !== 'undefined' && MDJ_LOGO_DATA_URI) ? MDJ_LOGO_DATA_URI : '';
  tpl.webAppUrl = ScriptApp.getService().getUrl();
  tpl.currentPage = page;
  tpl.asistenciaUrl = mdjAsisGetAsistenciaUrl_();
  tpl.asistenciaSpreadsheetId = MDJ_ASIS_APP.asistenciaSpreadsheetId;

  return tpl.evaluate()
    .setTitle(MDJ_APP.nombre + ' - ' + fileName)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* ============================================================
 * LOGS
 * ============================================================ */

function mdjLog_(nivel, mensaje, extra) {
  const ts = new Date().toISOString();
  try {
    if (extra !== undefined) {
      Logger.log('[' + ts + '] [MDJ ' + nivel + '] ' + mensaje + ' | ' + JSON.stringify(extra));
    } else {
      Logger.log('[' + ts + '] [MDJ ' + nivel + '] ' + mensaje);
    }
  } catch (e) {
    Logger.log('[' + ts + '] [MDJ ' + nivel + '] ' + mensaje);
  }
}

/* ============================================================
 * PREAUTORIZACIÓN
 * ============================================================ */

function mdjPreAutorizarSistema() {
  const out = {
    ok: true,
    activeUser: '',
    effectiveUser: '',
    controlSpreadsheetId: '',
    controlSpreadsheetName: '',
    rootFolderName: '',
    driveAdvancedEnabled: false,
    driveAdvancedTest: '',
    propertiesOk: false,
    cacheOk: false,
    message: ''
  };

  try {
    out.activeUser = Session.getActiveUser().getEmail() || '';
  } catch (e) {}

  try {
    out.effectiveUser = Session.getEffectiveUser().getEmail() || '';
  } catch (e) {}

  try {
    const ss = SpreadsheetApp.openById(MDJ_APP.hojaControlId);
    out.controlSpreadsheetId = ss.getId();
    out.controlSpreadsheetName = ss.getName();
  } catch (e) {
    throw new Error('No se pudo abrir la hoja de control. ' + (e && e.message ? e.message : e));
  }

  try {
    const root = DriveApp.getRootFolder();
    out.rootFolderName = root.getName();
  } catch (e) {
    throw new Error('No se pudo autorizar DriveApp. ' + (e && e.message ? e.message : e));
  }

  try {
    const props = PropertiesService.getScriptProperties();
    props.setProperty('MDJ_AUTH_TEST', new Date().toISOString());
    out.propertiesOk = true;
  } catch (e) {
    throw new Error('No se pudo autorizar PropertiesService. ' + (e && e.message ? e.message : e));
  }

  try {
    const cache = CacheService.getScriptCache();
    cache.put('MDJ_AUTH_TEST', 'OK', 60);
    out.cacheOk = true;
  } catch (e) {
    throw new Error('No se pudo autorizar CacheService. ' + (e && e.message ? e.message : e));
  }

  try {
    if (typeof Drive !== 'undefined' && Drive.Files && typeof Drive.Files.list === 'function') {
      Drive.Files.list({
        pageSize: 1,
        fields: 'files(id,name)'
      });
      out.driveAdvancedEnabled = true;
      out.driveAdvancedTest = 'Drive API avanzada OK';
    } else {
      out.driveAdvancedEnabled = false;
      out.driveAdvancedTest = 'Drive API avanzada no habilitada en Servicios.';
    }
  } catch (e) {
    out.driveAdvancedEnabled = false;
    out.driveAdvancedTest = 'Drive API avanzada con error: ' + (e && e.message ? e.message : e);
  }

  out.message = 'Pre-autorización completada.';
  Logger.log(JSON.stringify(out, null, 2));
  return out;
}

function mdjVerificarAutorizacionSistema() {
  const result = {
    ok: true,
    checks: {}
  };

  try {
    const ss = SpreadsheetApp.openById(MDJ_APP.hojaControlId);
    result.checks.spreadsheet = { ok: true, name: ss.getName(), id: ss.getId() };
  } catch (e) {
    result.ok = false;
    result.checks.spreadsheet = { ok: false, error: e.message };
  }

  try {
    const root = DriveApp.getRootFolder();
    result.checks.driveApp = { ok: true, name: root.getName() };
  } catch (e) {
    result.ok = false;
    result.checks.driveApp = { ok: false, error: e.message };
  }

  try {
    const props = PropertiesService.getScriptProperties();
    result.checks.properties = { ok: true, sample: props.getProperty('MDJ_AUTH_TEST') || '' };
  } catch (e) {
    result.ok = false;
    result.checks.properties = { ok: false, error: e.message };
  }

  try {
    const cache = CacheService.getScriptCache();
    result.checks.cache = { ok: true, sample: cache.get('MDJ_AUTH_TEST') || '' };
  } catch (e) {
    result.ok = false;
    result.checks.cache = { ok: false, error: e.message };
  }

  try {
    if (typeof Drive !== 'undefined' && Drive.Files && typeof Drive.Files.list === 'function') {
      Drive.Files.list({
        pageSize: 1,
        fields: 'files(id,name)'
      });
      result.checks.driveAdvanced = { ok: true, detail: 'Drive API avanzada OK' };
    } else {
      result.ok = false;
      result.checks.driveAdvanced = { ok: false, error: 'Drive API avanzada no habilitada.' };
    }
  } catch (e) {
    result.ok = false;
    result.checks.driveAdvanced = { ok: false, error: e.message };
  }

  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

/* ============================================================
 * HELPERS GENERALES
 * ============================================================ */

function mdjExtraerSpreadsheetId_(value) {
  const raw = String(value || '').trim();
  if (!raw) throw new Error('No se recibió ningún ID o URL.');

  const match = raw.match(/[-\w]{25,}/);
  if (!match) throw new Error('No se pudo extraer un ID válido desde: ' + raw);

  return match[0];
}

function mdjNormalizeText_(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^A-Za-z0-9_]+/g, '_')
    .replace(/^_+|_+$/g, '')
    .toUpperCase();
}

function mdjNormalizeHeaderName_(value) {
  return mdjNormalizeText_(String(value || '').replace(/\s+/g, '_'));
}

function mdjNormalizeRole_(value) {
  const role = mdjNormalizeText_(value).toLowerCase();
  if (role === 'ADMINISTRADOR') return 'admin';
  if (role === 'ADMIN') return 'admin';
  return 'maestro';
}

function mdjIsTruthy_(value) {
  const normalized = mdjNormalizeText_(value);
  return ['SI', 'TRUE', '1', 'ACTIVO', 'ACTIVA', 'YES'].indexOf(normalized) !== -1;
}

function mdjNormalizeGradeInput_(value) {
  if (value === null || value === undefined) return '';
  const raw = String(value).trim();
  if (raw === '') return '';
  if (raw === '-') return '-';

  const normalized = raw.replace(',', '.');
  if (/^-?\d+(\.\d+)?$/.test(normalized)) return Number(normalized);

  return raw;
}

function mdjToNumberSafe_(value) {
  if (value === null || value === undefined) return null;
  const raw = String(value).trim().replace(',', '.');
  if (!raw || raw === '-') return null;
  if (!/^-?\d+(\.\d+)?$/.test(raw)) return null;
  return Number(raw);
}

function mdjRound2_(value) {
  return Math.round(Number(value || 0) * 100) / 100;
}

function mdjGetBimestresOperativos_() {
  return (MDJ_APP.bimestres || []).filter(function(b) {
    return String(b) !== 'CONSOLIDADO';
  });
}

function mdjBuildStudentKey_(numero, nombre) {
  return mdjNormalizeText_(numero) + '__' + mdjNormalizeText_(nombre);
}

function mdjGetEscalaCurso_(cursoNombre) {
  const cursoNormalizado = mdjNormalizeText_(cursoNombre);
  const configurados = (
    typeof MDJ_APP !== 'undefined' &&
    MDJ_APP &&
    Array.isArray(MDJ_APP.cursosEscala10)
  ) ? MDJ_APP.cursosEscala10 : ['EXPRESION ARTISTICA'];

  const esEscala10 = configurados.some(function(nombre) {
    return mdjNormalizeText_(nombre) === cursoNormalizado;
  });

  return esEscala10 ? 10 : 100;
}

function mdjNormalizarNotaSobre10_(value) {
  const n = mdjToNumberSafe_(value);
  if (n === null) return null;
  if (n <= 10) return mdjRound2_(n);
  if (n <= 100) return mdjRound2_(n / 10);
  return null;
}

function mdjSumarComponentesCurso_(item) {
  const campos = [
    'a1', 'a2', 'a3', 'a4', 'a5', 'a6',
    'apObjetiva', 'evBimestral'
  ];

  let total = 0;
  let cantidad = 0;

  campos.forEach(function(campo) {
    const valor = mdjToNumberSafe_(item && item[campo]);
    if (valor !== null) {
      total += valor;
      cantidad++;
    }
  });

  return {
    tieneDatos: cantidad > 0,
    cantidad: cantidad,
    total: mdjRound2_(total)
  };
}

function mdjCalcularResultadoCurso_(item, cursoNombre) {
  item = item || {};

  const escala = mdjGetEscalaCurso_(cursoNombre);
  const componentes = mdjSumarComponentesCurso_(item);
  const remedial = mdjToNumberSafe_(item.evRemedial);
  const notaHoja = mdjToNumberSafe_(item.notaBimestral);

  let fuente = 'VACIO';
  let totalBase = null;
  let notaFinal = null;
  let advertencia = '';

  // EV. REM. reemplaza la nota cuando existe.
  if (remedial !== null) {
    fuente = 'EV_REMEDIAL';
    totalBase = remedial;
    notaFinal = escala === 10 && remedial > 10
      ? mdjRound2_(remedial / 10)
      : mdjRound2_(remedial);
  } else if (componentes.tieneDatos) {
    // La fuente principal siempre son las actividades visibles.
    fuente = 'SUMATORIA';
    totalBase = componentes.total;
    notaFinal = escala === 10
      ? mdjRound2_(componentes.total / 10)
      : mdjRound2_(componentes.total);
  } else if (notaHoja !== null && notaHoja >= 0 && notaHoja <= escala) {
    // Respaldo únicamente cuando no existen componentes.
    fuente = 'NOTA_BIMESTRAL';
    totalBase = escala === 10 ? mdjRound2_(notaHoja * 10) : notaHoja;
    notaFinal = mdjRound2_(notaHoja);
  }

  if (notaHoja !== null && (notaHoja < 0 || notaHoja > escala)) {
    advertencia = 'La nota almacenada en la hoja (' + notaHoja +
      ') está fuera de la escala ' + escala + ' y fue ignorada.';
  }

  // Protección final contra datos dañados.
  if (notaFinal !== null && notaFinal > escala) {
    advertencia = (advertencia ? advertencia + ' ' : '') +
      'La sumatoria supera la escala; se limitó a ' + escala + '.';
    notaFinal = escala;
  }

  if (notaFinal !== null && notaFinal < 0) notaFinal = 0;

  return {
    fuente: fuente,
    escala: escala,
    total100: totalBase,
    notaFinal: notaFinal,
    nota10: escala === 10 ? notaFinal : null,
    nota100: escala === 100 ? notaFinal : null,
    notaMostrar: notaFinal === null ? '' : Math.round(notaFinal),
    notaHojaOriginal: item.notaBimestral == null ? '' : item.notaBimestral,
    advertencia: advertencia
  };
}

function mdjGetConsolidatedCourseStudents_(libroKey, curso) {
  const ss = mdjGetSpreadsheetByKey_(libroKey);
  const mapa = {};
  const escala = mdjGetEscalaCurso_(curso);

  mdjGetBimestresOperativos_().forEach(function(bimestre) {
    const sh = ss.getSheetByName(bimestre);
    if (!sh) return;

    const courseMeta = mdjFindCourseInSheet_(sh, curso);
    if (!courseMeta) return;

    const students = mdjExtractStudentsForCourse_(sh, courseMeta);

    students.forEach(function(st) {
      const key = mdjBuildStudentKey_(st.numero, st.nombre);

      if (!mapa[key]) {
        mapa[key] = {
          rowNumber: st.rowNumber,
          numero: st.numero,
          nombre: st.nombre,
          iBimestre: '',
          iiBimestre: '',
          iiiBimestre: '',
          ivBimestre: ''
        };
      }

      const nota = st.notaFinalNumerica !== '' && st.notaFinalNumerica !== null
        ? Number(st.notaFinalNumerica)
        : null;

      if (bimestre === 'I BIMESTRE') mapa[key].iBimestre = nota === null ? '' : nota;
      if (bimestre === 'II BIMESTRE') mapa[key].iiBimestre = nota === null ? '' : nota;
      if (bimestre === 'III BIMESTRE') mapa[key].iiiBimestre = nota === null ? '' : nota;
      if (bimestre === 'IV BIMESTRE') mapa[key].ivBimestre = nota === null ? '' : nota;
    });
  });

  return Object.keys(mapa).map(function(key) {
    const item = mapa[key];
    const notas = [
      item.iBimestre,
      item.iiBimestre,
      item.iiiBimestre,
      item.ivBimestre
    ].filter(function(n) {
      return n !== '' && n !== null && n !== undefined && !isNaN(Number(n));
    }).map(Number);

    const promedio = notas.length
      ? mdjRound2_(notas.reduce(function(a, b) { return a + b; }, 0) / notas.length)
      : '';

    item.escala = escala;
    item.promedio = promedio;
    item.promedioNumerico = promedio;
    item.promedioMostrar = promedio === '' ? '' : Math.round(promedio);
    return item;
  }).sort(function(a, b) {
    return Number(a.rowNumber || 0) - Number(b.rowNumber || 0);
  });
}

function mdjBuildHeaderMap_(headerRow) {
  const map = {};
  (headerRow || []).forEach(function(name, index) {
    const key = mdjNormalizeHeaderName_(name);
    if (key) map[key] = index;
  });

  const requiredAliases = {
    usuario: ['usuario', 'user', 'correo'],
    password: ['password', 'contrasena', 'contraseña', 'clave'],
    nombre: ['nombre', 'nombre_completo'],
    rol: ['rol', 'role'],
    activo: ['activo', 'estado', 'habilitado'],
    libro_key: ['libro_key', 'libro', 'grado_key'],
    curso: ['curso', 'materia']
  };

  const finalMap = {};

  Object.keys(requiredAliases).forEach(function(target) {
    const aliases = requiredAliases[target].map(function(alias) {
      return mdjNormalizeHeaderName_(alias);
    });

    const found = aliases.find(function(alias) {
      return Object.prototype.hasOwnProperty.call(map, alias);
    });

    finalMap[target] = found ? map[found] : -1;
  });

  return finalMap;
}

function mdjEsNombreAlumnoValido_(nombre) {
  nombre = String(nombre || '').trim();
  if (!nombre) return false;

  // No cargar errores de fórmula como si fueran alumnos.
  if (/^#(ERROR|REF|N\/A|VALUE|NAME|DIV\/0|NUM|NULL)/i.test(nombre)) {
    return false;
  }

  const normalizado = mdjNormalizeText_(nombre);
  if (!normalizado) return false;

  const descartes = [
    'PROMEDIO',
    'PROMEDIOS',
    'TOTALES',
    'TOTAL',
    'ALUMNO',
    'ALUMNOS',
    'NOMBRE',
    'NOMBRE_DEL_ALUMNO',
    'ESTUDIANTES'
  ];

  return descartes.indexOf(normalizado) === -1;
}

/* ============================================================
 * CONTROL SHEET
 * ============================================================ */

function mdjGetControlSpreadsheet_() {
  const id = String(MDJ_APP.hojaControlId || '').trim();
  if (!id) {
    throw new Error('Debes configurar MDJ_APP.hojaControlId con el ID real de la hoja de control.');
  }
  return SpreadsheetApp.openById(id);
}

function mdjGetControlSheet_(sheetName) {
  const ss = mdjGetControlSpreadsheet_();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) {
    throw new Error('No existe la hoja "' + sheetName + '" dentro de la hoja de control.');
  }
  return sh;
}

function mdjInicializarHojaControlExistente() {
  const ss = mdjGetControlSpreadsheet_();

  let usuarios = ss.getSheetByName(MDJ_APP.hojaUsuarios);
  if (!usuarios) usuarios = ss.insertSheet(MDJ_APP.hojaUsuarios);

  let asignaciones = ss.getSheetByName(MDJ_APP.hojaAsignaciones);
  if (!asignaciones) asignaciones = ss.insertSheet(MDJ_APP.hojaAsignaciones);

  usuarios.clear();
  usuarios.getRange(1, 1, 3, 5).setValues([
    ['usuario', 'password', 'nombre', 'rol', 'activo'],
    ['admin', 'admin123', 'Administrador General', 'admin', 'SI'],
    ['maestro1', 'maestro123', 'Maestro Ejemplo', 'maestro', 'SI']
  ]);

  asignaciones.clear();
  asignaciones.getRange(1, 1, 3, 4).setValues([
    ['usuario', 'libro_key', 'curso', 'activo'],
    ['maestro1', 'primero_basico', '*', 'SI'],
    ['maestro1', 'segundo_basico', '*', 'SI']
  ]);

  usuarios.autoResizeColumns(1, 5);
  asignaciones.autoResizeColumns(1, 4);

  return {
    ok: true,
    spreadsheetId: ss.getId(),
    spreadsheetUrl: ss.getUrl(),
    sheets: ss.getSheets().map(function(sh) { return sh.getName(); }),
    message: 'Hoja de control inicializada correctamente.'
  };
}

/* ============================================================
 * DRIVE / CONVERSIÓN
 * ============================================================ */

function mdjEsMimeGoogleSheet_(mimeType) {
  return mimeType === MDJ_MIME_GOOGLE_SHEETS;
}

function mdjEsMimeExcel_(mimeType) {
  return mimeType === MDJ_MIME_XLS || mimeType === MDJ_MIME_XLSX;
}

function mdjGetDriveFileMetaById_(fileId) {
  const meta = {
    ok: false,
    id: fileId,
    name: '',
    mimeType: '',
    url: '',
    error: ''
  };

  try {
    const file = DriveApp.getFileById(fileId);
    meta.ok = true;
    meta.name = file.getName();
    meta.mimeType = file.getMimeType();
    meta.url = file.getUrl();
    return meta;
  } catch (e) {
    meta.error = e && e.message ? e.message : String(e);
    return meta;
  }
}

function mdjGetConvertedPropertyKey_(sourceId) {
  return MDJ_CONVERTED_PREFIX + sourceId;
}

function mdjGetConvertedSpreadsheetId_(sourceId) {
  return PropertiesService.getScriptProperties().getProperty(
    mdjGetConvertedPropertyKey_(sourceId)
  ) || '';
}

function mdjSetConvertedSpreadsheetId_(sourceId, targetId) {
  PropertiesService.getScriptProperties().setProperty(
    mdjGetConvertedPropertyKey_(sourceId),
    targetId
  );
}

function mdjDeleteConvertedSpreadsheetId_(sourceId) {
  PropertiesService.getScriptProperties().deleteProperty(
    mdjGetConvertedPropertyKey_(sourceId)
  );
}

function mdjGetParentFolderIdSafe_(file) {
  try {
    const parents = file.getParents();
    if (parents.hasNext()) return parents.next().getId();
  } catch (e) {}
  return '';
}

function mdjVerificarDriveAvanzado_() {
  if (typeof Drive === 'undefined' || !Drive.Files) {
    throw new Error(
      'La API avanzada de Drive no está habilitada. Ve a Servicios > + > Drive API y habilítala también en Google Cloud.'
    );
  }
}

function mdjInsertArchivoConvertido_(resource, blob) {
  mdjVerificarDriveAvanzado_();

  if (typeof Drive.Files.create === 'function') {
    return Drive.Files.create(resource, blob, {
      fields: 'id,name,mimeType'
    });
  }

  if (typeof Drive.Files.insert === 'function') {
    return Drive.Files.insert(resource, blob, {
      convert: true
    });
  }

  throw new Error('La API avanzada de Drive no expone ni Drive.Files.create ni Drive.Files.insert.');
}

function mdjConvertExcelToGoogleSheet_(sourceId, libroNombre, forceNew) {
  forceNew = !!forceNew;

  const cachedId = mdjGetConvertedSpreadsheetId_(sourceId);
  if (cachedId && !forceNew) {
    try {
      const cachedSs = SpreadsheetApp.openById(cachedId);
      mdjLog_('INFO', 'Usando conversión existente', {
        sourceId: sourceId,
        convertedId: cachedId,
        spreadsheetName: cachedSs.getName()
      });
      return {
        ok: true,
        reused: true,
        convertedId: cachedId,
        spreadsheet: cachedSs
      };
    } catch (e) {
      mdjLog_('WARN', 'Conversión cacheada inválida, se reconvertirá', {
        sourceId: sourceId,
        convertedId: cachedId,
        error: e && e.message ? e.message : String(e)
      });
    }
  }

  const file = DriveApp.getFileById(sourceId);
  const blob = file.getBlob();
  const originalName = String(file.getName() || libroNombre || 'Libro').trim();
  const cleanName = originalName.replace(/\.(xlsx|xls)$/i, '').trim();
  const targetTitle = cleanName + ' [GS]';

  const resource = {
    name: targetTitle,
    mimeType: MDJ_MIME_GOOGLE_SHEETS
  };

  const parentId = mdjGetParentFolderIdSafe_(file);
  if (parentId) {
    resource.parents = [parentId];
  }

  mdjLog_('INFO', 'Convirtiendo archivo Excel', {
    sourceId: sourceId,
    originalName: originalName,
    targetTitle: targetTitle,
    parentId: parentId || ''
  });

  const inserted = mdjInsertArchivoConvertido_(resource, blob);
  const newId = inserted.id || '';

  if (!newId) {
    throw new Error('No se obtuvo el ID del archivo convertido.');
  }

  mdjSetConvertedSpreadsheetId_(sourceId, newId);

  const ss = SpreadsheetApp.openById(newId);

  mdjLog_('INFO', 'Conversión completada', {
    sourceId: sourceId,
    convertedId: newId,
    spreadsheetName: ss.getName()
  });

  return {
    ok: true,
    reused: false,
    convertedId: newId,
    spreadsheet: ss
  };
}

function mdjGarantizarLibroComoGoogleSheet_(libroKey, forceNew) {
  const libro = mdjGetLibroByKey(libroKey);
  if (!libro) throw new Error('No se encontró el libro con key: ' + libroKey);

  const sourceId = mdjExtraerSpreadsheetId_(libro.driveId || libro.spreadsheetId || libro.url || '');
  const meta = mdjGetDriveFileMetaById_(sourceId);

  if (!meta.ok) {
    throw new Error('No se pudo acceder al archivo de "' + libro.nombre + '": ' + meta.error);
  }

  if (mdjEsMimeGoogleSheet_(meta.mimeType)) {
    mdjSetConvertedSpreadsheetId_(sourceId, sourceId);
    return {
      ok: true,
      libroKey: libro.key,
      libroNombre: libro.nombre,
      sourceId: sourceId,
      targetId: sourceId,
      reused: true,
      converted: false,
      alreadyGoogleSheet: true,
      fileName: meta.name,
      mimeType: meta.mimeType,
      spreadsheetName: meta.name
    };
  }

  if (mdjEsMimeExcel_(meta.mimeType)) {
    const result = mdjConvertExcelToGoogleSheet_(sourceId, libro.nombre, forceNew);
    const ss = result.spreadsheet;

    return {
      ok: true,
      libroKey: libro.key,
      libroNombre: libro.nombre,
      sourceId: sourceId,
      targetId: result.convertedId,
      reused: result.reused,
      converted: true,
      alreadyGoogleSheet: false,
      fileName: meta.name,
      mimeType: meta.mimeType,
      spreadsheetName: ss.getName()
    };
  }

  throw new Error('El archivo "' + meta.name + '" no es compatible. MIME detectado: ' + meta.mimeType);
}

function mdjConvertirTodosLosLibrosAGoogleSheets(forceNew) {
  forceNew = !!forceNew;

  const resultados = [];
  let okCount = 0;
  let failCount = 0;

  mdjLog_('INFO', 'Inicio conversión masiva', {
    totalLibros: MDJ_LIBROS.length,
    forceNew: forceNew
  });

  MDJ_LIBROS.forEach(function(libro, index) {
    const item = {
      index: index + 1,
      total: MDJ_LIBROS.length,
      libroKey: libro.key,
      libroNombre: libro.nombre,
      ok: false,
      sourceId: '',
      targetId: '',
      converted: false,
      reused: false,
      alreadyGoogleSheet: false,
      fileName: '',
      mimeType: '',
      spreadsheetName: '',
      error: ''
    };

    try {
      const result = mdjGarantizarLibroComoGoogleSheet_(libro.key, forceNew);
      item.ok = true;
      item.sourceId = result.sourceId || '';
      item.targetId = result.targetId || '';
      item.converted = !!result.converted;
      item.reused = !!result.reused;
      item.alreadyGoogleSheet = !!result.alreadyGoogleSheet;
      item.fileName = result.fileName || '';
      item.mimeType = result.mimeType || '';
      item.spreadsheetName = result.spreadsheetName || '';
      okCount++;
      mdjLog_('INFO', 'Libro procesado', item);
    } catch (e) {
      item.error = e && e.message ? e.message : String(e);
      failCount++;
      mdjLog_('ERROR', 'Error al convertir libro', item);
    }

    resultados.push(item);
  });

  const salida = {
    ok: failCount === 0,
    total: MDJ_LIBROS.length,
    procesados: okCount,
    errores: failCount,
    resultados: resultados
  };

  mdjLog_('INFO', 'Fin conversión masiva', salida);
  return salida;
}

function mdjVerLibrosConvertidos() {
  return MDJ_LIBROS.map(function(libro) {
    let sourceId = '';
    let targetId = '';
    let metaSource = null;
    let metaTarget = null;
    let error = '';

    try {
      sourceId = mdjExtraerSpreadsheetId_(libro.driveId || libro.spreadsheetId || libro.url || '');
      targetId = mdjGetConvertedSpreadsheetId_(sourceId);
      metaSource = mdjGetDriveFileMetaById_(sourceId);
      if (targetId) {
        metaTarget = mdjGetDriveFileMetaById_(targetId);
      }
    } catch (e) {
      error = e && e.message ? e.message : String(e);
    }

    return {
      libroKey: libro.key,
      libroNombre: libro.nombre,
      sourceId: sourceId,
      sourceName: metaSource && metaSource.name || '',
      sourceMimeType: metaSource && metaSource.mimeType || '',
      targetId: targetId,
      targetName: metaTarget && metaTarget.name || '',
      targetMimeType: metaTarget && metaTarget.mimeType || '',
      error: error
    };
  });
}

function mdjGenerarBloqueConfigConIdsConvertidos() {
  const lineas = [];
  lineas.push('const MDJ_LIBROS = [');

  MDJ_LIBROS.forEach(function(libro, idx) {
    const sourceId = mdjExtraerSpreadsheetId_(libro.driveId || libro.spreadsheetId || libro.url || '');
    const targetId = mdjGetConvertedSpreadsheetId_(sourceId) || sourceId;

    lineas.push('  {');
    lineas.push("    key: '" + libro.key + "',");
    lineas.push("    nombre: '" + String(libro.nombre || '').replace(/'/g, "\\'") + "',");
    lineas.push("    nivel: '" + String(libro.nivel || '').replace(/'/g, "\\'") + "',");
    lineas.push("    grado: '" + String(libro.grado || '').replace(/'/g, "\\'") + "',");
    lineas.push("    programa: '" + String(libro.programa || '').replace(/'/g, "\\'") + "',");
    lineas.push("    driveId: '" + targetId + "'");
    lineas.push(idx === MDJ_LIBROS.length - 1 ? '  }' : '  },');
  });

  lineas.push('];');

  const texto = lineas.join('\n');
  Logger.log(texto);

  return {
    ok: true,
    texto: texto
  };
}

function mdjLimpiarCacheConversiones() {
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  let deleted = 0;

  Object.keys(all).forEach(function(key) {
    if (key.indexOf(MDJ_CONVERTED_PREFIX) === 0) {
      props.deleteProperty(key);
      deleted++;
    }
  });

  return {
    ok: true,
    eliminadas: deleted
  };
}

function mdjGetSpreadsheetByKey_(libroKey) {
  const libro = mdjGetLibroByKey(libroKey);
  if (!libro) throw new Error('No se encontró el libro con key: ' + libroKey);

  const sourceId = mdjExtraerSpreadsheetId_(libro.driveId || libro.spreadsheetId || libro.url || '');
  const convertedId = mdjGetConvertedSpreadsheetId_(sourceId);
  const finalId = convertedId || sourceId;
  const meta = mdjGetDriveFileMetaById_(finalId);

  mdjLog_('INFO', 'mdjGetSpreadsheetByKey_', {
    libroKey: libroKey,
    libroNombre: libro.nombre,
    sourceId: sourceId,
    convertedId: convertedId,
    finalId: finalId,
    metaOk: meta.ok,
    fileName: meta.name,
    mimeType: meta.mimeType,
    metaError: meta.error
  });

  if (!meta.ok) {
    throw new Error('No se pudo acceder al archivo de "' + libro.nombre + '" (ID: ' + finalId + '). ' + meta.error);
  }

  if (!mdjEsMimeGoogleSheet_(meta.mimeType)) {
    if (mdjEsMimeExcel_(meta.mimeType)) {
      const conversion = mdjConvertExcelToGoogleSheet_(sourceId, libro.nombre, false);
      return conversion.spreadsheet;
    }
    throw new Error('El archivo final de "' + libro.nombre + '" no es Google Sheets nativo. MIME detectado: ' + meta.mimeType);
  }

  return SpreadsheetApp.openById(finalId);
}

function mdjDiagnosticarAccesoLibroPorKey(libroKey) {
  const libro = mdjGetLibroByKey(libroKey);
  if (!libro) {
    return {
      ok: false,
      libroKey: libroKey,
      error: 'Libro no encontrado.'
    };
  }

  const sourceId = mdjExtraerSpreadsheetId_(libro.driveId || libro.spreadsheetId || libro.url || '');
  const driveMeta = mdjGetDriveFileMetaById_(sourceId);
  const convertedId = mdjGetConvertedSpreadsheetId_(sourceId);

  const out = {
    ok: false,
    libroKey: libro.key,
    libroNombre: libro.nombre,
    sourceId: sourceId,
    convertedId: convertedId,
    driveMeta: driveMeta,
    spreadsheetName: '',
    totalSheets: 0,
    sheets: [],
    error: ''
  };

  try {
    const ss = mdjGetSpreadsheetByKey_(libro.key);
    out.ok = true;
    out.spreadsheetName = ss.getName();
    out.totalSheets = ss.getSheets().length;
    out.sheets = ss.getSheets().map(function(sh) { return sh.getName(); });
  } catch (e) {
    out.error = e && e.message ? e.message : String(e);
  }

  mdjLog_('INFO', 'mdjDiagnosticarAccesoLibroPorKey resultado', out);
  return out;
}

function mdjDiagnosticarAccesoTodosLosLibros() {
  return MDJ_LIBROS.map(function(libro) {
    return mdjDiagnosticarAccesoLibroPorKey(libro.key);
  });
}

/* ============================================================
 * PARSER REAL SEGÚN TUS HOJAS
 * ============================================================ */

function mdjExtractCoursesFromSheet_(sheet) {
  const lastCol = sheet.getLastColumn();
  const headerRow = sheet.getRange(4, 1, 1, lastCol).getDisplayValues()[0];
  const courses = [];

  for (let col = 3; col <= lastCol; col++) {
    const value = String(headerRow[col - 1] || '').trim();
    if (!value) continue;

    const normalized = mdjNormalizeText_(value);
    if (!normalized) continue;
    if (normalized === 'CONSOLIDADO') continue;

    courses.push({
      name: value,
      startCol: col,
      endCol: col + 9,
      columns: {
        a1: col,
        a2: col + 1,
        a3: col + 2,
        a4: col + 3,
        a5: col + 4,
        a6: col + 5,
        apObjetiva: col + 6,
        evBimestral: col + 7,
        evRemedial: col + 8,
        notaBimestral: col + 9
      }
    });

    col += 9;
  }

  return courses;
}

function mdjFindCourseInSheet_(sheet, cursoBuscado) {
  const courses = mdjExtractCoursesFromSheet_(sheet);
  return courses.find(function(item) {
    return mdjNormalizeText_(item.name) === mdjNormalizeText_(cursoBuscado);
  }) || null;
}

function mdjExtractStudentsForCourse_(sheet, courseMeta) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 7) return [];

  const values = sheet
    .getRange(7, 1, lastRow - 6, Math.max(2, courseMeta.endCol))
    .getDisplayValues();

  const students = [];

  values.forEach(function(row, index) {
    const rowNumber = index + 7;
    const numero = String(row[0] || '').trim();
    const nombre = String(row[1] || '').trim();

    if (!mdjEsNombreAlumnoValido_(nombre)) return;

    const item = {
      rowNumber: rowNumber,
      numero: numero,
      nombre: nombre,
      a1: row[courseMeta.startCol - 1] || '',
      a2: row[courseMeta.startCol] || '',
      a3: row[courseMeta.startCol + 1] || '',
      a4: row[courseMeta.startCol + 2] || '',
      a5: row[courseMeta.startCol + 3] || '',
      a6: row[courseMeta.startCol + 4] || '',
      apObjetiva: row[courseMeta.startCol + 5] || '',
      evBimestral: row[courseMeta.startCol + 6] || '',
      evRemedial: row[courseMeta.startCol + 7] || '',
      notaBimestral: row[courseMeta.startCol + 8] || ''
    };

    const notaBimestralHoja = item.notaBimestral;
    const resultado = mdjCalcularResultadoCurso_(item, courseMeta.name);

    item.resultado = resultado;
    item.escala = resultado.escala;
    item.totalCurso = resultado.total100 === null ? '' : resultado.total100;
    item.notaBimestralHoja = notaBimestralHoja;
    item.notaFinalNumerica = resultado.notaFinal === null ? '' : resultado.notaFinal;
    item.notaCalculada10 = resultado.nota10 === null ? '' : resultado.nota10;
    item.notaCalculada100 = resultado.nota100 === null ? '' : resultado.nota100;
    item.notaMostrar = resultado.notaMostrar;
    item.advertenciaNota = resultado.advertencia || '';

    // Compatibilidad con los HTML actuales, que imprimen item.notaBimestral.
    item.notaBimestral = resultado.notaMostrar;

    students.push(item);
  });

  return students;
}

function mdjGetBookSummary_(libroKey) {
  const libro = mdjGetLibroByKey(libroKey);
  if (!libro) return { totalCursos: 0, totalAlumnos: 0 };

  const ss = mdjGetSpreadsheetByKey_(libroKey);
  const sh = ss.getSheetByName('I BIMESTRE');
  if (!sh) return { totalCursos: 0, totalAlumnos: 0 };

  const courses = mdjExtractCoursesFromSheet_(sh);
  if (!courses.length) return { totalCursos: 0, totalAlumnos: 0 };

  const students = mdjExtractStudentsForCourse_(sh, courses[0]);

  return {
    totalCursos: courses.length,
    totalAlumnos: students.length
  };
}

/* ============================================================
 * ASIGNACIONES
 * ============================================================ */

function mdjGetAssignmentsForUser_(usuario) {
  const sh = mdjGetControlSheet_(MDJ_APP.hojaAsignaciones);
  const values = sh.getDataRange().getDisplayValues();
  if (!values || values.length < 2) return [];

  const header = mdjBuildHeaderMap_(values[0]);
  const usuarioBuscado = mdjNormalizeText_(usuario);

  return values.slice(1).filter(function(row) {
    const activo = mdjIsTruthy_(row[header.activo]);
    const usuarioFila = mdjNormalizeText_(row[header.usuario]);
    return activo && usuarioFila === usuarioBuscado;
  }).map(function(row) {
    const libroKey = String(row[header.libro_key] || '').trim();
    const curso = String(row[header.curso] || '').trim();
    const libro = mdjGetLibroByKey(libroKey);

    return {
      usuario: String(row[header.usuario] || ''),
      libroKey: libroKey,
      curso: curso,
      libroNombre: libro ? libro.nombre : libroKey,
      label: (libro ? libro.nombre : libroKey) + ' · ' + (curso || '*')
    };
  });
}

function mdjAssertCourseAccess_(user, libroKey, curso) {
  const role = String(user && user.rol || '').toLowerCase();
  if (role === 'admin') return true;

  const assignments = mdjGetAssignmentsForUser_(user.usuario);
  const ok = assignments.some(function(item) {
    const sameBook = mdjResolveLibroKey_(item.libroKey) === mdjResolveLibroKey_(libroKey);
    const sameCourse = item.curso === '*' || mdjNormalizeText_(item.curso) === mdjNormalizeText_(curso);
    return sameBook && sameCourse;
  });

  if (!ok) {
    throw new Error('No tienes permisos para acceder a este curso.');
  }

  return true;
}

/* ============================================================
 * BOOTSTRAPS
 * ============================================================ */

function mdjGetAdminBootstrap(token) {
  mdjLog_('INFO', 'mdjGetAdminBootstrap');

  try {
    const sess = mdjGetSession(token);
    if (!sess.ok) {
      return { ok: false, libros: [], bimestres: [], message: sess.message };
    }

    if (String(sess.user.rol || '').toLowerCase() !== 'admin') {
      return { ok: false, libros: [], bimestres: [], message: 'Acceso denegado.' };
    }

    const libros = MDJ_LIBROS.map(function(libro) {
      const info = {
        key: libro.key,
        nombre: libro.nombre,
        nivel: libro.nivel,
        grado: libro.grado,
        programa: libro.programa,
        totalCursos: 0,
        totalAlumnos: 0,
        ok: true,
        error: ''
      };

      try {
        const resumen = mdjGetBookSummary_(libro.key);
        info.totalCursos = Number(resumen.totalCursos || 0);
        info.totalAlumnos = Number(resumen.totalAlumnos || 0);
      } catch (e) {
        info.ok = false;
        info.error = e && e.message ? e.message : String(e);
        mdjLog_('ERROR', 'Error libro bootstrap: ' + libro.key, info);
      }

      return info;
    });

    return {
      ok: true,
      user: sess.user,
      libros: libros,
      bimestres: MDJ_APP.bimestres.filter(function(b) { return b !== 'CONSOLIDADO'; }),
      message: 'OK'
    };
  } catch (e) {
    mdjLog_('ERROR', 'mdjGetAdminBootstrap fatal', e && e.message ? e.message : String(e));
    return {
      ok: false,
      libros: [],
      bimestres: [],
      message: e && e.message ? e.message : String(e)
    };
  }
}

function mdjGetTeacherBootstrap(token) {
  mdjLog_('INFO', 'mdjGetTeacherBootstrap');

  try {
    const sess = mdjGetSession(token);
    if (!sess.ok) {
      return { ok: false, asignaciones: [], bimestres: [], message: sess.message };
    }

    const role = String(sess.user.rol || '').toLowerCase();
    if (role !== 'maestro' && role !== 'admin') {
      return { ok: false, asignaciones: [], bimestres: [], message: 'Acceso denegado.' };
    }

    const asignaciones = role === 'admin'
      ? []
      : mdjGetAssignmentsForUser_(sess.user.usuario);

    return {
      ok: true,
      user: sess.user,
      asignaciones: asignaciones,
      bimestres: MDJ_APP.bimestres.filter(function(b) { return b !== 'CONSOLIDADO'; }),
      message: 'OK'
    };
  } catch (e) {
    return {
      ok: false,
      asignaciones: [],
      bimestres: [],
      message: e && e.message ? e.message : String(e)
    };
  }
}

/* ============================================================
 * CURSOS / ALUMNOS / GUARDADO
 * ============================================================ */

function mdjGetCoursesByBook(token, libroKey, bimestre) {
  const sess = mdjGetSession(token);
  if (!sess.ok) return { ok: false, cursos: [], message: sess.message };

  const ss = mdjGetSpreadsheetByKey_(libroKey);

  // Para consolidado, tomamos la estructura de cursos desde I BIMESTRE
  if (String(bimestre || '').trim() === 'CONSOLIDADO') {
    const shRef = ss.getSheetByName('I BIMESTRE') || ss.getSheets()[0];
    if (!shRef) return { ok: false, cursos: [], message: 'No hay hojas disponibles.' };

    const courses = mdjExtractCoursesFromSheet_(shRef);
    return {
      ok: true,
      libro: mdjGetLibroByKey(libroKey),
      cursos: courses
    };
  }

  const sh = ss.getSheetByName(bimestre);
  if (!sh) return { ok: false, cursos: [], message: 'No existe la hoja ' + bimestre + '.' };

  const courses = mdjExtractCoursesFromSheet_(sh);
  return {
    ok: true,
    libro: mdjGetLibroByKey(libroKey),
    cursos: courses
  };
}

function mdjGetCourseStudents(token, libroKey, bimestre, curso) {
  const sess = mdjGetSession(token);
  if (!sess.ok) return { ok: false, alumnos: [], message: sess.message };

  mdjAssertCourseAccess_(sess.user, libroKey, curso);

  if (String(bimestre || '').trim() === 'CONSOLIDADO') {
    const alumnos = mdjGetConsolidatedCourseStudents_(libroKey, curso);
    return {
      ok: true,
      libro: mdjGetLibroByKey(libroKey),
      bimestre: 'CONSOLIDADO',
      curso: { name: curso },
      alumnos: alumnos
    };
  }

  const ss = mdjGetSpreadsheetByKey_(libroKey);
  const sh = ss.getSheetByName(bimestre);
  if (!sh) return { ok: false, alumnos: [], message: 'No existe la hoja ' + bimestre + '.' };

  const courseMeta = mdjFindCourseInSheet_(sh, curso);
  if (!courseMeta) {
    return { ok: false, alumnos: [], message: 'No se encontró el curso "' + curso + '" en ' + bimestre + '.' };
  }

  const students = mdjExtractStudentsForCourse_(sh, courseMeta);

  return {
    ok: true,
    libro: mdjGetLibroByKey(libroKey),
    bimestre: bimestre,
    curso: courseMeta,
    alumnos: students
  };
}

function mdjSaveCourseGrades(token, payload) {
  const sess = mdjGetSession(token);
  if (!sess.ok) return { ok: false, message: sess.message };

  payload = payload || {};
  const libroKey = String(payload.libroKey || '').trim();
  const bimestre = String(payload.bimestre || '').trim();
  const curso = String(payload.curso || '').trim();
  const alumnos = Array.isArray(payload.alumnos) ? payload.alumnos : [];

  if (!libroKey || !bimestre || !curso) {
    return { ok: false, message: 'Faltan datos para guardar notas.' };
  }

  if (bimestre === 'CONSOLIDADO') {
    return { ok: false, message: 'El consolidado se genera automáticamente; no se guarda manualmente.' };
  }

  mdjAssertCourseAccess_(sess.user, libroKey, curso);

  const ss = mdjGetSpreadsheetByKey_(libroKey);
  const sh = ss.getSheetByName(bimestre);
  if (!sh) return { ok: false, message: 'No se encontró la hoja ' + bimestre + '.' };

  const courseMeta = mdjFindCourseInSheet_(sh, curso);
  if (!courseMeta) return { ok: false, message: 'No se encontró el curso "' + curso + '".' };

  alumnos.forEach(function(item) {
    const row = Number(item.rowNumber || 0);
    if (!row) return;

    sh.getRange(row, courseMeta.startCol, 1, 9).setValues([[
      mdjNormalizeGradeInput_(item.a1),
      mdjNormalizeGradeInput_(item.a2),
      mdjNormalizeGradeInput_(item.a3),
      mdjNormalizeGradeInput_(item.a4),
      mdjNormalizeGradeInput_(item.a5),
      mdjNormalizeGradeInput_(item.a6),
      mdjNormalizeGradeInput_(item.apObjetiva),
      mdjNormalizeGradeInput_(item.evBimestral),
      mdjNormalizeGradeInput_(item.evRemedial)
    ]]);
  });

  SpreadsheetApp.flush();

  return {
    ok: true,
    message: 'Notas guardadas correctamente.',
    alumnos: mdjExtractStudentsForCourse_(sh, courseMeta)
  };
}

/* ============================================================
 * DIAGNÓSTICOS
 * ============================================================ */

function mdjDiagnosticoLogin() {
  try {
    const ss = mdjGetControlSpreadsheet_();
    const usuarios = ss.getSheetByName(MDJ_APP.hojaUsuarios);
    const asignaciones = ss.getSheetByName(MDJ_APP.hojaAsignaciones);

    const data = usuarios ? usuarios.getDataRange().getDisplayValues() : [];
    const usuariosEncontrados = data.slice(1).filter(function(r) {
      return String(r[0] || '').trim() !== '';
    }).map(function(r) {
      return {
        usuario: String(r[0] || ''),
        nombre: String(r[2] || ''),
        rol: String(r[3] || ''),
        activo: String(r[4] || '')
      };
    });

    return {
      ok: true,
      spreadsheetId: ss.getId(),
      spreadsheetName: ss.getName(),
      hojaUsuariosExiste: !!usuarios,
      hojaAsignacionesExiste: !!asignaciones,
      totalUsuarios: usuariosEncontrados.length,
      usuarios: usuariosEncontrados
    };
  } catch (error) {
    return {
      ok: false,
      message: error.message
    };
  }
}

/* ============================================================
 * CARPETAS DE BOLETAS
 * ============================================================ */

function mdjCrearEstructuraBoletasLibro(token, libroKey) {
  mdjLog_('INFO', 'mdjCrearEstructuraBoletasLibro', { libroKey: libroKey });

  try {
    const sess = mdjGetSession(token);
    if (!sess.ok) return { ok: false, message: sess.message };

    const folderId = MDJ_BOLETAS_DRIVE.foldersByLibro[libroKey];
    if (!folderId) return { ok: false, message: 'No hay carpeta Drive configurada para ' + libroKey };

    const rootFolder = DriveApp.getFolderById(folderId);
    const ss = mdjGetSpreadsheetByKey_(libroKey);
    const sh = ss.getSheetByName('I BIMESTRE');
    if (!sh) return { ok: false, message: 'Hoja I BIMESTRE no encontrada.' };

    const lastRow = sh.getLastRow();
    if (lastRow < 7) {
      return {
        ok: true,
        totalAlumnos: 0,
        folderUrl: rootFolder.getUrl(),
        message: 'No hay alumnos para procesar.'
      };
    }

    const data = sh.getRange(7, 1, lastRow - 6, 2).getDisplayValues();
    let count = 0;

    data.forEach(function(row) {
      const nombre = String(row[1] || '').trim();
      if (!mdjEsNombreAlumnoValido_(nombre)) return;

      const ex = rootFolder.getFoldersByName(nombre);
      const alumnoFolder = ex.hasNext() ? ex.next() : rootFolder.createFolder(nombre);

      MDJ_BOLETAS_DRIVE.subcarpetasAlumno.forEach(function(sub) {
        const s = alumnoFolder.getFoldersByName(sub);
        if (!s.hasNext()) alumnoFolder.createFolder(sub);
      });

      count++;
    });

    return {
      ok: true,
      totalAlumnos: count,
      folderUrl: rootFolder.getUrl(),
      message: 'Estructura lista.'
    };
  } catch (e) {
    mdjLog_('ERROR', 'mdjCrearEstructuraBoletasLibro error', e && e.message ? e.message : String(e));
    return {
      ok: false,
      message: e && e.message ? e.message : String(e)
    };
  }
}

/* ============================================================
 * DIAGNÓSTICO DE NOTAS POR CURSO
 * ============================================================ */
function mdjDiagnosticarCursoNotas_(libroKey, bimestre, curso) {
  const ss = mdjGetSpreadsheetByKey_(libroKey);
  const sh = ss.getSheetByName(bimestre);
  if (!sh) throw new Error('No existe la hoja ' + bimestre + '.');

  const meta = mdjFindCourseInSheet_(sh, curso);
  if (!meta) throw new Error('No se encontró el curso "' + curso + '".');

  const alumnos = mdjExtractStudentsForCourse_(sh, meta);

  return {
    ok: true,
    libroKey: libroKey,
    bimestre: bimestre,
    curso: curso,
    escala: mdjGetEscalaCurso_(curso),
    totalAlumnos: alumnos.length,
    muestra: alumnos.slice(0, 20).map(function(a) {
      return {
        fila: a.rowNumber,
        numero: a.numero,
        nombre: a.nombre,
        componentes: {
          a1: a.a1, a2: a.a2, a3: a.a3, a4: a.a4,
          a5: a.a5, a6: a.a6,
          apObjetiva: a.apObjetiva,
          evBimestral: a.evBimestral,
          evRemedial: a.evRemedial
        },
        notaHoja: a.notaBimestralHoja,
        notaCalculada: a.notaMostrar,
        fuente: a.resultado && a.resultado.fuente,
        advertencia: a.advertenciaNota
      };
    })
  };
}

/** Prueba directa desde el editor de Apps Script. */
function mdjPruebaInglesCuartoSecretariado() {
  const resultado = mdjDiagnosticarCursoNotas_(
    'cuarto_secretariado',
    'II BIMESTRE',
    'INGLES'
  );
  Logger.log(JSON.stringify(resultado, null, 2));
  return resultado;
}
