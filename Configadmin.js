/**
 * CONFIG ADMIN - CORRECCIONES
 * Reemplaza en tu configadmin.gs estas funciones
 */

/**
 * Construye un mapa de encabezados tolerante a variantes.
 */
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

/**
 * Normaliza nombres de encabezado.
 */
function mdjNormalizeHeaderName_(value) {
  return mdjNormalizeText_(value)
    .replace(/\s+/g, '_')
    .replace(/[^\w]/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_+|_+$/g, '');
}

/**
 * Extrae materias/cursos desde la fila 4.
 * Cada curso ocupa 10 columnas.
 */
function mdjExtractCoursesFromSheet_(sheet) {
  const lastCol = sheet.getLastColumn();
  const headerRow = sheet.getRange(4, 1, 1, lastCol).getDisplayValues()[0];
  const courses = [];

  for (let col = 3; col <= lastCol; col++) {
    const value = String(headerRow[col - 1] || '').trim();
    if (!value) continue;
    if (mdjNormalizeText_(value) === 'CONSOLIDADO') continue;

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

    col += 9; // el for suma 1 adicional y queda justo en la siguiente materia
  }

  return courses;
}

/**
 * Obtiene resumen general del libro.
 */
function mdjGetBookSummary_(libroKey) {
  const libro = mdjGetLibroByKey(libroKey);
  if (!libro) {
    return { totalCursos: 0, totalAlumnos: 0 };
  }

  const ss = mdjGetSpreadsheetByKey_(libroKey);
  const sh = ss.getSheetByName('I BIMESTRE');
  if (!sh) {
    return { totalCursos: 0, totalAlumnos: 0 };
  }

  const courses = mdjExtractCoursesFromSheet_(sh);
  if (!courses.length) {
    return { totalCursos: 0, totalAlumnos: 0 };
  }

  const students = mdjExtractStudentsForCourse_(sh, courses[0]);

  return {
    totalCursos: courses.length,
    totalAlumnos: students.length
  };
}

/**
 * Intenta detectar el año académico desde las primeras hojas.
 */
function mdjGetAcademicYearForLibro_(libroKey) {
  const ss = mdjGetSpreadsheetByKey_(libroKey);
  const firstSheets = ['CONSOLIDADO', 'I BIMESTRE'];
  let raw = '';

  firstSheets.some(function(name) {
    const sh = ss.getSheetByName(name);
    if (!sh) return false;

    raw = sh.getRange(1, 1, 2, Math.min(4, sh.getLastColumn()))
      .getDisplayValues()
      .reduce(function(acc, row) {
        return acc.concat(row);
      }, [])
      .join(' ');

    return !!raw;
  });

  const match = String(raw || '').match(/(20\d{2})/);
  return match ? Number(match[1]) : new Date().getFullYear();
}

/**
 * HELPERS CENTRALES MDJ
 * Pegar completo en un archivo nuevo: HelpersMDJ.gs
 * Si ya tienes funciones con estos mismos nombres, deja solo una versión.
 */

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
  if (role === 'administrador') return 'admin';
  if (role === 'admin') return 'admin';
  return 'maestro';
}

function mdjIsTruthy_(value) {
  const normalized = mdjNormalizeText_(value);
  return ['SI', 'TRUE', '1', 'ACTIVO', 'ACTIVA', 'YES'].indexOf(normalized) !== -1;
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

function mdjColumnToLetter_(column) {
  var letter = '';
  while (column > 0) {
    var temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
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

function mdjExtractCoursesFromSheet_(sheet) {
  const lastCol = sheet.getLastColumn();
  const headerRow = sheet.getRange(4, 1, 1, lastCol).getDisplayValues()[0];
  const courses = [];

  for (let col = 3; col <= lastCol; col++) {
    const value = String(headerRow[col - 1] || '').trim();
    if (!value) continue;
    if (mdjNormalizeText_(value) === 'CONSOLIDADO') continue;

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

    col += 9; // el for sumará 1 adicional, quedando bien en el siguiente bloque
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

  const values = sheet.getRange(7, 1, lastRow - 6, Math.max(2, courseMeta.endCol)).getDisplayValues();
  const students = [];

  values.forEach(function(row, index) {
    const rowNumber = index + 7;
    const nombre = String(row[1] || '').trim();
    if (!nombre) return;

    students.push({
      rowNumber: rowNumber,
      numero: row[0],
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
    });
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
  if (!courses.length) {
    return { totalCursos: 0, totalAlumnos: 0 };
  }

  const students = mdjExtractStudentsForCourse_(sh, courses[0]);

  return {
    totalCursos: courses.length,
    totalAlumnos: students.length
  };
}

function mdjGetAcademicYearForLibro_(libroKey) {
  const ss = mdjGetSpreadsheetByKey_(libroKey);
  const firstSheets = ['CONSOLIDADO', 'I BIMESTRE'];
  let raw = '';

  firstSheets.some(function(name) {
    const sh = ss.getSheetByName(name);
    if (!sh) return false;

    raw = sh.getRange(1, 1, 2, Math.min(4, sh.getLastColumn()))
      .getDisplayValues()
      .reduce(function(acc, row) { return acc.concat(row); }, [])
      .join(' ');

    return !!raw;
  });

  const match = String(raw || '').match(/(20\d{2})/);
  return match ? Number(match[1]) : new Date().getFullYear();
}

function mdjEsNombreAlumnoValido_(nombre) {
  nombre = String(nombre || '').trim();
  if (!nombre) return false;

  const normalizado = mdjNormalizeText_(nombre);
  if (!normalizado) return false;

  const descartes = [
    'PROMEDIO',
    'TOTALES',
    'TOTAL',
    'ALUMNO',
    'ALUMNOS',
    'NOMBRE',
    'NOMBRE_DEL_ALUMNO'
  ];

  return descartes.indexOf(normalizado) === -1;
}


function mdjLog_(nivel, mensaje, extra) {
  var ts = new Date().toISOString();
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

function mdjTestAccesoDocumentoPorId(id) {
  id = String(id || '').trim();

  var out = {
    ok: false,
    id: id,
    activeUser: '',
    effectiveUser: '',
    driveOk: false,
    spreadsheetOk: false,
    fileName: '',
    mimeType: '',
    fileUrl: '',
    totalSheets: 0,
    sheetNames: [],
    error: ''
  };

  try {
    out.activeUser = Session.getActiveUser().getEmail() || '';
  } catch (e1) {}

  try {
    out.effectiveUser = Session.getEffectiveUser().getEmail() || '';
  } catch (e2) {}

  mdjLog_('INFO', 'mdjTestAccesoDocumentoPorId inicio', {
    id: id,
    activeUser: out.activeUser,
    effectiveUser: out.effectiveUser
  });

  try {
    var file = DriveApp.getFileById(id);
    out.driveOk = true;
    out.fileName = file.getName();
    out.mimeType = file.getMimeType();
    out.fileUrl = file.getUrl();

    mdjLog_('INFO', 'DriveApp.getFileById OK', {
      id: id,
      fileName: out.fileName,
      mimeType: out.mimeType
    });
  } catch (e) {
    out.error = 'DriveApp.getFileById falló: ' + (e && e.message ? e.message : String(e));
    mdjLog_('ERROR', out.error);
    return out;
  }

  if (out.mimeType !== MimeType.GOOGLE_SHEETS &&
      out.mimeType !== 'application/vnd.google-apps.spreadsheet') {
    out.error = 'El archivo existe en Drive pero NO es una hoja de cálculo nativa de Google Sheets. MIME detectado: ' + out.mimeType;
    mdjLog_('ERROR', out.error, out);
    return out;
  }

  try {
    var ss = SpreadsheetApp.openById(id);
    out.spreadsheetOk = true;
    out.ok = true;
    out.fileName = ss.getName();

    var sheets = ss.getSheets();
    out.totalSheets = sheets.length;
    out.sheetNames = sheets.map(function(sh) {
      return sh.getName();
    });

    mdjLog_('INFO', 'SpreadsheetApp.openById OK', {
      id: id,
      spreadsheetName: out.fileName,
      totalSheets: out.totalSheets,
      sheetNames: out.sheetNames
    });

    return out;
  } catch (e3) {
    out.error = 'SpreadsheetApp.openById falló: ' + (e3 && e3.message ? e3.message : String(e3));
    mdjLog_('ERROR', out.error, out);
    return out;
  }
}

function mdjTestLibroPorKey(libroKey) {
  var libro = mdjGetLibroByKey(libroKey);
  if (!libro) {
    return {
      ok: false,
      libroKey: libroKey,
      error: 'No se encontró el libro con key: ' + libroKey
    };
  }

  var id = mdjExtraerSpreadsheetId_(libro.driveId || libro.spreadsheetId || libro.url || '');

  var result = mdjTestAccesoDocumentoPorId(id);
  result.libroKey = libro.key;
  result.libroNombre = libro.nombre;

  mdjLog_('INFO', 'mdjTestLibroPorKey resultado', result);
  return result;
}

function mdjTestLibroQuintoPerito() {
  return mdjTestLibroPorKey('quinto_perito');
}