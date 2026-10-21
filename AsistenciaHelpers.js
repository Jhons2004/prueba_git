/**
 * Helpers.gs - Utilidades compartidas
 * Colegio Mundo de Juguetes - Módulo Asistencia
 */

/**
 * Normaliza textos para comparar:
 * - quita tildes
 * - convierte caracteres raros a _
 * - pasa a mayúsculas
 */
function mdjAsisNormalizeText_(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^A-Za-z0-9_]+/g, '_')
    .replace(/^_+|_+$/g, '')
    .toUpperCase();
}


/**
 * Normaliza encabezados de hojas.
 */
function mdjAsisNormalizeHeaderName_(value) {
  return mdjAsisNormalizeText_(String(value || '').replace(/\s+/g, '_'));
}


/**
 * Normaliza roles.
 */
function mdjAsisNormalizeRole_(value) {
  const role = mdjAsisNormalizeText_(value).toLowerCase();

  if (
    role === 'administrador' ||
    role === 'admin' ||
    role === 'administrativo'
  ) {
    return 'admin';
  }

  return 'maestro';
}


/**
 * Interpreta valores booleanos desde hoja.
 */
function mdjAsisIsTruthy_(value) {
  const normalized = mdjAsisNormalizeText_(value);

  return [
    'SI',
    'TRUE',
    '1',
    'ACTIVO',
    'ACTIVA',
    'YES',
    'X'
  ].indexOf(normalized) !== -1;
}


/**
 * Construye mapa de encabezados para hoja USUARIOS.
 */
function mdjAsisBuildUserHeaderMap_(headerRow) {
  const map = {};

  (headerRow || []).forEach(function(name, index) {
    const key = mdjAsisNormalizeHeaderName_(name);
    if (key) map[key] = index;
  });

  const aliases = {
    usuario: [
      'usuario',
      'user',
      'correo',
      'email',
      'mail'
    ],
    password: [
      'password',
      'contrasena',
      'contraseña',
      'clave',
      'pass'
    ],
    nombre: [
      'nombre',
      'nombre_completo',
      'nombres',
      'docente',
      'maestro'
    ],
    rol: [
      'rol',
      'role',
      'perfil',
      'tipo'
    ],
    activo: [
      'activo',
      'estado',
      'habilitado',
      'status'
    ]
  };

  const out = {};

  Object.keys(aliases).forEach(function(target) {
    const found = aliases[target]
      .map(mdjAsisNormalizeHeaderName_)
      .find(function(a) {
        return Object.prototype.hasOwnProperty.call(map, a);
      });

    out[target] = found ? map[found] : -1;
  });

  const requeridos = ['usuario', 'password', 'nombre', 'rol', 'activo'];
  const faltantes = requeridos.filter(function(k) {
    return out[k] === -1;
  });

  if (faltantes.length) {
    throw new Error(
      'Faltan encabezados en hoja USUARIOS: ' + faltantes.join(', ')
    );
  }

  return out;
}


/**
 * Abre el Spreadsheet de control de usuarios.
 */
function mdjAsisGetControlSpreadsheet_() {
  const id = String(MDJ_ASIS_APP.hojaControlId || '').trim();

  if (!id) {
    throw new Error('Configura MDJ_ASIS_APP.hojaControlId.');
  }

  return SpreadsheetApp.openById(id);
}


/**
 * Abre una hoja dentro del Spreadsheet de control.
 */
function mdjAsisGetControlSheet_(sheetName) {
  const ss = mdjAsisGetControlSpreadsheet_();
  const sh = ss.getSheetByName(sheetName);

  if (!sh) {
    throw new Error('No existe la hoja "' + sheetName + '" en el control de usuarios.');
  }

  return sh;
}


/**
 * Abre el Spreadsheet maestro de asistencia.
 */
function mdjAsisGetAsistenciaSpreadsheet_() {
  const id = String(MDJ_ASIS_APP.asistenciaSpreadsheetId || '').trim();

  if (!id) {
    throw new Error(
      'Falta configurar MDJ_ASIS_APP.asistenciaSpreadsheetId.'
    );
  }

  return SpreadsheetApp.openById(id);
}


/**
 * Devuelve la hoja INDEX.
 */
function mdjAsisGetIndexSheet_() {
  const ss = mdjAsisGetAsistenciaSpreadsheet_();
  const sh = ss.getSheetByName(MDJ_ASIS_INDEX.sheetName || 'INDEX');

  if (!sh) {
    throw new Error('No existe la hoja INDEX en la hoja maestra de asistencia.');
  }

  return sh;
}


/**
 * Obtiene una hoja de grado.
 */
function mdjAsisGetSheetGrado_(libroKey) {
  const grado = mdjAsisGetGradoByKey_(libroKey);

  if (!grado) {
    throw new Error('Grado no configurado: ' + libroKey);
  }

  const ss = mdjAsisGetAsistenciaSpreadsheet_();
  const sh = ss.getSheetByName(grado.key);

  if (!sh) {
    throw new Error(
      'No existe la hoja del grado "' + grado.key +
      '". Ejecuta primero la creación de estructura.'
    );
  }

  return sh;
}


/**
 * Crea o devuelve una hoja de grado.
 */
function mdjAsisGetOrCreateSheetGrado_(libroKey) {
  const grado = mdjAsisGetGradoByKey_(libroKey);

  if (!grado) {
    throw new Error('Grado no configurado: ' + libroKey);
  }

  const ss = mdjAsisGetAsistenciaSpreadsheet_();
  let sh = ss.getSheetByName(grado.key);

  if (!sh) {
    sh = ss.insertSheet(grado.key);
  }

  return sh;
}


/**
 * Fecha de hoy en formato ISO.
 */
function mdjAsisFechaHoy_() {
  return Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd'
  );
}


/**
 * Fecha y hora legible.
 */
function mdjAsisFechaHoraAhora_() {
  return Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd HH:mm:ss'
  );
}


/**
 * Convierte fecha ISO a encabezado.
 * 2026-04-23 -> 23/04/2026
 */
function mdjAsisFechaToHeader_(fechaISO) {
  const clean = String(fechaISO || '').trim();
  const parts = clean.split('-');

  if (parts.length !== 3) {
    return clean;
  }

  return parts[2] + '/' + parts[1] + '/' + parts[0];
}


/**
 * Convierte encabezado a fecha ISO.
 * 23/04/2026 -> 2026-04-23
 */
function mdjAsisHeaderToFecha_(header) {
  const clean = String(header || '').trim();
  const parts = clean.split('/');

  if (parts.length !== 3) {
    return '';
  }

  return parts[2] + '-' + parts[1] + '-' + parts[0];
}


/**
 * Valida fecha ISO simple.
 */
function mdjAsisIsFechaISO_(fechaISO) {
  return /^\d{4}-\d{2}-\d{2}$/.test(String(fechaISO || '').trim());
}


/**
 * Busca grado por key.
 */
function mdjAsisGetGradoByKey_(key) {
  key = String(key || '').trim();

  return mdjAsisListarGradosConfigurados_().find(function(g) {
    return String(g.key || '') === key;
  }) || null;
}


/**
 * Escapa HTML.
 */
function mdjAsisEscapeHtml_(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}


/**
 * Limpia un nombre de estudiante.
 */
function mdjAsisCleanStudentName_(value) {
  return String(value || '')
    .replace(/\s+/g, ' ')
    .trim();
}


/**
 * Normaliza nombre de estudiante para evitar duplicados.
 */
function mdjAsisNormalizeStudentName_(value) {
  return mdjAsisNormalizeText_(
    mdjAsisCleanStudentName_(value)
  );
}


/**
 * Valida estado de asistencia.
 */
function mdjAsisNormalizeEstado_(estado) {
  const e = mdjAsisNormalizeText_(estado);

  if (e === 'SI' || e === 'SÍ') return 'SI';
  if (e === 'NO') return 'NO';
  if (e === 'JUSTIFICADO' || e === 'JUSTIFICADA' || e === 'J') return 'JUSTIFICADO';

  return '';
}


/**
 * Determina si el estado cuenta como asistencia.
 */
function mdjAsisEstadoCuentaAsistencia_(estado) {
  const e = mdjAsisNormalizeEstado_(estado);
  return e === 'SI' || e === 'JUSTIFICADO';
}


/**
 * URL de la hoja maestra.
 */
function mdjAsisBuildAsistenciaUrl_() {
  return 'https://docs.google.com/spreadsheets/d/' +
    MDJ_ASIS_APP.asistenciaSpreadsheetId +
    '/edit';
}


/**
 * Retorna metadata de asistencia.
 */
function mdjAsisGetAsistenciaMeta_() {
  return {
    spreadsheetId: MDJ_ASIS_APP.asistenciaSpreadsheetId,
    url: mdjAsisBuildAsistenciaUrl_(),
    index: MDJ_ASIS_INDEX.sheetName,
    totalGrados: mdjAsisListarGradosConfigurados_().length
  };
}


/**
 * Convierte valores de hoja en texto seguro.
 */
function mdjAsisToString_(value) {
  return String(value == null ? '' : value).trim();
}


/**
 * Convierte número de hoja en entero seguro.
 */
function mdjAsisToInt_(value, fallback) {
  const n = parseInt(value, 10);
  return isNaN(n) ? (fallback || 0) : n;
}


/**
 * Prueba rápida de helpers.
 */
function mdjAsisTestHelpers() {
  const ss = mdjAsisGetAsistenciaSpreadsheet_();

  return {
    ok: true,
    asistenciaSpreadsheetId: ss.getId(),
    asistenciaUrl: mdjAsisBuildAsistenciaUrl_(),
    fechaHoy: mdjAsisFechaHoy_(),
    grados: mdjAsisListarGradosConfigurados_().length
  };
}