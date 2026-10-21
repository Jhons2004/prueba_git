/**
 * Codigo.gs - Entrada web, bootstrap de páginas
 * Colegio Mundo de Juguetes - Módulo Asistencia
 */

/**
 * Entrada principal del Web App.
 */
// El enrutamiento y la función include están definidos en Code.js.

/**
 * Logger centralizado.
 */
function mdjAsisLog_(nivel, mensaje, extra) {
  const ts = new Date().toISOString();

  try {
    if (extra !== undefined) {
      Logger.log(
        '[' + ts + '] [ASIS ' + nivel + '] ' +
        mensaje + ' | ' + JSON.stringify(extra)
      );
    } else {
      Logger.log('[' + ts + '] [ASIS ' + nivel + '] ' + mensaje);
    }
  } catch (e) {
    Logger.log('[' + ts + '] [ASIS ' + nivel + '] ' + mensaje);
  }
}


/**
 * Bootstrap para la vista TomarAsistencia.html
 *
 * Permite:
 * - admin
 * - maestro
 */
function mdjAsisGetTomarBootstrap(token) {
  try {
    if (!MDJ_ASIS_APP.asistenciaSpreadsheetId) throw new Error('Configura MDJ_ASISTENCIA_SPREADSHEET_ID para la hoja de asistencia QA.');
    const sess = mdjAsisRequireSession_(token);
    const user = sess.user || {};
    const role = String(user.rol || '').trim().toLowerCase();

    if (role !== 'admin' && role !== 'maestro') {
      return {
        ok: false,
        message: 'No tienes permisos para tomar asistencia.'
      };
    }

    return {
      ok: true,
      user: user,
      role: role,
      grados: mdjAsisGetGradosPublicos_(user),
      fechaHoy: mdjAsisFechaHoy_(),
      asistenciaSpreadsheetId: MDJ_ASIS_APP.asistenciaSpreadsheetId,
      asistenciaUrl: mdjAsisGetAsistenciaUrl_()
    };

  } catch (e) {
    mdjAsisLog_('ERROR', 'mdjAsisGetTomarBootstrap', {
      message: e.message,
      stack: e.stack
    });

    return {
      ok: false,
      message: e.message || 'Error cargando módulo de asistencia.'
    };
  }
}


/**
 * Bootstrap para AdminAsistencia.html
 *
 * Permite:
 * - admin
 */
function mdjAsisGetAdminBootstrap(token) {
  try {
    if (!MDJ_ASIS_APP.asistenciaSpreadsheetId) throw new Error('Configura MDJ_ASISTENCIA_SPREADSHEET_ID para la hoja de asistencia QA.');
    const sess = mdjAsisRequireRole_(token, ['admin']);
    const user = sess.user || {};

    return {
      ok: true,
      user: user,
      role: 'admin',
      grados: mdjAsisGetGradosPublicos_(user),
      fechaHoy: mdjAsisFechaHoy_(),
      asistenciaSpreadsheetId: MDJ_ASIS_APP.asistenciaSpreadsheetId,
      asistenciaUrl: mdjAsisGetAsistenciaUrl_()
    };

  } catch (e) {
    mdjAsisLog_('ERROR', 'mdjAsisGetAdminBootstrap', {
      message: e.message,
      stack: e.stack
    });

    return {
      ok: false,
      message: e.message || 'Error cargando panel administrativo.'
    };
  }
}


/**
 * Devuelve información pública de grados para frontend.
 * No expone más de lo necesario, pero conserva driveId para admin si se necesita.
 */
function mdjAsisGetGradosPublicos_(user) {
  const role = String(user && user.rol || '').toLowerCase();
  let visibles = mdjAsisListarGradosConfigurados_();
  if (role !== 'admin') {
    const asignados = mdjGetAssignmentsForUser_(user && user.usuario || '');
    const permitidos = {};
    asignados.forEach(function(a) { permitidos[mdjResolveLibroKey_(a.libroKey)] = true; });
    visibles = visibles.filter(function(g) { return !!permitidos[g.key]; });
  }
  return visibles.map(function(g) {
    return {
      key: g.key, nombre: g.nombre, nivel: g.nivel || '', grado: g.grado || '',
      asistenciaSpreadsheetId: MDJ_ASIS_APP.asistenciaSpreadsheetId,
      asistenciaUrl: mdjAsisGetAsistenciaUrl_()
    };
  });
}

/** Protege APIs aunque alguien modifique las opciones del HTML. */
function mdjAsisAssertGradoAccess_(user, libroKey) {
  const role = String(user && user.rol || '').toLowerCase();
  mdjAsisGetLibroConfig_(libroKey);
  if (role === 'admin') return true;
  if (role !== 'maestro') throw new Error('Sin permisos de asistencia.');
  const permitidos = mdjAsisGetGradosPublicos_(user);
  if (!permitidos.some(function(g) { return g.key === libroKey; })) {
    throw new Error('No tienes asignado este grado para asistencia.');
  }
  return true;
}


/**
 * [NUEVA FUNCIÓN DE SEGURIDAD - MURO DE FUEGO]
 * Valida de forma estricta que un "libroKey" recibido del frontend
 * exista en nuestra configuración oficial. Si no existe, bloquea la ejecución.
 */
function mdjAsisGetLibroConfig_(libroKey) {
  const libro = mdjAsisListarGradosConfigurados_().find(function(g) {
    return g.key === libroKey;
  });

  if (!libro) {
    mdjAsisLog_('SECURITY', 'Intento de acceso a un grado/hoja no configurado', { key: libroKey });
    throw new Error('Acceso denegado: El grado solicitado no está autorizado en la configuración.');
  }

  return libro;
}


/**
 * URL de hoja maestra.
 */
function mdjAsisGetAsistenciaUrl_() {
  return MDJ_ASIS_APP.asistenciaSpreadsheetId ? 'https://docs.google.com/spreadsheets/d/' +
    MDJ_ASIS_APP.asistenciaSpreadsheetId + '/edit' : '';
}


/**
 * Validación rápida para probar desde Apps Script.
 */
function mdjAsisTestCodigoBootstrap() {
  return {
    ok: true,
    app: MDJ_ASIS_APP.nombre,
    version: MDJ_ASIS_APP.version,
    asistenciaSpreadsheetId: MDJ_ASIS_APP.asistenciaSpreadsheetId,
    asistenciaUrl: mdjAsisGetAsistenciaUrl_(),
    totalGrados: mdjAsisListarGradosConfigurados_().length,
    pages: {
      login: MDJ_ASIS_APP.loginPage,
      tomar: MDJ_ASIS_APP.tomarPage,
      admin: MDJ_ASIS_APP.adminPage
    }
  };
}


/**
 * Escape básico para errores HTML.
 */
function mdjAsisEscHtml_(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}