/**
 * Módulo de asistencia integrado. El catálogo académico, usuarios y logo
 * provienen del núcleo central (Config.js / Auth.js).
 * No modificar los ID de producción durante las pruebas de QA.
 */
const MDJ_ASIS_APP = {
  nombre: 'Colegio Mundo de Juguetes',
  subtitulo: 'Módulo de Control de Asistencia',
  version: '2.2.0-QA',
  sessionTtlSeconds: 60 * 60 * 6,

  // Login compartido con sistema académico.
  get hojaControlId() { return MDJ_APP.hojaControlId; },
  get hojaUsuarios() { return MDJ_APP.hojaUsuarios; },

  // Establecer MDJ_ASISTENCIA_SPREADSHEET_ID en las Propiedades del Script QA.
  asistenciaSpreadsheetId: PropertiesService.getScriptProperties().getProperty('MDJ_ASISTENCIA_SPREADSHEET_ID') || '',

  // Páginas HTML.
  loginPage: 'Login',
  tomarPage: 'TomarAsistencia',
  adminPage: 'AdminAsistencia',

  // Estados válidos.
  estadosAsistencia: ['SI', 'NO', 'JUSTIFICADO'],

  // Hoja fuente donde están los alumnos en los libros de notas.
  hojaFuenteAlumnos: 'I BIMESTRE',

  // Layout de alumnos en el libro original de notas.
  layoutFuenteAlumnos: {
    rowFirstStudent: 7,
    colNumero: 1,
    colNombre: 2
  },

};


function mdjAsisListarGradosConfigurados_() {
  return MDJ_LIBROS.map(function(libro) {
  return {
    key: libro.key, nombre: libro.nombre, nivel: libro.nivel,
    grado: libro.grado, programa: libro.programa,
    // El libro se resuelve dinámicamente por mdjGetSpreadsheetByKey_.
    driveId: libro.driveId,
    asistenciaSpreadsheetId: MDJ_ASIS_APP.asistenciaSpreadsheetId,
    asistenciaUrl: 'https://docs.google.com/spreadsheets/d/' + MDJ_ASIS_APP.asistenciaSpreadsheetId + '/edit'
  };
  });
}

const MDJ_ASIS_LAYOUT = {
  rowTitulo: 1,
  rowCiclo: 2,
  rowHeaders: 4,
  rowFirstStudent: 5,

  colNumero: 1,
  colNombre: 2,
  colTotalDias: 3,
  colDiasAsistidos: 4,
  colPorcentaje: 5,
  colFirstDate: 6
};


/**
 * Layout de hoja INDEX.
 */
const MDJ_ASIS_INDEX = {
  sheetName: 'INDEX',
  rowTitulo: 1,
  rowHeaders: 2,
  rowFirstData: 3,

  colGrado: 1,
  colKey: 2,
  colTotalAlumnos: 3,
  colUltimaSincronizacion: 4,
  colUltimaAsistencia: 5
};


/**
 * Cache para archivos convertidos, si aplica.
 */
const MDJ_ASIS_CONVERTED_PREFIX = 'MDJ_ASIS_CONVERTED_';


/**
 * Prueba rápida de configuración.
 */
function mdjAsisConfigCheck() {
  return {
    ok: true,
    app: MDJ_ASIS_APP.nombre,
    version: MDJ_ASIS_APP.version,
    asistenciaSpreadsheetId: MDJ_ASIS_APP.asistenciaSpreadsheetId,
    asistenciaUrl: 'https://docs.google.com/spreadsheets/d/' + MDJ_ASIS_APP.asistenciaSpreadsheetId + '/edit',
    totalLibros: MDJ_LIBROS.length,
    totalGradosAsistencia: mdjAsisListarGradosConfigurados_().length,
    hojaFuenteAlumnos: MDJ_ASIS_APP.hojaFuenteAlumnos,
    layoutFuenteAlumnos: MDJ_ASIS_APP.layoutFuenteAlumnos,
    layoutAsistencia: MDJ_ASIS_LAYOUT
  };
}