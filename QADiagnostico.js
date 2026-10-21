/** Diagnóstico de integración SIN escritura en hojas. */
function mdjQaDiagnosticoUnificacion() {
  const activos = MDJ_LIBROS.filter(function(l) { return !!l.driveId; });
  const claves = MDJ_LIBROS.map(function(l) { return l.key; });
  const duplicados = claves.filter(function(k, i) { return claves.indexOf(k) !== i; });
  const salida = {
    version: MDJ_APP.version,
    app: MDJ_APP.nombre,
    grados: MDJ_LIBROS.length,
    clavesDuplicadas: duplicados,
    fuenteAcademicaConfigurada: activos.length,
    controlConfigurado: !!MDJ_APP.hojaControlId,
    asistenciaConfigurada: !!MDJ_ASIS_APP.asistenciaSpreadsheetId,
    logoInstitucionalDisponible: !!mdjAsisGetLogoDataUri_(),
    mismaFuenteLibros: mdjAsisListarGradosConfigurados_().every(function(a) {
      return MDJ_LIBROS.some(function(l) { return l.key === a.key && l.driveId === a.driveId; });
    }),
    sesionUnica: true,
    mensajes: []
  };
  if (!salida.controlConfigurado) salida.mensajes.push('Configura MDJ_CONTROL_SPREADSHEET_ID (copia QA).');
  if (salida.fuenteAcademicaConfigurada < salida.grados) salida.mensajes.push('Configura los libros MDJ_LIBRO_* (copias QA).');
  if (!salida.asistenciaConfigurada) salida.mensajes.push('Configura MDJ_ASISTENCIA_SPREADSHEET_ID (copia QA).');
  return salida;
}
