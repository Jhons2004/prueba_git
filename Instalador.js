/**
 * Utilidades de instalación
 */

function mdjCrearPlantillaControl() {
  const ss = SpreadsheetApp.create('MDJ_CONTROL_ACADEMICO');
  const usuarios = ss.getActiveSheet();
  usuarios.setName(MDJ_APP.hojaUsuarios);
  usuarios.clear();

  usuarios.getRange(1, 1, 3, 5).setValues([
    ['usuario', 'password', 'nombre', 'rol', 'activo'],
    ['admin', 'admin123', 'Administrador General', 'admin', 'SI'],
    ['maestro1', 'maestro123', 'Maestro Ejemplo', 'maestro', 'SI']
  ]);

  const asignaciones = ss.insertSheet(MDJ_APP.hojaAsignaciones);
  asignaciones.getRange(1, 1, 3, 4).setValues([
    ['usuario', 'libro_key', 'curso', 'activo'],
    ['maestro1', 'primero_basico', 'MATEMATICA', 'SI'],
    ['maestro1', 'segundo_basico', 'COMUNICACIÓN', 'SI']
  ]);

  usuarios.autoResizeColumns(1, 5);
  asignaciones.autoResizeColumns(1, 4);

  return {
    ok: true,
    spreadsheetId: ss.getId(),
    spreadsheetUrl: ss.getUrl(),
    message: 'Plantilla creada correctamente. Copia el spreadsheetId en MDJ_APP.hojaControlId.'
  };
}

function mdjProbarConexionControl() {
  const ss = mdjGetControlSpreadsheet_();
  return {
    ok: true,
    spreadsheetId: ss.getId(),
    spreadsheetName: ss.getName(),
    sheets: ss.getSheets().map(function(sh) { return sh.getName(); })
  };
}

function mdjListarConfigBoletas() {
  return {
    ok: true,
    libros: mdjGetLibros().map(function(libro) {
      return {
        key: libro.key,
        nombre: libro.nombre,
        driveId: libro.driveId,
        boletasFolderId: libro.boletasFolderId || ''
      };
    })
  };
}