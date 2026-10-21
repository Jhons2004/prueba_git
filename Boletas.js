/**
 * Gestión de carpetas y generación de boletas PDF
 */

function mdjCrearEstructuraBoletasLibro(token, libroKey) {
  mdjRequireRole_(token, ['admin']);
  return mdjCrearEstructuraBoletasLibroInterno_(libroKey);
}

function mdjCrearEstructuraBoletasTodos(token) {
  mdjRequireRole_(token, ['admin']);
  const detalle = [];
  mdjGetLibros().forEach(function(libro) {
    detalle.push(mdjCrearEstructuraBoletasLibroInterno_(libro.key));
  });
  return {
    ok: true,
    totalLibros: detalle.length,
    detalle: detalle
  };
}

function mdjGenerarPdfBoletaAlumno(token, libroKey, rowNumber) {
  mdjRequireRole_(token, ['admin']);
  return mdjGenerarPdfBoletaAlumnoInterno_(libroKey, rowNumber);
}

function mdjGenerarPdfBoletasLibro(token, libroKey) {
  mdjRequireRole_(token, ['admin']);

  const libro = mdjGetLibroByKey(libroKey);
  if (!libro) throw new Error('Libro no encontrado.');

  const rows = mdjGetBoletaRowsByLibro_(libro.key);
  const results = [];
  let generados = 0;
  let errores = 0;

  rows.forEach(function(item) {
    try {
      const result = mdjGenerarPdfBoletaAlumnoInterno_(libro.key, item.rowNumber);
      generados++;
      results.push({
        ok: true,
        rowNumber: item.rowNumber,
        alumno: item.nombre,
        fileId: result.fileId,
        fileUrl: result.fileUrl
      });
    } catch (error) {
      errores++;
      results.push({
        ok: false,
        rowNumber: item.rowNumber,
        alumno: item.nombre,
        error: error.message
      });
    }
  });

  return {
    ok: true,
    libro: libro,
    total: rows.length,
    generados: generados,
    errores: errores,
    detalle: results
  };
}

function mdjGenerarPdfBoletasTodos(token) {
  mdjRequireRole_(token, ['admin']);

  const resumen = [];
  let totalLibros = 0;
  let totalGenerados = 0;
  let totalErrores = 0;

  mdjGetLibros().forEach(function(libro) {
    const result = mdjGenerarPdfBoletasLibro(token, libro.key);
    totalLibros++;
    totalGenerados += Number(result.generados || 0);
    totalErrores += Number(result.errores || 0);
    resumen.push({
      libroKey: libro.key,
      libroNombre: libro.nombre,
      total: result.total,
      generados: result.generados,
      errores: result.errores
    });
  });

  return {
    ok: true,
    totalLibros: totalLibros,
    totalGenerados: totalGenerados,
    totalErrores: totalErrores,
    resumen: resumen
  };
}

function mdjDiagnosticoBoletasDrive(token, libroKey) {
  mdjRequireRole_(token, ['admin']);
  const libro = mdjGetLibroByKey(libroKey);
  if (!libro) throw new Error('Libro no encontrado.');

  const gradeFolder = mdjGetBoletasGradeFolder_(libro.key);
  const rows = mdjGetBoletaRowsByLibro_(libro.key);

  return {
    ok: true,
    libro: libro,
    folderId: gradeFolder.getId(),
    folderUrl: gradeFolder.getUrl(),
    totalAlumnos: rows.length,
    ejemploAlumno: rows[0] || null
  };
}

/**
 * ==============================
 * Internas
 * ==============================
 */

function mdjGenerarPdfBoletaAlumnoInterno_(libroKey, rowNumber) {
  const data = mdjBuildBoletaAlumnoData_(libroKey, rowNumber);
  const folderInfo = mdjEnsureAlumnoBoletaFolders_(libroKey, data.alumno);

  const fileName = mdjBuildBoletaPdfName_(data);
  const html = mdjBuildBoletaPdfHtml_(data);
  const blob = HtmlService.createHtmlOutput(html).getBlob().getAs(MimeType.PDF).setName(fileName);

  mdjTrashExistingFilesByName_(folderInfo.targetFolder, fileName);
  const file = folderInfo.targetFolder.createFile(blob);

  return {
    ok: true,
    libro: data.libro,
    alumno: data.alumno,
    rowNumber: data.rowNumber,
    fileId: file.getId(),
    fileUrl: file.getUrl(),
    fileName: file.getName(),
    folderId: folderInfo.targetFolder.getId(),
    folderUrl: folderInfo.targetFolder.getUrl(),
    carpetaAlumnoId: folderInfo.studentFolder.getId(),
    carpetaAlumnoUrl: folderInfo.studentFolder.getUrl()
  };
}

function mdjCrearEstructuraBoletasLibroInterno_(libroKey) {
  const libro = mdjGetLibroByKey(libroKey);
  if (!libro) throw new Error('Libro no encontrado.');

  const gradeFolder = mdjGetBoletasGradeFolder_(libro.key);
  const rows = mdjGetBoletaRowsByLibro_(libro.key);
  const detalle = [];

  rows.forEach(function(item) {
    const folderInfo = mdjEnsureAlumnoBoletaFolders_(libro.key, item);
    detalle.push({
      rowNumber: item.rowNumber,
      numero: item.numero,
      nombre: item.nombre,
      carpetaAlumnoId: folderInfo.studentFolder.getId(),
      carpetaAlumnoUrl: folderInfo.studentFolder.getUrl(),
      carpetaConsolidadoId: folderInfo.targetFolder.getId(),
      carpetaConsolidadoUrl: folderInfo.targetFolder.getUrl()
    });
  });

  return {
    ok: true,
    libro: libro,
    folderId: gradeFolder.getId(),
    folderUrl: gradeFolder.getUrl(),
    totalAlumnos: rows.length,
    detalle: detalle
  };
}

function mdjGetBoletasGradeFolder_(libroKey) {
  const libro = mdjGetLibroByKey(libroKey);
  if (!libro) throw new Error('Libro no encontrado: ' + libroKey);

  const folderId = MDJ_BOLETAS_DRIVE.foldersByLibro[libro.key];
  if (!folderId) {
    throw new Error('No hay carpeta de boletas configurada para ' + libro.nombre + '.');
  }
  return DriveApp.getFolderById(folderId);
}

function mdjEnsureAlumnoBoletaFolders_(libroKey, alumno) {
  const gradeFolder = mdjGetBoletasGradeFolder_(libroKey);
  const studentFolderName = mdjBuildAlumnoFolderName_(alumno);
  const studentFolder = mdjGetOrCreateChildFolder_(gradeFolder, studentFolderName);

  MDJ_BOLETAS_DRIVE.subcarpetasAlumno.forEach(function(name) {
    mdjGetOrCreateChildFolder_(studentFolder, name);
  });

  const targetFolder = mdjGetOrCreateChildFolder_(studentFolder, MDJ_BOLETAS_DRIVE.carpetaDestinoPdf);

  return {
    gradeFolder: gradeFolder,
    studentFolder: studentFolder,
    targetFolder: targetFolder
  };
}

function mdjBuildBoletaPdfName_(data) {
  const anio = Number(data.anio || new Date().getFullYear());
  const alumno = mdjSanitizeFileName_(String(data.alumno && data.alumno.nombre || 'Alumno'));
  const libro = mdjSanitizeFileName_(String(data.libro && data.libro.nombre || 'Libro'));
  return MDJ_BOLETAS_DRIVE.nombreBoletaBase + ' - ' + alumno + ' - ' + libro + ' - ' + anio + '.pdf';
}

function mdjBuildBoletaPdfHtml_(data) {
  const alumno = data.alumno || {};
  const libro = data.libro || {};
  const grupos = Array.isArray(data.grupos) ? data.grupos : [];
  const anio = Number(data.anio || new Date().getFullYear());

  return [
    '<!DOCTYPE html>',
    '<html><head><meta charset="utf-8"><style>',
    '@page { size: A4; margin: 26px 24px; }',
    'body { font-family: Arial, sans-serif; color:#18324a; font-size:12px; }',
    '.header { display:flex; align-items:center; gap:16px; border-bottom:3px solid #0f4c81; padding-bottom:12px; margin-bottom:16px; }',
    '.logo { width:84px; height:84px; object-fit:contain; }',
    '.title h1 { margin:0; font-size:22px; color:#0f4c81; }',
    '.title h2 { margin:4px 0 0; font-size:13px; color:#55718d; font-weight:normal; }',
    '.meta { width:100%; border-collapse:collapse; margin:14px 0 18px; }',
    '.meta td { border:1px solid #d8e5f2; padding:8px 10px; }',
    '.meta .label { width:160px; font-weight:bold; background:#eff6fc; }',
    '.section { margin-top:14px; }',
    '.section h3 { margin:0 0 8px; padding:8px 10px; background:#0f4c81; color:#fff; font-size:13px; }',
    'table.tbl { width:100%; border-collapse:collapse; }',
    'table.tbl th, table.tbl td { border:1px solid #d8e5f2; padding:7px 8px; }',
    'table.tbl th { background:#eff6fc; text-align:left; color:#314b65; }',
    '.footer { margin-top:18px; font-size:10px; color:#64748b; border-top:1px solid #d8e5f2; padding-top:8px; }',
    '</style></head><body>',
    '<div class="header">',
    '<img class="logo" src="', MDJ_LOGO_DATA_URI, '" alt="Logo">',
    '<div class="title">',
    '<h1>', mdjEscapeHtml_(MDJ_APP.nombre), '</h1>',
    '<h2>', mdjEscapeHtml_(MDJ_APP.lema), '</h2>',
    '</div></div>',
    '<table class="meta">',
    '<tr><td class="label">Alumno</td><td>', mdjEscapeHtml_(alumno.nombre || ''), '</td><td class="label">No.</td><td>', mdjEscapeHtml_(alumno.numero || ''), '</td></tr>',
    '<tr><td class="label">Grado</td><td>', mdjEscapeHtml_(libro.nombre || ''), '</td><td class="label">Ciclo</td><td>', mdjEscapeHtml_(String(anio)), '</td></tr>',
    '<tr><td class="label">Programa</td><td>', mdjEscapeHtml_(libro.programa || ''), '</td><td class="label">Nivel</td><td>', mdjEscapeHtml_(libro.nivel || ''), '</td></tr>',
    '</table>',
    grupos.map(function(group) {
      return [
        '<div class="section">',
        '<h3>', mdjEscapeHtml_(group.section || 'General'), '</h3>',
        '<table class="tbl"><thead><tr><th>Indicador</th><th>Valor</th></tr></thead><tbody>',
        (group.items || []).map(function(item) {
          return '<tr><td>' + mdjEscapeHtml_(item.label || '') + '</td><td>' + mdjEscapeHtml_(item.value || '') + '</td></tr>';
        }).join(''),
        '</tbody></table></div>'
      ].join('');
    }).join(''),
    '<div class="footer">Documento generado automáticamente desde el sistema académico. Fecha de emisión: ',
    mdjEscapeHtml_(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')),
    '</div>',
    '</body></html>'
  ].join('');
}

function mdjTrashExistingFilesByName_(folder, fileName) {
  const files = folder.getFilesByName(fileName);
  while (files.hasNext()) {
    files.next().setTrashed(true);
  }
}

function mdjBuildAlumnoFolderName_(alumno) {
  const numero = String(alumno.numero || '').trim();
  const nombre = String(alumno.nombre || '').trim();
  const base = numero ? (numero + ' - ' + nombre) : nombre;
  return mdjSanitizeFolderName_(base);
}

function mdjGetOrCreateChildFolder_(parentFolder, folderName) {
  const cleanName = mdjSanitizeFolderName_(folderName);
  const folders = parentFolder.getFoldersByName(cleanName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return parentFolder.createFolder(cleanName);
}

function mdjSanitizeFolderName_(value) {
  return String(value || '')
    .replace(/[\\/:*?"<>|#%{}~]/g, '-')
    .replace(/\s+/g, ' ')
    .replace(/^\s+|\s+$/g, '')
    .substring(0, 180);
}

function mdjSanitizeFileName_(value) {
  return String(value || '')
    .replace(/[\\/:*?"<>|#%{}~]/g, '-')
    .replace(/\s+/g, ' ')
    .replace(/^\s+|\s+$/g, '');
}

function mdjEscapeHtml_(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function mdjGetBoletaRowsByLibro_(libroKey) {
  const ss = mdjGetSpreadsheetByKey_(libroKey);
  const sh = ss.getSheetByName('I BIMESTRE');
  if (!sh) throw new Error('No existe la hoja I BIMESTRE para ' + libroKey + '.');

  const courses = mdjExtractCoursesFromSheet_(sh);
  if (!courses.length) return [];

  return mdjExtractStudentsForCourse_(sh, courses[0]).map(function(item) {
    return {
      rowNumber: item.rowNumber,
      numero: item.numero,
      nombre: item.nombre
    };
  });
}

function mdjFindStudentByRowOrIdentity_(students, rowNumber, numero, nombre) {
  return (students || []).find(function(item) {
    return Number(item.rowNumber || 0) === Number(rowNumber || 0) ||
      mdjBuildStudentKey_(item.numero, item.nombre) === mdjBuildStudentKey_(numero, nombre);
  }) || null;
}

function mdjBuildBoletaConsolidadoItems_(libroKey, rowNumber, numero, nombre) {
  const ss = mdjGetSpreadsheetByKey_(libroKey);
  const refSheet = ss.getSheetByName('I BIMESTRE') || ss.getSheets()[0];
  if (!refSheet) return [];

  const courses = mdjExtractCoursesFromSheet_(refSheet);

  return courses.map(function(course) {
    const consolidados = mdjGetConsolidatedCourseStudents_(libroKey, course.name);
    const alumno = mdjFindStudentByRowOrIdentity_(consolidados, rowNumber, numero, nombre);

    return {
      label: course.name,
      value: alumno && alumno.promedioMostrar !== '' ? String(alumno.promedioMostrar) : '-'
    };
  });
}

function mdjBuildBoletaAlumnoData_(libroKey, rowNumber) {
  const libro = mdjGetLibroByKey(libroKey);
  if (!libro) throw new Error('Libro no encontrado.');

  const ss = mdjGetSpreadsheetByKey_(libroKey);
  const rows = mdjGetBoletaRowsByLibro_(libroKey);
  const alumno = rows.find(function(item) {
    return Number(item.rowNumber || 0) === Number(rowNumber || 0);
  });

  if (!alumno) {
    throw new Error('No se encontró el alumno en la fila ' + rowNumber + '.');
  }

  const grupos = [];

  mdjGetBimestresOperativos_().forEach(function(bimestre) {
    const sh = ss.getSheetByName(bimestre);
    if (!sh) return;

    const courses = mdjExtractCoursesFromSheet_(sh);

    const items = courses.map(function(course) {
      const students = mdjExtractStudentsForCourse_(sh, course);
      const st = mdjFindStudentByRowOrIdentity_(students, rowNumber, alumno.numero, alumno.nombre);

      return {
        label: course.name,
        value: st && st.notaMostrar !== '' ? String(st.notaMostrar) : '-'
      };
    });

    grupos.push({
      section: bimestre,
      items: items
    });
  });

  grupos.push({
    section: 'CONSOLIDADO',
    items: mdjBuildBoletaConsolidadoItems_(libroKey, rowNumber, alumno.numero, alumno.nombre)
  });

  return {
    libro: libro,
    alumno: alumno,
    rowNumber: rowNumber,
    grupos: grupos,
    anio: new Date().getFullYear()
  };
}