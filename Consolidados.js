const MDJ_CONSOLIDADOS_CONFIG = {
  rootFolderId: mdjConfigId_('MDJ_CONSOLIDADOS_ROOT_FOLDER_ID'),
  rootFolderName: 'Consolidados',
  gradeFolderPrefix: 'Consolidado - ',
  periodos: [
    'I BIMESTRE',
    'II BIMESTRE',
    'III BIMESTRE',
    'IV BIMESTRE',
    'FINAL'
  ],
  fileNamePrefix: 'Consolidado',
  signatureLabels: {
    maestro: 'Firma del maestro',
    direccion: 'Firma de dirección'
  }
};

function mdjConsolidadosRequireAdmin_(token) {
  if (typeof mdjRequireRole_ === 'function') {
    mdjRequireRole_(token, ['admin']);
    return;
  }

  if (typeof mdjGetSession !== 'function') {
    throw new Error('No existe mdjGetSession para validar permisos.');
  }

  const sess = mdjGetSession(token);
  if (!sess || !sess.ok) {
    throw new Error((sess && sess.message) || 'Sesión inválida.');
  }

  const role = String(sess.user && sess.user.rol || '').toLowerCase();
  if (role !== 'admin') {
    throw new Error('Acceso denegado.');
  }
}

function mdjConsolidadosNormalizeText_(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^A-Za-z0-9_]+/g, '_')
    .replace(/^_+|_+$/g, '')
    .toUpperCase();
}

function mdjConsolidadosSlug_(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^A-Za-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '');
}

function mdjConsolidadosSafeHtml_(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function mdjConsolidadosToNumber_(value) {
  if (value === null || value === undefined || value === '') return null;
  const raw = String(value).trim().replace(',', '.');
  if (!raw || raw === '-') return null;
  if (!/^-?\d+(\.\d+)?$/.test(raw)) return null;
  return Number(raw);
}

function mdjConsolidadosRound2_(value) {
  return Math.round(Number(value || 0) * 100) / 100;
}

function mdjConsolidadosResolveLibro_(libroKey) {
  if (typeof mdjGetLibroByKey === 'function') {
    return mdjGetLibroByKey(libroKey);
  }
  throw new Error('No existe mdjGetLibroByKey en el proyecto.');
}

function mdjConsolidadosGetSpreadsheet_(libroKey) {
  if (typeof mdjGetSpreadsheetByKey_ === 'function') {
    return mdjGetSpreadsheetByKey_(libroKey);
  }
  throw new Error('No existe mdjGetSpreadsheetByKey_ en el proyecto.');
}

function mdjConsolidadosGetRootFolder_() {
  const folder = DriveApp.getFolderById(MDJ_CONSOLIDADOS_CONFIG.rootFolderId);
  if (!folder) {
    throw new Error('No se encontró la carpeta raíz de consolidados.');
  }
  return folder;
}

function mdjConsolidadosGetOrCreateChildFolder_(parentFolder, name) {
  const it = parentFolder.getFoldersByName(name);
  return it.hasNext() ? it.next() : parentFolder.createFolder(name);
}

function mdjConsolidadosGetGradeFolder_(libroKey) {
  const libro = mdjConsolidadosResolveLibro_(libroKey);
  if (!libro) throw new Error('Libro no encontrado: ' + libroKey);

  const root = mdjConsolidadosGetRootFolder_();
  const folderName = MDJ_CONSOLIDADOS_CONFIG.gradeFolderPrefix + libro.nombre;
  const gradeFolder = mdjConsolidadosGetOrCreateChildFolder_(root, folderName);

  return {
    rootFolder: root,
    gradeFolder: gradeFolder,
    libro: libro
  };
}

function mdjConsolidadosBuildPeriodoLabel_(periodo) {
  const p = String(periodo || '').trim().toUpperCase();
  if (p === 'FINAL') return 'FINAL';
  return p;
}

function mdjConsolidadosBuildFileName_(libro, periodo, anio) {
  const grado = mdjConsolidadosSlug_(libro && libro.nombre || 'Grado');
  const per = mdjConsolidadosSlug_(mdjConsolidadosBuildPeriodoLabel_(periodo));
  return [MDJ_CONSOLIDADOS_CONFIG.fileNamePrefix, grado, String(anio || new Date().getFullYear()), per].join('-') + '.pdf';
}

function mdjConsolidadosTrashExistingFilesByName_(folder, fileName) {
  const files = folder.getFilesByName(fileName);
  while (files.hasNext()) {
    files.next().setTrashed(true);
  }
}

function mdjConsolidadosGetBaseStudentsFromSheet_(sheet) {
  if (!sheet) throw new Error('No se recibió hoja base para estudiantes.');

  const lastRow = sheet.getLastRow();
  if (lastRow < 7) return [];

  const values = sheet.getRange(7, 1, lastRow - 6, 2).getDisplayValues();
  const rows = [];

  values.forEach(function(row, idx) {
    const rowNumber = idx + 7;
    const numero = String(row[0] || '').trim();
    const nombre = String(row[1] || '').trim();

    if (typeof mdjEsNombreAlumnoValido_ === 'function') {
      if (!mdjEsNombreAlumnoValido_(nombre)) return;
    } else if (!nombre) {
      return;
    }

    rows.push({
      rowNumber: rowNumber,
      numero: numero,
      nombre: nombre,
      studentKey: mdjConsolidadosNormalizeText_(numero) + '__' + mdjConsolidadosNormalizeText_(nombre)
    });
  });

  return rows;
}

function mdjConsolidadosGetBaseStudents_(libroKey, ss) {
  const spreadsheet = ss || mdjConsolidadosGetSpreadsheet_(libroKey);
  const sh = spreadsheet.getSheetByName('I BIMESTRE') || spreadsheet.getSheets()[0];
  if (!sh) throw new Error('No hay hojas disponibles en el libro.');
  return mdjConsolidadosGetBaseStudentsFromSheet_(sh);
}

function mdjConsolidadosGetCoursesFromSheet_(sheet) {
  if (!sheet) throw new Error('No se recibió hoja base para cursos.');

  if (typeof mdjExtractCoursesFromSheet_ !== 'function') {
    throw new Error('No existe mdjExtractCoursesFromSheet_ en el proyecto.');
  }

  return mdjExtractCoursesFromSheet_(sheet) || [];
}

function mdjConsolidadosGetCourses_(libroKey, ss) {
  const spreadsheet = ss || mdjConsolidadosGetSpreadsheet_(libroKey);
  const sh = spreadsheet.getSheetByName('I BIMESTRE') || spreadsheet.getSheets()[0];
  if (!sh) throw new Error('No hay hojas disponibles para cursos.');
  return mdjConsolidadosGetCoursesFromSheet_(sh);
}

function mdjConsolidadosFindCourseMeta_(sheet, courseName, cachedCourses) {
  if (typeof mdjFindCourseInSheet_ === 'function') {
    return mdjFindCourseInSheet_(sheet, courseName);
  }

  const courses = cachedCourses || (typeof mdjExtractCoursesFromSheet_ === 'function'
    ? (mdjExtractCoursesFromSheet_(sheet) || [])
    : []);

  const wanted = mdjConsolidadosNormalizeText_(courseName);
  return courses.find(function(item) {
    return mdjConsolidadosNormalizeText_(item.name) === wanted;
  }) || null;
}

function mdjConsolidadosResolveStudentScore_(student, periodo) {
  if (!student) return '';

  const candidates = [
    student.finalMostrar,
    student.notaMostrar,
    student.promedioMostrar,
    student.finalNumerico,
    student.notaCalculada10,
    student.promedioNumerico,
    student.promedio,
    student.notaBimestral,
    student.totalCurso
  ];

  for (var i = 0; i < candidates.length; i++) {
    var val = candidates[i];
    if (val !== '' && val !== null && val !== undefined) {
      return val;
    }
  }

  return '';
}

function mdjConsolidadosBuildScoreMapByCourse_(libroKey, periodo, courseName, options) {
  const map = {};
  const opts = options || {};

  if (String(periodo).toUpperCase() === 'FINAL') {
    if (typeof mdjGetConsolidatedCourseStudents_ !== 'function') {
      throw new Error('No existe mdjGetConsolidatedCourseStudents_ en el proyecto.');
    }

    const students = mdjGetConsolidatedCourseStudents_(libroKey, courseName) || [];
    students.forEach(function(st) {
      const key = mdjConsolidadosNormalizeText_(st.numero) + '__' + mdjConsolidadosNormalizeText_(st.nombre);
      map[key] = mdjConsolidadosResolveStudentScore_(st, periodo);
    });
    return map;
  }

  const ss = opts.ss || mdjConsolidadosGetSpreadsheet_(libroKey);
  const sh = opts.sheet || ss.getSheetByName(periodo);
  if (!sh) throw new Error('No existe la hoja ' + periodo + '.');

  const courseMeta = opts.courseMeta || mdjConsolidadosFindCourseMeta_(sh, courseName, opts.courses);
  if (!courseMeta) return map;

  if (typeof mdjExtractStudentsForCourse_ !== 'function') {
    throw new Error('No existe mdjExtractStudentsForCourse_ en el proyecto.');
  }

  const students = mdjExtractStudentsForCourse_(sh, courseMeta) || [];
  students.forEach(function(st) {
    const key = mdjConsolidadosNormalizeText_(st.numero) + '__' + mdjConsolidadosNormalizeText_(st.nombre);
    map[key] = mdjConsolidadosResolveStudentScore_(st, periodo);
  });

  return map;
}

function mdjConsolidadosBuildRows_(libroKey, periodo, options) {
  const opts = options || {};
  const ss = opts.ss || mdjConsolidadosGetSpreadsheet_(libroKey);
  const baseStudents = opts.baseStudents || mdjConsolidadosGetBaseStudents_(libroKey, ss);
  const courses = opts.courses || mdjConsolidadosGetCourses_(libroKey, ss);

  const courseMaps = {};
  if (String(periodo).toUpperCase() === 'FINAL') {
    courses.forEach(function(course) {
      courseMaps[course.name] = mdjConsolidadosBuildScoreMapByCourse_(libroKey, periodo, course.name, { ss: ss });
    });
  } else {
    const sh = ss.getSheetByName(periodo);
    if (!sh) throw new Error('No existe la hoja ' + periodo + '.');

    courses.forEach(function(course) {
      courseMaps[course.name] = mdjConsolidadosBuildScoreMapByCourse_(libroKey, periodo, course.name, {
        ss: ss,
        sheet: sh,
        courses: courses,
        courseMeta: course
      });
    });
  }

  return baseStudents.map(function(student) {
    const row = {
      rowNumber: student.rowNumber,
      numero: student.numero,
      nombre: student.nombre,
      materias: {},
      promedioGeneral: ''
    };

    const nums = [];
    courses.forEach(function(course) {
      const val = courseMaps[course.name][student.studentKey];
      row.materias[course.name] = (val === undefined || val === null) ? '' : val;
      const n = mdjConsolidadosToNumber_(val);
      if (n !== null) nums.push(n);
    });

    row.promedioGeneral = nums.length ? mdjConsolidadosRound2_(nums.reduce(function(a, b) { return a + b; }, 0) / nums.length) : '';
    return row;
  });
}

function mdjConsolidadosBuildData_(libroKey, periodo) {
  const libro = mdjConsolidadosResolveLibro_(libroKey);
  if (!libro) throw new Error('Libro no encontrado.');

  const normalizedPeriodo = String(periodo || '').trim().toUpperCase();
  if (MDJ_CONSOLIDADOS_CONFIG.periodos.indexOf(normalizedPeriodo) === -1) {
    throw new Error('Periodo no válido: ' + periodo);
  }

  const ss = mdjConsolidadosGetSpreadsheet_(libroKey);
  const baseSheet = ss.getSheetByName('I BIMESTRE') || ss.getSheets()[0];
  if (!baseSheet) throw new Error('No hay hojas disponibles en el libro.');

  const courses = mdjConsolidadosGetCoursesFromSheet_(baseSheet);
  const baseStudents = mdjConsolidadosGetBaseStudentsFromSheet_(baseSheet);
  const rows = mdjConsolidadosBuildRows_(libroKey, normalizedPeriodo, {
    ss: ss,
    courses: courses,
    baseStudents: baseStudents
  });
  const anio = new Date().getFullYear();

  return {
    libro: libro,
    periodo: normalizedPeriodo,
    periodoLabel: mdjConsolidadosBuildPeriodoLabel_(normalizedPeriodo),
    anio: anio,
    generatedAt: new Date(),
    logoDataUri: (typeof MDJ_LOGO_DATA_URI !== 'undefined' && MDJ_LOGO_DATA_URI) ? MDJ_LOGO_DATA_URI : '',
    courses: courses.map(function(c) { return c.name; }),
    rows: rows,
    totalAlumnos: rows.length
  };
}

function mdjConsolidadosBuildHtml_(data) {
  const coursesHtml = data.courses.map(function(course) {
    return '<th>' + mdjConsolidadosSafeHtml_(course) + '</th>';
  }).join('');

  const rowsHtml = data.rows.map(function(row, idx) {
    const materiasHtml = data.courses.map(function(course) {
      return '<td>' + mdjConsolidadosSafeHtml_(row.materias[course]) + '</td>';
    }).join('');

    return '' +
      '<tr>' +
        '<td>' + (idx + 1) + '</td>' +
        '<td>' + mdjConsolidadosSafeHtml_(row.numero) + '</td>' +
        '<td class="student">' + mdjConsolidadosSafeHtml_(row.nombre) + '</td>' +
        materiasHtml +
        '<td class="avg">' + mdjConsolidadosSafeHtml_(row.promedioGeneral) + '</td>' +
      '</tr>';
  }).join('');

  const logoHtml = data.logoDataUri
    ? '<img class="logo" src="' + data.logoDataUri + '" alt="Logo">'
    : '<div class="logo-fallback">MDJ</div>';

  return '' +
    '<!DOCTYPE html>' +
    '<html lang="es">' +
    '<head>' +
      '<meta charset="utf-8">' +
      '<style>' +
        '@page{size:A4 landscape;margin:20px 18px 26px 18px;}' +
        'body{font-family:Arial,sans-serif;color:#1d2b3a;font-size:10px;margin:0;padding:0;}' +
        '.header{display:flex;align-items:center;gap:14px;margin-bottom:12px;border-bottom:2px solid #1f6fb2;padding-bottom:10px;}' +
        '.logo,.logo-fallback{width:62px;height:62px;border-radius:12px;}' +
        '.logo{object-fit:contain;border:1px solid #dbe7f3;padding:4px;background:#fff;}' +
        '.logo-fallback{display:flex;align-items:center;justify-content:center;background:#1f6fb2;color:#fff;font-weight:700;}' +
        '.header h1{font-size:20px;margin:0 0 4px 0;color:#0f4c81;}' +
        '.header p{margin:0;font-size:11px;color:#5b708a;}' +
        '.meta{display:grid;grid-template-columns:repeat(4,1fr);gap:8px;margin-bottom:12px;}' +
        '.meta-box{border:1px solid #dbe7f3;border-radius:10px;padding:8px;background:#f8fbff;}' +
        '.meta-label{font-size:9px;font-weight:700;color:#66788f;text-transform:uppercase;margin-bottom:3px;}' +
        '.meta-value{font-size:11px;font-weight:700;color:#0f4c81;}' +
        'table{width:100%;border-collapse:collapse;font-size:9px;}' +
        'thead th{background:#eaf3fb;color:#355270;border:1px solid #d4e2ef;padding:6px 4px;text-transform:uppercase;}' +
        'tbody td{border:1px solid #e2ebf4;padding:5px 4px;text-align:center;}' +
        'tbody td.student{text-align:left;min-width:160px;}' +
        'tbody td.avg{font-weight:700;color:#0f4c81;}' +
        'tbody tr:nth-child(even){background:#fbfdff;}' +
        '.footer{margin-top:18px;display:grid;grid-template-columns:1fr 1fr;gap:40px;}' +
        '.sign{padding-top:26px;text-align:center;}' +
        '.sign-line{border-top:1px solid #333;margin-bottom:6px;}' +
        '.generated{margin-top:8px;font-size:9px;color:#66788f;text-align:right;}' +
      '</style>' +
    '</head>' +
    '<body>' +
      '<div class="header">' +
        logoHtml +
        '<div>' +
          '<h1>' + mdjConsolidadosSafeHtml_(typeof MDJ_APP !== 'undefined' && MDJ_APP.nombre ? MDJ_APP.nombre : 'Colegio Mundo de Juguetes') + '</h1>' +
          '<p>Consolidado general de calificaciones por grado</p>' +
        '</div>' +
      '</div>' +
      '<div class="meta">' +
        '<div class="meta-box"><div class="meta-label">Grado</div><div class="meta-value">' + mdjConsolidadosSafeHtml_(data.libro.nombre) + '</div></div>' +
        '<div class="meta-box"><div class="meta-label">Periodo</div><div class="meta-value">' + mdjConsolidadosSafeHtml_(data.periodoLabel) + '</div></div>' +
        '<div class="meta-box"><div class="meta-label">Año</div><div class="meta-value">' + mdjConsolidadosSafeHtml_(data.anio) + '</div></div>' +
        '<div class="meta-box"><div class="meta-label">Total estudiantes</div><div class="meta-value">' + mdjConsolidadosSafeHtml_(data.totalAlumnos) + '</div></div>' +
      '</div>' +
      '<table>' +
        '<thead>' +
          '<tr>' +
            '<th>No.</th>' +
            '<th>Código</th>' +
            '<th>Estudiante</th>' +
            coursesHtml +
            '<th>Promedio</th>' +
          '</tr>' +
        '</thead>' +
        '<tbody>' + rowsHtml + '</tbody>' +
      '</table>' +
      '<div class="footer">' +
        '<div class="sign"><div class="sign-line"></div><div>' + mdjConsolidadosSafeHtml_(MDJ_CONSOLIDADOS_CONFIG.signatureLabels.maestro) + '</div></div>' +
        '<div class="sign"><div class="sign-line"></div><div>' + mdjConsolidadosSafeHtml_(MDJ_CONSOLIDADOS_CONFIG.signatureLabels.direccion) + '</div></div>' +
      '</div>' +
      '<div class="generated">Generado el ' + mdjConsolidadosSafeHtml_(Utilities.formatDate(data.generatedAt, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss')) + '</div>' +
    '</body>' +
    '</html>';
}

function mdjConsolidadosBuildBitacora_(file, data, folderInfo) {
  return {
    ok: true,
    libroKey: data.libro.key,
    libroNombre: data.libro.nombre,
    periodo: data.periodo,
    anio: data.anio,
    fileId: file.getId(),
    fileName: file.getName(),
    fileUrl: file.getUrl(),
    folderId: folderInfo.gradeFolder.getId(),
    folderUrl: folderInfo.gradeFolder.getUrl(),
    rootFolderId: folderInfo.rootFolder.getId(),
    rootFolderUrl: folderInfo.rootFolder.getUrl(),
    totalAlumnos: data.totalAlumnos,
    generatedAt: data.generatedAt
  };
}

function mdjConsolidadosPreview(token, libroKey, periodo) {
  mdjConsolidadosRequireAdmin_(token);
  const data = mdjConsolidadosBuildData_(libroKey, periodo);
  return {
    ok: true,
    libro: data.libro,
    periodo: data.periodo,
    periodoLabel: data.periodoLabel,
    anio: data.anio,
    totalAlumnos: data.totalAlumnos,
    courses: data.courses,
    rows: data.rows,
    html: mdjConsolidadosBuildHtml_(data)
  };
}

function mdjGenerarPdfConsolidadoGrado(token, libroKey, periodo) {
  mdjConsolidadosRequireAdmin_(token);

  if (typeof mdjLog_ === 'function') {
    mdjLog_('INFO', 'mdjGenerarPdfConsolidadoGrado:start', { libroKey: libroKey, periodo: periodo });
  }

  const data = mdjConsolidadosBuildData_(libroKey, periodo);
  const folderInfo = mdjConsolidadosGetGradeFolder_(libroKey);
  const fileName = mdjConsolidadosBuildFileName_(data.libro, data.periodo, data.anio);
  const html = mdjConsolidadosBuildHtml_(data);
  const blob = HtmlService.createHtmlOutput(html).getBlob().getAs(MimeType.PDF).setName(fileName);

  mdjConsolidadosTrashExistingFilesByName_(folderInfo.gradeFolder, fileName);
  const file = folderInfo.gradeFolder.createFile(blob);
  const bitacora = mdjConsolidadosBuildBitacora_(file, data, folderInfo);

  if (typeof mdjLog_ === 'function') {
    mdjLog_('INFO', 'mdjGenerarPdfConsolidadoGrado:ok', {
      libroKey: libroKey,
      periodo: data.periodo,
      fileId: bitacora.fileId,
      fileName: bitacora.fileName,
      totalAlumnos: bitacora.totalAlumnos
    });
  }

  return bitacora;
}

function mdjGenerarTodosPdfConsolidadoGrado(token, libroKey) {
  mdjConsolidadosRequireAdmin_(token);

  const detalle = [];
  let generados = 0;
  let errores = 0;

  MDJ_CONSOLIDADOS_CONFIG.periodos.forEach(function(periodo) {
    try {
      const result = mdjGenerarPdfConsolidadoGrado(token, libroKey, periodo);
      generados++;
      detalle.push(result);
    } catch (error) {
      errores++;
      detalle.push({
        ok: false,
        libroKey: libroKey,
        periodo: periodo,
        error: error.message
      });
    }
  });

  const folderInfo = mdjConsolidadosGetGradeFolder_(libroKey);

  return {
    ok: errores === 0,
    libro: folderInfo.libro,
    generados: generados,
    errores: errores,
    folderId: folderInfo.gradeFolder.getId(),
    folderUrl: folderInfo.gradeFolder.getUrl(),
    rootFolderId: folderInfo.rootFolder.getId(),
    rootFolderUrl: folderInfo.rootFolder.getUrl(),
    detalle: detalle
  };
}

function mdjListConsolidadosPorLibro(token, libroKey) {
  mdjConsolidadosRequireAdmin_(token);
  const folderInfo = mdjConsolidadosGetGradeFolder_(libroKey);
  const files = folderInfo.gradeFolder.getFiles();
  const items = [];

  while (files.hasNext()) {
    const file = files.next();
    if (file.isTrashed()) continue;
    items.push({
      id: file.getId(),
      name: file.getName(),
      url: file.getUrl(),
      dateCreated: file.getDateCreated(),
      lastUpdated: file.getLastUpdated(),
      size: file.getSize()
    });
  }

  items.sort(function(a, b) {
    return new Date(b.lastUpdated).getTime() - new Date(a.lastUpdated).getTime();
  });

  return {
    ok: true,
    libro: folderInfo.libro,
    folderId: folderInfo.gradeFolder.getId(),
    folderUrl: folderInfo.gradeFolder.getUrl(),
    rootFolderId: folderInfo.rootFolder.getId(),
    rootFolderUrl: folderInfo.rootFolder.getUrl(),
    total: items.length,
    items: items
  };
}

function mdjDiagnosticoConsolidados(token, libroKey) {
  mdjConsolidadosRequireAdmin_(token);
  const folderInfo = mdjConsolidadosGetGradeFolder_(libroKey);
  const data = mdjConsolidadosBuildData_(libroKey, 'I BIMESTRE');

  return {
    ok: true,
    libro: folderInfo.libro,
    rootFolderId: folderInfo.rootFolder.getId(),
    rootFolderName: folderInfo.rootFolder.getName(),
    rootFolderUrl: folderInfo.rootFolder.getUrl(),
    gradeFolderId: folderInfo.gradeFolder.getId(),
    gradeFolderName: folderInfo.gradeFolder.getName(),
    gradeFolderUrl: folderInfo.gradeFolder.getUrl(),
    periodos: MDJ_CONSOLIDADOS_CONFIG.periodos,
    totalAlumnosDetectados: data.totalAlumnos,
    totalMateriasDetectadas: data.courses.length,
    materias: data.courses
  };
}

function mdjAbrirCarpetaConsolidados(token, libroKey) {
  mdjConsolidadosRequireAdmin_(token);
  const folderInfo = mdjConsolidadosGetGradeFolder_(libroKey);
  return {
    ok: true,
    libro: folderInfo.libro,
    folderId: folderInfo.gradeFolder.getId(),
    folderUrl: folderInfo.gradeFolder.getUrl(),
    rootFolderId: folderInfo.rootFolder.getId(),
    rootFolderUrl: folderInfo.rootFolder.getUrl()
  };
}