/**
 * MaestroModulo.gs
 * Módulo abierto para docentes:
 * - usa la validación del login por rol maestro
 * - NO valida por asignaciones de cursos
 * - permite buscar por año, grado, bimestre y curso
 * - permite consultar y guardar notas
 *
 * NO tocar Auth.gs.
 */

function mdjGetTeacherOpenPanelData(token) {
  const session = mdjRequireRole_(token, ['maestro', 'admin']);

  const libros = mdjGetLibros().map(function(libro) {
    let academicYear = '';
    try {
      if (typeof mdjGetAcademicYearForLibro_ === 'function') {
        academicYear = mdjGetAcademicYearForLibro_(libro.key);
      }
    } catch (e) {
      academicYear = '';
    }

    return {
      key: libro.key,
      nombre: libro.nombre,
      nivel: libro.nivel,
      grado: libro.grado,
      programa: libro.programa,
      academicYear: academicYear || new Date().getFullYear()
    };
  });

  const years = Array.from(
    new Set(
      libros.map(function(item) {
        return String(item.academicYear || '');
      }).filter(Boolean)
    )
  ).sort();

  return {
    ok: true,
    user: session.user,
    libros: libros,
    years: years,
    bimestres: (MDJ_APP.bimestres || []).filter(function(b) {
      return b !== 'CONSOLIDADO';
    })
  };
}

function mdjGetTeacherOpenCourses(token, libroKey, bimestre) {
  mdjRequireRole_(token, ['maestro', 'admin']);

  try {
    const libro = mdjGetLibroByKey(libroKey);
    if (!libro) {
      return { ok: false, cursos: [], message: 'Libro no encontrado.' };
    }

    const ss = mdjGetSpreadsheetByKey_(libroKey);
    const sh = ss.getSheetByName(bimestre);
    if (!sh) {
      return { ok: false, cursos: [], message: 'No existe la hoja ' + bimestre + '.' };
    }

    const cursos = mdjExtractCoursesFromSheet_(sh);

    return {
      ok: true,
      libro: libro,
      cursos: cursos
    };
  } catch (e) {
    return {
      ok: false,
      cursos: [],
      message: e && e.message ? e.message : String(e)
    };
  }
}

function mdjGetTeacherOpenCourseStudents(token, libroKey, bimestre, curso) {
  mdjRequireRole_(token, ['maestro', 'admin']);

  try {
    const libro = mdjGetLibroByKey(libroKey);
    if (!libro) {
      return { ok: false, alumnos: [], message: 'Libro no encontrado.' };
    }

    const ss = mdjGetSpreadsheetByKey_(libroKey);
    const sh = ss.getSheetByName(bimestre);
    if (!sh) {
      return { ok: false, alumnos: [], message: 'No existe la hoja ' + bimestre + '.' };
    }

    const courseMeta = mdjFindCourseInSheet_(sh, curso);
    if (!courseMeta) {
      return { ok: false, alumnos: [], message: 'No se encontró el curso "' + curso + '" en ' + bimestre + '.' };
    }

    const students = mdjExtractStudentsForCourse_(sh, courseMeta);

    return {
      ok: true,
      libro: libro,
      bimestre: bimestre,
      curso: courseMeta,
      alumnos: students
    };
  } catch (e) {
    return {
      ok: false,
      alumnos: [],
      message: e && e.message ? e.message : String(e)
    };
  }
}

function mdjSaveTeacherOpenCourseGrades(token, payload) {
  mdjRequireRole_(token, ['maestro', 'admin']);

  payload = payload || {};
  const libroKey = String(payload.libroKey || '').trim();
  const bimestre = String(payload.bimestre || '').trim();
  const curso = String(payload.curso || '').trim();
  const alumnos = Array.isArray(payload.alumnos) ? payload.alumnos : [];

  if (!libroKey || !bimestre || !curso) {
    return { ok: false, message: 'Faltan datos para guardar notas.' };
  }

  try {
    const ss = mdjGetSpreadsheetByKey_(libroKey);
    const sh = ss.getSheetByName(bimestre);
    if (!sh) {
      return { ok: false, message: 'No se encontró la hoja ' + bimestre + '.' };
    }

    const courseMeta = mdjFindCourseInSheet_(sh, curso);
    if (!courseMeta) {
      return { ok: false, message: 'No se encontró el curso "' + curso + '".' };
    }

    const cols = courseMeta.columns;

    alumnos.forEach(function(item) {
      const row = Number(item.rowNumber || 0);
      if (!row) return;

      const updates = [];

      function pushIf(col, val) {
        if (col) {
          updates.push({
            col: col,
            value: mdjNormalizeGradeInput_(val)
          });
        }
      }

      pushIf(cols.a1, item.a1);
      pushIf(cols.a2, item.a2);
      pushIf(cols.a3, item.a3);
      pushIf(cols.a4, item.a4);
      pushIf(cols.a5, item.a5);
      pushIf(cols.a6, item.a6);
      pushIf(cols.apObjetiva, item.apObjetiva);
      pushIf(cols.evBimestral, item.evBimestral);
      pushIf(cols.evRemedial, item.evRemedial);

      updates.forEach(function(u) {
        sh.getRange(row, u.col).setValue(u.value);
      });
    });

    SpreadsheetApp.flush();

    return {
      ok: true,
      message: 'Notas guardadas correctamente.'
    };
  } catch (e) {
    return {
      ok: false,
      message: e && e.message ? e.message : String(e)
    };
  }
}