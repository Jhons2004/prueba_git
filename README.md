# Colegio Mundo de Juguetes — Sistema Central Unificado (QA)

Aplicación **única** de Google Apps Script que integra gestión académica, matrícula/padrón,
calificaciones, boletas, consolidados y control de asistencia. El módulo de asistencia
mantiene los registros diarios en una **hoja maestra propia**.

## Páginas

- `?page=login`: inicio de sesión central.
- `?page=admin`: Dirección / administración académica.
- `?page=maestro`: panel docente.
- `?page=asistencia`: toma diaria de asistencia (docentes autorizados / Dirección).
- `?page=asistencia-admin`: reportes, sincronización y usuarios (solo Dirección/admin).

El usuario y la sesión son **únicos** (`mdjLogin`, `mdjGetSession`, `mdj_token`).
La asistencia consulta `MDJ_LIBROS` del sistema central y usa el mismo logo `MDJ_LOGO_DATA_URI`.
**Asistencia administrativa y docente se muestran en sus respectivos paneles centrales**, sin iframe ni navegación hacia otro Web App.
En Dirección el menú se encuentra a la izquierda y puede plegarse/desplegarse con **☰**; su estado se conserva por navegador.

## Instalación segura en Apps Script QA

1. Crear un **proyecto nuevo** en script.google.com (no modificar los proyectos originales).
2. Clonar este repositorio y ejecutar `clasp create --type webapp` **solo para crear el archivo `.clasp.json`** en esta carpeta, o asociarlo al proyecto QA mediante `clasp clone-script <SCRIPT_ID_QA>` en una carpeta vacía y copiar aquí sus archivos fuente.
3. Activar *Google Apps Script API* y la API avanzada de **Google Drive** versión v3.
4. Configurar **Propiedades del script** en Configuración del proyecto:
   - `MDJ_ASISTENCIA_SPREADSHEET_ID`: **ID de una copia de pruebas** de la hoja maestra de asistencia. Sin esta propiedad, el módulo no puede abrir o guardar asistencias.
   - `MDJ_CONTROL_SPREADSHEET_ID` (**obligatorio para login QA**): ID de copia de la hoja de control de usuarios y asignaciones.
   - `MDJ_LIBRO_<CLAVE_EN_MAYÚSCULA>`: ID de la copia QA de cada libro de notas (12 claves, como `MDJ_LIBRO_PRIMERO_BASICO`).
   - `MDJ_BOLETAS_<CLAVE_EN_MAYÚSCULA>`: ID de carpeta Drive QA para boletas de cada grado.
   - `MDJ_CONSOLIDADOS_ROOT_FOLDER_ID`: carpeta raíz de consolidados QA.
   - **No existen IDs de producción incluidos en este repositorio público**. Si falta una propiedad, la función correspondiente no tendrá acceso al recurso.
5. Desplegar en Apps Script con acceso restringido a cuentas autorizadas, según corresponda.
6. Ejecutar `mdjQaDiagnosticoUnificacion()` (diagnóstico solo lectura, no toca registros).
7. Probar con un admin y docente, sincronizar un grado y validar que la hoja destino sea la copia QA.

> **Precaución:** No se incorporan identificadores de producción en el repositorio. Evita `clasp push` sobre cualquiera de los proyectos originales. Debes configurar copias de **todos** los libros y destinos antes de ensayar escrituras o generación masiva.

## Arquitectura

- `Code.js`, `Config.js`, `Auth.js`: sesión, navegación, grados y académico.
- `Administrativo.html`, `Maestro.html`, `Login.html`: interfaz central.
- `AsistenciaConfig.js`, `AsistenciaAuth.js`, `AsistenciaCodigo.js`, `AsistenciaHelpers.js`, `AsistenciaLogo.js`, `Asistencia.js`: módulo integrado.
- `AdminAsistenciaIntegrada.html`, `TomarAsistenciaIntegrada.html`: vistas incrustadas nativamente en `Administrativo.html` y `Maestro.html`, con estilos y JavaScript aislados por módulo.
- `AdminAsistencia.html`, `TomarAsistencia.html`: páginas originales conservadas únicamente como referencia/compatibilidad, no son el destino del menú central.
- Otros `.js`: boletas, consolidados y utilidades del sistema anterior.

## Validación

`node --check` para cada `.js`; inspeccionar duplicados de funciones globales.
La conexión, permisos, datos y generación real de documentos **requieren pruebas
contra copias de Sheets y Drive en Google Apps Script**.

## Navegación integrada (revisión octubre 2026)

- Dirección: ingresar normalmente a `?page=admin`, presionar **Asistencia administrativa** en la barra lateral. Dentro de la misma página aparecen **Reportes**, **Tomar asistencia**, **Sincronizar alumnos** y **Usuarios**. La barra lateral se pliega mediante **☰** sin perder el módulo seleccionado.
- Docentes: ingresar a `?page=maestro`, seleccionar **Tomar asistencia**; el registro aparece dentro del mismo panel. **Volver a calificaciones** restaura la vista anterior sin perder la sesión.
- Los enlaces antiguos `?page=asistencia-admin` y `?page=asistencia` abren el panel central correspondiente con el módulo seleccionado.
- Se conservaron el logo institucional, la autenticación, las funciones del servidor de asistencia y la configuración por propiedades de Apps Script.

### Actualizar proyecto Apps Script QA existente

```powershell
cd "$HOME\Documents\prueba_git"
git pull origin main
clasp status
clasp push
```

Después, **Implementar → Administrar implementaciones → Editar → Nueva versión → Implementar**. Si se usa la URL `/dev` en sesión de editor, los cambios guardados pueden revisarse sin generar una nueva versión. Para datos de producción se recomienda restringir permisos y comprobar con una cuenta de cada rol.

> Estas son modificaciones de interfaz verificadas con análisis estático y `node --check`. La prueba de navegación autenticada con Google Apps Script y las lecturas/escrituras reales requiere un despliegue de QA.