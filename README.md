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
No hay acceso mediante iframe a otro Web App.

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
- `AdminAsistencia.html`, `TomarAsistencia.html`: interfaces integradas.
- Otros `.js`: boletas, consolidados y utilidades del sistema anterior.

## Validación

`node --check` para cada `.js`; inspeccionar duplicados de funciones globales.
La conexión, permisos, datos y generación real de documentos **requieren pruebas
contra copias de Sheets y Drive en Google Apps Script**.
