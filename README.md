# Control_301_LID

Sistema web responsive para el control de estudiantes, docentes, horarios y asistencia del LID 301.

## Arquitectura
- Frontend público: GitHub Pages
- Aplicación/Backend: Google Apps Script
- Base de datos: Google Sheets
- Verificación física: Raspberry Pi en la red Wi‑Fi del laboratorio

## Acceso
La identidad sigue siendo la cuenta Google/UPCH seleccionada. El backend valida si corresponde a ADMIN, DOCENTE o ESTUDIANTE.

## GitHub Pages
URL prevista: https://ulewis.github.io/Control_301_LID/

El frontend es responsive para celular, tablet y escritorio. El Web App de Apps Script debe mantener `HtmlService.XFrameOptionsMode.ALLOWALL` para poder renderizarse dentro del sitio.
