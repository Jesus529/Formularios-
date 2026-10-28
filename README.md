# Sistema de Gestión de Préstamos — Laboratorio 11

Aplicación académica en Python + Flask + SQLite para el caso de solicitud de préstamo personal de una cooperativa financiera.

## Incluye
- Registro de clientes y solicitudes, con adjuntos de documentos.
- Revisión de documentación (completa/incompleta) y comentarios.
- Evaluación de riesgo (ingresos, deudas, puntaje, historial, recomendación y justificación).
- Decisión de aprobación/rechazo con validaciones básicas del flujo.
- Registro de aceptación y firmas por nombre (demostración; no es firma digital legal).
- Registro de desembolso simulado (no realiza transferencias reales).
- Panel con estados y exportación Excel (.xlsx).
- Base de datos SQLite creada automáticamente en `instance/prestamos.db`.

## Requisitos
Python 3.10 o superior.

## Ejecutar en Windows PowerShell
1. Abre PowerShell en la carpeta `prestamos_cooperativa`.
2. Crea entorno virtual:
   ```powershell
   py -m venv .venv
   ```
3. Instala dependencias:
   ```powershell
   .\.venv\Scripts\python.exe -m pip install -r requirements.txt
   ```
4. Ejecuta:
   ```powershell
   .\.venv\Scripts\python.exe app.py
   ```
5. Abre `http://127.0.0.1:5000`.

Si el comando `py` no funciona, prueba `python -m venv .venv`.

## Importante
Es una demo educativa: no procesa pagos reales, no envía notificaciones, no genera PDF contractual ni implementa firma digital certificada. Para uso real se necesitarían autenticación y roles, controles de acceso, cifrado y gestión segura de documentos, auditoría y revisión legal.
