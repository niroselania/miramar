# Miramar - gastos

App local para cargar, consultar y exportar los gastos mensuales de Miramar.

## Abrir

La app esta corriendo en:

http://127.0.0.1:8765/

Para iniciarla otro dia, ejecutar `iniciar_miramar.bat` en esta carpeta y dejar esa ventana abierta mientras se usa la app.

## Que guarda

- Base SQLite: `miramar.db`
- Exportaciones Excel: carpeta `exports`
- Importa la planilla original desde `C:\Users\ARIEL-ROSETI\Downloads\Miramar\Servicio Miramar 2026.xlsx` si la base esta vacia.

## Funciones

- Cargar un Excel `.xlsx`.
- Ver resumen mensual por concepto.
- Agregar o actualizar gastos por mes, anio y categoria.
- Borrar registros.
- Exportar nuevamente a Excel.
