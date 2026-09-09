# Automatizador de reportes de ventas con Python

Programa desarrollado en Python para procesar automáticamente archivos Excel de ventas.

## Funciones

- Seleccionar un archivo Excel.
- Calcular subtotales automáticamente.
- Calcular el total vendido.
- Calcular unidades vendidas.
- Identificar el producto más vendido.
- Crear una hoja de resumen.
- Aplicar formato profesional.
- Guardar el reporte generado en Excel.

## Tecnologías utilizadas

- Python
- openpyxl
- tkinter

## Formato esperado del Excel

El archivo debe contener una hoja llamada `Ventas` con las columnas:

| Producto | Cantidad | Precio |
|---|---:|---:|
| Coca Cola 3L | 5 | 12.50 |
| Arroz 5kg | 3 | 22.00 |

## Ejecución

Instalar la dependencia:

```bash
pip install openpyxl