from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from tkinter import Tk, filedialog

# =========================
# SELECCIONAR ARCHIVO EXCEL
# =========================

ventana = Tk()
ventana.withdraw()

archivo = filedialog.askopenfilename(
    title="Selecciona el archivo de ventas",
    filetypes=[("Archivos de Excel", "*.xlsx")]
)

if not archivo:
    print("No seleccionaste ningún archivo.")
    exit()

# =========================
# ABRIR ARCHIVO
# =========================

libro = load_workbook(archivo)

if "Ventas" not in libro.sheetnames:
    print("El archivo no tiene una hoja llamada 'Ventas'.")
    exit()

hoja = libro["Ventas"]

# =========================
# VARIABLES
# =========================

total_general = 0
mayor_cantidad = 0
producto_mas_vendido = ""
unidades_totales = 0

# Crear encabezado Subtotal
hoja["D1"] = "Subtotal"

# =========================
# PROCESAR VENTAS
# =========================

for fila in range(2, hoja.max_row + 1):

    producto = hoja[f"A{fila}"].value
    cantidad = hoja[f"B{fila}"].value
    precio = hoja[f"C{fila}"].value

    # Validar que cantidad y precio sean números
    if not isinstance(cantidad, (int, float)):
        continue

    if not isinstance(precio, (int, float)):
        continue

    subtotal = cantidad * precio

    hoja[f"D{fila}"] = subtotal

    total_general += subtotal
    unidades_totales += cantidad

    # Buscar producto con mayor cantidad
    if cantidad > mayor_cantidad:
        mayor_cantidad = cantidad
        producto_mas_vendido = producto

# =========================
# TOTAL GENERAL
# =========================

fila_total = hoja.max_row + 2

hoja[f"C{fila_total}"] = "TOTAL"
hoja[f"D{fila_total}"] = total_general

# =========================
# FORMATO HOJA VENTAS
# =========================

relleno_encabezado = PatternFill(
    fill_type="solid",
    fgColor="1F4E78"
)

borde_fino = Side(
    style="thin",
    color="808080"
)

borde = Border(
    left=borde_fino,
    right=borde_fino,
    top=borde_fino,
    bottom=borde_fino
)

# Formato encabezados
for celda in hoja[1]:

    celda.font = Font(
        bold=True,
        color="FFFFFF"
    )

    celda.fill = relleno_encabezado

    celda.alignment = Alignment(
        horizontal="center",
        vertical="center"
    )

    celda.border = borde

# Formato filas de ventas
for fila in range(2, hoja.max_row + 1):

    for columna in ["A", "B", "C", "D"]:
        hoja[f"{columna}{fila}"].border = borde

    hoja[f"B{fila}"].alignment = Alignment(horizontal="center")

    hoja[f"C{fila}"].number_format = '"S/" #,##0.00'
    hoja[f"D{fila}"].number_format = '"S/" #,##0.00'

# Formato total
hoja[f"C{fila_total}"].font = Font(bold=True)
hoja[f"D{fila_total}"].font = Font(bold=True)

hoja[f"C{fila_total}"].border = borde
hoja[f"D{fila_total}"].border = borde

hoja[f"D{fila_total}"].number_format = '"S/" #,##0.00'

# Ajustar columnas
hoja.column_dimensions["A"].width = 25
hoja.column_dimensions["B"].width = 12
hoja.column_dimensions["C"].width = 15
hoja.column_dimensions["D"].width = 15

hoja.row_dimensions[1].height = 25

# =========================
# CREAR HOJA RESUMEN
# =========================

if "Resumen" in libro.sheetnames:
    del libro["Resumen"]

resumen = libro.create_sheet("Resumen")

# =========================
# TÍTULO
# =========================

resumen["A1"] = "RESUMEN DE VENTAS"

resumen.merge_cells("A1:B1")

resumen["A1"].font = Font(
    bold=True,
    size=16,
    color="FFFFFF"
)

resumen["A1"].fill = relleno_encabezado

resumen["A1"].alignment = Alignment(
    horizontal="center",
    vertical="center"
)

resumen.row_dimensions[1].height = 30

# =========================
# DATOS DEL RESUMEN
# =========================

resumen["A3"] = "Total vendido"
resumen["B3"] = total_general

resumen["A4"] = "Producto más vendido"
resumen["B4"] = producto_mas_vendido

resumen["A5"] = "Cantidad vendida"
resumen["B5"] = mayor_cantidad

resumen["A6"] = "Unidades totales"
resumen["B6"] = unidades_totales

# =========================
# FORMATO RESUMEN
# =========================

resumen["B3"].number_format = '"S/" #,##0.00'

for fila in range(3, 7):

    resumen[f"A{fila}"].font = Font(bold=True)

    resumen[f"A{fila}"].border = borde
    resumen[f"B{fila}"].border = borde

    resumen[f"B{fila}"].alignment = Alignment(
        horizontal="center"
    )

resumen.column_dimensions["A"].width = 25
resumen.column_dimensions["B"].width = 20

# =========================
# GUARDAR ARCHIVO
# =========================

archivo_salida = filedialog.asksaveasfilename(
    title="Guardar reporte de ventas",
    defaultextension=".xlsx",
    filetypes=[("Archivos de Excel", "*.xlsx")],
    initialfile="reporte_ventas.xlsx"
)

if not archivo_salida:
    print("No se guardó el reporte.")
    exit()

libro.save(archivo_salida)

# =========================
# RESULTADO EN TERMINAL
# =========================

print("Reporte generado correctamente")
print(f"Total vendido: S/ {total_general:.2f}")
print(f"Producto más vendido: {producto_mas_vendido}")
print(f"Cantidad vendida: {mayor_cantidad}")
print(f"Unidades totales: {unidades_totales}")