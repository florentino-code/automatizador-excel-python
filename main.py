import openpyxl
from tkinter import Tk, filedialog
from collections import defaultdict
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


# Ocultar ventana principal de tkinter
root = Tk()
root.withdraw()


# ==============================
# SELECCIONAR ARCHIVO EXCEL
# ==============================

archivo = filedialog.askopenfilename(
    title="Selecciona el archivo de ventas",
    filetypes=[("Archivos Excel", "*.xlsx")]
)

if not archivo:
    print("No seleccionaste ningún archivo.")
    exit()


# ==============================
# ABRIR EXCEL
# ==============================

wb = openpyxl.load_workbook(archivo)

if "Ventas" not in wb.sheetnames:
    print("No existe una hoja llamada 'Ventas'.")
    exit()

ws = wb["Ventas"]

# ==============================
# DETECTAR COLUMNAS AUTOMÁTICAMENTE
# ==============================

columnas = {}

for celda in ws[1]:
    if celda.value:
        nombre = str(celda.value).strip().lower()
        columnas[nombre] = celda.column

print("COLUMNAS DETECTADAS:")
print(columnas) 


# ==============================
# VARIABLES
# ==============================

total_ventas = 0
numero_ventas = 0

ventas_producto = defaultdict(float)
ventas_cliente = defaultdict(float)


# ==============================
# LEER LAS 200 FILAS
# ==============================

for fila in range(2, ws.max_row + 1):

    fecha = ws.cell(fila, 1).value
    cliente = ws.cell(fila, 2).value
    producto = ws.cell(fila, 3).value
    venta = ws.cell(fila, 4).value

    if (
        cliente is None
        or producto is None
        or not isinstance(venta, (int, float))
    ):
        continue

    total_ventas += venta
    numero_ventas += 1

    ventas_producto[producto] += venta
    ventas_cliente[cliente] += venta


# ==============================
# COMPROBAR DATOS
# ==============================

if numero_ventas == 0:
    print("No se encontraron ventas válidas.")
    exit()


promedio_venta = total_ventas / numero_ventas

producto_mas_vendido = max(
    ventas_producto,
    key=ventas_producto.get
)

cliente_mayor_compra = max(
    ventas_cliente,
    key=ventas_cliente.get
)


# ==============================
# CREAR HOJA RESUMEN
# ==============================

if "Resumen" in wb.sheetnames:
    del wb["Resumen"]

resumen = wb.create_sheet("Resumen")


# Título
resumen.merge_cells("A1:B1")

resumen["A1"] = "RESUMEN DE VENTAS"

resumen["A1"].font = Font(
    bold=True,
    color="FFFFFF",
    size=16
)

resumen["A1"].fill = PatternFill(
    "solid",
    fgColor="1F4E78"
)

resumen["A1"].alignment = Alignment(
    horizontal="center"
)


# ==============================
# DATOS DEL RESUMEN
# ==============================

datos = [
    ("Total vendido", total_ventas),
    ("Número de ventas", numero_ventas),
    ("Promedio por venta", promedio_venta),
    ("Producto con mayores ventas", producto_mas_vendido),
    (
        "Ventas del producto",
        ventas_producto[producto_mas_vendido]
    ),
    ("Cliente con mayor compra", cliente_mayor_compra),
    (
        "Total comprado por cliente",
        ventas_cliente[cliente_mayor_compra]
    )
]


for fila, (concepto, valor) in enumerate(datos, start=3):

    resumen.cell(fila, 1).value = concepto
    resumen.cell(fila, 2).value = valor


# ==============================
# FORMATO
# ==============================

relleno = PatternFill(
    "solid",
    fgColor="D9EAF7"
)

borde = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin")
)


for fila in range(3, 10):

    resumen.cell(fila, 1).font = Font(bold=True)
    resumen.cell(fila, 1).fill = relleno

    resumen.cell(fila, 1).border = borde
    resumen.cell(fila, 2).border = borde


# Formato moneda
for fila in [3, 5, 7, 9]:
    resumen.cell(fila, 2).number_format = '"S/" #,##0.00'


# Tamaño de columnas
resumen.column_dimensions["A"].width = 30
resumen.column_dimensions["B"].width = 28


# ==============================
# GUARDAR REPORTE
# ==============================

archivo_salida = filedialog.asksaveasfilename(
    title="Guardar reporte",
    defaultextension=".xlsx",
    filetypes=[("Archivo Excel", "*.xlsx")],
    initialfile="reporte_ventas_200.xlsx"
)

if not archivo_salida:
    print("No seleccionaste dónde guardar el reporte.")
    exit()


wb.save(archivo_salida)


# ==============================
# RESULTADO
# ==============================

print()
print("REPORTE GENERADO CORRECTAMENTE")
print("-------------------------------")
print(f"Ventas procesadas: {numero_ventas}")
print(f"Total vendido: S/ {total_ventas:,.2f}")
print(f"Promedio por venta: S/ {promedio_venta:,.2f}")
print(f"Producto con mayores ventas: {producto_mas_vendido}")
print(f"Cliente con mayor compra: {cliente_mayor_compra}")
print()
print("Archivo guardado en:")
print(archivo_salida)