from openpyxl import Workbook

# Crear un nuevo libro de Excel
libro = Workbook()

# Seleccionar la hoja activa
hoja = libro.active

# Cambiar el nombre de la hoja
hoja.title = "Ventas"

# Crear encabezados
hoja["A1"] = "Producto"
hoja["B1"] = "Cantidad"
hoja["C1"] = "Precio"

# Agregar datos de ejemplo
hoja.append(["Coca Cola 3L", 5, 12.50])
hoja.append(["Arroz 5kg", 3, 22.00])
hoja.append(["Aceite 1L", 8, 9.50])
hoja.append(["Leche", 10, 4.50])

# Guardar el archivo
libro.save("ventas.xlsx")

print("Excel creado correctamente")