# ContSmart - Aplicación Excel en Flet

Aplicación de hoja de cálculo similar a Excel creada con Flet.

## 🎯 Versiones Disponibles

### 1. **excel_datatable.py** ⭐ RECOMENDADA
La versión más moderna y eficiente usando el componente `DataTable` de Flet:
- ✨ **Tabla profesional** con bordes y líneas divisorias
- 🖱️ **Clic para editar** - Haz clic en cualquier celda para abrir un diálogo de edición
- 📊 **Visualización clara** con encabezados destacados
- 💾 **Guardar/Cargar CSV** con notificaciones visuales
- ➕ **Agregar filas/columnas** dinámicamente
- 🎨 **Diseño moderno** y profesional
- ⚡ **Mejor rendimiento** que las versiones con TextFields

### 2. Versión Básica (excel_app.py)
- 📊 Tabla editable con múltiples filas y columnas
- ✏️ Barra de fórmulas para editar celdas
- ➕ Agregar filas y columnas dinámicamente
- 🔤 Encabezados de columnas (A, B, C...) y números de fila
- 🎨 Interfaz intuitiva similar a Excel
- 🧹 Función para limpiar todos los datos

### 3. Versión Mejorada (excel_app_mejorado.py)
Incluye todas las características de la versión básica más:
- 💾 **Guardar a CSV**: Exporta tus datos a archivo CSV
- 📂 **Cargar desde CSV**: Importa datos desde archivos CSV existentes
- 🔔 **Notificaciones**: Mensajes visuales de éxito/error
- 📝 **Nombre de archivo personalizable**: Especifica el nombre del archivo

## Instalación

1. Asegúrate de tener Python 3.7+ instalado
2. Instala las dependencias:
```bash
pip install flet
```

## Uso

### Versión DataTable (Recomendada) ⭐
```bash
python excel_datatable.py
```

### Versión Básica
Ejecuta la aplicación básica con:
```bash
python excel_app.py
```

### Versión Mejorada
Ejecuta la aplicación mejorada con:
```bash
python excel_app_mejorado.py
```

## 🎮 Cómo usar la aplicación

### Versión DataTable:
1. **Editar celda**: Haz clic en cualquier celda para abrir el diálogo de edición
2. **Guardar**: Escribe el valor y presiona "Guardar" o Enter
3. **Agregar filas/columnas**: Usa los botones en la parte superior
4. **Guardar archivo**: Escribe el nombre del archivo y haz clic en "Guardar CSV"
5. **Cargar archivo**: Escribe el nombre del archivo existente y haz clic en "Cargar CSV"

### Versiones con TextField:
- **Barra de fórmulas**: Muestra y edita el contenido de la celda seleccionada
- **Celdas editables**: Haz clic en cualquier celda para editarla directamente
- **Agregar Fila**: Añade una nueva fila al final de la tabla
- **Agregar Columna**: Añade una nueva columna al final de la tabla
- **Limpiar Todo**: Borra todos los datos de la tabla
- **Guardar CSV** (versión mejorada): Guarda los datos en un archivo CSV
- **Cargar CSV** (versión mejorada): Carga datos desde un archivo CSV

## Estructura del proyecto

```
ContSmart/
├── excel_datatable.py        # ⭐ Versión recomendada con DataTable
├── excel_app.py              # Aplicación básica con TextFields
├── excel_app_mejorado.py     # Aplicación con TextFields + CSV
└── README.md                 # Este archivo
```

## Formato CSV

Los archivos CSV se guardan con la siguiente estructura:
- Primera fila: Encabezados de columna (A, B, C, ...)
- Filas siguientes: Datos de las celdas

Ejemplo:
```csv
A,B,C,D
Juan,25,Madrid,España
María,30,Barcelona,España
Pedro,28,Valencia,España
```

## Próximas mejoras sugeridas

- [ ] Guardar y cargar datos desde archivos CSV/Excel
- [ ] Soporte para fórmulas básicas (SUM, AVG, etc.)
- [ ] Copiar y pegar celdas
- [ ] Formato de celdas (colores, negrita, etc.)
- [ ] Exportar a PDF
- [ ] Deshacer/Rehacer acciones
