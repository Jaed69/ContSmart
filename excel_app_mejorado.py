import flet as ft
from flet import Colors, Border
import csv
import os


class ExcelAppMejorado:
    def __init__(self, page: ft.Page):
        self.page = page
        self.page.title = "Excel Pro - ContSmart"
        self.page.padding = 10
        
        # Configuración inicial de filas y columnas
        self.num_rows = 20
        self.num_cols = 10
        
        # Almacén de datos
        self.data = {}
        
        # Controles
        self.selected_cell = None
        self.formula_bar = ft.TextField(
            hint_text="Celda seleccionada",
            on_submit=self.update_cell_value,
            expand=True,
            text_size=14,
        )
        
        # Nombre de archivo
        self.file_name = ft.TextField(
            hint_text="nombre_archivo.csv",
            width=200,
            text_size=12,
        )
        
        self.build_ui()
    
    def build_ui(self):
        # Barra de fórmulas
        formula_container = ft.Container(
            content=ft.Row([
                ft.Text("Celda:", weight=ft.FontWeight.BOLD, width=50),
                self.formula_bar,
            ]),
            padding=10,
            bgcolor=Colors.BLUE_GREY_50,
        )
        
        # Crear la tabla
        table = self.create_table()
        
        # Contenedor con scroll para la tabla
        table_container = ft.Container(
            content=table,
            border=Border.all(1, Colors.GREY_400),
            expand=True,
        )
        
        # Botones de acción principales
        actions = ft.Row([
            ft.FilledButton(
                "Agregar Fila",
                icon=ft.Icons.ADD,
                on_click=self.add_row,
            ),
            ft.FilledButton(
                "Agregar Columna",
                icon=ft.Icons.ADD,
                on_click=self.add_column,
            ),
            ft.FilledButton(
                "Limpiar Todo",
                icon=ft.Icons.CLEAR,
                on_click=self.clear_all,
                style=ft.ButtonStyle(
                    bgcolor=Colors.RED_400,
                    color=Colors.WHITE,
                ),
            ),
        ], spacing=10)
        
        # Botones de archivo
        file_actions = ft.Row([
            self.file_name,
            ft.FilledButton(
                "Guardar CSV",
                icon=ft.Icons.SAVE,
                on_click=self.save_to_csv,
                style=ft.ButtonStyle(bgcolor=Colors.GREEN_700),
            ),
            ft.FilledButton(
                "Cargar CSV",
                icon=ft.Icons.UPLOAD_FILE,
                on_click=self.load_from_csv,
                style=ft.ButtonStyle(bgcolor=Colors.BLUE_700),
            ),
        ], spacing=10)
        
        # Layout principal
        self.page.add(
            ft.Column([
                ft.Text("Hoja de Cálculo Pro - ContSmart", 
                       size=24, 
                       weight=ft.FontWeight.BOLD),
                actions,
                file_actions,
                formula_container,
                table_container,
            ], expand=True, spacing=10)
        )
    
    def create_table(self):
        # Crear encabezados de columnas (A, B, C, ...)
        headers = [ft.Container(
            content=ft.Text("", width=40, text_align=ft.TextAlign.CENTER),
            bgcolor=Colors.GREY_300,
            border=Border.all(1, Colors.GREY_400),
        )]
        
        for col in range(self.num_cols):
            col_letter = self.get_column_letter(col)
            headers.append(
                ft.Container(
                    content=ft.Text(
                        col_letter, 
                        width=100, 
                        text_align=ft.TextAlign.CENTER,
                        weight=ft.FontWeight.BOLD,
                    ),
                    bgcolor=Colors.BLUE_GREY_100,
                    border=Border.all(1, Colors.GREY_400),
                )
            )
        
        # Crear filas
        rows = [ft.Row(headers, spacing=0)]
        
        for row in range(self.num_rows):
            row_cells = []
            
            # Número de fila
            row_cells.append(
                ft.Container(
                    content=ft.Text(
                        str(row + 1), 
                        width=40, 
                        text_align=ft.TextAlign.CENTER,
                        weight=ft.FontWeight.BOLD,
                    ),
                    bgcolor=Colors.BLUE_GREY_100,
                    border=Border.all(1, Colors.GREY_400),
                )
            )
            
            # Celdas de datos
            for col in range(self.num_cols):
                cell_id = f"{row}_{col}"
                cell_value = self.data.get(cell_id, "")
                
                cell = ft.TextField(
                    value=cell_value,
                    width=100,
                    height=35,
                    text_size=12,
                    border_color=Colors.GREY_400,
                    focused_border_color=Colors.BLUE_400,
                    on_focus=lambda e, r=row, c=col: self.on_cell_focus(e, r, c),
                    on_change=lambda e, r=row, c=col: self.on_cell_change(e, r, c),
                    dense=True,
                )
                
                row_cells.append(cell)
            
            rows.append(ft.Row(row_cells, spacing=0))
        
        return ft.Column(rows, spacing=0, scroll=ft.ScrollMode.ALWAYS)
    
    def get_column_letter(self, col_num):
        """Convierte número de columna a letra (0=A, 1=B, ..., 26=AA)"""
        letter = ""
        while col_num >= 0:
            letter = chr(col_num % 26 + 65) + letter
            col_num = col_num // 26 - 1
        return letter
    
    def on_cell_focus(self, e, row, col):
        """Cuando una celda recibe el foco"""
        self.selected_cell = (row, col)
        cell_id = f"{row}_{col}"
        cell_value = self.data.get(cell_id, "")
        col_letter = self.get_column_letter(col)
        
        self.formula_bar.label = f"{col_letter}{row + 1}"
        self.formula_bar.value = cell_value
        self.formula_bar.update()
    
    def on_cell_change(self, e, row, col):
        """Cuando cambia el valor de una celda"""
        cell_id = f"{row}_{col}"
        self.data[cell_id] = e.control.value
        
        # Actualizar la barra de fórmulas si es la celda seleccionada
        if self.selected_cell == (row, col):
            self.formula_bar.value = e.control.value
            self.formula_bar.update()
    
    def update_cell_value(self, e):
        """Actualiza el valor de la celda desde la barra de fórmulas"""
        if self.selected_cell:
            row, col = self.selected_cell
            cell_id = f"{row}_{col}"
            self.data[cell_id] = self.formula_bar.value
            
            # Actualizar la interfaz
            self.rebuild_table()
    
    def add_row(self, e):
        """Agrega una nueva fila"""
        self.num_rows += 1
        self.rebuild_table()
    
    def add_column(self, e):
        """Agrega una nueva columna"""
        self.num_cols += 1
        self.rebuild_table()
    
    def clear_all(self, e):
        """Limpia todos los datos"""
        self.data.clear()
        self.formula_bar.value = ""
        self.formula_bar.label = ""
        self.rebuild_table()
    
    def save_to_csv(self, e):
        """Guarda los datos en un archivo CSV"""
        filename = self.file_name.value or "datos.csv"
        if not filename.endswith('.csv'):
            filename += '.csv'
        
        try:
            # Preparar datos para CSV
            csv_data = []
            for row in range(self.num_rows):
                row_data = []
                for col in range(self.num_cols):
                    cell_id = f"{row}_{col}"
                    cell_value = self.data.get(cell_id, "")
                    row_data.append(cell_value)
                csv_data.append(row_data)
            
            # Escribir CSV
            with open(filename, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                # Escribir encabezados
                headers = [self.get_column_letter(col) for col in range(self.num_cols)]
                writer.writerow(headers)
                # Escribir datos
                writer.writerows(csv_data)
            
            self.show_message(f"✓ Archivo guardado: {filename}", Colors.GREEN_700)
        except Exception as ex:
            self.show_message(f"✗ Error al guardar: {str(ex)}", Colors.RED_700)
    
    def load_from_csv(self, e):
        """Carga datos desde un archivo CSV"""
        filename = self.file_name.value or "datos.csv"
        if not filename.endswith('.csv'):
            filename += '.csv'
        
        try:
            if not os.path.exists(filename):
                self.show_message(f"✗ Archivo no encontrado: {filename}", Colors.RED_700)
                return
            
            # Leer CSV
            with open(filename, 'r', encoding='utf-8') as f:
                reader = csv.reader(f)
                rows = list(reader)
            
            # Limpiar datos actuales
            self.data.clear()
            
            # Cargar datos (omitir encabezados)
            if len(rows) > 1:
                data_rows = rows[1:]
                self.num_rows = len(data_rows)
                self.num_cols = len(data_rows[0]) if data_rows else 0
                
                for row_idx, row_data in enumerate(data_rows):
                    for col_idx, cell_value in enumerate(row_data):
                        cell_id = f"{row_idx}_{col_idx}"
                        self.data[cell_id] = cell_value
            
            self.rebuild_table()
            self.show_message(f"✓ Archivo cargado: {filename}", Colors.GREEN_700)
        except Exception as ex:
            self.show_message(f"✗ Error al cargar: {str(ex)}", Colors.RED_700)
    
    def show_message(self, message, color):
        """Muestra un mensaje temporal"""
        snack = ft.SnackBar(
            content=ft.Text(message, color=Colors.WHITE),
            bgcolor=color,
            duration=3000,
        )
        self.page.overlay.append(snack)
        snack.open = True
        self.page.update()
    
    def rebuild_table(self):
        """Reconstruye toda la tabla"""
        self.page.clean()
        self.build_ui()


def main(page: ft.Page):
    app = ExcelAppMejorado(page)


if __name__ == "__main__":
    ft.run(main)
