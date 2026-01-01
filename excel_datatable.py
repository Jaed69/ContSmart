import flet as ft
from flet import Colors, Border
import csv
import os


class ExcelDataTable:
    def __init__(self, page: ft.Page):
        self.page = page
        self.page.title = "Excel DataTable - ContSmart"
        self.page.padding = 20
        
        # Configuración inicial
        self.num_rows = 20
        self.num_cols = 10
        
        # Almacén de datos
        self.data = {}
        
        # Cell editing
        self.edit_dialog = None
        self.editing_cell = None
        
        # Nombre de archivo
        self.file_name = ft.TextField(
            hint_text="nombre_archivo.csv",
            width=200,
            text_size=12,
            value="datos.csv",
        )
        
        self.build_ui()
    
    def get_column_letter(self, col_num):
        """Convierte número de columna a letra"""
        letter = ""
        while col_num >= 0:
            letter = chr(col_num % 26 + 65) + letter
            col_num = col_num // 26 - 1
        return letter
    
    def create_datatable(self):
        """Crea el DataTable"""
        # Crear columnas
        columns = []
        columns.append(ft.DataColumn(ft.Text("#", weight=ft.FontWeight.BOLD)))
        
        for col in range(self.num_cols):
            col_letter = self.get_column_letter(col)
            columns.append(
                ft.DataColumn(
                    ft.Text(col_letter, weight=ft.FontWeight.BOLD),
                    numeric=False,
                )
            )
        
        # Crear filas
        rows = []
        for row_idx in range(self.num_rows):
            cells = []
            
            # Número de fila
            cells.append(
                ft.DataCell(
                    ft.Text(str(row_idx + 1), weight=ft.FontWeight.BOLD)
                )
            )
            
            # Celdas de datos
            for col_idx in range(self.num_cols):
                cell_id = f"{row_idx}_{col_idx}"
                cell_value = self.data.get(cell_id, "")
                
                cells.append(
                    ft.DataCell(
                        ft.Text(cell_value if cell_value else ""),
                        on_tap=lambda e, r=row_idx, c=col_idx: self.edit_cell(r, c),
                    )
                )
            
            rows.append(ft.DataRow(cells=cells))
        
        # Crear DataTable
        datatable = ft.DataTable(
            columns=columns,
            rows=rows,
            border=Border.all(1, Colors.GREY_400),
            border_radius=10,
            vertical_lines=ft.BorderSide(1, Colors.GREY_300),
            horizontal_lines=ft.BorderSide(1, Colors.GREY_300),
            heading_row_color=Colors.BLUE_GREY_100,
            heading_row_height=40,
            data_row_min_height=35,
            data_row_max_height=35,
            column_spacing=10,
        )
        
        return datatable
    
    def edit_cell(self, row, col):
        """Abre un diálogo para editar la celda"""
        cell_id = f"{row}_{col}"
        current_value = self.data.get(cell_id, "")
        col_letter = self.get_column_letter(col)
        
        # Campo de texto para editar
        edit_field = ft.TextField(
            value=current_value,
            label=f"Celda {col_letter}{row + 1}",
            autofocus=True,
            on_submit=lambda e: self.save_cell(row, col, edit_field.value),
        )
        
        # Crear diálogo
        self.edit_dialog = ft.AlertDialog(
            title=ft.Text(f"Editar {col_letter}{row + 1}"),
            content=edit_field,
            actions=[
                ft.TextButton("Cancelar", on_click=self.close_dialog),
                ft.FilledButton(
                    "Guardar",
                    on_click=lambda e: self.save_cell(row, col, edit_field.value)
                ),
            ],
            actions_alignment=ft.MainAxisAlignment.END,
        )
        
        self.page.overlay.append(self.edit_dialog)
        self.edit_dialog.open = True
        self.page.update()
    
    def save_cell(self, row, col, value):
        """Guarda el valor de la celda"""
        cell_id = f"{row}_{col}"
        self.data[cell_id] = value
        self.close_dialog()
        self.rebuild_table()
    
    def close_dialog(self, e=None):
        """Cierra el diálogo de edición"""
        if self.edit_dialog:
            self.edit_dialog.open = False
            self.page.update()
    
    def add_row(self, e):
        """Agrega una nueva fila"""
        self.num_rows += 1
        self.rebuild_table()
        self.show_message(f"✓ Fila {self.num_rows} agregada", Colors.GREEN_700)
    
    def add_column(self, e):
        """Agrega una nueva columna"""
        self.num_cols += 1
        col_letter = self.get_column_letter(self.num_cols - 1)
        self.rebuild_table()
        self.show_message(f"✓ Columna {col_letter} agregada", Colors.GREEN_700)
    
    def clear_all(self, e):
        """Limpia todos los datos"""
        self.data.clear()
        self.rebuild_table()
        self.show_message("✓ Todos los datos han sido limpiados", Colors.ORANGE_700)
    
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
            duration=2000,
        )
        self.page.overlay.append(snack)
        snack.open = True
        self.page.update()
    
    def rebuild_table(self):
        """Reconstruye toda la tabla"""
        # Limpiar solo el contenido principal, no los overlays
        self.page.controls.clear()
        self.build_ui()
    
    def build_ui(self):
        """Construye la interfaz de usuario"""
        # Título
        title = ft.Text(
            "📊 Hoja de Cálculo - ContSmart",
            size=28,
            weight=ft.FontWeight.BOLD,
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
        
        # Información
        info = ft.Text(
            f"📝 Haz clic en cualquier celda para editarla | Filas: {self.num_rows} | Columnas: {self.num_cols}",
            size=12,
            italic=True,
            color=Colors.GREY_700,
        )
        
        # Crear tabla
        datatable = self.create_datatable()
        
        # Contenedor con scroll
        table_container = ft.Container(
            content=ft.Column([datatable], scroll=ft.ScrollMode.ALWAYS),
            border=Border.all(2, Colors.BLUE_GREY_200),
            border_radius=10,
            padding=10,
            expand=True,
        )
        
        # Layout principal
        self.page.add(
            ft.Column([
                title,
                ft.Divider(height=1, color=Colors.GREY_300),
                actions,
                file_actions,
                info,
                table_container,
            ], expand=True, spacing=15)
        )
        
        self.page.update()


def main(page: ft.Page):
    page.theme_mode = ft.ThemeMode.LIGHT
    app = ExcelDataTable(page)


if __name__ == "__main__":
    ft.run(main)
