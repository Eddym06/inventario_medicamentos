import pandas as pd
from datetime import datetime
import json
import argparse
import os
import unicodedata
import sys
from threading import Thread
from PySide6.QtWidgets import *
from PySide6.QtCore import *
from PySide6.QtGui import *
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure

# Constantes
DEFAULT_FILEPATH = 'Inventario.xlsx'
ALIAS_FILE = 'alias.json'
REQUIRED_INVENTARIO_COLS = ['Medicamento', 'Componente', 'Cantidad', 'Precio Unitario']
VENTAS_COLS = ['Fecha', 'Medicamento', 'Cantidad', 'Precio Total']

# Diccionario de alias por defecto
DEFAULT_ALIAS_DICT = {
    "dolocordal": "acetaminofén",
    "aspirina": "ácido acetilsalicílico"
}

class DataManager:
    def __init__(self, filepath):
        self.filepath = filepath
        self.df_inventario, self.df_ventas = self.cargar_excel()
        self.stats_cache = None
        self.last_modified = self.get_file_modified_time()

    def get_file_modified_time(self):
        try:
            import os
            return os.path.getmtime(self.filepath)
        except:
            return 0

    def cargar_excel(self):
        try:
            with pd.ExcelFile(self.filepath) as xls:
                df_inventario = pd.read_excel(xls, sheet_name='Inventario')
                if not all(col in df_inventario.columns for col in REQUIRED_INVENTARIO_COLS):
                    raise ValueError(f"Columnas requeridas faltantes: {REQUIRED_INVENTARIO_COLS}")
                if not all(df_inventario['Cantidad'].apply(lambda x: isinstance(x, (int, float)) and x >= 0)):
                    raise ValueError("Cantidad debe ser un número no negativo")
                if not all(df_inventario['Precio Unitario'].apply(lambda x: isinstance(x, (int, float)) and x >= 0)):
                    raise ValueError("Precio Unitario debe ser un número no negativo")
                df_inventario.set_index('Medicamento', inplace=True)
                df_ventas = pd.read_excel(xls, sheet_name='Ventas', dtype={'Fecha': str}) if 'Ventas' in xls.sheet_names else pd.DataFrame(columns=VENTAS_COLS)
            return df_inventario, df_ventas
        except pd.errors.EmptyDataError:
            raise ValueError("El archivo Excel está vacío")
        except pd.errors.ParserError:
            raise ValueError("Formato de Excel inválido")
        except FileNotFoundError:
            raise FileNotFoundError(f"Archivo '{self.filepath}' no encontrado")
        except Exception as e:
            raise Exception(f"Error al cargar Excel: {str(e)}")

    def reload_if_changed(self):
        """Recarga el Excel solo si el archivo ha sido modificado"""
        current_modified = self.get_file_modified_time()
        if current_modified > self.last_modified:
            try:
                self.df_inventario, self.df_ventas = self.cargar_excel()
                self.stats_cache = None  # Forzar recálculo de estadísticas
                self.last_modified = current_modified
                return True  # Indica que se recargó
            except Exception as e:
                print(f"Error al recargar Excel: {str(e)}")
                return False
        return False  # No se recargó porque no cambió

    def save(self, df_inventario, df_ventas):
        self.stats_cache = generar_estadisticas(df_inventario, df_ventas)
        Thread(target=guardar_excel, args=(self.filepath, df_inventario, df_ventas, self.stats_cache), daemon=True).start()
        # Actualizar el tiempo de modificación después de guardar
        self.last_modified = self.get_file_modified_time()

    def get_stats(self):
        if self.stats_cache is None:
            self.stats_cache = generar_estadisticas(self.df_inventario, self.df_ventas)
        return self.stats_cache

def remove_accents(text: str) -> str:
    return ''.join(c for c in unicodedata.normalize('NFD', text) if unicodedata.category(c) != 'Mn')

def cargar_alias(filepath: str = ALIAS_FILE) -> dict:
    if os.path.exists(filepath):
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                return json.load(f)
        except json.JSONDecodeError:
            print(f"Error al cargar {filepath}. Usando alias por defecto.")
    return DEFAULT_ALIAS_DICT

def guardar_excel(filepath: str, df_inventario: pd.DataFrame, df_ventas: pd.DataFrame, stats: dict):
    try:
        with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
            df_inventario.reset_index().to_excel(writer, sheet_name='Inventario', index=False)
            df_ventas.to_excel(writer, sheet_name='Ventas', index=False)
            stats_data = [
                {'Métrica': 'Total vendido', 'Valor': stats['Total Vendidos']},
                {'Métrica': 'Ganancia total', 'Valor': f"${stats['Ganancia Total']:.2f}"}
            ]
            for i, (med, cant) in enumerate(stats['Top 10'].items(), 1):
                stats_data.append({'Métrica': f'Top {i}', 'Valor': f'{remove_accents(med)}: {cant} unidades'})
            for item in stats['Stock Bajo']:
                stats_data.append({'Métrica': 'Stock Bajo', 'Valor': f"{remove_accents(item['Medicamento'])}: {item['Cantidad']} unidades"})
            stats_df = pd.DataFrame(stats_data)
            stats_df.to_excel(writer, sheet_name='Estadisticas', index=False)
        print("Datos guardados exitosamente.")
    except Exception as e:
        print(f"Error al guardar Excel: {str(e)}")

def buscar_medicamento(df: pd.DataFrame, termino: str, alias_dict: dict) -> pd.DataFrame | None:
    if not termino.strip():
        return None
    termino = alias_dict.get(termino.lower().strip(), termino.lower().strip())
    mask = (
        df['Medicamento'].str.lower().str.contains(termino, na=False) |
        df['Componente'].str.lower().str.contains(termino, na=False)
    )
    resultados = df[mask]
    return resultados if not resultados.empty else None

def realizar_venta(df_inventario: pd.DataFrame, df_ventas: pd.DataFrame, medicamento: str, cantidad: int, data_manager: 'DataManager') -> tuple[bool, pd.DataFrame | None]:
    if medicamento not in df_inventario.index:
        print("Medicamento no encontrado en el inventario.")
        return False, None
    if df_inventario.loc[medicamento, 'Cantidad'] < cantidad:
        print(f"Alerta! No hay suficiente stock de {remove_accents(medicamento)}. Stock actual: {df_inventario.loc[medicamento, 'Cantidad']}")
        return False, None
    df_inventario.loc[medicamento, 'Cantidad'] -= cantidad
    precio_unitario = df_inventario.loc[medicamento, 'Precio Unitario']
    nueva_venta = pd.DataFrame({
        'Fecha': [datetime.now().strftime('%Y-%m-%d %H:%M:%S')],
        'Medicamento': [medicamento],
        'Cantidad': [cantidad],
        'Precio Total': [cantidad * precio_unitario]
    })
    df_ventas = pd.concat([df_ventas, nueva_venta], ignore_index=True)
    data_manager.save(df_inventario, df_ventas)
    print(f"Venta registrada: {cantidad} unidades de {remove_accents(medicamento)}")
    return True, df_ventas

def generar_estadisticas(df_inventario: pd.DataFrame, df_ventas: pd.DataFrame) -> dict:
    stats = {}
    if df_ventas.empty:
        stats['Total Vendidos'] = 0
        stats['Ganancia Total'] = 0.0
        stats['Top 10'] = {}
    else:
        # Asegurar que las columnas numéricas sean del tipo correcto
        df_ventas_copy = df_ventas.copy()
        df_ventas_copy['Cantidad'] = pd.to_numeric(df_ventas_copy['Cantidad'], errors='coerce').fillna(0)
        df_ventas_copy['Precio Total'] = pd.to_numeric(df_ventas_copy['Precio Total'], errors='coerce').fillna(0)

        stats['Total Vendidos'] = int(df_ventas_copy['Cantidad'].sum())
        stats['Ganancia Total'] = float(df_ventas_copy['Precio Total'].sum())
        top10 = df_ventas_copy.groupby('Medicamento')['Cantidad'].sum().nlargest(10)
        stats['Top 10'] = top10.to_dict()
    stock_bajo = df_inventario[df_inventario['Cantidad'] < 5]
    stats['Stock Bajo'] = stock_bajo.reset_index()[['Medicamento', 'Cantidad']].to_dict('records')
    return stats

def run_cli():
    parser = argparse.ArgumentParser(description="Sistema de Inventario Farmacéutico")
    parser.add_argument('--search', type=str, help="Buscar medicamento por nombre o componente")
    parser.add_argument('--sell', type=str, help="Vender medicamento (formato: nombre,cantidad)")
    parser.add_argument('--stats', action='store_true', help="Mostrar estadísticas")
    args = parser.parse_args()

    alias_dict = cargar_alias()
    data_manager = DataManager(DEFAULT_FILEPATH)

    if args.search:
        resultados = buscar_medicamento(data_manager.df_inventario.reset_index(), args.search, alias_dict)
        if resultados is not None and not resultados.empty:
            print(resultados.to_string())
        else:
            print("No se encontraron resultados.")
    elif args.sell:
        try:
            medicamento, cantidad = args.sell.split(',')
            cantidad = int(cantidad)
            success, df_ventas = realizar_venta(data_manager.df_inventario, data_manager.df_ventas, medicamento, cantidad, data_manager)
            if success:
                data_manager.df_ventas = df_ventas
                print(f"Venta exitosa: {cantidad} unidades de {medicamento}")
        except ValueError:
            print("Formato inválido. Use: --sell medicamento,cantidad")
    elif args.stats:
        stats = data_manager.get_stats()
        print(f"Total Vendidos: {stats['Total Vendidos']}")
        print(f"Ganancia Total: ${stats['Ganancia Total']:.2f}")
        print("Top 10:", stats['Top 10'])
        print("Stock Bajo:", stats['Stock Bajo'])

class InventarioTab(QWidget):
    def __init__(self, parent):
        super().__init__()
        self.parent = parent
        self.search_timer = QTimer()
        self.search_timer.setSingleShot(True)
        self.search_timer.timeout.connect(self.buscar)
        self.setup_ui()
        self.setup_animations()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(25, 25, 25, 25)
        layout.setSpacing(20)

        title = QLabel("📋 Inventario de Medicamentos")
        title.setObjectName("titleLabel")
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        search_card = QFrame()
        search_card.setObjectName("searchCard")
        search_layout = QVBoxLayout(search_card)
        search_layout.setContentsMargins(20, 15, 20, 15)

        search_label = QLabel("🔍 Buscar Medicamento")
        search_label.setObjectName("sectionLabel")
        search_layout.addWidget(search_label)

        input_layout = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setObjectName("modernInput")
        self.search_input.setPlaceholderText("Escribe el nombre del medicamento o componente...")
        self.search_input.setToolTip("Escribe el nombre o componente (mínimo 2 caracteres)")

        self.completer = QCompleter(self.parent.data_manager.df_inventario.index.tolist())
        self.completer.setCaseSensitivity(Qt.CaseInsensitive)
        self.completer.setFilterMode(Qt.MatchContains)
        self.completer.activated.connect(self.on_completer_activated)
        self.search_input.setCompleter(self.completer)

        search_button = QPushButton("Buscar")
        search_button.setObjectName("primaryButton")
        search_button.setToolTip("Iniciar búsqueda (Ctrl+B)")
        search_button.clicked.connect(self.buscar)

        clear_button = QPushButton("Limpiar")
        clear_button.setObjectName("secondaryButton")
        clear_button.setToolTip("Limpiar búsqueda (Ctrl+L)")
        clear_button.clicked.connect(self.clear_search)

        search_button.setFixedWidth(140)
        clear_button.setFixedWidth(140)

        input_layout.addWidget(self.search_input)
        input_layout.addWidget(search_button)
        input_layout.addWidget(clear_button)
        search_layout.addLayout(input_layout)
        
        # Botones para manejar plantillas
        template_buttons_layout = QHBoxLayout()
        template_buttons_layout.setSpacing(15)

        crear_plantilla_button = QPushButton("📄 Crear Plantilla")
        crear_plantilla_button.setObjectName("primaryButton")
        crear_plantilla_button.setToolTip("Crear una nueva plantilla Excel y reiniciar contadores")
        crear_plantilla_button.clicked.connect(self.crear_plantilla)
        crear_plantilla_button.setFixedWidth(180)

        abrir_plantilla_button = QPushButton("📂 Abrir Plantilla")
        abrir_plantilla_button.setObjectName("secondaryButton")
        abrir_plantilla_button.setToolTip("Abrir la plantilla Excel actual")
        abrir_plantilla_button.clicked.connect(self.abrir_plantilla)
        abrir_plantilla_button.setFixedWidth(180)

        refresh_button = QPushButton("🔄 Refrescar Datos")
        refresh_button.setObjectName("secondaryButton")
        refresh_button.setToolTip("Recargar datos desde Excel si han cambiado externamente")
        refresh_button.clicked.connect(self.refresh_data)
        refresh_button.setFixedWidth(180)

        template_buttons_layout.addWidget(crear_plantilla_button)
        template_buttons_layout.addWidget(abrir_plantilla_button)
        template_buttons_layout.addWidget(refresh_button)
        template_buttons_layout.addStretch()

        search_layout.addLayout(template_buttons_layout)
        layout.addWidget(search_card)

        results_card = QFrame()
        results_card.setObjectName("resultsCard")
        results_card.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        results_layout = QVBoxLayout(results_card)
        results_layout.setContentsMargins(20, 15, 20, 15)

        results_label = QLabel("📊 Resultados de Búsqueda")
        results_label.setObjectName("sectionLabel")
        results_layout.addWidget(results_label)

        self.table = QTableWidget()
        self.table.setObjectName("modernTable")
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["💊 Medicamento", "🧪 Componente", "📦 Cantidad", "💰 Precio"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setColumnWidth(0, 250)
        self.table.setColumnWidth(1, 200)
        self.table.setColumnWidth(2, 150)
        self.table.setColumnWidth(3, 80)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.verticalHeader().setVisible(False)
        self.table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.table.setSortingEnabled(True)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        
        # Hacer las filas más grandes
        self.table.verticalHeader().setDefaultSectionSize(45)
        self.table.verticalHeader().setMinimumSectionSize(40)
        
        results_layout.addWidget(self.table)
        layout.addWidget(results_card)
        layout.addStretch(1)

        QShortcut(QKeySequence("Ctrl+B"), self, self.buscar)
        QShortcut(QKeySequence("Ctrl+L"), self, self.clear_search)

        self.search_input.textChanged.connect(self.buscar_reactivo)

    def on_completer_activated(self, text):
        """Se ejecuta cuando el usuario selecciona un medicamento del autocompletado"""
        if text and text in self.parent.data_manager.df_inventario.index:
            self.selected_medicamento = text
            self.parent.show_inventario_clickable_toast("Ir a Ventas con este medicamento", "ventas")

    def setup_animations(self):
        self.fade_effect = QGraphicsOpacityEffect(self)
        self.setGraphicsEffect(self.fade_effect)
        self.fade_animation = QPropertyAnimation(self.fade_effect, b"opacity")
        self.fade_animation.setDuration(150)
        self.fade_animation.setStartValue(0.0)
        self.fade_animation.setEndValue(1.0)
        self.fade_animation.setEasingCurve(QEasingCurve.InOutQuad)

    def showEvent(self, event):
        super().showEvent(event)
        if hasattr(self, 'fade_animation'):
            self.fade_animation.start()

    def buscar_reactivo(self):
        self.search_timer.stop()
        if len(self.search_input.text()) >= 2:
            self.search_timer.start(300)
        elif len(self.search_input.text()) == 0:
            self.clear_search()

    def buscar(self):
        termino = self.search_input.text()
        if not termino.strip():
            self.table.setRowCount(0)
            return
        self.parent.show_loading_with_message("Buscando medicamentos...")
        resultados = buscar_medicamento(self.parent.data_manager.df_inventario.reset_index(), termino, self.parent.alias_dict)
        if resultados is not None and not resultados.empty:
            resultados_display = resultados.copy()
            resultados_display['Medicamento'] = resultados_display['Medicamento'].apply(remove_accents)
            resultados_display['Componente'] = resultados_display['Componente'].apply(remove_accents)
            self.table.setRowCount(len(resultados_display))
            for i, (_, row) in enumerate(resultados_display.iterrows()):
                self.table.setItem(i, 0, QTableWidgetItem(row['Medicamento']))
                self.table.setItem(i, 1, QTableWidgetItem(row['Componente']))
                self.table.setItem(i, 2, QTableWidgetItem(str(row['Cantidad'])))
                precio_item = QTableWidgetItem(f"${row['Precio Unitario']:.2f}")
                precio_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.table.setItem(i, 3, precio_item)
            QApplication.processEvents()
            self.selected_medicamento = resultados_display.iloc[0]['Medicamento']
            self.parent.show_inventario_clickable_toast("Ir a Ventas con este medicamento", "ventas")
        else:
            self.table.setRowCount(0)
            self.parent.show_inventario_toast("No se encontraron resultados", 2000)
        self.parent.hide_loading()

    def clear_search(self):
        self.search_input.clear()
        self.table.setRowCount(0)

    def crear_plantilla(self):
        """Crear una nueva plantilla Excel y reiniciar todos los contadores a cero"""
        try:
            import subprocess
            import sys
            import pandas as pd
            
            df_inventario = pd.DataFrame(columns=['Medicamento', 'Componente', 'Cantidad', 'Precio Unitario'])
            df_ventas = pd.DataFrame(columns=['Fecha', 'Medicamento', 'Cantidad', 'Precio Total'])
            df_estadisticas = pd.DataFrame(columns=['Métrica', 'Valor'])
            
            with pd.ExcelWriter('Inventario.xlsx', engine='openpyxl') as writer:
                df_inventario.to_excel(writer, sheet_name='Inventario', index=False)
                df_ventas.to_excel(writer, sheet_name='Ventas', index=False)
                df_estadisticas.to_excel(writer, sheet_name='Estadisticas', index=False)
            
            QMessageBox.information(self, "✅ Sistema Reiniciado", 
                                  "Plantilla 'Inventario.xlsx' creada correctamente.\n\n"
                                  "✅ Todos los contadores reiniciados a cero\n"
                                  "✅ Historial de ventas borrado\n"
                                  "✅ Estadísticas limpiadas\n\n"
                                  "El sistema ahora está completamente reiniciado.")
            
            self.parent.data_manager = DataManager('Inventario.xlsx')
            self.parent.alias_dict = cargar_alias()
            self.parent.data_manager.stats_cache = None
            
            self.completer = QCompleter(self.parent.data_manager.df_inventario.index.tolist())
            self.completer.setCaseSensitivity(Qt.CaseInsensitive)
            self.completer.setFilterMode(Qt.MatchContains)
            self.search_input.setCompleter(self.completer)
            
            self.clear_search()
            
            if hasattr(self.parent, 'estadisticas_tab'):
                self.parent.estadisticas_tab.update_stats()
                
        except Exception as e:
            QMessageBox.critical(self, "❌ Error", 
                               f"Error al crear plantilla y reiniciar sistema:\n{str(e)}")

    def abrir_plantilla(self):
        """Abrir la plantilla Excel actual"""
        try:
            import subprocess
            import os

            archivo_actual = 'Inventario.xlsx' if os.path.exists('Inventario.xlsx') else 'Inventario_Ejemplo.xlsx'

            if os.path.exists(archivo_actual):
                os.startfile(archivo_actual)
                QMessageBox.information(self, "📂 Excel Abierto",
                                      f"Archivo '{archivo_actual}' abierto en Excel.")
            else:
                QMessageBox.warning(self, "⚠️ Archivo no encontrado",
                                  "No se encontró ninguna plantilla.\n"
                                  "Por favor, crea una plantilla primero.")

        except Exception as e:
            QMessageBox.critical(self, "❌ Error",
                               f"Error al abrir Excel:\n{str(e)}")

    def refresh_data(self):
        """Refrescar datos desde Excel si han cambiado externamente"""
        self.parent.show_loading_with_message("Verificando cambios en Excel...")
        reloaded = self.parent.data_manager.reload_if_changed()
        if reloaded:
            self.update_completer()
            self.clear_search()
            self.parent.venta_tab.update_stock_display()
            self.parent.venta_tab.update_total()
            self.parent.estadisticas_tab.update_stats()
            self.parent.show_inventario_toast("✅ Datos actualizados desde Excel", 3000)
        else:
            self.parent.show_inventario_toast("ℹ️ No hay cambios en el Excel", 2000)
        self.parent.hide_loading()

    def update_completer(self):
        """Actualizar el autocompletado con los medicamentos actuales del Excel"""
        medicamentos_actuales = list(self.parent.data_manager.df_inventario.index)
        self.completer = QCompleter(medicamentos_actuales)
        self.completer.setCaseSensitivity(Qt.CaseInsensitive)
        self.completer.setFilterMode(Qt.MatchContains)
        self.completer.setCompletionMode(QCompleter.PopupCompletion)
        self.search_input.setCompleter(self.completer)
        self.search_input.update()

class VentaTab(QWidget):
    def __init__(self, parent):
        super().__init__()
        self.parent = parent
        self.setup_ui()
        self.setup_animations()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(25, 25, 25, 25)
        layout.setSpacing(20)

        title = QLabel("💰 Registro de Ventas")
        title.setObjectName("titleLabel")
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        form_card = QFrame()
        form_card.setObjectName("formCard")
        card_layout = QVBoxLayout(form_card)
        card_layout.setContentsMargins(25, 20, 25, 20)
        card_layout.setSpacing(15)

        form_title = QLabel("📝 Información de la Venta")
        form_title.setObjectName("sectionLabel")
        card_layout.addWidget(form_title)

        form_layout_inner = QFormLayout()
        form_layout_inner.setSpacing(15)
        form_layout_inner.setLabelAlignment(Qt.AlignRight)

        self.medicamento_input = QLineEdit()
        self.medicamento_input.setObjectName("modernInput")
        self.medicamento_input.setPlaceholderText("Escribe o selecciona el medicamento...")
        self.medicamento_input.setToolTip("Nombre exacto del medicamento (autocompletado)")
        med_completer = QCompleter(self.parent.data_manager.df_inventario.index.tolist())
        med_completer.setCaseSensitivity(Qt.CaseInsensitive)
        med_completer.setFilterMode(Qt.MatchContains)
        self.medicamento_input.setCompleter(med_completer)
        form_layout_inner.addRow(QLabel("Medicamento:", objectName="fieldLabel"), self.medicamento_input)

        self.stock_widget = QFrame()
        self.stock_widget.setObjectName("stockWidget")
        stock_layout = QHBoxLayout(self.stock_widget)
        stock_layout.setContentsMargins(15, 10, 15, 10)
        self.stock_icon = QLabel("📦")
        self.stock_icon.setObjectName("stockIcon")
        self.stock_label = QLabel("Stock disponible: -")
        self.stock_label.setObjectName("stockLabel")
        stock_layout.addWidget(self.stock_icon)
        stock_layout.addWidget(self.stock_label)
        stock_layout.addStretch()
        form_layout_inner.addRow(QLabel("Disponibilidad:", objectName="fieldLabel"), self.stock_widget)

        self.cantidad_spin = QSpinBox()
        self.cantidad_spin.setObjectName("modernSpinBox")
        self.cantidad_spin.setMinimum(1)
        self.cantidad_spin.setMaximum(1000)
        self.cantidad_spin.setValue(1)
        self.cantidad_spin.setAlignment(Qt.AlignRight)
        self.cantidad_spin.setToolTip("Cantidad a vender (máximo según stock)")
        form_layout_inner.addRow(QLabel("Cantidad a vender:", objectName="fieldLabel"), self.cantidad_spin)

        self.total_widget = QFrame()
        self.total_widget.setObjectName("totalWidget")
        total_layout = QHBoxLayout(self.total_widget)
        total_layout.setContentsMargins(15, 10, 15, 10)
        self.total_label = QLabel("💵 Total: $0.00")
        self.total_label.setObjectName("totalLabel")
        total_layout.addWidget(self.total_label)
        total_layout.addStretch()
        form_layout_inner.addRow(QLabel("Total:", objectName="fieldLabel"), self.total_widget)

        card_layout.addLayout(form_layout_inner)
        layout.addWidget(form_card)
        layout.addStretch(1)

        button_layout = QHBoxLayout()
        button_layout.setSpacing(15)
        self.sell_button = QPushButton("💰 Realizar Venta")
        self.sell_button.setObjectName("primaryButton")
        self.sell_button.setToolTip("Registrar venta (Ctrl+S)")
        self.sell_button.setEnabled(False)
        clear_button = QPushButton("🗑️ Limpiar")
        clear_button.setObjectName("secondaryButton")
        clear_button.setToolTip("Limpiar formulario")
        button_layout.addStretch()
        button_layout.addWidget(clear_button)
        button_layout.addWidget(self.sell_button)
        layout.addLayout(button_layout)

        QShortcut(QKeySequence("Ctrl+S"), self, self.vender)

        self.medicamento_input.textChanged.connect(self.update_stock_display)
        self.cantidad_spin.valueChanged.connect(self.update_total)
        self.sell_button.clicked.connect(self.vender)
        clear_button.clicked.connect(self.clear_form)

    def setup_animations(self):
        self.fade_effect = QGraphicsOpacityEffect(self)
        self.setGraphicsEffect(self.fade_effect)
        self.fade_animation = QPropertyAnimation(self.fade_effect, b"opacity")
        self.fade_animation.setDuration(150)
        self.fade_animation.setStartValue(0.0)
        self.fade_animation.setEndValue(1.0)
        self.fade_animation.setEasingCurve(QEasingCurve.InOutQuad)

    def showEvent(self, event):
        super().showEvent(event)
        if hasattr(self, 'fade_animation'):
            self.fade_animation.start()

    def clear_form(self):
        self.medicamento_input.clear()
        self.cantidad_spin.setValue(1)
        self.update_stock_display()

    def update_stock_display(self):
        medicamento = self.medicamento_input.text()
        if medicamento in self.parent.data_manager.df_inventario.index:
            stock = self.parent.data_manager.df_inventario.loc[medicamento, 'Cantidad']
            self.stock_label.setText(f"Stock disponible: {stock} unidades")
            self.stock_widget.setObjectName("stockWidgetValid")
            self.stock_icon.setText("✅")
            self.sell_button.setEnabled(True)
            self.cantidad_spin.setMaximum(min(stock, 1000))
            self.parent.show_ventas_toast("💊 Medicamento encontrado")
        else:
            self.stock_label.setText("Stock disponible: -")
            self.stock_widget.setObjectName("stockWidget")
            self.stock_icon.setText("📦")
            self.sell_button.setEnabled(False)
            self.cantidad_spin.setMaximum(1000)
        self.stock_widget.style().unpolish(self.stock_widget)
        self.stock_widget.style().polish(self.stock_widget)
        self.update_total()

    def update_total(self):
        medicamento = self.medicamento_input.text()
        if medicamento in self.parent.data_manager.df_inventario.index:
            precio_unitario = self.parent.data_manager.df_inventario.loc[medicamento, 'Precio Unitario']
            cantidad = self.cantidad_spin.value()
            total = precio_unitario * cantidad
            self.total_label.setText(f"💵 Total: ${total:.2f}")
        else:
            self.total_label.setText("💵 Total: $0.00")

    def vender(self):
        medicamento = self.medicamento_input.text()
        cantidad = self.cantidad_spin.value()
        self.parent.show_loading_with_message("Registrando venta...")
        success, updated_df_ventas = realizar_venta(
            self.parent.data_manager.df_inventario, self.parent.data_manager.df_ventas, medicamento, cantidad, self.parent.data_manager
        )
        if success:
            self.parent.data_manager.df_ventas = updated_df_ventas
            self.parent.show_ventas_toast("✅ Venta completada")
            self.parent.estadisticas_tab.update_stats()
            QTimer.singleShot(3000, lambda: self.parent.show_ventas_clickable_toast("🏠 Volver a inventario", "inventario"))
        else:
            self.parent.show_ventas_toast("⚠️ No se pudo realizar la venta", 3000)

        self.clear_form()
        self.parent.hide_loading()
        self.update_stock_display()
        self.update_total()

    def update_completer(self):
        """Actualizar el autocompletado con los medicamentos actuales del Excel"""
        medicamentos_actuales = list(self.parent.data_manager.df_inventario.index)
        med_completer = QCompleter(medicamentos_actuales)
        med_completer.setCaseSensitivity(Qt.CaseInsensitive)
        med_completer.setFilterMode(Qt.MatchContains)
        med_completer.setCompletionMode(QCompleter.PopupCompletion)
        self.medicamento_input.setCompleter(med_completer)
        self.medicamento_input.update()

class EstadisticasTab(QWidget):
    def __init__(self, parent):
        super().__init__()
        self.parent = parent
        self.setup_ui()
        self.setup_animations()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(25, 25, 25, 25)
        layout.setSpacing(20)

        title = QLabel("📊 Estadísticas y Reportes")
        title.setObjectName("titleLabel")
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        summary_layout = QHBoxLayout()
        summary_layout.setSpacing(20)

        sales_card = QFrame()
        sales_card.setObjectName("summaryCard")
        sales_layout = QHBoxLayout(sales_card)
        sales_layout.setContentsMargins(20, 15, 20, 15)
        sales_icon = QLabel("💰")
        sales_icon.setAlignment(Qt.AlignLeft)
        sales_icon.setObjectName("summaryIcon")
        self.total_label = QLabel("Total vendido: 0 unidades")
        self.total_label.setObjectName("summaryValue")
        self.total_label.setAlignment(Qt.AlignLeft)
        sales_layout.addWidget(sales_icon)
        sales_layout.addWidget(self.total_label)
        sales_layout.addStretch()

        revenue_card = QFrame()
        revenue_card.setObjectName("summaryCard")
        revenue_layout = QHBoxLayout(revenue_card)
        revenue_layout.setContentsMargins(20, 15, 20, 15)
        revenue_icon = QLabel("💵")
        revenue_icon.setAlignment(Qt.AlignLeft)
        revenue_icon.setObjectName("summaryIcon")
        self.ganancia_label = QLabel("Ganancia total: $0.00")
        self.ganancia_label.setObjectName("summaryValue")
        self.ganancia_label.setAlignment(Qt.AlignLeft)
        revenue_layout.addWidget(revenue_icon)
        revenue_layout.addWidget(self.ganancia_label)
        revenue_layout.addStretch()

        summary_layout.addWidget(sales_card)
        summary_layout.addWidget(revenue_card)
        layout.addLayout(summary_layout)

        data_grid = QGridLayout()
        data_grid.setSpacing(20)

        stock_card = QFrame()
        stock_card.setObjectName("alertCard")
        stock_layout = QVBoxLayout(stock_card)
        stock_layout.setContentsMargins(20, 15, 20, 15)
        stock_title = QLabel("⚠️ Alertas de Stock Bajo")
        stock_title.setObjectName("sectionLabel")
        stock_layout.addWidget(stock_title)
        self.stock_table = QTableWidget()
        self.stock_table.setObjectName("alertTable")
        self.stock_table.setColumnCount(2)
        self.stock_table.setHorizontalHeaderLabels(["💊 Medicamento", "📦 Stock Actual"])
        self.stock_table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.stock_table.horizontalHeader().setStretchLastSection(True)
        self.stock_table.setColumnWidth(0, 200)
        self.stock_table.setColumnWidth(1, 100)
        self.stock_table.verticalHeader().setVisible(False)
        self.stock_table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.stock_table.setMinimumHeight(300)
        self.stock_table.setMaximumHeight(500)
        self.stock_table.setSortingEnabled(True)
        self.stock_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.stock_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.stock_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.stock_table.setFocusPolicy(Qt.NoFocus)
        self.stock_table.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.stock_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        stock_layout.addWidget(self.stock_table)
        data_grid.addWidget(stock_card, 0, 1)

        graph_card = QFrame()
        graph_card.setObjectName("dataCard")
        graph_layout = QVBoxLayout(graph_card)
        graph_layout.setContentsMargins(20, 15, 20, 15)
        graph_title = QLabel("📈 Gráfico de Top 10 Medicamentos Más Vendidos")
        graph_title.setObjectName("sectionLabel")
        graph_layout.addWidget(graph_title)

        self.figure = Figure()
        self.canvas = FigureCanvas(self.figure)
        graph_layout.addWidget(self.canvas)
        data_grid.addWidget(graph_card, 0, 0)

        layout.addLayout(data_grid)
        layout.addStretch(1)

        refresh_layout = QHBoxLayout()
        refresh_button = QPushButton("🔄 Actualizar Estadísticas")
        refresh_button.setObjectName("primaryButton")
        refresh_button.setToolTip("Actualizar datos")
        refresh_button.clicked.connect(self.update_stats)
        refresh_layout.addStretch()
        refresh_layout.addWidget(refresh_button)
        layout.addLayout(refresh_layout)

        self.update_stats()

    def setup_animations(self):
        self.fade_effect = QGraphicsOpacityEffect(self)
        self.setGraphicsEffect(self.fade_effect)
        self.fade_animation = QPropertyAnimation(self.fade_effect, b"opacity")
        self.fade_animation.setDuration(150)
        self.fade_animation.setStartValue(0.0)
        self.fade_animation.setEndValue(1.0)
        self.fade_animation.setEasingCurve(QEasingCurve.InOutQuad)

    def showEvent(self, event):
        super().showEvent(event)
        if hasattr(self, 'fade_animation'):
            self.fade_animation.start()

    def update_stats(self):
        self.parent.show_loading()
        stats = self.parent.data_manager.get_stats()
        self.total_label.setText(f"Total vendido: {stats['Total Vendidos']} unidades")
        self.ganancia_label.setText(f"Ganancia total: ${stats['Ganancia Total']:.2f}")

        stock_bajo = stats['Stock Bajo']
        self.stock_table.setRowCount(len(stock_bajo))
        for i, item in enumerate(stock_bajo):
            self.stock_table.setItem(i, 0, QTableWidgetItem(remove_accents(item['Medicamento'])))
            self.stock_table.setItem(i, 1, QTableWidgetItem(str(item['Cantidad'])))
            
        # Actualizar gráfico de barras
        self.figure.clear()
        ax = self.figure.add_subplot(111)
        if stats['Top 10']:
            medicamentos = list(stats['Top 10'].keys())
            cantidades = list(stats['Top 10'].values())
            medicamentos = [m[:15] + '...' if len(m) > 15 else m for m in medicamentos]
            ax.bar(medicamentos, cantidades, color='#4d79ff')
            ax.set_title('Top 10 Medicamentos Más Vendidos')
            ax.set_ylabel('Unidades Vendidas')
            ax.tick_params(axis='x', rotation=45)
            self.figure.tight_layout()
        
        # Forzar actualización completa del gráfico
        self.canvas.draw()
        self.canvas.flush_events()
        QApplication.processEvents()
        
        # Forzar repintado del widget canvas
        self.canvas.update()
        self.canvas.repaint()

        self.parent.hide_loading()
        self.parent.show_estadisticas_toast("✅ Estadísticas actualizadas", 2000)


class HelpDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("🆘 Guía de Ayuda - Sistema de Inventario")
        self.setFixedSize(900, 700)
        self.setup_ui()
        self.apply_styles()
        
    def setup_ui(self):
        layout = QVBoxLayout(self)
        
        # Título
        title = QLabel("🆘 Centro de Ayuda")
        title.setObjectName("helpTitle")
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)
        
        # Pestañas de ayuda
        self.tab_widget = QTabWidget()
        
        # Pestaña de Guía de Uso
        usage_tab = self.create_usage_guide()
        self.tab_widget.addTab(usage_tab, "📖 Guía de Uso")
        
        # Pestaña de Problemas y Soluciones
        problems_tab = self.create_problems_guide()
        self.tab_widget.addTab(problems_tab, "🔧 Solución de Problemas")
        
        layout.addWidget(self.tab_widget)
        
        # Botón de cerrar
        close_button = QPushButton("Cerrar")
        close_button.clicked.connect(self.accept)
        close_button.setObjectName("helpCloseButton")
        layout.addWidget(close_button)
        
    def create_usage_guide(self):
        """Crea la pestaña con la guía de uso"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        # Crear área de scroll
        scroll = QScrollArea()
        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout(scroll_widget)
        
        usage_content = """
        <h2>📋 Pestaña de Inventario</h2>
        <p><b>Agregar medicamentos:</b></p>
        <ul>
            <li>Ingresa el nombre del medicamento en el campo "Medicamento"</li>
            <li>Especifica el componente activo principal</li>
            <li>Introduce la cantidad disponible</li>
            <li>Establece el precio unitario</li>
            <li>Haz clic en "Agregar Medicamento" para guardarlo</li>
        </ul>
        
        <p><b>Buscar medicamentos:</b></p>
        <ul>
            <li>Usa la barra de búsqueda en la parte superior</li>
            <li>El sistema buscará por nombre, componente o cantidad</li>
            <li>Los resultados aparecen automáticamente mientras escribes</li>
        </ul>
        
        <p><b>Editar inventario:</b></p>
        <ul>
            <li>Haz doble clic en cualquier celda de la tabla para editarla</li>
            <li>Los cambios se guardan automáticamente</li>
            <li>Puedes ordenar las columnas haciendo clic en los encabezados</li>
        </ul>
        
        <h2>💰 Pestaña de Ventas</h2>
        <p><b>Registrar una venta:</b></p>
        <ul>
            <li>Selecciona el medicamento de la lista desplegable</li>
            <li>Ajusta la cantidad con los botones + y -</li>
            <li>El precio total se calcula automáticamente</li>
            <li>Haz clic en "Registrar Venta" para confirmar</li>
        </ul>
        
        <p><b>Consultar stock:</b></p>
        <ul>
            <li>El stock disponible se muestra automáticamente</li>
            <li>Si no hay stock suficiente, aparece una alerta</li>
            <li>Los medicamentos sin stock aparecen marcados</li>
        </ul>
        
        <h2>📊 Pestaña de Estadísticas</h2>
        <p><b>Ver estadísticas:</b></p>
        <ul>
            <li>Las estadísticas se actualizan automáticamente</li>
            <li>Incluye medicamentos más vendidos (Top 10)</li>
            <li>Muestra gráficos de ventas y tendencias</li>
            <li>Calcula ingresos totales y por período</li>
        </ul>
        
        <h2>⚙️ Funciones Generales</h2>
        <p><b>Auto-actualización:</b></p>
        <ul>
            <li>Los datos se actualizan automáticamente cada 30 segundos</li>
            <li>Si modificas el Excel externamente, la app lo detectará</li>
        </ul>
        
        <p><b>Notificaciones:</b></p>
        <ul>
            <li>Las notificaciones toast aparecen para confirmar acciones</li>
            <li>Cambian de color según el tipo: verde (éxito), azul (info), rojo (error)</li>
        </ul>
        
        <p><b>Alias de medicamentos:</b></p>
        <ul>
            <li>Puedes usar nombres alternativos para los medicamentos</li>
            <li>Por ejemplo: "dolocordal" se convierte automáticamente en "acetaminofén"</li>
        </ul>
        """
        
        usage_label = QLabel(usage_content)
        usage_label.setWordWrap(True)
        usage_label.setTextFormat(Qt.RichText)
        usage_label.setObjectName("helpContent")
        
        scroll_layout.addWidget(usage_label)
        scroll_layout.addStretch()
        
        scroll.setWidget(scroll_widget)
        scroll.setWidgetResizable(True)
        
        layout.addWidget(scroll)
        return widget
        
    def create_problems_guide(self):
        """Crea la pestaña con la guía de problemas"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        # Crear área de scroll
        scroll = QScrollArea()
        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout(scroll_widget)
        
        problems_content = """
        <h2>🐛 Problemas Comunes de la Aplicación</h2>
        
        <h3>P: La aplicación no inicia o se cierra inesperadamente</h3>
        <p><b>R:</b> Verifica que tienes instaladas todas las dependencias:</p>
        <ul>
            <li><code>pip install pandas PySide6 matplotlib openpyxl</code></li>
            <li>Asegúrate de tener Python 3.8+ instalado</li>
            <li>Verifica que el archivo Inventario.xlsx esté en el mismo directorio</li>
        </ul>
        
        <h3>P: Los datos no se guardan correctamente</h3>
        <p><b>R:</b> Posibles causas y soluciones:</p>
        <ul>
            <li>Verifica que el archivo Excel no esté abierto en otro programa</li>
            <li>Asegúrate de tener permisos de escritura en la carpeta</li>
            <li>Revisa que el archivo no esté marcado como solo lectura</li>
        </ul>
        
        <h3>P: La búsqueda no funciona correctamente</h3>
        <p><b>R:</b> Sigue estos pasos:</p>
        <ul>
            <li>Asegúrate de escribir correctamente el nombre del medicamento</li>
            <li>Prueba con nombres parciales o componentes activos</li>
            <li>Verifica que el medicamento existe en el inventario</li>
        </ul>
        
        <h3>P: Las estadísticas no se actualizan</h3>
        <p><b>R:</b> Intenta lo siguiente:</p>
        <ul>
            <li>Cambia a otra pestaña y regresa a Estadísticas</li>
            <li>Las estadísticas se actualizan automáticamente cada 30 segundos</li>
            <li>Si persiste, reinicia la aplicación</li>
        </ul>
        
        <h3>P: Error al registrar ventas</h3>
        <p><b>R:</b> Revisa estos puntos:</p>
        <ul>
            <li>Verifica que hay stock suficiente del medicamento</li>
            <li>Asegúrate de que la cantidad sea un número válido</li>
            <li>Comprueba que el medicamento existe en el inventario</li>
        </ul>
        
        <h2>📊 Problemas con Excel</h2>
        
        <h3>P: Error al abrir el archivo Excel</h3>
        <p><b>R:</b> Soluciones posibles:</p>
        <ul>
            <li>Verifica que el archivo se llame exactamente "Inventario.xlsx"</li>
            <li>Comprueba que tenga las hojas "Inventario" y "Ventas"</li>
            <li>Asegúrate de que las columnas requeridas existen</li>
        </ul>
        
        <h3>P: Las columnas del Excel están mal organizadas</h3>
        <p><b>R:</b> Estructura requerida:</p>
        <ul>
            <li><b>Hoja "Inventario":</b> Medicamento, Componente, Cantidad, Precio Unitario</li>
            <li><b>Hoja "Ventas":</b> Fecha, Medicamento, Cantidad, Precio Total</li>
            <li>Los nombres deben coincidir exactamente (respeta mayúsculas)</li>
        </ul>
        
        <h3>P: Error de formato en datos numéricos</h3>
        <p><b>R:</b> Verifica que:</p>
        <ul>
            <li>Los precios sean números, no texto</li>
            <li>Las cantidades sean números enteros positivos</li>
            <li>No uses comas como separadores de miles (usa puntos)</li>
            <li>Las celdas estén formateadas como "Número"</li>
        </ul>
        
        <h3>P: Problemas con caracteres especiales</h3>
        <p><b>R:</b> Recomendaciones:</p>
        <ul>
            <li>Guarda el Excel como ".xlsx" (no .xls)</li>
            <li>Evita caracteres especiales raros en nombres de medicamentos</li>
            <li>Usa codificación UTF-8 si editas el archivo programáticamente</li>
        </ul>
        
        <h2>💡 Consejos y Mejores Prácticas</h2>
        
        <h3>Mantén organizado tu inventario:</h3>
        <ul>
            <li>Usa nombres consistentes para medicamentos similares</li>
            <li>Mantén actualizado el stock regularmente</li>
            <li>Revisa las estadísticas para identificar tendencias</li>
        </ul>
        
        <h3>Backup de datos:</h3>
        <ul>
            <li>Haz copias de seguridad regulares del archivo Excel</li>
            <li>Considera tener un sistema de respaldo en la nube</li>
        </ul>
        
        <h3>Rendimiento óptimo:</h3>
        <ul>
            <li>No tengas el Excel abierto mientras usas la aplicación</li>
            <li>Evita modificar el archivo Excel directamente durante operaciones</li>
            <li>Reinicia la aplicación si notas lentitud después de mucho uso</li>
        </ul>
        """
        
        problems_label = QLabel(problems_content)
        problems_label.setWordWrap(True)
        problems_label.setTextFormat(Qt.RichText)
        problems_label.setObjectName("helpContent")
        
        scroll_layout.addWidget(problems_label)
        scroll_layout.addStretch()
        
        scroll.setWidget(scroll_widget)
        scroll.setWidgetResizable(True)
        
        layout.addWidget(scroll)
        return widget
        
    def apply_styles(self):
        """Aplica estilos consistentes con el tema oscuro"""
        self.setStyleSheet("""
            QDialog {
                background-color: #0d1117;
                color: #ffffff;
            }
            
            #helpTitle {
                font-size: 24px;
                font-weight: bold;
                color: #00d4ff;
                margin: 20px 0px;
            }
            
            QTabWidget::pane {
                border-top: 2px solid #3a3f4e;
                background-color: #0d1117;
            }
            
            QTabBar::tab {
                background: #2a2f3e;
                color: #ffffff;
                border: 1px solid #3a3f4e;
                border-bottom: none;
                padding: 12px 20px;
                font-weight: bold;
                font-size: 13px;
                border-top-left-radius: 8px;
                border-top-right-radius: 8px;
            }
            
            QTabBar::tab:selected {
                background: #0d1117;
                color: #00d4ff;
                border: 1px solid #3a3f4e;
                border-bottom: 1px solid #0d1117;
            }
            
            QTabBar::tab:!selected:hover {
                background: #3a3f4e;
            }
            
            #helpContent {
                background-color: #1a1f2e;
                border: 1px solid #3a3f4e;
                border-radius: 8px;
                padding: 20px;
                font-size: 13px;
                line-height: 1.6;
            }
            
            QScrollArea {
                border: none;
                background-color: transparent;
            }
            
            QScrollBar:vertical {
                background-color: #2a2f3e;
                width: 12px;
                border-radius: 6px;
            }
            
            QScrollBar::handle:vertical {
                background-color: #00d4ff;
                border-radius: 6px;
                min-height: 20px;
            }
            
            QScrollBar::handle:vertical:hover {
                background-color: #0099cc;
            }
            
            #helpCloseButton {
                background-color: #00d4ff;
                color: #0d1117;
                border: none;
                padding: 12px 30px;
                border-radius: 6px;
                font-weight: bold;
                font-size: 14px;
                margin: 10px 0px;
            }
            
            #helpCloseButton:hover {
                background-color: #0099cc;
            }
            
            #helpCloseButton:pressed {
                background-color: #006699;
            }
        """)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("💊 Sistema de Inventario de Medicamentos - Versión Pro")
        self.setGeometry(100, 100, 1200, 800)
        self.setMinimumSize(1100, 750)

        # Sistema de gestión de toasts simplificado
        self.active_toast = None
        self.toast_timer = QTimer(self)
        self.toast_timer.setSingleShot(True)
        self.toast_timer.timeout.connect(self.hide_toast)

        self.setup_styles()
        self.alias_dict = cargar_alias()
        self.data_manager = DataManager(DEFAULT_FILEPATH)
        self.setup_ui()
        self.setup_animations()

    def setup_styles(self):
        self.setStyleSheet("""
            /* Estilo general para la ventana */
            QMainWindow {
                background-color: #0d1117;
            }
            
            /* Estilo general para etiquetas */
            QLabel {
                color: #ffffff;
                background-color: transparent !important;
                border: none !important;
                padding: 0px !important;
                margin: 0px !important;
            }
            
            /* Eliminar resaltados en todos los widgets */
            * {
                selection-background-color: #00d4ff;
                selection-color: #0d1117;
            }
            
            /* Estilos específicos para eliminar resaltados */
            QWidget {
                background-color: transparent;
            }
            
            QDialog QLabel {
                background-color: transparent !important;
                color: #ffffff !important;
            }
            
            QMessageBox QLabel {
                background-color: transparent !important;
                color: #ffffff !important;
            }

            /* Estilo para las pestañas */
            QTabWidget::pane {
                border-top: 2px solid #3a3f4e;
                background-color: #0d1117;
            }
            QTabBar::tab {
                background: #2a2f3e;
                color: #ffffff;
                border: 1px solid #3a3f4e;
                border-bottom: none;
                padding: 12px 25px;
                font-weight: bold;
                font-size: 14px;
                border-top-left-radius: 8px;
                border-top-right-radius: 8px;
            }
            QTabBar::tab:selected {
                background: #0d1117;
                color: #00d4ff;
                border: 1px solid #3a3f4e;
                border-bottom: 1px solid #0d1117; /* Conecta con el panel */
            }
            QTabBar::tab:!selected:hover {
                background: #3a3f4e;
            }

            /* Estilos para etiquetas y títulos */
            #titleLabel {
                font-size: 28px;
                font-weight: bold;
                color: #ffffff;
                background-color: transparent;
                padding-bottom: 10px;
            }
            #sectionLabel {
                font-size: 18px;
                font-weight: bold;
                color: #ffffff;
                background-color: transparent;
                border-bottom: 2px solid #3a3f4e;
                padding-bottom: 5px;
            }
            #fieldLabel {
                font-weight: bold;
                color: #ffffff;
                background-color: transparent;
            }

            /* Estilos para tarjetas y marcos */
            QFrame[objectName*="Card"] {
                background-color: #1a1f2e;
                border-radius: 12px;
                border: 1px solid #00d4ff;
            }
            QFrame {
                background-color: #0d1117;
            }
            QFrame[objectName="summaryCard"] {
                background-color: #1a1f2e;
                border-radius: 12px;
                border: 1px solid #404550;
            }
            QFrame[objectName="summaryCard"] QLabel {
                background-color: transparent;
                color: #ffffff;
            }

            /* Estilos para botones */
            QPushButton {
                padding: 12px 20px;
                border-radius: 8px;
                font-weight: bold;
                font-size: 14px;
                border: 1px solid transparent;
            }
            #primaryButton {
                background-color: #3498db; /* Azul moderno */
                color: white;
                border: 1px solid #2980b9;
            }
            #primaryButton:hover {
                background-color: #2980b9;
            }
            #primaryButton:pressed {
                background-color: #2c3e50;
            }
            #secondaryButton {
                background-color: #404550;
                color: #ffffff;
                border: 1px solid #555a67;
            }
            #secondaryButton:hover {
                background-color: #555a67;
                border-color: #6a6f7e;
            }
            #secondaryButton:pressed {
                background-color: #2a2f3e;
            }

            /* Estilos para campos de entrada */
            QLineEdit, QSpinBox {
                padding: 10px;
                border: 1px solid #404550;
                border-radius: 8px;
                font-size: 14px;
                background-color: #2a2f3e;
                color: #ffffff;
            }
            QLineEdit:focus, QSpinBox:focus {
                border-color: #00d4ff;
                background-color: #1a1f2e;
            }

            /* Estilos para tablas */
            QTableWidget {
                border: 1px solid #404550;
                border-radius: 8px;
                gridline-color: #404550;
                font-size: 14px;
                background-color: #1a1f2e;
                color: #ffffff;
            }
            QHeaderView::section {
                background-color: #2a2f3e;
                padding: 12px;
                border: none;
                border-bottom: 1px solid #404550;
                font-weight: bold;
                font-size: 16px;
                color: #ffffff;
            }
            QTableWidget::item {
                padding: 10px;
                border-bottom: 1px solid #404550;
                color: #ffffff;
            }
            QTableWidget::item:selected {
                background-color: #00d4ff;
                color: #0d1117;
            }
            QTableWidget::alternating-row-color {
                background-color: #2a2f3e;
            }
            
            /* Estilos específicos para tabla de resultados */
            #modernTable::item {
                padding: 15px 10px;
                border-bottom: 1px solid #404550;
                color: #ffffff;
                font-size: 14px;
            }

            /* Estilos para widgets de estado */
            #stockWidget {
                background-color: #4a2e30; /* Rojo oscuro para alerta */
                color: #ffffff;
                border-radius: 8px;
            }
            #stockWidgetValid {
                background-color: #2a4d3a; /* Verde oscuro para válido */
                color: #ffffff;
                border-radius: 8px;
            }
            #totalWidget {
                background-color: #404550;
                color: #ffffff;
                border-radius: 8px;
                font-weight: bold;
            }
            #totalLabel {
                font-size: 16px;
                color: #ffffff;
            }

            /* Estilo para la barra de progreso */
            QProgressBar {
                border: none;
                background-color: #404550;
            }
            QProgressBar::chunk {
                background-color: #00d4ff;
            }
            
            /* Estilos específicos para elementos problemáticos */
            #stockLabel {
                background-color: transparent;
                color: #ffffff;
                border: none;
                padding: 0px;
                font-size: 16px;
                font-weight: bold;
            }
            #stockIcon {
                background-color: transparent;
                color: #ffffff;
                border: none;
                padding: 0px;
                margin-right: 8px;
                font-size: 18px;
            }
            #totalLabel {
                background-color: transparent;
                color: #ffffff;
                border: none;
                padding: 0px;
                font-size: 16px;
                font-weight: bold;
            }
            
            /* Estilos para SpinBox moderno */
            #modernSpinBox {
                padding: 12px 15px;
                border: 1px solid #404550;
                border-radius: 8px;
                font-size: 14px;
                background-color: #2a2f3e;
                color: #ffffff;
                min-width: 120px;
                max-width: 200px;
            }
            #modernSpinBox:focus {
                border-color: #00d4ff;
                background-color: #1a1f2e;
            }
            #modernSpinBox::up-button {
                subcontrol-origin: border;
                subcontrol-position: top right;
                width: 24px;
                height: 24px;
                border-left: 1px solid #404550;
                border-bottom: 1px solid #404550;
                border-top-right-radius: 8px;
                background-color: #404550;
            }
            #modernSpinBox::up-button:hover {
                background-color: #555a67;
            }
            #modernSpinBox::up-button {
                subcontrol-origin: border;
                subcontrol-position: top right;
                width: 24px;
                height: 24px;
                border-left: 1px solid #404550;
                border-bottom: 1px solid #404550;
                border-top-right-radius: 8px;
                background-color: #404550;
                color: #ffffff;
                font-size: 16px;
                font-weight: bold;
            }
            #modernSpinBox::up-button:hover {
                background-color: #555a67;
            }
            #modernSpinBox::up-arrow {
                width: 0px;
                height: 0px;
                border-left: 4px solid transparent;
                border-right: 4px solid transparent;
                border-bottom: 6px solid #ffffff;
                margin: 6px;
            }
            #modernSpinBox::down-button {
                subcontrol-origin: border;
                subcontrol-position: bottom right;
                width: 24px;
                height: 24px;
                border-left: 1px solid #404550;
                border-top: 1px solid #404550;
                border-bottom-right-radius: 8px;
                background-color: #404550;
                color: #ffffff;
                font-size: 16px;
                font-weight: bold;
            }
            #modernSpinBox::down-button:hover {
                background-color: #555a67;
            }
            #modernSpinBox::down-arrow {
                width: 0px;
                height: 0px;
                border-left: 4px solid transparent;
                border-right: 4px solid transparent;
                border-top: 6px solid #ffffff;
                margin: 6px;
            }
            
            /* Estilos para elementos de estadísticas */
            #summaryValue {
                font-size: 16px;
                font-weight: bold;
                background-color: transparent;
                color: #ffffff;
                border: none;
                padding: 0px;
            }
            #summaryIcon {
                font-size: 20px;
                margin-right: 10px;
                background-color: transparent;
                color: #ffffff;
                border: none;
                padding: 0px;
            }
            
            /* Botón de ayuda */
            #helpButton {
                background-color: #00d4ff;
                color: #0d1117;
                border: 3px solid #00d4ff;
                border-radius: 25px;
                font-size: 20px;
                font-weight: bold;
                text-align: center;
            }
            #helpButton:hover {
                background-color: #0099cc;
                border-color: #0099cc;
                border-width: 3px;
            }
            #helpButton:pressed {
                background-color: #006699;
                border-color: #006699;
                border-width: 3px;
            }
        """)

    def setup_ui(self):
        central_widget = QWidget()
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        self.loading_bar = QProgressBar(self)
        self.loading_bar.setFixedHeight(4)
        self.loading_bar.setTextVisible(False)
        self.loading_bar.setVisible(False)
        main_layout.addWidget(self.loading_bar)

        self.tabs = QTabWidget()
        self.tabs.setTabPosition(QTabWidget.North)
        self.tabs.setDocumentMode(True)
        self.tabs.setMovable(False)
        main_layout.addWidget(self.tabs)

        self.inventario_tab = InventarioTab(self)
        self.venta_tab = VentaTab(self)
        self.estadisticas_tab = EstadisticasTab(self)

        self.tabs.addTab(self.inventario_tab, "📋 Inventario")
        self.tabs.addTab(self.venta_tab, "💰 Ventas")
        self.tabs.addTab(self.estadisticas_tab, "📊 Estadísticas")

        # Configurar recarga automática periódica cada 30 segundos
        self.auto_refresh_timer = QTimer(self)
        self.auto_refresh_timer.timeout.connect(self.auto_refresh_data)
        self.auto_refresh_timer.start(30000)  # 30 segundos

        self.tabs.currentChanged.connect(self.on_tab_changed)
        
        self.setCentralWidget(central_widget)
        
        # Botón de ayuda en la esquina inferior izquierda (crear después del central widget)
        self.help_button = QPushButton("?", self)
        self.help_button.setFixedSize(50, 50)  # Hacer el botón un poco más grande
        self.help_button.setObjectName("helpButton")
        self.help_button.clicked.connect(self.show_help_dialog)
        self.help_button.setToolTip("🆘 Ayuda y guía de uso")  # Añadir tooltip
        
        # Hacer que el botón esté siempre encima
        self.help_button.raise_()
        
        # Posicionar el botón en la esquina inferior izquierda
        self.position_help_button()

    def setup_animations(self):
        # Configuración simple sin efectos de opacidad para evitar problemas de renderizado
        pass

    def on_tab_changed(self, index):
        """Se dispara cuando el usuario cambia de pestaña para limpiar toasts."""
        # Ocultar cualquier toast activo inmediatamente al cambiar de pestaña
        self.hide_toast()

        # Auto-actualizar estadísticas si se abre esa pestaña (índice 2)
        if index == 2:  # Pestaña de estadísticas
            QTimer.singleShot(200, self.estadisticas_tab.update_stats)

        # Ejecutar animaciones de la nueva pestaña
        current_tab = self.tabs.widget(index)
        if hasattr(current_tab, 'fade_animation'):
            current_tab.fade_animation.start()
            
        # Forzar actualización de renderizado de tablas en la pestaña actual
        self.force_table_refresh(current_tab)

    def show_toast(self, message, duration=3000, is_clickable=False, action=None):
        """Muestra una notificación toast centralizada."""
        # Ocultar el toast anterior si existe
        self.hide_toast()

        self.active_toast = QLabel(message, self)
        self.active_toast.setAlignment(Qt.AlignCenter)
        
        style = """
            background-color: #4CAF50; /* Verde para éxito */
            color: white;
            border-radius: 15px;
            padding: 12px 24px;
            font-weight: bold;
            font-size: 14px;
        """
        if is_clickable:
            style = style.replace("#4CAF50", "#3498db") # Azul para clickeable
            self.active_toast.setCursor(Qt.PointingHandCursor)
            # Usamos un lambda para capturar el valor actual de 'action'
            self.active_toast.mousePressEvent = lambda event, act=action: self.on_toast_clicked(act)

        self.active_toast.setStyleSheet(style)
        self.active_toast.adjustSize()
        
        # Posicionar en la parte inferior central
        x = (self.width() - self.active_toast.width()) / 2
        y = self.height() - self.active_toast.height() - 20
        self.active_toast.move(int(x), int(y))
        
        self.active_toast.show()
        self.toast_timer.start(duration)

    def hide_toast(self):
        """Oculta y destruye el toast activo."""
        self.toast_timer.stop()
        if self.active_toast:
            self.active_toast.deleteLater()
            self.active_toast = None

    def on_toast_clicked(self, action):
        """Maneja el clic en un toast."""
        self.hide_toast()
        if action == "ventas":
            self.tabs.setCurrentIndex(1)
            # Pre-rellenar el campo de medicamento en la pestaña de ventas
            if hasattr(self.inventario_tab, 'selected_medicamento'):
                medicamento = self.inventario_tab.selected_medicamento
                self.venta_tab.medicamento_input.setText(medicamento)
                self.venta_tab.update_stock_display()
        elif action == "inventario":
            self.tabs.setCurrentIndex(0)

    def show_clickable_toast(self, message, action, duration=4000):
        """Muestra un toast que realiza una acción al hacer clic."""
        self.show_toast(message, duration, is_clickable=True, action=action)

    # --- Métodos específicos de ventana (verifican pestaña activa) ---

    def show_inventario_toast(self, message, duration=3000):
        """Muestra toast específico para la ventana de Inventario"""
        if self.tabs.currentIndex() == 0:  # Solo si estamos en inventario
            self.show_toast(message, duration)

    def show_ventas_toast(self, message, duration=3000):
        """Muestra toast específico para la ventana de Ventas"""
        if self.tabs.currentIndex() == 1:  # Solo si estamos en ventas
            self.show_toast(message, duration)

    def show_estadisticas_toast(self, message, duration=3000):
        """Muestra toast específico para la ventana de Estadísticas"""
        if self.tabs.currentIndex() == 2:  # Solo si estamos en estadísticas
            self.show_toast(message, duration)

    def show_inventario_clickable_toast(self, message, action, duration=4000):
        """Muestra toast clickeable específico para la ventana de Inventario"""
        if self.tabs.currentIndex() == 0:  # Solo si estamos en inventario
            self.show_clickable_toast(message, action, duration)

    def show_ventas_clickable_toast(self, message, action, duration=4000):
        """Muestra toast clickeable específico para la ventana de Ventas"""
        if self.tabs.currentIndex() == 1:  # Solo si estamos en ventas
            self.show_clickable_toast(message, action, duration)

    def force_table_refresh(self, tab_widget):
        """Fuerza el renderizado completo de las tablas en una pestaña"""
        if hasattr(tab_widget, 'table'):
            # Para InventarioTab
            table = tab_widget.table
            self.refresh_table_headers(table)
        elif hasattr(tab_widget, 'stock_table'):
            # Para EstadisticasTab
            table = tab_widget.stock_table
            self.refresh_table_headers(table)
            
    def refresh_table_headers(self, table):
        """Simula hover en las cabeceras para activar su renderizado"""
        if table and table.horizontalHeader():
            header = table.horizontalHeader()
            # Forzar repintado de las cabeceras
            header.repaint()
            # Simular eventos de entrada y salida del mouse para activar el renderizado
            for i in range(header.count()):
                # Crear eventos sintéticos para forzar renderizado
                enter_event = QEvent(QEvent.Enter)
                leave_event = QEvent(QEvent.Leave)
                QApplication.postEvent(header, enter_event)
                QApplication.postEvent(header, leave_event)
            QApplication.processEvents()

    def show_loading(self):
        self.loading_bar.setVisible(True)
        self.loading_bar.setRange(0, 0)

    def hide_loading(self):
        self.loading_bar.setVisible(False)
        
    def show_loading_with_message(self, message):
        """Muestra una barra de carga con un mensaje toast asociado"""
        self.show_loading()
        self.show_toast(message, 10000)  # Toast de larga duración

    def auto_refresh_data(self):
        """Refresca los datos automáticamente si el archivo Excel ha cambiado."""
        if self.data_manager.reload_if_changed():
            # Actualizar completers y datos en todas las pestañas
            self.inventario_tab.update_completer()
            self.venta_tab.update_completer()
            self.venta_tab.update_stock_display()
            self.estadisticas_tab.update_stats()
            
            # Mostrar toast en la pestaña actual
            current_tab_index = self.tabs.currentIndex()
            if current_tab_index == 0:
                self.show_inventario_toast("🔄 Datos actualizados automáticamente", 3000)
            elif current_tab_index == 1:
                self.show_ventas_toast("🔄 Datos actualizados automáticamente", 3000)
            elif current_tab_index == 2:
                self.show_estadisticas_toast("🔄 Datos actualizados automáticamente", 3000)

    def show_help_dialog(self):
        """Muestra el diálogo de ayuda con guía de uso y solución de problemas"""
        dialog = HelpDialog(self)
        dialog.exec()
    
    def position_help_button(self):
        """Posiciona el botón de ayuda en la esquina inferior izquierda"""
        if hasattr(self, 'help_button'):
            # Posicionar el botón con margen desde los bordes
            x = 20  # Más margen desde el borde izquierdo
            y = self.height() - 75  # Más margen desde el borde inferior
            self.help_button.move(x, y)
            # Asegurar que el botón esté visible y encima de otros elementos
            self.help_button.show()
            self.help_button.raise_()
    
    def resizeEvent(self, event):
        """Reposicionar el botón de ayuda cuando se redimensiona la ventana"""
        super().resizeEvent(event)
        self.position_help_button()

    def showEvent(self, event):
        """Asegurar que el botón se muestre correctamente cuando se muestra la ventana"""
        super().showEvent(event)
        # Usar QTimer para asegurar que el posicionamiento ocurra después del renderizado
        QTimer.singleShot(100, self.position_help_button)

if __name__ == "__main__":
    if len(sys.argv) > 1:
        run_cli()
    else:
        app = QApplication(sys.argv)
        
        # Optimizaciones para carga rápida y renderizado suave
        # Nota: AA_EnableHighDpiScaling y AA_UseHighDpiPixmaps están obsoletos en Qt6+
        # Qt6 maneja automáticamente el High DPI scaling
        
        # Crear ventana principal directamente sin splash para evitar problemas de renderizado
        window = MainWindow()
        window.show()
        
        sys.exit(app.exec())