"""
Interfaz gráfica PyQt5 para la Simulación de Biblioteca
Reemplaza todas las entradas por teclado del main.py
Incluye visualización de tabla del vector de estado
"""
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QLabel, QLineEdit, QPushButton, 
                             QGroupBox, QGridLayout, QTabWidget, QScrollArea,
                             QMessageBox, QProgressBar, QFileDialog, QDoubleSpinBox,
                             QSpinBox, QCheckBox, QTableWidget, QTableWidgetItem,
                             QHeaderView, QTextEdit)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QIcon, QColor
import sys
import json

from config import ConfigSimulacion
from simulador import Simulador
from visualizador import VisualizadorVectorEstado
from exportador import ExportadorExcel


class SimulacionThread(QThread):
    """Thread para ejecutar la simulación sin bloquear la UI"""
    finished = pyqtSignal(object, object, object)  # vector_estado, metricas, simulador
    error = pyqtSignal(str)
    
    def __init__(self, config):
        super().__init__()
        self.config = config
    
    def run(self):
        try:
            simulador = Simulador(self.config)
            vector_estado = simulador.ejecutar()
            metricas = simulador.calcular_metricas_finales()
            self.finished.emit(vector_estado, metricas, simulador)
        except Exception as e:
            self.error.emit(str(e))


class BibliotecaSimuladorGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.config = ConfigSimulacion()
        self.vector_estado = None
        self.metricas = None
        self.simulador = None
        self.thread = None
        
        self.initUI()
    
    def initUI(self):
        """Inicializa la interfaz de usuario"""
        self.setWindowTitle('Simulación de Sistema de Biblioteca - TP5 SIM')
        self.setGeometry(100, 100, 1400, 900)
        
        # Widget central
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Layout principal
        main_layout = QVBoxLayout()
        central_widget.setLayout(main_layout)
        
        # Título
        title_label = QLabel('🏛️ Simulación de Sistema de Biblioteca')
        title_font = QFont('Arial', 18, QFont.Bold)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("color: #2c3e50; padding: 15px; background-color: #ecf0f1; border-radius: 5px;")
        main_layout.addWidget(title_label)
        
        subtitle_label = QLabel('TP5 - Simulación Monte Carlo con Método de Euler')
        subtitle_label.setAlignment(Qt.AlignCenter)
        subtitle_label.setStyleSheet("color: #7f8c8d; padding: 5px;")
        main_layout.addWidget(subtitle_label)
        
        # Tabs para organizar parámetros y resultados
        self.tabs = QTabWidget()
        self.tabs.setStyleSheet("""
            QTabWidget::pane {
                border: 2px solid #bdc3c7;
                border-radius: 5px;
                background-color: white;
            }
            QTabBar::tab {
                background-color: #ecf0f1;
                padding: 10px 20px;
                margin-right: 2px;
                border-top-left-radius: 5px;
                border-top-right-radius: 5px;
            }
            QTabBar::tab:selected {
                background-color: #3498db;
                color: white;
            }
        """)
        
        # Tab 1: Parámetros de Simulación
        self.create_simulacion_tab()
        
        # Tab 2: Parámetros del Sistema
        self.create_sistema_tab()
        
        # Tab 3: Parámetros de Integración
        self.create_integracion_tab()
        
        # Tab 4: Visualización
        self.create_visualizacion_tab()
        
        # Tab 5: NUEVO - Vector de Estado (Tabla)
        self.create_tabla_tab()
        
        # Tab 6: NUEVO - Métricas y Resultados
        self.create_resultados_tab()
        
        main_layout.addWidget(self.tabs)
        
        # Barra de progreso
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 2px solid #bdc3c7;
                border-radius: 5px;
                text-align: center;
                height: 25px;
            }
            QProgressBar::chunk {
                background-color: #3498db;
            }
        """)
        main_layout.addWidget(self.progress_bar)
        
        # Botones de acción
        self.create_action_buttons(main_layout)
    
    def create_simulacion_tab(self):
        """Crea el tab de parámetros de simulación"""
        tab = QWidget()
        layout = QVBoxLayout()
        
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll_widget = QWidget()
        scroll_layout = QGridLayout()
        
        # Grupo: Parámetros temporales
        group = QGroupBox("⏱️ Parámetros Temporales")
        group.setStyleSheet("QGroupBox { font-weight: bold; font-size: 12pt; }")
        group_layout = QGridLayout()
        
        self.tiempo_max = self.create_spinbox(1, 10000, self.config.TIEMPO_MAXIMO_SIMULACION)
        self.max_iter = self.create_spinbox(100, 1000000, self.config.MAX_ITERACIONES)
        
        group_layout.addWidget(QLabel("Tiempo Máximo de Simulación (min):"), 0, 0)
        group_layout.addWidget(self.tiempo_max, 0, 1)
        group_layout.addWidget(QLabel("Máximo de Iteraciones:"), 1, 0)
        group_layout.addWidget(self.max_iter, 1, 1)
        
        group.setLayout(group_layout)
        scroll_layout.addWidget(group, 0, 0)
        
        scroll_widget.setLayout(scroll_layout)
        scroll.setWidget(scroll_widget)
        layout.addWidget(scroll)
        tab.setLayout(layout)
        self.tabs.addTab(tab, "🎯 Simulación")
    
    def create_sistema_tab(self):
        """Crea el tab de parámetros del sistema"""
        tab = QWidget()
        layout = QVBoxLayout()
        
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout()
        
        # Grupo: Llegadas
        group1 = QGroupBox("🚶 Parámetros de Llegadas")
        group1.setStyleSheet("QGroupBox { font-weight: bold; font-size: 12pt; }")
        group1_layout = QGridLayout()
        
        self.tiempo_llegadas = self.create_double_spinbox(0.1, 100, self.config.TIEMPO_ENTRE_LLEGADAS, 0.1)
        
        group1_layout.addWidget(QLabel("Tiempo Entre Llegadas (min):"), 0, 0)
        group1_layout.addWidget(self.tiempo_llegadas, 0, 1)
        
        group1.setLayout(group1_layout)
        scroll_layout.addWidget(group1)
        
        # Grupo: Distribución de acciones
        group2 = QGroupBox("📊 Distribución de Acciones de Clientes")
        group2.setStyleSheet("QGroupBox { font-weight: bold; font-size: 12pt; background-color: #e8f4f8; }")
        group2_layout = QGridLayout()
        
        self.prob_pedir = self.create_double_spinbox(0, 1, self.config.PROB_PEDIR_LIBRO, 0.01)
        self.prob_devolver = self.create_double_spinbox(0, 1, self.config.PROB_DEVOLVER_LIBRO, 0.01)
        self.prob_consultar = self.create_double_spinbox(0, 1, self.config.PROB_CONSULTAR, 0.01)
        
        group2_layout.addWidget(QLabel("Probabilidad Pedir Libro:"), 0, 0)
        group2_layout.addWidget(self.prob_pedir, 0, 1)
        group2_layout.addWidget(QLabel("Probabilidad Devolver Libro:"), 1, 0)
        group2_layout.addWidget(self.prob_devolver, 1, 1)
        group2_layout.addWidget(QLabel("Probabilidad Consultar:"), 2, 0)
        group2_layout.addWidget(self.prob_consultar, 2, 1)
        
        # Botón normalizar
        btn_normalizar = QPushButton("✓ Normalizar Probabilidades")
        btn_normalizar.clicked.connect(self.normalizar_probabilidades)
        btn_normalizar.setStyleSheet("background-color: #3498db; color: white; padding: 5px;")
        group2_layout.addWidget(btn_normalizar, 3, 0, 1, 2)
        
        group2.setLayout(group2_layout)
        scroll_layout.addWidget(group2)
        
        # Grupo: Tiempos de servicio
        group3 = QGroupBox("⏲️ Tiempos de Servicio")
        group3.setStyleSheet("QGroupBox { font-weight: bold; font-size: 12pt; }")
        group3_layout = QGridLayout()
        
        self.consulta_min = self.create_double_spinbox(0.1, 100, self.config.TIEMPO_CONSULTA_MIN, 0.1)
        self.consulta_max = self.create_double_spinbox(0.1, 100, self.config.TIEMPO_CONSULTA_MAX, 0.1)
        self.media_busqueda = self.create_double_spinbox(0.1, 100, self.config.MEDIA_BUSQUEDA, 0.1)
        
        group3_layout.addWidget(QLabel("Tiempo Consulta Mínimo (min):"), 0, 0)
        group3_layout.addWidget(self.consulta_min, 0, 1)
        group3_layout.addWidget(QLabel("Tiempo Consulta Máximo (min):"), 1, 0)
        group3_layout.addWidget(self.consulta_max, 1, 1)
        group3_layout.addWidget(QLabel("Media Búsqueda Exponencial (min):"), 2, 0)
        group3_layout.addWidget(self.media_busqueda, 2, 1)
        
        group3.setLayout(group3_layout)
        scroll_layout.addWidget(group3)
        
        # Grupo: Comportamiento
        group4 = QGroupBox("🎭 Comportamiento de Clientes")
        group4.setStyleSheet("QGroupBox { font-weight: bold; font-size: 12pt; }")
        group4_layout = QGridLayout()
        
        self.prob_retirarse = self.create_double_spinbox(0, 1, self.config.PROB_RETIRARSE, 0.01)
        self.capacidad = self.create_spinbox(1, 100, self.config.CAPACIDAD_MAXIMA)
        
        group4_layout.addWidget(QLabel("Probabilidad Retirarse con Libro:"), 0, 0)
        group4_layout.addWidget(self.prob_retirarse, 0, 1)
        group4_layout.addWidget(QLabel("Capacidad Máxima Biblioteca:"), 1, 0)
        group4_layout.addWidget(self.capacidad, 1, 1)
        
        group4.setLayout(group4_layout)
        scroll_layout.addWidget(group4)
        
        scroll_layout.addStretch()
        scroll_widget.setLayout(scroll_layout)
        scroll.setWidget(scroll_widget)
        layout.addWidget(scroll)
        tab.setLayout(layout)
        self.tabs.addTab(tab, "👥 Sistema")
    
    def create_integracion_tab(self):
        """Crea el tab de parámetros de integración"""
        tab = QWidget()
        layout = QVBoxLayout()
        
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout()
        
        # Info ecuación
        info_label = QLabel("📐 Ecuación Diferencial: dP/dt = K/5")
        info_label.setStyleSheet("""
            background-color: #fff3cd; 
            padding: 10px; 
            border: 2px solid #ffc107;
            border-radius: 5px;
            font-size: 11pt;
        """)
        scroll_layout.addWidget(info_label)
        
        # Grupo: Constantes K
        group1 = QGroupBox("🔢 Constantes K según Número de Páginas")
        group1.setStyleSheet("QGroupBox { font-weight: bold; font-size: 12pt; }")
        group1_layout = QGridLayout()
        
        self.k_100_200 = self.create_spinbox(1, 500, self.config.K_100_200_PAG)
        self.k_200_300 = self.create_spinbox(1, 500, self.config.K_200_300_PAG)
        self.k_mas_300 = self.create_spinbox(1, 500, self.config.K_MAS_300_PAG)
        
        group1_layout.addWidget(QLabel("K para libros 100-200 páginas:"), 0, 0)
        group1_layout.addWidget(self.k_100_200, 0, 1)
        group1_layout.addWidget(QLabel("K para libros 200-300 páginas:"), 1, 0)
        group1_layout.addWidget(self.k_200_300, 1, 1)
        group1_layout.addWidget(QLabel("K para libros >300 páginas:"), 2, 0)
        group1_layout.addWidget(self.k_mas_300, 2, 1)
        
        group1.setLayout(group1_layout)
        scroll_layout.addWidget(group1)
        
        # Grupo: Rango de páginas
        group2 = QGroupBox("📚 Rango de Páginas de Libros")
        group2.setStyleSheet("QGroupBox { font-weight: bold; font-size: 12pt; }")
        group2_layout = QGridLayout()
        
        self.paginas_min = self.create_spinbox(1, 1000, self.config.PAGINAS_MIN)
        self.paginas_max = self.create_spinbox(1, 1000, self.config.PAGINAS_MAX)
        
        group2_layout.addWidget(QLabel("Páginas Mínimas:"), 0, 0)
        group2_layout.addWidget(self.paginas_min, 0, 1)
        group2_layout.addWidget(QLabel("Páginas Máximas:"), 1, 0)
        group2_layout.addWidget(self.paginas_max, 1, 1)
        
        group2.setLayout(group2_layout)
        scroll_layout.addWidget(group2)
        
        # Grupo: Método de Euler
        group3 = QGroupBox("⚙️ Parámetros del Método de Euler")
        group3.setStyleSheet("QGroupBox { font-weight: bold; font-size: 12pt; background-color: #e8f8f5; }")
        group3_layout = QGridLayout()
        
        self.h_euler = self.create_double_spinbox(0.001, 1, self.config.H_EULER, 0.001)
        
        group3_layout.addWidget(QLabel("Paso de Integración (h):"), 0, 0)
        group3_layout.addWidget(self.h_euler, 0, 1)
        
        info_h = QLabel("⚠️ Valores más pequeños = mayor precisión pero más lento")
        info_h.setStyleSheet("color: #856404; font-size: 9pt; font-style: italic;")
        group3_layout.addWidget(info_h, 1, 0, 1, 2)
        
        group3.setLayout(group3_layout)
        scroll_layout.addWidget(group3)
        
        scroll_layout.addStretch()
        scroll_widget.setLayout(scroll_layout)
        scroll.setWidget(scroll_widget)
        layout.addWidget(scroll)
        tab.setLayout(layout)
        self.tabs.addTab(tab, "📈 Integración")
    
    def create_visualizacion_tab(self):
        """Crea el tab de parámetros de visualización"""
        tab = QWidget()
        layout = QVBoxLayout()
        
        group = QGroupBox("👁️ Parámetros de Visualización del Vector de Estado")
        group.setStyleSheet("QGroupBox { font-weight: bold; font-size: 12pt; }")
        group_layout = QGridLayout()
        
        self.hora_inicio = self.create_spinbox(0, 100000, self.config.HORA_INICIO_MOSTRAR)
        self.filas_mostrar = self.create_spinbox(1, 1000, self.config.FILAS_A_MOSTRAR)
        self.max_clientes = self.create_spinbox(1, 20, 7)
        
        group_layout.addWidget(QLabel("Fila de Inicio (j):"), 0, 0)
        group_layout.addWidget(self.hora_inicio, 0, 1)
        group_layout.addWidget(QLabel("Cantidad de Filas a Mostrar (i):"), 1, 0)
        group_layout.addWidget(self.filas_mostrar, 1, 1)
        group_layout.addWidget(QLabel("Max Clientes Visibles en Tabla:"), 2, 0)
        group_layout.addWidget(self.max_clientes, 2, 1)
        
        group.setLayout(group_layout)
        layout.addWidget(group)
        layout.addStretch()
        tab.setLayout(layout)
        self.tabs.addTab(tab, "📊 Visualización")
    
    def create_tabla_tab(self):
        """Crea el tab con la tabla del vector de estado"""
        tab = QWidget()
        layout = QVBoxLayout()
        
        # Toolbar con controles
        toolbar = QHBoxLayout()
        
        toolbar.addWidget(QLabel("Ir a fila:"))
        self.entry_ir_fila = QSpinBox()
        self.entry_ir_fila.setMinimum(0)
        self.entry_ir_fila.setMaximum(999999)
        toolbar.addWidget(self.entry_ir_fila)
        
        btn_buscar = QPushButton("🔍 Buscar")
        btn_buscar.clicked.connect(self.ir_a_fila)
        toolbar.addWidget(btn_buscar)
        
        btn_primera = QPushButton("⬆ Primera")
        btn_primera.clicked.connect(self.ir_primera_fila)
        toolbar.addWidget(btn_primera)
        
        btn_ultima = QPushButton("⬇ Última")
        btn_ultima.clicked.connect(self.ir_ultima_fila)
        toolbar.addWidget(btn_ultima)
        
        btn_actualizar = QPushButton("🔄 Actualizar")
        btn_actualizar.clicked.connect(self.actualizar_tabla)
        toolbar.addWidget(btn_actualizar)
        
        toolbar.addStretch()
        
        layout.addLayout(toolbar)
        
        # Tabla
        self.tabla_vector = QTableWidget()
        self.tabla_vector.setStyleSheet("""
            QTableWidget {
                gridline-color: #bdc3c7;
                background-color: white;
            }
            QTableWidget::item {
                padding: 5px;
            }
            QHeaderView::section {
                background-color: #3498db;
                color: white;
                padding: 8px;
                font-weight: bold;
                border: 1px solid #2980b9;
            }
        """)
        self.tabla_vector.setAlternatingRowColors(True)
        layout.addWidget(self.tabla_vector)
        
        tab.setLayout(layout)
        self.tabs.addTab(tab, "📋 Vector Estado")
    
    def create_resultados_tab(self):
        """Crea el tab de métricas y resultados"""
        tab = QWidget()
        layout = QVBoxLayout()
        
        # Área de resultados con scroll
        self.results_text = QTextEdit()
        self.results_text.setReadOnly(True)
        self.results_text.setStyleSheet("""
            QTextEdit {
                background-color: #f8f9fa;
                padding: 15px;
                border: 2px solid #dee2e6;
                border-radius: 5px;
                font-family: 'Courier New';
                font-size: 10pt;
            }
        """)
        layout.addWidget(self.results_text)
        
        tab.setLayout(layout)
        self.tabs.addTab(tab, "📊 Resultados")
    
    def create_action_buttons(self, parent_layout):
        """Crea los botones de acción"""
        buttons_layout = QHBoxLayout()
        
        # Botón ejecutar
        self.btn_ejecutar = QPushButton("▶️ Ejecutar Simulación")
        self.btn_ejecutar.setStyleSheet("""
            QPushButton {
                background-color: #27ae60;
                color: white;
                font-size: 14pt;
                font-weight: bold;
                padding: 15px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #229954;
            }
            QPushButton:disabled {
                background-color: #95a5a6;
            }
        """)
        self.btn_ejecutar.clicked.connect(self.ejecutar_simulacion)
        buttons_layout.addWidget(self.btn_ejecutar)
        
        # Botón exportar Excel
        self.btn_exportar = QPushButton("📊 Exportar a Excel")
        self.btn_exportar.setEnabled(False)
        self.btn_exportar.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                font-size: 12pt;
                padding: 15px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
            QPushButton:disabled {
                background-color: #95a5a6;
            }
        """)
        self.btn_exportar.clicked.connect(self.exportar_excel)
        buttons_layout.addWidget(self.btn_exportar)
        
        # Botón guardar config
        btn_guardar_config = QPushButton("💾 Guardar Config")
        btn_guardar_config.setStyleSheet("""
            QPushButton {
                background-color: #f39c12;
                color: white;
                font-size: 12pt;
                padding: 15px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #e67e22;
            }
        """)
        btn_guardar_config.clicked.connect(self.guardar_configuracion)
        buttons_layout.addWidget(btn_guardar_config)
        
        # Botón cargar config
        btn_cargar_config = QPushButton("📂 Cargar Config")
        btn_cargar_config.setStyleSheet("""
            QPushButton {
                background-color: #9b59b6;
                color: white;
                font-size: 12pt;
                padding: 15px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #8e44ad;
            }
        """)
        btn_cargar_config.clicked.connect(self.cargar_configuracion)
        buttons_layout.addWidget(btn_cargar_config)
        
        parent_layout.addLayout(buttons_layout)
    
    def create_spinbox(self, min_val, max_val, default_val):
        """Crea un QSpinBox configurado"""
        spinbox = QSpinBox()
        spinbox.setMinimum(min_val)
        spinbox.setMaximum(max_val)
        spinbox.setValue(int(default_val))
        spinbox.setStyleSheet("padding: 5px; font-size: 10pt;")
        return spinbox
    
    def create_double_spinbox(self, min_val, max_val, default_val, step):
        """Crea un QDoubleSpinBox configurado"""
        spinbox = QDoubleSpinBox()
        spinbox.setMinimum(min_val)
        spinbox.setMaximum(max_val)
        spinbox.setValue(float(default_val))
        spinbox.setSingleStep(step)
        spinbox.setDecimals(3)
        spinbox.setStyleSheet("padding: 5px; font-size: 10pt;")
        return spinbox
    
    def normalizar_probabilidades(self):
        """Normaliza las probabilidades para que sumen 1"""
        suma = self.prob_pedir.value() + self.prob_devolver.value() + self.prob_consultar.value()
        
        if abs(suma - 1.0) > 0.001:
            self.prob_pedir.setValue(self.prob_pedir.value() / suma)
            self.prob_devolver.setValue(self.prob_devolver.value() / suma)
            self.prob_consultar.setValue(self.prob_consultar.value() / suma)
            
            QMessageBox.information(self, "Normalización", 
                f"Probabilidades normalizadas.\nSuma anterior: {suma:.4f}\nSuma actual: 1.0000")
    
    def aplicar_configuracion(self):
        """Aplica los valores de la interfaz a la configuración"""
        # Simulación
        self.config.TIEMPO_MAXIMO_SIMULACION = self.tiempo_max.value()
        self.config.MAX_ITERACIONES = self.max_iter.value()
        self.config.HORA_INICIO_MOSTRAR = self.hora_inicio.value()
        self.config.FILAS_A_MOSTRAR = self.filas_mostrar.value()
        
        # Sistema
        self.config.TIEMPO_ENTRE_LLEGADAS = self.tiempo_llegadas.value()
        self.config.PROB_PEDIR_LIBRO = self.prob_pedir.value()
        self.config.PROB_DEVOLVER_LIBRO = self.prob_devolver.value()
        self.config.PROB_CONSULTAR = self.prob_consultar.value()
        self.config.TIEMPO_CONSULTA_MIN = self.consulta_min.value()
        self.config.TIEMPO_CONSULTA_MAX = self.consulta_max.value()
        self.config.MEDIA_BUSQUEDA = self.media_busqueda.value()
        self.config.PROB_RETIRARSE = self.prob_retirarse.value()
        self.config.PROB_QUEDARSE_LEER = 1.0 - self.prob_retirarse.value()
        self.config.CAPACIDAD_MAXIMA = self.capacidad.value()
        
        # Integración
        self.config.K_100_200_PAG = self.k_100_200.value()
        self.config.K_200_300_PAG = self.k_200_300.value()
        self.config.K_MAS_300_PAG = self.k_mas_300.value()
        self.config.PAGINAS_MIN = self.paginas_min.value()
        self.config.PAGINAS_MAX = self.paginas_max.value()
        self.config.H_EULER = self.h_euler.value()
    
    def ejecutar_simulacion(self):
        """Ejecuta la simulación en un thread separado"""
        self.aplicar_configuracion()
        
        self.btn_ejecutar.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)  # Modo indeterminado
        self.results_text.setText("⏳ Ejecutando simulación...")
        
        # Crear y ejecutar thread
        self.thread = SimulacionThread(self.config)
        self.thread.finished.connect(self.simulacion_completada)
        self.thread.error.connect(self.simulacion_error)
        self.thread.start()
    
    def simulacion_completada(self, vector_estado, metricas, simulador):
        """Callback cuando la simulación termina exitosamente"""
        self.vector_estado = vector_estado
        self.metricas = metricas
        self.simulador = simulador
        
        self.btn_ejecutar.setEnabled(True)
        self.btn_exportar.setEnabled(True)
        self.progress_bar.setVisible(False)
        
        # Mostrar resultados
        resultados = f"""
✅ SIMULACIÓN COMPLETADA EXITOSAMENTE

📊 MÉTRICAS FINALES:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

⏱️  Tiempo simulado: {metricas['tiempo_total_simulado']:.2f} minutos
📝 Filas generadas: {len(vector_estado)}

👥 Total personas llegadas: {metricas['total_personas_llegadas']}
🚪 Total personas salidas: {metricas['total_personas_salidas']}
⏳ Promedio de permanencia: {metricas['promedio_permanencia']:.2f} minutos

🔒 Tiempo biblioteca cerrada: {metricas['porcentaje_tiempo_cerrada']:.2f}%
⛔ Personas no entraron (cerrada): {metricas['personas_no_entraron']}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        """
        
        self.results_text.setText(resultados)
        
        # Actualizar tabla
        self.actualizar_tabla()
        
        # Cambiar al tab de resultados
        self.tabs.setCurrentIndex(5)  # Tab de resultados
        
        QMessageBox.information(self, "Simulación Completada", 
            f"✅ Simulación finalizada exitosamente!\n\n"
            f"Filas generadas: {len(vector_estado)}\n"
            f"Tiempo simulado: {simulador.reloj:.2f} min\n\n"
            f"Puedes ver los resultados en el tab 'Resultados' y la tabla en 'Vector Estado'")
    
    def simulacion_error(self, error_msg):
        """Callback cuando hay un error en la simulación"""
        self.btn_ejecutar.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.results_text.setText(f"❌ Error en la simulación:\n{error_msg}")
        
        QMessageBox.critical(self, "Error", f"Error en la simulación:\n\n{error_msg}")
    
    def actualizar_tabla(self):
        """Actualiza la tabla del vector de estado"""
        if not self.vector_estado:
            return
        
        # Determinar filas a mostrar
        inicio = int(self.config.HORA_INICIO_MOSTRAR)
        fin = min(inicio + self.config.FILAS_A_MOSTRAR, len(self.vector_estado))
        filas_a_mostrar = list(range(inicio, fin))
        
        # Agregar última fila si no está incluida
        ultima_idx = len(self.vector_estado) - 1
        if ultima_idx not in filas_a_mostrar and ultima_idx >= 0:
            filas_a_mostrar.append(ultima_idx)
        
        # Configurar columnas básicas
        columnas = [
            "Fila", "Reloj", "Evento", "Próx Evento", "Tiempo Próx",
            "Personas Dentro", "Cola", "Leyendo", "Estado Bib",
            "E1 Estado", "E1 Atendiendo", "E2 Estado", "E2 Atendiendo",
            "Llegadas", "Salidas", "L.Pedidos", "L.Devueltos", "Consultas"
        ]
        
        self.tabla_vector.setColumnCount(len(columnas))
        self.tabla_vector.setHorizontalHeaderLabels(columnas)
        self.tabla_vector.setRowCount(len(filas_a_mostrar))
        
        # Llenar tabla
        for row_idx, fila_idx in enumerate(filas_a_mostrar):
            fila = self.vector_estado[fila_idx]
            
            # Próximo evento
            prox_evento = ""
            tiempo_prox = ""
            if fila.proximos_eventos:
                prox_evento = fila.proximos_eventos[0]['tipo'][:15]
                tiempo_prox = f"{fila.proximos_eventos[0]['tiempo']:.2f}"
            
            # Datos de la fila
            datos = [
                str(fila.numero_fila),
                f"{fila.reloj:.2f}",
                fila.evento[:15],
                prox_evento,
                tiempo_prox,
                str(fila.biblioteca['personas_dentro']),
                str(len(fila.biblioteca['cola_atencion'])),
                str(len(fila.biblioteca['personas_leyendo'])),
                "CERR" if fila.biblioteca['cerrada'] else "ABIER",
                fila.biblioteca['empleados'][0]['estado'][:8],
                fila.biblioteca['empleados'][0]['persona_atendiendo'] or "",
                fila.biblioteca['empleados'][1]['estado'][:8] if len(fila.biblioteca['empleados']) > 1 else "",
                fila.biblioteca['empleados'][1]['persona_atendiendo'] or "" if len(fila.biblioteca['empleados']) > 1 else "",
                str(fila.acumuladores['total_personas_llegadas']),
                str(fila.acumuladores['total_personas_salidas']),
                str(fila.acumuladores['total_libros_pedidos']),
                str(fila.acumuladores['total_libros_devueltos']),
                str(fila.acumuladores['total_consultas'])
            ]
            
            # Agregar items a la tabla
            for col_idx, dato in enumerate(datos):
                item = QTableWidgetItem(dato)
                item.setTextAlignment(Qt.AlignCenter)
                
                # Colorear última fila
                if fila_idx == ultima_idx and fila_idx != filas_a_mostrar[0]:
                    item.setBackground(QColor("#FFFF00"))
                
                self.tabla_vector.setItem(row_idx, col_idx, item)
        
        # Ajustar columnas
        self.tabla_vector.resizeColumnsToContents()
        self.tabla_vector.horizontalHeader().setStretchLastSection(True)
    
    def ir_a_fila(self):
        """Va a una fila específica"""
        if not self.vector_estado:
            QMessageBox.warning(self, "Advertencia", "No hay datos para mostrar")
            return
        
        fila_num = self.entry_ir_fila.value()
        if 0 <= fila_num < len(self.vector_estado):
            self.config.HORA_INICIO_MOSTRAR = fila_num
            self.actualizar_tabla()
        else:
            QMessageBox.warning(self, "Advertencia", 
                f"La fila debe estar entre 0 y {len(self.vector_estado)-1}")
    
    def ir_primera_fila(self):
        """Va a la primera fila"""
        if not self.vector_estado:
            QMessageBox.warning(self, "Advertencia", "No hay datos para mostrar")
            return
        
        self.config.HORA_INICIO_MOSTRAR = 0
        self.actualizar_tabla()
    
    def ir_ultima_fila(self):
        """Va a la última fila"""
        if not self.vector_estado:
            QMessageBox.warning(self, "Advertencia", "No hay datos para mostrar")
            return
        
        self.config.HORA_INICIO_MOSTRAR = max(0, len(self.vector_estado) - self.config.FILAS_A_MOSTRAR)
        self.actualizar_tabla()
    
    def exportar_excel(self):
        """Exporta los resultados a Excel"""
        if not self.vector_estado:
            QMessageBox.warning(self, "Advertencia", "No hay resultados para exportar")
            return
        
        filename, _ = QFileDialog.getSaveFileName(
            self, 
            "Guardar archivo Excel", 
            "simulacion_biblioteca.xlsx",
            "Excel Files (*.xlsx)"
        )
        
        if filename:
            try:
                exportador = ExportadorExcel(self.vector_estado, self.metricas)
                exportador.exportar(filename)
                
                # Preguntar si desea exportar integraciones
                reply = QMessageBox.question(
                    self, 
                    "Exportar Integraciones",
                    "¿Desea exportar también las integraciones de Euler detalladas?",
                    QMessageBox.Yes | QMessageBox.No
                )
                
                if reply == QMessageBox.Yes:
                    int_filename, _ = QFileDialog.getSaveFileName(
                        self,
                        "Guardar integraciones detalladas",
                        "integraciones_detalladas.xlsx",
                        "Excel Files (*.xlsx)"
                    )
                    
                    if int_filename:
                        exportador.exportar_historial_integraciones_detallado(int_filename)
                        QMessageBox.information(self, "Éxito", 
                            f"✅ Archivos exportados correctamente:\n\n"
                            f"📊 Vector de estado: {filename}\n"
                            f"📈 Integraciones: {int_filename}")
                    else:
                        QMessageBox.information(self, "Éxito", 
                            f"✅ Vector de estado exportado:\n{filename}")
                else:
                    QMessageBox.information(self, "Éxito", 
                        f"✅ Archivo exportado correctamente:\n{filename}")
                        
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Error al exportar:\n\n{str(e)}")
    
    def guardar_configuracion(self):
        """Guarda la configuración actual en un archivo JSON"""
        self.aplicar_configuracion()
        
        filename, _ = QFileDialog.getSaveFileName(
            self,
            "Guardar configuración",
            "config_biblioteca.json",
            "JSON Files (*.json)"
        )
        
        if filename:
            try:
                config_dict = {}
                for attr in dir(self.config):
                    if attr.isupper() and not attr.startswith('_'):
                        config_dict[attr] = getattr(self.config, attr)
                
                with open(filename, 'w', encoding='utf-8') as f:
                    json.dump(config_dict, f, indent=4, ensure_ascii=False)
                
                QMessageBox.information(self, "Éxito", 
                    f"✅ Configuración guardada en:\n{filename}")
            except Exception as e:
                QMessageBox.critical(self, "Error", 
                    f"Error al guardar configuración:\n\n{str(e)}")
    
    def cargar_configuracion(self):
        """Carga una configuración desde un archivo JSON"""
        filename, _ = QFileDialog.getOpenFileName(
            self,
            "Cargar configuración",
            "",
            "JSON Files (*.json)"
        )
        
        if filename:
            try:
                with open(filename, 'r', encoding='utf-8') as f:
                    config_dict = json.load(f)
                
                # Aplicar valores a los widgets
                self.tiempo_max.setValue(int(config_dict.get('TIEMPO_MAXIMO_SIMULACION', 480)))
                self.max_iter.setValue(int(config_dict.get('MAX_ITERACIONES', 100000)))
                self.hora_inicio.setValue(int(config_dict.get('HORA_INICIO_MOSTRAR', 0)))
                self.filas_mostrar.setValue(int(config_dict.get('FILAS_A_MOSTRAR', 50)))
                
                self.tiempo_llegadas.setValue(float(config_dict.get('TIEMPO_ENTRE_LLEGADAS', 4)))
                self.prob_pedir.setValue(float(config_dict.get('PROB_PEDIR_LIBRO', 0.45)))
                self.prob_devolver.setValue(float(config_dict.get('PROB_DEVOLVER_LIBRO', 0.45)))
                self.prob_consultar.setValue(float(config_dict.get('PROB_CONSULTAR', 0.10)))
                self.consulta_min.setValue(float(config_dict.get('TIEMPO_CONSULTA_MIN', 2)))
                self.consulta_max.setValue(float(config_dict.get('TIEMPO_CONSULTA_MAX', 5)))
                self.media_busqueda.setValue(float(config_dict.get('MEDIA_BUSQUEDA', 6)))
                self.prob_retirarse.setValue(float(config_dict.get('PROB_RETIRARSE', 0.60)))
                self.capacidad.setValue(int(config_dict.get('CAPACIDAD_MAXIMA', 20)))
                
                self.k_100_200.setValue(int(config_dict.get('K_100_200_PAG', 100)))
                self.k_200_300.setValue(int(config_dict.get('K_200_300_PAG', 90)))
                self.k_mas_300.setValue(int(config_dict.get('K_MAS_300_PAG', 70)))
                self.paginas_min.setValue(int(config_dict.get('PAGINAS_MIN', 100)))
                self.paginas_max.setValue(int(config_dict.get('PAGINAS_MAX', 350)))
                self.h_euler.setValue(float(config_dict.get('H_EULER', 0.1)))
                
                QMessageBox.information(self, "Éxito", 
                    f"✅ Configuración cargada desde:\n{filename}")
            except Exception as e:
                QMessageBox.critical(self, "Error", 
                    f"Error al cargar configuración:\n\n{str(e)}")


def main():
    """Función principal"""
    app = QApplication(sys.argv)
    
    # Establecer estilo de aplicación
    app.setStyle('Fusion')
    
    ventana = BibliotecaSimuladorGUI()
    ventana.show()
    
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()