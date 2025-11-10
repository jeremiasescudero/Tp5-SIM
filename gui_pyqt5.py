"""
Aplicación PyQt5 para simulación de biblioteca con tabla de eventos
Replica la estructura exacta mostrada en la imagen Excel con encabezados agrupados

INTEGRACIÓN NUMÉRICA DE EULER:
- Se usa para calcular el tiempo de lectura de libros
- Resuelve la ecuación diferencial: dP/dt = K/5
- Donde P = páginas leídas, K = constante según el número de páginas del libro
- El método de Euler integra numéricamente hasta que P >= total_páginas
- Ver clase IntegradorEuler para la implementación
"""
import sys
import random
import heapq
import math
from enum import Enum
from dataclasses import dataclass
from typing import List, Optional, Dict

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QTableWidget, QTableWidgetItem, QHeaderView,
    QLabel, QSpinBox, QDoubleSpinBox, QGroupBox, QFormLayout,
    QProgressBar, QMessageBox, QGridLayout, QScrollArea, QFileDialog
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QColor, QFont

try:
    import openpyxl
    from openpyxl.styles import Font as ExcelFont, PatternFill, Alignment
    OPENPYXL_DISPONIBLE = True
except ImportError:
    OPENPYXL_DISPONIBLE = False


# ==================== INTEGRADOR DE EULER ====================

class IntegradorEuler:
    """
    INTEGRACIÓN NUMÉRICA POR MÉTODO DE EULER

    Resuelve la ecuación diferencial: dP/dt = K/5
    donde:
    - P = páginas leídas
    - t = tiempo
    - K = constante que depende del número de páginas del libro

    El método de Euler usa la fórmula:
    P(t + h) = P(t) + h * f(P, t)
    donde f(P, t) = dP/dt = K/5
    """

    def __init__(self, h: float, K: int, p_inicial: float = 0):
        """
        h: paso de integración (más pequeño = más preciso)
        K: constante según número de páginas
        p_inicial: páginas leídas al inicio
        """
        self.h = h
        self.K = K
        self.p = p_inicial
        self.t = 0.0

    def derivada(self, p: float, t: float) -> float:
        """Función derivada: dP/dt = K/5"""
        return self.K / 5.0

    def paso(self) -> float:
        """Ejecuta un paso del método de Euler"""
        # Euler: p_nuevo = p_actual + h * f(p_actual, t_actual)
        self.p = self.p + self.h * self.derivada(self.p, self.t)
        self.t += self.h
        return self.p

    def integrar_hasta_paginas(self, paginas_objetivo: float) -> float:
        """Integra hasta alcanzar las páginas objetivo y retorna el tiempo total"""
        while self.p < paginas_objetivo:
            self.paso()
        return self.t


# ==================== ENUMS Y DATACLASSES ====================

class TipoEvento(Enum):
    """Tipos de eventos en la simulación"""
    LLEGADA_CLIENTE = "llegada_alquiler"
    FIN_ATENCION = "fin_atencion_cliente"
    FIN_LECTURA = "fin_lectura"
    FIN_SIMULACION = "fin_simulacion"


class TipoObjetivo(Enum):
    """Objetivos del cliente"""
    PEDIR_LIBRO = "Pedir libro"
    DEVOLVER_LIBRO = "Devolver libro"
    CONSULTAR = "Consultar"


class EstadoEmpleado(Enum):
    """Estados del empleado"""
    LIBRE = "Libre"
    OCUPADO = "Ocupado"


@dataclass
class Evento:
    """Representa un evento en la simulación"""
    tiempo: float
    tipo: TipoEvento
    datos: dict

    def __lt__(self, otro):
        return self.tiempo < otro.tiempo


@dataclass
class Cliente:
    """Representa un cliente en el sistema"""
    id: int
    hora_llegada: float
    objetivo: TipoObjetivo
    rnd_objetivo: float
    # Tiempos de servicio según tipo
    rnd_tiempo_consulta: float = 0.0
    tiempo_consulta: float = 0.0
    rnd_tiempo_busqueda: float = 0.0
    tiempo_busqueda: float = 0.0
    rnd_tiempo_devolucion: float = 0.0
    tiempo_devolucion: float = 0.0
    # Fin de atención
    fin_atencion_emp1: Optional[float] = None
    fin_atencion_emp2: Optional[float] = None
    # Lectura
    se_retira: bool = False  # Si se retira o se queda a leer
    rnd_decision: Optional[float] = None
    paginas_a_leer: int = 0
    rnd_paginas: Optional[float] = None
    tiempo_lectura: float = 0.0
    fin_lectura: Optional[float] = None
    # Estado
    estado: str = "En cola"
    hora_salida: Optional[float] = None


class Empleado:
    """Representa un empleado del mostrador"""

    def __init__(self, id: int):
        self.id = id
        self.estado = EstadoEmpleado.LIBRE
        self.cliente_actual: Optional[Cliente] = None
        self.hora_fin_atencion: Optional[float] = None
        self.tiempo_acumulado_atencion = 0.0
        self.tiempo_acumulado_ocioso = 0.0
        self.ultimo_cambio_estado = 0.0

    def esta_libre(self):
        return self.estado == EstadoEmpleado.LIBRE

    def atender(self, cliente: Cliente, tiempo_fin: float, reloj: float):
        if self.estado == EstadoEmpleado.LIBRE:
            self.tiempo_acumulado_ocioso += reloj - self.ultimo_cambio_estado

        self.estado = EstadoEmpleado.OCUPADO
        self.cliente_actual = cliente
        self.hora_fin_atencion = tiempo_fin
        self.ultimo_cambio_estado = reloj

    def liberar(self, reloj: float):
        if self.estado == EstadoEmpleado.OCUPADO:
            self.tiempo_acumulado_atencion += reloj - self.ultimo_cambio_estado

        self.estado = EstadoEmpleado.LIBRE
        cliente = self.cliente_actual
        self.cliente_actual = None
        self.hora_fin_atencion = None
        self.ultimo_cambio_estado = reloj
        return cliente


# ==================== MOTOR DE SIMULACIÓN ====================

class Simulacion:
    """Motor de simulación de eventos discretos"""

    def __init__(self):
        # PARÁMETROS CONFIGURABLES (marcados en rojo en la imagen)
        self.tiempo_entre_llegadas = 4.0  # Tiempo entre llegadas (deterministico)

        # Probabilidades de objetivo
        self.prob_pedir_libro = 0.45
        self.prob_devolver_libro = 0.45
        self.prob_consultar = 0.10

        # Tiempos de consulta U[2, 5]
        self.tiempo_consulta_min = 2.0
        self.tiempo_consulta_max = 5.0

        # Tiempo de búsqueda EXP(media=6)
        self.media_busqueda = 6.0

        # Tiempo de devolución U[2±0.5] = [1.5, 2.5]
        self.tiempo_devolucion_min = 1.5
        self.tiempo_devolucion_max = 2.5

        # Probabilidad de retirarse (60% se retira, 40% se queda)
        self.prob_retirarse = 0.60

        # Páginas U[100, 350]
        self.paginas_min = 100
        self.paginas_max = 350

        # Constantes K para Euler según rango de páginas
        self.K_100_200 = 100
        self.K_200_300 = 90
        self.K_300_plus = 70

        # Sistema
        self.capacidad_maxima = 20
        self.tiempo_maximo = 480.0
        self.h_euler = 0.1  # Paso de integración de Euler (10 min)

        # Estado de la simulación
        self.reloj = 0.0
        self.eventos = []
        self.clientes_activos: List[Cliente] = []
        self.cola_espera: List[Cliente] = []
        self.clientes_leyendo: List[Cliente] = []
        self.empleados = [Empleado(1), Empleado(2)]
        self.contador_clientes = 0
        self.numero_fila = 0

        # Acumuladores
        self.ac_tiempo_permanencia = 0.0
        self.total_clientes_atendidos = 0
        self.total_clientes_leyendo = 0
        self.biblioteca_cerrada = False

        # Historial de filas para la tabla
        self.historial_filas: List[dict] = []

    def determinar_K(self, num_paginas: int) -> int:
        """Determina el valor de K según el número de páginas"""
        if 100 <= num_paginas <= 200:
            return self.K_100_200
        elif 200 < num_paginas <= 300:
            return self.K_200_300
        else:  # más de 300
            return self.K_300_plus

    def generar_objetivo(self) -> tuple:
        rnd = random.random()
        if rnd < self.prob_pedir_libro:
            return TipoObjetivo.PEDIR_LIBRO, rnd
        elif rnd < self.prob_pedir_libro + self.prob_devolver_libro:
            return TipoObjetivo.DEVOLVER_LIBRO, rnd
        else:
            return TipoObjetivo.CONSULTAR, rnd

    def generar_tiempo_consulta(self) -> tuple:
        """Genera tiempo de consulta: Uniforme [2, 5] minutos"""
        rnd = random.random()
        tiempo = self.tiempo_consulta_min + (self.tiempo_consulta_max - self.tiempo_consulta_min) * rnd
        return tiempo, rnd

    def generar_tiempo_busqueda(self) -> tuple:
        """Genera tiempo de búsqueda de libro: Exponencial con media 6 minutos"""
        rnd = random.random()
        tiempo = -self.media_busqueda * math.log(1 - rnd)
        return tiempo, rnd

    def generar_tiempo_devolucion(self) -> tuple:
        """Genera tiempo de devolución: Uniforme [1.5, 2.5] minutos"""
        rnd = random.random()
        tiempo = self.tiempo_devolucion_min + (self.tiempo_devolucion_max - self.tiempo_devolucion_min) * rnd
        return tiempo, rnd

    def agregar_evento(self, evento: Evento):
        heapq.heappush(self.eventos, evento)

    def proximo_evento(self) -> Optional[Evento]:
        if self.eventos:
            return heapq.heappop(self.eventos)
        return None

    def iniciar(self):
        """Inicia la simulación"""
        # Primera llegada a los 4 minutos (deterministico)
        tiempo_primera_llegada = self.tiempo_entre_llegadas
        self.agregar_evento(Evento(
            tiempo=tiempo_primera_llegada,
            tipo=TipoEvento.LLEGADA_CLIENTE,
            datos={}
        ))

        self.agregar_evento(Evento(
            tiempo=self.tiempo_maximo,
            tipo=TipoEvento.FIN_SIMULACION,
            datos={}
        ))

    def procesar_llegada_cliente(self, evento: Evento):
        self.contador_clientes += 1
        objetivo, rnd_objetivo = self.generar_objetivo()
        cliente = Cliente(
            id=self.contador_clientes,
            hora_llegada=self.reloj,
            objetivo=objetivo,
            rnd_objetivo=rnd_objetivo
        )

        self.clientes_activos.append(cliente)

        if len(self.clientes_activos) >= self.capacidad_maxima:
            self.biblioteca_cerrada = True

        empleado_libre = self._obtener_empleado_libre()
        if empleado_libre:
            self._atender_cliente(cliente, empleado_libre)
        else:
            self.cola_espera.append(cliente)
            cliente.estado = "En cola"

        # Próxima llegada cada 4 minutos exactos (deterministico)
        proxima_llegada = self.reloj + self.tiempo_entre_llegadas
        if proxima_llegada < self.tiempo_maximo:
            self.agregar_evento(Evento(
                tiempo=proxima_llegada,
                tipo=TipoEvento.LLEGADA_CLIENTE,
                datos={}
            ))

    def _obtener_empleado_libre(self) -> Optional[Empleado]:
        for emp in self.empleados:
            if emp.esta_libre():
                return emp
        return None

    def _atender_cliente(self, cliente: Cliente, empleado: Empleado):
        """Atiende un cliente según su objetivo"""
        cliente.estado = "Siendo atendido"

        # Determinar tiempo de atención según el tipo de acción
        if cliente.objetivo == TipoObjetivo.CONSULTAR:
            tiempo_atencion, rnd = self.generar_tiempo_consulta()
            cliente.tiempo_consulta = tiempo_atencion
            cliente.rnd_tiempo_consulta = rnd

        elif cliente.objetivo == TipoObjetivo.PEDIR_LIBRO:
            tiempo_atencion, rnd = self.generar_tiempo_busqueda()
            cliente.tiempo_busqueda = tiempo_atencion
            cliente.rnd_tiempo_busqueda = rnd

        else:  # DEVOLVER_LIBRO
            tiempo_atencion, rnd = self.generar_tiempo_devolucion()
            cliente.tiempo_devolucion = tiempo_atencion
            cliente.rnd_tiempo_devolucion = rnd

        tiempo_fin = self.reloj + tiempo_atencion

        if empleado.id == 1:
            cliente.fin_atencion_emp1 = tiempo_fin
        else:
            cliente.fin_atencion_emp2 = tiempo_fin

        empleado.atender(cliente, tiempo_fin, self.reloj)

        self.agregar_evento(Evento(
            tiempo=tiempo_fin,
            tipo=TipoEvento.FIN_ATENCION,
            datos={'cliente': cliente, 'empleado': empleado}
        ))

    def procesar_fin_atencion(self, evento: Evento):
        cliente = evento.datos['cliente']
        empleado = evento.datos['empleado']

        empleado.liberar(self.reloj)

        # Solo los que piden libros pueden quedarse a leer
        if cliente.objetivo == TipoObjetivo.PEDIR_LIBRO:
            rnd_decision = random.random()
            cliente.rnd_decision = rnd_decision

            # 60% se retira, 40% se queda a leer
            if rnd_decision < self.prob_retirarse:
                # Se retira
                cliente.se_retira = True
                cliente.estado = "Fuera del sistema"
                cliente.hora_salida = self.reloj
                self._cliente_sale(cliente)
            else:
                # Se queda a leer (40%)
                cliente.se_retira = False
                cliente.estado = "Leyendo"

                # Generar páginas a leer: U[100, 350]
                rnd_paginas = random.random()
                cliente.rnd_paginas = rnd_paginas
                cliente.paginas_a_leer = int(self.paginas_min +
                                              (self.paginas_max - self.paginas_min) * rnd_paginas)

                # *** APLICACIÓN DEL MÉTODO DE EULER ***
                # Calcular tiempo de lectura usando integración numérica
                K = self.determinar_K(cliente.paginas_a_leer)
                integrador = IntegradorEuler(h=self.h_euler, K=K, p_inicial=0)
                cliente.tiempo_lectura = integrador.integrar_hasta_paginas(cliente.paginas_a_leer)
                cliente.fin_lectura = self.reloj + cliente.tiempo_lectura

                self.clientes_leyendo.append(cliente)
                self.total_clientes_leyendo += 1

                self.agregar_evento(Evento(
                    tiempo=cliente.fin_lectura,
                    tipo=TipoEvento.FIN_LECTURA,
                    datos={'cliente': cliente}
                ))
        else:
            # Consultas y devoluciones se retiran directamente
            cliente.estado = "Fuera del sistema"
            cliente.hora_salida = self.reloj
            self._cliente_sale(cliente)

        # Atender siguiente en cola si hay
        if self.cola_espera and empleado.esta_libre():
            siguiente = self.cola_espera.pop(0)
            self._atender_cliente(siguiente, empleado)

    def procesar_fin_lectura(self, evento: Evento):
        cliente = evento.datos['cliente']
        cliente.estado = "Fuera del sistema"
        cliente.hora_salida = self.reloj

        if cliente in self.clientes_leyendo:
            self.clientes_leyendo.remove(cliente)

        self._cliente_sale(cliente)

    def _cliente_sale(self, cliente: Cliente):
        if cliente in self.clientes_activos:
            self.clientes_activos.remove(cliente)

        if cliente.hora_salida:
            tiempo_permanencia = cliente.hora_salida - cliente.hora_llegada
            self.ac_tiempo_permanencia += tiempo_permanencia

        self.total_clientes_atendidos += 1

        if len(self.clientes_activos) < self.capacidad_maxima:
            self.biblioteca_cerrada = False

    def ejecutar_paso(self) -> Optional[dict]:
        """Ejecuta un paso y retorna los datos para la tabla"""
        evento = self.proximo_evento()
        if not evento or evento.tipo == TipoEvento.FIN_SIMULACION:
            return None

        # Avanzar reloj
        self.reloj = evento.tiempo

        # Procesar evento
        if evento.tipo == TipoEvento.LLEGADA_CLIENTE:
            self.procesar_llegada_cliente(evento)
        elif evento.tipo == TipoEvento.FIN_ATENCION:
            self.procesar_fin_atencion(evento)
        elif evento.tipo == TipoEvento.FIN_LECTURA:
            self.procesar_fin_lectura(evento)

        # Capturar estado DESPUÉS de procesar el evento
        fila_datos = self._capturar_estado(evento)

        self.numero_fila += 1
        self.historial_filas.append(fila_datos)

        return fila_datos

    def ejecutar_completa(self):
        """Ejecuta la simulación completa de una vez"""
        self.iniciar()

        while True:
            resultado = self.ejecutar_paso()
            if resultado is None:
                break

        return self.historial_filas

    def _capturar_estado(self, evento: Evento) -> dict:
        """Captura el estado para la tabla"""
        proximos = sorted(self.eventos, key=lambda e: e.tiempo)[:3]
        cliente_actual = evento.datos.get('cliente')

        # Determinar RND y tiempo de búsqueda de libro (para PEDIR_LIBRO)
        rnd_busqueda = ''
        tiempo_busqueda = ''
        if cliente_actual and cliente_actual.objetivo == TipoObjetivo.PEDIR_LIBRO:
            rnd_busqueda = cliente_actual.rnd_tiempo_busqueda if cliente_actual.rnd_tiempo_busqueda > 0 else ''
            tiempo_busqueda = cliente_actual.tiempo_busqueda if cliente_actual.tiempo_busqueda > 0 else ''

        # Determinar RND y tiempo de devolución (para DEVOLVER_LIBRO)
        rnd_devolucion = ''
        tiempo_devolucion = ''
        if cliente_actual and cliente_actual.objetivo == TipoObjetivo.DEVOLVER_LIBRO:
            rnd_devolucion = cliente_actual.rnd_tiempo_devolucion if cliente_actual.rnd_tiempo_devolucion > 0 else ''
            tiempo_devolucion = cliente_actual.tiempo_devolucion if cliente_actual.tiempo_devolucion > 0 else ''

        # Determinar RND y tiempo de consulta (para CONSULTAR)
        rnd_consulta = ''
        tiempo_consulta = ''
        if cliente_actual and cliente_actual.objetivo == TipoObjetivo.CONSULTAR:
            rnd_consulta = cliente_actual.rnd_tiempo_consulta if cliente_actual.rnd_tiempo_consulta > 0 else ''
            tiempo_consulta = cliente_actual.tiempo_consulta if cliente_actual.tiempo_consulta > 0 else ''

        # Buscar próxima llegada en eventos futuros
        proxima_llegada = ''
        for e in proximos:
            if e.tipo == TipoEvento.LLEGADA_CLIENTE:
                proxima_llegada = e.tiempo
                break

        return {
            'n': self.numero_fila,
            'evento': evento.tipo.value,
            'reloj': self.reloj,
            'tiempo_entre_llegadas': self.tiempo_entre_llegadas if evento.tipo == TipoEvento.LLEGADA_CLIENTE else '',
            'proxima_llegada': proxima_llegada,
            'rnd_llegada': '',  # No hay RND para llegadas determinísticas
            'objetivo': cliente_actual.objetivo.value if cliente_actual else '',
            'rnd_objetivo': cliente_actual.rnd_objetivo if cliente_actual else '',
            # Pedir libro (búsqueda)
            'rnd_busqueda': rnd_busqueda,
            'tiempo_busqueda': tiempo_busqueda,
            'fin_atencion_alq1': cliente_actual.fin_atencion_emp1 if cliente_actual and cliente_actual.objetivo == TipoObjetivo.PEDIR_LIBRO else '',
            'fin_atencion_alq2': cliente_actual.fin_atencion_emp2 if cliente_actual and cliente_actual.objetivo == TipoObjetivo.PEDIR_LIBRO else '',
            # Decisión de quedarse a leer
            'rnd_decision': cliente_actual.rnd_decision if cliente_actual and cliente_actual.rnd_decision is not None else '',
            'se_retira': 'Sí' if cliente_actual and cliente_actual.se_retira else ('No' if cliente_actual and cliente_actual.rnd_decision is not None else ''),
            'rnd_paginas': cliente_actual.rnd_paginas if cliente_actual and cliente_actual.rnd_paginas else '',
            'paginas': cliente_actual.paginas_a_leer if cliente_actual and cliente_actual.paginas_a_leer > 0 else '',
            'tiempo_lectura': cliente_actual.tiempo_lectura if cliente_actual and cliente_actual.tiempo_lectura > 0 else '',
            # Devolución
            'rnd_devolucion': rnd_devolucion,
            'tiempo_devolucion': tiempo_devolucion,
            'fin_atencion_dev1': cliente_actual.fin_atencion_emp1 if cliente_actual and cliente_actual.objetivo == TipoObjetivo.DEVOLVER_LIBRO else '',
            'fin_atencion_dev2': cliente_actual.fin_atencion_emp2 if cliente_actual and cliente_actual.objetivo == TipoObjetivo.DEVOLVER_LIBRO else '',
            # Consulta
            'rnd_consulta': rnd_consulta,
            'tiempo_consulta': tiempo_consulta,
            'fin_atencion_cons': (cliente_actual.fin_atencion_emp1 or cliente_actual.fin_atencion_emp2) if cliente_actual and cliente_actual.objetivo == TipoObjetivo.CONSULTAR else '',
            # Empleados
            'empleado1_estado': self.empleados[0].estado.value,
            'empleado1_ac_atencion': self.empleados[0].tiempo_acumulado_atencion,
            'empleado1_ac_ocioso': self.empleados[0].tiempo_acumulado_ocioso,
            'empleado2_estado': self.empleados[1].estado.value,
            'empleado2_ac_atencion': self.empleados[1].tiempo_acumulado_atencion,
            'empleado2_ac_ocioso': self.empleados[1].tiempo_acumulado_ocioso,
            # Biblioteca
            'estado_biblioteca': 'Cerrada' if self.biblioteca_cerrada else 'Abierta',
            'cola': len(self.cola_espera),
            'ac_tiempo_permanencia': self.ac_tiempo_permanencia,
            'ac_clientes_leyendo': self.total_clientes_leyendo,
            'clientes': self.clientes_activos.copy()
        }


# ==================== THREAD DE SIMULACIÓN ====================

class SimulacionThread(QThread):
    """Thread para ejecutar la simulación sin bloquear la UI"""
    finished = pyqtSignal(object)  # historial_filas
    error = pyqtSignal(str)
    progress = pyqtSignal(int)

    def __init__(self, simulacion: Simulacion):
        super().__init__()
        self.simulacion = simulacion

    def run(self):
        try:
            self.simulacion.iniciar()
            contador = 0

            while True:
                resultado = self.simulacion.ejecutar_paso()
                if resultado is None:
                    break

                contador += 1
                if contador % 10 == 0:  # Actualizar progreso cada 10 eventos
                    self.progress.emit(contador)

            self.finished.emit(self.simulacion.historial_filas)

        except Exception as e:
            self.error.emit(str(e))


# ==================== INTERFAZ GRÁFICA ====================

class MainWindow(QMainWindow):
    """Ventana principal de la aplicación"""

    def __init__(self):
        super().__init__()
        self.simulacion = None
        self.historial_filas = []
        self.thread = None

        self.init_ui()

    def init_ui(self):
        """Inicializa la interfaz"""
        self.setWindowTitle("Simulación de Biblioteca - Sistema de Eventos Discretos con Método de Euler")
        self.setGeometry(50, 50, 1800, 950)

        # Widget central
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # Layout principal
        main_layout = QVBoxLayout()
        central_widget.setLayout(main_layout)

        # Título
        title = QLabel("🏛️ SIMULACIÓN DE BIBLIOTECA - EVENTOS DISCRETOS + INTEGRACIÓN DE EULER")
        title.setFont(QFont("Arial", 16, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("background-color: #4472C4; color: white; padding: 15px; border-radius: 5px;")
        main_layout.addWidget(title)

        # Info sobre Euler
        info_euler = QLabel("📐 Integración Numérica: dP/dt = K/5 (Método de Euler para calcular tiempo de lectura)")
        info_euler.setFont(QFont("Arial", 10))
        info_euler.setAlignment(Qt.AlignCenter)
        info_euler.setStyleSheet("background-color: #fff3cd; padding: 8px; border-radius: 3px; color: #856404;")
        main_layout.addWidget(info_euler)

        # Panel de control
        control_panel = self.crear_panel_control()
        main_layout.addWidget(control_panel)

        # Scroll para la tabla
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)

        self.tabla = QTableWidget()
        self.tabla.setStyleSheet("""
            QTableWidget {
                gridline-color: #d0d0d0;
                font-size: 9pt;
                background-color: white;
            }
            QTableWidget::item {
                padding: 4px;
            }
            QHeaderView::section {
                background-color: #4472C4;
                color: white;
                font-weight: bold;
                padding: 6px;
                border: 1px solid #2a52a4;
            }
            QTableWidget::item:alternate {
                background-color: #f9f9f9;
            }
        """)
        self.tabla.setAlternatingRowColors(True)

        scroll.setWidget(self.tabla)
        main_layout.addWidget(scroll)

        # Barra de progreso
        self.progress_bar = QProgressBar()
        self.progress_bar.setStyleSheet("QProgressBar { border: 2px solid #ccc; border-radius: 5px; text-align: center; } QProgressBar::chunk { background-color: #4CAF50; }")
        self.progress_bar.setVisible(False)
        main_layout.addWidget(self.progress_bar)

        # Labels de estado
        status_layout = QHBoxLayout()
        self.lbl_status = QLabel("⏰ Esperando...")
        self.lbl_status.setFont(QFont("Arial", 10, QFont.Bold))
        status_layout.addWidget(self.lbl_status)
        status_layout.addStretch()
        main_layout.addLayout(status_layout)

    def crear_panel_control(self) -> QGroupBox:
        """Crea el panel de control"""
        group = QGroupBox("🎮 Controles de Simulación")
        group.setFont(QFont("Arial", 11, QFont.Bold))
        layout = QHBoxLayout()

        # Botón ejecutar
        self.btn_ejecutar = QPushButton("▶ Ejecutar Simulación Completa")
        self.btn_ejecutar.setStyleSheet("background-color: #4CAF50; color: white; padding: 12px 24px; font-size: 13pt; font-weight: bold; border-radius: 5px;")
        self.btn_ejecutar.clicked.connect(self.ejecutar_simulacion)
        layout.addWidget(self.btn_ejecutar)

        # Botón exportar
        self.btn_exportar = QPushButton("📊 Exportar a Excel")
        self.btn_exportar.setStyleSheet("background-color: #2196F3; color: white; padding: 12px 24px; font-size: 13pt; font-weight: bold; border-radius: 5px;")
        self.btn_exportar.clicked.connect(self.exportar_excel)
        self.btn_exportar.setEnabled(False)
        layout.addWidget(self.btn_exportar)

        # Botón reiniciar
        self.btn_reiniciar = QPushButton("🔄 Nueva Simulación")
        self.btn_reiniciar.setStyleSheet("background-color: #F44336; color: white; padding: 12px 24px; font-size: 13pt; font-weight: bold; border-radius: 5px;")
        self.btn_reiniciar.clicked.connect(self.reiniciar)
        layout.addWidget(self.btn_reiniciar)

        layout.addStretch()
        group.setLayout(layout)
        return group

    def ejecutar_simulacion(self):
        """Ejecuta la simulación completa"""
        self.btn_ejecutar.setEnabled(False)
        self.btn_exportar.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)  # Modo indeterminado
        self.lbl_status.setText("⏳ Ejecutando simulación completa...")

        # Crear nueva simulación
        self.simulacion = Simulacion()

        # Ejecutar en thread
        self.thread = SimulacionThread(self.simulacion)
        self.thread.finished.connect(self.simulacion_completada)
        self.thread.error.connect(self.simulacion_error)
        self.thread.progress.connect(self.actualizar_progreso)
        self.thread.start()

    def actualizar_progreso(self, eventos):
        """Actualiza el progreso"""
        self.lbl_status.setText(f"⏳ Procesando... {eventos} eventos")

    def simulacion_completada(self, historial_filas):
        """Callback cuando termina la simulación"""
        self.historial_filas = historial_filas
        self.progress_bar.setVisible(False)
        self.btn_ejecutar.setEnabled(True)
        self.btn_exportar.setEnabled(True)

        # Poblar tabla con COLUMNAS DINÁMICAS
        self.poblar_tabla()

        self.lbl_status.setText(f"✅ Simulación completa: {len(historial_filas)} eventos | "
                                f"Tiempo: {self.simulacion.reloj:.2f} min | "
                                f"Clientes atendidos: {self.simulacion.total_clientes_atendidos} | "
                                f"Clientes que leyeron: {self.simulacion.total_clientes_leyendo}")

        QMessageBox.information(self, "Simulación Completa",
                                f"✅ Simulación finalizada exitosamente\n\n"
                                f"📊 Eventos procesados: {len(historial_filas)}\n"
                                f"⏱️ Tiempo simulado: {self.simulacion.reloj:.2f} min\n"
                                f"👥 Clientes atendidos: {self.simulacion.total_clientes_atendidos}\n"
                                f"📖 Clientes que leyeron: {self.simulacion.total_clientes_leyendo}")

    def simulacion_error(self, error_msg):
        """Callback cuando hay error"""
        self.progress_bar.setVisible(False)
        self.btn_ejecutar.setEnabled(True)
        self.lbl_status.setText("❌ Error en la simulación")
        QMessageBox.critical(self, "Error", f"Error en la simulación:\n\n{error_msg}")

    def poblar_tabla(self):
        """Pobla la tabla con columnas dinámicas según los clientes"""
        if not self.historial_filas:
            return

        # Determinar todos los clientes únicos que aparecen
        todos_clientes_ids = set()
        for fila in self.historial_filas:
            for cliente in fila['clientes']:
                todos_clientes_ids.add(cliente.id)

        clientes_ordenados = sorted(list(todos_clientes_ids))

        # COLUMNAS FIJAS - Actualizadas según la especificación
        columnas_fijas = [
            "n°", "Evento", "Reloj",
            "T.entre llegadas", "Próxima llegada", "RND llegada",
            "Objetivo", "RND obj",
            # Pedir libro (búsqueda exponencial)
            "RND búsq", "T.búsqueda (EXP)", "fin_búsq_emp1", "fin_búsq_emp2",
            # Decisión de leer
            "RND decisión", "Se retira?", "RND pág", "Páginas", "T.lectura (Euler)",
            # Devolución
            "RND dev", "T.devolución", "fin_dev_emp1", "fin_dev_emp2",
            # Consulta
            "RND cons", "T.consulta", "fin_cons",
            # Empleados
            "Emp1", "AC at1", "AC oc1", "Emp2", "AC at2", "AC oc2",
            # Biblioteca
            "Estado", "Cola", "AC perm", "AC leyendo"
        ]

        # COLUMNAS DINÁMICAS POR CADA CLIENTE
        columnas_clientes = []
        for cid in clientes_ordenados:
            columnas_clientes.extend([
                f"C{cid} Estado",
                f"C{cid} Hora",
                f"C{cid} T.at",
                f"C{cid} Fin lect"
            ])

        todas_columnas = columnas_fijas + columnas_clientes

        # Configurar tabla
        self.tabla.setColumnCount(len(todas_columnas))
        self.tabla.setHorizontalHeaderLabels(todas_columnas)
        self.tabla.setRowCount(len(self.historial_filas))

        # Configurar header
        header = self.tabla.horizontalHeader()
        header.setDefaultSectionSize(85)
        header.setSectionResizeMode(QHeaderView.Interactive)

        # Llenar filas
        for row, fila in enumerate(self.historial_filas):
            self.agregar_fila(row, fila, clientes_ordenados)

    def agregar_fila(self, row: int, datos: dict, clientes_ordenados: List[int]):
        """Agrega una fila a la tabla"""
        def fmt(val):
            if val == '' or val is None:
                return ''
            if isinstance(val, float):
                return f"{val:.2f}"
            return str(val)

        # Valores fijos - orden según las nuevas columnas
        valores = [
            datos['n'],                                 # n°
            datos['evento'],                            # Evento
            fmt(datos['reloj']),                        # Reloj
            fmt(datos['tiempo_entre_llegadas']),        # T.entre llegadas
            fmt(datos['proxima_llegada']),              # Próxima llegada
            fmt(datos['rnd_llegada']),                  # RND llegada
            datos['objetivo'],                          # Objetivo
            fmt(datos['rnd_objetivo']),                 # RND obj
            # Pedir libro (búsqueda)
            fmt(datos['rnd_busqueda']),                 # RND búsq
            fmt(datos['tiempo_busqueda']),              # T.búsqueda (EXP)
            fmt(datos['fin_atencion_alq1']),            # fin_búsq_emp1
            fmt(datos['fin_atencion_alq2']),            # fin_búsq_emp2
            # Decisión de leer
            fmt(datos['rnd_decision']),                 # RND decisión
            datos['se_retira'],                         # Se retira?
            fmt(datos['rnd_paginas']),                  # RND pág
            datos['paginas'],                           # Páginas
            fmt(datos['tiempo_lectura']),               # T.lectura (Euler) ***
            # Devolución
            fmt(datos['rnd_devolucion']),               # RND dev
            fmt(datos['tiempo_devolucion']),            # T.devolución
            fmt(datos['fin_atencion_dev1']),            # fin_dev_emp1
            fmt(datos['fin_atencion_dev2']),            # fin_dev_emp2
            # Consulta
            fmt(datos['rnd_consulta']),                 # RND cons
            fmt(datos['tiempo_consulta']),              # T.consulta
            fmt(datos['fin_atencion_cons']),            # fin_cons
            # Empleados
            datos['empleado1_estado'],                  # Emp1
            fmt(datos['empleado1_ac_atencion']),        # AC at1
            fmt(datos['empleado1_ac_ocioso']),          # AC oc1
            datos['empleado2_estado'],                  # Emp2
            fmt(datos['empleado2_ac_atencion']),        # AC at2
            fmt(datos['empleado2_ac_ocioso']),          # AC oc2
            # Biblioteca
            datos['estado_biblioteca'],                 # Estado
            str(datos['cola']),                         # Cola
            fmt(datos['ac_tiempo_permanencia']),        # AC perm
            str(datos['ac_clientes_leyendo'])           # AC leyendo
        ]

        # Valores dinámicos por cliente
        clientes_dict = {c.id: c for c in datos['clientes']}
        for cid in clientes_ordenados:
            if cid in clientes_dict:
                c = clientes_dict[cid]
                # Determinar tiempo de atención según objetivo
                tiempo_at = ''
                if c.tiempo_busqueda > 0:
                    tiempo_at = c.tiempo_busqueda
                elif c.tiempo_devolucion > 0:
                    tiempo_at = c.tiempo_devolucion
                elif c.tiempo_consulta > 0:
                    tiempo_at = c.tiempo_consulta

                valores.extend([
                    c.estado,
                    fmt(c.hora_llegada),
                    fmt(tiempo_at) if tiempo_at else '',
                    fmt(c.fin_lectura) if c.fin_lectura else ''
                ])
            else:
                valores.extend(['', '', '', ''])

        # Insertar valores
        for col, valor in enumerate(valores):
            item = QTableWidgetItem(str(valor))
            item.setTextAlignment(Qt.AlignCenter)

            # Colores
            if row % 2 == 0:
                item.setBackground(QColor(249, 249, 249))

            if col == 1:  # Columna Evento
                if 'llegada' in str(valor).lower():
                    item.setBackground(QColor(200, 230, 201))
                elif 'fin_atencion' in str(valor).lower():
                    item.setBackground(QColor(255, 224, 178))
                elif 'fin_lectura' in str(valor).lower():
                    item.setBackground(QColor(187, 222, 251))

            self.tabla.setItem(row, col, item)

    def exportar_excel(self):
        """Exporta la tabla a Excel"""
        if not OPENPYXL_DISPONIBLE:
            QMessageBox.warning(self, "Advertencia",
                                "La librería 'openpyxl' no está instalada.\n\n"
                                "Instálala con: pip install openpyxl")
            return

        if not self.historial_filas:
            QMessageBox.warning(self, "Advertencia", "No hay datos para exportar")
            return

        filename, _ = QFileDialog.getSaveFileName(
            self,
            "Guardar archivo Excel",
            "simulacion_biblioteca.xlsx",
            "Excel Files (*.xlsx)"
        )

        if not filename:
            return

        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Simulación"

            # Estilos
            header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
            header_font = ExcelFont(bold=True, color="FFFFFF")
            center_align = Alignment(horizontal="center", vertical="center")

            # Headers
            for col in range(self.tabla.columnCount()):
                cell = ws.cell(row=1, column=col + 1)
                cell.value = self.tabla.horizontalHeaderItem(col).text()
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = center_align

            # Datos
            for row in range(self.tabla.rowCount()):
                for col in range(self.tabla.columnCount()):
                    item = self.tabla.item(row, col)
                    cell = ws.cell(row=row + 2, column=col + 1)
                    if item:
                        cell.value = item.text()
                        cell.alignment = center_align

            wb.save(filename)
            QMessageBox.information(self, "Éxito",
                                    f"✅ Archivo exportado correctamente:\n{filename}")

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al exportar:\n\n{str(e)}")

    def reiniciar(self):
        """Reinicia la aplicación"""
        self.simulacion = None
        self.historial_filas = []
        self.tabla.setRowCount(0)
        self.tabla.setColumnCount(0)
        self.btn_exportar.setEnabled(False)
        self.lbl_status.setText("⏰ Esperando...")


def main():
    """Función principal"""
    app = QApplication(sys.argv)
    app.setStyle('Fusion')

    window = MainWindow()
    window.show()

    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
