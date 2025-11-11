"""
Aplicación PyQt5 para simulación de biblioteca con tabla de eventos
CON INTERFAZ DE CONFIGURACIÓN MÍNIMA Y PARAMETRIZADA.
Incluye:
1. Parametrización de los 4 datos requeridos (llegadas, objetivos, consulta, retiro).
2. Uso de copia profunda para corregir el bug de los estados de los clientes en la tabla.
3. Eliminación de la columna "Estado" y actualización de índices.
4. Resaltado en rojo del próximo evento.
5. CORRECCIÓN DEL ERROR: NameError: name 'clientes_dict' is not defined.
"""
import sys
import random
import heapq
import math
import json
import os
from enum import Enum
from dataclasses import dataclass
from typing import List, Optional, Dict
import copy 

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QTableWidget, QTableWidgetItem, QHeaderView,
    QLabel, QSpinBox, QDoubleSpinBox, QGroupBox, QFormLayout,
    QProgressBar, QMessageBox, QGridLayout, QScrollArea, QFileDialog,
    QDialog, QLineEdit, QTabWidget
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
    """INTEGRACIÓN NUMÉRICA POR MÉTODO DE EULER"""
    def __init__(self, h: float, K: int, p_inicial: float = 0):
        self.h = h
        self.K = K
        self.p = p_inicial
        self.t = 0.0
        self.historial_pasos = []  # Guardar el historial de pasos

    def derivada(self, p: float, t: float) -> float:
        return self.K / 5.0

    def paso(self) -> float:
        derivada_actual = self.derivada(self.p, self.t)
        # Guardar estado antes del paso
        self.historial_pasos.append({
            't': self.t,
            'p': self.p,
            'derivada': derivada_actual,
            'h': self.h
        })
        self.p = self.p + self.h * derivada_actual
        self.t += self.h
        return self.p

    def integrar_hasta_paginas(self, paginas_objetivo: float) -> float:
        while self.p < paginas_objetivo:
            self.paso()
        # Agregar el estado final
        self.historial_pasos.append({
            't': self.t,
            'p': self.p,
            'derivada': self.derivada(self.p, self.t),
            'h': self.h
        })
        return self.t


# ==================== ENUMS Y DATACLASSES ====================

class TipoEvento(Enum):
    INICIALIZACION = "inicializacion"
    LLEGADA_CLIENTE = "llegada_cliente"
    FIN_ATENCION = "fin_atencion_cliente"
    FIN_LECTURA = "fin_lectura"
    FIN_SIMULACION = "fin_simulacion"
class TipoObjetivo(Enum):
    PEDIR_LIBRO = "Pedir libro"
    DEVOLVER_LIBRO = "Devolver libro"
    CONSULTAR = "Consultar"
class EstadoEmpleado(Enum):
    LIBRE = "Libre"
    OCUPADO = "Ocupado"
@dataclass
class Evento:
    tiempo: float
    tipo: TipoEvento
    datos: dict
    def __lt__(self, otro):
        return self.tiempo < otro.tiempo
@dataclass
class Cliente:
    id: int
    hora_llegada: float
    objetivo: TipoObjetivo
    rnd_objetivo: float
    rnd_tiempo_consulta: float = 0.0
    tiempo_consulta: float = 0.0
    rnd_tiempo_busqueda: float = 0.0
    tiempo_busqueda: float = 0.0
    rnd_tiempo_devolucion: float = 0.0
    tiempo_devolucion: float = 0.0
    fin_atencion_emp1: Optional[float] = None
    fin_atencion_emp2: Optional[float] = None
    se_retira: bool = False
    rnd_decision: Optional[float] = None
    paginas_a_leer: int = 0
    rnd_paginas: Optional[float] = None
    tiempo_lectura: float = 0.0
    fin_lectura: Optional[float] = None
    estado: str = "En cola"
    hora_salida: Optional[float] = None
class Empleado:
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

    def __init__(self, parametros: dict): # Acepta parámetros
        
        # --- PARÁMETROS CONFIGURABLES ---
        self.tiempo_entre_llegadas = parametros.get('tiempo_entre_llegadas', 4.0)
        self.prob_pedir_libro = parametros.get('prob_pedir_libro', 0.45)
        self.prob_devolver_libro = parametros.get('prob_devolver_libro', 0.45)
        self.prob_consultar = parametros.get('prob_consultar', 0.10)
        self.tiempo_consulta_min = parametros.get('tiempo_consulta_min', 2.0)
        self.tiempo_consulta_max = parametros.get('tiempo_consulta_max', 5.0)
        self.prob_retirarse = parametros.get('prob_retirarse', 0.60)
        
        # --- PARÁMETROS FIJOS (VALORES PREDETERMINADOS) ---
        self.media_busqueda = 6.0
        self.tiempo_devolucion_min = 1.5
        self.tiempo_devolucion_max = 2.5
        self.paginas_min = 100
        self.paginas_max = 350
        self.K_100_200 = parametros.get('k_100_200', 100)
        self.K_200_300 = parametros.get('k_200_300', 90)
        self.K_300_plus = parametros.get('k_300_plus', 70)
        self.capacidad_maxima = 20
        self.tiempo_maximo = parametros.get('tiempo_simulacion', 480.0)
        self.max_iteraciones = parametros.get('max_iteraciones', 100000)
        self.hora_desde = parametros.get('hora_desde', 0)
        self.hora_hasta = parametros.get('hora_hasta', 100000.0)
        self.h_euler = 0.1

        # Estado de la simulación
        self.reloj = 0.0
        self.eventos = []
        self.clientes_activos: List[Cliente] = []
        self.cola_espera: List[Cliente] = []
        self.clientes_leyendo: List[Cliente] = []
        self.empleados = [Empleado(1), Empleado(2)]
        self.contador_clientes = 0
        self.numero_fila = 0
        self.ultimo_reloj = 0.0
        self.iteracion_actual = 0

        # Acumuladores
        self.ac_tiempo_permanencia = 0.0
        self.total_clientes_atendidos = 0
        self.total_clientes_leyendo = 0
        self.total_clientes_generados = 0
        self.total_clientes_rechazados = 0
        self.biblioteca_cerrada = False
        self.tiempo_biblioteca_cerrada_ac = 0.0
        self.tiempo_inicio_cerrada: Optional[float] = None

        self.historial_filas: List[dict] = []
        self.historial_integraciones: List[dict] = []  # Guardar todas las integraciones

    def determinar_K(self, num_paginas: int) -> int:
        if 100 <= num_paginas <= 200:
            return self.K_100_200
        elif 200 < num_paginas <= 300:
            return self.K_200_300
        else:
            return self.K_300_plus

    def contar_personas_leyendo(self) -> int:
        """Cuenta las personas que están LEYENDO dentro de la biblioteca"""
        # Según el enunciado, solo se cuentan las personas LEYENDO
        return len(self.clientes_leyendo)

    def generar_objetivo(self) -> tuple:
        rnd = random.random()
        if rnd < self.prob_pedir_libro:
            return TipoObjetivo.PEDIR_LIBRO, rnd
        elif rnd < self.prob_pedir_libro + self.prob_devolver_libro:
            return TipoObjetivo.DEVOLVER_LIBRO, rnd
        else:
            return TipoObjetivo.CONSULTAR, rnd

    def generar_tiempo_consulta(self) -> tuple:
        rnd = random.random()
        tiempo = self.tiempo_consulta_min + (self.tiempo_consulta_max - self.tiempo_consulta_min) * rnd
        return tiempo, rnd

    def generar_tiempo_busqueda(self) -> tuple:
        rnd = random.random()
        media = self.media_busqueda if self.media_busqueda > 0 else 1e-6
        tiempo = -media * math.log(1 - rnd)
        return tiempo, rnd

    def generar_tiempo_devolucion(self) -> tuple:
        rnd = random.random()
        tiempo = self.tiempo_devolucion_min + (self.tiempo_devolucion_max - self.tiempo_devolucion_min) * rnd
        return tiempo, rnd

    def agregar_evento(self, evento: Evento):
        heapq.heappush(self.eventos, evento)

    def proximo_evento(self) -> Optional[Evento]:
        if self.eventos:
            return heapq.heappop(self.eventos)
        return None

    def actualizar_tiempo_cerrada(self):
        if self.biblioteca_cerrada and self.tiempo_inicio_cerrada is not None:
            self.tiempo_biblioteca_cerrada_ac += self.reloj - self.tiempo_inicio_cerrada
            self.tiempo_inicio_cerrada = self.reloj

    def iniciar(self):
        self.numero_fila = 0
        self.iteracion_actual = 0
        self.historial_filas.append(self._capturar_estado(
            Evento(tiempo=0.0, tipo=TipoEvento.INICIALIZACION, datos={})
        ))
        self.numero_fila += 1
        self.ultimo_reloj = 0.0

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
        self.total_clientes_generados += 1

        # Verificar si la biblioteca está llena (20 personas LEYENDO)
        personas_leyendo = self.contar_personas_leyendo()

        if personas_leyendo >= self.capacidad_maxima:
            self.total_clientes_rechazados += 1

            proxima_llegada = self.reloj + self.tiempo_entre_llegadas
            if proxima_llegada < self.tiempo_maximo:
                self.agregar_evento(Evento(
                    tiempo=proxima_llegada,
                    tipo=TipoEvento.LLEGADA_CLIENTE,
                    datos={}
                ))

            cliente_rechazado = Cliente(
                id=self.total_clientes_generados,
                hora_llegada=self.reloj,
                objetivo=TipoObjetivo.CONSULTAR,
                rnd_objetivo=0.0,
                estado="RECHAZADO"
            )
            evento.datos['cliente'] = cliente_rechazado
            return

        self.contador_clientes += 1
        objetivo, rnd_objetivo = self.generar_objetivo()
        cliente = Cliente(
            id=self.contador_clientes,
            hora_llegada=self.reloj,
            objetivo=objetivo,
            rnd_objetivo=rnd_objetivo
        )

        self.clientes_activos.append(cliente)
        evento.datos['cliente'] = cliente

        # Actualizar estado de la biblioteca basado en personas leyendo
        if self.contar_personas_leyendo() >= self.capacidad_maxima:
            self.biblioteca_cerrada = True
            self.tiempo_inicio_cerrada = self.reloj

        empleado_libre = self._obtener_empleado_libre()
        if empleado_libre:
            self._atender_cliente(cliente, empleado_libre)
        else:
            self.cola_espera.append(cliente)
            cliente.estado = "En cola"

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
        cliente.estado = "Siendo atendido"

        if cliente.objetivo == TipoObjetivo.CONSULTAR:
            tiempo_atencion, rnd = self.generar_tiempo_consulta()
            cliente.tiempo_consulta = tiempo_atencion
            cliente.rnd_tiempo_consulta = rnd

        elif cliente.objetivo == TipoObjetivo.PEDIR_LIBRO:
            tiempo_atencion, rnd = self.generar_tiempo_busqueda()
            cliente.tiempo_busqueda = tiempo_atencion
            cliente.rnd_tiempo_busqueda = rnd

        else:
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

        if cliente.objetivo == TipoObjetivo.PEDIR_LIBRO:
            rnd_decision = random.random()
            cliente.rnd_decision = rnd_decision

            if rnd_decision < self.prob_retirarse:
                cliente.se_retira = True
                cliente.estado = "Fuera del sistema" 
                cliente.hora_salida = self.reloj
                self._cliente_sale(cliente)
            else:
                cliente.se_retira = False
                cliente.estado = "Leyendo" 

                rnd_paginas = random.random()
                cliente.rnd_paginas = rnd_paginas
                cliente.paginas_a_leer = int(self.paginas_min +
                                             (self.paginas_max - self.paginas_min) * rnd_paginas)

                K = self.determinar_K(cliente.paginas_a_leer)
                integrador = IntegradorEuler(h=self.h_euler, K=K, p_inicial=0)
                cliente.tiempo_lectura = integrador.integrar_hasta_paginas(cliente.paginas_a_leer)
                cliente.fin_lectura = self.reloj + cliente.tiempo_lectura

                # Guardar el historial de integración para este cliente
                self.historial_integraciones.append({
                    'cliente_id': cliente.id,
                    'paginas_objetivo': cliente.paginas_a_leer,
                    'K': K,
                    'h': self.h_euler,
                    'tiempo_total': cliente.tiempo_lectura,
                    'pasos': integrador.historial_pasos
                })

                self.clientes_leyendo.append(cliente)
                self.total_clientes_leyendo += 1

                self.agregar_evento(Evento(
                    tiempo=cliente.fin_lectura,
                    tipo=TipoEvento.FIN_LECTURA,
                    datos={'cliente': cliente}
                ))
        else:
            cliente.estado = "Fuera del sistema" 
            cliente.hora_salida = self.reloj
            self._cliente_sale(cliente)

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

        # Reabrir biblioteca si hay espacio disponible (menos de 20 leyendo)
        if self.biblioteca_cerrada and self.contar_personas_leyendo() < self.capacidad_maxima:
            self.biblioteca_cerrada = False
            if self.tiempo_inicio_cerrada is not None:
                self.tiempo_biblioteca_cerrada_ac += self.reloj - self.tiempo_inicio_cerrada
            self.tiempo_inicio_cerrada = None

    def ejecutar_paso(self) -> Optional[dict]:
        # Verificar límite de iteraciones
        if self.iteracion_actual >= self.max_iteraciones:
            if self.biblioteca_cerrada and self.tiempo_inicio_cerrada is not None:
                self.tiempo_biblioteca_cerrada_ac += self.tiempo_maximo - self.reloj
            return None

        evento = self.proximo_evento()
        if not evento or evento.tipo == TipoEvento.FIN_SIMULACION:
            if self.biblioteca_cerrada and self.tiempo_inicio_cerrada is not None:
                self.tiempo_biblioteca_cerrada_ac += self.tiempo_maximo - self.reloj

            return None

        if self.biblioteca_cerrada and self.tiempo_inicio_cerrada is not None:
            tiempo_transcurrido = evento.tiempo - self.reloj
            self.tiempo_biblioteca_cerrada_ac += tiempo_transcurrido

        self.reloj = evento.tiempo
        self.ultimo_reloj = self.reloj

        if evento.tipo == TipoEvento.LLEGADA_CLIENTE:
            self.procesar_llegada_cliente(evento)
        elif evento.tipo == TipoEvento.FIN_ATENCION:
            self.procesar_fin_atencion(evento)
        elif evento.tipo == TipoEvento.FIN_LECTURA:
            self.procesar_fin_lectura(evento)

        fila_datos = self._capturar_estado(evento)

        self.historial_filas.append(fila_datos)
        self.numero_fila += 1
        self.iteracion_actual += 1

        return fila_datos

    def ejecutar_completa(self):
        self.iniciar()
        while True:
            resultado = self.ejecutar_paso()
            if resultado is None:
                break
        return self.historial_filas

    def _capturar_estado(self, evento: Evento) -> dict:
        proximos = sorted(self.eventos, key=lambda e: e.tiempo)[:3]
        cliente_actual = evento.datos.get('cliente')

        evento_str = evento.tipo.value
        if cliente_actual and evento.tipo != TipoEvento.INICIALIZACION:
            evento_str = f"{evento.tipo.value} C{cliente_actual.id}"


        rnd_busqueda = cliente_actual.rnd_tiempo_busqueda if cliente_actual and cliente_actual.objetivo == TipoObjetivo.PEDIR_LIBRO and cliente_actual.rnd_tiempo_busqueda > 0 else ''
        tiempo_busqueda = cliente_actual.tiempo_busqueda if cliente_actual and cliente_actual.objetivo == TipoObjetivo.PEDIR_LIBRO and cliente_actual.tiempo_busqueda > 0 else ''
        rnd_devolucion = cliente_actual.rnd_tiempo_devolucion if cliente_actual and cliente_actual.objetivo == TipoObjetivo.DEVOLVER_LIBRO and cliente_actual.rnd_tiempo_devolucion > 0 else ''
        tiempo_devolucion = cliente_actual.tiempo_devolucion if cliente_actual and cliente_actual.objetivo == TipoObjetivo.DEVOLVER_LIBRO and cliente_actual.tiempo_devolucion > 0 else ''
        rnd_consulta = cliente_actual.rnd_tiempo_consulta if cliente_actual and cliente_actual.objetivo == TipoObjetivo.CONSULTAR and cliente_actual.rnd_tiempo_consulta > 0 else ''
        tiempo_consulta = cliente_actual.tiempo_consulta if cliente_actual and cliente_actual.objetivo == TipoObjetivo.CONSULTAR and cliente_actual.tiempo_consulta > 0 else ''

        proxima_llegada = next((e.tiempo for e in proximos if e.tipo == TipoEvento.LLEGADA_CLIENTE), '')

        objetivo_val = cliente_actual.objetivo.value if cliente_actual and cliente_actual.estado != "RECHAZADO" else ('RECHAZADO' if cliente_actual and cliente_actual.estado == "RECHAZADO" else '')
        rnd_obj_val = cliente_actual.rnd_objetivo if cliente_actual and cliente_actual.estado != "RECHAZADO" else ''

        clientes_copiados = [copy.deepcopy(c) for c in self.clientes_activos]

        return {
            'n': self.numero_fila,
            'evento': evento_str,
            'reloj': self.reloj,
            'tiempo_entre_llegadas': self.tiempo_entre_llegadas if evento.tipo == TipoEvento.LLEGADA_CLIENTE or evento.tipo == TipoEvento.INICIALIZACION else '',
            'proxima_llegada': proxima_llegada if evento.tipo != TipoEvento.INICIALIZACION else self.tiempo_entre_llegadas,
            'rnd_llegada': '',
            'rnd_objetivo': rnd_obj_val,
            'objetivo': objetivo_val,
            'rnd_busqueda': rnd_busqueda, 'tiempo_busqueda': tiempo_busqueda,
            'fin_atencion_alq1': cliente_actual.fin_atencion_emp1 if cliente_actual and cliente_actual.objetivo == TipoObjetivo.PEDIR_LIBRO else '',
            'fin_atencion_alq2': cliente_actual.fin_atencion_emp2 if cliente_actual and cliente_actual.objetivo == TipoObjetivo.PEDIR_LIBRO else '',
            'rnd_decision': cliente_actual.rnd_decision if cliente_actual and cliente_actual.rnd_decision is not None else '',
            'se_retira': 'Sí' if cliente_actual and cliente_actual.se_retira else ('No' if cliente_actual and cliente_actual.rnd_decision is not None else ''),
            'rnd_paginas': cliente_actual.rnd_paginas if cliente_actual and cliente_actual.rnd_paginas else '',
            'paginas': cliente_actual.paginas_a_leer if cliente_actual and cliente_actual.paginas_a_leer > 0 else '',
            'tiempo_lectura': cliente_actual.tiempo_lectura if cliente_actual and cliente_actual.tiempo_lectura > 0 else '',
            'rnd_devolucion': rnd_devolucion, 'tiempo_devolucion': tiempo_devolucion,
            'fin_atencion_dev1': cliente_actual.fin_atencion_emp1 if cliente_actual and cliente_actual.objetivo == TipoObjetivo.DEVOLVER_LIBRO else '',
            'fin_atencion_dev2': cliente_actual.fin_atencion_emp2 if cliente_actual and cliente_actual.objetivo == TipoObjetivo.DEVOLVER_LIBRO else '',
            'rnd_consulta': rnd_consulta, 'tiempo_consulta': tiempo_consulta,
            'fin_atencion_cons': (cliente_actual.fin_atencion_emp1 or cliente_actual.fin_atencion_emp2) if cliente_actual and cliente_actual.objetivo == TipoObjetivo.CONSULTAR else '',
            'empleado1_estado': self.empleados[0].estado.value,
            'empleado1_ac_atencion': self.empleados[0].tiempo_acumulado_atencion,
            'empleado1_ac_ocioso': self.empleados[0].tiempo_acumulado_ocioso,
            'empleado2_estado': self.empleados[1].estado.value,
            'empleado2_ac_atencion': self.empleados[1].tiempo_acumulado_atencion,
            'empleado2_ac_ocioso': self.empleados[1].tiempo_acumulado_ocioso,
            'estado_biblioteca': 'Cerrada' if self.biblioteca_cerrada else 'Abierta',
            'personas_leyendo': self.contar_personas_leyendo(),
            'cola': len(self.cola_espera),
            'ac_tiempo_permanencia': self.ac_tiempo_permanencia,
            'ac_clientes_leyendo': self.total_clientes_leyendo,
            'clientes': clientes_copiados
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
                if contador % 10 == 0:
                    self.progress.emit(contador)

            self.finished.emit(self.simulacion.historial_filas)

        except Exception as e:
            self.error.emit(str(e))


# ==================== VENTANA DE CONFIGURACIÓN (MÍNIMA) ====================

class ConfiguracionWindow(QDialog):
    """Ventana para configurar los parámetros esenciales de la simulación."""
    CONFIG_FILE = "config_simulacion.json"

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("⚙️ Configuración de Parámetros Esenciales")
        self.setGeometry(100, 100, 700, 500)
        self.parametros = None

        main_layout = QVBoxLayout()
        self.setLayout(main_layout)

        tabs = QTabWidget()
        tabs.addTab(self.crear_tab_parametros(), "Parámetros Esenciales")
        tabs.addTab(self.crear_tab_fijos(), "Parámetros Fijos (Enunciado)")
        main_layout.addWidget(tabs)

        btn_iniciar = QPushButton("🚀 Iniciar Simulación con Parámetros")
        btn_iniciar.setStyleSheet("background-color: #2196F3; color: white; padding: 10px; font-size: 12pt; font-weight: bold; border-radius: 5px;")
        btn_iniciar.clicked.connect(self.aceptar)
        main_layout.addWidget(btn_iniciar)

        # Cargar configuración guardada
        self.cargar_configuracion()

    def crear_tab_parametros(self):
        widget = QWidget()
        layout = QFormLayout()
        widget.setLayout(layout)
        
        # 1. Frecuencia de Llegada (Default: 4 min)
        group_llegadas = QGroupBox("1. Frecuencia de Llegada")
        group_llegadas.setLayout(QFormLayout())
        self.s_tiempo_entre_llegadas = QDoubleSpinBox(value=4.0, decimals=1, minimum=0.1, maximum=30.0, singleStep=0.5)
        group_llegadas.layout().addRow("Tiempo Entre Llegadas (min):", self.s_tiempo_entre_llegadas)
        layout.addRow(group_llegadas)

        # 2. Probabilidades de Objetivo (Default: 45%, 45%, 10%)
        group_objetivo = QGroupBox("2. Probabilidades de Objetivo (Suma debe ser 1.0)")
        group_objetivo.setLayout(QFormLayout())
        self.s_prob_pedir = QDoubleSpinBox(value=0.45, decimals=2, minimum=0.0, maximum=1.0, singleStep=0.01)
        self.s_prob_devolver = QDoubleSpinBox(value=0.45, decimals=2, minimum=0.0, maximum=1.0, singleStep=0.01)
        self.s_prob_consultar = QDoubleSpinBox(value=0.10, decimals=2, minimum=0.0, maximum=1.0, singleStep=0.01)
        group_objetivo.layout().addRow("P(Pedir libro):", self.s_prob_pedir)
        group_objetivo.layout().addRow("P(Devolver libro):", self.s_prob_devolver)
        group_objetivo.layout().addRow("P(Consultar):", self.s_prob_consultar)
        layout.addRow(group_objetivo)
        
        # 3. Tiempos de Consulta (Default: 2 y 5 min)
        group_consulta = QGroupBox("3. Tiempo de Consulta (U[Min, Max])")
        group_consulta.setLayout(QFormLayout())
        self.s_cons_min = QDoubleSpinBox(value=2.0, decimals=1, minimum=0.1, maximum=10.0, singleStep=0.1)
        self.s_cons_max = QDoubleSpinBox(value=5.0, decimals=1, minimum=0.1, maximum=10.0, singleStep=0.1)
        group_consulta.layout().addRow("Consulta Min (min):", self.s_cons_min)
        group_consulta.layout().addRow("Consulta Max (min):", self.s_cons_max)
        layout.addRow(group_consulta)

        # 4. Decisión de Lectura (Default: 60% retira)
        group_lectura_decision = QGroupBox("4. Decisión de Retirarse/Leer")
        group_lectura_decision.setLayout(QFormLayout())
        self.s_prob_retirarse = QDoubleSpinBox(value=0.60, decimals=2, minimum=0.0, maximum=1.0, singleStep=0.01)
        group_lectura_decision.layout().addRow("P(Se retira después de pedir):", self.s_prob_retirarse)
        layout.addRow(group_lectura_decision)

        # 5. Parámetros de Simulación
        group_simulacion = QGroupBox("5. Parámetros de Simulación")
        group_simulacion.setLayout(QFormLayout())
        self.s_tiempo_simulacion = QDoubleSpinBox(value=480.0, decimals=1, minimum=1.0, maximum=100000.0, singleStep=10.0)
        self.s_max_iteraciones = QSpinBox(value=100000, minimum=1, maximum=1000000, singleStep=1000)
        group_simulacion.layout().addRow("X - Tiempo de Simulación (min):", self.s_tiempo_simulacion)
        group_simulacion.layout().addRow("N - Cantidad máxima de iteraciones:", self.s_max_iteraciones)
        layout.addRow(group_simulacion)

        # 6. Constantes K de Integración Numérica
        group_k = QGroupBox("6. Constantes K de Integración Numérica (Método de Euler)")
        group_k.setLayout(QFormLayout())

        info_k = QLabel("dP/dt = K/5, donde K depende del rango de páginas:")
        info_k.setStyleSheet("color: #333; font-size: 9pt; font-style: italic;")
        group_k.layout().addRow(info_k)

        self.s_k_100_200 = QSpinBox(value=100, minimum=1, maximum=1000, singleStep=10)
        self.s_k_200_300 = QSpinBox(value=90, minimum=1, maximum=1000, singleStep=10)
        self.s_k_300_plus = QSpinBox(value=70, minimum=1, maximum=1000, singleStep=10)

        group_k.layout().addRow("K [100-200 páginas]:", self.s_k_100_200)
        group_k.layout().addRow("K [201-300 páginas]:", self.s_k_200_300)
        group_k.layout().addRow("K [301+ páginas]:", self.s_k_300_plus)
        layout.addRow(group_k)

        # 7. Rango de Visualización (Vector de Estado)
        group_visualizacion = QGroupBox("7. Rango de Visualización en Tabla (Filtro de Eventos)")
        group_visualizacion.setLayout(QFormLayout())
        self.s_hora_desde = QDoubleSpinBox(value=0.0, decimals=1, minimum=0.0, maximum=100000.0, singleStep=10.0)
        self.s_hora_hasta = QDoubleSpinBox(value=100000.0, decimals=1, minimum=0.0, maximum=100000.0, singleStep=10.0)

        info_label = QLabel("⚠️ Estos valores filtran qué eventos se MUESTRAN en la tabla.\n"
                           "Para ver TODOS los eventos, usa valores muy grandes (ej: j=100000)")
        info_label.setStyleSheet("color: #856404; font-size: 9pt; font-style: italic;")
        info_label.setWordWrap(True)

        group_visualizacion.layout().addRow(info_label)
        group_visualizacion.layout().addRow("i - Desde hora (min):", self.s_hora_desde)
        group_visualizacion.layout().addRow("j - Hasta hora (min):", self.s_hora_hasta)
        layout.addRow(group_visualizacion)

        return widget
    
    def crear_tab_fijos(self):
        widget = QWidget()
        layout = QFormLayout()
        widget.setLayout(layout)

        group_fijos = QGroupBox("Parámetros Fijos (No Parametrizables por Requerimiento)")
        group_fijos.setStyleSheet("QGroupBox { font-style: italic; } QLabel { color: #6c757d; }")

        fixed_layout = QFormLayout()
        fixed_layout.addRow("T. Búsqueda (Media EXP):", QLabel("6.0 min"))
        fixed_layout.addRow("T. Devolución (U[Min, Max]):", QLabel("U[1.5, 2.5] min"))
        fixed_layout.addRow("Páginas (U[Min, Max]):", QLabel("U[100, 350]"))
        fixed_layout.addRow("Capacidad Máxima:", QLabel("20 personas leyendo"))
        fixed_layout.addRow("Paso de integración (h):", QLabel("0.1 = 10 min"))

        group_fijos.setLayout(fixed_layout)
        layout.addRow(group_fijos)
        return widget


    def aceptar(self):
        """Valida y guarda los parámetros antes de cerrar."""
        p_pedir = self.s_prob_pedir.value()
        p_devolver = self.s_prob_devolver.value()
        p_consultar = self.s_prob_consultar.value()

        # Validación de la suma de probabilidades
        if not math.isclose(p_pedir + p_devolver + p_consultar, 1.0, abs_tol=0.001):
            QMessageBox.critical(self, "Error de Probabilidad", 
                                "La suma de las probabilidades de objetivo (Pedir, Devolver, Consultar) debe ser 1.0.")
            return
        
        # Validación de rangos
        if self.s_cons_min.value() > self.s_cons_max.value():
            QMessageBox.critical(self, "Error de Rango",
                                "El valor mínimo de consulta no puede ser mayor que el máximo.")
            return

        # Validación de rango de visualización
        if self.s_hora_desde.value() > self.s_hora_hasta.value():
            QMessageBox.critical(self, "Error de Rango",
                                "La hora 'Desde' (i) no puede ser mayor que la hora 'Hasta' (j).")
            return

        # Validación de tiempo de simulación
        if self.s_tiempo_simulacion.value() < self.s_hora_hasta.value():
            QMessageBox.warning(self, "Advertencia",
                               "El tiempo de simulación (X) es menor que el rango 'Hasta' (j).\n"
                               "Es posible que no se generen todos los eventos esperados.")

        self.parametros = {
            'tiempo_entre_llegadas': self.s_tiempo_entre_llegadas.value(),
            'prob_pedir_libro': p_pedir,
            'prob_devolver_libro': p_devolver,
            'prob_consultar': p_consultar,
            'tiempo_consulta_min': self.s_cons_min.value(),
            'tiempo_consulta_max': self.s_cons_max.value(),
            'prob_retirarse': self.s_prob_retirarse.value(),
            'tiempo_simulacion': self.s_tiempo_simulacion.value(),
            'max_iteraciones': self.s_max_iteraciones.value(),
            'k_100_200': self.s_k_100_200.value(),
            'k_200_300': self.s_k_200_300.value(),
            'k_300_plus': self.s_k_300_plus.value(),
            'hora_desde': self.s_hora_desde.value(),
            'hora_hasta': self.s_hora_hasta.value(),
        }

        # Guardar configuración
        self.guardar_configuracion()
        self.accept()

    def guardar_configuracion(self):
        """Guarda la configuración actual en un archivo JSON"""
        try:
            config = {
                'tiempo_entre_llegadas': self.s_tiempo_entre_llegadas.value(),
                'prob_pedir_libro': self.s_prob_pedir.value(),
                'prob_devolver_libro': self.s_prob_devolver.value(),
                'prob_consultar': self.s_prob_consultar.value(),
                'tiempo_consulta_min': self.s_cons_min.value(),
                'tiempo_consulta_max': self.s_cons_max.value(),
                'prob_retirarse': self.s_prob_retirarse.value(),
                'tiempo_simulacion': self.s_tiempo_simulacion.value(),
                'max_iteraciones': self.s_max_iteraciones.value(),
                'k_100_200': self.s_k_100_200.value(),
                'k_200_300': self.s_k_200_300.value(),
                'k_300_plus': self.s_k_300_plus.value(),
                'hora_desde': self.s_hora_desde.value(),
                'hora_hasta': self.s_hora_hasta.value(),
            }
            with open(self.CONFIG_FILE, 'w') as f:
                json.dump(config, f, indent=2)
        except Exception as e:
            print(f"Error al guardar configuración: {e}")

    def cargar_configuracion(self):
        """Carga la configuración guardada desde el archivo JSON"""
        if not os.path.exists(self.CONFIG_FILE):
            return

        try:
            with open(self.CONFIG_FILE, 'r') as f:
                config = json.load(f)

            # Aplicar valores guardados
            self.s_tiempo_entre_llegadas.setValue(config.get('tiempo_entre_llegadas', 4.0))
            self.s_prob_pedir.setValue(config.get('prob_pedir_libro', 0.45))
            self.s_prob_devolver.setValue(config.get('prob_devolver_libro', 0.45))
            self.s_prob_consultar.setValue(config.get('prob_consultar', 0.10))
            self.s_cons_min.setValue(config.get('tiempo_consulta_min', 2.0))
            self.s_cons_max.setValue(config.get('tiempo_consulta_max', 5.0))
            self.s_prob_retirarse.setValue(config.get('prob_retirarse', 0.60))
            self.s_tiempo_simulacion.setValue(config.get('tiempo_simulacion', 480.0))
            self.s_max_iteraciones.setValue(config.get('max_iteraciones', 100000))
            self.s_k_100_200.setValue(config.get('k_100_200', 100))
            self.s_k_200_300.setValue(config.get('k_200_300', 90))
            self.s_k_300_plus.setValue(config.get('k_300_plus', 70))
            self.s_hora_desde.setValue(config.get('hora_desde', 0.0))
            self.s_hora_hasta.setValue(config.get('hora_hasta', 100000.0))
        except Exception as e:
            print(f"Error al cargar configuración: {e}")


# ==================== INTERFAZ GRÁFICA PRINCIPAL ====================

class MainWindow(QMainWindow):
    """Ventana principal de la aplicación"""

    def __init__(self):
        super().__init__()
        self.simulacion: Optional[Simulacion] = None
        self.historial_filas = []
        self.thread = None
        self.parametros_simulacion = None 

        self.init_ui()
        # Inicia la ventana de configuración inmediatamente al arrancar
        self.mostrar_configuracion() 

    def init_ui(self):
        """Inicializa la interfaz"""
        self.setWindowTitle("Simulación de Biblioteca - Sistema de Eventos Discretos con Método de Euler")
        self.setGeometry(50, 50, 1800, 950)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout()
        central_widget.setLayout(main_layout)

        title = QLabel("🏛️ SIMULACIÓN DE BIBLIOTECA - EVENTOS DISCRETOS + INTEGRACIÓN DE EULER")
        title.setFont(QFont("Arial", 16, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("background-color: #4472C4; color: white; padding: 15px; border-radius: 5px;")
        main_layout.addWidget(title)

        info_euler = QLabel("📐 Integración Numérica: dP/dt = K/5 (Método de Euler para calcular tiempo de lectura)")
        info_euler.setFont(QFont("Arial", 10))
        info_euler.setAlignment(Qt.AlignCenter)
        info_euler.setStyleSheet("background-color: #fff3cd; padding: 8px; border-radius: 3px; color: #856404;")
        main_layout.addWidget(info_euler)

        self.control_panel = self.crear_panel_control()
        main_layout.addWidget(self.control_panel)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)

        self.tabla = QTableWidget()
        self.tabla.setAlternatingRowColors(True)
        self.tabla.setStyleSheet("""
            QTableWidget { gridline-color: #d0d0d0; font-size: 9pt; background-color: white; }
            QHeaderView::section { background-color: #4472C4; color: white; font-weight: bold; padding: 6px; border: 1px solid #2a52a4; }
            QTableWidget::item:alternate { background-color: #f9f9f9; }
        """)

        scroll.setWidget(self.tabla)
        main_layout.addWidget(scroll)

        self.progress_bar = QProgressBar()
        self.progress_bar.setStyleSheet("QProgressBar { border: 2px solid #ccc; border-radius: 5px; text-align: center; } QProgressBar::chunk { background-color: #4CAF50; }")
        self.progress_bar.setVisible(False)
        main_layout.addWidget(self.progress_bar)

        status_layout = QHBoxLayout()
        self.lbl_status = QLabel("⏰ Esperando configuración...")
        self.lbl_status.setFont(QFont("Arial", 10, QFont.Bold))
        status_layout.addWidget(self.lbl_status)
        status_layout.addStretch()
        main_layout.addLayout(status_layout)
        
        self.set_botones_habilitados(False)

    def set_botones_habilitados(self, enabled: bool):
        self.btn_ejecutar.setEnabled(enabled)
        self.btn_exportar.setEnabled(False)
        self.btn_reiniciar.setEnabled(True)

    def crear_panel_control(self) -> QGroupBox:
        """Crea el panel de control"""
        group = QGroupBox("🎮 Controles de Simulación")
        group.setFont(QFont("Arial", 11, QFont.Bold))
        layout = QHBoxLayout()

        self.btn_ejecutar = QPushButton("▶ Ejecutar Simulación Completa")
        self.btn_ejecutar.setStyleSheet("background-color: #4CAF50; color: white; padding: 12px 24px; font-size: 13pt; font-weight: bold; border-radius: 5px;")
        self.btn_ejecutar.clicked.connect(self.ejecutar_simulacion)
        layout.addWidget(self.btn_ejecutar)

        self.btn_exportar = QPushButton("📊 Exportar a Excel")
        self.btn_exportar.setStyleSheet("background-color: #2196F3; color: white; padding: 12px 24px; font-size: 13pt; font-weight: bold; border-radius: 5px;")
        self.btn_exportar.clicked.connect(self.exportar_excel)
        layout.addWidget(self.btn_exportar)

        self.btn_reiniciar = QPushButton("🔄 Reconfigurar / Nueva Simulación")
        self.btn_reiniciar.setStyleSheet("background-color: #F44336; color: white; padding: 12px 24px; font-size: 13pt; font-weight: bold; border-radius: 5px;")
        self.btn_reiniciar.clicked.connect(self.mostrar_configuracion)
        layout.addWidget(self.btn_reiniciar)

        layout.addStretch()
        group.setLayout(layout)
        return group
    
    def mostrar_configuracion(self):
        """Abre la ventana de diálogo para configurar parámetros."""
        config_dialog = ConfiguracionWindow(self)
        if config_dialog.exec_() == QDialog.Accepted:
            self.parametros_simulacion = config_dialog.parametros
            self.lbl_status.setText(f"✅ Configuración cargada. T.sim: {self.parametros_simulacion['tiempo_simulacion']:.0f} min | "
                                  f"Max iter: {self.parametros_simulacion['max_iteraciones']} | "
                                  f"Rango: [{self.parametros_simulacion['hora_desde']:.0f}, {self.parametros_simulacion['hora_hasta']:.0f}]")
            self.set_botones_habilitados(True)
            self.reiniciar_tabla()
        else:
            self.lbl_status.setText("❌ Configuración cancelada. Presione 'Reconfigurar / Nueva Simulación'.")
            self.set_botones_habilitados(False)


    def ejecutar_simulacion(self):
        """Ejecuta la simulación completa"""
        if not self.parametros_simulacion:
            QMessageBox.warning(self, "Advertencia", "Debe configurar y cargar los parámetros primero.")
            self.mostrar_configuracion()
            return
            
        self.reiniciar_tabla()

        self.btn_ejecutar.setEnabled(False)
        self.btn_exportar.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)
        self.lbl_status.setText("⏳ Ejecutando simulación completa...")

        # Inyectamos los parámetros en la simulación
        self.simulacion = Simulacion(self.parametros_simulacion)

        self.thread = SimulacionThread(self.simulacion)
        self.thread.finished.connect(self.simulacion_completada)
        self.thread.error.connect(self.simulacion_error)
        self.thread.progress.connect(self.actualizar_progreso)
        self.thread.start()

    def actualizar_progreso(self, eventos):
        """Actualiza el progreso"""
        self.lbl_status.setText(f"⏳ Procesando... {eventos} eventos (sin contar Inicialización)")

    def simulacion_completada(self, historial_filas):
        """Callback cuando termina la simulación"""
        self.historial_filas = historial_filas
        self.progress_bar.setVisible(False)
        self.set_botones_habilitados(True)
        self.btn_ejecutar.setEnabled(True)
        self.btn_exportar.setEnabled(True)

        self.poblar_tabla()
        
        # CÁLCULO DE MÉTRICAS FINALES
        total_clientes = self.simulacion.total_clientes_generados
        promedio_permanencia = (self.simulacion.ac_tiempo_permanencia / 
                                self.simulacion.total_clientes_atendidos) if self.simulacion.total_clientes_atendidos > 0 else 0
        
        porcentaje_rechazados = (self.simulacion.total_clientes_rechazados / 
                                total_clientes) * 100 if total_clientes > 0 else 0
        
        porcentaje_tiempo_cerrada = (self.simulacion.tiempo_biblioteca_cerrada_ac / 
                                     self.simulacion.tiempo_maximo) * 100 if self.simulacion.tiempo_maximo > 0 else 0

        # Determinar razón de finalización
        razon_finalizacion = ""
        if self.simulacion.iteracion_actual >= self.simulacion.max_iteraciones:
            razon_finalizacion = f"⚠️ Límite de iteraciones alcanzado: {self.simulacion.max_iteraciones}\n"
        elif self.simulacion.reloj >= self.simulacion.tiempo_maximo:
            razon_finalizacion = f"✅ Tiempo de simulación completado: {self.simulacion.tiempo_maximo:.2f} min\n"

        self.lbl_status.setText(f"✅ Simulación completa: {len(historial_filas)} eventos | "
                                f"T. Sim: {self.simulacion.reloj:.2f} min | "
                                f"Iter: {self.simulacion.iteracion_actual}")

        QMessageBox.information(self, "Simulación Completa",
                                f"✅ Simulación finalizada exitosamente\n\n"
                                f"{razon_finalizacion}"
                                f"📊 Eventos procesados: {len(historial_filas)}\n"
                                f"🔄 Iteraciones ejecutadas: {self.simulacion.iteracion_actual}\n"
                                f"⏱️ Tiempo simulado: {self.simulacion.reloj:.2f} min\n"
                                f"--- MÉTRICAS ---\n"
                                f"👥 Clientes generados: {total_clientes}\n"
                                f"❌ Clientes rechazados: {self.simulacion.total_clientes_rechazados} ({porcentaje_rechazados:.2f}%)\n"
                                f"⏳ Promedio de permanencia: {promedio_permanencia:.2f} min\n"
                                f"🚪 % Tiempo cerrada por capacidad: {porcentaje_tiempo_cerrada:.2f}%")

    def simulacion_error(self, error_msg):
        """Callback cuando hay error"""
        self.progress_bar.setVisible(False)
        self.set_botones_habilitados(True)
        self.btn_ejecutar.setEnabled(True)
        self.lbl_status.setText("❌ Error en la simulación")
        QMessageBox.critical(self, "Error", f"Error en la simulación:\n\n{error_msg}")

    def reiniciar_tabla(self):
        """Limpia los datos y la tabla de la simulación anterior."""
        self.simulacion = None
        self.historial_filas = []
        self.tabla.setRowCount(0)
        self.tabla.setColumnCount(0)
        self.btn_exportar.setEnabled(False)
        self.lbl_status.setText("⏰ Esperando...")


    def poblar_tabla(self):
        if not self.historial_filas: return

        # Filtrar filas según el rango de tiempo especificado
        hora_desde = self.simulacion.hora_desde if self.simulacion else 0
        hora_hasta = self.simulacion.hora_hasta if self.simulacion else float('inf')

        filas_filtradas = [fila for fila in self.historial_filas
                          if hora_desde <= fila['reloj'] <= hora_hasta]

        if not filas_filtradas:
            QMessageBox.warning(self, "Sin datos",
                              f"No hay eventos en el rango de tiempo especificado:\n"
                              f"Desde {hora_desde} hasta {hora_hasta} minutos.\n\n"
                              f"Prueba ajustar los parámetros 'i' y 'j' en la configuración.")
            return

        # OPTIMIZACIÓN: Limitar cantidad de filas a mostrar para evitar congelamientos
        MAX_FILAS_TABLA = 5000  # Reducido para mejor rendimiento
        total_filas = len(filas_filtradas)

        if total_filas > MAX_FILAS_TABLA:
            # Mostrar advertencia
            respuesta = QMessageBox.question(
                self, "Demasiados eventos",
                f"Se encontraron {total_filas:,} eventos en el rango especificado.\n\n"
                f"Mostrar todos los eventos puede hacer que la aplicación se congele.\n"
                f"Se mostrarán solo los primeros {MAX_FILAS_TABLA:,} eventos.\n\n"
                f"💡 Para ver todos los datos, use la exportación a Excel.\n\n"
                f"¿Desea continuar?\n\n"
                f"Sugerencia: Ajuste los parámetros 'i' y 'j' para un rango más pequeño.",
                QMessageBox.Yes | QMessageBox.No
            )
            if respuesta == QMessageBox.No:
                return

            filas_filtradas = filas_filtradas[:MAX_FILAS_TABLA]

        todos_clientes_ids = set()
        for fila in filas_filtradas:
            for cliente in fila['clientes']:
                if cliente.estado != "RECHAZADO":
                    todos_clientes_ids.add(cliente.id)
        clientes_ordenados = sorted(list(todos_clientes_ids))

        columnas_fijas = [
            "n°", "Evento", "Reloj",
            "T.entre llegadas", "Próxima llegada",
            "RND obj", "Objetivo",
            "RND búsq", "T.búsqueda (EXP)", "fin_búsq_emp1", "fin_búsq_emp2",
            "RND decisión", "Se retira?", "RND pág", "Páginas", "T.lectura (Euler)",
            "RND dev", "T.devolución", "fin_dev_emp1", "fin_dev_emp2",
            "RND cons", "T.consulta", "fin_cons",
            "Emp1", "AC at1", "AC oc1", "Emp2", "AC at2", "AC oc2",
            "Estado Bib", "Leyendo", "Cola", "AC perm", "AC leyendo"
        ]

        columnas_clientes = []
        for cid in clientes_ordenados:
            columnas_clientes.extend([
                f"C{cid} Estado", f"C{cid} Hora", f"C{cid} T.at", f"C{cid} Fin lect"
            ])

        todas_columnas = columnas_fijas + columnas_clientes

        # OPTIMIZACIÓN: Deshabilitar actualizaciones durante el llenado
        self.tabla.setUpdatesEnabled(False)
        self.tabla.setSortingEnabled(False)

        try:
            self.tabla.setColumnCount(len(todas_columnas))
            self.tabla.setHorizontalHeaderLabels(todas_columnas)
            self.tabla.setRowCount(len(filas_filtradas))

            header = self.tabla.horizontalHeader()
            header.setDefaultSectionSize(85)
            header.setSectionResizeMode(QHeaderView.Interactive)

            # Mostrar progreso para tablas grandes
            if len(filas_filtradas) > 1000:
                self.lbl_status.setText(f"⏳ Poblando tabla con {len(filas_filtradas):,} filas...")
                QApplication.processEvents()

            for row, fila in enumerate(filas_filtradas):
                self.agregar_fila(row, fila, clientes_ordenados)

                # Actualizar progreso cada 500 filas
                if len(filas_filtradas) > 1000 and row % 500 == 0:
                    progreso = int((row / len(filas_filtradas)) * 100)
                    self.lbl_status.setText(f"⏳ Poblando tabla: {progreso}% ({row:,}/{len(filas_filtradas):,})")
                    QApplication.processEvents()

        finally:
            # Rehabilitar actualizaciones
            self.tabla.setUpdatesEnabled(True)
            self.tabla.setSortingEnabled(True)

            if total_filas > MAX_FILAS_TABLA:
                self.lbl_status.setText(f"⚠️ Mostrando {len(filas_filtradas):,} de {total_filas:,} eventos")
            else:
                self.lbl_status.setText(f"✅ Tabla poblada: {len(filas_filtradas):,} eventos")

    def agregar_fila(self, row: int, datos: dict, clientes_ordenados: List[int]):
        # OPTIMIZACIÓN: Pre-cache de objetos de color comunes para evitar recrearlos
        if not hasattr(self, '_cached_colors'):
            self._cached_colors = {
                'row_even': QColor(249, 249, 249),
                'row_odd': QColor(255, 255, 255),
                'init': QColor(230, 230, 250),
                'llegada': QColor(200, 230, 201),
                'fin_atencion': QColor(255, 224, 178),
                'fin_lectura': QColor(187, 222, 251),
                'rechazado': QColor(255, 199, 206),
                'red': QColor(255, 0, 0)
            }
            self._cached_font_bold = QFont("Arial", 9, QFont.Bold)

        def fmt(val):
            if val == '' or val is None: return ''
            if isinstance(val, float): return f"{val:.2f}"
            return str(val)

        es_llegada = datos['evento'].startswith(TipoEvento.LLEGADA_CLIENTE.value)
        es_rechazado = datos['objetivo'] == 'RECHAZADO'
        mostrar_rnd_obj = es_llegada and not es_rechazado

        # Preparar valores
        valores = [
            datos['n'], datos['evento'], fmt(datos['reloj']), fmt(datos['tiempo_entre_llegadas']), fmt(datos['proxima_llegada']),
            fmt(datos['rnd_objetivo']) if mostrar_rnd_obj else '', datos['objetivo'] if es_llegada else '',
            fmt(datos['rnd_busqueda']), fmt(datos['tiempo_busqueda']), fmt(datos['fin_atencion_alq1']), fmt(datos['fin_atencion_alq2']),
            fmt(datos['rnd_decision']), datos['se_retira'], fmt(datos['rnd_paginas']), datos['paginas'], fmt(datos['tiempo_lectura']),
            fmt(datos['rnd_devolucion']), fmt(datos['tiempo_devolucion']), fmt(datos['fin_atencion_dev1']), fmt(datos['fin_atencion_dev2']),
            fmt(datos['rnd_consulta']), fmt(datos['tiempo_consulta']), fmt(datos['fin_atencion_cons']),
            datos['empleado1_estado'], fmt(datos['empleado1_ac_atencion']), fmt(datos['empleado1_ac_ocioso']),
            datos['empleado2_estado'], fmt(datos['empleado2_ac_atencion']), fmt(datos['empleado2_ac_ocioso']),
            datos['estado_biblioteca'], str(datos.get('personas_leyendo', 0)),
            str(datos['cola']), fmt(datos['ac_tiempo_permanencia']), str(datos['ac_clientes_leyendo'])
        ]

        col_offset = len(valores)
        clientes_dict = {c.id: c for c in datos['clientes']}

        for cid in clientes_ordenados:
            if cid in clientes_dict:
                c = clientes_dict[cid]
                tiempo_at = next((t for t in [c.tiempo_busqueda, c.tiempo_devolucion, c.tiempo_consulta] if t > 0), '')
                valores.extend([c.estado, fmt(c.hora_llegada), fmt(tiempo_at) if tiempo_at else '', fmt(c.fin_lectura) if c.fin_lectura else ''])
            else:
                valores.extend(['', '', '', ''])
            col_offset += 4

        # OPTIMIZACIÓN CRÍTICA: Minimizar el styling por celda
        # Solo coloreamos el evento y los rechazados, sin revisar tiempos próximos
        color_fondo_base = self._cached_colors['row_even'] if row % 2 == 0 else self._cached_colors['row_odd']

        for col, valor in enumerate(valores):
            item = QTableWidgetItem(str(valor))
            item.setTextAlignment(Qt.AlignCenter)

            color_fondo = color_fondo_base

            # Solo aplicar colores especiales a columnas críticas
            if col == 1:  # Columna evento
                valor_lower = str(valor).lower()
                if 'inicializacion' in valor_lower:
                    color_fondo = self._cached_colors['init']
                elif 'llegada' in valor_lower:
                    color_fondo = self._cached_colors['llegada']
                elif 'fin_atencion' in valor_lower:
                    color_fondo = self._cached_colors['fin_atencion']
                elif 'fin_lectura' in valor_lower:
                    color_fondo = self._cached_colors['fin_lectura']
            elif es_llegada and es_rechazado and col == 6:  # Columna objetivo cuando es rechazado
                color_fondo = self._cached_colors['rechazado']

            item.setBackground(color_fondo)
            self.tabla.setItem(row, col, item)

    def exportar_excel(self):
        if not OPENPYXL_DISPONIBLE:
            QMessageBox.warning(self, "Advertencia", "La librería 'openpyxl' no está instalada.\nInstálala con: pip install openpyxl")
            return
        if not self.historial_filas:
            QMessageBox.warning(self, "Advertencia", "No hay datos para exportar")
            return

        filename, _ = QFileDialog.getSaveFileName(self, "Guardar archivo Excel", "simulacion_biblioteca.xlsx", "Excel Files (*.xlsx)")
        if not filename: return

        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Simulación"

            header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
            header_font = ExcelFont(bold=True, color="FFFFFF")
            center_align = Alignment(horizontal="center", vertical="center")

            # Exportar tabla principal de simulación
            for col in range(self.tabla.columnCount()):
                cell = ws.cell(row=1, column=col + 1)
                cell.value = self.tabla.horizontalHeaderItem(col).text()
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = center_align

            for row in range(self.tabla.rowCount()):
                for col in range(self.tabla.columnCount()):
                    item = self.tabla.item(row, col)
                    cell = ws.cell(row=row + 2, column=col + 1)
                    if item:
                        cell.value = item.text()
                        cell.alignment = center_align

            # Crear hoja de integración numérica si hay datos
            if self.simulacion and self.simulacion.historial_integraciones:
                ws_integracion = wb.create_sheet(title="Integración Numérica")

                header_integracion = ["Cliente ID", "Páginas Objetivo", "K", "h", "Paso", "t (min)", "P(t) (páginas)", "dP/dt", "Δt", "ΔP"]
                for col, header_text in enumerate(header_integracion, 1):
                    cell = ws_integracion.cell(row=1, column=col)
                    cell.value = header_text
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = center_align

                row_idx = 2
                for integracion in self.simulacion.historial_integraciones:
                    cliente_id = integracion['cliente_id']
                    paginas_obj = integracion['paginas_objetivo']
                    K = integracion['K']
                    h = integracion['h']

                    for paso_num, paso in enumerate(integracion['pasos'], 1):
                        ws_integracion.cell(row=row_idx, column=1).value = cliente_id
                        ws_integracion.cell(row=row_idx, column=2).value = paginas_obj
                        ws_integracion.cell(row=row_idx, column=3).value = K
                        ws_integracion.cell(row=row_idx, column=4).value = h
                        ws_integracion.cell(row=row_idx, column=5).value = paso_num
                        ws_integracion.cell(row=row_idx, column=6).value = round(paso['t'], 4)
                        ws_integracion.cell(row=row_idx, column=7).value = round(paso['p'], 4)
                        ws_integracion.cell(row=row_idx, column=8).value = round(paso['derivada'], 4)
                        ws_integracion.cell(row=row_idx, column=9).value = h
                        ws_integracion.cell(row=row_idx, column=10).value = round(h * paso['derivada'], 4)

                        for col in range(1, 11):
                            ws_integracion.cell(row=row_idx, column=col).alignment = center_align

                        row_idx += 1

                    # Fila en blanco entre clientes
                    row_idx += 1

            wb.save(filename)
            msg = f"✅ Archivo exportado correctamente:\n{filename}"
            if self.simulacion and self.simulacion.historial_integraciones:
                msg += f"\n\n📊 Se exportaron {len(self.simulacion.historial_integraciones)} integraciones numéricas"
            QMessageBox.information(self, "Éxito", msg)

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al exportar:\n\n{str(e)}")

    def reiniciar(self):
        self.mostrar_configuracion()


def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()