"""
Sistema de eventos de la simulación
"""
from enum import Enum

class TipoEvento(Enum):
    """Tipos de eventos en la simulación"""
    LLEGADA_PERSONA = "llegada_cliente"
    FIN_ATENCION = "fin_atencion_cliente"
    FIN_LECTURA = "fin_lectura"
    FIN_SIMULACION = "Fin Simulación"


class Evento:
    """Representa un evento en la simulación"""

    def __init__(self, tipo, tiempo, datos=None):
        """
        tipo: TipoEvento
        tiempo: momento en que ocurre el evento (en minutos)
        datos: diccionario con información adicional del evento
        """
        self.tipo = tipo
        self.tiempo = tiempo
        self.datos = datos if datos is not None else {}

    def __lt__(self, otro):
        """Permite ordenar eventos por tiempo"""
        return self.tiempo < otro.tiempo

    def __repr__(self):
        return f"Evento({self.tipo.value}, t={self.tiempo:.2f})"


class ListaEventos:
    """Administra la lista de eventos futuros (FEL - Future Event List)"""

    def __init__(self):
        self.eventos = []

    def agregar_evento(self, evento):
        """Agrega un evento y mantiene la lista ordenada por tiempo"""
        self.eventos.append(evento)
        self.eventos.sort()

    def proximo_evento(self):
        """Retorna y elimina el próximo evento a procesar"""
        if self.eventos:
            return self.eventos.pop(0)
        return None

    def cancelar_evento(self, tipo, condicion=None):
        """
        Cancela eventos de un tipo específico que cumplan una condición
        condicion: función que recibe un evento y retorna True si debe cancelarse
        """
        if condicion is None:
            self.eventos = [e for e in self.eventos if e.tipo != tipo]
        else:
            self.eventos = [e for e in self.eventos if not (e.tipo == tipo and condicion(e))]

    def obtener_proximos_eventos(self, n=5):
        """Retorna los próximos n eventos sin eliminarlos"""
        return self.eventos[:n]

    def tiene_eventos(self):
        """Verifica si hay eventos pendientes"""
        return len(self.eventos) > 0

    def __len__(self):
        return len(self.eventos)

    def __repr__(self):
        return f"ListaEventos({len(self.eventos)} eventos)"
