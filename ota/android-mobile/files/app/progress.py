import time
import threading
from datetime import datetime

tareas_progreso = {}
lock_progreso = threading.Lock()

def iniciar_tarea(tarea_id, descripcion, total_pasos=100):
    with lock_progreso:
        tareas_progreso[tarea_id] = {
            'id': tarea_id,
            'descripcion': descripcion,
            'total': total_pasos,
            'actual': 0,
            'porcentaje': 0,
            'estado': 'ejecutando',
            'mensaje': 'Iniciando...',
            'inicio': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'fin': None
        }

def actualizar_progreso(tarea_id, actual, mensaje=''):
    with lock_progreso:
        if tarea_id in tareas_progreso:
            tarea = tareas_progreso[tarea_id]
            tarea['actual'] = actual
            tarea['porcentaje'] = round((actual / tarea['total']) * 100, 1) if tarea['total'] > 0 else 0
            if mensaje:
                tarea['mensaje'] = mensaje

def completar_tarea(tarea_id, mensaje='Completado'):
    with lock_progreso:
        if tarea_id in tareas_progreso:
            tarea = tareas_progreso[tarea_id]
            tarea['estado'] = 'completado'
            tarea['porcentaje'] = 100
            tarea['actual'] = tarea['total']
            tarea['mensaje'] = mensaje
            tarea['fin'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

def error_tarea(tarea_id, mensaje='Error'):
    with lock_progreso:
        if tarea_id in tareas_progreso:
            tarea = tareas_progreso[tarea_id]
            tarea['estado'] = 'error'
            tarea['mensaje'] = mensaje
            tarea['fin'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

def obtener_progreso(tarea_id):
    with lock_progreso:
        if tarea_id in tareas_progreso:
            return tareas_progreso[tarea_id].copy()
    return None

def limpiar_tarea(tarea_id):
    with lock_progreso:
        if tarea_id in tareas_progreso:
            del tareas_progreso[tarea_id]

def limpiar_antiguas(max_edad_segundos=300):
    """Limpia tareas completadas hace más de X segundos"""
    with lock_progreso:
        ahora = datetime.now()
        tareas_viejas = []
        for tid, tarea in tareas_progreso.items():
            if tarea['estado'] in ('completado', 'error') and tarea['fin']:
                # ✅ CORREGIDO: Usar strptime en lugar de fromisoformat (Python 3.6)
                try:
                    fin = datetime.strptime(tarea['fin'], '%Y-%m-%d %H:%M:%S')
                    if (ahora - fin).total_seconds() > max_edad_segundos:
                        tareas_viejas.append(tid)
                except Exception:
                    tareas_viejas.append(tid)
        for tid in tareas_viejas:
            del tareas_progreso[tid]
