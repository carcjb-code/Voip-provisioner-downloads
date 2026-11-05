import os
import json
import time
import shutil
import subprocess
import logging
import threading
from datetime import datetime, timedelta
from app.config import DATA_FOLDER, cargar_config, cargar_perfiles, buscar_perfil, get_excel_file
from app.ssh_client import ejecutar_comando_ssh, conectar_sftp
from app.asterisk import obtener_credenciales_mysql, es_host_local

logger = logging.getLogger(__name__)

# Rutas base de respaldos.
# En Android/Windows, los respaldos del cliente deben quedar en datos privados,
# nunca intentar crear /opt (que no existe y es de solo lectura en Android).
_BACKUP_OVERRIDE = os.environ.get('VOIP_PROVISIONER_BACKUP_DIR', '').strip()
if _BACKUP_OVERRIDE:
    BACKUP_DIR = _BACKUP_OVERRIDE
    PROVISIONER_BACKUP_DIR = BACKUP_DIR
    ISSABEL_BACKUP_DIR = os.path.join(BACKUP_DIR, 'Issabel')
elif os.name == 'nt' or os.environ.get('VOIP_PROVISIONER_DESKTOP') == '1':
    BACKUP_DIR = os.path.join(DATA_FOLDER, 'backups')
    PROVISIONER_BACKUP_DIR = os.path.join(BACKUP_DIR, 'voip-provisioner')
    ISSABEL_BACKUP_DIR = os.path.join(BACKUP_DIR, 'Issabel')
else:
    # Instalación del servidor Issabel: conservar las rutas históricas.
    BACKUP_DIR = '/opt/backup'
    PROVISIONER_BACKUP_DIR = os.path.join(BACKUP_DIR, 'voip-provisioner')
    ISSABEL_BACKUP_DIR = os.path.join(BACKUP_DIR, 'Issabel')

# Crear carpetas si no existen
os.makedirs(PROVISIONER_BACKUP_DIR, exist_ok=True)
os.makedirs(ISSABEL_BACKUP_DIR, exist_ok=True)

def asegurar_directorio(perfil_nombre=''):
    """Crea y retorna el directorio de backup según el tipo"""
    if not perfil_nombre or perfil_nombre == '_provisioner':
        dir_path = PROVISIONER_BACKUP_DIR
    else:
        dir_path = ISSABEL_BACKUP_DIR
    os.makedirs(dir_path, exist_ok=True)
    return dir_path

def crear_backup(perfil_nombre='', destino='local'):
    """
    Crea un backup.
    - Si perfil_nombre está vacío: backup del propio provisioner
    - Si perfil_nombre tiene valor: backup del servidor Issabel (BD asterisk + /etc/asterisk)
    """
    try:
        if not perfil_nombre:
            return _crear_backup_provisioner()
        else:
            return _crear_backup_issabel(perfil_nombre, destino)
    except Exception as e:
        logger.error(f"Error en crear_backup: {e}", exc_info=True)
        return False, str(e), 0

def _crear_backup_provisioner():
    """Backup del propio servidor provisioner en /opt/backup/voip-provisioner"""
    fecha = datetime.now().strftime("%Y%m%d_%H%M%S")
    nombre = f"provisioner_{fecha}.tar.gz"
    dir_path = asegurar_directorio('')
    ruta = os.path.join(dir_path, nombre)
    
    cmd = f"tar -czf '{ruta}' --exclude='voip-provisioner/*.log' -C /opt voip-provisioner"
    result = subprocess.run(cmd, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=120)
    
    if result.returncode > 1:
        return False, f"Error tar: {result.stderr.decode()[:200]}", 0
    
    size = os.path.getsize(ruta) if os.path.exists(ruta) else 0
    logger.info(f"Backup provisioner creado: {nombre} ({size} bytes)")
    return True, nombre, size

def _crear_backup_issabel(perfil_nombre, destino):
    """Backup de Issabel (troncales, internos, colas, grupos) en /opt/backup/Issabel"""
    pf = buscar_perfil(perfil_nombre)
    if not pf:
        return False, f"Perfil '{perfil_nombre}' no encontrado", 0
    
    fecha = datetime.now().strftime("%Y%m%d_%H%M%S")
    host_safe = pf['host'].replace('.', '_').replace(':', '_')
    nombre = f"issabel_{host_safe}_{fecha}.tar.gz"
    temp_dir = f"/tmp/backup_issabel_{int(time.time())}"
    
    try:
        db_user, db_pass = obtener_credenciales_mysql(pf)
    except Exception as e:
        return False, f"Error obteniendo credenciales MySQL: {str(e)[:100]}", 0

    # Crear carpeta temporal remota y copiar ahi el Excel del perfil (edificio/piso/S-N)
    # ANTES de armar el tar, para que quede todo en un solo paquete autocontenido.
    ejecutar_comando_ssh(pf, f"mkdir -p {temp_dir}")
    excel_path = get_excel_file(perfil_nombre)
    if os.path.exists(excel_path):
        try:
            if es_host_local(pf['host']):
                shutil.copy2(excel_path, os.path.join(temp_dir, os.path.basename(excel_path)))
            else:
                s_excel, _ = conectar_sftp(
                    pf['host'], int(pf.get('puerto', 22)),
                    pf['usuario'], pf['password'],
                    '/tmp', pf.get('ssh_compat', False)
                )
                with s_excel:
                    s_excel.put(excel_path, f"{temp_dir}/{os.path.basename(excel_path)}")
        except Exception as e:
            logger.warning(f"BACKUP: no se pudo incluir el Excel del perfil ({excel_path}): {e}")
    else:
        logger.info(f"BACKUP: no existe todavia Excel para el perfil '{perfil_nombre}' ({excel_path}), se omite")

    # Extrae BD asterisk (extensiones, troncales, colas, rutas, grupos), /etc/asterisk,
    # /tftpboot (configs de telefonos: MAC, provisioning), buzones de voz, sonidos custom
    # y el Excel del perfil ya copiado arriba.
    cmd_backup = f"""
set -e
cd {temp_dir}

mysqldump -u {db_user} -p'{db_pass}' --single-transaction --quick asterisk > asterisk_db.sql 2>/dev/null || true

if [ -d /etc/asterisk ]; then
    tar -czf asterisk_etc.tar.gz -C / etc/asterisk 2>/dev/null || true
fi

if [ -d /tftpboot ]; then
    tar -czf tftpboot.tar.gz -C / tftpboot 2>/dev/null || true
fi

if [ -d /var/spool/asterisk/voicemail ]; then
    tar -czf voicemail.tar.gz -C / var/spool/asterisk/voicemail 2>/dev/null || true
fi

if [ -d /var/lib/asterisk/sounds/custom ]; then
    tar -czf sounds_custom.tar.gz -C / var/lib/asterisk/sounds/custom 2>/dev/null || true
fi

tar -czf {nombre} asterisk_db.sql asterisk_etc.tar.gz tftpboot.tar.gz voicemail.tar.gz sounds_custom.tar.gz *.xlsx 2>/dev/null || \
tar -czf {nombre} asterisk_db.sql *.xlsx 2>/dev/null || \
tar -czf {nombre} asterisk_db.sql 2>/dev/null

rm -f asterisk_db.sql asterisk_etc.tar.gz tftpboot.tar.gz voicemail.tar.gz sounds_custom.tar.gz *.xlsx 2>/dev/null || true

echo "{temp_dir}/{nombre}"
"""

    logger.info(f"BACKUP: Ejecutando backup completo en {pf['host']}...")
    out, err, code = ejecutar_comando_ssh(pf, cmd_backup, timeout=180)
    
    if code != 0 or not out:
        return False, f"Error ejecutando backup remoto (code={code}): {err[:200]}", 0
    
    ruta_remota = out.strip().split('\n')[-1].strip()
    
    if destino == 'remoto':
        cmd_mover = f"mkdir -p /root/backups_issabel && mv '{ruta_remota}' /root/backups_issabel/ && rm -rf {temp_dir}"
        ejecutar_comando_ssh(pf, cmd_mover)
        logger.info(f"BACKUP: Guardado en servidor remoto: /root/backups_issabel/{nombre}")
        return True, f"[REMOTO] /root/backups_issabel/{nombre}", 0
    else:
        try:
            dir_path = asegurar_directorio(perfil_nombre)
            ruta_local = os.path.join(dir_path, nombre)
            
            s, _ = conectar_sftp(
                pf['host'], int(pf.get('puerto', 22)), 
                pf['usuario'], pf['password'], 
                '/tmp', pf.get('ssh_compat', False)
            )
            with s:
                s.get(ruta_remota, ruta_local)
            
            ejecutar_comando_ssh(pf, f"rm -rf {temp_dir}")
            
            size = os.path.getsize(ruta_local)
            logger.info(f"BACKUP: Descargado a {ruta_local} ({size} bytes)")
            return True, nombre, size
        except Exception as e:
            ejecutar_comando_ssh(pf, f"rm -rf {temp_dir}")
            return False, f"Error descargando backup: {str(e)[:200]}", 0

def listar_backups(perfil_nombre=''):
    """Lista los backups de ambas carpetas"""
    backups = []
    try:
        dirs_a_revisar = [
            ('Provisioner', PROVISIONER_BACKUP_DIR),
            ('Issabel', ISSABEL_BACKUP_DIR)
        ]
        
        for perfil_tag, dir_path in dirs_a_revisar:
            if not os.path.exists(dir_path):
                continue
            for f in os.listdir(dir_path):
                if f.endswith('.tar.gz') or f.endswith('.tgz'):
                    filepath = os.path.join(dir_path, f)
                    try:
                        stat = os.stat(filepath)
                        backups.append({
                            'nombre': f,
                            'perfil': perfil_tag,
                            'fecha': datetime.fromtimestamp(stat.st_mtime).isoformat(),
                            'tamano': stat.st_size,
                            'ruta': filepath
                        })
                    except Exception as e:
                        logger.warning(f"No se pudo leer {filepath}: {e}")
        
        backups.sort(key=lambda x: x['fecha'], reverse=True)
        return backups
    except Exception as e:
        logger.error(f"Error listando backups: {e}", exc_info=True)
        return []

def eliminar_backup(nombre, perfil_nombre=''):
    """Elimina un backup buscando en ambas carpetas"""
    try:
        for carpeta in [PROVISIONER_BACKUP_DIR, ISSABEL_BACKUP_DIR]:
            filepath = os.path.join(carpeta, nombre)
            if os.path.exists(filepath):
                os.remove(filepath)
                return True, "Backup eliminado"
        return False, "Backup no encontrado"
    except Exception as e:
        return False, str(e)

def restaurar_backup(ruta_archivo, reiniciar_servicio=False):
    """Restaura un backup del provisioner"""
    try:
        if not os.path.exists(ruta_archivo):
            return False, "Archivo no encontrado"
        
        cmd = f"tar -xzf '{ruta_archivo}' -C /opt 2>&1"
        result = subprocess.run(cmd, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=120)
        
        if result.returncode > 1:
            return False, f"Error extrayendo: {result.stderr.decode()[:200]}"
        
        if reiniciar_servicio:
            threading.Thread(
                target=lambda: (time.sleep(2), subprocess.run(['systemctl', 'restart', 'voip-provisioner'])),
                daemon=True
            ).start()
        
        return True, "Backup restaurado correctamente"
    except Exception as e:
        return False, str(e)

def cargar_programacion():
    config_path = os.path.join(BACKUP_DIR, '.schedule.json')
    try:
        if os.path.exists(config_path):
            with open(config_path, 'r') as f:
                return json.load(f)
    except Exception as e:
        logger.warning(f"Error cargando programación: {e}")
    return {'activo': False, 'frecuencia': 'diario', 'hora': '02:00', 'retencion_dias': 30}

def guardar_programacion(schedule):
    try:
        config_path = os.path.join(BACKUP_DIR, '.schedule.json')
        with open(config_path, 'w') as f:
            json.dump(schedule, f, indent=2)
        return True
    except Exception as e:
        logger.error(f"Error guardando programación: {e}")
        return False

# ============================================================
# SCHEDULER DE BACKUPS
# ============================================================

_scheduler_thread = None
_scheduler_lock = threading.Lock()
_scheduler_ejecutando = False
_scheduler_estado = {
    'activo': False,
    'ejecutando': False,
    'proxima_ejecucion': None,
    'proximo_ejecutado': None,
    'ultima_ejecucion': None,
    'ultimo_resultado': None,
    'ultimo_mensaje': None
}


def calcular_proxima_ejecucion(frecuencia, hora):
    """
    Calcula la próxima ejecución del backup.

    diario   -> todos los días a HH:MM
    semanal  -> mismo día de la semana a HH:MM
    mensual  -> día 1 del próximo mes a HH:MM
    """
    try:
        ahora = datetime.now()
        hora_dt = datetime.strptime(hora, '%H:%M')

        candidato = ahora.replace(
            hour=hora_dt.hour,
            minute=hora_dt.minute,
            second=0,
            microsecond=0
        )

        frecuencia = (frecuencia or 'diario').lower()

        if frecuencia == 'diario':
            if candidato <= ahora:
                candidato += timedelta(days=1)

        elif frecuencia == 'semanal':
            if candidato <= ahora:
                candidato += timedelta(days=7)

        elif frecuencia == 'mensual':
            # El scheduler no tiene día configurable.
            # Se usa el día 1 de cada mes.
            if ahora.month == 12:
                siguiente_mes = datetime(
                    ahora.year + 1, 1, 1,
                    hora_dt.hour, hora_dt.minute
                )
            else:
                siguiente_mes = datetime(
                    ahora.year,
                    ahora.month + 1,
                    1,
                    hora_dt.hour,
                    hora_dt.minute
                )

            candidato = siguiente_mes

        else:
            logger.warning(
                "Frecuencia de backup desconocida: %s",
                frecuencia
            )
            return None

        return candidato

    except Exception as e:
        logger.error(
            "Error calculando próxima ejecución: %s",
            e,
            exc_info=True
        )
        return None


def _limpiar_backups_antiguos(retencion_dias):
    """
    Elimina backups locales cuya antigüedad supere la retención.
    """
    try:
        dias = int(retencion_dias)
    except Exception:
        dias = 30

    if dias < 1:
        dias = 1

    limite = time.time() - (dias * 86400)
    eliminados = 0

    for carpeta in (
        PROVISIONER_BACKUP_DIR,
        ISSABEL_BACKUP_DIR
    ):
        if not os.path.isdir(carpeta):
            continue

        try:
            for nombre in os.listdir(carpeta):
                if not (
                    nombre.endswith('.tar.gz') or
                    nombre.endswith('.tgz')
                ):
                    continue

                ruta = os.path.join(carpeta, nombre)

                try:
                    if os.path.isfile(ruta):
                        if os.path.getmtime(ruta) < limite:
                            os.remove(ruta)
                            eliminados += 1
                            logger.info(
                                "Backup antiguo eliminado: %s",
                                ruta
                            )
                except Exception as e:
                    logger.warning(
                        "No se pudo eliminar backup antiguo %s: %s",
                        ruta,
                        e
                    )

        except Exception as e:
            logger.warning(
                "Error revisando retención en %s: %s",
                carpeta,
                e
            )

    return eliminados


def ejecutar_backup_programado():
    """
    Ejecuta inmediatamente el backup definido en .schedule.json.

    Retorna:
        (True, mensaje)
        (False, mensaje)
    """
    global _scheduler_ejecutando

    with _scheduler_lock:
        if _scheduler_ejecutando:
            return False, 'Ya hay un backup en ejecución'

        _scheduler_ejecutando = True
        _scheduler_estado['ejecutando'] = True

    try:
        schedule = cargar_programacion()

        if not schedule.get('activo', False):
            return False, 'La programación automática está desactivada'

        tipo = schedule.get('tipo', 'provisioner')
        perfil_nombre = schedule.get('perfil_nombre', '')
        destino = schedule.get('destino', 'local')

        inicio = datetime.now()

        if tipo == 'issabel':
            if not perfil_nombre:
                return False, 'No hay un perfil Issabel seleccionado'

            ok, mensaje, tamano = crear_backup(
                perfil_nombre=perfil_nombre,
                destino=destino
            )

        else:
            ok, mensaje, tamano = crear_backup(
                perfil_nombre='',
                destino='local'
            )

        retencion = schedule.get('retencion_dias', 30)
        eliminados = _limpiar_backups_antiguos(retencion)

        fin = datetime.now()

        if ok:
            texto = str(mensaje)

            if eliminados:
                texto += ' | Retención: {} backup(s) eliminado(s)'.format(
                    eliminados
                )

            _scheduler_estado['ultima_ejecucion'] = fin.isoformat()
            _scheduler_estado['ultimo_resultado'] = 'ok'
            _scheduler_estado['ultimo_mensaje'] = texto

            logger.info(
                "Backup programado completado en %s: %s",
                fin - inicio,
                texto
            )

            return True, texto

        _scheduler_estado['ultima_ejecucion'] = fin.isoformat()
        _scheduler_estado['ultimo_resultado'] = 'error'
        _scheduler_estado['ultimo_mensaje'] = str(mensaje)

        logger.error(
            "Backup programado falló: %s",
            mensaje
        )

        return False, str(mensaje)

    except Exception as e:
        mensaje = str(e)

        _scheduler_estado['ultima_ejecucion'] = datetime.now().isoformat()
        _scheduler_estado['ultimo_resultado'] = 'error'
        _scheduler_estado['ultimo_mensaje'] = mensaje

        logger.error(
            "Error ejecutando backup programado: %s",
            e,
            exc_info=True
        )

        return False, mensaje

    finally:
        with _scheduler_lock:
            _scheduler_ejecutando = False
            _scheduler_estado['ejecutando'] = False


def obtener_estado_scheduler():
    """
    Devuelve el estado actual del scheduler.
    """
    try:
        schedule = cargar_programacion()

        activo = bool(schedule.get('activo', False))

        proxima = calcular_proxima_ejecucion(
            schedule.get('frecuencia', 'diario'),
            schedule.get('hora', '02:00')
        )

        proxima_iso = None
        if proxima:
            proxima_iso = proxima.isoformat()

        _scheduler_estado['activo'] = activo
        _scheduler_estado['proxima_ejecucion'] = proxima_iso
        _scheduler_estado['proximo_ejecutado'] = proxima_iso

        estado = dict(_scheduler_estado)

        estado['frecuencia'] = schedule.get(
            'frecuencia',
            'diario'
        )

        estado['hora'] = schedule.get(
            'hora',
            '02:00'
        )

        estado['tipo'] = schedule.get(
            'tipo',
            'provisioner'
        )

        estado['perfil_nombre'] = schedule.get(
            'perfil_nombre',
            ''
        )

        estado['destino'] = schedule.get(
            'destino',
            'local'
        )

        estado['retencion_dias'] = schedule.get(
            'retencion_dias',
            30
        )

        return estado

    except Exception as e:
        logger.error(
            "Error obteniendo estado del scheduler: %s",
            e,
            exc_info=True
        )

        return {
            'activo': False,
            'ejecutando': False,
            'proxima_ejecucion': None,
            'proximo_ejecutado': None,
            'ultima_ejecucion': None,
            'ultimo_resultado': 'error',
            'ultimo_mensaje': str(e)
        }


def _scheduler_loop():
    """
    Hilo principal del scheduler de backups.

    La próxima ejecución se toma de proximo_ejecutado
    guardado en /opt/backup/.schedule.json.

    Cuando llega la hora:
      1. Ejecuta el backup.
      2. Calcula la próxima ejecución.
      3. Guarda nuevamente proximo_ejecutado.
    """
    logger.info("Scheduler de backups iniciado")

    ultima_clave = None

    while True:
        try:
            schedule = cargar_programacion()

            if not schedule.get('activo', False):
                time.sleep(30)
                continue

            frecuencia = schedule.get('frecuencia', 'diario')
            hora = schedule.get('hora', '02:00')

            ahora = datetime.now()

            # ---------------------------------------------------------
            # 1. Obtener la ejecución programada desde el archivo.
            # ---------------------------------------------------------
            proxima = None
            proxima_str = schedule.get('proximo_ejecutado')

            if proxima_str:
                try:
                    proxima = datetime.strptime(str(proxima_str).strip(), '%Y-%m-%dT%H:%M:%S')
                except Exception:
                    logger.warning(
                        "proximo_ejecutado inválido: %s",
                        proxima_str
                    )
                    proxima = None

            # ---------------------------------------------------------
            # 2. Si no existe una fecha válida, calcularla una vez.
            # ---------------------------------------------------------
            if proxima is None:
                proxima = calcular_proxima_ejecucion(
                    frecuencia,
                    hora
                )

                if proxima:
                    schedule['proximo_ejecutado'] = proxima.isoformat()

                    try:
                        guardar_programacion(schedule)
                    except Exception as e:
                        logger.error(
                            "No se pudo guardar proximo_ejecutado: %s",
                            e,
                            exc_info=True
                        )

            if not proxima:
                time.sleep(30)
                continue

            # ---------------------------------------------------------
            # 3. Si la ejecución llegó o quedó pendiente, ejecutar.
            # ---------------------------------------------------------
            if ahora >= proxima:

                clave = proxima.strftime('%Y-%m-%d %H:%M')

                # Evita ejecutar dos veces la misma programación.
                if clave != ultima_clave:

                    ultima_clave = clave

                    logger.info(
                        "Ejecutando backup programado: %s",
                        clave
                    )

                    ok, mensaje = ejecutar_backup_programado()

                    if ok:
                        logger.info(
                            "Backup programado completado: %s",
                            mensaje
                        )

                        # -------------------------------------------------
                        # 4. Calcular la siguiente ejecución.
                        #
                        # Como ya estamos dentro del horario programado,
                        # calcular_proxima_ejecucion() avanzará al siguiente
                        # día/semana/mes.
                        # -------------------------------------------------
                        siguiente = calcular_proxima_ejecucion(
                            frecuencia,
                            hora
                        )

                        if siguiente:
                            schedule['proximo_ejecutado'] = (
                                siguiente.isoformat()
                            )

                            try:
                                guardar_programacion(schedule)

                                logger.info(
                                    "Próximo backup programado: %s",
                                    siguiente.isoformat()
                                )

                            except Exception as e:
                                logger.error(
                                    "No se pudo guardar la próxima "
                                    "ejecución: %s",
                                    e,
                                    exc_info=True
                                )

                    else:
                        logger.error(
                            "Error en backup programado: %s",
                            mensaje
                        )

            time.sleep(30)

        except Exception as e:
            logger.error(
                "Error en loop del scheduler: %s",
                e,
                exc_info=True
            )

            time.sleep(30)


def iniciar_scheduler():
    """
    Inicia el scheduler una sola vez.
    """
    global _scheduler_thread

    with _scheduler_lock:
        if (
            _scheduler_thread is not None and
            _scheduler_thread.is_alive()
        ):
            return

        _scheduler_thread = threading.Thread(
            target=_scheduler_loop,
            name='backup-scheduler',
            daemon=True
        )

        _scheduler_thread.start()

        logger.info(
            "Scheduler de backups iniciado correctamente"
        )


# Arrancar automáticamente al cargar el módulo.
iniciar_scheduler()
