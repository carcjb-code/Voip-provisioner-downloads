import shutil
import tempfile
import os, re, json, time, queue, threading, logging, subprocess
import posixpath
import threading
from datetime import datetime
import paramiko
import uuid
import secrets
import select
import base64
import inspect
from flask import Blueprint, request, jsonify, render_template, send_file, Response
from flask import current_app
from openpyxl import Workbook, load_workbook
from app.config import (COLUMNAS_EXCEL, CARPETA_CONFIGS, cargar_config, guardar_config,
                        cargar_perfiles, guardar_perfiles, buscar_perfil, get_excel_file,
                        ALARMAS_PERFILES, PLANTILLAS_FOLDER, cargar_plantillas, PLANTILLAS_PERMITIDAS,
                        cargar_plantillas_eliminadas, guardar_plantillas_eliminadas,
                        cargar_ubicaciones_perfiles, guardar_ubicaciones_perfiles,
                        get_excel_file_legacy, SFTP_SERVERS_FILE)
from app.ssh_client import ejecutar_comando_ssh, conectar_sftp
from app.asterisk import (interno_existe_mysql, crear_peer_en_asterisk,
                          crear_interno_individual_en_asterisk,
                          crear_internos_lote_en_asterisk,
                          obtener_password_interno_issabel,
                          eliminar_interno_asterisk, sincronizar_internos_rapido,
                          obtener_credenciales_mysql, recargar_asterisk,
                          habilitar_paralelo_pjsip,
                          actualizar_nombres_interno_asterisk)
from app.monitor import (
    profile_monitors,
    monitor_lock,
    monitor_asterisk,
    obtener_estado_completo,
    colgar_canales
)
from app.progress import (iniciar_tarea, actualizar_progreso, completar_tarea,
                          error_tarea, obtener_progreso, limpiar_tarea, limpiar_antiguas)
from app.backup import (crear_backup, listar_backups, eliminar_backup, restaurar_backup, 
                        cargar_programacion, guardar_programacion, iniciar_scheduler)

logger = logging.getLogger(__name__)


def _vp_servidores_backup_habilitado():
    """Permite respaldar perfiles en Android y en la app de escritorio Windows."""
    return updater_enabled() or os.environ.get('VOIP_PROVISIONER_DESKTOP') == '1'

def _vp_host_es_local(host):
    """Detecta si el perfil apunta a esta misma máquina, sin IPs de laboratorio."""
    if os.environ.get('VOIP_PROVISIONER_DESKTOP') == '1':
        return False
    host = str(host or '').strip().lower().strip('[]')
    if host in ('127.0.0.1', 'localhost', '::1'):
        return True
    if not host:
        return False
    try:
        import socket
        nombres_locales = set()
        ips_locales = set()
        for nombre in (socket.gethostname(), socket.getfqdn()):
            nombre = str(nombre or '').strip().lower()
            if nombre:
                nombres_locales.add(nombre)
                try:
                    for info in socket.getaddrinfo(nombre, None):
                        ips_locales.add(str(info[4][0]).split('%', 1)[0].lower())
                except Exception:
                    pass
        if host in nombres_locales:
            return True
        try:
            p = subprocess.run(
                ['hostname', '-I'],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                universal_newlines=True,
                timeout=2
            )
            if p.returncode == 0:
                ips_locales.update(
                    str(ip).split('%', 1)[0].lower()
                    for ip in (p.stdout or '').split()
                )
        except Exception:
            pass
        ips_host = set(
            str(info[4][0]).split('%', 1)[0].lower()
            for info in socket.getaddrinfo(host, None)
        )
        return bool(ips_host.intersection(ips_locales))
    except Exception:
        return False

from app.mac_vendor_db import lookup_mac_vendor, refresh_ieee_database, refresh_ieee_database_from_files
from app.ota_updater import (
    updater_enabled, updater_status, check_for_updates, apply_updates
)

main_bp = Blueprint('main', __name__)

REGEX_LIMPIAR_MAC = re.compile(r'[:\-\.]')

def aplicar_formato_excel(wb, ws):
    """Aplica formato al Excel: encabezado con color primario, anchos adaptados, texto centrado"""
    try:
        from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
        from openpyxl.utils import get_column_letter
        
        cfg = cargar_config()
        color_hex = cfg.get('empresa', {}).get('color_primario', '#4da6ff').lstrip('#')
        
        header_fill = PatternFill(start_color=color_hex, end_color=color_hex, fill_type='solid')
        header_font = Font(color='FFFFFF', bold=True, size=11)
        header_align = Alignment(horizontal='center', vertical='center')
        cell_align = Alignment(horizontal='center', vertical='center')
        thin_border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
        
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_align
            cell.border = thin_border
        
        for col_idx, col in enumerate(ws.columns, 1):
            max_length = 0
            col_letter = get_column_letter(col_idx)
            for cell in col:
                try:
                    if cell.value:
                        cell_len = len(str(cell.value))
                        if cell_len > max_length:
                            max_length = cell_len
                    cell.alignment = cell_align
                    cell.border = thin_border
                except:
                    pass
            adjusted_width = max(max_length + 3, 10)
            if adjusted_width > 40:
                adjusted_width = 40
            ws.column_dimensions[col_letter].width = adjusted_width
        
        ws.freeze_panes = 'A2'
    except Exception as e:
        logger.error("Error aplicando formato Excel: {}".format(e))


def verificar_duplicados_mac_sn(ep, interno_actual, mac_nueva, sn_nuevo):
    """Verifica si existe otra fila con la misma MAC o S/N."""
    if not os.path.exists(ep):
        return None, None
    
    try:
        wb = load_workbook(ep)
        ws = wb.active
        h = [c.value for c in ws[1]]
        tn = len(h) >= 8 and h[3] == "Nombre Mostrar"
        ci = 5 if tn else 3
        cm = 6 if tn else 4
        
        mac_nueva_clean = REGEX_LIMPIAR_MAC.sub('', str(mac_nueva or '').lower())
        sn_nuevo_clean = str(sn_nuevo or '').strip().upper()
        
        for row in ws.iter_rows(min_row=2):
            if len(row) <= cm:
                continue
            
            row_int = str(row[ci].value or '').strip()
            if row_int == interno_actual:
                continue
            
            if mac_nueva_clean:
                row_mac = REGEX_LIMPIAR_MAC.sub('', str(row[cm].value or '').lower())
                if row_mac and row_mac == mac_nueva_clean:
                    wb.close()
                    return "MAC_DUPLICADA", row_int
            
            if sn_nuevo_clean and tn and len(row) > 8:
                row_sn = str(row[8].value or '').strip().upper()
                if row_sn and row_sn == sn_nuevo_clean:
                    wb.close()
                    return "SN_DUPLICADO", row_int
        
        wb.close()
    except Exception as e:
        logger.error("Error verificando duplicados: " + str(e))
    
    return None, None


def macs_asignadas_interno(ep, interno):
    """MACs (12 hex, minusculas) que el Excel del perfil tiene para un interno."""
    macs = []
    interno = str(interno or '').strip()
    if not interno or not os.path.exists(ep):
        return macs
    wb = None
    try:
        wb = load_workbook(ep)
        ws = wb.active
        h = [c.value for c in ws[1]]
        tn = "Nombre Mostrar" in h
        ci = h.index("Interno") if "Interno" in h else (5 if tn else 3)
        cm = h.index("MAC") if "MAC" in h else (6 if tn else 4)
        for row in ws.iter_rows(min_row=2):
            if len(row) <= max(ci, cm):
                continue
            if str(row[ci].value or '').strip() != interno:
                continue
            m = REGEX_LIMPIAR_MAC.sub('', str(row[cm].value or '').lower())
            if m and m not in macs:
                macs.append(m)
    except Exception as e:
        logger.error("Error leyendo MACs del interno {}: {}".format(interno, e))
    finally:
        if wb is not None:
            wb.close()
    return macs


def recargar_issabel_completo(pf):
    if _vp_host_es_local(pf['host']):
        try:
            subprocess.run(['/var/lib/asterisk/bin/retrieve_conf'], stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=30)
            subprocess.run(['/usr/sbin/asterisk', '-rx', 'module reload'], stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=10)
            subprocess.run(['/usr/sbin/asterisk', '-rx', 'dialplan reload'], stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=10)
            subprocess.run(['/usr/sbin/asterisk', '-rx', 'sip reload'], stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=10)
            subprocess.run(['/usr/sbin/asterisk', '-rx', 'pjsip reload'], stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=10)
            subprocess.run(['/usr/sbin/asterisk', '-rx', 'core reload'], stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=10)
            return True
        except Exception as e:
            logger.error("Error en recarga local: {}".format(e))
            return False
    else:
        try:
            cmd = "/var/lib/asterisk/bin/retrieve_conf && asterisk -rx 'module reload' && asterisk -rx 'dialplan reload' && asterisk -rx 'sip reload' && asterisk -rx 'pjsip reload' && asterisk -rx 'core reload'"
            out, err, code = ejecutar_comando_ssh(pf, cmd)
            return code == 0
        except Exception as e:
            logger.error("Error en recarga remota: {}".format(e))
            return False

@main_bp.route('/api/progreso/<tarea_id>')
def api_progreso(tarea_id):
    limpiar_antiguas()
    progreso = obtener_progreso(tarea_id)
    if progreso:
        return jsonify({'status': 'ok', 'progreso': progreso})
    return jsonify({'status': 'not_found'})

@main_bp.route('/api/progreso/limpiar/<tarea_id>', methods=['POST'])
def api_limpiar_progreso(tarea_id):
    limpiar_tarea(tarea_id)
    return jsonify({'status': 'ok'})

def migrar_excel(path):
    if not os.path.exists(path):
        return
    wb = None
    try:
        wb = load_workbook(path)
        ws = wb.active
        h = [c.value for c in ws[1]]

        # Migracion del formato antiguo de seis columnas.
        if len(h) == 6 and h[3] == "Interno":
            ws.insert_cols(4, 2)
            ws.cell(1, 4, "Nombre Mostrar")
            ws.cell(1, 5, "Nombre Saliente")
            for r in ws.iter_rows(min_row=2):
                if len(r) > 5 and r[5].value:
                    r[3].value = r[5].value
                    r[4].value = r[5].value

        # Agregar columnas faltantes sin borrar los datos existentes.
        h = [c.value for c in ws[1]]
        if "Numero Serie" not in h:
            ws.cell(1, ws.max_column + 1, "Numero Serie")
            h = [c.value for c in ws[1]]

        if "Modelo Telefono" not in h:
            ws.cell(1, ws.max_column + 1, "Modelo Telefono")

        aplicar_formato_excel(wb, ws)
        wb.save(path)
    except Exception as e:
        logger.error("Error migrando Excel: {}".format(e))
        raise
    finally:
        if wb:
            wb.close()

def crear_excel_si_no_existe(ep, pn, permitir_vacio=False):
    # Recupera inventario real primero; nunca reemplaza un archivo existente
    # por una planilla vacía si la recuperación no pudo completarse.
    if _vp_excel_tiene_registros(ep):
        return False

    existia = os.path.isfile(ep)
    destino_recuperacion = ep
    temporal_recuperacion = None

    # Si ya había archivo, la recuperación se escribe aparte. Esto impide que
    # una respuesta vacía de Asterisk destruya el Excel activo.
    if existia:
        respaldo_previo = ep + '.bak_recuperacion_' + time.strftime('%Y%m%d_%H%M%S') + '_' + str(os.getpid())
        try:
            shutil.copy2(ep, respaldo_previo)
        except Exception as e:
            raise RuntimeError(
                'No se pudo respaldar el Excel existente; se canceló la operación para protegerlo: {}'.format(e)
            )
        if not os.path.isfile(respaldo_previo):
            raise RuntimeError('No se pudo verificar el respaldo del Excel existente; operación cancelada.')
        logger.info('Excel: respaldo previo a recuperación guardado en %s', respaldo_previo)
        temporal_recuperacion = ep + '.recuperacion_' + time.strftime('%Y%m%d_%H%M%S') + '_' + str(os.getpid()) + '.xlsx'
        destino_recuperacion = temporal_recuperacion

    try:
        recuperado = _vp_actualizar_excel_perfil(
            pn, forzar=True, ep_override=destino_recuperacion
        )
    except Exception:
        recuperado = False
        logger.exception('No se pudo reconstruir el Excel del perfil %s', pn)

    if recuperado and _vp_excel_tiene_registros(destino_recuperacion):
        if existia:
            os.replace(destino_recuperacion, ep)
        return True

    if temporal_recuperacion and os.path.exists(temporal_recuperacion):
        try:
            os.remove(temporal_recuperacion)
        except Exception:
            logger.warning('No se pudo quitar el archivo temporal de recuperación %s', temporal_recuperacion)

    # Un libro existente y legible, aunque solo tenga encabezados, se conserva.
    # migrar_excel() podrá completar su esquema sin descartar datos.
    if os.path.isfile(ep):
        wb_check = None
        try:
            wb_check = load_workbook(ep, read_only=True, data_only=True)
            ws_check = wb_check.active
            primera_fila = next(ws_check.iter_rows(min_row=1, max_row=1, values_only=True), ())
            encabezados = [str(v or '').strip().lower() for v in primera_fila]
            if 'interno' in encabezados:
                return False
        except Exception as e:
            logger.warning('Excel existente no recuperable para %s; se conserva intacto: %s', pn, e)
        finally:
            if wb_check is not None:
                try:
                    wb_check.close()
                except Exception:
                    pass
        raise RuntimeError(
            'No se pudo recuperar el Excel del perfil {}. El archivo existente se conservó intacto: {}. '
            'Verifique el respaldo y el acceso a Asterisk antes de reintentar.'.format(pn, ep)
        )

    # Solo se crea un libro nuevo si realmente no existe uno previo.
    if permitir_vacio:
        directorio = os.path.dirname(ep)
        if directorio and not os.path.isdir(directorio):
            os.makedirs(directorio, exist_ok=True)
        wb = Workbook()
        temporal = None
        try:
            ws = wb.active
            ws.append(COLUMNAS_EXCEL)
            aplicar_formato_excel(wb, ws)
            fd, temporal = tempfile.mkstemp(
                prefix='.vp_excel_vacio_', suffix='.xlsx', dir=(directorio or None)
            )
            os.close(fd)
            wb.save(temporal)
            os.replace(temporal, ep)
            temporal = None
        finally:
            wb.close()
            if temporal and os.path.exists(temporal):
                try:
                    os.remove(temporal)
                except Exception:
                    pass
        return True

    raise RuntimeError(
        'No se pudo recuperar el Excel del perfil {}. Se canceló la operación '
        'para no borrar ni descargar una planilla vacía.'.format(pn)
    )

def _vp_excel_ruta_remota(perfil, ruta_remota=None):
    """Ruta única del Excel maestro en el servidor Issabel del perfil."""
    host = str((perfil or {}).get('host', '') or '').strip()
    if not host or '/' in host or '\\' in host:
        raise RuntimeError('El perfil no contiene un host válido para el Excel remoto.')
    carpeta = str(
        ruta_remota or (perfil or {}).get('ruta_remota', '/tftpboot') or '/tftpboot'
    ).rstrip('/')
    return (carpeta + '/' if carpeta else '/') + host + '.xlsx'


def _vp_sftp_archivo_inexistente(error):
    mensaje = str(error or '').lower()
    return (
        getattr(error, 'errno', None) == 2 or
        'no such file' in mensaje or
        'errno 2' in mensaje or
        'file not found' in mensaje
    )


def subir_excel_remoto(pn, rr, source_path=None, backup=False, raise_on_error=False):
    """Publica el Excel de trabajo como maestro remoto, usando reemplazo seguro."""
    sftp = None
    transporte = None
    temporal_remoto = None
    respaldo_remoto = None
    original_movido = False
    try:
        pf = buscar_perfil(pn, cargar_perfiles())
        if not pf:
            raise RuntimeError('Perfil no encontrado al subir el Excel remoto.')
        if _vp_host_es_local(pf.get('host', '')):
            return True

        origen = source_path or get_excel_file(pn)
        if not os.path.isfile(origen):
            raise RuntimeError('No existe el Excel de trabajo que se debe guardar en el servidor.')
        if not _vp_excel_tiene_registros(origen):
            raise RuntimeError('El Excel de trabajo no pasó la validación; se conservó el maestro remoto.')

        sftp, transporte = conectar_sftp(
            pf['host'], int(pf.get('puerto', 22)), pf['usuario'],
            pf['password'], rr, pf.get('ssh_compat', False)
        )
        if sftp is None:
            raise RuntimeError('No se pudo abrir SFTP contra el servidor remoto.')

        remoto = _vp_excel_ruta_remota(pf, rr)
        sello = time.strftime('%Y%m%d_%H%M%S') + '_' + str(os.getpid()) + '_' + uuid.uuid4().hex[:8]
        temporal_remoto = remoto + '.vp_upload_' + sello + '_' + uuid.uuid4().hex[:8]
        sftp.put(origen, temporal_remoto)
        tam_origen = os.path.getsize(origen)
        stat_temporal = sftp.stat(temporal_remoto)
        if int(getattr(stat_temporal, 'st_size', -1)) != int(tam_origen):
            raise RuntimeError('La copia temporal subida al servidor no coincide en tamaño.')

        existe = True
        try:
            sftp.stat(remoto)
        except Exception as exc:
            if _vp_sftp_archivo_inexistente(exc):
                existe = False
            else:
                raise

        if existe and backup:
            respaldo_remoto = remoto + '.bak_sync_' + sello
            sftp.rename(remoto, respaldo_remoto)
            original_movido = True
            existe = False

        if existe:
            # OpenSSH permite reemplazo atómico con posix_rename. Si no está
            # disponible, moveremos el original a un backup antes del rename.
            try:
                sftp.posix_rename(temporal_remoto, remoto)
                temporal_remoto = None
            except Exception:
                respaldo_remoto = remoto + '.bak_update_' + sello
                sftp.rename(remoto, respaldo_remoto)
                original_movido = True
                sftp.rename(temporal_remoto, remoto)
                temporal_remoto = None
        else:
            sftp.rename(temporal_remoto, remoto)
            temporal_remoto = None

        logger.info('Excel maestro publicado en %s', remoto)
        return True
    except Exception as exc:
        logger.error('Error subiendo Excel maestro del perfil %s: %s', pn, exc, exc_info=True)
        if sftp is not None and original_movido and respaldo_remoto:
            try:
                try:
                    sftp.stat(_vp_excel_ruta_remota(pf, rr))
                    remoto_presente = True
                except Exception as comprobar:
                    remoto_presente = not _vp_sftp_archivo_inexistente(comprobar)
                if not remoto_presente:
                    sftp.rename(respaldo_remoto, _vp_excel_ruta_remota(pf, rr))
            except Exception:
                logger.critical('No se pudo restaurar automáticamente el Excel remoto desde %s', respaldo_remoto, exc_info=True)
        if raise_on_error:
            raise RuntimeError('No se pudo guardar el Excel maestro en el servidor remoto: {}'.format(exc))
        return False
    finally:
        if sftp is not None and temporal_remoto:
            try:
                sftp.remove(temporal_remoto)
            except Exception:
                pass
        if sftp is not None:
            try:
                sftp.close()
            except Exception:
                pass
        if transporte is not None:
            try:
                transporte.close()
            except Exception:
                pass


# === VP EXCEL AUTO-RECUPERACION V1 ===
# El Excel maestro permanece en VoIP Provisioner. Si falta, se intenta recuperar
# su copia remota; si no hay copia, se reconstruye desde Asterisk y configs por MAC.
# Nunca elimina filas solo porque un interno no aparecio en una consulta remota.

_VP_EXCEL_LOCKS = {}
_VP_EXCEL_LOCKS_GUARD = threading.Lock()
_VP_EXCEL_ULTIMA_ACTUALIZACION = {}
_VP_EXCEL_TTL_SEGUNDOS = 20
_VP_EXCEL_MODELOS_CACHE = {}
_VP_EXCEL_MODELOS_CACHE_LOCK = threading.Lock()
_VP_EXCEL_MODELOS_CACHE_TTL = 60


def _vp_excel_lock(nombre):
    with _VP_EXCEL_LOCKS_GUARD:
        lock = _VP_EXCEL_LOCKS.get(nombre)
        if lock is None:
            lock = threading.RLock()
            _VP_EXCEL_LOCKS[nombre] = lock
        return lock


def _vp_excel_mac(valor):
    mac = re.sub(r'[^0-9a-f]', '', str(valor or '').lower())
    return mac if re.match(r'^[0-9a-f]{12}$', mac) else ''


def _vp_excel_header(valor):
    return str(valor or '').strip().lower().replace('í', 'i').replace('ó', 'o').replace('_', ' ')


def _vp_excel_info(path):
    if not os.path.isfile(path):
        return False, 0
    wb = None
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        headers = [_vp_excel_header(c.value) for c in ws[1]]
        if 'interno' not in headers or 'mac' not in headers:
            return False, 0
        ci = headers.index('interno')
        count = 0
        for row in ws.iter_rows(min_row=2, values_only=True):
            if len(row) > ci and str(row[ci] or '').strip():
                count += 1
        return count > 0, count
    except Exception as e:
        logger.warning('Excel no valido %s: %s', path, e)
        return False, 0
    finally:
        if wb is not None:
            try:
                wb.close()
            except Exception:
                pass


def _vp_excel_tiene_registros(path):
    valido, cantidad = _vp_excel_info(path)
    return valido and cantidad > 0


def _vp_excel_preparar_maestro_remoto(perfil_nombre, destino=None, ruta_remota=None):
    """
    Refresca la caché de trabajo desde el Excel maestro del Issabel por SFTP.
    Si no existe todavía en el remoto, migra una copia local antigua una sola vez.
    Un error de conexión/permiso no se interpreta como archivo inexistente.
    """
    nombre = str(perfil_nombre or '').strip()
    destino = destino or get_excel_file(nombre)
    if os.environ.get('VOIP_PROVISIONER_REMOTE_EXCEL') != '1':
        return _vp_excel_tiene_registros(destino)

    perfil = buscar_perfil(nombre)
    if not perfil:
        raise RuntimeError('Perfil no encontrado para consultar el Excel remoto.')

    carpeta = os.path.dirname(os.path.abspath(destino))
    os.makedirs(carpeta, exist_ok=True)
    fd, temporal = tempfile.mkstemp(prefix='.vp_excel_remoto_', suffix='.xlsx', dir=carpeta)
    os.close(fd)
    sftp = None
    transporte = None
    temporal_remoto = None
    try:
        sftp, transporte = conectar_sftp(
            perfil.get('host', ''), int(perfil.get('puerto', 22)),
            perfil.get('usuario', 'root'), perfil.get('password', ''),
            ruta_remota or perfil.get('ruta_remota', '/tftpboot'),
            perfil.get('ssh_compat', False)
        )
        if sftp is None:
            raise RuntimeError('No se pudo abrir SFTP para leer el Excel maestro remoto.')

        remoto = _vp_excel_ruta_remota(perfil, ruta_remota)
        try:
            sftp.stat(remoto)
            remoto_existe = True
        except Exception as exc:
            if _vp_sftp_archivo_inexistente(exc):
                remoto_existe = False
            else:
                raise RuntimeError('No se pudo consultar el Excel remoto {}: {}'.format(remoto, exc))

        if remoto_existe:
            sftp.get(remoto, temporal)
            if not _vp_excel_tiene_registros(temporal):
                raise RuntimeError(
                    'El Excel maestro remoto existe, pero no contiene una estructura válida. '
                    'No se reemplazó ni se regeneró.'
                )
            migrar_excel(temporal)
            if not _vp_excel_tiene_registros(temporal):
                raise RuntimeError('El Excel remoto no pasó la validación después de migrarlo.')
            os.replace(temporal, destino)
            temporal = None
            logger.info('Excel maestro leído desde el servidor remoto: %s', remoto)
            return True

        # Migración inicial: si una versión antigua ya guardó el Excel localmente,
        # se publica como maestro solo cuando se ha confirmado que el remoto no existe.
        candidatos = [get_excel_file_legacy(nombre), destino]
        origen = None
        for candidato in candidatos:
            if candidato and os.path.isfile(candidato) and _vp_excel_tiene_registros(candidato):
                origen = candidato
                break
        if origen:
            shutil.copy2(origen, temporal)
            migrar_excel(temporal)
            if not _vp_excel_tiene_registros(temporal):
                raise RuntimeError('El Excel local antiguo no pasó la validación de migración.')
            temporal_remoto = remoto + '.vp_migrate_' + uuid.uuid4().hex[:10]
            sftp.put(temporal, temporal_remoto)
            stat_temporal = sftp.stat(temporal_remoto)
            if int(getattr(stat_temporal, 'st_size', -1)) != int(os.path.getsize(temporal)):
                raise RuntimeError('La migración del Excel no se verificó correctamente en el remoto.')
            sftp.rename(temporal_remoto, remoto)
            temporal_remoto = None
            os.replace(temporal, destino)
            temporal = None
            logger.info('Excel antiguo migrado al maestro remoto %s desde %s', remoto, origen)
            return True

        logger.info('No existe aún un Excel maestro para el perfil %s; se creará tras consultar Asterisk.', nombre)
        return False
    except Exception:
        logger.exception('No se pudo preparar el Excel maestro remoto del perfil %s', nombre)
        raise
    finally:
        if sftp is not None and temporal_remoto:
            try:
                sftp.remove(temporal_remoto)
            except Exception:
                pass
        if sftp is not None:
            try:
                sftp.close()
            except Exception:
                pass
        if transporte is not None:
            try:
                transporte.close()
            except Exception:
                pass
        if temporal and os.path.isfile(temporal):
            try:
                os.remove(temporal)
            except Exception:
                pass


def _vp_excel_descargar_respaldo(pf, destino):
    import tempfile
    host = str(pf.get('host', '')).strip()
    rr = str(pf.get('ruta_remota', '/tftpboot') or '/tftpboot').rstrip('/')
    if not host or not rr:
        return False
    remoto = rr + '/' + host + '.xlsx'
    folder = os.path.dirname(destino)
    if not os.path.isdir(folder):
        os.makedirs(folder, exist_ok=True)
    fd, temporal = tempfile.mkstemp(prefix='.vp_excel_restore_', suffix='.xlsx', dir=folder)
    os.close(fd)
    sftp = None
    transporte = None
    try:
        if _vp_host_es_local(host):
            if not os.path.isfile(remoto):
                return False
            shutil.copy2(remoto, temporal)
        else:
            sftp, transporte = conectar_sftp(
                host, int(pf.get('puerto', 22)), pf.get('usuario', 'root'),
                pf.get('password', ''), rr, pf.get('ssh_compat', False)
            )
            if sftp is None:
                return False
            sftp.get(remoto, temporal)
        if not _vp_excel_tiene_registros(temporal):
            logger.warning('La copia remota %s no tiene un inventario valido.', remoto)
            return False
        migrar_excel(temporal)
        if not _vp_excel_tiene_registros(temporal):
            return False
        os.replace(temporal, destino)
        logger.info('Excel del perfil recuperado desde %s', remoto)
        return True
    except Exception as e:
        logger.info('No se pudo recuperar %s: %s', remoto, e)
        return False
    finally:
        if sftp is not None:
            try:
                sftp.close()
            except Exception:
                pass
        if transporte is not None:
            try:
                transporte.close()
            except Exception:
                pass
        if os.path.isfile(temporal):
            try:
                os.remove(temporal)
            except Exception:
                pass


def _vp_excel_extensiones_config(contenido):
    # Extraer solo numeros de interno en campos conocidos; nunca leer modelo ni serie.
    patrones = [
        r'<\s*SIP\d+_Phone_Number\s*>\s*["\']?([0-9]{2,10})',
        r'(?im)^\s*SIP\d+\s+Phone\s+Number\s*[:=]\s*["\']?([0-9]{2,10})\b',
        r'<\s*Phone_Number\s*>\s*["\']?([0-9]{2,10})',
        r'<\s*PhoneNumber\s*>\s*["\']?([0-9]{2,10})',
        r'<\s*P35\s*>\s*["\']?([0-9]{2,10})',
        # Fanvil CFG: las variantes de firmware usan line.1 / line1 y user_id / bUserId.
        r'(?im)^\s*(?:sip\.)?line\.?\d+\.(?:bUserId|user_id|userId|username|user_name|PhoneNumber|phone_number|extension)\s*[:=]\s*["\']?([0-9]{2,10})\b',
        r'(?im)^\s*(?:sip\.)?account\.\d+\.(?:user_id|username|user_name)\s*[:=]\s*["\']?([0-9]{2,10})\b',
        r'(?im)^\s*P35\s*[:=]\s*["\']?([0-9]{2,10})\b',
        r'(?im)^\s*(?:user_id|username|user_name|extension|phone_number)\s*[:=]\s*["\']?([0-9]{2,10})\b',
    ]
    encontrados = []
    for patron in patrones:
        for dato in re.findall(patron, contenido or ''):
            interno = str(dato or '').strip()
            if interno and interno not in encontrados:
                encontrados.append(interno)
    return encontrados[:8]


def _vp_excel_detectar_modelo_config(nombre, contenido):
    """Reconoce uno de los tres modelos soportados por su archivo real de aprovisionamiento."""
    bajo = str(contenido or '').lower()
    compacto = ''.join(bajo.split())
    if 'fanvilv61g' in compacto or 'net.dhcp.hostname=v61g' in compacto:
        return 'Fanvil V61G'
    if ('denwadw-320g' in compacto or
            'net.dhcp.hostname=dw-320g' in compacto or
            'net.dhcp.vendorclassid=denwa' in compacto):
        return 'Denwa DW-320G'
    if ('<gs_provision' in bajo and '<p271>' in bajo and
            '<p270>' in bajo and '<p35>' in bajo):
        return 'Grandstream GXP1628'
    return ''


def _vp_excel_macs_remotas(pf):
    mapa = {}
    modelos_por_mac = {}
    modelos_por_interno = {}
    host = str(pf.get('host', '')).strip()
    rr = str(pf.get('ruta_remota', '/tftpboot') or '/tftpboot').rstrip('/')
    if not host or not rr:
        return mapa
    extensiones = ('.xml', '.cfg', '.conf', '.ini', '.txt')
    sftp = None
    transporte = None

    def procesar(nombre, leer):
        if not nombre.lower().endswith(extensiones):
            return
        base = os.path.splitext(os.path.basename(nombre))[0]
        if base.lower().startswith('cfg'):
            base = base[3:]
        mac = _vp_excel_mac(base)
        if not mac:
            return
        try:
            raw = leer()
            if isinstance(raw, bytes):
                raw = raw.decode('utf-8', 'replace')
            contenido = str(raw or '')[:524288]
            internos_config = _vp_excel_extensiones_config(contenido)
            modelo_config = _vp_excel_detectar_modelo_config(nombre, contenido)
            if modelo_config:
                modelos_por_mac[mac] = modelo_config
                for interno_modelo in internos_config:
                    lista_modelos = modelos_por_interno.setdefault(interno_modelo, [])
                    if modelo_config not in lista_modelos:
                        lista_modelos.append(modelo_config)
            for interno in internos_config:
                valores = mapa.setdefault(interno, [])
                if mac not in valores:
                    valores.append(mac)
        except Exception as e:
            logger.debug('No se pudo leer config remota %s: %s', nombre, e)

    try:
        local = _vp_host_es_local(host)
        if local:
            if not os.path.isdir(rr):
                return mapa
            nombres = os.listdir(rr)
            procesados = 0
            for nombre in nombres:
                base = os.path.splitext(nombre)[0]
                if base.lower().startswith('cfg'):
                    base = base[3:]
                if not _vp_excel_mac(base) or not nombre.lower().endswith(extensiones):
                    continue
                ruta = os.path.join(rr, nombre)
                if not os.path.isfile(ruta):
                    continue
                def leer_local(ruta=ruta):
                    with open(ruta, 'rb') as fh:
                        return fh.read(524288)
                procesar(nombre, leer_local)
                procesados += 1
                if procesados >= 1500:
                    break
        else:
            sftp, transporte = conectar_sftp(
                host, int(pf.get('puerto', 22)), pf.get('usuario', 'root'),
                pf.get('password', ''), rr, pf.get('ssh_compat', False)
            )
            if sftp is None:
                return mapa
            procesados = 0
            for item in sftp.listdir_attr(rr):
                nombre = str(item.filename)
                base = os.path.splitext(nombre)[0]
                if base.lower().startswith('cfg'):
                    base = base[3:]
                if not _vp_excel_mac(base) or not nombre.lower().endswith(extensiones):
                    continue
                ruta = rr + '/' + nombre
                def leer(ruta=ruta):
                    with sftp.open(ruta, 'rb') as fh:
                        return fh.read(524288)
                procesar(nombre, leer)
                procesados += 1
                if procesados >= 1500:
                    break
    except Exception as e:
        logger.warning('No se pudo buscar MAC en configs de %s:%s: %s', host, rr, e)
    finally:
        if sftp is not None:
            try:
                sftp.close()
            except Exception:
                pass
        if transporte is not None:
            try:
                transporte.close()
            except Exception:
                pass
    clave_modelos = (host.lower(), rr)
    with _VP_EXCEL_MODELOS_CACHE_LOCK:
        _VP_EXCEL_MODELOS_CACHE[clave_modelos] = {
            'actualizado': time.time(),
            'mac': dict(modelos_por_mac),
            'interno': {k: list(v) for k, v in modelos_por_interno.items()}
        }
    logger.info('Configs remotas: MAC asociables a %s internos; modelos identificados: %s',
                len(mapa), len(modelos_por_mac))
    return mapa


def _vp_excel_modelos_config_cache(perfil):
    host = str((perfil or {}).get('host', '')).strip().lower()
    ruta = str((perfil or {}).get('ruta_remota', '/tftpboot') or '/tftpboot').rstrip('/')
    clave = (host, ruta)
    with _VP_EXCEL_MODELOS_CACHE_LOCK:
        dato = _VP_EXCEL_MODELOS_CACHE.get(clave)
        vigente = bool(dato and time.time() - dato.get('actualizado', 0) < _VP_EXCEL_MODELOS_CACHE_TTL)
    if not vigente:
        try:
            _vp_excel_macs_remotas(perfil)
        except Exception:
            logger.warning('No se pudieron actualizar los modelos desde las configuraciones remotas.', exc_info=True)
    with _VP_EXCEL_MODELOS_CACHE_LOCK:
        dato = _VP_EXCEL_MODELOS_CACHE.get(clave) or {}
        return {
            'mac': dict(dato.get('mac', {})),
            'interno': {k: list(v) for k, v in dato.get('interno', {}).items()}
        }


def _vp_excel_aplicar_modelos_hoja(ws, modelos):
    """Completa modelos reconocibles sin pisar los valores administrativos concretos."""
    encabezados = [_vp_excel_header(c.value) for c in ws[1]]
    if 'interno' not in encabezados or 'mac' not in encabezados or 'modelo telefono' not in encabezados:
        return 0
    ci = encabezados.index('interno')
    cm = encabezados.index('mac')
    cmodelo = encabezados.index('modelo telefono')
    por_mac = (modelos or {}).get('mac', {})
    por_interno = (modelos or {}).get('interno', {})
    cambios = 0
    for num_fila in range(2, ws.max_row + 1):
        interno = str(ws.cell(num_fila, ci + 1).value or '').strip()
        mac = _vp_excel_mac(ws.cell(num_fila, cm + 1).value)
        actual = str(ws.cell(num_fila, cmodelo + 1).value or '').strip()
        detectado = por_mac.get(mac) if mac else ''
        if not detectado and interno:
            candidatos = sorted(set(por_interno.get(interno, []) or []))
            if len(candidatos) == 1:
                detectado = candidatos[0]
        # El OUI de IEEE identifica fabricante/marca aunque no permita deducir
        # el modelo exacto. Solo usamos esa marca como respaldo si no hay modelo
        # leído de la configuración y el campo está vacío o contiene esa marca.
        if not detectado and mac:
            try:
                detectado = str(lookup_mac_vendor(mac) or '').strip()
            except Exception:
                detectado = ''
        if not detectado:
            continue
        marca = ''
        if mac:
            try:
                marca = str(lookup_mac_vendor(mac) or '').strip()
            except Exception:
                marca = ''
        # Los modelos manuales se conservan; solo reemplazar campo vacío o
        # el fabricante genérico que el sistema había inferido desde la MAC.
        if actual and not (marca and actual.casefold() == marca.casefold()):
            continue
        if actual != detectado:
            ws.cell(num_fila, cmodelo + 1).value = detectado
            cambios += 1
    return cambios


def _vp_excel_rellenar_modelos_desde_configs(ep, perfil):
    """Enriquece el Excel con el modelo de las configs reales y publica el maestro remoto."""
    if not perfil or not ep or not os.path.isfile(ep):
        return 0
    nombre = str(perfil.get('nombre', '') or perfil.get('host', '')).strip()
    temporal = None
    wb = None
    cambios = 0
    with _vp_excel_lock(nombre):
        try:
            wb = load_workbook(ep)
            ws = wb.active
            encabezados = [_vp_excel_header(c.value) for c in ws[1]]
            if 'mac' not in encabezados or 'interno' not in encabezados or 'modelo telefono' not in encabezados:
                wb.close()
                wb = None
                return 0
            ci = encabezados.index('interno')
            cm = encabezados.index('mac')
            cmodelo = encabezados.index('modelo telefono')
            necesita = False
            for fila in range(2, ws.max_row + 1):
                actual = str(ws.cell(fila, cmodelo + 1).value or '').strip()
                if not actual:
                    necesita = True
                    break
                mac = _vp_excel_mac(ws.cell(fila, cm + 1).value)
                if mac:
                    try:
                        marca = str(lookup_mac_vendor(mac) or '').strip()
                    except Exception:
                        marca = ''
                    if marca and actual.casefold() == marca.casefold():
                        necesita = True
                        break
            wb.close()
            wb = None
            if not necesita:
                return 0

            # Aunque no se haya encontrado el modelo exacto en los archivos CFG/XML,
            # el catálogo IEEE permite completar al menos fabricante/marca por MAC.
            modelos = _vp_excel_modelos_config_cache(perfil)

            wb = load_workbook(ep)
            ws = wb.active
            cambios = _vp_excel_aplicar_modelos_hoja(ws, modelos)
            if not cambios:
                wb.close()
                wb = None
                return 0
            aplicar_formato_excel(wb, ws)
            carpeta = os.path.dirname(os.path.abspath(ep)) or '.'
            fd, temporal = tempfile.mkstemp(prefix='.vp_modelos_', suffix='.xlsx', dir=carpeta)
            os.close(fd)
            wb.save(temporal)
            wb.close()
            wb = None
            if not _vp_excel_tiene_registros(temporal):
                raise RuntimeError('El Excel con modelos no pasó la validación.')
            if os.environ.get('VOIP_PROVISIONER_REMOTE_EXCEL') == '1':
                try:
                    subir_excel_remoto(
                        nombre, perfil.get('ruta_remota', '/tftpboot'),
                        source_path=temporal, backup=True, raise_on_error=True
                    )
                except Exception:
                    # La copia de descarga igualmente se puede completar; el siguiente
                    # acceso vuelve a leer el maestro remoto como fuente de verdad.
                    logger.warning('Se detectaron modelos, pero no se pudo publicar el Excel enriquecido en el remoto.', exc_info=True)
            os.replace(temporal, ep)
            temporal = None
            logger.info('Excel %s: %s modelos de teléfono completados desde configs remotas.', nombre, cambios)
            return cambios
        except Exception:
            logger.warning('No se pudieron completar los modelos de teléfono desde configs remotas.', exc_info=True)
            return 0
        finally:
            if wb is not None:
                try:
                    wb.close()
                except Exception:
                    pass
            if temporal and os.path.isfile(temporal):
                try:
                    os.remove(temporal)
                except OSError:
                    pass


def _vp_actualizar_excel_perfil(perfil_nombre, forzar=False, ep_override=None):
    import shutil
    import tempfile
    nombre = str(perfil_nombre or '').strip()
    if not nombre:
        return False
    perfil = buscar_perfil(nombre)
    if not perfil:
        return False
    ep = ep_override or get_excel_file(nombre)
    with _vp_excel_lock(nombre):
        try:
            if os.environ.get('VOIP_PROVISIONER_REMOTE_EXCEL') == '1':
                _vp_excel_preparar_maestro_remoto(nombre, ep)
            ahora = time.time()
            if (not forzar and _vp_excel_tiene_registros(ep) and
                    ahora - _VP_EXCEL_ULTIMA_ACTUALIZACION.get(nombre, 0) < _VP_EXCEL_TTL_SEGUNDOS):
                return True

            local_valido = _vp_excel_tiene_registros(ep)
            if not local_valido and os.path.isfile(ep):
                try:
                    migrar_excel(ep)
                except Exception:
                    logger.warning('No se pudo migrar Excel de %s', nombre, exc_info=True)
                local_valido = _vp_excel_tiene_registros(ep)

            if not local_valido:
                local_valido = _vp_excel_descargar_respaldo(perfil, ep)

            # Fuente de verdad primero: obtener internos programados en Asterisk
            # (chan_sip y PJSIP) antes de recorrer archivos de aprovisionamiento.
            datos_remotos = {}
            error_sync = ''
            try:
                resultados, error_sync = _vp_excel_consultar_asterisk_rapido(perfil)
                if isinstance(resultados, dict):
                    for key, value in resultados.items():
                        interno = str(key or '').strip()
                        if re.match(r'^[0-9]{1,10}$', interno):
                            datos_remotos[interno] = value if isinstance(value, dict) else {}
            except Exception as e:
                error_sync = str(e)
                logger.warning('Error consultando internos remotos de %s: %s', nombre, e)

            if not datos_remotos:
                if local_valido:
                    logger.warning('Se conserva el Excel actual de %s: consulta Asterisk vacia/fallida (%s).', nombre, error_sync)
                    return True
                logger.error('No se pudo reconstruir Excel de %s: Asterisk no devolvio internos (%s).', nombre, error_sync)
                return False

            # Solo después de consultar Asterisk se buscan las MAC. Se descartan
            # archivos cuyo interno no esté programado en Asterisk.
            mapa_archivos = _vp_excel_mapa_mac_remoto(perfil)
            mapa_macs = {}
            for interno, macs in mapa_archivos.items():
                interno = str(interno or '').strip()
                if interno in datos_remotos and isinstance(macs, (list, tuple, set)):
                    validas = []
                    for mac in macs:
                        mac = _vp_excel_mac(mac)
                        if mac and mac not in validas:
                            validas.append(mac)
                    if validas:
                        mapa_macs[interno] = sorted(validas)
            _vp_excel_cachear_mapa_mac_index(perfil, datos_remotos, mapa_macs)

            if local_valido:
                migrar_excel(ep)
                if not _vp_excel_tiene_registros(ep):
                    local_valido = False

            if local_valido:
                wb = load_workbook(ep)
                ws = wb.active
                creado = False
            else:
                wb = Workbook()
                ws = wb.active
                ws.title = 'Registro'
                ws.append(COLUMNAS_EXCEL)
                creado = True

            headers = [_vp_excel_header(c.value) for c in ws[1]]
            if 'interno' not in headers or 'mac' not in headers:
                wb.close()
                raise RuntimeError('El Excel de {} no tiene columnas Interno/MAC.'.format(nombre))

            ix = {}
            defaults = {'edificio': 0, 'piso': 1, 'dependencia': 2, 'nombre mostrar': 3,
                        'nombre saliente': 4, 'interno': 5, 'mac': 6, 'ip': 7,
                        'numero serie': 8, 'modelo telefono': 9}
            for key, default in defaults.items():
                ix[key] = headers.index(key) if key in headers else default

            def getv(row_num, key):
                col = ix[key]
                if col < 0 or col >= ws.max_column:
                    return ''
                return ws.cell(row_num, col + 1).value or ''

            def setv(row_num, key, value):
                col = ix[key]
                if col < 0:
                    return False
                actual = getv(row_num, key)
                if str(actual or '').strip() == str(value or '').strip():
                    return False
                ws.cell(row_num, col + 1).value = value
                return True

            filas_por_interno = {}
            for n in range(2, ws.max_row + 1):
                interno = str(getv(n, 'interno') or '').strip()
                if interno:
                    filas_por_interno.setdefault(interno, []).append(n)

            modificado = creado

            def nueva_fila(interno, datos, mac, base=None, ip=''):
                cantidad = max(ws.max_column, len(COLUMNAS_EXCEL))
                valores = [None] * cantidad
                if base:
                    for c in range(min(ws.max_column, cantidad)):
                        valores[c] = ws.cell(base, c + 1).value
                    for key in ('ip', 'numero serie', 'modelo telefono'):
                        if ix[key] >= 0 and ix[key] < cantidad:
                            valores[ix[key]] = ''
                else:
                    nombre_mostrar = str(datos.get('display_name') or interno).strip()
                    if not nombre_mostrar or nombre_mostrar.lower() == 'device':
                        nombre_mostrar = interno
                    valores[ix['nombre mostrar']] = nombre_mostrar
                    valores[ix['nombre saliente']] = interno
                valores[ix['interno']] = interno
                valores[ix['mac']] = str(mac or '').upper()
                valores[ix['ip']] = str(ip or '').strip()
                marca_mac = lookup_mac_vendor(mac)
                if marca_mac and ix.get('modelo telefono', -1) >= 0 and ix['modelo telefono'] < cantidad:
                    valores[ix['modelo telefono']] = marca_mac
                ws.append(valores)
                return ws.max_row

            for interno, datos in sorted(datos_remotos.items(), key=lambda item: item[0]):
                datos = datos or {}
                macs = list(mapa_macs.get(interno, []))
                if not macs:
                    mac = _vp_excel_mac(datos.get('mac', ''))
                    if mac:
                        macs = [mac]
                macs = list(dict.fromkeys([_vp_excel_mac(m) for m in macs if _vp_excel_mac(m)]))
                ip = str(datos.get('ip') or '').strip()
                if ip.lower() in ('', 'dynamic', 'unknown', '0.0.0.0', 'unavailable'):
                    ip = ''
                filas = filas_por_interno.setdefault(interno, [])

                # Fabricante desde base local/offline. Nunca reemplazar un modelo manual.
                for fila_existente in filas:
                    mac_existente = _vp_excel_mac(getv(fila_existente, 'mac'))
                    modelo_existente = str(getv(fila_existente, 'modelo telefono') or '').strip()
                    if mac_existente and not modelo_existente:
                        marca_existente = lookup_mac_vendor(mac_existente)
                        if marca_existente and setv(fila_existente, 'modelo telefono', marca_existente):
                            modificado = True

                # Si un interno tiene varios teléfonos y varias MAC nuevas, no
                # asignarlas arbitrariamente a filas que ya tienen modelo/serie
                # manuales: eso podría cambiar qué IP queda asociada a cada modelo.
                if len(macs) > 1 and filas:
                    macs_ya_asignadas = set(
                        _vp_excel_mac(getv(nf, 'mac')) for nf in filas
                        if _vp_excel_mac(getv(nf, 'mac'))
                    )
                    macs_pendientes = [m for m in macs if m not in macs_ya_asignadas]
                    filas_sin_mac = [nf for nf in filas if not _vp_excel_mac(getv(nf, 'mac'))]
                    hay_datos_manuales = any(
                        str(getv(nf, 'modelo telefono') or '').strip() or
                        str(getv(nf, 'numero serie') or '').strip()
                        for nf in filas_sin_mac
                    )
                    if len(macs_pendientes) > 1 and filas_sin_mac and hay_datos_manuales:
                        logger.warning(
                            'MAC no asignadas automáticamente para interno %s del perfil %s: '
                            'hay varios teléfonos y datos manuales; se evita alterar la asociación modelo/IP.',
                            interno, nombre
                        )
                        continue

                if not filas:
                    if macs:
                        for mac in macs:
                            filas.append(nueva_fila(interno, datos, mac, None, ip if len(macs) == 1 else ''))
                    else:
                        filas.append(nueva_fila(interno, datos, '', None, ip))
                    modificado = True
                    continue

                for mac in macs:
                    fila_match = None
                    fila_sin_mac = None
                    for nf in filas:
                        mac_fila = _vp_excel_mac(getv(nf, 'mac'))
                        if mac_fila == mac:
                            fila_match = nf
                            break
                        if not mac_fila and fila_sin_mac is None:
                            fila_sin_mac = nf
                    if fila_match is not None:
                        if len(macs) == 1 and ip and setv(fila_match, 'ip', ip):
                            modificado = True
                        continue
                    if fila_sin_mac is not None:
                        if setv(fila_sin_mac, 'mac', mac.upper()):
                            modificado = True
                        modelo_sin_mac = str(getv(fila_sin_mac, 'modelo telefono') or '').strip()
                        if not modelo_sin_mac:
                            marca_sin_mac = lookup_mac_vendor(mac)
                            if marca_sin_mac and setv(fila_sin_mac, 'modelo telefono', marca_sin_mac):
                                modificado = True
                        if len(macs) == 1 and ip and setv(fila_sin_mac, 'ip', ip):
                            modificado = True
                        continue
                    # Nunca sobreescribir una MAC distinta ya asignada:
                    # puede tratarse de un segundo telefono del mismo interno.
                    filas.append(nueva_fila(interno, datos, mac, filas[0], ip if len(macs) == 1 else ''))
                    modificado = True

                if not macs and ip and len(filas) == 1:
                    if setv(filas[0], 'ip', ip):
                        modificado = True

            # Actualiza también los modelos: primero usa el contenido de las
            # configuraciones y, como respaldo, el fabricante de la MAC.
            try:
                modelos_actualizados = _vp_excel_modelos_config_cache(perfil)
                cambios_modelo = _vp_excel_aplicar_modelos_hoja(ws, modelos_actualizados)
                if cambios_modelo:
                    modificado = True
                    logger.info('Excel %s: se completaron %s modelos/marcas desde configs/OUI.',
                                nombre, cambios_modelo)
            except Exception:
                logger.warning('No se pudieron enriquecer los modelos del Excel %s.', nombre, exc_info=True)

            if not modificado:
                wb.close()
                return _vp_excel_tiene_registros(ep)

            fd, temporal = tempfile.mkstemp(prefix='.vp_excel_update_', suffix='.xlsx', dir=os.path.dirname(ep))
            os.close(fd)
            try:
                aplicar_formato_excel(wb, ws)
                wb.save(temporal)
                wb.close()
                if not _vp_excel_tiene_registros(temporal):
                    raise RuntimeError('El Excel reconstruido no paso la validacion.')
                if os.path.isfile(ep) and os.environ.get('VOIP_PROVISIONER_REMOTE_EXCEL') != '1':
                    respaldo = ep + '.bak_auto_' + str(int(time.time()))
                    shutil.copy2(ep, respaldo)
                    logger.info('Respaldo de Excel generado: %s', respaldo)
                os.replace(temporal, ep)
            finally:
                if os.path.isfile(temporal):
                    try:
                        os.remove(temporal)
                    except Exception:
                        pass

            if os.environ.get('VOIP_PROVISIONER_REMOTE_EXCEL') == '1':
                subir_excel_remoto(
                    nombre, perfil.get('ruta_remota', '/tftpboot'),
                    source_path=ep, backup=True, raise_on_error=True
                )
            else:
                try:
                    subir_excel_remoto(nombre, perfil.get('ruta_remota', '/tftpboot'), source_path=ep)
                except Exception:
                    logger.warning('Excel actualizado localmente pero no se pudo respaldar en el remoto.', exc_info=True)
            valido = _vp_excel_tiene_registros(ep)
            logger.info('Excel %s actualizado: %s internos consultados; %s internos con MAC en configs.', nombre, len(datos_remotos), len(mapa_macs))
            return valido
        except Exception:
            logger.exception('Error actualizando/reconstruyendo Excel del perfil %s', nombre)
            if os.environ.get('VOIP_PROVISIONER_REMOTE_EXCEL') == '1':
                return False
            return _vp_excel_tiene_registros(ep)
        finally:
            _VP_EXCEL_ULTIMA_ACTUALIZACION[nombre] = time.time()

# === FIN VP EXCEL AUTO-RECUPERACION V1 ===


# === VP ASTERISK PRIMERO Y BORRADO SIN EXCEL V1 ===
# La fuente primaria del inventario es Asterisk. El Excel central se reconstruye
# desde Asterisk si falta. Las bajas pueden ejecutarse consultando Asterisk aunque
# no exista un Excel valido; no se sube una planilla vacia al servidor remoto.

def _vp_excel_mapa_mac_remoto(perfil):
    for nombre_funcion in ('_vp_excel_macs_desde_archivos', '_vp_excel_macs_remotas'):
        funcion = globals().get(nombre_funcion)
        if callable(funcion):
            try:
                resultado = funcion(perfil)
                if isinstance(resultado, dict):
                    limpio = {}
                    for interno, macs in resultado.items():
                        interno = str(interno or '').strip()
                        if not interno:
                            continue
                        if not isinstance(macs, (list, tuple, set)):
                            macs = [macs]
                        valores = []
                        for mac in macs:
                            mac = re.sub(r'[^0-9a-f]', '', str(mac or '').lower())
                            if re.match(r'^[0-9a-f]{12}$', mac) and mac not in valores:
                                valores.append(mac)
                        if valores:
                            limpio[interno] = valores
                    return limpio
            except Exception:
                logger.warning('No se pudo leer el mapa remoto de MAC mediante %s', nombre_funcion, exc_info=True)
    return {}



# === VP INDEX: MAC DESDE ASTERISK + ARCHIVOS DE APROVISIONAMIENTO ===
# Cache breve: la lectura SSH/SFTP se hace desde el actualizador en background,
# nunca bloqueando la respuesta habitual del Index.
_VP_INDEX_MAC_CACHE = {}
_VP_INDEX_MAC_CACHE_LOCK = threading.Lock()
_VP_INDEX_MAC_CACHE_TTL = 180

def _vp_excel_cachear_mapa_mac_index(perfil, datos_asterisk, mapa_macs):
    try:
        host = str(perfil.get('host', '') or '').strip()
        ruta = str(perfil.get('ruta_remota', '/tftpboot') or '/tftpboot').rstrip('/')
        nombre = str(perfil.get('nombre', '') or '').strip()
        clave = (nombre, host, ruta)
        internos_asterisk = set(
            str(k or '').strip() for k in (datos_asterisk or {}).keys()
            if re.match(r'^[0-9]{1,10}$', str(k or '').strip())
        )
        resultado = {}
        for interno, macs in (mapa_macs or {}).items():
            interno = str(interno or '').strip()
            if interno not in internos_asterisk:
                continue
            if not isinstance(macs, (list, tuple, set)):
                macs = [macs]
            validas = []
            for mac in macs:
                mac = re.sub(r'[^0-9a-f]', '', str(mac or '').lower())
                if re.match(r'^[0-9a-f]{12}$', mac) and mac not in validas:
                    validas.append(mac)
            if validas:
                resultado[interno] = sorted(validas)
        with _VP_INDEX_MAC_CACHE_LOCK:
            _VP_INDEX_MAC_CACHE[clave] = (time.time(), dict((k, list(v)) for k, v in resultado.items()))
        logger.info('Index MAC cache perfil=%s: %s internos cruzados con Asterisk.', nombre, len(resultado))
    except Exception:
        logger.warning('No se pudo guardar el cache de MAC para Index.', exc_info=True)

def _vp_excel_mapa_mac_para_index(perfil):
    # Solo leer el cache; nunca abre SSH/SFTP dentro de la petición HTTP.
    try:
        host = str(perfil.get('host', '') or '').strip()
        ruta = str(perfil.get('ruta_remota', '/tftpboot') or '/tftpboot').rstrip('/')
        nombre = str(perfil.get('nombre', '') or '').strip()
        clave = (nombre, host, ruta)
        with _VP_INDEX_MAC_CACHE_LOCK:
            entrada = _VP_INDEX_MAC_CACHE.get(clave)
            if entrada and time.time() - entrada[0] < _VP_INDEX_MAC_CACHE_TTL:
                return dict((k, list(v)) for k, v in entrada[1].items())
    except Exception:
        logger.warning('No se pudo consultar cache de MAC para Index.', exc_info=True)
    return {}

# === FIN VP INDEX MAC ===


# === VP EXCEL RAPIDO SIN BLOQUEOS V1 ===
# Consulta las fuentes SIP/PJSIP con un solo comando SSH, en lugar de abrir una
# sesion SSH por cada AOR PJSIP. La actualizacion enriquecida corre en segundo plano.

def _vp_excel_consultar_asterisk_rapido(perfil):
    host = str(perfil.get('host', '') or '').strip()
    local = _vp_host_es_local(host)
    cmd = (
        "printf '\\n__VP_SIP_CONF_BEGIN__\\n'; "
        "cat /etc/asterisk/sip_additional.conf 2>/dev/null; "
        "printf '\\n__VP_PJSIP_CONF_BEGIN__\\n'; "
        "cat /etc/asterisk/pjsip_additional.conf 2>/dev/null; "
        "printf '\\n__VP_SIP_PEERS_BEGIN__\\n'; "
        "asterisk -rx 'sip show peers' 2>/dev/null; "
        "printf '\\n__VP_PJSIP_CONTACTS_BEGIN__\\n'; "
        "asterisk -rx 'pjsip show contacts' 2>/dev/null; "
        "printf '\\n__VP_DEVICE_DB_BEGIN__\\n'; "
        "asterisk -rx 'database show DEVICE' 2>/dev/null; "
        "printf '\\n__VP_FAST_END__\\n'"
    )

    try:
        if local:
            result = subprocess.run(
                cmd, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                universal_newlines=True, timeout=45
            )
            salida = result.stdout or ''
            rc = result.returncode
            error = result.stderr or ''
        else:
            salida, error, rc = ejecutar_comando_ssh(perfil, cmd, timeout=45)

        if rc != 0 and not salida:
            return {}, str(error or 'SSH/CLI no devolvio datos')

        def seccion(marca, siguiente):
            token = '__VP_' + marca + '_BEGIN__'
            token_sig = '__VP_' + siguiente + '_BEGIN__' if siguiente else '__VP_FAST_END__'
            if token not in salida:
                return ''
            trozo = salida.split(token, 1)[1]
            if token_sig in trozo:
                trozo = trozo.split(token_sig, 1)[0]
            return trozo.strip()

        sip_conf = seccion('SIP_CONF', 'PJSIP_CONF')
        pjsip_conf = seccion('PJSIP_CONF', 'SIP_PEERS')
        sip_peers = seccion('SIP_PEERS', 'PJSIP_CONTACTS')
        pjsip_contacts = seccion('PJSIP_CONTACTS', 'DEVICE_DB')
        device_db = seccion('DEVICE_DB', None)

        datos = {}

        def asegurar(interno, tecnologia):
            interno = str(interno or '').strip()
            if not re.match(r'^[0-9]{1,10}$', interno):
                return None
            if interno not in datos:
                datos[interno] = {
                    'ip': '', 'mac': '', 'display_name': interno,
                    'tecnologia': tecnologia
                }
            return datos[interno]

        # CHAN_SIP: las secciones numericas de sip_additional.conf corresponden a peers.
        actual = None
        for linea in sip_conf.splitlines():
            linea = linea.strip()
            m = re.match(r'^\[([0-9]{1,10})\]', linea)
            if m:
                actual = m.group(1)
                asegurar(actual, 'SIP')
                continue
            if actual and '=' in linea:
                clave, _, valor = linea.partition('=')
                if clave.strip().lower() == 'callerid':
                    m = re.match(r'^\s*["\']?(.*?)\s*<\s*[0-9]+\s*>', valor)
                    if m:
                        nombre = m.group(1).strip().strip('"').strip("'")
                        if nombre and nombre.lower() != 'device':
                            datos[actual]['display_name'] = nombre

        # PJSIP: tomar solo secciones numericas cuyo tipo sea endpoint.
        for m in re.finditer(
            r'(?ms)^\[([0-9]{1,10})\]\s*(?:\([^\n]*\))?\s*\n(.*?)(?=^\[|\Z)',
            pjsip_conf
        ):
            interno = m.group(1)
            cuerpo = m.group(2)
            mt = re.search(r'(?im)^\s*type\s*=\s*endpoint\s*$', cuerpo)
            if not mt:
                continue
            item = asegurar(interno, 'PJSIP')
            mc = re.search(r'(?im)^\s*callerid\s*=\s*(.*?)\s*<[^>]+>', cuerpo)
            if mc:
                nombre = mc.group(1).strip().strip('"').strip("'")
                if nombre and nombre.lower() != 'device':
                    item['display_name'] = nombre

        # IPs SIP: una sola salida global de sip show peers.
        for linea in sip_peers.splitlines():
            partes = linea.split()
            if len(partes) < 2 or '/' not in partes[0]:
                continue
            interno = partes[0].split('/', 1)[0]
            ip = partes[1]
            if interno in datos and re.match(r'^\d{1,3}(?:\.\d{1,3}){3}$', ip):
                datos[interno]['ip'] = ip

        # Contactos PJSIP: leer la lista global; no ejecutar pjsip show aor por cada extension.
        for linea in pjsip_contacts.splitlines():
            m = re.search(
                r'(?i)\b([0-9]{1,10})/(?:[^\s]*?)(?:sip:)[^@\s]+@([0-9]{1,3}(?:\.\d{1,3}){3})(?::\d+)?',
                linea
            )
            if not m:
                m = re.search(
                    r'(?i)\bsip:([0-9]{1,10})@([0-9]{1,3}(?:\.\d{1,3}){3})(?::\d+)?',
                    linea
                )
            if m:
                interno, ip = m.group(1), m.group(2)
                if interno in datos:
                    datos[interno]['ip'] = ip

        # MAC registrada en AstDB, si existe.
        for linea in device_db.splitlines():
            m = re.match(r'^\s*/DEVICE/([0-9]{1,10})/mac\s*:\s*(\S+)', linea)
            if m and m.group(1) in datos:
                mac = re.sub(r'[^0-9a-f]', '', m.group(2).lower())
                if re.match(r'^[0-9a-f]{12}$', mac):
                    datos[m.group(1)]['mac'] = mac

        logger.info(
            'Consulta Asterisk rapida %s: %s internos, SSH rc=%s',
            host, len(datos), rc
        )
        if not datos:
            return {}, str(error or 'No se detectaron internos numericos en sip_additional.conf ni pjsip_additional.conf')
        return datos, ''

    except Exception as e:
        logger.exception('Consulta Asterisk rapida fallida para %s', host)
        return {}, str(e)


def _vp_excel_filas_desde_asterisk(perfil_nombre, ep=None, ips_dict=None, guardar=True, incluir_macs=True):
    """Reconstruye el inventario con una sola consulta SSH; el escaneo MAC es opcional."""
    nombre = str(perfil_nombre or '').strip()
    perfil = buscar_perfil(nombre)
    if not perfil:
        return []
    if ep is None:
        ep = get_excel_file(nombre)
    if guardar and os.environ.get('VOIP_PROVISIONER_REMOTE_EXCEL') == '1':
        _vp_excel_preparar_maestro_remoto(nombre, ep)
    ips_dict = ips_dict or {}

    datos, error_sync = _vp_excel_consultar_asterisk_rapido(perfil)
    if not isinstance(datos, dict):
        datos = {}

    # Buscar archivos por MAC solo para la reconstruccion enriquecida en background.
    mapa_mac = _vp_excel_mapa_mac_remoto(perfil) if incluir_macs else {}
    if not datos and mapa_mac:
        for interno in mapa_mac:
            if re.match(r'^[0-9]{1,10}$', str(interno)):
                datos[str(interno)] = {
                    'ip': '', 'mac': '', 'display_name': str(interno)
                }

    filas = []
    for interno in sorted(datos.keys(), key=lambda x: (len(x), x)):
        if not re.match(r'^[0-9]{1,10}$', str(interno)):
            continue
        item = datos[interno] or {}
        macs = list(mapa_mac.get(interno, [])) if incluir_macs else []
        if not macs:
            mac = re.sub(r'[^0-9a-f]', '', str(item.get('mac') or '').lower())
            if re.match(r'^[0-9a-f]{12}$', mac):
                macs = [mac]
        if not macs:
            macs = ['']

        ip = str(ips_dict.get(interno) or item.get('ip') or '').strip()
        if ip.lower() in ('dynamic', 'unknown', 'unavailable', '0.0.0.0'):
            ip = ''
        nombre_mostrar = str(item.get('display_name') or interno).strip()
        if not nombre_mostrar or nombre_mostrar.lower() == 'device':
            nombre_mostrar = interno

        for mac in macs:
            mac_norm = re.sub(r'[^0-9a-f]', '', str(mac or '').lower())
            filas.append({
                'edificio': '', 'piso': '', 'dependencia': '',
                'nombre_mostrar': nombre_mostrar,
                'nombre_saliente': interno,
                'interno': interno,
                'mac': mac_norm.upper() if re.match(r'^[0-9a-f]{12}$', mac_norm) else '',
                'ip': ip if len(macs) == 1 else '',
                'numero_serie': '', 'modelo': '',
                'fila_excel': len(filas) + 2
            })

    if not filas:
        logger.error('Consulta rapida no devolvio internos para %s: %s', nombre, error_sync)
        return []

    if guardar and not _vp_excel_tiene_registros(ep):
        import tempfile
        import shutil
        directorio = os.path.dirname(ep)
        if not os.path.isdir(directorio):
            os.makedirs(directorio, exist_ok=True)
        if os.path.isfile(ep) and os.environ.get('VOIP_PROVISIONER_REMOTE_EXCEL') != '1':
            try:
                shutil.copy2(ep, ep + '.bak_antes_reconstruccion_' + time.strftime('%Y%m%d_%H%M%S'))
            except Exception:
                logger.warning('No se pudo respaldar el Excel previo %s', ep)

        wb = Workbook()
        ws = wb.active
        ws.title = 'Registro'
        ws.append(COLUMNAS_EXCEL)
        for fila in filas:
            ws.append([
                fila.get('edificio', ''), fila.get('piso', ''), fila.get('dependencia', ''),
                fila.get('nombre_mostrar', ''), fila.get('nombre_saliente', ''),
                fila.get('interno', ''), fila.get('mac', ''), fila.get('ip', ''),
                fila.get('numero_serie', ''), fila.get('modelo', '')
            ])
        fd, temporal = tempfile.mkstemp(
            prefix='.vp_asterisk_excel_', suffix='.xlsx', dir=directorio
        )
        os.close(fd)
        try:
            aplicar_formato_excel(wb, ws)
            wb.save(temporal)
            wb.close()
            if not _vp_excel_tiene_registros(temporal):
                raise RuntimeError('La planilla reconstruida no paso la validacion.')
            os.replace(temporal, ep)
        except Exception:
            try:
                wb.close()
            except Exception:
                pass
            raise
        finally:
            if os.path.isfile(temporal):
                try:
                    os.remove(temporal)
                except Exception:
                    pass

        logger.info('Excel central %s reconstruido desde Asterisk: %s filas', nombre, len(filas))
        try:
            if incluir_macs or os.environ.get('VOIP_PROVISIONER_REMOTE_EXCEL') == '1':
                subir_excel_remoto(
                    nombre, perfil.get('ruta_remota', '/tftpboot'),
                    source_path=ep,
                    backup=(os.environ.get('VOIP_PROVISIONER_REMOTE_EXCEL') == '1'),
                    raise_on_error=(os.environ.get('VOIP_PROVISIONER_REMOTE_EXCEL') == '1')
                )
        except Exception:
            logger.warning('Excel reconstruido, pero no se pudo guardar el maestro remoto.', exc_info=True)
            if os.environ.get('VOIP_PROVISIONER_REMOTE_EXCEL') == '1':
                raise

    return filas


_VP_EXCEL_REFRESH_EN_CURSO = set()
_VP_EXCEL_REFRESH_LOCK = threading.Lock()


def _vp_excel_iniciar_actualizacion(perfil_nombre, forzar=False, ep_override=None):
    """Inicia el refresco del perfil en background y regresa sin bloquear HTTP."""
    nombre = str(perfil_nombre or '').strip()
    if not nombre or not buscar_perfil(nombre):
        return False

    ep = ep_override or get_excel_file(nombre)
    ahora = time.time()
    with _VP_EXCEL_REFRESH_LOCK:
        if nombre in _VP_EXCEL_REFRESH_EN_CURSO:
            return False
        if (not forzar and _vp_excel_tiene_registros(ep) and
                ahora - _VP_EXCEL_ULTIMA_ACTUALIZACION.get(nombre, 0) < _VP_EXCEL_TTL_SEGUNDOS):
            return False
        _VP_EXCEL_REFRESH_EN_CURSO.add(nombre)

    def trabajador():
        try:
            _vp_actualizar_excel_perfil(nombre, forzar=forzar, ep_override=ep)
        except Exception:
            logger.exception('Actualizacion de Excel en background fallo para %s', nombre)
        finally:
            with _VP_EXCEL_REFRESH_LOCK:
                _VP_EXCEL_REFRESH_EN_CURSO.discard(nombre)

    try:
        threading.Thread(target=trabajador, daemon=True).start()
        return True
    except Exception:
        with _VP_EXCEL_REFRESH_LOCK:
            _VP_EXCEL_REFRESH_EN_CURSO.discard(nombre)
        logger.exception('No se pudo iniciar refresco background para %s', nombre)
        return False


def _vp_excel_actualizar_al_conectar(perfil_nombre):
    ep = get_excel_file(perfil_nombre)
    if os.environ.get('VOIP_PROVISIONER_REMOTE_EXCEL') == '1':
        _vp_excel_preparar_maestro_remoto(perfil_nombre, ep)
    if not _vp_excel_tiene_registros(ep):
        # Respuesta rapida desde Asterisk; la busqueda de archivos MAC va aparte.
        _vp_excel_filas_desde_asterisk(
            perfil_nombre, ep=ep, guardar=True, incluir_macs=False
        )
        if _vp_excel_tiene_registros(ep):
            _vp_excel_iniciar_actualizacion(
                perfil_nombre, forzar=True, ep_override=ep
            )
            return True
        return _vp_actualizar_excel_perfil(
            perfil_nombre, forzar=True, ep_override=ep
        )
    _vp_excel_actualizar_excel = _vp_excel_iniciar_actualizacion(
        perfil_nombre, forzar=False, ep_override=ep
    )
    return _vp_excel_tiene_registros(ep)

def _vp_borrar_configuraciones_mac_sin_excel(perfil, ruta_remota, macs):
    import shlex
    host = str(perfil.get('host', '')).strip()
    rr = str(ruta_remota or perfil.get('ruta_remota', '/tftpboot') or '/tftpboot').rstrip('/')
    local = _vp_host_es_local(host)
    errores = []
    rutas_por_mac = {}

    for mac in sorted(set(macs)):
        archivos_locales = [
            os.path.join(CARPETA_CONFIGS, mac + '.cfg'),
            os.path.join(CARPETA_CONFIGS, 'cfg' + mac + '.xml'),
            os.path.join(CARPETA_CONFIGS, mac + '.xml')
        ]
        rutas_remotas = [
            rr + '/' + mac + '.cfg',
            rr + '/cfg' + mac + '.xml',
            rr + '/' + mac + '.xml'
        ]
        for ruta in archivos_locales:
            try:
                if os.path.isfile(ruta):
                    os.remove(ruta)
            except Exception as e:
                errores.append('local ' + os.path.basename(ruta) + ': ' + str(e))
        if local:
            for ruta in rutas_remotas:
                try:
                    if os.path.isfile(ruta):
                        os.remove(ruta)
                except Exception as e:
                    errores.append(ruta + ': ' + str(e))
        else:
            cmd = 'rm -f ' + ' '.join(shlex.quote(ruta) for ruta in rutas_remotas)
            try:
                out, err, rc = ejecutar_comando_ssh(perfil, cmd)
                if rc != 0:
                    errores.append('remoto ' + mac + ': ' + str(err or out or rc))
            except Exception as e:
                errores.append('remoto ' + mac + ': ' + str(e))
        rutas_por_mac[mac] = rutas_remotas

    return errores


def _vp_borrar_sin_excel(nombre_perfil, perfil, seleccion, ruta_remota, tarea_id):
    """Ejecuta bajas verificando Asterisk y configs; no necesita libro Excel."""
    try:
        if not perfil:
            error_tarea(tarea_id, 'Perfil no encontrado; no se elimino nada.')
            return

        normalizados = []
        for item in seleccion:
            interno = str(item.get('interno') or '').strip()
            mac = re.sub(r'[^0-9a-f]', '', str(item.get('mac') or '').lower())
            if not re.match(r'^[0-9]{1,10}$', interno):
                error_tarea(tarea_id, 'Interno invalido; no se elimino nada.')
                return
            if mac and not re.match(r'^[0-9a-f]{12}$', mac):
                error_tarea(tarea_id, 'MAC invalida; no se elimino nada.')
                return
            normalizados.append({'interno': interno, 'mac': mac})

        if not normalizados:
            error_tarea(tarea_id, 'No hay internos seleccionados.')
            return

        actualizar_progreso(tarea_id, 10, 'Consultando Asterisk; no se requiere Excel...')
        try:
            inventario, error_sync = _vp_excel_consultar_asterisk_rapido(perfil)
            if not isinstance(inventario, dict):
                inventario = {}
        except Exception as e:
            inventario, error_sync = {}, str(e)

        mapa = _vp_excel_mapa_mac_remoto(perfil)
        por_interno = {}
        for item in normalizados:
            por_interno.setdefault(item['interno'], set())
            if item['mac']:
                por_interno[item['interno']].add(item['mac'])

        # Validar todos los internos antes de borrar configuraciones o MySQL.
        for interno in sorted(por_interno):
            existe = interno in inventario or interno in mapa
            if not existe:
                try:
                    info = obtener_password_interno_issabel(perfil, interno)
                    existe = bool(isinstance(info, dict) and info.get('existe'))
                except Exception:
                    existe = False
            if not existe:
                error_tarea(
                    tarea_id,
                    'No se pudo confirmar que el interno {} exista en Asterisk. '
                    'No se borro nada. Verifique SSH/CLI. {}'.format(interno, str(error_sync or '')[:100])
                )
                logger.error('Baja sin Excel cancelada: no se confirmo %s en %s', interno, nombre_perfil)
                return

        # Evitar eliminar el interno compartido si quedan MAC/configs no seleccionadas.
        restantes_por_interno = {}
        eliminar_por_interno = {}
        for interno, macs_seleccionadas in por_interno.items():
            conocidas = set(mapa.get(interno, []))
            if macs_seleccionadas:
                quitar = set(macs_seleccionadas)
                restantes = conocidas - quitar
                eliminar_por_interno[interno] = not bool(restantes)
                restantes_por_interno[interno] = restantes
            else:
                if len(conocidas) > 1:
                    error_tarea(
                        tarea_id,
                        'El interno {} tiene varios telefonos remotos y la seleccion no trae MAC. '
                        'No se borro nada por seguridad.'.format(interno)
                    )
                    return
                eliminar_por_interno[interno] = True
                restantes_por_interno[interno] = set()

        # Si se borra un interno completo sin MAC seleccionada, retirar la unica config conocida.
        macs_quitar = set()
        for interno, macs_seleccionadas in por_interno.items():
            if macs_seleccionadas:
                macs_quitar.update(macs_seleccionadas)
            else:
                macs_quitar.update(mapa.get(interno, []))

        actualizar_progreso(tarea_id, 35, 'Eliminando archivos de aprovisionamiento seleccionados...')
        errores_cfg = _vp_borrar_configuraciones_mac_sin_excel(perfil, ruta_remota, macs_quitar)

        actualizar_progreso(tarea_id, 65, 'Eliminando o conservando internos en Asterisk...')
        eliminados = []
        conservados = []
        fallidos = []
        for interno in sorted(por_interno):
            if eliminar_por_interno.get(interno, True):
                try:
                    ok, mensaje = eliminar_interno_asterisk(perfil, interno)
                except Exception as e:
                    ok, mensaje = False, str(e)
                if ok:
                    eliminados.append(interno)
                else:
                    fallidos.append('{}: {}'.format(interno, mensaje))
            else:
                conservados.append(interno)
                logger.info(
                    'Baja sin Excel: se conserva %s porque quedan MAC/configs remotas: %s',
                    interno, ','.join(sorted(restantes_por_interno.get(interno, set()))))

        actualizar_progreso(tarea_id, 90, 'Aplicando cambios remotos...')
        if eliminados:
            try:
                recargar_issabel_completo(perfil)
            except Exception:
                logger.warning('No se pudo recargar Asterisk al terminar la baja sin Excel.', exc_info=True)

        if fallidos:
            error_tarea(tarea_id, 'Baja parcial. Errores Asterisk: ' + '; '.join(fallidos)[:300])
        else:
            detalles = []
            if eliminados:
                detalles.append('{} interno(s) eliminados desde Asterisk'.format(len(eliminados)))
            if conservados:
                detalles.append('{} interno(s) conservados porque tienen otro telefono'.format(len(conservados)))
            detalles.append('Excel no disponible; no se uso ni subio una planilla vacia')
            if errores_cfg:
                detalles.append('{} aviso(s) al borrar archivos de configuracion'.format(len(errores_cfg)))
            completar_tarea(tarea_id, '; '.join(detalles))

    except Exception as e:
        logger.exception('Error en baja independiente del Excel (%s)', nombre_perfil)
        error_tarea(tarea_id, 'Error en baja sin Excel: {}'.format(str(e)[:180]))

# === FIN VP ASTERISK PRIMERO Y BORRADO SIN EXCEL V1 ===

@main_bp.route('/')
def index():
    return render_template('index.html')

@main_bp.route('/provisionamiento')
def provisionamiento():
    return render_template('index.html')

@main_bp.route('/admin')
def admin():
    return render_template('admin.html')

@main_bp.route('/monitoreo')
def monitoreo():
    return render_template('monitoreo.html')


# ===== BLOQUE 1 - API CDR PAGINADO =====

# ===== CRITERIO DE TRONCAL PARA EL CDR =====
#
# En el CDR, el canal de un interno es <TECNOLOGIA>/<numero>-<hex>
# (PJSIP/1201-0000001d). Una troncal es cualquier canal que NO tenga
# esa forma: PJSIP/synway-e1-..., SIP/movistar-..., IAX2/proveedor-...,
# mas las tecnologias que solo se usan para troncales.

_RE_INTERNO_SQL = "^(pjsip|sip|iax2)/[0-9]+-[0-9a-f]+$"
_RE_TRONCAL_NOMBRADA_SQL = "^(pjsip|sip|iax2)/[^/]+-[0-9a-f]+$"
_TECNOLOGIAS_TRONCAL = (
    'ooh323/', 'dahdi/', 'zap/', 'dongle/', 'custom/'
)


def _sql_canal_es_troncal(columna):
    c = "LOWER({0})".format(columna)
    partes = ["{0} LIKE '{1}%'".format(c, t) for t in _TECNOLOGIAS_TRONCAL]
    partes.append(
        "({0} REGEXP '{1}' AND NOT {0} REGEXP '{2}')".format(
            c, _RE_TRONCAL_NOMBRADA_SQL, _RE_INTERNO_SQL
        )
    )
    return "(" + " OR ".join(partes) + ")"


def _sql_llamada_recibida():
    # Entra por una troncal (CHANNEL) y no sale por otra (DSTCHANNEL).
    return "{0} AND NOT {1}".format(
        _sql_canal_es_troncal("channel"),
        _sql_canal_es_troncal("dstchannel")
    )


def _canal_es_troncal(canal):
    c = str(canal or '').strip().lower()
    if not c:
        return False
    if c.startswith(_TECNOLOGIAS_TRONCAL):
        return True
    return bool(
        re.match(_RE_TRONCAL_NOMBRADA_SQL, c)
        and not re.match(_RE_INTERNO_SQL, c)
    )


def _ejecutar_sql_cdr(perfil, sql):
    """
    Ejecuta SQL contra asteriskcdrdb.

    Local:
        ejecuta directamente.

    Remoto:
        ejecuta mediante SSH en el Issabel remoto.

    Las credenciales se obtienen del propio servidor remoto.
    """

    try:

        cred = obtener_credenciales_mysql(
            perfil
        )

        if not cred:
            return (
                "",
                "No se pudieron obtener las credenciales de MariaDB",
                1
            )

        if isinstance(
            cred,
            (tuple, list)
        ):

            if len(cred) < 2:
                return (
                    "",
                    "Credenciales MariaDB incompletas",
                    1
                )

            usuario = cred[0]
            password = cred[1]

        else:

            usuario = (
                cred.get("usuario")
                or cred.get("user")
                or cred.get("AMPDBUSER")
            )

            password = (
                cred.get("password")
                or cred.get("pass")
                or cred.get("AMPDBPASS")
            )

        if (
            not usuario
            or password is None
        ):
            return (
                "",
                "Credenciales MariaDB incompletas",
                1
            )

        import shlex

        user_q = shlex.quote(
            str(usuario)
        )

        pass_q = shlex.quote(
            str(password)
        )

        sql_q = shlex.quote(
            str(sql)
        )

        comando = (
            "if [ -x /usr/bin/mysql ]; then "
            "MYSQL=/usr/bin/mysql; "
            "elif command -v mysql >/dev/null 2>&1; then "
            "MYSQL=$(command -v mysql); "
            "else "
            "echo 'ERROR: mysql no encontrado' >&2; "
            "exit 127; "
            "fi; "
            "$MYSQL -N -B "
            "-u {0} "
            "-p{1} "
            "asteriskcdrdb "
            "-e {2}"
        ).format(
            user_q,
            pass_q,
            sql_q
        )

        resultado = _ejecutar_comando_monitor_cdr(
            perfil,
            comando
        )

        if isinstance(
            resultado,
            tuple
        ):

            stdout = (
                resultado[0]
                if len(resultado) > 0
                else ""
            )

            stderr = (
                resultado[1]
                if len(resultado) > 1
                else ""
            )

            rc = (
                resultado[2]
                if len(resultado) > 2
                else 0
            )

        else:

            stdout = resultado or ""
            stderr = ""
            rc = 0

        if stdout is None:
            stdout = ""

        if stderr is None:
            stderr = ""

        if rc not in (
            0,
            None
        ):

            mensaje = str(
                stderr
            ).strip()

            if not mensaje:
                mensaje = (
                    "MariaDB devolvio "
                    "codigo %s" % rc
                )

            logger.error(
                "CDR SQL error perfil=%s host=%s: %s",
                perfil.get("nombre", ""),
                perfil.get("host", ""),
                mensaje
            )

            return (
                stdout,
                mensaje,
                rc
            )

        if (
            stderr
            and (
                "ERROR" in stderr.upper()
                or "ACCESS DENIED" in stderr.upper()
                or "UNKNOWN DATABASE" in stderr.upper()
                or "NOT FOUND" in stderr.upper()
            )
        ):

            logger.error(
                "CDR MariaDB stderr perfil=%s host=%s: %s",
                perfil.get("nombre", ""),
                perfil.get("host", ""),
                stderr.strip()
            )

            return (
                stdout,
                stderr,
                1
            )

        return (
            stdout,
            stderr,
            0
        )

    except Exception as e:

        logger.exception(
            "Error ejecutando SQL CDR"
        )

        return (
            "",
            str(e),
            1
        )

def _ejecutar_comando_monitor_cdr(perfil, comando):
    """
    Ejecuta localmente si el perfil apunta al Issabel local.
    En perfiles remotos utiliza SSH.
    """
    try:
        host = str(perfil.get('host', '')).strip()

        if _vp_host_es_local(host):
            p = subprocess.Popen(
                comando,
                shell=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE
            )

            out, err = p.communicate(timeout=30)

            return (
                out.decode('utf-8', 'replace'),
                err.decode('utf-8', 'replace'),
                p.returncode
            )

        return ejecutar_comando_ssh(perfil, comando)

    except Exception as e:
        return "", str(e), 1


def _normalizar_estado_cdr(disposition):
    """
    Normaliza los estados de Asterisk/Issabel para que el frontend
    pueda trabajar con tres categorías principales.
    """
    estado = str(disposition or '').strip().upper()

    if estado == 'ANSWERED':
        return 'ANSWERED'

    if estado in (
        'BUSY',
        'FAILED',
        'CONGESTION',
        'CHANUNAVAIL',
        'DONTCALL',
        'TORTURE'
    ):
        return 'REJECTED'

    if estado in (
        'NO ANSWER',
        'NOANSWER',
        'CANCEL'
    ):
        return 'NO ANSWER'

    return estado or 'UNKNOWN'




# ===== CDR - BORRADO PROTEGIDO =====

def _verificar_password_admin_cdr(data):
    """
    Valida la misma contraseña de administración utilizada
    por admin.html y por las demás operaciones protegidas.
    """
    try:
        cfg = cargar_config()

        password_admin = (
            cfg.get('seguridad', {})
               .get('admin_password', '')
        )

        password_recibida = ''

        if isinstance(data, dict):
            password_recibida = str(
                data.get('password', '')
            )

        return bool(
            password_admin
            and password_recibida == str(password_admin)
        )

    except Exception as e:
        logger.error(
            "Error verificando contraseña CDR: %s",
            e
        )
        return False


@main_bp.route('/api/monitoreo/cdr/eliminar', methods=['POST'])
def api_monitoreo_cdr_eliminar():
    """
    Elimina un único registro CDR utilizando su uniqueid.

    Requiere la contraseña de administración.
    """

    try:
        data = request.get_json(
            silent=True
        ) or {}

        if not _verificar_password_admin_cdr(data):
            return jsonify({
                'status': 'error',
                'message': 'Contraseña de administrador incorrecta'
            }), 401

        perfil_nombre = str(
            data.get('perfil_nombre', '')
        ).strip()

        uniqueid = str(
            data.get('uniqueid', '')
        ).strip()

        if not perfil_nombre:
            return jsonify({
                'status': 'error',
                'message': 'Falta perfil_nombre'
            }), 400

        if not uniqueid:
            return jsonify({
                'status': 'error',
                'message': 'Falta uniqueid'
            }), 400

        perfil = buscar_perfil(
            perfil_nombre,
            cargar_perfiles()
        )

        if not perfil:
            return jsonify({
                'status': 'error',
                'message': 'Perfil no encontrado'
            }), 404

        # Escapado SQL.
        uniqueid_sql = (
            uniqueid
            .replace('\\', '\\\\')
            .replace("'", "''")
        )

        sql = (
            "DELETE FROM cdr "
            "WHERE uniqueid = '{0}' "
            "LIMIT 1"
        ).format(
            uniqueid_sql
        )

        out, err, code = _ejecutar_sql_cdr(
            perfil,
            sql
        )

        if code != 0:
            logger.error(
                "Error eliminando CDR uniqueid=%s: %s",
                uniqueid,
                err
            )

            return jsonify({
                'status': 'error',
                'message': 'No se pudo eliminar el registro CDR',
                'detail': err
            }), 500

        return jsonify({
            'status': 'ok',
            'message': 'Registro CDR eliminado'
        })

    except Exception as e:
        logger.exception(
            "Error eliminando CDR"
        )

        return jsonify({
            'status': 'error',
            'message': 'Error eliminando registro CDR',
            'detail': str(e)
        }), 500


@main_bp.route('/api/monitoreo/cdr/limpiar', methods=['POST'])
def api_monitoreo_cdr_limpiar():
    """
    Elimina todos los registros de la tabla CDR.

    Requiere la contraseña de administración.
    """

    try:
        data = request.get_json(
            silent=True
        ) or {}

        if not _verificar_password_admin_cdr(data):
            return jsonify({
                'status': 'error',
                'message': 'Contraseña de administrador incorrecta'
            }), 401

        perfil_nombre = str(
            data.get('perfil_nombre', '')
        ).strip()

        if not perfil_nombre:
            return jsonify({
                'status': 'error',
                'message': 'Falta perfil_nombre'
            }), 400

        perfil = buscar_perfil(
            perfil_nombre,
            cargar_perfiles()
        )

        if not perfil:
            return jsonify({
                'status': 'error',
                'message': 'Perfil no encontrado'
            }), 404

        sql = "DELETE FROM cdr"

        out, err, code = _ejecutar_sql_cdr(
            perfil,
            sql
        )

        if code != 0:
            logger.error(
                "Error limpiando CDR: %s",
                err
            )

            return jsonify({
                'status': 'error',
                'message': 'No se pudo limpiar el historial CDR',
                'detail': err
            }), 500

        return jsonify({
            'status': 'ok',
            'message': 'Historial CDR limpiado correctamente'
        })

    except Exception as e:
        logger.exception(
            "Error limpiando CDR"
        )

        return jsonify({
            'status': 'error',
            'message': 'Error limpiando historial CDR',
            'detail': str(e)
        }), 500


@main_bp.route('/api/monitoreo/cdr', methods=['GET'])
def api_monitoreo_cdr():
    """
    Historial CDR paginado y filtrable.

    Parámetros:

      perfil_nombre = perfil de Issabel

      pagina       = número de página, comienza en 1
      por_pagina   = registros por página

      buscar       = busca simultáneamente en src y dst

      estado       = ALL / ANSWERED / NO ANSWER / REJECTED

      desde        = YYYY-MM-DD
      hasta        = YYYY-MM-DD

    Devuelve:

      registros
      paginacion
      estadisticas
    """

    try:
        perfil_nombre = request.args.get(
            'perfil_nombre',
            ''
        ).strip()

        if not perfil_nombre:
            return jsonify({
                'status': 'error',
                'message': 'Falta perfil_nombre'
            }), 400

        perfil = buscar_perfil(
            perfil_nombre,
            cargar_perfiles()
        )

        if not perfil:
            return jsonify({
                'status': 'error',
                'message': 'Perfil no encontrado'
            }), 404

        try:
            pagina = int(
                request.args.get('pagina', '1')
            )
        except Exception:
            pagina = 1

        try:
            por_pagina = int(
                request.args.get('por_pagina', '10')
            )
        except Exception:
            por_pagina = 10

        pagina = max(1, pagina)
        por_pagina = max(1, min(100, por_pagina))

        buscar = str(
            request.args.get('buscar', '')
        ).strip()

        estado_filtro = str(
            request.args.get('estado', 'ALL')
        ).strip().upper()

        # Filtros especiales de llamadas realizadas.
        #
        # REALIZADAS:
        #   llamadas originadas por un interno.
        #
        # REALIZADAS_EXTERIOR:
        #   llamadas originadas por un interno que utilizaron
        #   una troncal/canal externo.

        desde = str(
            request.args.get('desde', '')
        ).strip()

        hasta = str(
            request.args.get('hasta', '')
        ).strip()

        # Validar fechas básicas para evitar introducir SQL.
        if desde and not re.match(
            r'^\d{4}-\d{2}-\d{2}$',
            desde
        ):
            return jsonify({
                'status': 'error',
                'message': 'Fecha desde inválida'
            }), 400

        if hasta and not re.match(
            r'^\d{4}-\d{2}-\d{2}$',
            hasta
        ):
            return jsonify({
                'status': 'error',
                'message': 'Fecha hasta inválida'
            }), 400

        # Escapado SQL.
        def sql_escape(value):
            return str(value).replace(
                '\\\\',
                '\\\\\\\\'
            ).replace(
                "'",
                "''"
            )

        condiciones = []

        if desde:
            condiciones.append(
                "calldate >= '{0} 00:00:00'".format(
                    sql_escape(desde)
                )
            )

        if hasta:
            condiciones.append(
                "calldate <= '{0} 23:59:59'".format(
                    sql_escape(hasta)
                )
            )

        if buscar:
            b = sql_escape(buscar)

            condiciones.append(
                "(CAST(src AS CHAR) LIKE '%{0}%' "
                "OR CAST(dst AS CHAR) LIKE '%{0}%')".format(b)
            )

        where = ""

        if condiciones:
            where = " WHERE " + " AND ".join(condiciones)

        # Para el filtro de estado usamos la misma clasificación
        # que devuelve el endpoint.
        estado_sql = ""

        if estado_filtro == 'ANSWERED':
            estado_sql = " AND disposition = 'ANSWERED'"

        elif estado_filtro == 'NO ANSWER':
            estado_sql = (
                " AND disposition IN "
                "('NO ANSWER','NOANSWER','CANCEL')"
            )

        elif estado_filtro == 'REJECTED':
            estado_sql = (
                " AND disposition IN "
                "('BUSY','FAILED','CONGESTION',"
                "'CHANUNAVAIL','DONTCALL','TORTURE')"
            )

        elif estado_filtro in (
            'REALIZADAS',
            'REALIZADA'
        ):
            # Llamadas realizadas:
            #
            # Cualquier CDR cuyo origen (src) sea numérico.
            # No dependemos de dcontext ni del tipo de troncal,
            # porque una llamada interna o externa puede viajar
            # por OOH323/IAX2/etc.
            estado_sql = (
                " AND CAST(src AS CHAR) REGEXP '^[0-9]+$'"
            )

        elif estado_filtro in (
            'RECIBIDAS',
            'RECIBIDA',
            'ENTRANTES',
            'ENTRANTE'
        ):
            # ==================================================
            # LLAMADAS RECIBIDAS
            # ==================================================
            #
            # La llamada externa entra al PBX por el CHANNEL.
            #
            # Una salida por una troncal PJSIP/Synway, en cambio,
            # tiene normalmente la troncal en DSTCHANNEL.
            #
            # Por eso NO buscamos simplemente "PJSIP":
            #
            #   interno -> Synway
            #       channel    = PJSIP/57364-...
            #       dstchannel = PJSIP/synway-e1-...
            #
            #   Synway -> interno
            #       channel    = PJSIP/synway-e1-...
            #       dstchannel = PJSIP/57364-...
            #
            # Se contemplan los tipos de troncal conocidos por
            # el Monitor y los nombres habituales de Synway/E1.
            #
            estado_sql = " AND (" + _sql_llamada_recibida() + ")"

        elif estado_filtro in (
            'REALIZADAS_EXTERIOR',
            'EXTERIOR',
            'EXTERNAS'
        ):
            # Llamada externa realizada:
            #
            #   - originada desde from-internal
            #   - destino comienza con 9
            #
            # Regla definitiva:
            #   9 = salida a línea externa.
            #
            # Los números cortos (100, 101, 113, 144, 147, etc.)
            # NO se consideran externos.
            #
            # No dependemos del tipo de troncal, ya que OOH323/IAX2
            # también pueden transportar llamadas internas.

            estado_sql = (
                " AND CAST(src AS CHAR) REGEXP '^[0-9]+$'"
                " AND CAST(dst AS CHAR) LIKE '9%'"
            )

        elif estado_filtro not in (
            '',
            'ALL',
            'TODOS'
        ):
            estado_filtro = 'ALL'

        # Combinar correctamente los filtros.
        # Si ya existe WHERE, el estado se agrega con AND.
        # Si no existe ningún otro filtro, el estado debe comenzar con WHERE.
        if estado_sql:
            if where:
                where_filtrado = where + estado_sql
            else:
                where_filtrado = " WHERE " + estado_sql[5:]
        else:
            where_filtrado = where

        offset = (pagina - 1) * por_pagina

        # Total de registros filtrados.
        sql_total = (
            "SELECT COUNT(*) "
            "FROM cdr"
            + where_filtrado
        )

        out, err, code = _ejecutar_sql_cdr(
            perfil,
            sql_total
        )

        if code != 0:
            logger.error(
                "Error contando CDR: %s",
                err
            )

            return jsonify({
                'status': 'error',
                'message': 'No se pudo consultar el historial CDR',
                'detail': err
            }), 500

        try:
            total = int(
                (out or '0').strip().splitlines()[0]
            )
        except Exception:
            total = 0

        # Estadísticas generales sobre el mismo período/búsqueda,
        # pero sin aplicar el filtro de estado.
        sql_stats = (
            "SELECT "
            "SUM(CASE WHEN disposition='ANSWERED' "
            "THEN 1 ELSE 0 END), "
            "SUM(CASE WHEN disposition IN "
            "('NO ANSWER','NOANSWER','CANCEL') "
            "THEN 1 ELSE 0 END), "
            "SUM(CASE WHEN disposition IN "
            "('BUSY','FAILED','CONGESTION','CHANUNAVAIL',"
            "'DONTCALL','TORTURE') "
            "THEN 1 ELSE 0 END), "
            "SUM(CASE WHEN CAST(src AS CHAR) "
            "REGEXP '^[0-9]+$' "
            "THEN 1 ELSE 0 END), "
            "SUM(CASE WHEN CAST(src AS CHAR) REGEXP '^[0-9]+$' "
            "AND CAST(dst AS CHAR) LIKE '9%' "
            "THEN 1 ELSE 0 END), "

            # RECIBIDAS: entra por troncal y no sale por otra.
            "SUM(CASE WHEN (" + _sql_llamada_recibida() + ") "
            "THEN 1 ELSE 0 END), "

            "COUNT(*) "
            "FROM cdr"
            + where
        )

        out_stats, err_stats, code_stats = _ejecutar_sql_cdr(
            perfil,
            sql_stats
        )

        contestadas = 0
        perdidas = 0
        rechazadas = 0
        realizadas = 0
        realizadas_exterior = 0
        recibidas = 0
        total_estadisticas = 0

        if code_stats == 0 and out_stats:
            try:
                valores = (
                    out_stats.strip()
                    .splitlines()[0]
                    .split('\t')
                )

                if len(valores) >= 7:
                    contestadas = int(
                        valores[0] or 0
                    )

                    perdidas = int(
                        valores[1] or 0
                    )

                    rechazadas = int(
                        valores[2] or 0
                    )

                    realizadas = int(
                        valores[3] or 0
                    )

                    realizadas_exterior = int(
                        valores[4] or 0
                    )

                    recibidas = int(
                        valores[5] or 0
                    )

                    total_estadisticas = int(
                        valores[6] or 0
                    )

            except Exception as e:
                logger.error(
                    "Error procesando estadísticas CDR: %s",
                    e
                )

        # Registros paginados.
        sql_registros = (
            "SELECT "
            "DATE_FORMAT(calldate,'%Y-%m-%d %H:%i:%s'),"
            "src,"
            "dst,"
            "duration,"
            "billsec,"
            "disposition,"
            "channel,"
            "dstchannel,"
            "uniqueid "
            "FROM cdr"
            + where_filtrado +
            " ORDER BY calldate DESC "
            "LIMIT {0} OFFSET {1}"
        ).format(
            int(por_pagina),
            int(offset)
        )

        out, err, code = _ejecutar_sql_cdr(
            perfil,
            sql_registros
        )

        if code != 0:
            logger.error(
                "Error obteniendo CDR: %s",
                err
            )

            return jsonify({
                'status': 'error',
                'message': 'No se pudo obtener el historial CDR',
                'detail': err
            }), 500

        registros = []

        for linea in (out or '').splitlines():

            partes = linea.split('\t')

            if len(partes) < 9:
                continue

            fecha = partes[0]
            origen = partes[1]
            destino = partes[2]

            try:
                duracion = int(
                    float(partes[3] or 0)
                )
            except Exception:
                duracion = 0

            try:
                billsec = int(
                    float(partes[4] or 0)
                )
            except Exception:
                billsec = 0

            disposition = partes[5]
            channel = partes[6]
            dstchannel = partes[7]
            uniqueid = partes[8]

            estado = _normalizar_estado_cdr(
                disposition
            )

            # Texto amigable para el frontend.
            if estado == 'ANSWERED':
                estado_texto = 'Contestada'

            elif estado == 'NO ANSWER':
                estado_texto = 'Perdida'

            elif estado == 'REJECTED':
                estado_texto = 'Rechazada'

            else:
                estado_texto = disposition or 'Desconocido'

            origen_limpio = str(
                origen or ''
            ).strip()

            canal_limpio = str(
                channel or ''
            ).strip().lower()

            dstcanal_limpio = str(
                dstchannel or ''
            ).strip().lower()

            origen_numerico = re.sub(
                r'[^0-9]',
                '',
                origen_limpio
            )

            es_realizada = bool(
                re.match(
                    r'^[0-9]{1,5}$',
                    origen_numerico
                )
            )

            # Identificadores habituales de canales/troncales
            # externos en Asterisk/Issabel.
            es_troncal = any(
                x in canal_limpio or
                x in dstcanal_limpio
                for x in (
                    'ooh323/',
                    'ooh323',
                    'ipoffice',
                    'iax2/',
                    'iax2',
                    'dahdi/',
                    'zap/',
                    'custom/',
                    'dongle/',
                    'sip/trunk',
                    'pjsip/trunk',
                    'trunk/'
                )
            )

            es_realizada_exterior = (
                es_realizada and
                es_troncal
            )

            # ==================================================
            # DETERMINAR LLAMADA RECIBIDA
            # ==================================================
            #
            # Para una llamada entrante la troncal aparece en
            # CHANNEL.
            #
            # Para una llamada saliente hacia una troncal,
            # aparece normalmente en DSTCHANNEL.
            #
            # Esto es especialmente importante para PJSIP /
            # Synway-E1.
            #
            es_recibida = (
                _canal_es_troncal(channel) and
                not _canal_es_troncal(dstchannel)
            )

            registros.append({
                'fecha': fecha,
                'hora': fecha[11:] if len(fecha) >= 19 else fecha,
                'origen': origen or '-',
                'destino': destino or '-',
                'duracion': duracion,
                'billsec': billsec,
                'estado': estado,
                'estado_original': disposition,
                'estado_texto': estado_texto,
                'uniqueid': uniqueid,
                'channel': channel,
                'dstchannel': dstchannel,
                'realizada': es_realizada,
                'realizada_exterior': es_realizada_exterior,
                'recibida': es_recibida
            })

        paginas = (
            (total + por_pagina - 1)
            // por_pagina
            if total
            else 1
        )

        return jsonify({
            'status': 'ok',

            'registros': registros,

            'paginacion': {
                'pagina': pagina,
                'por_pagina': por_pagina,
                'total': total,
                'paginas': paginas,
                'hay_anterior': pagina > 1,
                'hay_siguiente': pagina < paginas
            },

            'estadisticas': {
                'total': total_estadisticas,
                'contestadas': contestadas,
                'perdidas': perdidas,
                'rechazadas': rechazadas,
                'realizadas': realizadas,
                'realizadas_exterior': realizadas_exterior,
                'recibidas': recibidas
            },

            'filtros': {
                'buscar': buscar,
                'estado': estado_filtro,
                'desde': desde,
                'hasta': hasta
            }
        })

    except Exception as e:
        logger.exception(
            "Error en api_monitoreo_cdr"
        )

        return jsonify({
            'status': 'error',
            'message': str(e)
        }), 500


@main_bp.route('/api/opciones', methods=['GET', 'POST'])
def api_opciones():
    try:
        if request.method == 'POST':
            d = request.get_json(silent=True) or {}
            pn = str(d.get('perfil_nombre', '')).strip()
            campo = str(d.get('campo', '')).strip()
            valor = str(d.get('valor', '')).strip()
            if not pn or campo not in ('edificio', 'piso', 'dependencia') or not valor:
                return jsonify({'status':'error','message':'Perfil, campo y valor son obligatorios'}), 400
            if not buscar_perfil(pn):
                return jsonify({'status':'error','message':'Perfil no encontrado'}), 404
            data = cargar_ubicaciones_perfiles()
            perfil = data.setdefault(pn, {'edificio':[], 'piso':[], 'dependencia':[]})
            lista = perfil.setdefault(campo, [])
            if not any(str(x).strip().casefold() == valor.casefold() for x in lista):
                lista.append(valor)
                lista.sort(key=lambda x: str(x).casefold())
                if not guardar_ubicaciones_perfiles(data):
                    return jsonify({'status':'error','message':'No se pudo guardar el catálogo'}), 500
            return jsonify({'status':'ok','message':'Valor guardado','opciones':perfil})

        pn = request.args.get('perfil_nombre', '').strip()
        data = cargar_ubicaciones_perfiles()
        perfil = data.get(pn, {}) if pn else {}
        opciones = {'edificio':[], 'piso':[], 'dependencia':[]}
        for campo in opciones:
            valores = perfil.get(campo, []) if isinstance(perfil, dict) else []
            if isinstance(valores, list):
                opciones[campo] = sorted({str(v).strip() for v in valores if str(v).strip()}, key=str.casefold)
        return jsonify(opciones)
    except Exception as e:
        logger.exception('Error en api_opciones')
        return jsonify({'status':'error','message':str(e)}), 500

@main_bp.route('/api/opciones/admin', methods=['GET', 'POST'])
def api_opciones_admin():
    try:
        if request.method == 'GET':
            pn = request.args.get('perfil_nombre', '').strip()
            data = cargar_ubicaciones_perfiles()
            perfil = data.get(pn, {}) if pn else {}
            return jsonify({
                'edificio': list(perfil.get('edificio', [])) if isinstance(perfil, dict) else [],
                'piso': list(perfil.get('piso', [])) if isinstance(perfil, dict) else [],
                'dependencia': list(perfil.get('dependencia', [])) if isinstance(perfil, dict) else []
            })

        d = request.get_json(silent=True) or {}
        cfg = cargar_config()
        if d.get('password') != cfg.get('seguridad', {}).get('admin_password'):
            return jsonify({'status':'error','message':'Contraseña de administrador incorrecta'}), 401
        pn = str(d.get('perfil_nombre', '')).strip()
        if not pn or not buscar_perfil(pn):
            return jsonify({'status':'error','message':'Perfil no encontrado'}), 404
        opciones = d.get('opciones') or {}
        normal = {}
        for campo in ('edificio','piso','dependencia'):
            vals = opciones.get(campo, [])
            if not isinstance(vals, list): vals = []
            normal[campo] = sorted({str(v).strip() for v in vals if str(v).strip()}, key=str.casefold)
        data = cargar_ubicaciones_perfiles()
        data[pn] = normal
        if not guardar_ubicaciones_perfiles(data):
            return jsonify({'status':'error','message':'No se pudo guardar el catálogo'}), 500
        return jsonify({'status':'ok','message':'Catálogo guardado','opciones':normal})
    except Exception as e:
        logger.exception('Error en api_opciones_admin')
        return jsonify({'status':'error','message':str(e)}), 500

@main_bp.route('/api/mantenimiento/servidores/exportar', methods=['GET'])
def api_mantenimiento_exportar_servidores():
    if not _vp_servidores_backup_habilitado():
        return jsonify({'status': 'error', 'message': 'La función está disponible en las aplicaciones instalables.'}), 404
    if not os.path.isfile(SFTP_SERVERS_FILE):
        return jsonify({'status': 'error', 'message': 'Todavía no hay servidores guardados para exportar.'}), 404
    try:
        return send_file(
            SFTP_SERVERS_FILE,
            as_attachment=True,
            download_name='voip_provisioner_servidores.json',
            mimetype='application/json'
        )
    except Exception as exc:
        logger.exception('No se pudo exportar el respaldo de servidores')
        return jsonify({'status': 'error', 'message': str(exc)[:220]}), 500


@main_bp.route('/api/mantenimiento/servidores/importar', methods=['POST'])
def api_mantenimiento_importar_servidores():
    if not _vp_servidores_backup_habilitado():
        return jsonify({'status': 'error', 'message': 'La función está disponible en las aplicaciones instalables.'}), 404
    archivo = request.files.get('file')
    if archivo is None:
        return jsonify({'status': 'error', 'message': 'Seleccioná el archivo JSON de respaldo.'}), 400
    try:
        raw = archivo.read(1024 * 1024 + 1)
        if not raw or len(raw) > 1024 * 1024:
            raise ValueError('El respaldo está vacío o supera 1 MB.')
        contenido = json.loads(raw.decode('utf-8-sig'))
        if isinstance(contenido, dict):
            contenido = contenido.get('perfiles', [])
        if not isinstance(contenido, list) or not contenido or len(contenido) > 1000:
            raise ValueError('El JSON no contiene una lista válida de servidores.')

        perfiles = []
        nombres = set()
        for item in contenido:
            if not isinstance(item, dict):
                raise ValueError('El respaldo contiene un registro de servidor inválido.')
            nombre = str(item.get('nombre', '') or '').strip()
            host = str(item.get('host', '') or '').strip()
            usuario = str(item.get('usuario', '') or '').strip()
            password = str(item.get('password', '') or '')
            ruta = str(item.get('ruta_remota', '/tftpboot') or '/tftpboot').strip()
            tecnologia = str(item.get('tecnologia', 'sip') or 'sip').strip().lower()
            try:
                puerto = int(item.get('puerto', 22) or 22)
            except Exception:
                raise ValueError('El puerto de un servidor no es numérico.')
            if not nombre or not host or not usuario:
                raise ValueError('Cada servidor debe tener nombre, host y usuario.')
            if len(nombre) > 120 or len(host) > 255 or len(usuario) > 255 or len(password) > 4096 or len(ruta) > 1024:
                raise ValueError('Un campo del respaldo excede el tamaño permitido.')
            if puerto < 1 or puerto > 65535:
                raise ValueError('El puerto SSH debe estar entre 1 y 65535.')
            if tecnologia not in ('sip', 'pjsip'):
                tecnologia = 'sip'
            clave_nombre = nombre.casefold()
            if clave_nombre in nombres:
                raise ValueError('Hay nombres de perfil duplicados: ' + nombre)
            nombres.add(clave_nombre)
            perfil = {
                'nombre': nombre,
                'host': host,
                'usuario': usuario,
                'password': password,
                'ruta_remota': ruta or '/tftpboot',
                'puerto': str(puerto),
                'tecnologia': tecnologia,
                'ssh_compat': bool(item.get('ssh_compat', False)),
                'dialplan': _vp_perfil_dialplan(item.get('dialplan', [])),
            }
            perfiles.append(perfil)

        if os.path.isfile(SFTP_SERVERS_FILE):
            try:
                respaldo = SFTP_SERVERS_FILE + '.bak_import_' + time.strftime('%Y%m%d_%H%M%S')
                shutil.copy2(SFTP_SERVERS_FILE, respaldo)
            except Exception:
                logger.warning('No se pudo generar un respaldo previo de los perfiles antes de importar.', exc_info=True)
        temporal = SFTP_SERVERS_FILE + '.tmp'
        os.makedirs(os.path.dirname(SFTP_SERVERS_FILE), exist_ok=True)
        with open(temporal, 'w', encoding='utf-8') as destino:
            json.dump(perfiles, destino, indent=2, ensure_ascii=False)
        os.replace(temporal, SFTP_SERVERS_FILE)
        return jsonify({
            'status': 'ok',
            'importados': len(perfiles),
            'message': 'Se importaron {} servidor(es). Las contraseñas se guardaron en el almacenamiento privado de la app.'.format(len(perfiles))
        })
    except Exception as exc:
        logger.warning('No se pudo importar el respaldo de servidores: %s', exc)
        return jsonify({'status': 'error', 'message': str(exc)[:300]}), 400


@main_bp.route('/api/mantenimiento/actualizacion-app/estado', methods=['GET'])
def api_actualizacion_app_estado():
    return jsonify(updater_status())


@main_bp.route('/api/mantenimiento/actualizacion-app/comprobar', methods=['POST'])
def api_actualizacion_app_comprobar():
    if not updater_enabled():
        return jsonify({'status': 'error', 'message': 'El actualizador público está disponible solo en Android y Windows.'}), 404
    try:
        return jsonify(check_for_updates())
    except Exception as exc:
        logger.warning('No se pudo comprobar actualizaciones públicas: %s', exc)
        return jsonify({'status': 'error', 'message': str(exc)[:350]}), 400


@main_bp.route('/api/mantenimiento/actualizacion-app/aplicar', methods=['POST'])
def api_actualizacion_app_aplicar():
    if not updater_enabled():
        return jsonify({'status': 'error', 'message': 'El actualizador público está disponible solo en Android y Windows.'}), 404
    try:
        return jsonify(apply_updates())
    except Exception as exc:
        logger.warning('No se pudo aplicar la actualización OTA pública: %s', exc, exc_info=True)
        return jsonify({'status': 'error', 'message': str(exc)[:400]}), 400


@main_bp.route('/api/modelos')
def api_modelos():
    return jsonify(list(cargar_plantillas().keys()))


@main_bp.route('/api/modelo_por_mac')
def api_modelo_por_mac():
    """Devuelve marca por OUI y plantilla solo si el archivo remoto confirma el modelo exacto."""
    mac = _vp_excel_mac(request.args.get('mac', ''))
    perfil_nombre = str(request.args.get('perfil_nombre', '') or '').strip()
    if not mac:
        return jsonify({'status': 'error', 'message': 'MAC inválida'}), 400

    marca = ''
    try:
        marca = str(lookup_mac_vendor(mac) or '').strip()
    except Exception:
        logger.warning('No se pudo consultar la marca OUI de la MAC %s', mac, exc_info=True)

    modelo = ''
    pf = buscar_perfil(perfil_nombre) if perfil_nombre else None
    if pf:
        try:
            modelos = _vp_excel_modelos_config_cache(pf)
            modelo = str((modelos.get('mac') or {}).get(mac, '') or '').strip()
        except Exception:
            logger.warning('No se pudo confirmar el modelo remoto de la MAC %s', mac, exc_info=True)

    # No inferir el modelo a partir del fabricante: solo devolver una plantilla
    # cuando el XML/CFG remoto haya confirmado uno de los tres modelos probados.
    plantillas = {
        'grandstream gxp1628': 'GrandstreamGXP1628',
        'fanvil v61g': 'Fanvil V61G',
        'denwa dw-320g': 'Denwa DW-320G',
    }
    plantilla = plantillas.get(modelo.lower(), '')
    if not marca and modelo:
        marca = modelo.split()[0]
    return jsonify({
        'status': 'ok',
        'mac': mac.upper(),
        'marca': marca,
        'modelo': modelo,
        'plantilla': plantilla,
        'modelo_confirmado': bool(plantilla),
    })


def _vp_perfil_dialplan(valor):
    if not isinstance(valor, list):
        return []

    resultado = []
    for item in valor:
        item = str(item or '').strip()
        if item and item not in resultado:
            resultado.append(item)

    return resultado


@main_bp.route('/api/perfiles', methods=['GET', 'POST'])
def api_perfiles():
    if request.method == 'GET':
        resultado = []

        for x in cargar_perfiles():
            item = dict(x)
            item['password'] = str(x.get('password', ''))[:4] + '****'
            item['dialplan'] = _vp_perfil_dialplan(
                x.get('dialplan', [])
            )
            resultado.append(item)

        return jsonify(resultado)

    d = request.json or {}
    p = cargar_perfiles()

    dialplan = _vp_perfil_dialplan(
        d.get('dialplan', [])
    )

    orig = d.get('original_nombre', '')

    if orig:
        encontrado = False

        for x in p:
            if x['nombre'] == orig:
                x.update({
                    'nombre': d['nombre'],
                    'host': d['host'],
                    'usuario': d['usuario'],
                    'password': (
                        d.get('password') or
                        x.get('password', '')
                    ),
                    'ruta_remota': d.get(
                        'ruta_remota',
                        '/tftpboot'
                    ),
                    'puerto': d.get(
                        'puerto',
                        '22'
                    ),
                    'ssh_compat': d.get(
                        'ssh_compat',
                        False
                    ),
                    'tecnologia': (
                        d.get('tecnologia') or
                        x.get('tecnologia', 'sip')
                    ),
                    'dialplan': dialplan
                })
                encontrado = True
                break

        if not encontrado:
            return jsonify({
                'status': 'error',
                'message': 'Perfil original no encontrado'
            })

    else:
        if any(x['nombre'] == d.get('nombre') for x in p):
            return jsonify({
                'status': 'error',
                'message': 'Ya existe un perfil con ese nombre'
            })

        p.append({
            'nombre': d['nombre'],
            'host': d['host'],
            'usuario': d['usuario'],
            'password': d.get('password', ''),
            'ruta_remota': d.get('ruta_remota', '/tftpboot'),
            'puerto': d.get('puerto', '22'),
            'ssh_compat': d.get('ssh_compat', False),
            'tecnologia': d.get('tecnologia', 'sip'),
            'dialplan': dialplan
        })

    guardar_perfiles(p)
    return jsonify({'status': 'ok'})


@main_bp.route('/api/verificar_password', methods=['POST'])
def verificar_pw():
    return jsonify({'status': 'ok' if request.json.get('password') == cargar_config()['seguridad']['admin_password'] else 'error'})


def _vp_mac_vendor_enriquecer_excel(ep, perfil_nombre=''):
    """Completa solo modelos vacíos a partir de la MAC; no conecta al Issabel."""
    import shutil
    import tempfile
    wb = None
    temporal = None
    try:
        if not ep or not os.path.isfile(ep):
            return 0, False
        with _vp_excel_lock(str(perfil_nombre or ep)):
            wb = load_workbook(ep)
            ws = wb.active
            headers = [_vp_excel_header(c.value) for c in ws[1]]
            if 'mac' not in headers or 'modelo telefono' not in headers:
                wb.close()
                wb = None
                return 0, False
            col_mac = headers.index('mac') + 1
            col_modelo = headers.index('modelo telefono') + 1
            cambios = 0
            for fila in range(2, ws.max_row + 1):
                actual = str(ws.cell(fila, col_modelo).value or '').strip()
                if actual:
                    continue
                mac = ws.cell(fila, col_mac).value
                marca = lookup_mac_vendor(mac)
                if marca:
                    ws.cell(fila, col_modelo).value = marca
                    cambios += 1
            if cambios == 0:
                wb.close()
                wb = None
                return 0, False
            carpeta = os.path.dirname(ep) or '.'
            fd, temporal = tempfile.mkstemp(prefix='.vp_mac_vendor_', suffix='.xlsx', dir=carpeta)
            os.close(fd)
            wb.save(temporal)
            wb.close()
            wb = None
            # Valida el archivo antes de respaldar y reemplazar el original.
            prueba = load_workbook(temporal, read_only=True, data_only=False)
            prueba_ws = prueba.active
            prueba_headers = [_vp_excel_header(c.value) for c in prueba_ws[1]]
            valido = 'mac' in prueba_headers and 'modelo telefono' in prueba_headers
            prueba.close()
            if not valido:
                raise RuntimeError('El Excel temporal no pasó la validación de encabezados.')
            respaldo = 'servidor remoto' if os.environ.get('VOIP_PROVISIONER_REMOTE_EXCEL') == '1' else ''
            if os.environ.get('VOIP_PROVISIONER_REMOTE_EXCEL') == '1':
                pf_vendor = buscar_perfil(perfil_nombre)
                if not pf_vendor:
                    raise RuntimeError('No se encontró el perfil del Excel remoto.')
                subir_excel_remoto(
                    perfil_nombre, pf_vendor.get('ruta_remota', '/tftpboot'),
                    source_path=temporal, backup=True, raise_on_error=True
                )
            else:
                respaldo = ep + '.bak_mac_vendor_' + time.strftime('%Y%m%d_%H%M%S') + '_' + str(os.getpid())
                shutil.copy2(ep, respaldo)
            os.replace(temporal, ep)
            temporal = None
            logger.info('Fabricantes MAC: completadas %s filas del Excel del perfil %s; respaldo %s', cambios, perfil_nombre, respaldo)
            return cambios, False
    except Exception:
        logger.exception('No se pudo completar Modelo por MAC en el Excel del perfil %s', perfil_nombre or ep)
        return 0, True
    finally:
        if wb is not None:
            try:
                wb.close()
            except Exception:
                pass
        if temporal and os.path.isfile(temporal):
            try:
                os.remove(temporal)
            except OSError:
                pass


@main_bp.route('/api/admin/mac-vendors/refresh', methods=['POST'])
def api_admin_mac_vendors_refresh():
    uploaded_mode = bool(request.files)
    data = request.form if uploaded_mode else (request.get_json(silent=True) or {})
    password = str(data.get('password', '') or '')
    cfg = cargar_config()
    admin_password = str(cfg.get('seguridad', {}).get('admin_password', '') or '')
    if not password or password != admin_password:
        return jsonify({'status': 'error', 'message': 'Contraseña de administrador incorrecta.'}), 403
    try:
        if uploaded_mode:
            archivos = {}
            for registry, field in (('MA-L', 'ma_l'), ('MA-M', 'ma_m'), ('MA-S', 'ma_s')):
                upload = request.files.get(field)
                if upload is None or not str(upload.filename or '').strip():
                    return jsonify({'status': 'error', 'message': 'Falta el archivo CSV del registro {}.'.format(registry)}), 400
                raw = upload.read(25 * 1024 * 1024 + 1)
                if not raw:
                    return jsonify({'status': 'error', 'message': 'El archivo {} está vacío.'.format(registry)}), 400
                if len(raw) > 25 * 1024 * 1024:
                    return jsonify({'status': 'error', 'message': 'El archivo {} excede 25 MB.'.format(registry)}), 400
                archivos[registry] = raw
            base = refresh_ieee_database_from_files(archivos)
        else:
            base = refresh_ieee_database()
    except Exception as e:
        logger.exception('No se pudo actualizar la base local de fabricantes MAC desde IEEE.')
        return jsonify({
            'status': 'error',
            'message': 'No se pudo actualizar la base IEEE. La base local anterior se conserva. Detalle: {}'.format(str(e)),
        }), 502

    filas_completadas = 0
    archivos_actualizados = 0
    archivos_con_error = []
    perfiles = cargar_perfiles() or []
    vistos = set()
    for perfil in perfiles:
        if not isinstance(perfil, dict):
            continue
        nombre = str(perfil.get('nombre', '') or '').strip()
        if not nombre:
            continue
        try:
            ep = get_excel_file(nombre)
            if os.environ.get('VOIP_PROVISIONER_REMOTE_EXCEL') == '1':
                _vp_excel_preparar_maestro_remoto(nombre, ep)
            clave = os.path.abspath(ep)
            if clave in vistos:
                continue
            vistos.add(clave)
            filas, error = _vp_mac_vendor_enriquecer_excel(ep, nombre)
            filas_completadas += filas
            if filas:
                archivos_actualizados += 1
            if error:
                archivos_con_error.append(nombre)
        except Exception:
            logger.exception('Error recorriendo el Excel del perfil %s para completar fabricantes.', nombre)
            archivos_con_error.append(nombre)

    mensaje = 'Base IEEE actualizada: {} asignaciones. Se completaron {} celdas Modelo en {} archivos Excel.'.format(
        int(base.get('total_entries', 0)), filas_completadas, archivos_actualizados
    )
    if archivos_con_error:
        mensaje += ' No se pudieron procesar {} perfiles: {}.'.format(len(archivos_con_error), ', '.join(archivos_con_error[:8]))
    return jsonify({
        'status': 'ok',
        'message': mensaje,
        'updated_at': base.get('updated_at', ''),
        'total_entries': int(base.get('total_entries', 0)),
        'registry_counts': base.get('counts', {}),
        'profiles_checked': len(vistos),
        'files_updated': archivos_actualizados,
        'rows_filled': filas_completadas,
        'profiles_with_errors': archivos_con_error,
    })


@main_bp.route('/api/config', methods=['GET'])
def api_get_config():
    c = cargar_config()
    c['seguridad']['admin_password'] = '****'
    return jsonify(c)

@main_bp.route('/api/config', methods=['POST'])
def api_update_config():
    d = request.json
    cfg = cargar_config()
    if d.get('password') != cfg['seguridad']['admin_password']:
        return jsonify({'status': 'error', 'message': 'Contrasena incorrecta'})
    if 'empresa' in d:
        cfg['empresa'].update(d['empresa'])
    if 'seguridad' in d and 'admin_password' in d['seguridad']:
        cfg['seguridad']['admin_password'] = d['seguridad']['admin_password']
    if 'aplicacion' in d:
        cfg['aplicacion'].update(d['aplicacion'])
    guardar_config(cfg)
    return jsonify({'status': 'ok', 'message': 'Configuracion actualizada'})

@main_bp.route('/api/plantillas', methods=['GET'])
def api_plt_list():
    # Mostrar solamente los tres modelos/formatos permitidos, aunque haya
    # archivos antiguos en la carpeta de plantillas del usuario.
    disponibles = cargar_plantillas()
    r = []
    for nombre, extension in PLANTILLAS_PERMITIDAS.items():
        if nombre not in disponibles:
            continue
        ruta = os.path.join(PLANTILLAS_FOLDER, "{}.{}".format(nombre, extension))
        if os.path.isfile(ruta):
            tamano = os.path.getsize(ruta)
        else:
            tamano = len(disponibles[nombre].encode('utf-8'))
        r.append({
            'nombre': nombre,
            'tipo': extension.upper(),
            'tamano': tamano
        })
    return jsonify(r)


@main_bp.route('/api/plantillas', methods=['POST'])
def api_plt_save():
    cfg = cargar_config()
    d = request.json or {}

    if d.get('password') != cfg['seguridad']['admin_password']:
        return jsonify({
            'status': 'error',
            'message': 'Contrasena incorrecta'
        })

    n = str(d.get('nombre', '') or '').strip()
    c = d.get('contenido', '')
    tipo = str(d.get('tipo', 'CFG') or 'CFG').strip().upper()

    if not n:
        return jsonify({
            'status': 'error',
            'message': 'Nombre obligatorio'
        })

    if tipo not in ('CFG', 'XML'):
        tipo = 'CFG'

    # El nombre ingresado en Admin es el nombre real
    # de la plantilla.
    if n.lower().endswith('.cfg') or n.lower().endswith('.xml'):
        n = os.path.splitext(n)[0].strip()

    if PLANTILLAS_PERMITIDAS.get(n) != tipo.lower():
        return jsonify({
            'status': 'error',
            'message': 'Solo se admiten GrandstreamGXP1628 (XML), Fanvil V61G (CFG) y Denwa DW-320G (CFG).'
        }), 400

    nombre_archivo = "{}.{}".format(n, tipo.lower())
    ruta = os.path.join(PLANTILLAS_FOLDER, nombre_archivo)

    try:
        with open(ruta, 'w', encoding='utf-8') as f:
            f.write(c)
    except Exception as e:
        logger.exception("Error guardando plantilla")
        return jsonify({
            'status': 'error',
            'message': 'No se pudo guardar la plantilla: {}'.format(e)
        })

    eliminadas = cargar_plantillas_eliminadas()
    if n in eliminadas:
        eliminadas.discard(n)
        if not guardar_plantillas_eliminadas(eliminadas):
            return jsonify({'status': 'error', 'message': 'La plantilla se guardó, pero no se pudo actualizar su estado persistente.'}), 500

    return jsonify({
        'status': 'ok',
        'message': 'Guardada',
        'nombre': n,
        'tipo': tipo,
        'archivo': ruta
    })


@main_bp.route('/api/plantillas/<nombre>', methods=['GET'])
def api_plt_get(nombre):
    tipo = str(
        request.args.get('tipo', 'CFG') or 'CFG'
    ).strip().upper()

    if tipo not in ('CFG', 'XML'):
        tipo = 'CFG'

    if nombre.lower().endswith('.cfg') or nombre.lower().endswith('.xml'):
        nombre = os.path.splitext(nombre)[0]

    if PLANTILLAS_PERMITIDAS.get(nombre) != tipo.lower():
        return jsonify({'status': 'error', 'message': 'Plantilla no admitida'}), 404
    if nombre in cargar_plantillas_eliminadas():
        return jsonify({'status': 'error', 'message': 'La plantilla fue eliminada por el usuario.'}), 404

    p = os.path.join(
        PLANTILLAS_FOLDER,
        "{}.{}".format(nombre, tipo.lower())
    )

    if os.path.isfile(p):
        try:
            with open(p, 'r', encoding='utf-8') as f:
                contenido = f.read()
        except Exception:
            return jsonify({'status': 'error'}), 500
    else:
        # Permitir visualizar plantillas incluidas en el cliente antes de
        # que exista una copia editable en el directorio privado.
        contenido = cargar_plantillas().get(nombre)
        if contenido is None:
            return jsonify({'status': 'error'}), 404

    return jsonify({
        'status': 'ok',
        'contenido': contenido,
        'nombre': nombre,
        'tipo': tipo
    })


@main_bp.route('/api/plantillas/<nombre>', methods=['DELETE'])
def api_plt_del(nombre):
    cfg = cargar_config()
    d = request.json or {}

    if d.get('password') != cfg['seguridad']['admin_password']:
        return jsonify({
            'status': 'error',
            'message': 'Contrasena incorrecta'
        }), 401

    tipo = str(
        d.get('tipo', 'CFG') or 'CFG'
    ).strip().upper()

    if tipo not in ('CFG', 'XML'):
        tipo = 'CFG'

    if nombre.lower().endswith('.cfg') or nombre.lower().endswith('.xml'):
        nombre = os.path.splitext(nombre)[0]

    if PLANTILLAS_PERMITIDAS.get(nombre) != tipo.lower():
        return jsonify({'status': 'error', 'message': 'Plantilla no admitida'}), 404

    # Registrar primero el borrado permanente; cargar_plantillas respetará
    # esta marca incluso si la plantilla original sigue dentro del APK/EXE.
    eliminadas = cargar_plantillas_eliminadas()
    eliminadas.add(nombre)
    if not guardar_plantillas_eliminadas(eliminadas):
        return jsonify({
            'status': 'error',
            'message': 'No se pudo guardar el borrado; la plantilla no se quitó.'
        }), 500

    ruta = os.path.join(PLANTILLAS_FOLDER, '{}.{}'.format(nombre, tipo.lower()))
    # En Issabel, PLANTILLAS_FOLDER puede coincidir con el directorio fuente
    # incluido en el proyecto. No borrar el archivo versionado: la marca
    # persistente basta para ocultarlo, también después de una actualización.
    raiz_proyecto = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    carpetas_incluidas = set()
    for carpeta in (
        os.environ.get('VOIP_PROVISIONER_BUNDLED_PLANTILLAS', ''),
        os.environ.get('VOIP_PROVISIONER_BUNDLED_PLANTILLAS_EXTRA', ''),
        os.path.join(raiz_proyecto, 'data', 'plantillas'),
        os.path.join(raiz_proyecto, 'plantillas')
    ):
        if carpeta:
            carpetas_incluidas.add(os.path.abspath(carpeta))
    if os.path.abspath(PLANTILLAS_FOLDER) not in carpetas_incluidas:
        try:
            if os.path.isfile(ruta):
                os.remove(ruta)
        except Exception:
            # La marca persistente oculta la plantilla aunque no se pueda
            # borrar físicamente el archivo local.
            logger.exception('No se pudo borrar el archivo local de la plantilla %s', ruta)

    return jsonify({
        'status': 'ok',
        'message': 'Plantilla eliminada. No volverá a aparecer al reiniciar la aplicación.'
    })


@main_bp.route('/api/test_sftp', methods=['POST'])
def api_test_sftp():
    pf = buscar_perfil(request.json.get('perfil_nombre', ''))
    if not pf:
        return jsonify({'status': 'error', 'message': 'Perfil no encontrado'})
    try:
        s, _ = conectar_sftp(pf['host'], int(pf.get('puerto', 22)), pf['usuario'],
                             pf['password'], pf.get('ruta_remota', '/tftpboot'),
                             pf.get('ssh_compat', False))
        s.close()
        return jsonify({'status': 'ok', 'message': 'Conectado a ' + pf['host']})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)[:200]})

@main_bp.route('/api/perfiles/agregar', methods=['POST'])
def api_perfiles_add():
    d = request.json
    p = cargar_perfiles()
    if any(x['nombre'] == d.get('nombre') for x in p):
        return jsonify({'status': 'error', 'message': 'Ya existe'})
    p.append({
        'nombre': d['nombre'], 'host': d['host'], 'usuario': d['usuario'],
        'password': d.get('password', ''), 'ruta_remota': d.get('ruta_remota', '/tftpboot'),
        'puerto': d.get('puerto', '22'), 'ssh_compat': d.get('ssh_compat', False),
        'tecnologia': d.get('tecnologia', 'sip')
    })
    guardar_perfiles(p)
    return jsonify({'status': 'ok'})

@main_bp.route('/api/perfiles/editar', methods=['POST'])
def api_perfiles_edit():
    d = request.json
    p = cargar_perfiles()
    for x in p:
        if x['nombre'] == d.get('nombre_original'):
            x.update({
                'nombre': d['nombre'], 'host': d['host'], 'usuario': d['usuario'],
                'password': d.get('password', ''), 'ruta_remota': d.get('ruta_remota', '/tftpboot'),
                'puerto': d.get('puerto', '22'), 'ssh_compat': d.get('ssh_compat', False),
                'tecnologia': d.get('tecnologia') or x.get('tecnologia', 'sip')
            })
            break
    guardar_perfiles(p)
    return jsonify({'status': 'ok'})

@main_bp.route('/api/perfiles/eliminar', methods=['POST'])
def api_perfiles_del():
    d = request.get_json(silent=True) or {}

    # IMPORTANTE:
    # El nombre puede ser vacío porque puede existir un perfil
    # mal creado con "nombre": "".
    #
    # Lo que NO permitimos es una petición sin el campo nombre.
    if 'nombre' not in d:
        return jsonify({
            'status': 'error',
            'message': 'Nombre de servidor no informado.'
        }), 400

    n = str(d.get('nombre') or '').strip()
    password = d.get('password', '')

    cfg = cargar_config()

    admin_password = (
        cfg.get('seguridad', {}).get('admin_password', '')
    )

    if password != admin_password:
        return jsonify({
            'status': 'error',
            'message': 'Contraseña de administrador incorrecta.'
        }), 401

    perfiles = cargar_perfiles()

    encontrados = [
        x for x in perfiles
        if str(x.get('nombre') or '').strip() == n
    ]

    if not encontrados:
        return jsonify({
            'status': 'error',
            'message': 'Servidor no encontrado.'
        }), 404

    perfiles = [
        x for x in perfiles
        if str(x.get('nombre') or '').strip() != n
    ]

    guardar_perfiles(perfiles)

    return jsonify({
        'status': 'ok',
        'message': (
            'Perfil vacío eliminado correctamente.'
            if not n
            else 'Servidor eliminado correctamente.'
        )
    })

@main_bp.route('/api/perfiles/<nombre>', methods=['DELETE'])
def api_perfiles_del_directo(nombre):
    p = [x for x in cargar_perfiles() if x['nombre'] != nombre]
    guardar_perfiles(p)
    return jsonify({'status': 'ok'})


# === VP INDEX TFTP EDITOR V1 ===
def _vp_index_tftp_profile(data):
    nombre = str((data or {}).get('perfil_nombre') or '').strip()
    if not nombre:
        return None, '', 'Debe seleccionar un perfil SSH.'
    perfil = buscar_perfil(nombre)
    if not perfil:
        return None, '', 'No se encontró el perfil SSH seleccionado.'
    if not str(perfil.get('host') or '').strip():
        return None, '', 'El perfil no tiene host configurado.'
    root = str(perfil.get('ruta_remota') or '/tftpboot').strip() or '/tftpboot'
    if not root.startswith('/'):
        root = '/' + root
    root = posixpath.normpath(root)
    if root in ('/', '.', '') or '\x00' in root:
        return None, '', 'La ruta TFTP del perfil no es válida.'
    return dict(perfil), root, None


def _vp_index_tftp_parse_name(nombre):
    nombre = str(nombre or '').strip()
    if not nombre or nombre in ('.', '..') or '/' in nombre or '\\' in nombre or '\x00' in nombre or posixpath.basename(nombre) != nombre:
        return None
    ext = posixpath.splitext(nombre)[1].lower()
    if ext not in ('.cfg', '.xml', '.conf', '.ini', '.txt'):
        return None
    base = posixpath.splitext(nombre)[0]
    if base.lower().startswith('cfg'):
        base = base[3:]
    base = re.sub(r'[^0-9a-fA-F]', '', base)
    if not re.match(r'(?i)^[0-9a-f]{12}$', base):
        return None
    return base.lower()


def _vp_index_tftp_local(perfil):
    return _vp_host_es_local(perfil.get('host'))


def _vp_index_tftp_connect(perfil, root):
    sftp, transporte = conectar_sftp(
        str(perfil.get('host') or '').strip(),
        int(perfil.get('puerto', 22) or 22),
        str(perfil.get('usuario') or 'root'),
        str(perfil.get('password') or ''),
        root,
        perfil.get('ssh_compat', False)
    )
    if sftp is None:
        raise RuntimeError('No se pudo establecer conexión SFTP con el servidor del perfil.')
    return sftp, transporte


@main_bp.route('/api/provisioner/tftp/list', methods=['POST'])
def vp_index_tftp_list():
    import stat as _vp_stat
    data = request.get_json(silent=True) or {}
    perfil, root, error = _vp_index_tftp_profile(data)
    if error:
        return jsonify({'status': 'error', 'message': error}), 400
    archivos = []
    sftp = transporte = None
    try:
        if _vp_index_tftp_local(perfil):
            if not os.path.isdir(root):
                return jsonify({'status': 'error', 'message': 'No existe o no se puede acceder al directorio TFTP: ' + root}), 404
            with os.scandir(root) as elementos:
                for elemento in elementos:
                    if not elemento.is_file(follow_symlinks=False):
                        continue
                    mac = _vp_index_tftp_parse_name(elemento.name)
                    if not mac:
                        continue
                    try:
                        info = elemento.stat(follow_symlinks=False)
                    except Exception:
                        continue
                    archivos.append({'nombre': elemento.name, 'mac': mac, 'tamano': int(info.st_size or 0), 'mtime': int(info.st_mtime or 0)})
        else:
            sftp, transporte = _vp_index_tftp_connect(perfil, root)
            for elemento in sftp.listdir_attr(root):
                nombre = str(elemento.filename)
                mac = _vp_index_tftp_parse_name(nombre)
                if not mac or not _vp_stat.S_ISREG(int(getattr(elemento, 'st_mode', 0) or 0)):
                    continue
                archivos.append({'nombre': nombre, 'mac': mac, 'tamano': int(getattr(elemento, 'st_size', 0) or 0), 'mtime': int(getattr(elemento, 'st_mtime', 0) or 0)})
        archivos.sort(key=lambda x: (x['mac'], x['nombre'].lower()))
        return jsonify({'status': 'ok', 'ruta': root, 'archivos': archivos})
    except Exception as e:
        logger.exception('Error listando archivos TFTP del perfil')
        return jsonify({'status': 'error', 'message': str(e)[:240]}), 500
    finally:
        if sftp is not None:
            try: sftp.close()
            except Exception: pass
        if transporte is not None:
            try: transporte.close()
            except Exception: pass


@main_bp.route('/api/provisioner/tftp/read', methods=['POST'])
def vp_index_tftp_read():
    import hashlib as _vp_hashlib
    import stat as _vp_stat
    data = request.get_json(silent=True) or {}
    perfil, root, error = _vp_index_tftp_profile(data)
    if error:
        return jsonify({'status': 'error', 'message': error}), 400
    nombre = str(data.get('nombre') or '').strip()
    if not _vp_index_tftp_parse_name(nombre):
        return jsonify({'status': 'error', 'message': 'Nombre de archivo no admitido.'}), 400
    ruta = posixpath.join(root, nombre)
    sftp = transporte = None
    try:
        if _vp_index_tftp_local(perfil):
            if os.path.dirname(os.path.abspath(ruta)) != os.path.abspath(root) or not os.path.isfile(ruta) or os.path.islink(ruta):
                return jsonify({'status': 'error', 'message': 'El archivo ya no existe o no es regular dentro de TFTP.'}), 404
            with open(ruta, 'rb') as f:
                contenido_bytes = f.read(2 * 1024 * 1024 + 1)
        else:
            sftp, transporte = _vp_index_tftp_connect(perfil, root)
            attrs = sftp.lstat(ruta)
            if not _vp_stat.S_ISREG(int(getattr(attrs, 'st_mode', 0) or 0)):
                return jsonify({'status': 'error', 'message': 'El elemento no es un archivo regular.'}), 400
            with sftp.open(ruta, 'rb') as f:
                contenido_bytes = f.read(2 * 1024 * 1024 + 1)
        if len(contenido_bytes) > 2 * 1024 * 1024:
            return jsonify({'status': 'error', 'message': 'El archivo supera el límite de edición de 2 MB.'}), 413
        return jsonify({'status': 'ok', 'nombre': nombre, 'contenido': contenido_bytes.decode('utf-8', 'replace'), 'sha256': _vp_hashlib.sha256(contenido_bytes).hexdigest(), 'tamano': len(contenido_bytes), 'ruta': root})
    except Exception as e:
        logger.exception('Error leyendo archivo TFTP %s', nombre)
        return jsonify({'status': 'error', 'message': str(e)[:240]}), 500
    finally:
        if sftp is not None:
            try: sftp.close()
            except Exception: pass
        if transporte is not None:
            try: transporte.close()
            except Exception: pass


@main_bp.route('/api/provisioner/tftp/write', methods=['POST'])
def vp_index_tftp_write():
    import hashlib as _vp_hashlib
    import shutil as _vp_shutil
    import stat as _vp_stat
    import tempfile as _vp_tempfile
    import time as _vp_time
    import uuid as _vp_uuid
    data = request.get_json(silent=True) or {}
    cfg = cargar_config()
    supplied = str(data.get('admin_password') or '')
    expected = str((cfg.get('seguridad') or {}).get('admin_password') or '')
    if not expected or supplied != expected:
        return jsonify({'status': 'error', 'message': 'Contraseña de administrador incorrecta.'}), 401
    perfil, root, error = _vp_index_tftp_profile(data)
    if error:
        return jsonify({'status': 'error', 'message': error}), 400
    nombre = str(data.get('nombre') or '').strip()
    if not _vp_index_tftp_parse_name(nombre):
        return jsonify({'status': 'error', 'message': 'Nombre de archivo no admitido.'}), 400
    contenido = data.get('contenido')
    if contenido is None:
        contenido = ''
    if not isinstance(contenido, str):
        contenido = str(contenido)
    nuevo = contenido.encode('utf-8')
    if len(nuevo) > 2 * 1024 * 1024:
        return jsonify({'status': 'error', 'message': 'El archivo supera el límite de edición de 2 MB.'}), 413
    sha_cliente = str(data.get('sha256') or '').strip().lower()
    if not re.match(r'^[0-9a-f]{64}$', sha_cliente):
        return jsonify({'status': 'error', 'message': 'No se recibió la huella de la versión abierta. Volvé a abrir el archivo.'}), 400
    ruta = posixpath.join(root, nombre)
    marca = _vp_time.strftime('%Y%m%d_%H%M%S') + '_' + _vp_uuid.uuid4().hex[:6]
    nombre_backup = nombre + '.bak_' + marca
    ruta_backup = posixpath.join(root, nombre_backup)
    sftp = transporte = temporal = None
    try:
        if _vp_index_tftp_local(perfil):
            if os.path.dirname(os.path.abspath(ruta)) != os.path.abspath(root) or not os.path.isfile(ruta) or os.path.islink(ruta):
                return jsonify({'status': 'error', 'message': 'El archivo ya no existe o no es regular dentro de TFTP.'}), 404
            info = os.stat(ruta, follow_symlinks=False)
            with open(ruta, 'rb') as f:
                original = f.read(2 * 1024 * 1024 + 1)
            if len(original) > 2 * 1024 * 1024:
                return jsonify({'status': 'error', 'message': 'El archivo supera el límite de edición de 2 MB.'}), 413
            if _vp_hashlib.sha256(original).hexdigest() != sha_cliente:
                return jsonify({'status': 'error', 'conflict': True, 'message': 'El archivo cambió en el servidor desde que lo abriste. Cerralo y volvé a abrirlo antes de guardar.'}), 409
            _vp_shutil.copy2(ruta, ruta_backup)
            fd, temporal = _vp_tempfile.mkstemp(prefix='.vp_tftp_edit_', dir=root)
            with os.fdopen(fd, 'wb') as f:
                f.write(nuevo)
                f.flush()
                os.fchmod(f.fileno(), _vp_stat.S_IMODE(info.st_mode))
                try: os.fchown(f.fileno(), info.st_uid, info.st_gid)
                except Exception: pass
            _vp_shutil.copystat(ruta, temporal, follow_symlinks=False)
            try: os.chown(temporal, info.st_uid, info.st_gid)
            except Exception: pass
            os.replace(temporal, ruta)
            temporal = None
        else:
            sftp, transporte = _vp_index_tftp_connect(perfil, root)
            attrs = sftp.lstat(ruta)
            if not _vp_stat.S_ISREG(int(getattr(attrs, 'st_mode', 0) or 0)):
                return jsonify({'status': 'error', 'message': 'El elemento ya no es un archivo regular.'}), 400
            with sftp.open(ruta, 'rb') as f:
                original = f.read(2 * 1024 * 1024 + 1)
            if len(original) > 2 * 1024 * 1024:
                return jsonify({'status': 'error', 'message': 'El archivo supera el límite de edición de 2 MB.'}), 413
            if _vp_hashlib.sha256(original).hexdigest() != sha_cliente:
                return jsonify({'status': 'error', 'conflict': True, 'message': 'El archivo cambió en el servidor desde que lo abriste. Cerralo y volvé a abrirlo antes de guardar.'}), 409
            with sftp.open(ruta_backup, 'wb') as f:
                f.write(original)
            modo = int(getattr(attrs, 'st_mode', 0) or 0) & 0o777
            if modo:
                try: sftp.chmod(ruta_backup, modo)
                except Exception: pass
            try:
                with sftp.open(ruta, 'wb') as f:
                    f.write(nuevo)
                if modo: sftp.chmod(ruta, modo)
            except Exception:
                try:
                    with sftp.open(ruta, 'wb') as f: f.write(original)
                    if modo: sftp.chmod(ruta, modo)
                except Exception:
                    logger.exception('Falló también la restauración del archivo TFTP %s', nombre)
                raise
        return jsonify({'status': 'ok', 'message': 'Archivo guardado con respaldo.', 'backup': nombre_backup, 'sha256': _vp_hashlib.sha256(nuevo).hexdigest()})
    except Exception as e:
        logger.exception('Error guardando archivo TFTP %s', nombre)
        return jsonify({'status': 'error', 'message': 'No se pudo guardar el archivo: ' + str(e)[:200]}), 500
    finally:
        if temporal and os.path.exists(temporal):
            try: os.remove(temporal)
            except Exception: pass
        if sftp is not None:
            try: sftp.close()
            except Exception: pass
        if transporte is not None:
            try: transporte.close()
            except Exception: pass
# === FIN VP INDEX TFTP EDITOR V1 ===


@main_bp.route('/api/tabla', methods=['GET'])
def api_tabla():
    """
    Devuelve los registros del Excel junto con la IP real del teléfono.

    Compatibilidad:
      - chan_sip -> sip show peers
      - PJSIP    -> pjsip show contacts

    La tecnología del perfil determina qué fuente consultar.
    No modifica el servidor Issabel.
    """
    try:
        pn = request.args.get('perfil_nombre', '')
        ep = get_excel_file(pn)
        pf = buscar_perfil(pn)
        if pn and pf and os.environ.get('VOIP_PROVISIONER_REMOTE_EXCEL') == '1':
            try:
                _vp_excel_preparar_maestro_remoto(pn, ep)
            except Exception as exc:
                return jsonify({
                    'status': 'error',
                    'message': 'No se pudo consultar el Excel maestro del servidor remoto: {}'.format(str(exc)[:220])
                }), 503

        # Si el Excel ya es valido, mantener la actualizacion normal con limite de frecuencia.
        # Si falta, se consulta Asterisk en el bloque de recuperacion inferior.
        if pn and pf and _vp_excel_tiene_registros(ep):
            _vp_excel_iniciar_actualizacion(pn, forzar=False, ep_override=ep)

        ips_dict = {}

        if pf:

            host = str(
                pf.get('host', '')
            ).strip()

            tecnologia = str(
                pf.get('tecnologia', 'sip')
            ).strip().lower()

            es_local = _vp_host_es_local(host)

            def ejecutar_estado(cmd):
                if es_local:
                    try:
                        result = subprocess.run(
                            cmd,
                            shell=True,
                            stdout=subprocess.PIPE,
                            stderr=subprocess.PIPE,
                            universal_newlines=True,
                            timeout=5
                        )

                        return (
                            result.stdout or '',
                            result.returncode
                        )

                    except Exception as e:
                        logger.error(
                            "api_tabla estado local: %s",
                            e
                        )

                        return '', -1

                try:
                    out, _, rc = ejecutar_comando_ssh(
                        pf,
                        cmd
                    )

                    return (
                        out or '',
                        rc
                    )

                except Exception as e:
                    logger.error(
                        "api_tabla estado SSH: %s",
                        e
                    )

                    return '', -1

            # ==================================================
            # PJSIP
            # ==================================================
            if tecnologia == 'pjsip':

                cmd_contacts = (
                    "/usr/sbin/asterisk -rx "
                    "'pjsip show contacts' 2>/dev/null"
                )

                out_contacts, rc_contacts = ejecutar_estado(
                    cmd_contacts
                )

                if (
                    rc_contacts == 0
                    and out_contacts
                ):

                    for line in out_contacts.splitlines():

                        line = line.strip()

                        if not line:
                            continue

                        # Formato habitual:
                        #
                        # Contact: 1993/sip:1993@10.12.34.56:5060
                        #
                        # También soporta:
                        #
                        # 1993/sip:1993@10.12.34.56:5060
                        #
                        # y sips:.

                        m = re.search(
                            r'(?:Contact:\s*)?'
                            r'([A-Za-z0-9_.:-]+)'
                            r'/'
                            r'(?:sip|sips):'
                            r'[^@\s>]+@'
                            r'('
                            r'\d{1,3}\.'
                            r'\d{1,3}\.'
                            r'\d{1,3}\.'
                            r'\d{1,3}'
                            r')'
                            r'(?::\d+)?',
                            line,
                            re.IGNORECASE
                        )

                        if not m:
                            continue

                        interno_pjsip = (
                            m.group(1)
                            .strip()
                        )

                        ip_pjsip = (
                            m.group(2)
                            .strip()
                        )

                        if (
                            interno_pjsip
                            and ip_pjsip
                        ):
                            ips_dict[
                                interno_pjsip
                            ] = ip_pjsip

                logger.info(
                    "api_tabla PJSIP: %s IPs encontradas",
                    len(ips_dict)
                )

            # ==================================================
            # CHAN_SIP
            # ==================================================
            elif tecnologia == 'sip':

                cmd_peers = (
                    "/usr/sbin/asterisk -rx "
                    "'sip show peers' 2>/dev/null"
                )

                out_peers, rc_peers = ejecutar_estado(
                    cmd_peers
                )

                if (
                    rc_peers == 0
                    and out_peers
                ):

                    for line in out_peers.splitlines():

                        if (
                            "Name/username" in line
                            or "----" in line
                            or "sip peers" in line
                            or not line.strip()
                        ):
                            continue

                        parts = line.split()

                        if len(parts) < 2:
                            continue

                        peer_name = (
                            parts[0]
                            .split('/')[0]
                            .strip()
                        )

                        peer_host = (
                            parts[1]
                            .strip()
                        )

                        if (
                            peer_name
                            and peer_host
                            and re.match(
                                r'^\d{1,3}(?:\.\d{1,3}){3}$',
                                peer_host
                            )
                            and 'Unspecified' not in peer_host
                        ):
                            ips_dict[
                                peer_name
                            ] = peer_host

                logger.info(
                    "api_tabla SIP: %s IPs encontradas",
                    len(ips_dict)
                )

            # ==================================================
            # Perfil sin tecnología definida:
            # consultar ambos sin romper compatibilidad.
            # ==================================================
            else:

                # SIP
                cmd_peers = (
                    "/usr/sbin/asterisk -rx "
                    "'sip show peers' 2>/dev/null"
                )

                out_peers, rc_peers = ejecutar_estado(
                    cmd_peers
                )

                if (
                    rc_peers == 0
                    and out_peers
                ):

                    for line in out_peers.splitlines():

                        if (
                            "Name/username" in line
                            or "----" in line
                            or "sip peers" in line
                            or not line.strip()
                        ):
                            continue

                        parts = line.split()

                        if len(parts) < 2:
                            continue

                        peer_name = (
                            parts[0]
                            .split('/')[0]
                            .strip()
                        )

                        peer_host = (
                            parts[1]
                            .strip()
                        )

                        if (
                            peer_name
                            and re.match(
                                r'^\d{1,3}(?:\.\d{1,3}){3}$',
                                peer_host
                            )
                            and 'Unspecified' not in peer_host
                        ):
                            ips_dict[
                                peer_name
                            ] = peer_host

                # PJSIP
                cmd_contacts = (
                    "/usr/sbin/asterisk -rx "
                    "'pjsip show contacts' 2>/dev/null"
                )

                out_contacts, rc_contacts = ejecutar_estado(
                    cmd_contacts
                )

                if (
                    rc_contacts == 0
                    and out_contacts
                ):

                    for line in out_contacts.splitlines():

                        line = line.strip()

                        m = re.search(
                            r'(?:Contact:\s*)?'
                            r'([A-Za-z0-9_.:-]+)'
                            r'/'
                            r'(?:sip|sips):'
                            r'[^@\s>]+@'
                            r'('
                            r'\d{1,3}\.'
                            r'\d{1,3}\.'
                            r'\d{1,3}\.'
                            r'\d{1,3}'
                            r')'
                            r'(?::\d+)?',
                            line,
                            re.IGNORECASE
                        )

                        if not m:
                            continue

                        interno_pjsip = (
                            m.group(1).strip()
                        )

                        ip_pjsip = (
                            m.group(2).strip()
                        )

                        if (
                            interno_pjsip
                            and ip_pjsip
                        ):
                            ips_dict[
                                interno_pjsip
                            ] = ip_pjsip

        # === VP IP POR INTERNO SIN ARP V1 ===
        # MAC e interno se cruzan mediante archivos de aprovisionamiento.
        # La IP se consulta en Asterisk y se entrega por numero de interno,
        # sin exigir que servidor y telefono compartan subred para resolver ARP.
        def _vp_ip_mac_normalizar(valor):
            return re.sub(r'[^0-9A-F]', '', str(valor or '').upper())

        # Compatibilidad: algunos perfiles figuran como SIP pero usan PJSIP.
        # En ese caso, las IP activas se consultan con pjsip show contacts.
        if pf and tecnologia == 'sip':
            _vp_cmd_pjsip_por_interno = (
                "/usr/sbin/asterisk -rx "
                "'pjsip show contacts' 2>/dev/null"
            )
            _vp_out_pjsip_por_interno, _vp_rc_pjsip_por_interno = ejecutar_estado(
                _vp_cmd_pjsip_por_interno
            )
            if _vp_rc_pjsip_por_interno == 0 and _vp_out_pjsip_por_interno:
                _vp_regex_pjsip_por_interno = re.compile(
                    r'(?:Contact:\s*)?([A-Za-z0-9_.:-]+)/'
                    r'(?:sip|sips):[^@\s>]+@'
                    r'(\d{1,3}(?:\.\d{1,3}){3})(?::\d+)?',
                    re.IGNORECASE
                )
                for _vp_linea_pjsip_por_interno in _vp_out_pjsip_por_interno.splitlines():
                    _vp_match_pjsip_por_interno = _vp_regex_pjsip_por_interno.search(
                        _vp_linea_pjsip_por_interno.strip()
                    )
                    if not _vp_match_pjsip_por_interno:
                        continue
                    _vp_int_pjsip_por_interno = _vp_match_pjsip_por_interno.group(1).strip()
                    _vp_ip_pjsip_por_interno = _vp_match_pjsip_por_interno.group(2).strip()
                    if _vp_int_pjsip_por_interno and _vp_ip_pjsip_por_interno:
                        # Un valor por interno, como en la consulta principal PJSIP.
                        ips_dict[_vp_int_pjsip_por_interno] = _vp_ip_pjsip_por_interno

        # ======================================================
        # Excel
        # ======================================================

        if not _vp_excel_tiene_registros(ep):
            # Responder con Asterisk sin recorrer archivos remotos ni esperar el refresco completo.
            filas_directas = _vp_excel_filas_desde_asterisk(
                pn, ep=ep, ips_dict=ips_dict, guardar=False, incluir_macs=False
            )
            if pf:
                _vp_excel_iniciar_actualizacion(
                    pn, forzar=True, ep_override=ep
                )
            return jsonify(filas_directas), 200

        wb = load_workbook(ep)
        ws = wb.active

        h = [
            c.value
            for c in ws[1]
        ]

        tn = (
            len(h) >= 8
            and h[3] == "Nombre Mostrar"
        )

        hn = [str(v or '').strip().lower() for v in h]
        ci = hn.index("interno") if "interno" in hn else (5 if tn else 3)
        cm = hn.index("mac") if "mac" in hn else (6 if tn else 4)
        cp = hn.index("ip") if "ip" in hn else (7 if tn else 5)

        idx_serie = -1
        for pos, nombre in enumerate(hn):
            nombre = nombre.replace("ú", "u").replace("ó", "o")
            if nombre in ("numero serie", "numero de serie", "serie"):
                idx_serie = pos
                break
        if idx_serie < 0 and len(h) > 8:
            idx_serie = 8

        filas_excel_por_interno_api = {}
        for _vp_fila_conteo in ws.iter_rows(min_row=2, min_col=ci + 1, max_col=ci + 1, values_only=True):
            _vp_interno_conteo = str((_vp_fila_conteo[0] if _vp_fila_conteo else '') or '').strip()
            if _vp_interno_conteo:
                filas_excel_por_interno_api[_vp_interno_conteo] = filas_excel_por_interno_api.get(_vp_interno_conteo, 0) + 1

        r = []

        for row in ws.iter_rows(
            min_row=2
        ):

            if len(row) <= cp:
                continue

            interno = str(
                row[ci].value or ''
            ).strip()

            if not interno:
                continue

            # IP siempre por interno desde Asterisk, no por MAC ni ARP.
            # Si el mismo interno tiene varios contactos activos, esta API
            # muestra la IP elegida para el interno; el Index conserva su
            # mecanismo existente de distribuir contactos reales por interno.
            ip_real = ips_dict.get(interno, '')

            r.append({
                'edificio': (
                    str(row[0].value or '').strip()
                    if len(row) > 0
                    else ''
                ),

                'piso': (
                    str(row[1].value or '').strip()
                    if len(row) > 1
                    else ''
                ),

                'dependencia': (
                    str(row[2].value or '').strip()
                    if len(row) > 2
                    else ''
                ),

                'nombre_mostrar': (
                    str(row[3].value or '').strip()
                    if tn
                    else str(row[0].value or '').strip()
                ),

                'nombre_saliente': (
                    str(row[4].value or '').strip()
                    if tn
                    else str(row[1].value or '').strip()
                ),

                'interno': interno,

                'mac': (
                    str(row[cm].value or '').strip()
                ),

                'ip': ip_real,

                'numero_serie': (
                    str(row[idx_serie].value or '').strip()
                    if idx_serie >= 0 and len(row) > idx_serie
                    else ''
                ),
                'modelo': (
                    str(row[h.index("Modelo Telefono")].value or '').strip()
                    if "Modelo Telefono" in h and len(row) > h.index("Modelo Telefono")
                    else ''
                ),
                'fila_excel': row[0].row,
            })

        wb.close()

        # Completar la MAC que el Index muestra cuando falta en el Excel, usando
        # solo la intersección Asterisk + archivos. No se sustituyen MAC existentes.
        # En extensiones con varios equipos, solo asignar si la correspondencia por
        # fila no es ambigua; no tocar modelo ni serie manuales.
        if pf and r:
            try:
                mapa_macs_index = _vp_excel_mapa_mac_para_index(pf)
                filas_por_interno_index = {}
                for registro in r:
                    filas_por_interno_index.setdefault(str(registro.get('interno') or '').strip(), []).append(registro)

                for interno_index, registros in filas_por_interno_index.items():
                    candidatas = sorted(set(mapa_macs_index.get(interno_index, [])))
                    if not candidatas:
                        continue
                    existentes = set(
                        _vp_ip_mac_normalizar(registro.get('mac'))
                        for registro in registros if _vp_ip_mac_normalizar(registro.get('mac'))
                    )
                    faltantes = [mac for mac in candidatas if _vp_ip_mac_normalizar(mac) not in existentes]
                    vacias = [registro for registro in registros if not _vp_ip_mac_normalizar(registro.get('mac'))]
                    asignaciones = []

                    if len(faltantes) == 1 and len(vacias) == 1:
                        asignaciones = [(vacias[0], faltantes[0])]
                    elif len(faltantes) > 1 and len(vacias) == len(faltantes) and vacias:
                        # Asignación estable permitida solo cuando esas filas no tienen
                        # datos manuales que pudieran quedar ligados al teléfono equivocado.
                        sin_datos_manuales = all(
                            not str(registro.get('modelo') or '').strip() and
                            not str(registro.get('numero_serie') or '').strip()
                            for registro in vacias
                        )
                        if sin_datos_manuales:
                            vacias_ordenadas = sorted(vacias, key=lambda x: int(x.get('fila_excel') or 0))
                            asignaciones = list(zip(vacias_ordenadas, faltantes))

                    for registro, mac_nueva in asignaciones:
                        registro['mac'] = str(mac_nueva).upper()
                        ip_del_interno = ips_dict.get(interno_index, '')
                        if ip_del_interno:
                            registro['ip'] = ip_del_interno
            except Exception:
                logger.warning('api_tabla: no se pudieron completar MAC pendientes desde Asterisk/configs.', exc_info=True)

        logger.info(
            "api_tabla perfil=%s tecnologia=%s registros=%s ips=%s",

            pn,
            (
                pf.get('tecnologia', 'sip')
                if pf
                else 'sin perfil'
            ),
            len(r),
            len(ips_dict)
        )

        return jsonify(r)

    except Exception as e:

        logger.error(
            "Error en api_tabla: %s",
            e,
            exc_info=True
        )

        return jsonify([])


@main_bp.route('/api/exportar_excel')
def exportar_excel():
    from openpyxl.styles import PatternFill, Font, Alignment
    from openpyxl.utils import get_column_letter
    import tempfile
    
    pn = request.args.get('perfil_nombre', '')
    ep = get_excel_file(pn)
    pf_export = buscar_perfil(pn)
    if pn and pf_export and os.environ.get('VOIP_PROVISIONER_REMOTE_EXCEL') == '1':
        try:
            _vp_excel_preparar_maestro_remoto(pn, ep)
        except Exception as exc:
            return jsonify({
                'status': 'error',
                'message': 'No se pudo consultar el Excel maestro del servidor remoto: {}'.format(str(exc)[:220])
            }), 503

    if not _vp_excel_tiene_registros(ep):
        _vp_excel_filas_desde_asterisk(
            pn, ep=ep, guardar=True, incluir_macs=False
        )
        if not _vp_excel_tiene_registros(ep):
            return jsonify({
                'status': 'error',
                'message': 'Asterisk no devolvio registros; no se genero un Excel vacio.'
            }), 503
        if pf_export:
            _vp_excel_iniciar_actualizacion(pn, forzar=True, ep_override=ep)
    elif pf_export:
        _vp_excel_iniciar_actualizacion(pn, forzar=False, ep_override=ep)

    migrar_excel(ep)
    if pf_export:
        try:
            _vp_excel_rellenar_modelos_desde_configs(ep, pf_export)
        except Exception:
            logger.warning('Exportación Excel: no se pudieron completar los modelos desde configs.', exc_info=True)
    wb = load_workbook(ep)
    ws = wb.active
    
    pf = buscar_perfil(pn)
    ips_actuales = {}
    if pf:
        cmd_peers = "asterisk -rx 'sip show peers' 2>/dev/null"
        es_local = _vp_host_es_local(pf['host'])
        if es_local:
            try:
                result = subprocess.run(cmd_peers, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, universal_newlines=True, timeout=5)
                out_peers = result.stdout
                c_peers = result.returncode
            except:
                out_peers, c_peers = "", -1
        else:
            out_peers, _, c_peers = ejecutar_comando_ssh(pf, cmd_peers)
        
        if out_peers and c_peers == 0:
            for line in out_peers.splitlines():
                if "Name/username" in line or "----" in line or "sip peers" in line or not line.strip():
                    continue
                parts = line.split()
                if len(parts) >= 2:
                    peer_name = parts[0].split('/')[0]
                    peer_host = parts[1]
                    if peer_host and '.' in peer_host and 'Unspecified' not in peer_host:
                        ips_actuales[peer_name] = peer_host
    
    fill_verde = PatternFill(start_color='27AE60', end_color='27AE60', fill_type='solid')
    fill_rojo = PatternFill(start_color='E74C3C', end_color='E74C3C', fill_type='solid')
    font_blanco = Font(color='FFFFFF', bold=True)
    
    h = [c.value for c in ws[1]]
    tn = len(h) >= 8 and h[3] == "Nombre Mostrar"
    ci = 5 if tn else 3
    cp = 7 if tn else 5
    
    for row in ws.iter_rows(min_row=2):
        if len(row) > cp:
            interno = str(row[ci].value or '').strip()
            ip_cell = row[cp]
            
            if interno in ips_actuales:
                ip_cell.value = ips_actuales[interno]
                ip_cell.hyperlink = 'https://' + ips_actuales[interno]
                ip_cell.fill = fill_verde
                ip_cell.font = Font(color='FFFFFF', bold=True, underline='single')
                ip_cell.alignment = Alignment(horizontal='center')
            elif ip_cell.value and str(ip_cell.value).strip():
                ip_cell.hyperlink = 'https://' + str(ip_cell.value)
                ip_cell.fill = fill_rojo
                ip_cell.font = Font(color='FFFFFF', bold=True, underline='single')
                ip_cell.alignment = Alignment(horizontal='center')
            else:
                ip_cell.value = '-'
                ip_cell.alignment = Alignment(horizontal='center')
    
    try:
        cfg = cargar_config()
        color_hex = cfg.get('empresa', {}).get('color_primario', '#4da6ff').lstrip('#')
        header_fill = PatternFill(start_color=color_hex, end_color=color_hex, fill_type='solid')
        header_font = Font(color='FFFFFF', bold=True, size=11)
        header_align = Alignment(horizontal='center', vertical='center')
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_align
    except:
        pass
    
    for col_idx, col in enumerate(ws.columns, 1):
        max_length = 0
        col_letter = get_column_letter(col_idx)
        for cell in col:
            try:
                if cell.value:
                    cell_len = len(str(cell.value))
                    if cell_len > max_length:
                        max_length = cell_len
                cell.alignment = Alignment(horizontal='center', vertical='center')
            except:
                pass
        adjusted_width = max(max_length + 3, 10)
        if adjusted_width > 40:
            adjusted_width = 40
        ws.column_dimensions[col_letter].width = adjusted_width
    
    fd, temp_path = tempfile.mkstemp(suffix='.xlsx')
    os.close(fd)
    wb.save(temp_path)
    wb.close()
    
    resp = send_file(temp_path, as_attachment=True, download_name='Registro_' + pn + '.xlsx')
    resp.headers['Access-Control-Allow-Origin'] = '*'
    
    @resp.call_on_close
    def cleanup():
        try:
            os.remove(temp_path)
        except:
            pass
    
    return resp

@main_bp.route('/api/sync_full', methods=['POST'])
def api_sync_full():
    d = request.json
    pn = d.get('perfil_nombre', '')
    rr = d.get('ruta_remota', '/tftpboot')
    pf = buscar_perfil(pn)
    tarea_id = "sync_{}_{}".format(pn, int(time.time()))
    if not pf:
        return jsonify({'status': 'error', 'message': 'Perfil no encontrado', 'tarea_id': tarea_id})
    
    iniciar_tarea(tarea_id, "Sincronizando {}".format(pn), 100)
    
    def ejecutar_sync():
        try:
            actualizar_progreso(tarea_id, 5, 'Conectando al servidor...')
            actualizar_progreso(tarea_id, 10, 'Obteniendo internos de Asterisk...')
            ints, err = sincronizar_internos_rapido(pf)
            if err or not ints:
                detalle = str(err or 'Asterisk no devolvió internos válidos; la consulta está vacía.')
                mensaje_error = 'Sincronización cancelada sin modificar el Excel: {}'.format(detalle[:180])
                logger.warning('Sync cancelado para %s: %s', pn, mensaje_error)
                error_tarea(tarea_id, mensaje_error)
                return
            
            # === VP SYNC FULL MAC DESDE CONFIGS V1 ===
            # El boton "Sincronizar con Asterisk" primero obtiene los internos de
            # Asterisk y luego cruza solo esos internos con las MAC de /tftpboot.
            mapa_macs_archivos_sync = {}
            def _vp_normalizar_mac_sync(valor):
                return re.sub(r'[^0-9a-f]', '', str(valor or '').lower())
            try:
                actualizar_progreso(tarea_id, 27, 'Buscando archivos de aprovisionamiento y asociando MAC...')
                mapa_archivos_sync = _vp_excel_mapa_mac_remoto(pf)
                internos_asterisk_sync = set(
                    str(k or '').strip() for k in (ints or {}).keys()
                    if re.match(r'^[0-9]{1,10}$', str(k or '').strip())
                )
                for _vp_int_sync, _vp_lista_sync in (mapa_archivos_sync or {}).items():
                    _vp_int_sync = str(_vp_int_sync or '').strip()
                    if _vp_int_sync not in internos_asterisk_sync:
                        continue
                    if not isinstance(_vp_lista_sync, (list, tuple, set)):
                        _vp_lista_sync = [_vp_lista_sync]
                    _vp_validas_sync = []
                    for _vp_mac_sync in _vp_lista_sync:
                        _vp_mac_sync = _vp_normalizar_mac_sync(_vp_mac_sync)
                        if re.match(r'^[0-9a-f]{12}$', _vp_mac_sync) and _vp_mac_sync not in _vp_validas_sync:
                            _vp_validas_sync.append(_vp_mac_sync.upper())
                    if _vp_validas_sync:
                        mapa_macs_archivos_sync[_vp_int_sync] = sorted(_vp_validas_sync)
                    else:
                        mapa_macs_archivos_sync.pop(_vp_int_sync, None)
                logger.info('Sync Index: %s internos de Asterisk cruzados con archivos MAC.', len(mapa_macs_archivos_sync))
            except Exception:
                mapa_macs_archivos_sync = {}
                logger.warning('Sync Index: fallo al consultar archivos MAC; continúa la sincronización Asterisk habitual.', exc_info=True)
            # === FIN VP SYNC FULL MAC DESDE CONFIGS V1 ===

            actualizar_progreso(tarea_id, 30, "Obtenidos {} internos. Consultando IPs...".format(len(ints)))
            db_user, db_pass = obtener_credenciales_mysql(pf)
            cmd_ips = "/usr/bin/mysql -N -B -u {} -p'{}' -e \"SELECT id, data FROM asterisk.sip WHERE keyword='ip' AND data NOT IN ('dynamic', '', '0.0.0.0');\" asterisk 2>/dev/null".format(db_user, db_pass)
            out_ips, _, _ = ejecutar_comando_ssh(pf, cmd_ips)
            ips_dict = {}
            if out_ips:
                for line in out_ips.splitlines():
                    parts = line.split('\t')
                    if len(parts) >= 2:
                        ips_dict[parts[0].strip()] = parts[1].strip()
            
            # Segunda fuente MAC: tabla de vecinos Linux y tablas ARP.
            # Se consultan varios formatos porque cambian entre versiones/distribuciones.
            # La fuente principal sigue siendo el archivo de aprovisionamiento en tftpboot.
            cmd_mac_linux = (
                "printf '__VP_NEIGH_BEGIN__\n'; ip -4 neigh show nud all 2>/dev/null; "
                "printf '\n__VP_ARP_AN_BEGIN__\n'; arp -an 2>/dev/null; "
                "printf '\n__VP_ARP_A_BEGIN__\n'; arp -a 2>/dev/null; "
                "printf '\n__VP_PROC_ARP_BEGIN__\n'; cat /proc/net/arp 2>/dev/null"
            )
            o, err_mac_linux, rc_mac_linux = ejecutar_comando_ssh(pf, cmd_mac_linux)
            macs = {}

            def _vp_mac_agregar_salida_linux(salida_mac):
                for linea_mac in (salida_mac or '').splitlines():
                    linea_mac = str(linea_mac or '').strip()
                    if not linea_mac:
                        continue
                    # arp -a/-an: ? (192.0.2.10) at aa:bb:cc:dd:ee:ff
                    m_ip = re.search(r'\((\d{1,3}(?:\.\d{1,3}){3})\)', linea_mac)
                    # ip neigh y /proc/net/arp comienzan con la IP.
                    if not m_ip:
                        m_ip = re.match(r'^(\d{1,3}(?:\.\d{1,3}){3})(?:\s|$)', linea_mac)
                    m_mac = re.search(r'(?i)(?:[0-9a-f]{2}[:-]){5}[0-9a-f]{2}', linea_mac)
                    if not m_ip or not m_mac:
                        continue
                    ip_mac = m_ip.group(1)
                    mac_mac = m_mac.group(0).replace(':', '').replace('-', '').upper()
                    if mac_mac != '000000000000' and ip_mac not in macs:
                        macs[ip_mac] = mac_mac

            if o:
                _vp_mac_agregar_salida_linux(o)
            logger.info(
                'Sync MAC Linux: %s relaciones IP-MAC detectadas por ip neigh/arp (%s, rc=%s)',
                len(macs), pf.get('host', ''), rc_mac_linux
            )
            if err_mac_linux:
                logger.debug('Consulta MAC Linux con aviso en %s: %s', pf.get('host', ''), str(err_mac_linux)[:200])
            
            actualizar_progreso(tarea_id, 40, 'Procesando Excel remoto...')
            ep = get_excel_file(pn)
            directorio_excel = os.path.dirname(os.path.abspath(ep))
            os.makedirs(directorio_excel, exist_ok=True)
            if os.environ.get('VOIP_PROVISIONER_REMOTE_EXCEL') == '1':
                # Descarga del maestro remoto a una caché temporal escribible.
                # El respaldo se realizará en el servidor remoto durante la publicación.
                _vp_excel_preparar_maestro_remoto(pn, ep, ruta_remota=rr)
            elif os.path.isfile(ep):
                respaldo_sync = ep + '.bak_sync_' + time.strftime('%Y%m%d_%H%M%S')
                shutil.copy2(ep, respaldo_sync)
                if not os.path.isfile(respaldo_sync):
                    raise RuntimeError('No se pudo verificar la copia previa del Excel; sync cancelado.')
                logger.info('Sync: respaldo local previo guardado en %s', respaldo_sync)
            crear_excel_si_no_existe(ep, pn, permitir_vacio=True)
            migrar_excel(ep)
            wb = load_workbook(ep) if os.path.exists(ep) else Workbook()
            ws = wb.active
            if not os.path.exists(ep):
                ws.append(COLUMNAS_EXCEL)
            
            nip = nmac = ni = 0
            ex = {}
            for row in ws.iter_rows(min_row=2, values_only=False):
                if len(row) >= 6 and row[5].value:
                    ex[str(row[5].value).strip()] = row
            
            internos_excel = set()
            for row in ws.iter_rows(min_row=2, values_only=False):
                if len(row) >= 6 and row[5].value:
                    internos_excel.add(str(row[5].value).strip())
            
            internos_a_eliminar = internos_excel - set(ints.keys())
            internos_protegidos = set()
            if internos_a_eliminar:
                filas_a_eliminar = []
                for idx_row, row in enumerate(ws.iter_rows(min_row=2, values_only=False), start=2):
                    if len(row) >= 6 and row[5].value:
                        int_excel = str(row[5].value).strip()
                        if int_excel in internos_a_eliminar:
                            tiene_sn = len(row) > 8 and str(row[8].value or '').strip()
                            if tiene_sn:
                                internos_protegidos.add(int_excel)
                            else:
                                filas_a_eliminar.append(idx_row)
                for fila in sorted(filas_a_eliminar, reverse=True):
                    ws.delete_rows(fila, 1)
                if internos_protegidos:
                    logger.warning('Sync: {} interno(s) no detectados pero con S/N cargado, NO se borraron: {}'.format(len(internos_protegidos), ', '.join(sorted(internos_protegidos))))
            
            actualizar_progreso(tarea_id, 50, 'Actualizando registros...')
            total_ints = len(ints)
            for idx, (i, di) in enumerate(ints.items()):
                ip = ips_dict.get(i, di.get('ip', ''))
                mac = di.get('mac', '')
                nm = di.get('display_name', i)
                _vp_candidatas_mac_sync = mapa_macs_archivos_sync.get(str(i or '').strip(), [])
                # Si los archivos muestran una sola MAC para este interno, esa
                # asociacion es mas fiable que inferirla por una ARP antigua.
                if not mac and len(_vp_candidatas_mac_sync) == 1:
                    mac = _vp_candidatas_mac_sync[0]
                if not mac and ip and ip in macs:
                    _vp_mac_arp_sync = _vp_normalizar_mac_sync(macs[ip])
                    _vp_candidatas_norm_sync = set(_vp_normalizar_mac_sync(x) for x in _vp_candidatas_mac_sync)
                    if not _vp_candidatas_norm_sync or _vp_mac_arp_sync in _vp_candidatas_norm_sync:
                        mac = _vp_mac_arp_sync.upper()
                
                if i in ex:
                    f = ex[i]
                    if len(f) >= 8 and ip:
                        current_ip = str(f[7].value or '').strip()
                        if current_ip != ip:
                            f[7].value = ip
                            nip += 1
                    if len(f) >= 7 and mac:
                        current_mac = str(f[6].value or '').strip().upper().replace(':', '').replace('-', '')
                        clean_mac = mac.upper().replace(':', '').replace('-', '')
                        if not current_mac:
                            f[6].value = mac
                            nmac += 1
                else:
                    ws.append(["", "", "", nm, nm, i, mac, ip, ""])
                    ni += 1
                
                if idx % 10 == 0 or idx == total_ints - 1:
                    progreso = 50 + int((idx / total_ints) * 30)
                    actualizar_progreso(tarea_id, progreso, "Procesando {}/{}...".format(idx+1, total_ints))
            
            # Completar filas existentes con MAC de archivos sin tocar modelo/serie.
            # Primero se aprovecha una correspondencia IP->MAC ya comprobada por ARP;
            # si no existe, solo se asigna por interno cuando la correspondencia es unica.
            try:
                _vp_h_sync = [_vp_excel_header(c.value) for c in ws[1]]
                _vp_ci_sync = _vp_h_sync.index('interno') if 'interno' in _vp_h_sync else 5
                _vp_cm_sync = _vp_h_sync.index('mac') if 'mac' in _vp_h_sync else 6
                _vp_cp_sync = _vp_h_sync.index('ip') if 'ip' in _vp_h_sync else 7
                _vp_cs_sync = _vp_h_sync.index('numero serie') if 'numero serie' in _vp_h_sync else 8
                _vp_cmod_sync = _vp_h_sync.index('modelo telefono') if 'modelo telefono' in _vp_h_sync else 9
                _vp_filas_sync = {}
                for _vp_num_fila_sync in range(2, ws.max_row + 1):
                    _vp_num_sync = str(ws.cell(_vp_num_fila_sync, _vp_ci_sync + 1).value or '').strip()
                    if _vp_num_sync:
                        _vp_filas_sync.setdefault(_vp_num_sync, []).append(_vp_num_fila_sync)

                for _vp_int_sync, _vp_candidatas_sync in mapa_macs_archivos_sync.items():
                    # Segunda barrera: no se incorpora una MAC si Asterisk no lista ese interno.
                    if _vp_int_sync not in internos_asterisk_sync:
                        continue
                    _vp_filas_int_sync = _vp_filas_sync.get(_vp_int_sync, [])
                    if not _vp_filas_int_sync:
                        continue
                    _vp_existentes_sync = set()
                    for _vp_f_sync in _vp_filas_int_sync:
                        _vp_val_mac_sync = _vp_normalizar_mac_sync(ws.cell(_vp_f_sync, _vp_cm_sync + 1).value)
                        if _vp_val_mac_sync:
                            _vp_existentes_sync.add(_vp_val_mac_sync)
                    _vp_pendientes_sync = [
                        _vp_normalizar_mac_sync(m).upper() for m in _vp_candidatas_sync
                        if _vp_normalizar_mac_sync(m) not in _vp_existentes_sync
                    ]
                    if not _vp_pendientes_sync:
                        continue

                    # Vinculo seguro por IP cuando ARP confirma la misma MAC
                    # que ya figura en el archivo de configuracion del interno.
                    for _vp_f_sync in _vp_filas_int_sync:
                        if not _vp_pendientes_sync:
                            break
                        if _vp_normalizar_mac_sync(ws.cell(_vp_f_sync, _vp_cm_sync + 1).value):
                            continue
                        _vp_ip_f_sync = str(ws.cell(_vp_f_sync, _vp_cp_sync + 1).value or '').strip()
                        _vp_mac_ip_sync = _vp_normalizar_mac_sync(macs.get(_vp_ip_f_sync, '')) if _vp_ip_f_sync else ''
                        if _vp_mac_ip_sync and _vp_mac_ip_sync.upper() in _vp_pendientes_sync:
                            ws.cell(_vp_f_sync, _vp_cm_sync + 1).value = _vp_mac_ip_sync.upper()
                            _vp_pendientes_sync.remove(_vp_mac_ip_sync.upper())
                            nmac += 1

                    _vp_filas_vacias_sync = [
                        f for f in _vp_filas_int_sync
                        if not _vp_normalizar_mac_sync(ws.cell(f, _vp_cm_sync + 1).value)
                    ]
                    if len(_vp_pendientes_sync) == 1 and len(_vp_filas_vacias_sync) == 1:
                        # Un solo archivo MAC sin asignar y una sola fila disponible.
                        ws.cell(_vp_filas_vacias_sync[0], _vp_cm_sync + 1).value = _vp_pendientes_sync[0]
                        nmac += 1
                        _vp_pendientes_sync = []
                    elif _vp_pendientes_sync and len(_vp_pendientes_sync) == len(_vp_filas_vacias_sync) and _vp_filas_vacias_sync:
                        # Orden estable solo si las filas vacias no tienen IP, modelo o serie
                        # manual; asi no se cambia una asociacion administrativa existente.
                        _vp_sin_datos_manuales_sync = all(
                            not str(ws.cell(f, _vp_cs_sync + 1).value or '').strip() and
                            not str(ws.cell(f, _vp_cmod_sync + 1).value or '').strip() and
                            not str(ws.cell(f, _vp_cp_sync + 1).value or '').strip()
                            for f in _vp_filas_vacias_sync
                        )
                        if _vp_sin_datos_manuales_sync:
                            for _vp_f_sync, _vp_mac_sync in zip(sorted(_vp_filas_vacias_sync), sorted(_vp_pendientes_sync)):
                                ws.cell(_vp_f_sync, _vp_cm_sync + 1).value = _vp_mac_sync
                                nmac += 1
                            _vp_pendientes_sync = []

                    if _vp_pendientes_sync:
                        # No reasignar una MAC a una fila con datos manuales ambiguos.
                        logger.warning(
                            'Sync Index: MAC pendientes para interno %s: %s; se preservan modelo/serie/IP por seguridad.',
                            _vp_int_sync, ','.join(sorted(_vp_pendientes_sync))
                        )
            except Exception:
                logger.warning('Sync Index: no se pudieron completar todas las MAC desde configs.', exc_info=True)

            actualizar_progreso(tarea_id, 85, 'Completando modelos de teléfono...')
            try:
                modelos_sync = _vp_excel_modelos_config_cache(pf)
                cambios_modelo_sync = _vp_excel_aplicar_modelos_hoja(ws, modelos_sync)
                if cambios_modelo_sync:
                    logger.info('Sync %s: completados %s modelos/marcas por configuración/OUI.',
                                pn, cambios_modelo_sync)
            except Exception:
                logger.warning('Sync %s: no se pudieron completar los modelos/marcas.', pn, exc_info=True)
            actualizar_progreso(tarea_id, 88, 'Guardando Excel...')
            aplicar_formato_excel(wb, ws)
            wb.save(ep)
            wb.close()
            actualizar_progreso(tarea_id, 90, 'Publicando Excel maestro en el servidor...')
            subir_excel_remoto(
                pn, rr, source_path=ep,
                backup=(os.environ.get('VOIP_PROVISIONER_REMOTE_EXCEL') == '1'),
                raise_on_error=(os.environ.get('VOIP_PROVISIONER_REMOTE_EXCEL') == '1')
            )
            
            detalles = "{} nuevos, {} IPs actualizadas, {} MACs actualizadas".format(ni, nip, nmac)
            if internos_a_eliminar:
                detalles += ", {} obsoletos eliminados".format(len(internos_a_eliminar))
            
            completar_tarea(tarea_id, "Sync: {}".format(detalles))
        except Exception as e:
            error_tarea(tarea_id, "Error: {}: {}".format(e.__class__.__name__, str(e)[:600]))
            logger.exception("Error en sync para perfil %s", pn)
    
    thread = threading.Thread(target=ejecutar_sync, daemon=True)
    thread.start()
    return jsonify({'status': 'ok', 'message': 'Sincronizacion iniciada', 'tarea_id': tarea_id})


@main_bp.route('/api/interno_issabel', methods=['POST'])
def api_interno_issabel():
    """
    Consulta silenciosamente si un interno ya existe en el Issabel
    seleccionado.

    NUNCA devuelve la contraseña.
    """

    try:
        d = request.json or {}

        perfil_nombre = str(
            d.get('perfil_nombre', '') or ''
        ).strip()

        interno = str(
            d.get('interno', '') or ''
        ).strip()

        if not perfil_nombre:
            return jsonify({
                'status': 'error',
                'message': 'Perfil no indicado'
            })

        if not interno:
            return jsonify({
                'status': 'ok',
                'existe': False,
                'tecnologia': '',
                'tiene_password': False
            })

        pf = buscar_perfil(perfil_nombre)

        if not pf:
            return jsonify({
                'status': 'error',
                'message': 'Perfil no encontrado'
            })

        if os.environ.get('VOIP_PROVISIONER_REMOTE_EXCEL') == '1':
            _vp_excel_preparar_maestro_remoto(perfil_nombre, get_excel_file(perfil_nombre))
        resultado = obtener_password_interno_issabel(
            pf,
            interno
        )

        # ----------------------------------------------------
        # IMPORTANTE:
        # nunca incluir resultado['password']
        # ----------------------------------------------------
        return jsonify({
            'status': 'ok',
            'existe': bool(resultado.get('existe')),
            'tecnologia': resultado.get('tecnologia', ''),
            'tiene_password': bool(resultado.get('tiene_password')),
            'macs': (
                macs_asignadas_interno(get_excel_file(perfil_nombre), interno)
                if resultado.get('existe') else []
            )
        })

    except Exception as e:
        logger.exception(
            "Error en /api/interno_issabel: %s",
            e
        )

        return jsonify({
            'status': 'error',
            'message': 'Error consultando el interno'
        })


@main_bp.route('/api/generar', methods=['POST'])
@main_bp.route('/api/editar', methods=['POST'])
def api_generar():
    d = request.json or {}
    cfg = cargar_config()

    # Edicion de datos: no ejecutar aprovisionamiento de telefonos.
    if request.path.rstrip('/').endswith('/api/editar'):
        if d.get('password') != cfg.get('seguridad', {}).get('admin_password'):
            return jsonify({'status': 'error', 'message': 'Contraseña de administrador incorrecta'}), 401

        perfil_nombre = d.get('perfil_nombre', '')
        pf_edicion = buscar_perfil(perfil_nombre)
        if not pf_edicion:
            return jsonify({'status': 'error', 'message': 'Perfil no encontrado'}), 404

        ep_edicion = get_excel_file(perfil_nombre)
        if os.environ.get('VOIP_PROVISIONER_REMOTE_EXCEL') == '1':
            try:
                _vp_excel_preparar_maestro_remoto(perfil_nombre, ep_edicion)
            except Exception as exc:
                return jsonify({'status': 'error', 'message': 'No se pudo leer el Excel remoto: {}'.format(str(exc)[:220])}), 503
        if not os.path.exists(ep_edicion):
            return jsonify({'status': 'error', 'message': 'No existe el Excel del perfil seleccionado'}), 404

        wb_edicion = load_workbook(ep_edicion)
        try:
            ws_edicion = wb_edicion.active
            encabezados = [c.value for c in ws_edicion[1]]
            if 'Interno' not in encabezados:
                return jsonify({'status': 'error', 'message': 'Falta la columna Interno en el Excel'}), 400

            interno_buscar = str(d.get('interno_original', '') or d.get('interno', '')).strip()
            if not interno_buscar:
                return jsonify({'status': 'error', 'message': 'Falta el interno a editar'}), 400

            # VP_NOMBRES_COMPARTIDOS_V1
            idx_interno = encabezados.index('Interno')
            idx_mac = encabezados.index('MAC') if 'MAC' in encabezados else None
            filas_interno = []

            for fila in ws_edicion.iter_rows(min_row=2):
                if (len(fila) > idx_interno and
                    str(fila[idx_interno].value or '').strip() == interno_buscar):
                    filas_interno.append(fila)

            if not filas_interno:
                return jsonify({'status': 'error', 'message': 'No se encontro el interno en el Excel'}), 404

            def normalizar_mac_edicion(valor):
                return ''.join(
                    c.lower() for c in str(valor or '')
                    if c.lower() in '0123456789abcdef'
                )

            mac_original = normalizar_mac_edicion(
                d.get('mac_original') or d.get('nueva_mac') or ''
            )
            fila_encontrada = None

            # Los datos propios del telefono se guardan en su fila,
            # identificada por MAC, aunque el interno este compartido.
            if idx_mac is not None and mac_original:
                for fila in filas_interno:
                    if (len(fila) > idx_mac and
                        normalizar_mac_edicion(fila[idx_mac].value) == mac_original):
                        fila_encontrada = fila
                        break

            if fila_encontrada is None:
                fila_encontrada = filas_interno[0]

            # Los datos particulares (edificio, piso, modelo) solo
            # cambian en la fila/MAC seleccionada.
            cambios = {
                'Edificio': 'nuevo_edificio',
                'Piso': 'nuevo_piso',
                'Dependencia': 'nueva_dependencia',
                'Modelo Telefono': 'nuevo_modelo'
            }

            for columna, clave in cambios.items():
                if columna in encabezados and clave in d:
                    fila_encontrada[encabezados.index(columna)].value = str(d.get(clave) or '').strip()

            actualizar_mostrar = d.get('actualizar_nombre_mostrar') is True
            actualizar_saliente = d.get('actualizar_nombre_saliente') is True

            # La identidad pertenece al interno compartido, no a una MAC.
            # Propagar cada nombre modificado a todas sus filas del Excel.
            if actualizar_mostrar and 'Nombre Mostrar' in encabezados:
                idx_nombre = encabezados.index('Nombre Mostrar')
                nombre = str(d.get('nuevo_nombre_mostrar') or '').strip() or interno_buscar
                for fila in filas_interno:
                    fila[idx_nombre].value = nombre

            if actualizar_saliente and 'Nombre Saliente' in encabezados:
                idx_saliente = encabezados.index('Nombre Saliente')
                nombre = str(d.get('nuevo_nombre_saliente') or '').strip() or interno_buscar
                for fila in filas_interno:
                    fila[idx_saliente].value = nombre
            mensaje_edicion = 'Datos guardados en el Excel del perfil'
            if actualizar_mostrar or actualizar_saliente:
                ok_nombres, err_nombres = actualizar_nombres_interno_asterisk(
                    pf_edicion,
                    interno_buscar,
                    d.get('nuevo_nombre_mostrar', ''),
                    d.get('nuevo_nombre_saliente', ''),
                    actualizar_mostrar=actualizar_mostrar,
                    actualizar_saliente=actualizar_saliente
                )
                if not ok_nombres:
                    return jsonify({
                        'status': 'error',
                        'message': 'No se confirmó la actualización en Issabel. El Excel todavía no se guardó. ' + str(err_nombres)
                    }), 502
                mensaje_edicion = 'Datos guardados en Excel y nombres actualizados en Issabel.'

            wb_edicion.save(ep_edicion)
        finally:
            wb_edicion.close()

        try:
            subir_excel_remoto(
                perfil_nombre,
                d.get('ruta_remota', pf_edicion.get('ruta_remota', '/tftpboot')),
                source_path=ep_edicion, backup=True, raise_on_error=True
            )
        except Exception as e:
            logger.exception("No se pudo guardar la edición del Excel en el servidor remoto.")
            return jsonify({
                'status': 'error',
                'message': 'La edición no quedó guardada en el servidor remoto: {}'.format(str(e)[:220])
            }), 502

        return jsonify({'status': 'ok', 'message': mensaje_edicion})

    pf = buscar_perfil(d.get('perfil_nombre', ''))
    tarea_id = "gen_{}_{}".format(d.get('interno', ''), int(time.time()))
    if not pf:
        return jsonify({'status': 'error', 'message': 'Perfil no encontrado', 'tarea_id': tarea_id})
    
    interno = str(d.get('interno', '')).strip()
    ml = REGEX_LIMPIAR_MAC.sub('', d.get('mac', '').upper())
    sn = d.get('numero_serie', '').strip().upper()
    
    ep = get_excel_file(d.get('perfil_nombre', ''))
    try:
        if os.environ.get('VOIP_PROVISIONER_REMOTE_EXCEL') == '1':
            _vp_excel_preparar_maestro_remoto(d.get('perfil_nombre', ''), ep)
        crear_excel_si_no_existe(ep, d.get('perfil_nombre', ''), permitir_vacio=True)
        migrar_excel(ep)
    except Exception as e:
        logger.exception('No se pudo preparar el Excel antes de generar el interno')
        return jsonify({
            'status': 'error',
            'message': 'No se pudo preparar el Excel del perfil. No se continuó el alta para proteger el inventario. Detalle: {}'.format(str(e)[:220])
        }), 500
    
    if ml or sn:
        error_tipo, interno_conflicto = verificar_duplicados_mac_sn(ep, interno, ml, sn)
        if error_tipo == "MAC_DUPLICADA":
            return jsonify({
                'status': 'error',
                'message': 'La MAC {} ya está asignada al interno {}. No se pueden crear dos internos con la misma MAC.'.format(ml, interno_conflicto),
                'tarea_id': tarea_id
            })
        elif error_tipo == "SN_DUPLICADO":
            return jsonify({
                'status': 'error',
                'message': 'El número de serie {} ya está asignado al interno {}. Cada equipo debe tener un S/N único.'.format(sn, interno_conflicto),
                'tarea_id': tarea_id
            })
    
    plts = cargar_plantillas()
    modelo_req = str(d.get('modelo', '') or '').strip()
    modelo_req_lower = modelo_req.lower()

    es_grandstream_req = (
        'grandstream' in modelo_req_lower or
        'gxp' in modelo_req_lower or
        'grp' in modelo_req_lower
    )

    pl = plts.get(modelo_req)

    # Grandstream utiliza exclusivamente su plantilla XML.
    # No debe depender de que exista un .cfg.
    if ml and not es_grandstream_req and not pl:
        return jsonify({
            'status': 'error',
            'message': 'Modelo no disponible',
            'tarea_id': tarea_id
        })
    
    # --------------------------------------------------------
    # Contraseña para CREACION INDIVIDUAL
    #
    # Si el usuario eligio usar la contraseña ya existente en
    # Issabel, la obtenemos DIRECTAMENTE desde el servidor
    # seleccionado.
    #
    # La contraseña nunca viaja al navegador.
    # --------------------------------------------------------
    usar_password_issabel = bool(
        d.get('usar_password_issabel', False)
    )

    tecnologia_individual = str(
        d.get('tecnologia', '') or pf.get('tecnologia', 'sip')
    ).strip().lower()

    if tecnologia_individual not in ('sip', 'pjsip'):
        tecnologia_individual = 'sip'

    pw = d.get('password', '').strip()

    if usar_password_issabel:
        resultado_issabel = obtener_password_interno_issabel(
            pf,
            interno
        )

        if not resultado_issabel.get('existe'):
            return jsonify({
                'status': 'error',
                'message': 'El interno ya no existe en Issabel o no pudo ser consultado.',
                'tarea_id': tarea_id
            })

        password_issabel = str(
            resultado_issabel.get('password', '') or ''
        ).strip()

        if not password_issabel:
            return jsonify({
                'status': 'error',
                'message': 'El interno existe en Issabel pero no tiene una contraseña almacenada.',
                'tarea_id': tarea_id
            })

        # Si estamos reutilizando el interno existente,
        # actuamos con la tecnologia que realmente tiene Issabel.
        tecnologia_detectada = str(
            resultado_issabel.get('tecnologia', '') or ''
        ).strip().lower()

        if tecnologia_detectada in ('sip', 'pjsip'):
            tecnologia_individual = tecnologia_detectada

        pw = password_issabel

        # Nunca registrar pw.
        logger.info(
            "Interno %s: reutilizando contraseña existente de Issabel "
            "tecnologia=%s host=%s",
            interno,
            tecnologia_individual,
            pf.get('host', '')
        )

    else:
        if not pw:
            pw = "pjn{}".format(interno)

    # --------------------------------------------------------
    # SEGUNDO TELEFONO (PARALELO) - automatico, solo PJSIP.
    #
    # Se activa solo si el interno ya existe en Issabel como
    # PJSIP y el Excel ya le tiene asignada OTRA MAC.
    # Si la MAC es la misma es un reaprovisionamiento normal.
    # SIP tradicional conserva el comportamiento actual.
    # --------------------------------------------------------
    es_paralelo = False
    if ml:
        macs_previas = macs_asignadas_interno(ep, interno)
        if macs_previas and ml.lower() not in macs_previas:
            res_par = obtener_password_interno_issabel(pf, interno)
            tec_par = str(res_par.get('tecnologia', '') or '').strip().lower()
            if res_par.get('existe') and tec_par == 'pjsip':
                if len(macs_previas) >= 2:
                    return jsonify({
                        'status': 'error',
                        'message': 'El interno {} ya tiene 2 telefonos (maximo). Elegi otro numero.'.format(interno),
                        'tarea_id': tarea_id
                    })
                pw_par = str(res_par.get('password', '') or '').strip()
                if not pw_par:
                    return jsonify({
                        'status': 'error',
                        'message': 'El interno {} existe pero no tiene contraseña almacenada.'.format(interno),
                        'tarea_id': tarea_id
                    })
                es_paralelo = True
                pw = pw_par
                tecnologia_individual = 'pjsip'
                logger.info(
                    "Interno %s: segundo telefono (paralelo) host=%s",
                    interno, pf.get('host', '')
                )

    base_url_global = request.host_url.rstrip('/')
    
    iniciar_tarea(tarea_id, "Generando interno {}".format(interno), 100)
    
    def ejecutar_generar():
        try:
            if es_paralelo:
                actualizar_progreso(tarea_id, 10, 'Habilitando segundo telefono (paralelo)...')
                ok, err = habilitar_paralelo_pjsip(pf, interno)
            else:
                actualizar_progreso(tarea_id, 10, 'Actualizando peer en Asterisk...')
                ok, err = crear_interno_individual_en_asterisk(
                    pf,
                    interno,
                    pw,
                    'from-internal',
                    d.get('nombre_mostrar', ''),
                    d.get('nombre_saliente', ''),
                    tecnologia_individual
                )

            if not ok:
                logger.warning(
                    "Error en crear_interno_individual_en_asterisk: {}".format(err)
                )
                error_tarea(
                    tarea_id,
                    "Error creando interno {}: {}".format(
                        interno,
                        str(err)[:180]
                    )
                )
                return
            
            if ml:
                actualizar_progreso(
                    tarea_id,
                    50,
                    'Generando archivo de configuracion...'
                )

                mac_limpia = ml.lower()

                modelo_cfg = str(
                    modelo_req or d.get('modelo') or ''
                ).strip()

                modelo_cfg_lower = modelo_cfg.lower()

                es_grandstream = (
                    'grandstream' in modelo_cfg_lower or
                    'gxp' in modelo_cfg_lower or
                    'grp' in modelo_cfg_lower
                )

                # ====================================================
                # GRANDSTREAM
                # ====================================================
                #
                # Grandstream utiliza exclusivamente XML.
                #
                # La plantilla se toma de:
                #
                # /opt/voip-provisioner/data/plantillas/
                #
                # y es exactamente la que fue cargada desde Admin.
                #
                if es_grandstream:

                    from xml.sax.saxutils import escape as _xesc

                    nombre_cfg = (
                        'cfg' + mac_limpia.lower() + '.xml'
                    )

                    # Nombre del modelo ingresado en Admin.
                    _modelo_admin = str(
                        modelo_req or ''
                    ).strip()

                    if _modelo_admin.lower().endswith('.cfg'):
                        _modelo_admin = _modelo_admin[:-4]

                    if _modelo_admin.lower().endswith('.xml'):
                        _modelo_admin = _modelo_admin[:-4]

                    # Normalización solamente para encontrar
                    # la plantilla correspondiente.
                    # El CONTENIDO utilizado será exactamente
                    # el que esté guardado en Admin.
                    def _normalizar_nombre_plantilla(_v):
                        return ''.join(
                            ch.lower()
                            for ch in str(_v or '')
                            if ch.isalnum()
                        )

                    _modelo_normalizado = (
                        _normalizar_nombre_plantilla(
                            _modelo_admin
                        )
                    )

                    _tpl_xml = None

                    _carpetas_plantillas_xml = [
                        PLANTILLAS_FOLDER,
                        os.environ.get('VOIP_PROVISIONER_BUNDLED_PLANTILLAS', ''),
                        os.environ.get('VOIP_PROVISIONER_BUNDLED_PLANTILLAS_EXTRA', ''),
                        os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'data', 'plantillas'),
                        os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'plantillas')
                    ]
                    _carpetas_vistas_xml = set()
                    for _carpeta_xml in _carpetas_plantillas_xml:
                        if not _carpeta_xml:
                            continue
                        _carpeta_xml = os.path.abspath(_carpeta_xml)
                        if _carpeta_xml in _carpetas_vistas_xml:
                            continue
                        _carpetas_vistas_xml.add(_carpeta_xml)
                        if not os.path.isdir(_carpeta_xml):
                            continue
                        try:
                            _archivos_xml = sorted(os.listdir(_carpeta_xml), key=lambda _x: _x.lower())
                        except Exception:
                            continue
                        for _fn in _archivos_xml:
                            if not _fn.lower().endswith('.xml'):
                                continue
                            _base_fn = _fn[:-4]
                            if _normalizar_nombre_plantilla(_base_fn) == _modelo_normalizado:
                                _cand = os.path.join(_carpeta_xml, _fn)
                                if os.path.isfile(_cand):
                                    _tpl_xml = _cand
                                    break
                        if _tpl_xml:
                            break
                    if not _tpl_xml:
                        raise RuntimeError(
                            "Grandstream: no existe una plantilla XML "
                            "guardada desde Administracion para el "
                            "modelo '{}' en {}".format(
                                _modelo_admin,
                                PLANTILLAS_FOLDER
                            )
                        )

                    try:
                        with open(
                            _tpl_xml,
                            'r',
                            encoding='utf-8'
                        ) as _fx:
                            _t = _fx.read()

                        # IMPORTANTE:
                        # no generamos ningun XML propio.
                        # se utiliza exactamente el contenido
                        # cargado desde Admin.
                        cont = _t.format(
                            mac=mac_limpia.lower(),
                            interno=_xesc(str(interno)),
                            password=_xesc(str(pw)),
                            nombre_mostrar=_xesc(
                                str(
                                    d.get(
                                        'nombre_mostrar',
                                        ''
                                    ) or interno
                                )
                            ),
                            servidor_sip=_xesc(
                                str(pf['host'])
                            ),
                            servidor_tftp=_xesc(
                                str(pf['host'])
                            ),
                            base_url=base_url_global
                        )

                    except Exception as e_x:
                        raise RuntimeError(
                            "Plantilla XML Grandstream invalida "
                            "'{}': {}".format(
                                _tpl_xml,
                                e_x
                            )
                        )

                    logger.info(
                        "Grandstream: usando plantilla XML de Admin %s",
                        _tpl_xml
                    )

                # ====================================================
                # RESTO DE MODELOS
                # ====================================================
                #
                # Este es el flujo anterior para CFG.
                # Fanvil queda intacto.
                #
                else:

                    try:
                        cont = pl.format(
                            interno=interno,
                            servidor_sip=pf['host'],
                            password=pw,
                            base_url=base_url_global,
                            mac=mac_limpia,
                            nombre_mostrar=(
                                d.get(
                                    'nombre_mostrar',
                                    ''
                                ) or interno
                            )
                        )

                    except Exception as e_fmt:
                        logger.error(
                            "Error al formatear plantilla: {}".format(
                                e_fmt
                            )
                        )
                        cont = pl

                    # Fanvil: el servidor de Flash/Autoprovisioning
                    # siempre debe ser el host del perfil seleccionado.
                    cont = re.sub(
                        r'(?m)^\s*ap\.FlashServerIP\s*=.*$',
                        'ap.FlashServerIP = {}'.format(
                            pf['host']
                        ),
                        cont
                    )

                    nombre_cfg = (
                        mac_limpia + '.cfg'
                    )

                rl = os.path.join(
                    CARPETA_CONFIGS,
                    nombre_cfg
                )

                with open(
                    rl,
                    "w",
                    encoding='utf-8'
                ) as f:
                    f.write(cont)

                actualizar_progreso(
                    tarea_id,
                    70,
                    'Subiendo configuracion al servidor...'
                )

                try:
                    s, _ = conectar_sftp(
                        pf['host'],
                        int(pf.get('puerto', 22)),
                        pf['usuario'],
                        pf['password'],
                        pf.get(
                            'ruta_remota',
                            '/tftpboot'
                        ),
                        pf.get(
                            'ssh_compat',
                            False
                        )
                    )

                    with s:
                        s.put(
                            rl,
                            "{}/{}".format(
                                pf.get(
                                    'ruta_remota',
                                    '/tftpboot'
                                ).rstrip('/'),
                                os.path.basename(rl)
                            )
                        )

                except Exception as e:
                    logger.warning(
                        "Error subiendo config: {}".format(e)
                    )

            actualizar_progreso(tarea_id, 85, 'Actualizando Excel...')
            wb = load_workbook(ep) if os.path.exists(ep) else Workbook()
            ws = wb.active
            if not os.path.exists(ep):
                ws.append(COLUMNAS_EXCEL)
            h = [c.value for c in ws[1]]
            tn = "Nombre Mostrar" in h
            ci = h.index("Interno") if "Interno" in h else (5 if tn else 3)
            cm = h.index("MAC") if "MAC" in h else (6 if tn else 4)
            csn = h.index("Numero Serie") if "Numero Serie" in h else 8
            cmodelo = h.index("Modelo Telefono") if "Modelo Telefono" in h else 9
            encontrado = False
            fila_base = None
            if es_paralelo:
                # Segundo telefono: la fila del telefono original no se toca.
                for row in ws.iter_rows(min_row=2):
                    if len(row) > ci and str(row[ci].value or '').strip() == interno:
                        fila_base = row
                        break

            def _base_val(col):
                if fila_base is not None and len(fila_base) > col:
                    return fila_base[col].value or ''
                return ''

            for row in ws.iter_rows(min_row=2):
                if len(row) > ci and not es_paralelo:
                    row_int = str(row[ci].value or '').strip()
                    if row_int == interno:
                        row[0].value = d.get('edificio', '') or row[0].value
                        row[1].value = d.get('piso', '') or row[1].value
                        row[2].value = d.get('dependencia', '') or row[2].value

                        if tn:
                            row[h.index("Nombre Mostrar")].value = (
                                d.get('nombre_mostrar', '') or
                                row[h.index("Nombre Mostrar")].value or interno
                            )
                            row[h.index("Nombre Saliente")].value = (
                                d.get('nombre_saliente', '') or
                                row[h.index("Nombre Saliente")].value or interno
                            )

                        if ml:
                            row[cm].value = ml
                        if sn:
                            row[csn].value = sn
                        if modelo_req:
                            row[cmodelo].value = modelo_req

                        encontrado = True
                        break

            if not encontrado:
                nueva_fila = [''] * max(len(h), cmodelo + 1)
                nueva_fila[0] = d.get('edificio', '') or _base_val(0)
                nueva_fila[1] = d.get('piso', '') or _base_val(1)
                nueva_fila[2] = d.get('dependencia', '') or _base_val(2)

                if "Nombre Mostrar" in h:
                    nueva_fila[h.index("Nombre Mostrar")] = (
                        d.get('nombre_mostrar', '') or _base_val(h.index("Nombre Mostrar")) or interno
                    )
                if "Nombre Saliente" in h:
                    nueva_fila[h.index("Nombre Saliente")] = (
                        d.get('nombre_saliente', '') or _base_val(h.index("Nombre Saliente")) or interno
                    )

                nueva_fila[ci] = interno
                nueva_fila[cm] = ml
                nueva_fila[csn] = sn
                nueva_fila[cmodelo] = modelo_req
                ws.append(nueva_fila)

            aplicar_formato_excel(wb, ws)
            wb.save(ep)
            wb.close()
            actualizar_progreso(tarea_id, 95, 'Publicando Excel maestro en el servidor...')
            subir_excel_remoto(
                d.get('perfil_nombre', ''), d.get('ruta_remota', '/tftpboot'),
                source_path=ep, backup=(os.environ.get('VOIP_PROVISIONER_REMOTE_EXCEL') == '1'),
                raise_on_error=(os.environ.get('VOIP_PROVISIONER_REMOTE_EXCEL') == '1')
            )
            
            actualizar_progreso(tarea_id, 98, 'Aplicando cambios en Issabel...')
            pf_reload = buscar_perfil(d.get('perfil_nombre', ''))
            if pf_reload:
                recargar_issabel_completo(pf_reload)
            
            if es_paralelo:
                mensaje = "Segundo teléfono agregado al interno {} (paralelo) y aplicado".format(interno)
            elif encontrado:
                mensaje = "Interno {} actualizado y aplicado".format(interno)
            else:
                mensaje = "Interno {} generado y aplicado".format(interno)
                if sn:
                    mensaje += " (S/N: {})".format(sn)
            completar_tarea(tarea_id, mensaje)
        except Exception as e:
            error_tarea(tarea_id, "Error: {}".format(str(e)[:100]))
            logger.error("Error en generar: {}".format(e))
    
    thread = threading.Thread(target=ejecutar_generar, daemon=True)
    thread.start()
    return jsonify({'status': 'ok', 'message': 'Generacion iniciada', 'tarea_id': tarea_id})

MAX_INTERNOS_RANGO = 200

_LOCK_RANGO = threading.Lock()


@main_bp.route('/api/generar_rango', methods=['POST'])
def api_generar_rango():
    # Carga masiva de internos por rango.
    #
    # - Los internos que ya existen (tabla o Issabel) NO se modifican.
    # - La creacion en Issabel se hace por tandas (crear_internos_lote_en_asterisk).
    # - Los internos que existen en Issabel pero faltan en la tabla se suman a la tabla.
    # - Se guarda un CSV de registro (permisos 600) en data/cargas_masivas/.
    # - Issabel se recarga UNA vez, al final, con recargar_issabel_completo().
    d = request.json or {}
    cfg = cargar_config()
    desde_txt = str(d.get('desde', '')).strip()
    hasta_txt = str(d.get('hasta', '')).strip()
    tarea_id = "rango_{}_{}".format(desde_txt, int(time.time()))

    def _error(msg):
        return jsonify({'status': 'error', 'message': msg, 'tarea_id': tarea_id})

    if d.get('password_admin') != cfg['seguridad']['admin_password']:
        return _error('Contrasena incorrecta')

    pn = d.get('perfil_nombre', '')
    pf = buscar_perfil(pn)
    if not pf:
        return _error('Perfil no encontrado')

    if not (re.fullmatch(r'[0-9]{1,10}', desde_txt) and re.fullmatch(r'[0-9]{1,10}', hasta_txt)):
        return _error('Rango invalido: usar solo numeros')
    desde = int(desde_txt)
    hasta = int(hasta_txt)
    if desde <= 0 or hasta < desde:
        return _error('Rango invalido: "hasta" debe ser mayor o igual a "desde"')

    limite_rango = 200
    if hasta - desde + 1 > limite_rango:
        return _error('Maximo {} internos por rango'.format(limite_rango))

    ancho = len(desde_txt) if desde_txt.startswith('0') else 0
    internos = [str(n).zfill(ancho) for n in range(desde, hasta + 1)]

    modo_pw = str(d.get('modo_password', 'pjn') or 'pjn').strip().lower()
    texto_pw = str(d.get('password_sip', '') or '').strip()
    if modo_pw not in ('pjn', 'prefijo', 'fija'):
        return _error('Modo de contrasena invalido')
    if modo_pw in ('prefijo', 'fija') and not re.fullmatch(r'[A-Za-z0-9_.@#%+=:-]{1,40}', texto_pw):
        return _error('Contrasena SIP: 1 a 40 caracteres (letras, numeros y _ . @ # % + = : -)')

    tecnologia = str(d.get('tecnologia', '') or pf.get('tecnologia', 'sip')).strip().lower()
    if tecnologia not in ('sip', 'pjsip'):
        return _error('Tecnologia invalida (sip o pjsip)')
    pf = dict(pf)
    pf['tecnologia'] = tecnologia

    def _pw_para(it):
        if modo_pw == 'prefijo':
            return texto_pw + it
        if modo_pw == 'fija':
            return texto_pw
        return "pjn{}".format(it)

    edificio = str(d.get('edificio', '') or '').strip()
    piso = str(d.get('piso', '') or '').strip()
    dependencia = str(d.get('dependencia', '') or '').strip()
    rr = d.get('ruta_remota', '/tftpboot')

    if not _LOCK_RANGO.acquire(False):
        return _error('Ya hay una carga masiva en curso. Esperar a que termine.')

    iniciar_tarea(tarea_id, "Carga masiva {}-{}".format(desde_txt, hasta_txt), 100)

    def ejecutar_rango():
        try:
            import shutil
            import csv

            ep = get_excel_file(pn)
            if os.environ.get('VOIP_PROVISIONER_REMOTE_EXCEL') == '1':
                _vp_excel_preparar_maestro_remoto(pn, ep, ruta_remota=rr)
            crear_excel_si_no_existe(ep, pn, permitir_vacio=True)
            migrar_excel(ep)

            def _internos_en_excel():
                # Devuelve el conjunto de internos del Excel, o None si el formato no sirve.
                _wb = load_workbook(ep)
                try:
                    _ws = _wb.active
                    _h = [c.value for c in _ws[1]]
                    if not (len(_h) >= 8 and _h[3] == "Nombre Mostrar"):
                        return None
                    return set(str(r[5]).strip()
                               for r in _ws.iter_rows(min_row=2, values_only=True)
                               if len(r) > 5 and r[5])
                finally:
                    _wb.close()

            en_tabla = _internos_en_excel()
            if en_tabla is None:
                error_tarea(tarea_id, "Formato de Excel no soportado")
                return

            # ---------------- CSV de registro (permisos 600) ----------------
            carpeta_csv = os.path.join(os.path.dirname(ep), "cargas_masivas")
            csv_path = os.path.join(carpeta_csv, "rango_{}_{}_{}.csv".format(
                desde_txt, hasta_txt, time.strftime("%Y%m%d_%H%M%S")))
            csv_ok = {'ok': False}

            def _escribir_csv(estados):
                try:
                    if not os.path.isdir(carpeta_csv):
                        os.makedirs(carpeta_csv)
                    os.chmod(carpeta_csv, 0o700)
                    fd = os.open(csv_path, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
                    with os.fdopen(fd, 'w', newline='', encoding='utf-8') as fh:
                        w = csv.writer(fh)
                        w.writerow(['interno', 'password', 'tecnologia', 'callerid', 'contexto', 'estado'])
                        for it in internos:
                            w.writerow([it, _pw_para(it), tecnologia, 'device <{}>'.format(it),
                                        'from-internal', estados.get(it, 'PENDIENTE')])
                    csv_ok['ok'] = True
                except Exception as e_csv:
                    logger.warning("RANGO: no se pudo escribir el CSV: {}".format(e_csv))

            omitidos = [it for it in internos if it in en_tabla]
            candidatos = [it for it in internos if it not in en_tabla]

            estados = {}
            for it in internos:
                estados[it] = 'EXISTENTE_TABLA' if it in en_tabla else 'PENDIENTE'
            _escribir_csv(estados)

            # ---------------- Creacion en Issabel (por tandas) ----------------
            actualizar_progreso(tarea_id, 10, 'Verificando internos existentes en Issabel...')
            if candidatos:
                resultado = crear_internos_lote_en_asterisk(
                    pf, candidatos, _pw_para, 'from-internal', tecnologia,
                    progreso=lambda pct, msg: actualizar_progreso(tarea_id, pct, msg))
            else:
                resultado = {'creados': [], 'existentes': [], 'fallidos': [],
                             'no_procesados': [], 'error': ''}

            creados = resultado.get('creados', [])
            existentes_db = resultado.get('existentes', [])
            fallidos = resultado.get('fallidos', [])
            no_procesados = resultado.get('no_procesados', [])
            detalle = resultado.get('error', '')

            for it in existentes_db:
                estados[it] = 'EXISTENTE_ISSABEL'
            for it in creados:
                estados[it] = 'CREADO'
            for it in fallidos:
                estados[it] = 'ERROR'
            for it in no_procesados:
                estados[it] = 'NO_PROCESADO'

            # ---------------- Tabla (Excel): releer justo antes de escribir ----------------
            para_tabla = list(creados) + list(existentes_db)
            sumados = []
            if para_tabla:
                actualizar_progreso(tarea_id, 88, 'Guardando Excel...')
                try:
                    if os.path.exists(ep) and os.environ.get('VOIP_PROVISIONER_REMOTE_EXCEL') != '1':
                        shutil.copy2(ep, ep + '.bak_rango')
                    ya = _internos_en_excel()
                    if ya is None:
                        raise Exception("formato de Excel no soportado")
                    wb = load_workbook(ep)
                    try:
                        ws = wb.active
                        for it in para_tabla:
                            if it in ya:
                                continue
                            ws.append([edificio, piso, dependencia, it, it, it, '', '', ''])
                            sumados.append(it)
                        if sumados:
                            aplicar_formato_excel(wb, ws)
                            wb.save(ep)
                    finally:
                        wb.close()
                    if sumados:
                        verificados = _internos_en_excel() or set()
                        faltan = [x for x in sumados if x not in verificados]
                        if faltan:
                            raise Exception("{} internos no quedaron en el Excel".format(len(faltan)))
                        subir_excel_remoto(
                            pn, rr, source_path=ep,
                            backup=(os.environ.get('VOIP_PROVISIONER_REMOTE_EXCEL') == '1'),
                            raise_on_error=(os.environ.get('VOIP_PROVISIONER_REMOTE_EXCEL') == '1')
                        )
                except Exception as e_xl:
                    logger.error("RANGO: error guardando Excel: {}".format(e_xl), exc_info=True)
                    _escribir_csv(estados)
                    error_tarea(tarea_id, (
                        "Los internos se procesaron en Issabel pero no se pudo guardar el Excel ({}). "
                        "Usar 'Sincronizar con Asterisk' para traerlos."
                    ).format(str(e_xl)[:80]))
                    return

            # ---------------- Recarga de Issabel (una sola vez) ----------------
            aviso = ''
            if creados:
                actualizar_progreso(tarea_id, 95, 'Aplicando cambios en Issabel...')
                try:
                    recargar_issabel_completo(pf)
                except Exception as e_rl:
                    logger.error("RANGO: error recargando Issabel: {}".format(e_rl), exc_info=True)
                    aviso = " ATENCION: fallo la recarga de Issabel ({}); aplicarla manualmente.".format(str(e_rl)[:60])

            _escribir_csv(estados)

            # ---------------- Resultado ----------------
            partes = ["{} creados ({})".format(len(creados), tecnologia.upper())]
            if omitidos:
                partes.append("{} ya estaban en la tabla".format(len(omitidos)))
            if existentes_db:
                partes.append("{} ya existian en Issabel (sumados a la tabla)".format(len(existentes_db)))
            if fallidos:
                partes.append("{} con error".format(len(fallidos)))
            if no_procesados:
                partes.append("{} sin procesar".format(len(no_procesados)))
            mensaje = "Carga {}-{}: {}.".format(desde_txt, hasta_txt, ", ".join(partes))
            if detalle:
                mensaje += " Detalle: {}".format(detalle)
            if csv_ok['ok']:
                mensaje += " Registro: {}".format(os.path.basename(csv_path))
            mensaje += aviso

            logger.info("RANGO {}-{} tec={} excel={}: {}".format(desde_txt, hasta_txt, tecnologia, ep, mensaje))

            if (fallidos or no_procesados) and not creados and not sumados:
                error_tarea(tarea_id, mensaje)
            else:
                completar_tarea(tarea_id, mensaje)

        except Exception as e:
            logger.error("RANGO: excepcion: {}".format(e), exc_info=True)
            error_tarea(tarea_id, "Error: {}".format(str(e)[:150]))
        finally:
            _LOCK_RANGO.release()

    try:
        threading.Thread(target=ejecutar_rango, daemon=True).start()
    except Exception:
        _LOCK_RANGO.release()
        raise
    return jsonify({'status': 'ok', 'message': 'Carga masiva iniciada', 'tarea_id': tarea_id})
@main_bp.route('/api/eliminar', methods=['POST'])
def api_eliminar():
    d = request.json
    cfg = cargar_config()
    tarea_id = "del_{}_{}".format(d.get('interno', ''), int(time.time()))
    if d.get('password') != cfg['seguridad']['admin_password']:
        return jsonify({'status': 'error', 'message': 'Contrasena incorrecta', 'tarea_id': tarea_id})
    
    iniciar_tarea(tarea_id, "Eliminando interno {}".format(d.get("interno", "")), 100)
    
    def ejecutar_eliminar():
        try:
            actualizar_progreso(tarea_id, 20, 'Procesando Excel...')
            ml = REGEX_LIMPIAR_MAC.sub('', d.get('mac', '').upper())
            nombre_perfil_borrado = str(d.get('perfil_nombre', '') or '').strip()
            ep = get_excel_file(nombre_perfil_borrado)
            if os.environ.get('VOIP_PROVISIONER_REMOTE_EXCEL') == '1':
                _vp_excel_preparar_maestro_remoto(nombre_perfil_borrado, ep, ruta_remota=d.get('ruta_remota', ''))
            if not _vp_excel_tiene_registros(ep):
                _vp_borrar_sin_excel(
                    nombre_perfil_borrado,
                    buscar_perfil(nombre_perfil_borrado),
                    [{'interno': str(d.get('interno', '') or '').strip(), 'mac': ml}],
                    d.get('ruta_remota', ''),
                    tarea_id
                )
                return
            migrar_excel(ep)
            wb = load_workbook(ep)
            ws = wb.active
            h = [c.value for c in ws[1]]
            tn = len(h) >= 8 and h[3] == "Nombre Mostrar"
            hn = [str(v or '').strip().lower() for v in h]
            ci = hn.index("interno") if "interno" in hn else (5 if tn else 3)
            cm = hn.index("mac") if "mac" in hn else (6 if tn else 4)
            fe = None

            for idx, row in enumerate(ws.iter_rows(min_row=2, values_only=False), 2):
                if len(row) > cm and len(row) > ci:
                    ie = str(row[ci].value or '').strip()
                    me = REGEX_LIMPIAR_MAC.sub(
                        '', str(row[cm].value or '').upper()
                    )
                    if ie == str(d.get('interno', '')).strip() and (
                        (ml and me == ml) or (not ml)
                    ):
                        fe = idx
                        break

            # Si se solicita borrar una MAC concreta, exigir la fila exacta.
            # Nunca borrar CFG ni tocar Issabel si no existe esa fila.
            if ml and not fe:
                wb.close()
                error_tarea(
                    tarea_id,
                    "No se encontró la fila exacta del teléfono (interno y MAC); no se eliminó nada."
                )
                logger.warning(
                    "Eliminacion cancelada: interno %s MAC %s no encontrado en Excel",
                    d.get('interno', ''), ml
                )
                return

            # (parche) antes, si no habia fila en el Excel (interno que solo existe
            # en Issabel / Monitoreo) no se hacia nada y la tarea nunca terminaba.
            if True:
                if fe:
                    ws.delete_rows(fe, 1)
                    aplicar_formato_excel(wb, ws)
                    wb.save(ep)
                # Si quedan otras filas del mismo interno (segundo telefono en
                # paralelo), el interno NO se elimina de Asterisk.
                interno_del = str(d.get('interno', '')).strip()
                ci_del = ci
                quedan_otros = False
                for _r in ws.iter_rows(min_row=2):
                    if len(_r) > ci_del and str(_r[ci_del].value or '').strip() == interno_del:
                        quedan_otros = True
                        break
                wb.close()
                actualizar_progreso(tarea_id, 50, 'Eliminando configuracion...')
                if ml:
                    try:
                        lp_local = os.path.join(CARPETA_CONFIGS, "{}.cfg".format(ml.lower()))
                        lp_local_xml = os.path.join(CARPETA_CONFIGS, "cfg{}.xml".format(ml.lower()))
                        if os.path.exists(lp_local_xml):
                            os.remove(lp_local_xml)
                        if os.path.exists(lp_local):
                            os.remove(lp_local)
                        pf = buscar_perfil(d.get('perfil_nombre', ''))
                        if pf:
                            rr = d.get('ruta_remota', '') or pf.get('ruta_remota', '/tftpboot')
                            cfg_remoto = "{}/{}.cfg".format(rr, ml.lower())
                            cfg_remoto_xml = "{}/cfg{}.xml".format(rr, ml.lower())
                            if _vp_host_es_local(pf['host']):
                                if os.path.exists(cfg_remoto):
                                    os.remove(cfg_remoto)
                                if os.path.exists(cfg_remoto_xml):
                                    os.remove(cfg_remoto_xml)
                            else:
                                cmd_borrar = "rm -f {} {} 2>&1".format(cfg_remoto, cfg_remoto_xml)
                                ejecutar_comando_ssh(pf, cmd_borrar)
                    except Exception as e:
                        logger.error("Error al borrar CFG: " + str(e))
                
                actualizar_progreso(tarea_id, 70, 'Eliminando de Asterisk...')
                pf = buscar_perfil(d.get('perfil_nombre', ''))
                if pf and not quedan_otros:
                    eliminar_interno_asterisk(pf, d.get('interno', ''))
                elif pf:
                    logger.info("Interno %s: queda otro telefono en el Excel; se conserva en Asterisk", interno_del)
                
                actualizar_progreso(tarea_id, 90, 'Subiendo Excel actualizado...')
                subir_excel_remoto(
                    d.get('perfil_nombre', ''), d.get('ruta_remota', '/tftpboot'),
                    source_path=ep, backup=(os.environ.get('VOIP_PROVISIONER_REMOTE_EXCEL') == '1'),
                    raise_on_error=(os.environ.get('VOIP_PROVISIONER_REMOTE_EXCEL') == '1')
                )
                
                pf_reload = buscar_perfil(d.get('perfil_nombre', ''))
                if pf_reload:
                    recargar_issabel_completo(pf_reload)
                
                if quedan_otros:
                    completar_tarea(tarea_id, "Teléfono del interno {} eliminado (el interno sigue activo con otro equipo)".format(d.get("interno", "")))
                else:
                    completar_tarea(tarea_id, "Interno {} eliminado y aplicado".format(d.get("interno", "")))
        except Exception as e:
            error_tarea(tarea_id, "Error: {}".format(str(e)[:100]))
            logger.error("Error en eliminar: {}".format(e))
    
    thread = threading.Thread(target=ejecutar_eliminar, daemon=True)
    thread.start()
    return jsonify({'status': 'ok', 'message': 'Eliminacion iniciada', 'tarea_id': tarea_id})

@main_bp.route('/api/eliminar_varios', methods=['POST'])
def api_eliminar_varios():
    d = request.json or {}
    cfg = cargar_config()
    tarea_id = "delvarios_{}".format(int(time.time()))

    if d.get('password') != cfg['seguridad']['admin_password']:
        return jsonify({'status': 'error', 'message': 'Contraseña incorrecta'})

    nombre_perfil = d.get('perfil_nombre', '')
    pf = buscar_perfil(nombre_perfil)
    if not pf:
        return jsonify({'status': 'error', 'message': 'Perfil no encontrado'})

    regs = d.get('registros', [])
    if not regs:
        return jsonify({'status': 'error', 'message': 'No hay registros seleccionados'})

    # Normalizar y validar la selección antes de iniciar cualquier borrado.
    seleccion = []
    vistos = set()
    for r in regs:
        interno = str(r.get('interno', '')).strip()
        mac = re.sub(r'[^0-9a-f]', '', str(r.get('mac', '')).lower())
        if not interno:
            return jsonify({'status': 'error', 'message': 'Hay un interno inválido'})
        if mac and not re.match(r'^[0-9a-f]{12}$', mac):
            return jsonify({'status': 'error', 'message': 'Hay una MAC inválida'})
        try:
            fila_excel = int(r.get('fila_excel') or 0)
        except (TypeError, ValueError):
            fila_excel = 0
        if fila_excel < 2:
            fila_excel = 0

        clave = (interno, mac, fila_excel)
        if clave not in vistos:
            vistos.add(clave)
            seleccion.append({
                'interno': interno,
                'mac': mac,
                'fila_excel': fila_excel
            })

    iniciar_tarea(
        tarea_id,
        "Eliminando {} registros seleccionados".format(len(seleccion)),
        100
    )

    def ejecutar_eliminar_varios():
        wb = None
        try:
            actualizar_progreso(tarea_id, 10, 'Verificando registros...')
            ep = get_excel_file(nombre_perfil)
            if os.environ.get('VOIP_PROVISIONER_REMOTE_EXCEL') == '1':
                _vp_excel_preparar_maestro_remoto(nombre_perfil, ep, ruta_remota=d.get('ruta_remota', ''))
            if not _vp_excel_tiene_registros(ep):
                _vp_borrar_sin_excel(
                    nombre_perfil,
                    pf,
                    seleccion,
                    d.get('ruta_remota', ''),
                    tarea_id
                )
                return
            migrar_excel(ep)
            wb = load_workbook(ep)
            ws = wb.active
            encabezados = [
                str(c.value or '').strip().lower()
                for c in ws[1]
            ]

            if 'interno' not in encabezados or 'mac' not in encabezados:
                raise Exception('El Excel no contiene las columnas Interno y MAC')

            ci = encabezados.index('interno')
            cm = encabezados.index('mac')

            filas = {}
            for idx, row in enumerate(ws.iter_rows(min_row=2), 2):
                interno_fila = str(row[ci].value or '').strip()
                mac_fila = re.sub(
                    r'[^0-9a-f]', '',
                    str(row[cm].value or '').lower()
                )
                filas.setdefault((interno_fila, mac_fila), []).append(idx)

            # Prevalidar TODAS las filas antes de modificar el Excel.
            # Con MAC: exigir coincidencia exacta interno + MAC.
            # Sin MAC: permitir solo si el interno tiene una única fila.
            filas_borrar = []
            claves_borrar = set()

            for item in seleccion:
                interno_sel = item['interno']
                mac_sel = item['mac']
                fila_excel = item.get('fila_excel', 0)

                if fila_excel:
                    if fila_excel >= ws.max_row + 1:
                        raise Exception(
                            'La fila seleccionada ya no existe; no se borró ningún registro'
                        )
                    valor_interno = str(ws.cell(fila_excel, ci + 1).value or '').strip()
                    valor_mac = re.sub(
                        r'[^0-9a-f]', '',
                        str(ws.cell(fila_excel, cm + 1).value or '').lower()
                    )
                    if valor_interno != interno_sel or (mac_sel and valor_mac != mac_sel):
                        raise Exception(
                            'La fila seleccionada cambió. Actualice la tabla e intente nuevamente'
                        )
                    candidatas = [fila_excel]
                elif mac_sel:
                    candidatas = filas.get((interno_sel, mac_sel), [])
                else:
                    candidatas = []
                    for (interno_fila, mac_fila), indices in filas.items():
                        if interno_fila == interno_sel:
                            candidatas.extend(indices)

                if len(candidatas) != 1:
                    raise Exception(
                        'El interno {} no identifica una única fila. '
                        'Actualice la tabla y seleccione nuevamente; no se borró '
                        'ningún registro'.format(interno_sel)
                    )

                fila = candidatas[0]
                if fila in claves_borrar:
                    raise Exception(
                        'La selección contiene una fila duplicada; '
                        'no se borró ningún registro'
                    )

                claves_borrar.add(fila)
                filas_borrar.append(fila)

            actualizar_progreso(tarea_id, 25, 'Eliminando filas seleccionadas...')
            for idx in sorted(filas_borrar, reverse=True):
                ws.delete_rows(idx, 1)

            aplicar_formato_excel(wb, ws)
            wb.save(ep)
            wb.close()
            wb = None

            # Consultar las filas restantes por interno. Solo eliminar de
            # Asterisk los internos que ya no tengan ningún teléfono en Excel.
            wb = load_workbook(ep, read_only=True, data_only=True)
            ws = wb.active
            restantes = set()
            for row in ws.iter_rows(min_row=2, values_only=True):
                if len(row) > ci:
                    valor = str(row[ci] or '').strip()
                    if valor:
                        restantes.add(valor)
            wb.close()
            wb = None

            actualizar_progreso(tarea_id, 45, 'Eliminando configuraciones...')
            import shlex
            rr = d.get('ruta_remota', '') or pf.get('ruta_remota', '/tftpboot')
            es_local = _vp_host_es_local(pf.get('host'))

            for item in seleccion:
                mac = item['mac']
                if not mac:
                    continue

                for ruta in (
                    os.path.join(CARPETA_CONFIGS, mac + '.cfg'),
                    os.path.join(CARPETA_CONFIGS, 'cfg' + mac + '.xml')
                ):
                    if os.path.exists(ruta):
                        os.remove(ruta)

                remotos = [
                    rr.rstrip('/') + '/' + mac + '.cfg',
                    rr.rstrip('/') + '/cfg' + mac + '.xml'
                ]
                if es_local:
                    for ruta in remotos:
                        if os.path.exists(ruta):
                            os.remove(ruta)
                else:
                    comando = 'rm -f ' + ' '.join(
                        shlex.quote(ruta) for ruta in remotos
                    )
                    ejecutar_comando_ssh(pf, comando)

            actualizar_progreso(tarea_id, 65, 'Actualizando Asterisk...')
            for interno in sorted(set(x['interno'] for x in seleccion)):
                if interno not in restantes:
                    eliminar_interno_asterisk(pf, interno)
                else:
                    logger.info(
                        'Se conserva el interno %s: quedan otros teléfonos asignados',
                        interno
                    )

            actualizar_progreso(tarea_id, 85, 'Aplicando cambios...')
            subir_excel_remoto(
                nombre_perfil, rr, source_path=ep,
                backup=(os.environ.get('VOIP_PROVISIONER_REMOTE_EXCEL') == '1'),
                raise_on_error=(os.environ.get('VOIP_PROVISIONER_REMOTE_EXCEL') == '1')
            )
            recargar_issabel_completo(pf)

            completar_tarea(
                tarea_id,
                '{} registro(s) eliminado(s); los internos con otros teléfonos '
                'asignados se conservaron'.format(len(seleccion))
            )

        except Exception as e:
            if wb is not None:
                try:
                    wb.close()
                except Exception:
                    pass
            error_tarea(tarea_id, "Error: {}".format(str(e)[:180]))
            logger.error("Error en eliminar_varios: {}".format(e))

    thread = threading.Thread(target=ejecutar_eliminar_varios, daemon=True)
    thread.start()
    return jsonify({
        'status': 'ok',
        'message': 'Eliminación iniciada',
        'tarea_id': tarea_id
    })

@main_bp.route('/api/logs/audit')
def api_audit_logs():
    logs = []

    try:
        res = subprocess.run(
            [
                'journalctl',
                '-u', 'voip-provisioner',
                '-n', '500',
                '--no-pager',
                '-o', 'short-iso'
            ],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            universal_newlines=True,
            timeout=5
        )

        for line in res.stdout.splitlines():
            if not line.strip():
                continue

            # Ignorar solamente tráfico interno que ensucia los eventos
            if 'GET /api/progreso' in line:
                continue
            if 'heartbeat' in line.lower():
                continue

            # Separar timestamp del mensaje
            partes = line.split(' ', 1)

            if len(partes) == 2:
                timestamp = partes[0]
                mensaje = partes[1]
            else:
                timestamp = datetime.now().isoformat()
                mensaje = line

            texto = mensaje.strip()

            tipo = 'info'
            accion = 'Sistema'
            estado = 'success'

            texto_lower = texto.lower()

            if (
                'error' in texto_lower
                or 'exception' in texto_lower
                or 'traceback' in texto_lower
                or 'failed' in texto_lower
            ):
                tipo = 'error'
                accion = 'Error del sistema'
                estado = 'error'

            elif (
                'editar:' in texto_lower
                or 'edit_' in texto_lower
                or 'edicion' in texto_lower
            ):
                tipo = 'edit'
                accion = 'Edicion de interno'

            elif (
                'generar:' in texto_lower
                or 'gen_' in texto_lower
                or 'generando interno' in texto_lower
                or 'generacion' in texto_lower
            ):
                tipo = 'create'
                accion = 'Generacion de interno'

            elif (
                'sync' in texto_lower
                or 'sincronizando' in texto_lower
                or 'sincronizacion' in texto_lower
            ):
                tipo = 'sync'
                accion = 'Sincronizacion con Asterisk'

            elif (
                'eliminando' in texto_lower
                or 'del_' in texto_lower
                or 'eliminado' in texto_lower
                or 'eliminacion' in texto_lower
            ):
                tipo = 'delete'
                accion = 'Eliminacion de interno'

            elif 'tabla:' in texto_lower:
                tipo = 'info'
                accion = 'Consulta de tabla'

            logs.append({
                'timestamp': timestamp,
                'tipo': tipo,
                'accion': accion,
                'detalles': texto[:500],
                'estado': estado
            })

    except Exception as e:
        logger.error("Error leyendo logs: " + str(e))

    # Los más recientes primero
    logs.sort(
        key=lambda x: x.get('timestamp', ''),
        reverse=True
    )

    limit = int(request.args.get('limit', '100'))

    return jsonify({
        'status': 'ok',
        'logs': logs[:limit]
    })



@main_bp.route('/api/clear_logs', methods=['POST'])
@main_bp.route('/api/logs/clear', methods=['POST'])
def api_clear_logs():
    try:
        data = request.json or {}
        password = data.get('password', '')
        cfg = cargar_config()
        if password != cfg['seguridad']['admin_password']:
            return jsonify({'status': 'error', 'message': 'Contrasena incorrecta'})
        
        subprocess.run(['journalctl', '--rotate'], stdout=subprocess.PIPE, stderr=subprocess.PIPE, universal_newlines=True)
        subprocess.run(['journalctl', '--vacuum-time=1s'], stdout=subprocess.PIPE, stderr=subprocess.PIPE, universal_newlines=True)
        
        log_file = '/var/log/voip-provisioner.log'
        if os.path.exists(log_file):
            open(log_file, 'w').close()
        
        return jsonify({'status': 'ok', 'message': 'Logs limpiados correctamente'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})

@main_bp.route('/api/stream')
def api_stream():
    pn = request.args.get('perfil_nombre', '')

    # Al conectar el Monitor, refrescar el Excel sin bloquear el SSE.
    if pn:
        try:
            threading.Thread(
                target=_vp_excel_actualizar_al_conectar,
                args=(pn,),
                daemon=True
            ).start()
        except Exception:
            logger.exception('No se pudo iniciar la actualizacion del Excel del Monitor')

    def gen():
        cq = queue.Queue()

        with monitor_lock:
            if pn not in profile_monitors:
                t = threading.Thread(
                    target=monitor_asterisk,
                    args=(pn, False),
                    daemon=True
                )

                profile_monitors[pn] = {
                    'thread': t,
                    'queues': [cq],
                    'ultimo_estado': None
                }

                t.start()

            else:
                profile_monitors[pn]['queues'].append(cq)

        try:
            while True:
                try:
                    yield "data: " + json.dumps(
                        cq.get(timeout=15)
                    ) + "\n\n"

                except queue.Empty:
                    yield ": heartbeat\n\n"

        except GeneratorExit:
            with monitor_lock:

                data = profile_monitors.get(pn)

                if data:
                    try:
                        data.get('queues', []).remove(cq)
                    except:
                        pass

                    # Si ya no queda ningun navegador conectado
                    # al monitor de este perfil, detener el monitor.
                    if not data.get('queues'):
                        profile_monitors.pop(pn, None)

    resp = Response(
        gen(),
        mimetype="text/event-stream"
    )

    resp.headers['Cache-Control'] = 'no-cache'
    resp.headers['Access-Control-Allow-Origin'] = '*'

    return resp

@main_bp.route('/api/alarmas')
def api_alarmas():
    return jsonify(ALARMAS_PERFILES.get(request.args.get('perfil_nombre', ''), []))

@main_bp.route('/api/alarmas/clear', methods=['POST'])
def api_alarmas_clear():
    ALARMAS_PERFILES[request.args.get('perfil_nombre', '')] = []
    return jsonify({"status": "ok"})

@main_bp.route('/api/monitoreo')
def api_monitoreo_api():
    pn = request.args.get('perfil_nombre', '')
    pf = buscar_perfil(pn)
    if not pf:
        return jsonify({'status': 'error'})
    
    est = obtener_estado_completo(pf, pn)
    peers_online = sum(
        1 for p in est['peers']
        if str(p.get('status', '') or '').lower() in ('online', 'ok')
        or str(p.get('status_raw', '') or '').lower() == 'ok'
    )
    peers_offline = len(est['peers']) - peers_online
    
    return jsonify({
        'totales': {
            'canales_activos': len(est['canales']),
            'peers_online': peers_online,
            'peers_offline': peers_offline,
            'total_peers': len(est['peers'])
        },
        'canales': est['canales'],
        'troncales': est['troncales'],
        'peers': est['peers']
    })



# ===== BLOQUE RUTA HISTORICA CDR =====

@main_bp.route('/api/monitoreo/colgar', methods=['POST'])
def api_monitoreo_colgar():
    """
    Finaliza una llamada activa utilizando los canales
    reales proporcionados por el monitor.
    """
    try:
        datos = request.get_json(silent=True) or {}

        perfil_nombre = str(
            datos.get('perfil_nombre') or
            datos.get('perfil') or
            ''
        ).strip()

        canales = datos.get('canales')

        if not perfil_nombre:
            return jsonify({
                "ok": False,
                "mensaje": "Falta el perfil"
            }), 400

        if not isinstance(canales, list) or not canales:
            return jsonify({
                "ok": False,
                "mensaje": "No se recibieron canales de la llamada"
            }), 400

        perfil = None

        try:
            perfiles = cargar_perfiles()
        except Exception:
            perfiles = []

        for pf in perfiles:
            nombre_pf = str(
                pf.get("nombre") or
                pf.get("name") or
                ''
            ).strip()

            if nombre_pf == perfil_nombre:
                perfil = pf
                break

        if perfil is None:
            return jsonify({
                "ok": False,
                "mensaje": "Perfil no encontrado"
            }), 404

        resultado = colgar_canales(
            perfil,
            canales
        )

        if not resultado.get("ok"):
            return jsonify(resultado), 500

        return jsonify(resultado)

    except Exception as e:
        logger.exception(
            "Error intentando colgar llamada"
        )

        return jsonify({
            "ok": False,
            "mensaje": str(e)
        }), 500


@main_bp.route('/api/monitoreo/cdr/ruta', methods=['GET'])
def api_monitoreo_cdr_ruta():
    """
    Reconstruye la ruta histórica de una llamada usando linkedid.

    Identifica:
      SIP
      PJSIP
      IAX2
      OOH323
      DAHDI
      ZAP
      LOCAL
      DONGLE
      CUSTOM

    Para OOH323/IAX2 intenta identificar además el nombre de la troncal.
    """

    try:
        pn = request.args.get('perfil_nombre', '')
        uniqueid = str(
            request.args.get('uniqueid', '') or ''
        ).strip()

        if not uniqueid:
            return jsonify({
                'status': 'error',
                'message': 'Falta uniqueid'
            }), 400

        pf = buscar_perfil(pn)

        if not pf:
            return jsonify({
                'status': 'error',
                'message': 'Perfil no encontrado'
            }), 404

        def sql_escape(valor):
            return (
                str(valor)
                .replace('\\', '\\\\')
                .replace("'", "''")
            )

        uid = sql_escape(uniqueid)

        # -------------------------------------------------
        # Compatibilidad con CDR antiguos
        #
        # Algunas versiones de Asterisk/Issabel pueden no
        # tener linkedid, sequence o peeraccount.
        # Detectamos las columnas antes de utilizarlas.
        # -------------------------------------------------

        columnas_cdr = set()

        try:
            out_cols, err_cols, code_cols = _ejecutar_sql_cdr(
                pf,
                "SHOW COLUMNS FROM cdr"
            )

            if code_cols == 0 and out_cols:
                for linea_col in str(out_cols).splitlines():
                    partes_col = linea_col.split('\t')
                    if partes_col:
                        nombre_col = str(partes_col[0]).strip().lower()
                        if nombre_col:
                            columnas_cdr.add(nombre_col)

        except Exception:
            columnas_cdr = set()

        tiene_linkedid = 'linkedid' in columnas_cdr
        tiene_sequence = 'sequence' in columnas_cdr
        tiene_peeraccount = 'peeraccount' in columnas_cdr

        # -------------------------------------------------
        # Buscar CDR seleccionado
        # -------------------------------------------------

        columnas_base = [
            'calldate',
            'clid',
            'src',
            'dst',
            'dcontext',
            'channel',
            'dstchannel',
            'lastapp',
            'lastdata',
            'duration',
            'billsec',
            'disposition',
            'uniqueid'
        ]

        if tiene_linkedid:
            columnas_base.append('linkedid')

        if tiene_peeraccount:
            columnas_base.append('peeraccount')

        if tiene_sequence:
            columnas_base.append('sequence')

        sql = (
            "SELECT " +
            ",".join(columnas_base) +
            " FROM cdr "
            "WHERE uniqueid = '%s' "
            "LIMIT 1"
        ) % uid

        def normalizar_resultado_sql(resultado):
            # _ejecutar_sql_cdr devuelve:
            #   (stdout, stderr, returncode)
            if isinstance(resultado, tuple):
                if len(resultado) >= 3:
                    stdout = resultado[0]
                    stderr = resultado[1]
                    codigo = resultado[2]

                    if codigo != 0:
                        return []

                    texto = str(stdout or '').strip()

                    if not texto:
                        return []

                    lineas = texto.splitlines()

                    # La consulta usa mysql -N -B:
                    # no hay encabezados, solamente filas TSV.
                    columnas = [
                        'calldate',
                        'clid',
                        'src',
                        'dst',
                        'dcontext',
                        'channel',
                        'dstchannel',
                        'lastapp',
                        'lastdata',
                        'duration',
                        'billsec',
                        'disposition',
                        'uniqueid',
                        'linkedid',
                        'peeraccount',
                        'sequence'
                    ]

                    filas = []

                    for linea in lineas:
                        valores = linea.split('\t')

                        if len(valores) < len(columnas):
                            valores += [''] * (
                                len(columnas) - len(valores)
                            )

                        fila = {}

                        for i, columna in enumerate(columnas):
                            fila[columna] = valores[i]

                        filas.append(fila)

                    return filas

            if isinstance(resultado, dict):
                if 'rows' in resultado:
                    return resultado['rows']

                if 'data' in resultado:
                    return resultado['data']

                return [resultado]

            if isinstance(resultado, list):
                return resultado

            return []


        resultado = _ejecutar_sql_cdr(pf, sql)
        resultado = normalizar_resultado_sql(resultado)

        if isinstance(resultado, dict):
            if 'rows' in resultado:
                resultado = resultado['rows']
            elif 'data' in resultado:
                resultado = resultado['data']
            else:
                resultado = [resultado]

        if not isinstance(resultado, list):
            resultado = []

        if not resultado:
            return jsonify({
                'status': 'ok',
                'encontrada': False,
                'ruta': []
            })

        principal = resultado[0]

        linkedid = str(
            principal.get('linkedid') or
            principal.get('uniqueid') or
            ''
        ).strip()

        # -------------------------------------------------
        # Buscar todas las patas de la llamada
        # -------------------------------------------------

        lid = sql_escape(linkedid)

        columnas_ruta = [
            'calldate',
            'clid',
            'src',
            'dst',
            'dcontext',
            'channel',
            'dstchannel',
            'lastapp',
            'lastdata',
            'duration',
            'billsec',
            'disposition',
            'uniqueid'
        ]

        if tiene_linkedid:
            columnas_ruta.append('linkedid')

        if tiene_peeraccount:
            columnas_ruta.append('peeraccount')

        if tiene_sequence:
            columnas_ruta.append('sequence')

        if tiene_linkedid and linkedid:
            condicion_ruta = "linkedid = '%s'" % lid
        else:
            condicion_ruta = "uniqueid = '%s'" % uid

        orden_ruta = "calldate ASC"

        if tiene_sequence:
            orden_ruta += ", sequence ASC"

        orden_ruta += ", uniqueid ASC"

        sql = (
            "SELECT " +
            ",".join(columnas_ruta) +
            " FROM cdr "
            "WHERE " + condicion_ruta + " "
            "ORDER BY " + orden_ruta
        )

        registros = _ejecutar_sql_cdr(pf, sql)
        registros = normalizar_resultado_sql(registros)

        if not registros:
            registros = [principal]

        # -------------------------------------------------
        # Tecnología
        # -------------------------------------------------

        def tecnologia(canal):
            c = str(canal or '').upper().strip()

            if c.startswith('SIP/'):
                return 'SIP'

            if c.startswith('PJSIP/'):
                return 'PJSIP'

            if c.startswith('IAX2/'):
                return 'IAX2'

            if c.startswith('OOH323/'):
                return 'OOH323'

            if c.startswith('DAHDI/'):
                return 'DAHDI'

            if c.startswith('ZAP/'):
                return 'ZAP'

            if c.startswith('LOCAL/'):
                return 'LOCAL'

            if c.startswith('DONGLE/'):
                return 'DONGLE'

            if c.startswith('CUSTOM/'):
                return 'CUSTOM'

            if '/' in c:
                return c.split('/', 1)[0]

            return 'OTRO'

        # -------------------------------------------------
        # Extraer peer / troncal
        # -------------------------------------------------

        def peer_de_canal(canal):
            c = str(canal or '').strip()

            if '/' not in c:
                return c

            resto = c.split('/', 1)[1]

            # Ejemplos:
            #
            # SIP/57320-00000104
            # OOH323/ipoffice-229
            # IAX2/voipcentral1-123
            #
            if '-' in resto:
                resto = resto.rsplit('-', 1)[0]

            return resto

        def es_troncal(tech):
            return tech in (
                'OOH323',
                'IAX2',
                'CUSTOM',
                'DAHDI',
                'ZAP',
                'DONGLE'
            )

        def descripcion(tech):
            return {
                'SIP': 'Canal SIP / interno',
                'PJSIP': 'Canal PJSIP',
                'IAX2': 'Troncal IAX2',
                'OOH323': 'Troncal H.323 / OOH323',
                'DAHDI': 'DAHDI / posible FXO',
                'ZAP': 'ZAP / DAHDI / posible FXO',
                'LOCAL': 'Canal Local',
                'DONGLE': 'Gateway GSM / Dongle',
                'CUSTOM': 'Troncal Custom'
            }.get(
                tech,
                'Canal ' + tech
            )

        # -------------------------------------------------
        # Construcción de ruta
        # -------------------------------------------------

        ruta = []
        vistos = set()

        def agregar_paso(row, canal, origen):
            if not canal:
                return

            canal = str(canal).strip()

            tech = tecnologia(canal)
            peer = peer_de_canal(canal)

            clave = (
                str(row.get('uniqueid') or ''),
                canal,
                origen
            )

            if clave in vistos:
                return

            vistos.add(clave)

            troncal = ''

            if es_troncal(tech):
                troncal = peer

            paso = {
                'orden': len(ruta) + 1,
                'tecnologia': tech,
                'canal': canal,
                'peer': peer,
                'troncal': troncal,
                'es_troncal': bool(troncal),
                'descripcion': descripcion(tech),
                'contexto': str(
                    row.get('dcontext') or ''
                ),
                'aplicacion': str(
                    row.get('lastapp') or ''
                ),
                'datos': str(
                    row.get('lastdata') or ''
                ),
                'origen': str(
                    row.get('src') or ''
                ),
                'destino': str(
                    row.get('dst') or ''
                ),
                'resultado': str(
                    row.get('disposition') or ''
                ),
                'uniqueid': str(
                    row.get('uniqueid') or ''
                ),
                'linkedid': str(
                    row.get('linkedid') or ''
                ),
                'sequence': row.get(
                    'sequence',
                    0
                )
            }

            ruta.append(paso)

        for row in registros:

            agregar_paso(
                row,
                row.get('channel'),
                'channel'
            )

            agregar_paso(
                row,
                row.get('dstchannel'),
                'dstchannel'
            )

        # -------------------------------------------------
        # Orden lógico:
        # canal de origen -> troncal -> destino
        # -------------------------------------------------

        def prioridad(paso):
            tech = paso.get('tecnologia')

            if tech in ('SIP', 'PJSIP'):
                return 10

            if tech == 'LOCAL':
                return 20

            if tech in (
                'OOH323',
                'IAX2',
                'CUSTOM'
            ):
                return 30

            if tech in (
                'DAHDI',
                'ZAP',
                'DONGLE'
            ):
                return 40

            return 50

        ruta.sort(
            key=lambda x: (
                prioridad(x),
                x.get('sequence', 0),
                x.get('orden', 0)
            )
        )

        # Renumerar
        for i, paso in enumerate(ruta, 1):
            paso['orden'] = i

        # -------------------------------------------------
        # Tecnologías y troncales encontradas
        # -------------------------------------------------

        tecnologias = []
        troncales = []

        for paso in ruta:

            tech = paso.get('tecnologia')
            troncal = paso.get('troncal')

            if tech and tech not in tecnologias:
                tecnologias.append(tech)

            if troncal and troncal not in troncales:
                troncales.append(troncal)

        # -------------------------------------------------
        # Resumen de ruta
        # -------------------------------------------------

        resumen = []

        for paso in ruta:

            if paso.get('es_troncal'):
                resumen.append(
                    'TRONCAL ' +
                    paso.get('tecnologia', '') +
                    '/' +
                    paso.get('troncal', '')
                )
            else:
                resumen.append(
                    paso.get('canal', '')
                )

        return jsonify({
            'status': 'ok',
            'encontrada': True,

            'uniqueid': str(
                principal.get('uniqueid') or ''
            ),

            'linkedid': linkedid,

            'calldate': str(
                principal.get('calldate') or ''
            ),

            'src': str(
                principal.get('src') or ''
            ),

            'dst': str(
                principal.get('dst') or ''
            ),

            'dcontext': str(
                principal.get('dcontext') or ''
            ),

            'disposition': str(
                principal.get('disposition') or ''
            ),

            'duration': principal.get(
                'duration',
                0
            ),

            'billsec': principal.get(
                'billsec',
                0
            ),

            'registros': len(registros),

            'tecnologias': tecnologias,

            'troncales': troncales,

            'resumen_ruta': resumen,

            'ruta': ruta,

            'cdr': registros
        })

    except Exception as e:

        import traceback
        traceback.print_exc()

        return jsonify({
            'status': 'error',
            'message': str(e)
        }), 500


@main_bp.route('/api/monitoreo/ruta')
def api_monitoreo_ruta():
    """
    Devuelve la ruta real de una llamada activa asociada a un interno.
    Se basa en los canales actuales de Asterisk.
    """
    try:
        pn = request.args.get('perfil_nombre', '')
        interno = str(request.args.get('interno', '') or '').strip()

        pf = buscar_perfil(pn)

        if not pf:
            return jsonify({
                'status': 'error',
                'message': 'Perfil no encontrado'
            }), 404

        if not interno:
            return jsonify({
                'status': 'error',
                'message': 'Interno no informado'
            }), 400

        # Obtener canales activos exactamente en este momento.
        try:
            out = _ejecutar_comando_monitor_cdr(
                pf,
                "asterisk -rx 'core show channels concise' 2>/dev/null"
            )
        except Exception as e:
            logger.error("Error obteniendo canales para ruta: {}".format(e))
            return jsonify({
                'status': 'error',
                'message': 'No se pudieron consultar los canales de Asterisk'
            }), 500

        if isinstance(out, tuple):
            salida = out[0]
        else:
            salida = out

        salida = salida or ''

        canales = []

        for linea in salida.splitlines():
            linea = linea.strip()

            if not linea or '!' not in linea:
                continue

            partes = linea.split('!')

            # Formato esperado de "core show channels concise"
            if len(partes) < 8:
                continue

            canal = partes[0].strip()
            contexto = partes[1].strip()
            extension = partes[2].strip()
            estado = partes[4].strip()
            aplicacion = partes[5].strip()
            datos = partes[6].strip()
            callerid = partes[7].strip()

            origen = callerid

            if '<' in callerid and '>' in callerid:
                try:
                    origen = callerid.split('<', 1)[1].split('>', 1)[0].strip()
                except Exception:
                    pass

            canales.append({
                'canal': canal,
                'contexto': contexto,
                'extension': extension,
                'estado': estado,
                'aplicacion': aplicacion,
                'datos': datos,
                'callerid': callerid,
                'origen': origen
            })

        # --------------------------------------------------------
        # Buscar canales donde aparezca el interno.
        # --------------------------------------------------------

        relacionados = []

        for c in canales:
            texto = ' '.join([
                c.get('canal', ''),
                c.get('extension', ''),
                c.get('datos', ''),
                c.get('callerid', ''),
                c.get('origen', '')
            ])

            if interno in texto:
                relacionados.append(c)

        if not relacionados:
            return jsonify({
                'status': 'ok',
                'activo': False,
                'interno': interno,
                'message': 'El interno no tiene una llamada activa',
                'ruta': [],
                'canales': []
            })

        # --------------------------------------------------------
        # Construcción de una ruta legible.
        #
        # No inventamos el troncal. Mostramos los canales reales
        # encontrados y los datos de Dial/Aplicación.
        # --------------------------------------------------------

        pasos = []

        vistos = set()

        for c in relacionados:
            canal = c.get('canal') or '-'
            contexto = c.get('contexto') or '-'
            extension = c.get('extension') or '-'
            aplicacion = c.get('aplicacion') or '-'
            datos = c.get('datos') or '-'
            origen = c.get('origen') or '-'

            clave = (
                canal,
                contexto,
                extension,
                aplicacion,
                datos
            )

            if clave in vistos:
                continue

            vistos.add(clave)

            pasos.append({
                'canal': canal,
                'origen': origen,
                'destino': extension,
                'contexto': contexto,
                'aplicacion': aplicacion,
                'datos': datos,
                'estado': c.get('estado') or '-'
            })

        # --------------------------------------------------------
        # Determinar dirección aproximada a partir de los canales.
        # --------------------------------------------------------

        direccion = 'ACTIVA'

        for c in relacionados:
            canal = (c.get('canal') or '').upper()
            if 'SIP/' + interno in canal:
                direccion = 'LLAMADA DEL INTERNO'
                break

        # Si hay un canal SIP del interno y además otro canal,
        # probablemente existe una salida hacia un troncal.
        if len(relacionados) > 1:
            for c in relacionados:
                canal = c.get('canal') or ''
                if not canal.startswith('SIP/{}'.format(interno)):
                    direccion = 'ENTRANTE / SALIENTE'
                    break

        return jsonify({
            'status': 'ok',
            'activo': True,
            'interno': interno,
            'direccion': direccion,
            'ruta': pasos,
            'canales': relacionados
        })

    except Exception as e:
        logger.exception("Error obteniendo ruta de llamada")
        return jsonify({
            'status': 'error',
            'message': str(e)
        }), 500

@main_bp.route('/api/backups', methods=['GET'])
def api_listar_backups():
    try:
        perfil_nombre = request.args.get('perfil_nombre', '')
        backups = listar_backups(perfil_nombre)
        return jsonify({'status': 'ok', 'backups': backups})
    except Exception as e:
        logger.error("Error listando backups: {}".format(e))
        return jsonify({'status': 'error', 'message': str(e)}), 500

@main_bp.route('/api/backup/crear', methods=['POST'])
def api_crear_backup():
    try:
        d = request.json or {}
        tipo = d.get('tipo', 'provisioner')
        perfil_nombre = d.get('perfil_nombre', '')
        destino = d.get('destino', 'local')
        
        try:
            ok, result, size = crear_backup(tipo=tipo, perfil_nombre=perfil_nombre, destino=destino)
        except TypeError as te:
            if 'unexpected keyword argument' in str(te):
                try:
                    ok, result, size = crear_backup(perfil_nombre=perfil_nombre, destino=destino)
                except TypeError:
                    ok, result, size = crear_backup(perfil_nombre)
            else:
                raise te

        if ok:
            return jsonify({
                'status': 'ok',
                'message': 'Backup creado exitosamente: {}'.format(result),
                'nombre': result,
                'tamano': size
            })
        else:
            return jsonify({'status': 'error', 'message': result}), 500
    except Exception as e:
        logger.error("Error creando backup: {}".format(e))
        return jsonify({'status': 'error', 'message': str(e)}), 500

@main_bp.route('/api/backup/descargar/<nombre>', methods=['GET'])
def api_descargar_backup(nombre):
    try:
        from flask import send_from_directory
        perfil = request.args.get('perfil', '')
        if perfil and perfil != 'Provisioner':
            backup_dir = os.path.join('/opt/voip-provisioner/backups', perfil)
        else:
            backup_dir = None
            if os.path.exists('/opt/voip-provisioner/backups'):
                for d in os.listdir('/opt/voip-provisioner/backups'):
                    candidate = os.path.join('/opt/voip-provisioner/backups', d, nombre)
                    if os.path.exists(candidate):
                        backup_dir = os.path.join('/opt/voip-provisioner/backups', d)
                        break
            if not backup_dir:
                return jsonify({'status': 'error', 'message': 'Backup no encontrado'}), 404
        nombre_seguro = os.path.basename(nombre)
        filepath = os.path.join(backup_dir, nombre_seguro)
        if not os.path.exists(filepath):
            return jsonify({'status': 'error', 'message': 'Backup no encontrado'}), 404
        return send_from_directory(backup_dir, nombre_seguro, as_attachment=True)
    except Exception as e:
        logger.error("Error descargando backup: {}".format(e))
        return jsonify({'status': 'error', 'message': str(e)}), 500

@main_bp.route('/api/backup/eliminar/<nombre>', methods=['POST'])
def api_eliminar_backup(nombre):
    try:
        d = request.json or {}
        cfg = cargar_config()
        if d.get('password') != cfg['seguridad']['admin_password']:
            return jsonify({'status': 'error', 'message': 'Contrasena incorrecta'}), 401
        perfil_nombre = d.get('perfil_nombre', '')
        ok, msg = eliminar_backup(nombre, perfil_nombre)
        if ok:
            return jsonify({'status': 'ok', 'message': msg})
        else:
            return jsonify({'status': 'error', 'message': msg}), 400
    except Exception as e:
        logger.error("Error eliminando backup: {}".format(e))
        return jsonify({'status': 'error', 'message': str(e)}), 500

@main_bp.route('/api/backup/restaurar', methods=['POST'])
def api_restaurar_backup():
    try:
        if 'backup' not in request.files:
            return jsonify({'status': 'error', 'message': 'No se recibio archivo'}), 400
        password = request.form.get('password', '')
        cfg = cargar_config()
        if password != cfg['seguridad']['admin_password']:
            return jsonify({'status': 'error', 'message': 'Contrasena incorrecta'}), 401
        file = request.files['backup']
        if file.filename == '':
            return jsonify({'status': 'error', 'message': 'Archivo vacio'}), 400
        import tempfile
        fd, temp_path = tempfile.mkstemp(suffix='.tar.gz')
        os.close(fd)
        file.save(temp_path)
        try:
            ok, msg = restaurar_backup(temp_path, reiniciar_servicio=True)
            if ok:
                return jsonify({'status': 'ok', 'message': msg + '. El servicio se reiniciara en unos segundos.'})
            else:
                return jsonify({'status': 'error', 'message': msg}), 400
        finally:
            try:
                os.remove(temp_path)
            except:
                pass
    except Exception as e:
        logger.error("Error restaurando backup: {}".format(e))
        return jsonify({'status': 'error', 'message': str(e)}), 500

@main_bp.route('/api/backup/programacion', methods=['GET'])
def api_get_programacion():
    try:
        schedule = cargar_programacion()
        return jsonify({'status': 'ok', 'schedule': schedule})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

@main_bp.route('/api/backup/programacion', methods=['POST'])
def api_set_programacion():
    try:
        d = request.json or {}
        cfg = cargar_config()
        if d.get('password') != cfg['seguridad']['admin_password']:
            return jsonify({'status': 'error', 'message': 'Contrasena incorrecta'}), 401
        schedule = {
            'tipo': d.get('tipo', 'provisioner'),
            'perfil_nombre': d.get('perfil_nombre', ''),
            'destino': d.get('destino', 'local'),
            'activo': d.get('activo', False),
            'frecuencia': d.get('frecuencia', 'diario'),
            'hora': d.get('hora', '02:00'),
            'retencion_dias': int(d.get('retencion_dias', 30))
        }
        if schedule['tipo'] == 'issabel' and not schedule['perfil_nombre']:
            return jsonify({'status': 'error', 'message': 'Selecciona un perfil Issabel'}), 400
        if schedule['frecuencia'] not in ['diario', 'semanal', 'mensual']:
            return jsonify({'status': 'error', 'message': 'Frecuencia invalida'}), 400
        try:
            datetime.strptime(schedule['hora'], '%H:%M')
        except:
            return jsonify({'status': 'error', 'message': 'Hora invalida (formato HH:MM)'}), 400
        if schedule['activo']:
            from app.backup import calcular_proxima_ejecucion
            proxima = calcular_proxima_ejecucion(schedule['frecuencia'], schedule['hora'])
            if proxima:
                schedule['proximo_ejecutado'] = proxima.isoformat()
        if guardar_programacion(schedule):
            tipo_txt = 'Issabel ({})'.format(schedule['perfil_nombre']) if schedule['tipo'] == 'issabel' else 'Provisioner'
            destino_txt = 'en servidor remoto' if schedule['destino'] == 'remoto' else 'en este servidor'
            msg = 'Programacion {}: {} - {} a las {}, guardado {}'.format(
                'activada' if schedule['activo'] else 'desactivada',
                schedule['frecuencia'],
                tipo_txt,
                schedule['hora'],
                destino_txt
            )
            return jsonify({'status': 'ok', 'message': msg, 'schedule': schedule})
        else:
            return jsonify({'status': 'error', 'message': 'Error guardando programacion'}), 500
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

@main_bp.route('/api/backup/programacion/estado', methods=['GET'])
def api_estado_programacion():
    try:
        from app.backup import obtener_estado_scheduler
        return jsonify({'status': 'ok', 'estado': obtener_estado_scheduler()})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

@main_bp.route('/api/backup/ejecutar-ahora', methods=['POST'])
def api_ejecutar_backup_ahora():
    try:
        from app.backup import ejecutar_backup_programado
        ok, msg = ejecutar_backup_programado()
        if ok:
            return jsonify({'status': 'ok', 'message': 'Backup ejecutado: {}'.format(msg)})
        else:
            return jsonify({'status': 'error', 'message': msg}), 500
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

SSH_IDLE_TIMEOUT = 600
_ssh_channels = {}
_ssh_lock = threading.Lock()
_explorer_sessions = {}
_explorer_lock = threading.Lock()


def _vp_close_terminal_session(session):
    if not session:
        return
    for key in ('channel', 'ssh'):
        try:
            value = session.get(key)
            if value:
                value.close()
        except Exception:
            pass


def _vp_terminal_session(session_id, touch=False):
    now = time.time()
    expired_session = None
    with _ssh_lock:
        session = _ssh_channels.get(session_id)
        if session and now - float(session.get('last_activity', now)) >= SSH_IDLE_TIMEOUT:
            expired_session = _ssh_channels.pop(session_id, None)
            session = None
        elif session and touch:
            session['last_activity'] = now
    if expired_session:
        _vp_close_terminal_session(expired_session)
    return session


def _vp_remove_terminal_session(session_id):
    with _ssh_lock:
        session = _ssh_channels.pop(session_id, None)
    if session:
        _vp_close_terminal_session(session)
    return bool(session)


def _vp_ssh_session_cleanup_loop():
    while True:
        try:
            now = time.time()
            expired = []
            with _ssh_lock:
                for session_id, session in list(_ssh_channels.items()):
                    if now - float(session.get('last_activity', now)) >= SSH_IDLE_TIMEOUT:
                        expired.append(_ssh_channels.pop(session_id, None))
            for session in expired:
                _vp_close_terminal_session(session)
            with _explorer_lock:
                for session_id, session in list(_explorer_sessions.items()):
                    if now - float(session.get('last_activity', now)) >= SSH_IDLE_TIMEOUT:
                        _explorer_sessions.pop(session_id, None)
        except Exception:
            logger.exception("Error limpiando sesiones SSH inactivas")
        time.sleep(15)


try:
    _vp_ssh_cleanup_thread = threading.Thread(
        target=_vp_ssh_session_cleanup_loop,
        name='voip-ssh-session-cleanup',
        daemon=True
    )
    _vp_ssh_cleanup_thread.start()
except Exception:
    logger.exception("No se pudo iniciar el limpiador de sesiones SSH")


@main_bp.route('/admin/terminal')
def admin_terminal():
    response = current_app.make_response(render_template('terminal.html'))
    response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
    response.headers['Pragma'] = 'no-cache'
    return response


@main_bp.route('/ws/terminal/connect', methods=['POST'])
def ws_terminal_connect():
    data = request.get_json(silent=True) or {}
    perfil_nombre = str(data.get('perfil_nombre') or '').strip()
    username = str(data.get('username_ssh') or data.get('usuario_ssh') or '').strip()
    password_ssh = str(data.get('password_ssh') or '')
    if not perfil_nombre:
        return jsonify({'status': 'error', 'message': 'Perfil requerido'}), 400
    if not username or not password_ssh:
        return jsonify({'status': 'error', 'message': 'Ingrese nuevamente el usuario y la contraseña SSH.'}), 400
    perfil = buscar_perfil(perfil_nombre, cargar_perfiles())
    if not perfil:
        return jsonify({'status': 'error', 'message': 'Perfil no encontrado'}), 404
    host = str(perfil.get('host') or '').strip()
    try:
        port = int(perfil.get('puerto', 22))
    except Exception:
        return jsonify({'status': 'error', 'message': 'Puerto SSH no válido.'}), 400
    if not host or port < 1 or port > 65535:
        return jsonify({'status': 'error', 'message': 'Servidor SSH no válido.'}), 400
    ssh = None
    channel = None
    try:
        ssh = paramiko.SSHClient()
        ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        ssh.connect(hostname=host, port=port, username=username, password=password_ssh,
                    timeout=8, banner_timeout=8, auth_timeout=8,
                    allow_agent=False, look_for_keys=False)
        channel = ssh.invoke_shell(term='xterm-256color', width=120, height=40)
        channel.setblocking(0)
        session_id = secrets.token_urlsafe(32)
        now = time.time()
        with _ssh_lock:
            _ssh_channels[session_id] = {
                'ssh': ssh, 'channel': channel, 'perfil': perfil_nombre,
                'username': username, 'created_at': now, 'last_activity': now
            }
        logger.info("TERMINAL: Sesion SSH creada para perfil: {}".format(perfil_nombre))
        return jsonify({'status': 'ok', 'session_id': session_id, 'host': host, 'usuario': username})
    except paramiko.AuthenticationException:
        _vp_close_terminal_session({'ssh': ssh, 'channel': channel})
        return jsonify({'status': 'error', 'message': 'Usuario o contraseña SSH incorrectos.'}), 401
    except Exception:
        _vp_close_terminal_session({'ssh': ssh, 'channel': channel})
        logger.exception("TERMINAL: Error conectando al perfil %s", perfil_nombre)
        return jsonify({'status': 'error', 'message': 'No se pudo establecer la conexión SSH.'}), 502


@main_bp.route('/ws/terminal/stream/<session_id>')
def ws_terminal_stream(session_id):
    def generate():
        session = _vp_terminal_session(session_id)
        if not session:
            yield "data: ERROR:Sesion no encontrada o ya cerrada\n\n"
            return
        channel = session['channel']
        yield "data: CONNECTED\n\n"
        ultimo_keepalive = time.time()
        while True:
            try:
                if not _vp_terminal_session(session_id):
                    yield "data: EXPIRED\n\n"
                    return
                readable, _, _ = select.select([channel], [], [], 0.2)
                if readable:
                    data = channel.recv(4096)
                    if data:
                        encoded = base64.b64encode(data).decode('ascii')
                        yield "data: OUTPUT:{}\n\n".format(encoded)
                    else:
                        yield "data: DISCONNECTED\n\n"
                        _vp_remove_terminal_session(session_id)
                        return
                elif time.time() - ultimo_keepalive > 15:
                    yield ": keepalive\n\n"
                    ultimo_keepalive = time.time()
            except Exception:
                yield "data: DISCONNECTED\n\n"
                _vp_remove_terminal_session(session_id)
                return
    response = Response(generate(), mimetype="text/event-stream")
    response.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
    response.headers['Pragma'] = 'no-cache'
    response.headers['X-Accel-Buffering'] = 'no'
    return response


@main_bp.route('/ws/terminal/input/<session_id>', methods=['POST'])
def ws_terminal_input(session_id):
    session = _vp_terminal_session(session_id, touch=True)
    if not session:
        return jsonify({'status': 'error', 'message': 'Sesión SSH vencida o cerrada. Vuelva a autenticarse.'}), 401
    try:
        data = request.get_data(as_text=True)
        if not data or len(data) > 8192:
            return jsonify({'status': 'error', 'message': 'Entrada de terminal no válida.'}), 400
        session['channel'].send(data)
        return jsonify({'status': 'ok'})
    except Exception:
        _vp_remove_terminal_session(session_id)
        logger.exception("Error enviando datos a la terminal SSH")
        return jsonify({'status': 'error', 'message': 'La sesión SSH se cerró.'}), 410


@main_bp.route('/ws/terminal/disconnect/<session_id>', methods=['POST'])
def ws_terminal_disconnect(session_id):
    _vp_remove_terminal_session(session_id)
    return jsonify({'status': 'ok'})


@main_bp.route('/ws/terminal/resize/<session_id>', methods=['POST'])
def ws_terminal_resize(session_id):
    session = _vp_terminal_session(session_id, touch=True)
    if not session:
        return jsonify({'status': 'error', 'message': 'Sesión SSH vencida o cerrada. Vuelva a autenticarse.'}), 401
    try:
        data = request.get_json(silent=True) or {}
        cols = max(20, min(int(data.get('cols', 120)), 400))
        rows = max(5, min(int(data.get('rows', 40)), 150))
        session['channel'].resize_pty(width=cols, height=rows)
        return jsonify({'status': 'ok'})
    except Exception:
        return jsonify({'status': 'error', 'message': 'No se pudo cambiar el tamaño de la terminal.'}), 400


# ===== SSH FILE EXPLORER / SFTP =====

def _explorer_auth_profile(data):
    data = data or {}
    profile_name = str(data.get('perfil') or data.get('nombre') or '').strip()
    username = str(data.get('username_ssh') or data.get('usuario_ssh') or '').strip()
    password = str(data.get('password_ssh') or data.get('password') or '')
    if not profile_name:
        return None, None, None, "Debe seleccionar un servidor."
    if not username or not password:
        return None, None, None, "Ingrese nuevamente el usuario y la contraseña SSH."
    profile = buscar_perfil(profile_name, cargar_perfiles())
    if not profile:
        return None, None, None, "No se encontró el perfil seleccionado."
    host = str(profile.get('host') or '').strip()
    try:
        port = int(profile.get('puerto', 22))
    except Exception:
        return None, None, None, "El puerto SSH del perfil no es válido."
    if not host or port < 1 or port > 65535:
        return None, None, None, "El servidor SSH del perfil no es válido."
    return dict(profile), username, password, None


@main_bp.route('/api/terminal/explorer/auth', methods=['POST'])
def terminal_explorer_auth():
    data = request.get_json(silent=True) or {}
    profile, username, password, error = _explorer_auth_profile(data)
    if error:
        return jsonify({'status': 'error', 'message': error}), 400
    host = str(profile.get('host') or '').strip()
    port = int(profile.get('puerto', 22))
    sftp = None
    transport = None
    try:
        sftp, transport = conectar_sftp(host, port, username, password, '/',
                                        profile.get('ssh_compat', False))
        if not sftp or not transport:
            return jsonify({'status': 'error', 'message': 'No se pudo autenticar. Verifique usuario, contraseña y servidor SSH.'}), 401
        session_id = secrets.token_urlsafe(32)
        now = time.time()
        with _explorer_lock:
            _explorer_sessions[session_id] = {
                'perfil': str(profile.get('nombre') or data.get('perfil') or '').strip(),
                'host': host, 'port': port, 'username': username,
                'password': password, 'created_at': now, 'last_activity': now
            }
        response = jsonify({'status': 'ok', 'session_id': session_id, 'host': host,
                            'usuario': username, 'expires_in': SSH_IDLE_TIMEOUT})
        response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
        response.headers['Pragma'] = 'no-cache'
        return response
    except Exception:
        logger.exception("Error autenticando sesión del explorador SSH")
        return jsonify({'status': 'error', 'message': 'No se pudo autenticar contra el servidor SSH.'}), 502
    finally:
        try:
            if sftp: sftp.close()
        except Exception:
            pass
        try:
            if transport: transport.close()
        except Exception:
            pass


def _explorer_sftp(data):
    data = data or {}
    session_id = str(data.get('session_id') or '').strip()
    requested_profile = str(data.get('perfil') or data.get('nombre') or '').strip()
    if not session_id:
        return None, None, None, "Sesión SFTP vencida o no válida. Ingrese nuevamente usuario y contraseña."
    now = time.time()
    with _explorer_lock:
        auth = _explorer_sessions.get(session_id)
        if auth and now - float(auth.get('last_activity', now)) >= SSH_IDLE_TIMEOUT:
            _explorer_sessions.pop(session_id, None)
            auth = None
        elif auth:
            if requested_profile and requested_profile != auth.get('perfil'):
                return None, None, None, "La sesión SFTP no corresponde al servidor seleccionado."
            auth['last_activity'] = now
            auth = dict(auth)
    if not auth:
        return None, None, None, "Sesión SFTP vencida o no válida. Ingrese nuevamente usuario y contraseña."
    profile = buscar_perfil(auth.get('perfil'), cargar_perfiles())
    if not profile:
        with _explorer_lock:
            _explorer_sessions.pop(session_id, None)
        return None, None, None, "El perfil SSH cambió o ya no existe. Autentíquese nuevamente."
    profile = dict(profile)
    host = str(profile.get('host') or '').strip()
    try:
        port = int(profile.get('puerto', 22))
    except Exception:
        port = -1
    if host != auth.get('host') or port != auth.get('port'):
        with _explorer_lock:
            _explorer_sessions.pop(session_id, None)
        return None, None, None, "El servidor del perfil cambió. Autentíquese nuevamente."
    profile['usuario'] = auth.get('username')
    sftp, transport = conectar_sftp(host, port, auth.get('username'), auth.get('password'),
                                    '/', profile.get('ssh_compat', False))
    if not sftp or not transport:
        return None, None, None, "No se pudo abrir SFTP. Autentíquese nuevamente."
    return sftp, transport, profile, None


@main_bp.route('/api/terminal/explorer/disconnect', methods=['POST'])
def terminal_explorer_disconnect():
    data = request.get_json(silent=True) or {}
    session_id = str(data.get('session_id') or '').strip()
    if session_id:
        with _explorer_lock:
            _explorer_sessions.pop(session_id, None)
    return jsonify({'status': 'ok'})


@main_bp.route('/api/terminal/explorer/list', methods=['POST'])

def terminal_explorer_list():
    try:
        data = request.get_json(silent=True) or {}

        path = str(
            data.get('path')
            or '/'
        ).strip()

        if not path.startswith('/'):
            path = '/' + path

        sftp, transport, perfil, error = _explorer_sftp(data)

        if error:
            return jsonify({
                'status': 'error',
                'message': error
            }), 401

        try:
            entries = []

            for item in sftp.listdir_attr(path):

                name = str(item.filename)

                if name in ('.', '..'):
                    continue

                full = path.rstrip('/') + '/' + name

                mode = int(item.st_mode or 0)

                # S_ISDIR sin importar el tipo de archivo.
                is_dir = bool(mode & 0o040000)

                if is_dir:
                    tipo = 'directory'
                elif mode & 0o120000:
                    tipo = 'symlink'
                else:
                    tipo = 'file'

                try:
                    permisos = oct(mode & 0o777)
                except Exception:
                    permisos = ''

                entries.append({
                    'name': name,
                    'path': full,
                    'type': tipo,
                    'size': int(item.st_size or 0),
                    'mtime': int(item.st_mtime or 0),
                    'permissions': permisos
                })

            entries.sort(
                key=lambda x: (
                    0 if x['type'] == 'directory' else 1,
                    x['name'].lower()
                )
            )

            return jsonify({
                'status': 'ok',
                'path': path,
                'host': perfil.get('host', ''),
                'usuario': perfil.get('usuario', ''),
                'items': entries
            })

        finally:
            try:
                sftp.close()
            except Exception:
                pass

            try:
                transport.close()
            except Exception:
                pass

    except Exception as e:
        logger.exception("Error explorador SSH")
        return jsonify({
            'status': 'error',
            'message': str(e)
        }), 500


@main_bp.route('/api/terminal/explorer/mkdir', methods=['POST'])
def terminal_explorer_mkdir():
    try:
        data = request.get_json(silent=True) or {}

        path = str(data.get('path') or '/').strip()
        name = str(data.get('name') or '').strip()

        if not name:
            return jsonify({
                'status': 'error',
                'message': 'Debe indicar el nombre de la carpeta.'
            }), 400

        if '/' in name or name in ('.', '..'):
            return jsonify({
                'status': 'error',
                'message': 'Nombre de carpeta no válido.'
            }), 400

        if not path.startswith('/'):
            path = '/' + path

        sftp, transport, perfil, error = _explorer_sftp(data)

        if error:
            return jsonify({
                'status': 'error',
                'message': error
            }), 401

        try:
            destino = path.rstrip('/') + '/' + name
            sftp.mkdir(destino)

            return jsonify({
                'status': 'ok',
                'message': 'Carpeta creada correctamente.',
                'path': destino
            })

        finally:
            try:
                sftp.close()
            except Exception:
                pass

            try:
                transport.close()
            except Exception:
                pass

    except Exception as e:
        logger.exception("Error creando carpeta SSH")
        return jsonify({
            'status': 'error',
            'message': str(e)
        }), 500


@main_bp.route('/api/terminal/explorer/delete', methods=['POST'])
def terminal_explorer_delete():
    try:
        data = request.get_json(silent=True) or {}

        path = str(data.get('path') or '').strip()

        if not path or path == '/':
            return jsonify({
                'status': 'error',
                'message': 'No se puede eliminar el directorio raíz.'
            }), 400

        sftp, transport, perfil, error = _explorer_sftp(data)

        if error:
            return jsonify({
                'status': 'error',
                'message': error
            }), 401

        try:
            try:
                attrs = sftp.stat(path)
                mode = int(attrs.st_mode or 0)
                is_dir = bool(mode & 0o040000)
            except Exception:
                is_dir = False

            if is_dir:
                # Primero intentamos borrar directamente.
                # Si contiene archivos, SFTP devolverá error y se informa.
                sftp.rmdir(path)
            else:
                sftp.remove(path)

            return jsonify({
                'status': 'ok',
                'message': 'Elemento eliminado correctamente.'
            })

        finally:
            try:
                sftp.close()
            except Exception:
                pass

            try:
                transport.close()
            except Exception:
                pass

    except Exception as e:
        logger.exception("Error eliminando elemento SSH")

        mensaje = str(e)

        if 'not empty' in mensaje.lower():
            mensaje = (
                'La carpeta no está vacía. '
                'Elimine primero su contenido.'
            )

        return jsonify({
            'status': 'error',
            'message': mensaje
        }), 500


@main_bp.route('/api/terminal/explorer/rename', methods=['POST'])
def terminal_explorer_rename():
    try:
        data = request.get_json(silent=True) or {}

        old = str(data.get('old') or '').strip()
        new_name = str(data.get('new_name') or '').strip()

        if not old or old == '/':
            return jsonify({
                'status': 'error',
                'message': 'Elemento no válido.'
            }), 400

        if not new_name or '/' in new_name:
            return jsonify({
                'status': 'error',
                'message': 'Nuevo nombre no válido.'
            }), 400

        parent = old.rsplit('/', 1)[0]

        if not parent:
            parent = '/'

        new_path = parent.rstrip('/') + '/' + new_name

        sftp, transport, perfil, error = _explorer_sftp(data)

        if error:
            return jsonify({
                'status': 'error',
                'message': error
            }), 401

        try:
            sftp.rename(old, new_path)

            return jsonify({
                'status': 'ok',
                'message': 'Elemento renombrado correctamente.',
                'path': new_path
            })

        finally:
            try:
                sftp.close()
            except Exception:
                pass

            try:
                transport.close()
            except Exception:
                pass

    except Exception as e:
        logger.exception("Error renombrando elemento SSH")
        return jsonify({
            'status': 'error',
            'message': str(e)
        }), 500


@main_bp.route('/api/terminal/explorer/read', methods=['POST'])
def terminal_explorer_read():
    try:
        data = request.get_json(silent=True) or {}

        path = str(data.get('path') or '').strip()

        if not path or path == '/':
            return jsonify({
                'status': 'error',
                'message': 'Debe seleccionar un archivo.'
            }), 400

        sftp, transport, perfil, error = _explorer_sftp(data)

        if error:
            return jsonify({
                'status': 'error',
                'message': error
            }), 401

        try:
            with sftp.open(path, 'rb') as f:
                raw = f.read(2 * 1024 * 1024 + 1)

            if len(raw) > 2 * 1024 * 1024:
                return jsonify({
                    'status': 'error',
                    'message': (
                        'El archivo supera el límite de edición '
                        'de 2 MB.'
                    )
                }), 400

            text = raw.decode('utf-8', 'replace')

            return jsonify({
                'status': 'ok',
                'path': path,
                'content': text
            })

        finally:
            try:
                sftp.close()
            except Exception:
                pass

            try:
                transport.close()
            except Exception:
                pass

    except Exception as e:
        logger.exception("Error leyendo archivo SSH")
        return jsonify({
            'status': 'error',
            'message': str(e)
        }), 500


@main_bp.route('/api/terminal/explorer/write', methods=['POST'])
def terminal_explorer_write():
    try:
        data = request.get_json(silent=True) or {}

        path = str(data.get('path') or '').strip()
        content = data.get('content')

        if not path or path == '/':
            return jsonify({
                'status': 'error',
                'message': 'Ruta de archivo no válida.'
            }), 400

        if content is None:
            content = ''

        content = str(content)

        if len(content.encode('utf-8')) > 2 * 1024 * 1024:
            return jsonify({
                'status': 'error',
                'message': 'El archivo supera el límite de 2 MB.'
            }), 400

        sftp, transport, perfil, error = _explorer_sftp(data)

        if error:
            return jsonify({
                'status': 'error',
                'message': error
            }), 401

        try:
            with sftp.open(path, 'w') as f:
                f.write(content)

            return jsonify({
                'status': 'ok',
                'message': 'Archivo guardado correctamente.'
            })

        finally:
            try:
                sftp.close()
            except Exception:
                pass

            try:
                transport.close()
            except Exception:
                pass

    except Exception as e:
        logger.exception("Error escribiendo archivo SSH")
        return jsonify({
            'status': 'error',
            'message': str(e)
        }), 500


@main_bp.route('/api/terminal/explorer/upload', methods=['POST'])
def terminal_explorer_upload():
    try:
        perfil_nombre = str(
            request.form.get('perfil') or ''
        ).strip()

        password = str(
            request.form.get('password') or ''
        )

        path = str(
            request.form.get('path') or '/'
        ).strip()

        archivo = request.files.get('file')

        if not perfil_nombre:
            return jsonify({
                'status': 'error',
                'message': 'Debe seleccionar un servidor.'
            }), 400

        if not archivo or not archivo.filename:
            return jsonify({
                'status': 'error',
                'message': 'No se recibió ningún archivo.'
            }), 400

        if not path.startswith('/'):
            path = '/' + path

        data = {
            'perfil': perfil_nombre,
            'password': password
        }

        sftp, transport, perfil, error = _explorer_sftp(data)

        if error:
            return jsonify({
                'status': 'error',
                'message': error
            }), 401

        try:
            nombre = os.path.basename(archivo.filename)

            if not nombre:
                return jsonify({
                    'status': 'error',
                    'message': 'Nombre de archivo no válido.'
                }), 400

            destino = path.rstrip('/') + '/' + nombre

            # Guardamos temporalmente solo en memoria/disk temporal
            # y lo subimos mediante SFTP.
            import tempfile

            tmp = tempfile.NamedTemporaryFile(
                delete=False,
                prefix='voip-explorer-'
            )

            tmp_path = tmp.name

            try:
                archivo.save(tmp_path)
                sftp.put(tmp_path, destino)
            finally:
                try:
                    os.unlink(tmp_path)
                except Exception:
                    pass

            return jsonify({
                'status': 'ok',
                'message': 'Archivo subido correctamente.',
                'path': destino
            })

        finally:
            try:
                sftp.close()
            except Exception:
                pass

            try:
                transport.close()
            except Exception:
                pass

    except Exception as e:
        logger.exception("Error subiendo archivo SSH")
        return jsonify({
            'status': 'error',
            'message': str(e)
        }), 500


@main_bp.route('/api/terminal/explorer/download', methods=['POST'])
def terminal_explorer_download():
    try:
        data = request.get_json(silent=True) or {}

        path = str(data.get('path') or '').strip()

        if not path or path == '/':
            return jsonify({
                'status': 'error',
                'message': 'Debe seleccionar un archivo.'
            }), 400

        sftp, transport, perfil, error = _explorer_sftp(data)

        if error:
            return jsonify({
                'status': 'error',
                'message': error
            }), 401

        import tempfile

        tmp_path = None

        try:
            nombre = os.path.basename(path)

            tmp = tempfile.NamedTemporaryFile(
                delete=False,
                prefix='voip-download-'
            )

            tmp_path = tmp.name
            tmp.close()

            sftp.get(path, tmp_path)

            return send_file(
                tmp_path,
                as_attachment=True,
                download_name=nombre
            )

        finally:
            try:
                sftp.close()
            except Exception:
                pass

            try:
                transport.close()
            except Exception:
                pass

            # El archivo temporal será limpiado por el sistema.
            # No se elimina antes de send_file.

    except Exception as e:
        logger.exception("Error descargando archivo SSH")
        return jsonify({
            'status': 'error',
            'message': str(e)
        }), 500


# ===== VOIP PROVISIONER - GESTION DE TRONCALES DESDE MONITOR =====
#
# Estas funciones administran únicamente las troncales creadas desde
# Monitor. No modifican el parser existente de app/monitor.py.
#
# Las troncales se almacenan en un archivo personalizado de Asterisk
# para evitar que retrieve_conf de Issabel las elimine.
#

import shlex
import base64


def _vp_trunk_local(perfil):
    """
    Determina si el perfil apunta al Issabel local.
    Mantiene exactamente el mismo criterio utilizado por el proyecto.
    """
    try:
        host = str(perfil.get('host', '')).strip()
    except Exception:
        host = ''

    return _vp_host_es_local(host)


def _vp_trunk_exec(perfil, comando, timeout=30):
    """
    Ejecuta un comando local o mediante SSH según el perfil.
    """
    try:
        if _vp_trunk_local(perfil):
            p = subprocess.Popen(
                comando,
                shell=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE
            )

            out, err = p.communicate(timeout=timeout)

            return (
                out.decode('utf-8', 'replace'),
                err.decode('utf-8', 'replace'),
                p.returncode
            )

        return ejecutar_comando_ssh(perfil, comando)

    except Exception as e:
        return '', str(e), 1


def _vp_trunk_tipo(tipo):
    tipo = str(tipo or '').strip().upper()

    if tipo in ('SIP', 'IAX2', 'OOH323'):
        return tipo

    return ''


def _vp_trunk_default_port(tipo):
    tipo = _vp_trunk_tipo(tipo)

    if tipo == 'SIP':
        return 5060

    if tipo == 'IAX2':
        return 4569

    if tipo == 'OOH323':
        return 1720

    return 0


def _vp_trunk_safe_name(nombre):
    """
    Nombre de sección Asterisk seguro.
    Permitimos letras, números, guion, punto y underscore.
    """
    nombre = str(nombre or '').strip()

    if not nombre:
        return ''

    if len(nombre) > 64:
        return ''

    if not re.match(r'^[A-Za-z0-9_.-]+$', nombre):
        return ''

    return nombre


def _vp_trunk_description(descripcion):
    """
    Limpia la descripción para que pueda utilizarse como comentario
    dentro del archivo personalizado.
    """
    descripcion = str(descripcion or '').strip()

    if len(descripcion) > 180:
        descripcion = descripcion[:180]

    descripcion = (
        descripcion
        .replace('\r', ' ')
        .replace('\n', ' ')
        .replace('#', '')
    )

    return descripcion.strip()


def _vp_trunk_host(host):
    """
    Host/IP básico.
    No permite espacios ni caracteres que puedan romper el comando.
    """
    host = str(host or '').strip()

    if not host:
        return ''

    if len(host) > 255:
        return ''

    if re.match(r'^[A-Za-z0-9_.:-]+$', host):
        return host

    return ''


def _vp_trunk_custom_paths(tipo):
    """
    Archivo personalizado recomendado por Issabel/Asterisk.
    """
    tipo = _vp_trunk_tipo(tipo)

    if tipo == 'SIP':
        return (
            '/etc/asterisk/sip_custom.conf',
            '/etc/asterisk/sip.conf'
        )

    if tipo == 'IAX2':
        return (
            '/etc/asterisk/iax_custom.conf',
            '/etc/asterisk/iax.conf'
        )

    if tipo == 'OOH323':
        return (
            '/etc/asterisk/ooh323_custom.conf',
            '/etc/asterisk/ooh323.conf'
        )

    return '', ''


def _vp_trunk_config(
    nombre,
    descripcion,
    tipo,
    host,
    port,
    usuario='',
    clave=''
):
    """
    Genera la configuración de la troncal.

    Usuario y clave son opcionales para SIP/IAX2.
    OOH323 conserva exactamente su configuración anterior.
    """

    comentario = descripcion or 'Troncal creada desde VoIP Provisioner'

    usuario = str(usuario or '').strip()
    clave = str(clave or '')

    credenciales_sip = ''

    if usuario:
        credenciales_sip += 'defaultuser={}'.format(usuario)

    if clave:
        if credenciales_sip:
            credenciales_sip += '\n'
        credenciales_sip += 'secret={}'.format(clave)

    if credenciales_sip:
        credenciales_sip += '\n'

    credenciales_iax = ''

    if usuario:
        credenciales_iax += 'username={}'.format(usuario)

    if clave:
        if credenciales_iax:
            credenciales_iax += '\n'
        credenciales_iax += 'secret={}'.format(clave)

    if credenciales_iax:
        credenciales_iax += '\n'

    if tipo == 'SIP':

        return """; ============================================================
; VOIP PROVISIONER - TRONCAL
; Nombre: {nombre}
; Descripcion: {descripcion}
; ============================================================

[{nombre}]
type=peer
host={host}
port={port}
{credenciales}qualify=yes
context=from-trunk
disallow=all
allow=ulaw
allow=alaw

""".format(
            nombre=nombre,
            descripcion=comentario,
            host=host,
            port=port,
            credenciales=credenciales_sip
        )

    if tipo == 'IAX2':

        return """; ============================================================
; VOIP PROVISIONER - TRONCAL
; Nombre: {nombre}
; Descripcion: {descripcion}
; ============================================================

[{nombre}]
type=peer
host={host}
port={port}
{credenciales}qualify=yes
context=from-trunk
disallow=all
allow=ulaw
allow=alaw

""".format(
            nombre=nombre,
            descripcion=comentario,
            host=host,
            port=port,
            credenciales=credenciales_iax
        )

    if tipo == 'OOH323':

        return """; ============================================================
; VOIP PROVISIONER - TRONCAL
; Nombre: {nombre}
; Descripcion: {descripcion}
; ============================================================

[{nombre}]
type=friend
host={host}
port={port}
faststart=no
h245tunneling=yes
dtmfmode=rfc2833
context=from-trunk
disallow=all
allow=ulaw
allow=alaw

""".format(
            nombre=nombre,
            descripcion=comentario,
            host=host,
            port=port
        )

    return ''


def _vp_trunk_ensure_include(perfil, tipo):
    """
    Comprueba si el archivo custom ya está incluido.
    Si no lo está, realiza un backup remoto y agrega un único include.
    """

    custom, base = _vp_trunk_custom_paths(tipo)

    if not custom or not base:
        return False, '', 'Tipo de troncal no válido'

    custom_q = shlex.quote(custom)
    base_q = shlex.quote(base)

    comando = r"""set -e

CUSTOM={custom}
BASE={base}

if [ ! -f "$BASE" ]; then
    echo "No existe $BASE" >&2
    exit 20
fi

mkdir -p /etc/asterisk

if [ ! -f "$CUSTOM" ]; then
    touch "$CUSTOM"
    chmod 644 "$CUSTOM"
fi

if grep -Fqs "$CUSTOM" "$BASE"; then
    echo "INCLUDE_OK"
    exit 0
fi

if grep -Fqs "$(basename "$CUSTOM")" "$BASE"; then
    echo "INCLUDE_OK"
    exit 0
fi

STAMP=$(date +%Y%m%d_%H%M%S)

cp -a "$BASE" "$BASE.voip-provisioner-backup-$STAMP"

printf '\n\n; VOIP PROVISIONER - INCLUDE AUTOMATICO\n#include %s\n' "$CUSTOM" >> "$BASE"

echo "INCLUDE_ADDED"
""".format(
        custom=custom_q,
        base=base_q
    )

    out, err, code = _vp_trunk_exec(perfil, comando)

    if code != 0:
        return False, out, err or 'No se pudo preparar el include de Asterisk'

    return True, out, ''


def _vp_trunk_existing(perfil, tipo, nombre, host):
    """
    Comprueba si ya existe una troncal con el mismo IP/host.

    La comprobación se realiza por IP/host y no por nombre.

    Si el archivo personalizado todavía no existe, se considera
    que no hay duplicado y se permite continuar con la creación.
    """

    custom, base = _vp_trunk_custom_paths(tipo)

    host = str(host or '').strip()

    if not host:
        return False, ''

    archivos = []

    if custom:
        archivos.append(custom)

    if base and base != custom:
        archivos.append(base)

    # Eliminar duplicados manteniendo el orden.
    archivos = list(dict.fromkeys(archivos))

    for archivo in archivos:

        archivo_q = shlex.quote(archivo)

        patron = (
            r'^[[:space:]]*(host|address)[[:space:]]*=[[:space:]]*'
            + re.escape(host)
            + r'([[:space:]]|$)'
        )

        patron_q = shlex.quote(patron)

        # Si el archivo no existe, no es un error:
        # simplemente no hay nada que comprobar.
        #
        # "test -f ... || exit 0" evita que grep intente abrir
        # un archivo inexistente.
        comando = (
            "test -f {archivo} || exit 0; "
            "grep -Eiq {patron} {archivo} "
            "&& echo EXISTE_IP"
        ).format(
            archivo=archivo_q,
            patron=patron_q
        )

        out, err, code = _vp_trunk_exec(
            perfil,
            comando
        )

        # Archivo inexistente o IP no encontrada.
        if 'EXISTE_IP' not in out:
            if code not in (0, 1):
                return False, (
                    err or
                    'No se pudo consultar la configuración '
                    'de la troncal.'
                )
            continue

        return True, (
            'Ya existe una troncal configurada con la IP/host '
            '{} para {}. Utilizá otra IP/host.'.format(
                host,
                tipo.upper()
            )
        )

    return False, ''

def _vp_trunk_write(perfil, tipo, nombre, contenido):
    """
    Agrega una sección al archivo custom sin reemplazar otras troncales.
    """
    custom, base = _vp_trunk_custom_paths(tipo)

    contenido_b64 = base64.b64encode(
        contenido.encode('utf-8')
    ).decode('ascii')

    custom_q = shlex.quote(custom)

    comando = r"""set -e

CUSTOM={custom}

mkdir -p /etc/asterisk

if [ ! -f "$CUSTOM" ]; then
    touch "$CUSTOM"
    chmod 644 "$CUSTOM"
fi

STAMP=$(date +%Y%m%d_%H%M%S)

cp -a "$CUSTOM" "$CUSTOM.voip-provisioner-backup-$STAMP"

printf '%s' '{contenido}' | base64 -d >> "$CUSTOM"

echo TRONCAL_ESCRITA
""".format(
        custom=custom_q,
        contenido=contenido_b64
    )

    return _vp_trunk_exec(perfil, comando)



def _vp_trunk_register_issabel(
    perfil,
    tipo,
    nombre,
    host,
    port
):
    """
    Registra la troncal en la tabla asterisk.trunks de Issabel.

    No ejecuta retrieve_conf.
    No regenera archivos de Asterisk.
    """

    tipo = _vp_trunk_tipo(tipo)

    nombre = str(nombre or '').strip()
    host = str(host or '').strip()

    if not nombre:
        return False, '', 'Nombre de troncal vacío.'

    if not host:
        return False, '', 'Host de troncal vacío.'

    if tipo == 'OOH323':
        tech = 'custom'
        channelid = 'OOH323/{}/$OUTNUM$'.format(nombre)
        usercontext = 'notneeded'

    elif tipo == 'SIP':
        tech = 'sip'
        channelid = nombre
        usercontext = ''

    elif tipo == 'IAX2':
        tech = 'iax'
        channelid = nombre
        usercontext = ''

    else:
        return False, '', 'Tipo de troncal no válido.'

    def _b64(valor):
        return base64.b64encode(
            str(valor or '').encode('utf-8')
        ).decode('ascii')

    nombre_b64 = _b64(nombre)
    tech_b64 = _b64(tech)
    channelid_b64 = _b64(channelid)
    usercontext_b64 = _b64(usercontext)

    comando = """set -e

TMP_CNF=$(mktemp)
trap 'rm -f \"$TMP_CNF\"' EXIT

AMPDBUSER=$(grep '^AMPDBUSER=' /etc/amportal.conf | head -1 | cut -d= -f2-)
AMPDBPASS=$(grep '^AMPDBPASS=' /etc/amportal.conf | head -1 | cut -d= -f2-)
AMPDBHOST=$(grep '^AMPDBHOST=' /etc/amportal.conf | head -1 | cut -d= -f2-)
AMPDBNAME=$(grep '^AMPDBNAME=' /etc/amportal.conf | head -1 | cut -d= -f2-)

if [ -z \"$AMPDBUSER\" ] || [ -z \"$AMPDBPASS\" ] || [ -z \"$AMPDBNAME\" ]; then
    echo \"No se pudieron obtener las credenciales de Issabel desde /etc/amportal.conf\" >&2
    exit 20
fi

cat > \"$TMP_CNF\" <<MYSQLCONF
[client]
user=$AMPDBUSER
password=$AMPDBPASS
host=$AMPDBHOST
database=$AMPDBNAME
MYSQLCONF

NAME=$(printf '%s' '__NAME_B64__' | base64 -d)
TECH=$(printf '%s' '__TECH_B64__' | base64 -d)
CHANNELID=$(printf '%s' '__CHANNELID_B64__' | base64 -d)
USERCONTEXT=$(printf '%s' '__USERCONTEXT_B64__' | base64 -d)

EXISTE=$(mysql --defaults-extra-file=\"$TMP_CNF\" -N -B -e \"SELECT COUNT(*) FROM trunks WHERE name='$NAME';\")

if [ \"$EXISTE\" -gt 0 ]; then
    mysql --defaults-extra-file=\"$TMP_CNF\" -N -B -e \"SELECT CONCAT('TRONCAL_YA_EXISTE|',trunkid,'|',name,'|',tech,'|',channelid) FROM trunks WHERE name='$NAME' LIMIT 1;\"
    exit 0
fi

NEXT_ID=$(mysql --defaults-extra-file=\"$TMP_CNF\" -N -B -e \"SELECT COALESCE(MAX(trunkid),0)+1 FROM trunks;\")

mysql --defaults-extra-file=\"$TMP_CNF\" -e \"INSERT INTO trunks (trunkid,name,tech,outcid,keepcid,maxchans,failscript,dialoutprefix,channelid,usercontext,provider,disabled,\\`continue\\`) VALUES ($NEXT_ID,'$NAME','$TECH','','off','','','','$(printf '%s' \"$CHANNELID\")','$USERCONTEXT','','off','off');\"

mysql --defaults-extra-file=\"$TMP_CNF\" -N -B -e \"SELECT CONCAT('TRONCAL_REGISTRADA|',trunkid,'|',name,'|',tech,'|',channelid) FROM trunks WHERE trunkid=$NEXT_ID LIMIT 1;\"
"""

    comando = comando.replace('__NAME_B64__', shlex.quote(nombre_b64))
    comando = comando.replace('__TECH_B64__', shlex.quote(tech_b64))
    comando = comando.replace('__CHANNELID_B64__', shlex.quote(channelid_b64))
    comando = comando.replace('__USERCONTEXT_B64__', shlex.quote(usercontext_b64))

    out, err, code = _vp_trunk_exec(
        perfil,
        comando,
        timeout=30
    )

    if code != 0:
        return False, out, err or 'No se pudo registrar la troncal en Issabel.'

    if 'TRONCAL_REGISTRADA|' in out:
        return True, out, ''

    if 'TRONCAL_YA_EXISTE|' in out:
        return True, out, ''

    return False, out, 'Issabel no confirmó el registro de la troncal.'


def _vp_trunk_reload(perfil, tipo):
    """
    Recarga únicamente el módulo correspondiente.
    Nunca reinicia Asterisk ni Issabel completo.
    """
    tipo = _vp_trunk_tipo(tipo)

    if tipo == 'SIP':
        comando = (
            "asterisk -rx 'sip reload'"
        )

    elif tipo == 'IAX2':
        comando = (
            "asterisk -rx 'iax2 reload'"
        )

    elif tipo == 'OOH323':
        comando = (
            "if asterisk -rx 'module show like chan_ooh323' "
            "| grep -qi chan_ooh323; then "
            "asterisk -rx 'module reload chan_ooh323.so'; "
            "else "
            "asterisk -rx 'ooh323 reload'; "
            "fi"
        )

    else:
        return '', 'Tipo no válido', 1

    return _vp_trunk_exec(perfil, comando, timeout=30)


def _vp_trunk_verify(perfil, tipo, nombre):
    """
    Comprueba que la nueva troncal aparece en Asterisk.
    """
    tipo = _vp_trunk_tipo(tipo)
    nombre_q = shlex.quote(nombre)

    if tipo == 'SIP':
        comando = "asterisk -rx 'sip show peers'"

    elif tipo == 'IAX2':
        comando = "asterisk -rx 'iax2 show peers'"

    elif tipo == 'OOH323':
        comando = "asterisk -rx 'ooh323 show peers'"

    else:
        return False, '', 'Tipo no válido'

    out, err, code = _vp_trunk_exec(perfil, comando)

    if code != 0:
        return False, out, err or 'No se pudo consultar Asterisk'

    encontrado = False

    for linea in str(out or '').splitlines():
        if nombre.lower() in linea.lower():
            encontrado = True
            break

    return encontrado, out, ''


def _vp_trunk_admin_password(data):
    """
    Usa exactamente la contraseña administrativa existente.
    """
    try:
        data = data or {}

        password = str(
            data.get('password', '')
        )

        cfg = cargar_config()

        admin_password = str(
            cfg.get('seguridad', {})
               .get('admin_password', '')
        )

        return bool(
            password
            and admin_password
            and password == admin_password
        )

    except Exception as e:
        logger.error(
            'Error verificando contraseña de troncales: %s',
            e
        )
        return False


@main_bp.route(
    '/api/monitoreo/troncal/recargar',
    methods=['POST']
)
def api_monitoreo_troncal_recargar():
    """Recarga únicamente el módulo Asterisk de la troncal seleccionada."""
    try:
        data = request.get_json(silent=True) or {}

        if not _vp_trunk_admin_password(data):
            return jsonify({
                'status': 'error',
                'message': 'Contraseña de administrador incorrecta.'
            }), 401

        perfil_nombre = str(data.get('perfil_nombre', '')).strip()
        nombre = str(data.get('nombre', '')).strip()
        tipo_recibido = str(data.get('tipo', '')).strip().upper()

        if tipo_recibido == 'CUSTOM':
            tipo_recibido = 'OOH323'
        elif tipo_recibido == 'IAX':
            tipo_recibido = 'IAX2'

        tipo = _vp_trunk_tipo(tipo_recibido)

        if not perfil_nombre:
            return jsonify({
                'status': 'error',
                'message': 'Falta seleccionar el servidor/perfil de la troncal.'
            }), 400

        if not nombre:
            return jsonify({
                'status': 'error',
                'message': 'Falta el nombre de la troncal.'
            }), 400

        if not tipo:
            return jsonify({
                'status': 'error',
                'message': 'Tipo de troncal no válido para recargar.'
            }), 400

        perfil = buscar_perfil(perfil_nombre, cargar_perfiles())
        if not perfil:
            return jsonify({
                'status': 'error',
                'message': 'Perfil no encontrado.'
            }), 404

        salida, error, codigo = _vp_trunk_reload(perfil, tipo)

        logger.info(
            '[TRONCAL RELOAD] perfil=%s tipo=%s nombre=%s codigo=%s salida=%s error=%s',
            perfil_nombre,
            tipo,
            nombre,
            codigo,
            str(salida or '').strip()[:500],
            str(error or '').strip()[:500]
        )

        if codigo != 0:
            detalle = str(error or salida or '').strip()
            return jsonify({
                'status': 'error',
                'message': detalle or 'No se pudo recargar el módulo de la troncal.'
            }), 500

        return jsonify({
            'status': 'ok',
            'message': 'Se recargó el módulo {} correctamente.'.format(tipo)
        })

    except Exception as exc:
        logger.exception('[TRONCAL RELOAD] Error recargando la troncal')
        return jsonify({
            'status': 'error',
            'message': 'Error al recargar la troncal: {}'.format(str(exc))
        }), 500


@main_bp.route(
    '/api/monitoreo/troncal/agregar',
    methods=['POST']
)
def api_monitoreo_troncal_agregar():

    try:

        data = request.get_json(
            silent=True
        ) or {}

        if not _vp_trunk_admin_password(data):
            return jsonify({
                'status': 'error',
                'message': 'Contraseña de administrador incorrecta'
            }), 401

        perfil_nombre = str(
            data.get('perfil_nombre', '')
        ).strip()

        nombre = _vp_trunk_safe_name(
            data.get('nombre', '')
        )

        descripcion = _vp_trunk_description(
            data.get('descripcion', '')
        )

        tipo = _vp_trunk_tipo(
            data.get('tipo', '')
        )

        host = _vp_trunk_host(
            data.get('host', '')
        )

        usuario = str(
            data.get('usuario', '')
        ).strip()

        clave = str(
            data.get('clave', '')
        )

        # Evitar inyección de nuevas líneas en el archivo
        # de configuración de Asterisk.
        if '\n' in usuario or '\r' in usuario:
            return jsonify({
                'status': 'error',
                'message': 'El usuario contiene caracteres no válidos.'
            }), 400

        if '\n' in clave or '\r' in clave:
            return jsonify({
                'status': 'error',
                'message': 'La contraseña de la troncal contiene caracteres no válidos.'
            }), 400

        if bool(usuario) != bool(clave):
            return jsonify({
                'status': 'error',
                'message': 'Debe completar usuario y contraseña de la troncal.'
            }), 400

        port_raw = str(
            data.get('port', '')
        ).strip()

        if not perfil_nombre:
            return jsonify({
                'status': 'error',
                'message': 'Debe seleccionar un servidor.'
            }), 400

        if not nombre:
            return jsonify({
                'status': 'error',
                'message': (
                    'Nombre no válido. Use letras, números, '
                    'guion, punto o underscore.'
                )
            }), 400

        if not tipo:
            return jsonify({
                'status': 'error',
                'message': 'Debe seleccionar el tipo de troncal.'
            }), 400

        # IAX2 bilateral:
        # la IP/host se obtiene automáticamente del perfil destino.
        # SIP y OOH323 continúan requiriendo host explícito.
        if tipo != 'IAX2' and not host:
            return jsonify({
                'status': 'error',
                'message': 'Debe indicar la IP o host de la troncal.'
            }), 400

        perfil = buscar_perfil(
            perfil_nombre,
            cargar_perfiles()
        )

        if not perfil:
            return jsonify({
                'status': 'error',
                'message': 'Perfil de servidor no encontrado.'
            }), 404

        if port_raw:
            try:
                port = int(port_raw)
            except Exception:
                return jsonify({
                    'status': 'error',
                    'message': 'Puerto no válido.'
                }), 400
        else:
            port = _vp_trunk_default_port(tipo)

        if port < 1 or port > 65535:
            return jsonify({
                'status': 'error',
                'message': 'Puerto fuera de rango.'
            }), 400

        existe, error = _vp_trunk_existing(
            perfil,
            tipo,
            nombre,
            host
        )

        if error:
            return jsonify({
                'status': 'error',
                'message': error
            }), 500

        if existe:
            return jsonify({
                'status': 'error',
                'message': (
                    'Ya existe una troncal con el nombre '
                    '"{}" en este tipo de configuración.'
                ).format(nombre)
            }), 409

                # ====================================================
        # IAX2 BILATERAL
        #
        # El técnico selecciona:
        #   - servidor origen: perfil actual
        #   - servidor destino: perfil_destino
        #
        # Las IP se obtienen de los perfiles.
        # Las credenciales NO se guardan en JSON.
        # Se generan únicamente durante esta operación.
        # ====================================================

        
        if tipo == 'IAX2':
            # ====================================================
            # IAX2 BILATERAL
            # ====================================================
            #
            # Credenciales GENERICAS para ambos extremos:
            #
            #   username = *voippjn
            #   secret   = *voippjn
            #
            # No dependen del nombre del perfil.
            # No se guardan en JSON.
            # ====================================================

            perfil_destino_nombre = str(
                data.get('perfil_destino', '')
            ).strip()

            if not perfil_destino_nombre:
                return jsonify({
                    'status': 'error',
                    'message':
                        'Seleccioná el servidor destino para IAX2.'
                }), 400

            if perfil_destino_nombre == perfil_nombre:
                return jsonify({
                    'status': 'error',
                    'message':
                        'El servidor destino debe ser diferente al origen.'
                }), 400

            perfiles = cargar_perfiles()

            perfil_destino = buscar_perfil(
                perfil_destino_nombre,
                perfiles
            )

            if not perfil_destino:
                return jsonify({
                    'status': 'error',
                    'message':
                        'No se encontró el perfil destino.'
                }), 404

            # ----------------------------------------------------
            # HOSTS
            # ----------------------------------------------------

            host_origen = str(
                perfil.get('host', '')
            ).strip()

            host_destino = str(
                perfil_destino.get('host', '')
            ).strip()

            if not host_origen:
                return jsonify({
                    'status': 'error',
                    'message':
                        'El perfil origen no tiene host configurado.'
                }), 400

            if not host_destino:
                return jsonify({
                    'status': 'error',
                    'message':
                        'El perfil destino no tiene host configurado.'
                }), 400

            # ----------------------------------------------------
            # CREDENCIALES GENERICAS
            # ----------------------------------------------------

            iax_usuario = '*voippjn'
            iax_clave = '*voippjn'

            # ----------------------------------------------------
            # PUERTO FIJO IAX2
            # ----------------------------------------------------

            port = 4569

            # ----------------------------------------------------
            # NOMBRE DE LA SEGUNDA PUNTA
            #
            # En el servidor destino se conserva como nombre
            # la identificación del servidor origen.
            # ----------------------------------------------------

            nombre_destino = str(
                perfil.get('nombre', '')
            ).strip()

            if not nombre_destino:
                return jsonify({
                    'status': 'error',
                    'message':
                        'El perfil origen no tiene nombre válido.'
                }), 400

            # ----------------------------------------------------
            # VALIDAR PRIMERO AMBOS EXTREMOS
            # ----------------------------------------------------

            existe_origen, error_origen = _vp_trunk_existing(
                perfil,
                'IAX2',
                nombre,
                host_destino
            )

            if error_origen:
                return jsonify({
                    'status': 'error',
                    'message':
                        'Error validando origen: ' +
                        str(error_origen)
                }), 500

            if existe_origen:
                return jsonify({
                    'status': 'error',
                    'message':
                        'La troncal "' +
                        nombre +
                        '" ya existe en ' +
                        perfil_nombre +
                        '.'
                }), 409

            existe_destino, error_destino = _vp_trunk_existing(
                perfil_destino,
                'IAX2',
                nombre_destino,
                host_origen
            )

            if error_destino:
                return jsonify({
                    'status': 'error',
                    'message':
                        'Error validando destino: ' +
                        str(error_destino)
                }), 500

            if existe_destino:
                return jsonify({
                    'status': 'error',
                    'message':
                        'La troncal "' +
                        nombre_destino +
                        '" ya existe en ' +
                        perfil_destino_nombre +
                        '.'
                }), 409

            # ----------------------------------------------------
            # CREAR EXTREMO ORIGEN
            #
            # Peer:
            #   host = destino
            #   username = *voippjn
            #   secret = *voippjn
            #
            # User:
            #   account = *voippjn
            #   secret  = *voippjn
            # ----------------------------------------------------

            ok_origen, out_origen, err_origen = (
                _vp_trunk_create_iax_issabel(
                    perfil,
                    nombre,
                    descripcion,
                    host_destino,
                    port,
                    iax_usuario,
                    iax_clave,
                    iax_usuario,
                    iax_clave
                )
            )

            if not ok_origen:
                return jsonify({
                    'status': 'error',
                    'message':
                        'No se pudo crear IAX2 en ' +
                        perfil_nombre +
                        ': ' +
                        (err_origen or out_origen or '')
                }), 500

            # ----------------------------------------------------
            # CREAR EXTREMO DESTINO
            #
            # Peer:
            #   host = origen
            #   username = *voippjn
            #   secret = *voippjn
            #
            # User:
            #   account = *voippjn
            #   secret  = *voippjn
            # ----------------------------------------------------

            ok_destino, out_destino, err_destino = (
                _vp_trunk_create_iax_issabel(
                    perfil_destino,
                    nombre_destino,
                    descripcion,
                    host_origen,
                    port,
                    iax_usuario,
                    iax_clave,
                    iax_usuario,
                    iax_clave
                )
            )

            if not ok_destino:
                # ------------------------------------------------
                # ROLLBACK ORIGEN
                # ------------------------------------------------

                try:
                    _vp_trunk_run_sql(
                        perfil,
                        "DELETE FROM iax "
                        "WHERE id IN ("
                        "SELECT CONCAT('tr-peer-',trunkid) "
                        "FROM trunks "
                        "WHERE name=" +
                        _vp_trunk_sql_q(nombre) +
                        ") "
                        "OR id IN ("
                        "SELECT CONCAT('tr-user-',trunkid) "
                        "FROM trunks "
                        "WHERE name=" +
                        _vp_trunk_sql_q(nombre) +
                        "); "
                        "DELETE FROM trunks "
                        "WHERE name=" +
                        _vp_trunk_sql_q(nombre) +
                        ";"
                    )
                except Exception:
                    pass

                return jsonify({
                    'status': 'error',
                    'message':
                        'No se pudo crear IAX2 en ' +
                        perfil_destino_nombre +
                        '. Se revirtió la creación en ' +
                        perfil_nombre +
                        '. ' +
                        (err_destino or out_destino or '')
                }), 500

            # ----------------------------------------------------
            # RUTAS SALIENTES SEGUN DIAL PLAN
            #
            # Cada servidor utiliza SU PROPIO Dial Plan y SU PROPIA
            # punta IAX2 para salir hacia el otro servidor.
            # ----------------------------------------------------

            # ----------------------------------------------------
            # IMPORTANTE:
            #
            # La ruta de cada servidor debe permitir marcar los
            # internos que pertenecen al servidor DESTINO.
            #
            # Por lo tanto:
            #
            #   A -> B = Dial Plan de B
            #   B -> A = Dial Plan de A
            # ----------------------------------------------------

            dialplan_origen = _vp_perfil_dialplan(
                perfil_destino.get('dialplan', [])
            )

            dialplan_destino = _vp_perfil_dialplan(
                perfil.get('dialplan', [])
            )

            ruta_origen = _vp_iax_crear_ruta_dialplan(
                perfil,
                nombre,
                dialplan_origen
            )

            if not ruta_origen[0]:
                # Si la ruta falla, eliminar ambas troncales para
                # no dejar una configuración bilateral incompleta.
                try:
                    _vp_trunk_run_sql(
                        perfil,
                        "DELETE FROM outbound_route_trunks "
                        "WHERE trunk_id IN ("
                        "SELECT trunkid FROM trunks "
                        "WHERE name=" +
                        _vp_trunk_sql_q(nombre) +
                        "); "
                        "DELETE FROM outbound_route_patterns "
                        "WHERE route_id NOT IN "
                        "(SELECT route_id FROM outbound_route_trunks); "
                        "DELETE FROM iax "
                        "WHERE id IN ("
                        "SELECT CONCAT('tr-peer-',trunkid) "
                        "FROM trunks "
                        "WHERE name=" +
                        _vp_trunk_sql_q(nombre) +
                        ") "
                        "OR id IN ("
                        "SELECT CONCAT('tr-user-',trunkid) "
                        "FROM trunks "
                        "WHERE name=" +
                        _vp_trunk_sql_q(nombre) +
                        "); "
                        "DELETE FROM trunks "
                        "WHERE name=" +
                        _vp_trunk_sql_q(nombre) +
                        ";"
                    )
                except Exception:
                    pass

                try:
                    _vp_trunk_run_sql(
                        perfil_destino,
                        "DELETE FROM outbound_route_trunks "
                        "WHERE trunk_id IN ("
                        "SELECT trunkid FROM trunks "
                        "WHERE name=" +
                        _vp_trunk_sql_q(nombre_destino) +
                        "); "
                        "DELETE FROM iax "
                        "WHERE id IN ("
                        "SELECT CONCAT('tr-peer-',trunkid) "
                        "FROM trunks "
                        "WHERE name=" +
                        _vp_trunk_sql_q(nombre_destino) +
                        ") "
                        "OR id IN ("
                        "SELECT CONCAT('tr-user-',trunkid) "
                        "FROM trunks "
                        "WHERE name=" +
                        _vp_trunk_sql_q(nombre_destino) +
                        "); "
                        "DELETE FROM trunks "
                        "WHERE name=" +
                        _vp_trunk_sql_q(nombre_destino) +
                        ";"
                    )
                except Exception:
                    pass

                return jsonify({
                    'status': 'error',
                    'message':
                        'Se crearon las puntas IAX2 pero falló la ruta '
                        'saliente en ' + perfil_nombre + ': ' +
                        (ruta_origen[2] or ruta_origen[1] or '')
                }), 500

            ruta_destino = _vp_iax_crear_ruta_dialplan(
                perfil_destino,
                nombre_destino,
                dialplan_destino
            )

            if not ruta_destino[0]:
                return jsonify({
                    'status': 'error',
                    'message':
                        'Se creó la ruta en ' +
                        perfil_nombre +
                        ' pero falló la ruta en ' +
                        perfil_destino_nombre +
                        ': ' +
                        (ruta_destino[2] or ruta_destino[1] or '')
                }), 500

            # ----------------------------------------------------
            # APLICAR EN AMBOS ISSABEL
            # ----------------------------------------------------

            apply_origen = _vp_trunk_apply_issabel(
                perfil,
                'IAX2'
            )

            apply_destino = _vp_trunk_apply_issabel(
                perfil_destino,
                'IAX2'
            )

            errores_apply = []

            if len(apply_origen) >= 3 and apply_origen[2] != 0:
                errores_apply.append(
                    'No se pudo aplicar IAX2 en ' +
                    perfil_nombre +
                    ': ' +
                    str(apply_origen[1] or apply_origen[0] or '')
                )

            if len(apply_destino) >= 3 and apply_destino[2] != 0:
                errores_apply.append(
                    'No se pudo aplicar IAX2 en ' +
                    perfil_destino_nombre +
                    ': ' +
                    str(apply_destino[1] or apply_destino[0] or '')
                )

            # ----------------------------------------------------
            # VERIFICAR QUE REALMENTE QUEDARON CARGADAS
            # ----------------------------------------------------

            if not errores_apply:

                ver_origen = _vp_iax_verificar_aplicacion(
                    perfil,
                    nombre
                )

                if not ver_origen[0]:
                    errores_apply.append(
                        'IAX2 no quedó cargado en ' +
                        perfil_nombre +
                        ': ' +
                        ver_origen[1]
                    )

                ver_destino = _vp_iax_verificar_aplicacion(
                    perfil_destino,
                    nombre_destino
                )

                if not ver_destino[0]:
                    errores_apply.append(
                        'IAX2 no quedó cargado en ' +
                        perfil_destino_nombre +
                        ': ' +
                        ver_destino[1]
                    )

            if errores_apply:
                return jsonify({
                    'status': 'warning',
                    'message':
                        'Las dos puntas fueron creadas y las rutas '
                        'fueron generadas, pero: ' +
                        ' '.join(errores_apply)
                }), 200

            return jsonify({
                'status': 'ok',
                'message':
                    'Troncal IAX2 bilateral creada correctamente entre ' +
                    perfil_nombre +
                    ' y ' +
                    perfil_destino_nombre +
                    '. Rutas salientes generadas según el Dial Plan.'
            })

        if tipo in ('SIP', 'OOH323'):
            # SIP y OOH323 usan archivos custom; no regeneramos archivos
            # administrados por Issabel ni reiniciamos Asterisk completo.
            ok_include, salida_include, error_include = _vp_trunk_ensure_include(
                perfil,
                tipo
            )
            if not ok_include:
                return jsonify({
                    'status': 'error',
                    'message': (
                        'No se pudo preparar la inclusión de la configuración '
                        '{}: {}'.format(tipo, error_include or salida_include or '')
                    )
                }), 500

            contenido = _vp_trunk_config(
                nombre,
                descripcion,
                tipo,
                host,
                port,
                usuario,
                clave
            )
            if not contenido:
                return jsonify({
                    'status': 'error',
                    'message': 'No se pudo generar la configuración de la troncal {}.'.format(tipo)
                }), 400

            salida_escritura, error_escritura, codigo_escritura = _vp_trunk_write(
                perfil,
                tipo,
                nombre,
                contenido
            )
            if codigo_escritura != 0 or 'TRONCAL_ESCRITA' not in str(salida_escritura or ''):
                detalle = str(error_escritura or salida_escritura or '').strip()
                return jsonify({
                    'status': 'error',
                    'message': (
                        'No se pudo escribir la configuración {}: {}'.format(
                            tipo,
                            detalle or 'Asterisk no confirmó la escritura.'
                        )
                    )
                }), 500

            registrada, salida_db, error_db = _vp_trunk_register_issabel(
                perfil,
                tipo,
                nombre,
                host,
                port
            )
            if not registrada:
                detalle = str(error_db or salida_db or '').strip()
                logger.error(
                    '[TRONCAL ADD] Configuración %s escrita pero no registrada en Issabel: %s',
                    nombre,
                    detalle
                )
                return jsonify({
                    'status': 'error',
                    'message': (
                        'Se escribió la configuración de {} en el servidor, '
                        'pero Issabel no confirmó el registro en su base de datos: {}'
                    ).format(tipo, detalle or 'sin detalle')
                }), 500

            salida_reload, error_reload, codigo_reload = _vp_trunk_reload(
                perfil,
                tipo
            )
            if codigo_reload != 0:
                detalle = str(error_reload or salida_reload or '').strip()
                logger.warning(
                    '[TRONCAL ADD] %s creada; no se pudo recargar %s: %s',
                    nombre,
                    tipo,
                    detalle
                )
                return jsonify({
                    'status': 'warning',
                    'message': (
                        'La troncal {} fue guardada y registrada en Issabel, '
                        'pero no se pudo recargar el módulo {}: {}. '
                        'Podés intentar Recargar desde la lista de troncales.'
                    ).format(nombre, tipo, detalle or 'sin detalle')
                })

            verificada, salida_verify, error_verify = _vp_trunk_verify(
                perfil,
                tipo,
                nombre
            )
            if not verificada:
                detalle = str(error_verify or salida_verify or '').strip()
                logger.warning(
                    '[TRONCAL ADD] %s creada; no se pudo confirmar su presencia en Asterisk: %s',
                    nombre,
                    detalle
                )
                return jsonify({
                    'status': 'warning',
                    'message': (
                        'La troncal {} fue guardada, registrada y se solicitó '
                        'la recarga de {}, pero Asterisk no confirmó su presencia. '
                        'Verificá el peer y el módulo OOH323. {}'
                    ).format(nombre, tipo, detalle)
                })

            logger.info(
                '[TRONCAL ADD] Troncal %s tipo=%s creada en perfil=%s host=%s puerto=%s',
                nombre,
                tipo,
                perfil_nombre,
                host,
                port
            )
            return jsonify({
                'status': 'ok',
                'message': 'Troncal {} creada correctamente en {} ({}:{}).'.format(
                    nombre,
                    perfil_nombre,
                    tipo,
                    port
                )
            })

    except Exception as e:
        logger.exception(
            "[TRONCAL ADD] Error creando troncal"
        )

        return jsonify({
            'status': 'error',
            'message': str(e)
        }), 500


@main_bp.route(
    '/api/monitoreo/troncal/eliminar',
    methods=['POST']
)
def api_monitoreo_troncal_eliminar():
    """
    Elimina una troncal creada desde el Monitor.

    SIP:
      - elimina su bloque de sip_custom.conf

    IAX2:
      - elimina su bloque de iax_custom.conf

    OOH323:
      - mantiene la lógica específica existente.

    La eliminación de configuración se realiza antes de borrar
    la fila de Issabel para evitar que el Monitor vuelva a detectar
    el peer desde Asterisk.
    """

    try:
        data = request.get_json(silent=True) or {}

        nombre = str(
            data.get("nombre")
            or data.get("name")
            or data.get("troncal")
            or ""
        ).strip()

        tipo_recibido = str(
            data.get("tipo")
            or data.get("tech")
            or ""
        ).strip()

        if not nombre:
            return jsonify({
                "ok": False,
                "message": "Falta el nombre de la troncal."
            }), 400

        # ----------------------------------------------------
        # Contraseña administrativa
        # ----------------------------------------------------
        password = (
            data.get("password")
            or data.get("admin_password")
            or data.get("clave")
            or ""
        )

        try:
            expected_password = _vp_trunk_admin_password()
        except Exception:
            expected_password = ""

        if not _vp_trunk_admin_password(data):
            return jsonify({
                "ok": False,
                "message": "Contraseña de administrador incorrecta."
            }), 403

        tipo_in = tipo_recibido.upper()

        if tipo_in == "IAX":
            tipo = "IAX2"
        elif tipo_in in ("SIP", "IAX2", "OOH323", "CUSTOM"):
            tipo = tipo_in
        else:
            tipo = tipo_in

        perfil_nombre = str(
            data.get("perfil_nombre", "")
        ).strip()

        if not perfil_nombre:
            return jsonify({
                "status": "error",
                "message": "Falta seleccionar el servidor/perfil de la troncal."
            }), 400

        perfil = buscar_perfil(
            perfil_nombre,
            cargar_perfiles()
        )

        if not perfil:
            return jsonify({
                "status": "error",
                "message": "Perfil no encontrado."
            }), 404

        def exec_remote(command):
            return _vp_trunk_exec(perfil, command)

        import base64

        nombre_b64 = base64.b64encode(
            nombre.encode("utf-8")
        ).decode("ascii")

        # ----------------------------------------------------
        # Buscar la troncal exacta en Issabel
        # ----------------------------------------------------
        comando_db = r"""
AMPDBUSER="$(awk -F= '$1=="AMPDBUSER"{print substr($0,index($0,"=")+1)}' /etc/amportal.conf)"
AMPDBPASS="$(awk -F= '$1=="AMPDBPASS"{print substr($0,index($0,"=")+1)}' /etc/amportal.conf)"
AMPDBHOST="$(awk -F= '$1=="AMPDBHOST"{print substr($0,index($0,"=")+1)}' /etc/amportal.conf)"
AMPDBNAME="$(awk -F= '$1=="AMPDBNAME"{print substr($0,index($0,"=")+1)}' /etc/amportal.conf)"

NOMBRE="$(printf '%s' '""" + nombre_b64 + r"""' | base64 -d)"

mysql -N -B \
  -u"$AMPDBUSER" \
  -p"$AMPDBPASS" \
  -h"$AMPDBHOST" \
  "$AMPDBNAME" \
  -e "
SELECT trunkid,
       COALESCE(name,''),
       COALESCE(tech,''),
       COALESCE(channelid,'')
FROM trunks
WHERE name='${NOMBRE}'
ORDER BY trunkid
LIMIT 1;
"
"""

        db_out = exec_remote(comando_db)

        if isinstance(db_out, tuple):
            db_out = db_out[0]

        db_out = str(db_out or "").strip()

        logger.info(
            "[TRONCAL DELETE] DB lookup nombre=%s resultado=%s",
            nombre,
            db_out
        )

        # Si SIP/IAX2 ya no existe en asterisk.trunks:
        #
        # NO es un error. El Monitor obtiene estas troncales
        # desde el estado/configuración de Asterisk, por lo que
        # puede quedar un peer cargado aunque Issabel ya no
        # tenga la fila correspondiente.
        #
        # En ese caso seguimos directamente con la eliminación
        # del bloque administrado por VoIP Provisioner.
        #
        # OOH323 mantiene la validación DB existente.
        # ====================================================

        if not db_out:

            if tipo not in ("SIP", "IAX2"):
                return jsonify({
                    "ok": False,
                    "message": (
                        "No se encontró la troncal '%s' en Issabel."
                        % nombre
                    )
                }), 404

            trunkid = ""
            db_name = nombre
            tech_db = tipo
            channelid = ""
            tipo_real = tipo

            logger.info(
                "[TRONCAL DELETE] %s no existe en trunks; "
                "se continuará con eliminación de configuración.",
                nombre
            )

        else:

            primera_linea = db_out.splitlines()[0].strip()
            partes = primera_linea.split("\t")

            if len(partes) < 4:
                partes = primera_linea.split(None, 3)

            if len(partes) < 4:
                return jsonify({
                    "ok": False,
                    "message": (
                        "No se pudo interpretar la troncal '%s'."
                        % nombre
                    )
                }), 500

            trunkid = partes[0].strip()
            db_name = partes[1].strip()
            tech_db = partes[2].strip()
            channelid = partes[3].strip()

            tipo_real = tipo

        if tech_db.lower() == "custom":
            if channelid.upper().startswith("OOH323/"):
                tipo_real = "OOH323"
            elif channelid.upper().startswith("SIP/"):
                tipo_real = "SIP"
            elif channelid.upper().startswith("IAX2/"):
                tipo_real = "IAX2"
            elif channelid.upper().startswith("IAX/"):
                tipo_real = "IAX2"

        logger.info(
            "[TRONCAL DELETE] trunkid=%s name=%s tech=%s channelid=%s tipo=%s",
            trunkid,
            db_name,
            tech_db,
            channelid,
            tipo_real
        )

        # ====================================================
        # OOH323
        #
        # NO MODIFICAR LA LÓGICA EXISTENTE.
        # ====================================================
        if tipo_real == "OOH323":

            import re

            peer_name = ""

            m = re.match(
                r"^OOH323/([^/]+)/",
                channelid,
                re.IGNORECASE
            )

            if m:
                peer_name = m.group(1).strip()

            if not peer_name:
                return jsonify({
                    "ok": False,
                    "message": (
                        "No se pudo determinar el peer OOH323 "
                        "de la troncal '%s'."
                    ) % nombre
                }), 409

            peer_b64 = base64.b64encode(
                peer_name.encode("utf-8")
            ).decode("ascii")

            comando_ooh323 = r"""
PEER="$(printf '%s' '""" + peer_b64 + r"""' | base64 -d)"

FOUND=0

for FILE in /etc/asterisk/ooh323.conf /etc/asterisk/ooh323_custom.conf
do
    if [ -f "$FILE" ] && grep -q -E "^[[:space:]]*\[$PEER\][[:space:]]*$" "$FILE"
    then
        cp -p "$FILE" "$FILE.vp-backup-$(date +%Y%m%d-%H%M%S)"

        awk -v target="$PEER" '
        BEGIN {
            skip=0
            pending_header=0
        }

        /^; ============================================================$/ {
            if (skip == 1) {
                next
            }

            pending_header=1
            header=$0
            next
        }

        pending_header == 1 && /^; VOIP PROVISIONER - TRONCAL[[:space:]]*$/ {
            vp_header=1
            pending_header=0
            next
        }

        vp_header == 1 && /^; Nombre:[[:space:]]*/ {
            name=$0
            sub(/^; Nombre:[[:space:]]*/, "", name)

            if (name == target) {
                skip=1
                vp_header=0
                next
            }

            vp_header=0

            print header
            print "; VOIP PROVISIONER - TRONCAL"
            print $0
            next
        }

        pending_header == 1 {
            print header
            pending_header=0
        }

        skip == 1 {
            if ($0 ~ /^; ============================================================$/) {
                skip=0
                pending_header=1
                header=$0
            }
            next
        }

        {
            print
        }
        ' "$FILE" > "$FILE.vp.tmp"

        if [ $? -ne 0 ]
        then
            rm -f "$FILE.vp.tmp"
            echo "ERROR:AWK:$FILE"
            exit 20
        fi

        mv "$FILE.vp.tmp" "$FILE"

        FOUND=1
    fi
done

if [ "$FOUND" -eq 0 ]
then
    echo "PEER_NOT_FOUND:$PEER"
    exit 21
fi

echo "PEER_REMOVED:$PEER"
"""

            config_out = exec_remote(comando_ooh323)

            if isinstance(config_out, tuple):
                config_out = config_out[0]

            config_out = str(config_out or "").strip()

            if "PEER_NOT_FOUND:" in config_out:
                return jsonify({
                    "ok": False,
                    "message": (
                        "La troncal '%s' usa OOH323 pero no se encontró "
                        "el peer en la configuración. No se eliminó "
                        "la troncal de la base de datos."
                    ) % nombre
                }), 409

            if "ERROR:AWK:" in config_out:
                return jsonify({
                    "ok": False,
                    "message": (
                        "No se pudo modificar la configuración OOH323. "
                        "No se eliminó la troncal de la base de datos."
                    )
                }), 500

            try:
                _vp_trunk_reload(perfil, "OOH323")
            except Exception as exc:
                logger.warning(
                    "[TRONCAL DELETE] Error recargando OOH323: %s",
                    exc
                )

        # ====================================================
        # SIP / IAX2
        # ====================================================
        else:

            if tipo_real not in ("SIP", "IAX2"):
                return jsonify({
                    "ok": False,
                    "message": (
                        "Tipo de troncal no soportado para eliminación: %s"
                        % tipo_real
                    )
                }), 400

            if db_name and db_name != nombre:
                return jsonify({
                    "ok": False,
                    "message": (
                        "La troncal encontrada no coincide "
                        "con la solicitada."
                    )
                }), 409

            if tipo_real == "SIP":
                custom_file = "/etc/asterisk/sip_custom.conf"
            else:
                custom_file = "/etc/asterisk/iax_custom.conf"

            custom_b64 = base64.b64encode(
                custom_file.encode("utf-8")
            ).decode("ascii")

            # ------------------------------------------------
            # Eliminación eficiente:
            #
            # Solo elimina bloques administrados por VP cuyo
            # comentario "; Nombre:" coincide exactamente.
            #
            # No toca otros peers del archivo.
            # ------------------------------------------------
            comando_remove = r"""
FILE="$(printf '%s' '""" + custom_b64 + r"""' | base64 -d)"
TARGET="$(printf '%s' '""" + nombre_b64 + r"""' | base64 -d)"

if [ ! -f "$FILE" ]
then
    echo "FILE_NOT_FOUND:$FILE"
    exit 0
fi

if ! grep -q -F "; Nombre: $TARGET" "$FILE"
then
    echo "BLOCK_NOT_FOUND:$TARGET"
    exit 0
fi

cp -p "$FILE" "$FILE.vp-backup-$(date +%Y%m%d-%H%M%S)"

awk -v target="$TARGET" '
BEGIN {
    skip=0
}

# Inicio de un bloque VP.
$0 == "; ============================================================" {
    if (skip == 1) {
        skip=0
    }

    header=$0
    in_header=1
    next
}

# Identificación del bloque VP.
in_header == 1 && $0 == "; VOIP PROVISIONER - TRONCAL" {
    vp=1
    in_header=0
    next
}

# Nombre del bloque.
vp == 1 && $0 ~ /^; Nombre:[[:space:]]*/ {
    nombre=$0
    sub(/^; Nombre:[[:space:]]*/, "", nombre)

    if (nombre == target) {
        skip=1
        vp=0
        in_header=0
        next
    }

    print header
    print "; VOIP PROVISIONER - TRONCAL"
    print $0

    vp=0
    in_header=0
    next
}

# Si no era bloque VP, conservar el encabezado.
in_header == 1 {
    print header
    print
    in_header=0
    next
}

# Mientras eliminamos el bloque, no imprimir nada.
skip == 1 {
    next
}

{
    print
}
' "$FILE" > "$FILE.vp.tmp"

if [ $? -ne 0 ]
then
    rm -f "$FILE.vp.tmp"
    echo "ERROR_AWK:$FILE"
    exit 20
fi

mv "$FILE.vp.tmp" "$FILE"

echo "BLOCK_REMOVED:$FILE"
"""

            config_out = exec_remote(comando_remove)

            if isinstance(config_out, tuple):
                config_out = config_out[0]

            config_out = str(config_out or "").strip()

            logger.info(
                "[TRONCAL DELETE] Config %s: %s",
                tipo_real,
                config_out
            )

            if "ERROR_AWK:" in config_out:
                return jsonify({
                    "ok": False,
                    "message": (
                        "No se pudo modificar %s. "
                        "La troncal no fue eliminada de la base de datos."
                    ) % custom_file
                }), 500

            # -----------------------------------------------
            # Recargar solamente la tecnología eliminada.
            # -----------------------------------------------
            try:
                reload_out = _vp_trunk_reload(
                    perfil,
                    tipo_real
                )
            except Exception as exc:
                reload_out = "ERROR_RELOAD:%s" % exc

            logger.info(
                "[TRONCAL DELETE] Reload %s: %s",
                tipo_real,
                reload_out
            )

        # VP_PATCH_TRONCALES_V1: limpieza completa antes de borrar la fila.
        try:
            _vp_resumen = _vp_trunk_purge(perfil, nombre, tipo_real, trunkid)
            logger.info("[TRONCAL DELETE] Purga %s: %s", nombre, _vp_resumen)
        except Exception as _vp_exc:
            logger.warning("[TRONCAL DELETE] Error en purga: %s", _vp_exc)

        # ====================================================
        # ELIMINAR DE ISSABEL
        # ====================================================
        # ----------------------------------------------------
        # Borrar de la DB de Issabel solamente si existe
        # una fila real. Si trunkid está vacío, ya fue eliminada
        # previamente y no debe considerarse un error.
        # ----------------------------------------------------

        if trunkid:

            trunkid_safe = str(trunkid).replace("'", "\\'")

            comando_delete_db = r"""
    AMPDBUSER="$(awk -F= '$1=="AMPDBUSER"{print substr($0,index($0,"=")+1)}' /etc/amportal.conf)"
    AMPDBPASS="$(awk -F= '$1=="AMPDBPASS"{print substr($0,index($0,"=")+1)}' /etc/amportal.conf)"
    AMPDBHOST="$(awk -F= '$1=="AMPDBHOST"{print substr($0,index($0,"=")+1)}' /etc/amportal.conf)"
    AMPDBNAME="$(awk -F= '$1=="AMPDBNAME"{print substr($0,index($0,"=")+1)}' /etc/amportal.conf)"

    TRUNKID='""" + trunkid_safe + r"""'

    mysql -N -B \
      -u"$AMPDBUSER" \
      -p"$AMPDBPASS" \
      -h"$AMPDBHOST" \
      "$AMPDBNAME" \
      -e "
    DELETE FROM trunk_dialpatterns
    WHERE trunkid='${TRUNKID}';

    DELETE FROM trunks
    WHERE trunkid='${TRUNKID}';

    SELECT ROW_COUNT();
    "
    """

            delete_out = exec_remote(comando_delete_db)

            if isinstance(delete_out, tuple):
                delete_out = delete_out[0]

            delete_out = str(delete_out or "").strip()

            logger.info(
                "[TRONCAL DELETE] DB delete trunkid=%s resultado=%s",
                trunkid,
                delete_out
            )

            # ----------------------------------------------------
            # Verificación final
            # ----------------------------------------------------
            comando_verify = r"""
    AMPDBUSER="$(awk -F= '$1=="AMPDBUSER"{print substr($0,index($0,"=")+1)}' /etc/amportal.conf)"
    AMPDBPASS="$(awk -F= '$1=="AMPDBPASS"{print substr($0,index($0,"=")+1)}' /etc/amportal.conf)"
    AMPDBHOST="$(awk -F= '$1=="AMPDBHOST"{print substr($0,index($0,"=")+1)}' /etc/amportal.conf)"
    AMPDBNAME="$(awk -F= '$1=="AMPDBNAME"{print substr($0,index($0,"=")+1)}' /etc/amportal.conf)"

    NOMBRE="$(printf '%s' '""" + nombre_b64 + r"""' | base64 -d)"

    mysql -N -B \
      -u"$AMPDBUSER" \
      -p"$AMPDBPASS" \
      -h"$AMPDBHOST" \
      "$AMPDBNAME" \
      -e "
    SELECT COUNT(*)
    FROM trunks
    WHERE name='${NOMBRE}';
    "
    """

            verify_db = exec_remote(comando_verify)

            if isinstance(verify_db, tuple):
                verify_db = verify_db[0]

            verify_db = str(verify_db or "").strip()

            if verify_db != "0":
                return jsonify({
                    "ok": False,
                    "message": (
                        "La configuración fue procesada pero la troncal "
                        "todavía existe en la base de datos."
                    ),
                    "tipo": tipo_real,
                    "trunkid": trunkid
                }), 500

            logger.info(
                "[TRONCAL DELETE] OK nombre=%s tipo=%s trunkid=%s",
                nombre,
                tipo_real,
                trunkid
            )

        # VP_PATCH_TRONCALES_V1: regenerar config de Issabel y recargar.
        try:
            _vp_trunk_finalize(perfil, tipo_real, nombre)
        except Exception as _vp_exc:
            logger.warning("[TRONCAL DELETE] Error finalizando: %s", _vp_exc)

        return jsonify({
            "ok": True,
            "status": "ok",
            "message": "Troncal eliminada correctamente.",
            "nombre": nombre,
            "tipo": tipo_real,
            "trunkid": trunkid
        })

    except Exception as e:
        logger.exception(
            "[TRONCAL DELETE] Error eliminando troncal"
        )

        return jsonify({
            "ok": False,
            "status": "error",
            "message": str(e)
        }), 500



# ===== VP_PATCH_TRONCALES_V1 =====
# Eliminacion COMPLETA de troncales y alta real de troncales IAX2 en Issabel.

def _vp_trunk_run_sql(perfil, sql, timeout=60):
    """Ejecuta SQL en la base de Issabel (via base64, sin problemas de comillas)."""
    sql_b64 = base64.b64encode(sql.encode('utf-8')).decode('ascii')
    comando = r"""
AMPDBUSER="$(awk -F= '$1=="AMPDBUSER"{print substr($0,index($0,"=")+1)}' /etc/amportal.conf)"
AMPDBPASS="$(awk -F= '$1=="AMPDBPASS"{print substr($0,index($0,"=")+1)}' /etc/amportal.conf)"
AMPDBHOST="$(awk -F= '$1=="AMPDBHOST"{print substr($0,index($0,"=")+1)}' /etc/amportal.conf)"
AMPDBNAME="$(awk -F= '$1=="AMPDBNAME"{print substr($0,index($0,"=")+1)}' /etc/amportal.conf)"
printf '%s' '""" + sql_b64 + r"""' | base64 -d | mysql -f -N -B -u"$AMPDBUSER" -p"$AMPDBPASS" ${AMPDBHOST:+-h"$AMPDBHOST"} "$AMPDBNAME" 2>&1
echo "VP_SQL_RC=$?"
"""
    out, err, code = _vp_trunk_exec(perfil, comando, timeout=timeout)
    out = str(out or '')
    ok = ('VP_SQL_RC=0' in out)
    return ok, out.replace('VP_SQL_RC=0', '').strip(), str(err or '')


def _vp_trunk_sql_q(valor):
    v = str(valor if valor is not None else '')
    return "'" + v.replace('\\', '\\\\').replace("'", "\\'") + "'"


def _vp_trunk_apply_issabel(perfil, tipo):
    """
    Regenera la configuración de Issabel y recarga la tecnología indicada.

    Para IAX2 además:
      - ejecuta retrieve_conf
      - hace iax2 reload
      - hace dialplan reload
      - devuelve correctamente el código de retorno
    """
    tipo = _vp_trunk_tipo(tipo)

    if tipo == 'IAX2':
        comando = (
            "/var/lib/asterisk/bin/retrieve_conf 2>&1; "
            "RC_RETRIEVE=$?; "
            "asterisk -rx 'iax2 reload' 2>&1; "
            "RC_IAX=$?; "
            "asterisk -rx 'dialplan reload' 2>&1; "
            "RC_DP=$?; "
            "echo VP_APPLY_RETRIEVE=$RC_RETRIEVE; "
            "echo VP_APPLY_IAX=$RC_IAX; "
            "echo VP_APPLY_DIALPLAN=$RC_DP; "
            "if [ $RC_RETRIEVE -ne 0 ] || [ $RC_IAX -ne 0 ] || [ $RC_DP -ne 0 ]; "
            "then exit 1; fi"
        )
    else:
        recarga = {
            'SIP': 'sip reload'
        }.get(tipo, 'core reload')

        comando = (
            "/var/lib/asterisk/bin/retrieve_conf 2>&1; "
            "RC_RETRIEVE=$?; "
            "asterisk -rx '" + recarga + "' 2>&1; "
            "RC_RELOAD=$?; "
            "if [ $RC_RETRIEVE -ne 0 ] || [ $RC_RELOAD -ne 0 ]; "
            "then exit 1; fi"
        )

    return _vp_trunk_exec(
        perfil,
        comando,
        timeout=120
    )


def _vp_iax_obtener_trunkid(perfil, nombre):
    """
    Obtiene el trunkid real de una troncal IAX2 por nombre.
    """
    ok, out, err = _vp_trunk_run_sql(
        perfil,
        "SELECT trunkid "
        "FROM trunks "
        "WHERE name=" +
        _vp_trunk_sql_q(nombre) +
        " AND tech='iax' "
        "ORDER BY trunkid DESC "
        "LIMIT 1;"
    )

    if not ok:
        return '', (
            err or
            out or
            'No se pudo obtener el trunkid de IAX2.'
        )

    trunkid = str(out or '').strip().splitlines()[0].strip()         if str(out or '').strip() else ''

    return trunkid, ''


def _vp_iax_crear_ruta_dialplan(perfil, nombre_troncal, dialplan):
    """
    Crea/reutiliza una ruta saliente para una troncal IAX2.

    Reglas:
      - Si ya existe una ruta con el mismo nombre, se reutiliza.
      - Nunca borra patrones ni asociaciones existentes de una ruta reutilizada.
      - La troncal IAX2 nueva se agrega al FINAL de la ruta.
      - Una ruta nueva se agrega al FINAL del orden de rutas.
      - Los patrones del Dial Plan se cargan solamente al crear una ruta nueva.
    """

    patrones = []

    if isinstance(dialplan, (list, tuple)):
        valores = dialplan
    else:
        valores = str(dialplan or '').splitlines()

    for valor in valores:
        valor = str(valor or '').strip()

        if not valor:
            continue

        partes = [
            x.strip()
            for x in valor.split(',')
            if x.strip()
        ]

        for patron in partes:
            if patron not in patrones:
                patrones.append(patron)

    if not patrones:
        return True, '', 'DIALPLAN_VACIO'

    nombre_troncal = str(nombre_troncal or '').strip()

    if not nombre_troncal:
        return False, '', 'Nombre de troncal vacío.'

    trunkid, err = _vp_iax_obtener_trunkid(
        perfil,
        nombre_troncal
    )

    if not trunkid:
        return False, '', err or (
            'No se encontró el trunkid de la troncal IAX2.'
        )

    # --------------------------------------------------------
    # BUSCAR RUTA EXISTENTE POR NOMBRE REAL
    #
    # Primero se busca exactamente el nombre de la troncal.
    # Esto evita crear una ruta duplicada.
    #
    # Como compatibilidad con rutas creadas por versiones
    # anteriores, también se acepta VP-IAX2-<nombre>.
    # --------------------------------------------------------

    nombres_ruta = [
        nombre_troncal,
        'VP-IAX2-' + nombre_troncal
    ]

    route_id = ''
    ruta_existente_nombre = ''

    for candidato in nombres_ruta:

        ok, out, err = _vp_trunk_run_sql(
            perfil,
            "SELECT route_id,name "
            "FROM outbound_routes "
            "WHERE name=" +
            _vp_trunk_sql_q(candidato) +
            " "
            "ORDER BY route_id DESC "
            "LIMIT 1;"
        )

        if not ok:
            return False, out, (
                'No se pudo consultar outbound_routes: ' +
                (err or out or '')
            )

        linea = str(out or '').strip()

        if linea:
            partes = linea.splitlines()[0].split('\t')

            route_id = str(partes[0]).strip()

            if len(partes) > 1:
                ruta_existente_nombre = str(partes[1]).strip()
            else:
                ruta_existente_nombre = candidato

            break

    # --------------------------------------------------------
    # SI LA RUTA YA EXISTE
    #
    # NO tocar sus patrones.
    # NO borrar sus otras troncales.
    # Solamente asegurar que la nueva troncal quede asociada
    # al final.
    # --------------------------------------------------------

    if route_id:

        sql = ""

        # ----------------------------------------------------
        # Verificar si la troncal ya está asociada.
        # ----------------------------------------------------

        ok, out, err = _vp_trunk_run_sql(
            perfil,
            "SELECT COUNT(*) "
            "FROM outbound_route_trunks "
            "WHERE route_id=" + route_id +
            " AND trunk_id=" + trunkid + ";"
        )

        if not ok:
            return False, out, (
                'No se pudo verificar la asociación ruta/troncal: ' +
                (err or out or '')
            )

        asociada = str(out or '').strip() not in ('', '0')

        if not asociada:

            # La nueva troncal siempre queda última.
            sql += (
                "INSERT INTO outbound_route_trunks "
                "(route_id,trunk_id,seq) "
                "SELECT "
                + route_id + ","
                + trunkid + ","
                "COALESCE(MAX(seq),-1)+1 "
                "FROM outbound_route_trunks "
                "WHERE route_id=" + route_id + ";"
            )

        # ----------------------------------------------------
        # Asegurar que la ruta tenga posición.
        #
        # Si ya la tiene, no se modifica.
        # Si no la tiene, se agrega al final.
        # ----------------------------------------------------

        ok, out, err = _vp_trunk_run_sql(
            perfil,
            "SELECT COUNT(*) "
            "FROM outbound_route_sequence "
            "WHERE route_id=" + route_id + ";"
        )

        if not ok:
            return False, out, (
                'No se pudo verificar el orden de la ruta: ' +
                (err or out or '')
            )

        tiene_secuencia = str(out or '').strip() not in ('', '0')

        if not tiene_secuencia:
            sql += (
                "INSERT INTO outbound_route_sequence "
                "(route_id,seq) "
                "SELECT "
                + route_id + ","
                "COALESCE(MAX(seq),-1)+1 "
                "FROM outbound_route_sequence;"
            )

        if not sql:
            sql = (
                "SELECT CONCAT("
                "'VP_ROUTE_OK|',"
                + route_id +
                ",'|TRUNK|'," +
                trunkid +
                ");"
            )
        else:
            sql += (
                "\nSELECT CONCAT("
                "'VP_ROUTE_OK|',"
                + route_id +
                ",'|TRUNK|'," +
                trunkid +
                ");"
            )

        ok, out, err = _vp_trunk_run_sql(
            perfil,
            sql
        )

        if not ok:
            return False, out, (
                'No se pudo asociar la troncal a la ruta existente: ' +
                (err or out or '')
            )

        if 'VP_ROUTE_OK|' not in str(out or ''):
            return False, out, (
                'La asociación ruta/troncal no devolvió confirmación.'
            )

        return True, out, ''

    # --------------------------------------------------------
    # RUTA NUEVA
    #
    # Se crea con el nombre real de la troncal.
    # --------------------------------------------------------

    ruta_nombre = nombre_troncal

    sql = (
        "INSERT INTO outbound_routes "
        "("
        "name,"
        "outcid,"
        "outcid_mode,"
        "password,"
        "emergency_route,"
        "intracompany_route,"
        "mohclass,"
        "time_group_id"
        ") "
        "VALUES "
        "("
        + _vp_trunk_sql_q(ruta_nombre) + ","
        "'',"
        "'',"
        "'',"
        "'NO',"
        "'NO',"
        "'',"
        "0"
        ");\n"

        "SET @VP_ROUTE_ID=LAST_INSERT_ID();\n"

        "SELECT CONCAT('VP_ROUTE_ID|',@VP_ROUTE_ID);"
    )

    ok_create, out_create, err_create = _vp_trunk_run_sql(
        perfil,
        sql
    )

    if not ok_create:
        return False, out_create, (
            'No se pudo crear la ruta saliente: ' +
            (err_create or out_create or '')
        )

    match = re.search(
        r'VP_ROUTE_ID\|(\d+)',
        str(out_create or '')
    )

    if not match:
        return False, out_create, (
            'Se creó la ruta pero no se pudo obtener su route_id.'
        )

    route_id = match.group(1)

    # --------------------------------------------------------
    # PATRONES DEL DIAL PLAN
    # --------------------------------------------------------

    sql = ""

    for patron in patrones:

        patron = re.sub(
            r'[^0-9A-Za-zXxNnZz\*#\+\.\[\]-]',
            '',
            patron
        )

        if not patron:
            continue

        sql += (
            "INSERT INTO outbound_route_patterns "
            "("
            "route_id,"
            "match_pattern_prefix,"
            "match_pattern_pass,"
            "match_cid,"
            "prepend_digits"
            ") VALUES ("
            + route_id + ","
            "'',"
            + _vp_trunk_sql_q(patron) + ","
            "'',"
            "''"
            ");\n"
        )

    # --------------------------------------------------------
    # TRONCAL DE LA RUTA
    #
    # Al ser nueva, queda como primera troncal de esa ruta.
    # Si posteriormente se agregan otras, las nuevas irán
    # al final.
    # --------------------------------------------------------

    sql += (
        "INSERT INTO outbound_route_trunks "
        "(route_id,trunk_id,seq) "
        "VALUES ("
        + route_id + ","
        + trunkid + ","
        "0"
        ");\n"
    )

    # --------------------------------------------------------
    # POSICIÓN DE LA RUTA
    #
    # MAX(seq)+1 = última posición.
    # --------------------------------------------------------

    sql += (
        "INSERT INTO outbound_route_sequence "
        "(route_id,seq) "
        "SELECT "
        + route_id + ","
        "COALESCE(MAX(seq),-1)+1 "
        "FROM outbound_route_sequence;\n"
    )

    sql += (
        "SELECT CONCAT("
        "'VP_ROUTE_OK|',"
        + route_id +
        ",'|TRUNK|'," +
        trunkid +
        ");"
    )

    ok, out, err = _vp_trunk_run_sql(
        perfil,
        sql
    )

    if not ok:
        return False, out, (
            'No se pudo configurar la ruta saliente: ' +
            (err or out or '')
        )

    if 'VP_ROUTE_OK|' not in str(out or ''):
        return False, out, (
            'La ruta saliente no devolvió confirmación.'
        )

    return True, out, ''

def _vp_iax_verificar_aplicacion(perfil, nombre):
    """
    Verifica que la troncal IAX2 esté realmente cargada en Asterisk.
    """
    comando = (
        "asterisk -rx 'iax2 show peers' 2>&1"
    )

    out, err, code = _vp_trunk_exec(
        perfil,
        comando,
        timeout=30
    )

    salida = str(out or '')

    if code != 0:
        return False, (
            'Asterisk devolvió código ' +
            str(code) +
            ': ' +
            str(err or salida)
        )

    objetivo = str(nombre or '').strip().lower()

    for linea in salida.splitlines():

        partes = linea.split()

        if not partes:
            continue

        peer = partes[0].split('/')[0].strip().lower()

        if peer == objetivo:
            return True, ''

    return False, (
        'La troncal "' +
        str(nombre) +
        '" no aparece en "iax2 show peers" '
        'después de aplicar la configuración.'
    )


def _vp_trunk_purge(perfil, nombre, tipo, trunkid):
    """
    Borra TODO rastro de la troncal (salvo la fila de 'trunks', que borra
    el codigo original):
      - bloque completo en sip_custom.conf / iax_custom.conf
      - filas de las tablas sip / iax de Issabel (peer, user, register)
      - asignaciones a rutas salientes
    """
    tipo = _vp_trunk_tipo(tipo)
    resumen = []

    custom = {
        'SIP': '/etc/asterisk/sip_custom.conf',
        'IAX2': '/etc/asterisk/iax_custom.conf'
    }.get(tipo, '')

    if custom:
        comando = r"""
FILE=""" + shlex.quote(custom) + r"""
if [ -f "$FILE" ]; then
    cp -p "$FILE" "$FILE.vp-backup-$(date +%Y%m%d-%H%M%S)"
    VP_T=""" + shlex.quote(nombre) + r""" perl -0777 -i -pe 'BEGIN{$t=quotemeta($ENV{VP_T})} s/(?:^;[ \t]*=+[ \t]*\n;[ \t]*VOIP PROVISIONER - TRONCAL[ \t]*\n;[ \t]*Nombre:[ \t]*$t[ \t]*\n(?:;[^\n]*\n)*?;[ \t]*=+[ \t]*\n)?\n*^\[$t\][ \t]*\n(?:(?!\[|;[ \t]*=)[^\n]*\n?)*//mg' "$FILE"
    echo "CUSTOM_LIMPIO"
fi
"""
        out, err, code = _vp_trunk_exec(perfil, comando, timeout=30)
        resumen.append(str(out or '').strip() or str(err or '').strip())

    tid = re.sub(r'[^0-9]', '', str(trunkid or ''))

    if tid:
        sql = (
            "DELETE FROM outbound_route_trunks WHERE trunk_id=" + tid + ";\n"
            "DELETE FROM trunk_dialpatterns WHERE trunkid=" + tid + ";\n"
        )

        if tipo == 'IAX2':
            sql += (
                "DELETE FROM iax WHERE id IN "
                "('tr-peer-" + tid + "','tr-user-" + tid + "','tr-reg-" + tid + "');\n"
            )
        elif tipo == 'SIP':
            sql += (
                "DELETE FROM sip WHERE id IN "
                "('tr-peer-" + tid + "','tr-user-" + tid + "','tr-reg-" + tid + "');\n"
            )

        ok, out, err = _vp_trunk_run_sql(perfil, sql)
        resumen.append(out or err)

    return resumen


def _vp_trunk_finalize(perfil, tipo, nombre):
    """Regenera la config de Issabel, recarga y comprueba que ya no exista en memoria."""
    tipo = _vp_trunk_tipo(tipo)
    _vp_trunk_apply_issabel(perfil, tipo)

    if tipo not in ('SIP', 'IAX2'):
        return True

    show = 'iax2 show peers' if tipo == 'IAX2' else 'sip show peers'
    mod = 'chan_iax2.so' if tipo == 'IAX2' else 'chan_sip.so'

    def sigue():
        out, err, code = _vp_trunk_exec(
            perfil, "asterisk -rx '" + show + "' 2>/dev/null", timeout=20
        )
        for linea in str(out or '').splitlines():
            partes = linea.split()
            if partes and partes[0].split('/')[0].lower() == nombre.lower():
                return True
        return False

    if sigue():
        _vp_trunk_exec(perfil, "asterisk -rx 'module reload " + mod + "' 2>&1", timeout=30)
        time.sleep(2)

    return not sigue()


def _vp_iax_normalizar_nombre_servidor(nombre):
    """
    Normaliza únicamente para generar las credenciales IAX2:
    - minúsculas
    - elimina espacios
    No guarda este valor en configuración.
    """
    nombre = str(nombre or '').strip().lower()
    nombre = re.sub(r'\s+', '', nombre)
    return nombre


def _vp_iax_credenciales_perfil(perfil):
    """
    Genera las credenciales IAX2 determinísticas a partir del nombre
    del perfil. No se almacenan en JSON.
    """
    nombre = _vp_iax_normalizar_nombre_servidor(
        (perfil or {}).get('nombre', '')
    )

    if not nombre:
        return '', ''

    return (
        '*voip' + nombre + 'USER',
        '*voip' + nombre + 'PASS'
    )


def _vp_trunk_create_iax_issabel(
    perfil,
    nombre,
    descripcion,
    host,
    port,
    peer_usuario=None,
    peer_clave=None,
    user_usuario=None,
    user_clave=None
):
    """
    Crea una troncal IAX2 Issabel.

    TODAS las troncales IAX2 del VoIP Provisioner utilizan:

        username = *voippjn
        secret   = *voippjn

    Tanto el peer como el user utilizan las mismas credenciales.

    No se utilizan nombres de perfiles para generar usuarios.
    No se guardan credenciales en JSON.
    """

    IAX_USER = "*voippjn"
    IAX_PASS = "*voippjn"

    nombre = str(nombre or '').strip()
    host = str(host or '').strip()

    try:
        port = int(port or 4569)
    except Exception:
        return False, '', 'Puerto IAX2 inválido.'

    # --------------------------------------------------------
    # VALIDACION COMPLETA ANTES DE INSERTAR
    # --------------------------------------------------------

    errores = []

    if not nombre:
        errores.append('nombre de troncal vacío')

    if not host:
        errores.append('host remoto vacío')

    if port < 1 or port > 65535:
        errores.append('puerto fuera de rango')

    if not IAX_USER or not IAX_PASS:
        errores.append('credenciales IAX2 inválidas')

    if errores:
        return False, '', (
            'Validación IAX2: ' +
            ', '.join(errores)
        )

    # --------------------------------------------------------
    # VERIFICAR QUE NO EXISTA LA TRONCAL
    # --------------------------------------------------------

    ok, out, err = _vp_trunk_run_sql(
        perfil,
        "SELECT COUNT(*) FROM trunks WHERE name=" +
        _vp_trunk_sql_q(nombre) +
        ";"
    )

    if not ok:
        return False, out, (
            'No se pudo validar la troncal IAX2: ' +
            (err or out or '')[:500]
        )

    if out.strip() not in ('0', ''):
        return False, out, (
            'Ya existe una troncal con nombre "' +
            nombre +
            '".'
        )

    # --------------------------------------------------------
    # PEER
    #
    # Siempre usa *voippjn
    # --------------------------------------------------------

    peer = [
        ('account', nombre),
        ('host', host),
        ('username', IAX_USER),
        ('secret', IAX_PASS),
        ('encryption', 'aes128'),
        ('auth', 'md5'),
        ('type', 'friend'),
        ('context', 'from-internal'),
        ('qualify', 'yes'),
        ('transfer', 'yes'),
        ('allow', 'gsm'),
        ('disallow', 'all'),
        ('trunk', 'yes'),
        ('requirecalltoken', 'no')
    ]

    # --------------------------------------------------------
    # USER
    #
    # Siempre usa *voippjn
    # --------------------------------------------------------

    user = [
        ('account', IAX_USER),
        ('secret', IAX_PASS),
        ('type', 'user'),
        ('context', 'from-internal')
    ]

    # --------------------------------------------------------
    # ARMAR SQL COMPLETO
    # --------------------------------------------------------

    sql = (
        "SET @id = (SELECT COALESCE(MAX(trunkid),0)+1 FROM trunks);\n"
        "INSERT INTO trunks "
        "(trunkid,name,tech,outcid,keepcid,maxchans,failscript,"
        "dialoutprefix,channelid,usercontext,provider,disabled,`continue`) "
        "VALUES ("
        "@id," +
        _vp_trunk_sql_q(nombre) +
        ",'iax','','off','','',''," +
        _vp_trunk_sql_q(nombre) +
        "," +
        _vp_trunk_sql_q(IAX_USER) +
        ",'','off','off');\n"
    )

    filas = []

    # Issabel:
    # flags=1 = deshabilitado
    # flags=2 en adelante = activo

    for i, (k, v) in enumerate(peer, start=1):
        flags = i + 1

        filas.append(
            "(CONCAT('tr-peer-',@id)," +
            _vp_trunk_sql_q(k) +
            "," +
            _vp_trunk_sql_q(v) +
            "," +
            str(flags) +
            ")"
        )

    for i, (k, v) in enumerate(user, start=1):
        flags = i + 1

        filas.append(
            "(CONCAT('tr-user-',@id)," +
            _vp_trunk_sql_q(k) +
            "," +
            _vp_trunk_sql_q(v) +
            "," +
            str(flags) +
            ")"
        )

    sql += (
        "INSERT INTO iax (id,keyword,data,flags) VALUES\n" +
        ",\n".join(filas) +
        ";\n"
    )

    sql += (
        "SELECT CONCAT('TRONCAL_IAX_OK|',@id);\n"
    )

    # --------------------------------------------------------
    # EJECUTAR
    # --------------------------------------------------------

    ok, out, err = _vp_trunk_run_sql(
        perfil,
        sql
    )

    if not ok or 'TRONCAL_IAX_OK|' not in out:
        return False, out, (
            'No se pudo crear la troncal IAX2: ' +
            (err or out or '')[:500]
        )

    return True, out, ''




# === VOIP PROVISIONER - ADMIN DESVIO INTERNOS V1 ===

def _vp_admin_desvio_perfil(perfil_nombre):
    """
    Busca el perfil por nombre usando la misma estructura
    de perfiles existente en la aplicación.
    """
    nombre = str(perfil_nombre or '').strip()

    if not nombre:
        return None

    try:
        perfiles = cargar_perfiles()

        if isinstance(perfiles, dict):
            # Formato posible: {"nombre": {...}}
            if nombre in perfiles and isinstance(perfiles[nombre], dict):
                return perfiles[nombre]

            # Formato posible: {"perfiles": [...]}
            lista = perfiles.get('perfiles')
            if isinstance(lista, list):
                perfiles = lista

        if isinstance(perfiles, list):
            for p in perfiles:
                if not isinstance(p, dict):
                    continue

                if str(p.get('nombre', '')).strip() == nombre:
                    return p

    except Exception as e:
        logger.exception("Error buscando perfil para admin de desvio: %s", e)

    return None


def _vp_admin_desvio_password_valida(data):
    try:
        cfg = cargar_config()
        seguridad = cfg.get('seguridad', {}) if isinstance(cfg, dict) else {}
        password_real = str(
            seguridad.get('admin_password', '')
        )

        password_recibida = str(
            (data or {}).get('password', '')
        )

        return bool(password_real) and password_recibida == password_real

    except Exception:
        logger.exception(
            "Error verificando password de admin para desvio"
        )
        return False


def _vp_admin_desvio_interno(valor):
    valor = str(valor or '').strip()

    if not re.match(r'^[0-9]+$', valor):
        return None

    return valor


def _vp_admin_desvio_destino(valor):
    valor = str(valor or '').strip()

    if not valor:
        return None

    # Permitimos internos, números externos y los caracteres
    # normalmente utilizados en destinos telefónicos.
    if not re.match(r'^[0-9*#+]+$', valor):
        return None

    return valor


@main_bp.route('/api/monitoreo/internos/desvios', methods=['GET'])
def api_monitoreo_internos_desvios():
    """
    Devuelve los desvíos incondicionales almacenados en AstDB.

    FreePBX/Issabel utiliza la familia CF para Call Forward All.
    Ejemplo:
        database get CF 1001
    """

    perfil_nombre = str(
        request.args.get('perfil_nombre', '')
    ).strip()

    perfil = _vp_admin_desvio_perfil(perfil_nombre)

    if not perfil:
        return jsonify({
            'ok': False,
            'mensaje': 'Perfil no encontrado.'
        }), 404

    comando = (
        "asterisk -rx "
        "\"database show CF\""
    )

    out, err, rc = ejecutar_comando_ssh(
        perfil,
        comando
    )

    if rc != 0:
        return jsonify({
            'ok': False,
            'mensaje': err or 'No se pudo consultar AstDB.'
        }), 500

    resultado = {}

    for linea in str(out or '').splitlines():

        # Formato:
        # /CF/1001 : 1155551234
        m = re.search(
            r'/CF/([0-9]+)\s*:\s*(.+?)\s*$',
            linea
        )

        if not m:
            continue

        interno = m.group(1).strip()
        destino = m.group(2).strip()

        if interno and destino:
            resultado[interno] = destino

    return jsonify({
        'ok': True,
        'desvios': resultado
    })


@main_bp.route('/api/monitoreo/interno/desvio', methods=['POST'])
def api_monitoreo_interno_desvio():
    """
    Crea/modifica el desvío incondicional de un interno.

    Utiliza:
        database put CF <interno> <destino>

    No requiere retrieve_conf porque AstDB es consultada
    directamente por Asterisk.
    """

    data = request.get_json(silent=True) or {}

    if not _vp_admin_desvio_password_valida(data):
        return jsonify({
            'ok': False,
            'mensaje': 'Contraseña de administrador incorrecta.'
        }), 403

    perfil_nombre = str(
        data.get('perfil_nombre', '')
    ).strip()

    interno = _vp_admin_desvio_interno(
        data.get('interno')
    )

    destino = _vp_admin_desvio_destino(
        data.get('destino')
    )

    if not perfil_nombre:
        return jsonify({
            'ok': False,
            'mensaje': 'Falta el perfil.'
        }), 400

    if not interno:
        return jsonify({
            'ok': False,
            'mensaje': 'Interno inválido.'
        }), 400

    if not destino:
        return jsonify({
            'ok': False,
            'mensaje': 'Destino de desvío inválido.'
        }), 400

    perfil = _vp_admin_desvio_perfil(
        perfil_nombre
    )

    if not perfil:
        return jsonify({
            'ok': False,
            'mensaje': 'Perfil no encontrado.'
        }), 404

    comando = (
        "asterisk -rx "
        "\"database put CF %s %s\""
        % (interno, destino)
    )

    out, err, rc = ejecutar_comando_ssh(
        perfil,
        comando
    )

    if rc != 0:
        return jsonify({
            'ok': False,
            'mensaje': err or out or 'No se pudo guardar el desvío.'
        }), 500

    # Verificación inmediata.
    comando_verificar = (
        "asterisk -rx "
        "\"database get CF %s\""
        % interno
    )

    out2, err2, rc2 = ejecutar_comando_ssh(
        perfil,
        comando_verificar
    )

    if rc2 != 0:
        return jsonify({
            'ok': False,
            'mensaje': 'El desvío fue enviado pero no se pudo verificar.'
        }), 500

    if destino not in str(out2 or ''):
        return jsonify({
            'ok': False,
            'mensaje': (
                'Asterisk respondió pero el desvío no quedó '
                'verificado en AstDB.'
            )
        }), 500

    return jsonify({
        'ok': True,
        'interno': interno,
        'destino': destino,
        'mensaje': 'Desvío guardado correctamente.'
    })


@main_bp.route('/api/monitoreo/interno/desvio', methods=['DELETE'])
def api_monitoreo_interno_desvio_delete():
    """
    Elimina el desvío incondicional de un interno.
    """

    data = request.get_json(silent=True) or {}

    if not _vp_admin_desvio_password_valida(data):
        return jsonify({
            'ok': False,
            'mensaje': 'Contraseña de administrador incorrecta.'
        }), 403

    perfil_nombre = str(
        data.get('perfil_nombre', '')
    ).strip()

    interno = _vp_admin_desvio_interno(
        data.get('interno')
    )

    if not perfil_nombre:
        return jsonify({
            'ok': False,
            'mensaje': 'Falta el perfil.'
        }), 400

    if not interno:
        return jsonify({
            'ok': False,
            'mensaje': 'Interno inválido.'
        }), 400

    perfil = _vp_admin_desvio_perfil(
        perfil_nombre
    )

    if not perfil:
        return jsonify({
            'ok': False,
            'mensaje': 'Perfil no encontrado.'
        }), 404

    comando = (
        "asterisk -rx "
        "\"database del CF %s\""
        % interno
    )

    out, err, rc = ejecutar_comando_ssh(
        perfil,
        comando
    )

    if rc != 0:
        return jsonify({
            'ok': False,
            'mensaje': err or out or 'No se pudo eliminar el desvío.'
        }), 500

    # Verificación.
    comando_verificar = (
        "asterisk -rx "
        "\"database get CF %s\""
        % interno
    )

    out2, err2, rc2 = ejecutar_comando_ssh(
        perfil,
        comando_verificar
    )

    if rc2 == 0 and 'Database entry not found' not in str(out2 or ''):
        return jsonify({
            'ok': False,
            'mensaje': (
                'Asterisk respondió pero el desvío '
                'todavía aparece en AstDB.'
            )
        }), 500

    return jsonify({
        'ok': True,
        'interno': interno,
        'mensaje': 'Desvío eliminado correctamente.'
    })


# === VOIP PROVISIONER - GRUPO CAPTURA ASTERISK V1 ===

# ============================================================
# GRUPOS DE CAPTURA - PERSISTENCIA POR SERVIDOR
#
# El archivo vive en el Issabel remoto:
#
#   /var/lib/voip-provisioner/grupos.json
#
# Si no existe, se crea automáticamente.
# ============================================================

_VP_GRUPOS_JSON_REMOTO = (
    "/var/lib/voip-provisioner/grupos.json"
)


def _vp_grupos_json_remoto_leer(perfil):
    """
    Lee los grupos persistidos en el Issabel remoto.

    Si el archivo no existe:
      - crea el directorio
      - crea grupos.json
      - deja una estructura JSON válida
    """

    import json
    import base64
    import shlex

    try:

        comando = (
            "mkdir -p /var/lib/voip-provisioner && "
            "if [ ! -f {0} ]; then "
            "printf '%s' '{1}' | base64 -d > {0}.tmp && "
            "chmod 600 {0}.tmp && "
            "mv {0}.tmp {0}; "
            "fi; "
            "chmod 600 {0} 2>/dev/null || true; "
            "cat {0}"
        ).format(
            shlex.quote(_VP_GRUPOS_JSON_REMOTO),
            base64.b64encode(
                b'{"version":1,"grupos":[]}'
            ).decode("ascii")
        )

        out, err, rc = ejecutar_comando_ssh(
            perfil,
            comando
        )

        if rc != 0:
            logger.warning(
                "No se pudo leer grupos.json remoto: %s",
                err or out
            )
            return []

        contenido = str(out or "").strip()

        if not contenido:
            return []

        datos = json.loads(contenido)

        if not isinstance(datos, dict):
            return []

        grupos = datos.get("grupos", [])

        if not isinstance(grupos, list):
            return []

        resultado = []

        for grupo in grupos:

            nombre = str(
                grupo or ""
            ).strip()

            if not nombre:
                continue

            if nombre not in resultado:
                resultado.append(nombre)

        return resultado

    except Exception:
        logger.exception(
            "Error leyendo grupos.json remoto"
        )
        return []


def _vp_grupos_numeros_leer(perfil):
    """Lee los numeros administrativos sin alterar el catalogo remoto."""
    import json
    import shlex

    cmd = "cat " + shlex.quote(_VP_GRUPOS_JSON_REMOTO) + " 2>/dev/null"
    try:
        out, err, rc = ejecutar_comando_ssh(perfil, cmd)
        if rc != 0:
            return None

        datos = json.loads(str(out or "").strip() or "{}")
        if not isinstance(datos, dict):
            return None

        numeros = datos.get("numeros", {})
        if not isinstance(numeros, dict):
            return None

        return numeros
    except Exception:
        logger.exception("Error leyendo numeros de grupos")
        return None


def _vp_grupos_numeros_normalizar(grupos, numeros):
    """Conserva numeros validos y asigna el menor libre desde 100."""
    if not isinstance(numeros, dict):
        return None

    nombres = []
    for grupo in grupos or []:
        nombre = str(grupo or "").strip()
        if nombre and nombre not in nombres:
            nombres.append(nombre)

    resultado = {}
    usados = set()

    for nombre in nombres:
        valor = numeros.get(nombre)
        try:
            numero = int(valor)
        except (TypeError, ValueError):
            continue

        if numero >= 100 and numero not in usados:
            resultado[nombre] = numero
            usados.add(numero)

    siguiente = 100
    for nombre in nombres:
        if nombre in resultado:
            continue
        while siguiente in usados:
            siguiente += 1
        resultado[nombre] = siguiente
        usados.add(siguiente)
        siguiente += 1

    return resultado


def _vp_grupos_json_remoto_guardar(perfil, grupos, numeros=None):
    """Guarda nombres y numeros en forma atomica, preservando los existentes."""
    import json
    import base64
    import shlex

    try:
        grupos_limpios = []
        for grupo in grupos or []:
            nombre = str(grupo or "").strip()
            if nombre and nombre not in grupos_limpios:
                grupos_limpios.append(nombre)

        grupos_limpios.sort(key=lambda x: x.lower())

        if numeros is None:
            actuales = _vp_grupos_numeros_leer(perfil)
            if actuales is None:
                logger.error("No se guardan grupos: no se pudo leer la numeracion actual.")
                return False
            numeros = actuales

        numeros_limpios = _vp_grupos_numeros_normalizar(
            grupos_limpios, numeros
        )
        if numeros_limpios is None:
            return False

        datos = {
            "version": 1,
            "grupos": grupos_limpios,
            "numeros": numeros_limpios
        }

        contenido = json.dumps(datos, ensure_ascii=False, indent=2)
        contenido_b64 = base64.b64encode(
            contenido.encode("utf-8")
        ).decode("ascii")

        comando = (
            "mkdir -p /var/lib/voip-provisioner && "
            "printf '%s' {0} | base64 -d > {1}.tmp && "
            "chmod 600 {1}.tmp && "
            "mv {1}.tmp {1}"
        ).format(
            shlex.quote(contenido_b64),
            shlex.quote(_VP_GRUPOS_JSON_REMOTO)
        )

        out, err, rc = ejecutar_comando_ssh(perfil, comando)
        if rc != 0:
            logger.error(
                "No se pudo guardar grupos.json remoto: %s",
                err or out
            )
            return False

        return True

    except Exception:
        logger.exception("Error guardando grupos.json remoto")
        return False


def _vp_grupos_json_remoto_actualizar(
    perfil,
    grupos_nuevos
):
    """
    Une los grupos existentes con los nuevos y guarda
    el resultado en el Issabel remoto.
    """

    actuales = _vp_grupos_json_remoto_leer(
        perfil
    )

    todos = list(actuales)

    for grupo in grupos_nuevos or []:

        nombre = str(
            grupo or ""
        ).strip()

        if nombre and nombre not in todos:
            todos.append(nombre)

    return _vp_grupos_json_remoto_guardar(
        perfil,
        todos
    )



def _vp_grupo_captura_perfil(perfil_nombre):
    nombre = str(perfil_nombre or '').strip()

    if not nombre:
        return None

    try:
        perfiles = cargar_perfiles()

        if isinstance(perfiles, dict):
            if nombre in perfiles and isinstance(perfiles[nombre], dict):
                return perfiles[nombre]

            lista = perfiles.get('perfiles')
            if isinstance(lista, list):
                perfiles = lista

        if isinstance(perfiles, list):
            for p in perfiles:
                if not isinstance(p, dict):
                    continue

                if str(p.get('nombre', '')).strip() == nombre:
                    return p

    except Exception:
        logger.exception(
            "Error buscando perfil para grupo de captura"
        )

    return None


def _vp_grupo_captura_password_valida(data):
    try:
        cfg = cargar_config()
        seguridad = (
            cfg.get('seguridad', {})
            if isinstance(cfg, dict)
            else {}
        )

        password_real = str(
            seguridad.get('admin_password', '')
        )

        password_recibida = str(
            (data or {}).get('password', '')
        )

        return (
            bool(password_real)
            and password_recibida == password_real
        )

    except Exception:
        logger.exception(
            "Error verificando password para grupo de captura"
        )
        return False


def _vp_grupo_captura_interno(valor):
    valor = str(valor or '').strip()

    if not re.match(r'^[0-9]+$', valor):
        return None

    return valor


def _vp_grupo_captura_nombre(valor):
    valor = str(valor or '').strip()

    if not valor:
        return None

    # Permitimos nombres habituales de ubicaciones/grupos:
    # letras, numeros, espacios y algunos separadores.
    if len(valor) > 64:
        return None

    if not re.match(
        r'^[A-Za-z0-9ÁÉÍÓÚáéíóúÑñÜü _.\-]+$',
        valor
    ):
        return None

    return valor


def _vp_grupo_captura_mysql(perfil, sql):
    """
    Ejecuta SQL en la base asterisk del perfil.
    Usa el mismo mecanismo de credenciales del proyecto.
    """
    import base64
    import shlex

    db_user, db_pass = obtener_credenciales_mysql(perfil)

    sql_b64 = base64.b64encode(
        sql.encode('utf-8')
    ).decode('ascii')

    cmd = (
        "SQL_B64={}; "
        "printf '%s' \"$SQL_B64\" | base64 -d | "
        "/usr/bin/mysql -N -B -u {} -p{} "
        "-e \"$(cat)\" asterisk 2>&1"
    ).format(
        shlex.quote(sql_b64),
        shlex.quote(str(db_user)),
        shlex.quote(str(db_pass))
    )

    return ejecutar_comando_ssh(
        perfil,
        cmd
    )



# Ubicación administrativa de internos y enrutamiento IP Office.
_VP_UBICACIONES_INTERNOS_LOCK = threading.Lock()
_VP_UBICACIONES_INTERNOS_VALIDAS = ('Issabel', 'IP Office')
_VP_UBICACIONES_ARCHIVO_REMOTO = '/etc/asterisk/extensions_custom.conf'
_VP_UBICACIONES_MARCA_INICIO = '; BEGIN VOIP-PROVISIONER UBICACIONES INTERNOS'
_VP_UBICACIONES_MARCA_FIN = '; END VOIP-PROVISIONER UBICACIONES INTERNOS'


def _vp_ubicaciones_internos_archivo():
    return os.path.abspath(os.path.join(
        os.path.dirname(__file__), '..', 'data',
        'ubicaciones_internos.json'
    ))


def _vp_ubicaciones_internos_leer():
    ruta = _vp_ubicaciones_internos_archivo()
    if not os.path.exists(ruta):
        return {}
    with open(ruta, 'r') as archivo:
        datos = json.load(archivo)
    if not isinstance(datos, dict):
        raise ValueError('El archivo de ubicaciones no contiene un objeto JSON.')
    return datos


def _vp_ubicaciones_internos_guardar(datos):
    ruta = _vp_ubicaciones_internos_archivo()
    carpeta = os.path.dirname(ruta)
    if not os.path.isdir(carpeta):
        os.makedirs(carpeta)
    temporal = ruta + '.tmp'
    try:
        with open(temporal, 'w') as archivo:
            json.dump(datos, archivo, ensure_ascii=False, indent=2,
                      sort_keys=True)
            archivo.write('\n')
        os.replace(temporal, ruta)
    finally:
        if os.path.exists(temporal):
            os.remove(temporal)



def _vp_ubicaciones_troncal_id(valor):
    valor = str(valor or '').strip()
    if valor == 'IP Office':
        return '2'
    import re
    coincidencia = re.match(r'^IP Office:([0-9]+)$', valor)
    return coincidencia.group(1) if coincidencia else ''


def _vp_ubicaciones_troncales_leer(perfil):
    """Consulta las troncales OOH323 registradas en Issabel."""
    sql = (
        "SELECT trunkid, name, tech, channelid "
        "FROM asterisk.trunks "
        "WHERE UPPER(channelid) LIKE 'OOH323/%' "
        "ORDER BY name, trunkid;"
    )

    out, err, rc = _vp_grupo_captura_mysql(perfil, sql)

    if rc != 0:
        raise RuntimeError(
            str(err or out or
                "No se pudieron consultar las troncales OOH323.")
        )

    troncales = []
    vistos = set()

    for linea in str(out or "").splitlines():
        partes = linea.split("\t")

        if len(partes) < 4:
            continue

        trunkid = str(partes[0] or "").strip()
        nombre = str(partes[1] or "").strip()
        channelid = str(partes[3] or "").strip()

        coincidencia = re.match(
            r"^OOH323/([^/]+)/", channelid, re.IGNORECASE
        )

        if not trunkid.isdigit() or not nombre or not coincidencia:
            continue

        if trunkid in vistos:
            continue

        vistos.add(trunkid)

        troncales.append({
            "id": trunkid,
            "nombre": nombre,
            "peer": coincidencia.group(1),
            "channelid": channelid
        })

    troncales.sort(
        key=lambda t: (t["nombre"].lower(), int(t["id"]))
    )

    return troncales



def _vp_ubicaciones_generar_bloque(
    internos_ipoffice, salto, troncales_internos=None
):
    troncales_internos = troncales_internos or {}

    lineas = [
        _VP_UBICACIONES_MARCA_INICIO,
        '; Administrado por VoIP Provisioner. No editar este bloque.'
    ]

    for interno in sorted(internos_ipoffice, key=lambda valor: int(valor)):
        troncal_id = str(
            troncales_internos.get(str(interno), '2')
        ).strip()

        if not re.match(r'^[0-9]+$', troncal_id):
            raise ValueError(
                'trunkid inválido para el interno {}.'.format(interno)
            )

        lineas.extend([
            'exten => {},1,NoOp(VoIP Provisioner: interno {} en IP Office)'.format(
                interno, interno
            ),
            ' same => n,Macro(user-callerid,LIMIT,EXTERNAL,)',
            ' same => n,Set(MOHCLASS=${IF($["${MOHCLASS}"=""]?default:${MOHCLASS})})',
            ' same => n,Set(_NODEST=)',
            ' same => n,Gosub(sub-record-check,s,1(out,${EXTEN},))',
            ' same => n,Macro(dialout-trunk,' + troncal_id + ',${EXTEN},,off)',
            ' same => n,Macro(outisbusy,)'
        ])

    lineas.append(_VP_UBICACIONES_MARCA_FIN)
    return salto.join(lineas)


def _vp_ubicaciones_preparar_archivo_remoto_base(contenido, internos_ipoffice, troncales_internos=None):
    salto = '\r\n' if '\r\n' in contenido else '\n'
    inicio = contenido.count(_VP_UBICACIONES_MARCA_INICIO)
    fin = contenido.count(_VP_UBICACIONES_MARCA_FIN)

    if inicio != fin or inicio > 1:
        raise ValueError(
            'Los marcadores del bloque administrado no son válidos. '
            'Se cancela para proteger extensions_custom.conf.'
        )

    # Quitar temporalmente el bloque administrado del análisis para no
    # confundir nuestras propias reglas con reglas preexistentes.
    contenido_revision = contenido
    if inicio == 1:
        p_inicio = contenido.find(_VP_UBICACIONES_MARCA_INICIO)
        p_fin = contenido.find(_VP_UBICACIONES_MARCA_FIN, p_inicio)
        if p_fin < 0:
            raise ValueError('No se encontró el cierre del bloque administrado.')
        p_fin += len(_VP_UBICACIONES_MARCA_FIN)
        contenido_revision = (
            contenido[:p_inicio] + contenido[p_fin:]
        )

    patron_seccion = r'(?mi)^\[from-internal-custom\][ \t]*\r?$'
    secciones = list(re.finditer(patron_seccion, contenido_revision))

    if len(secciones) == 0:
        if contenido.strip():
            raise ValueError(
                'El archivo extensions_custom.conf contiene datos, '
                'pero no tiene [from-internal-custom]. '
                'Se cancela para proteger la configuración existente.'
            )

        contenido = '[from-internal-custom]' + salto
        contenido_revision = contenido
        secciones = list(
            re.finditer(patron_seccion, contenido_revision)
        )

    if len(secciones) != 1:
        raise ValueError(
            'Se esperaba exactamente una sección [from-internal-custom]. '
            'No se modificará el archivo remoto.'
        )

    inicio_contexto = secciones[0].end()
    siguiente = re.search(
        r'(?mi)^\[[^\]\r\n]+\][ \t]*\r?$',
        contenido_revision[inicio_contexto:]
    )
    fin_contexto = (
        inicio_contexto + siguiente.start()
        if siguiente else len(contenido_revision)
    )
    contexto = contenido_revision[inicio_contexto:fin_contexto]

    for interno in internos_ipoffice:
        if not re.match(r'^[0-9]+$', str(interno)):
            raise ValueError('Se detectó un interno inválido.')

        patron = (
            r'(?mi)^[ \t]*exten[ \t]*=>[ \t]*["\']?'
            + re.escape(str(interno))
            + r'["\']?[ \t]*,'
        )
        if re.search(patron, contexto):
            raise ValueError(
                'Ya existe una regla exacta para el interno {} fuera '
                'del bloque administrado. No se modificó el dialplan. '
                'Revisar primero esa regla.'.format(interno)
            )

    bloque = _vp_ubicaciones_generar_bloque(internos_ipoffice, salto, troncales_internos)

    if inicio == 1:
        p_inicio = contenido.find(_VP_UBICACIONES_MARCA_INICIO)
        p_fin = contenido.find(_VP_UBICACIONES_MARCA_FIN, p_inicio)
        p_fin += len(_VP_UBICACIONES_MARCA_FIN)
        return contenido[:p_inicio] + bloque + contenido[p_fin:]

    fin_cabecera = secciones[0].end()
    return (
        contenido[:fin_cabecera] + salto + salto +
        bloque + salto + salto + contenido[fin_cabecera:]
    )

def _vp_ubicaciones_restaurar_remoto(perfil, respaldo):
    comando = (
        'test -f "{}" && '
        'cp -p "{}" "{}" && '
        'asterisk -rx "dialplan reload"'
    ).format(
        respaldo,
        respaldo,
        _VP_UBICACIONES_ARCHIVO_REMOTO
    )
    salida, error, estado = ejecutar_comando_ssh(
        perfil, comando, timeout=30
    )
    if estado != 0:
        raise RuntimeError(
            'Falló la restauración remota. Salida: {} Error: {}'.format(
                salida, error
            )
        )


# VP OOH323 ENTRANTES WRAPPER V1
def _vp_ubicaciones_entrantes_incorporar(contenido, internos_ipoffice, troncales_internos=None):
    """Agrega solo extensiones asignadas a OOH323 en from-did-direct."""
    import re
    contenido = str(contenido or '')
    ini = '; BEGIN VOIP-PROVISIONER UBICACIONES ENTRANTES'
    fin = '; END VOIP-PROVISIONER UBICACIONES ENTRANTES'
    n_ini = contenido.count(ini)
    n_fin = contenido.count(fin)
    if n_ini != n_fin or n_ini > 1:
        raise ValueError('Marcadores de ubicaciones entrantes incompletos o duplicados; no se modifica el dialplan.')
    if n_ini == 1:
        a = contenido.find(ini)
        b = contenido.find(fin, a + len(ini))
        contenido = contenido[:a].rstrip() + '\n' + contenido[b + len(fin):].lstrip('\r\n')

    internos = []
    for valor in (internos_ipoffice or []):
        interno = str(valor or '').strip()
        if re.match(r'^[0-9]+$', interno) and interno not in internos:
            internos.append(interno)
    internos.sort(key=lambda x: (len(x), x))
    if not internos:
        return contenido.rstrip() + ('\n' if contenido.strip() else '')

    mapa = troncales_internos if isinstance(troncales_internos, dict) else {}
    rutas = []
    for interno in internos:
        tid = str(mapa.get(interno, '') or '').strip()
        if not tid:
            tid = '2'  # Compatibilidad con asignaciones antiguas "IP Office".
        if not re.match(r'^[0-9]+$', tid):
            raise ValueError('ID de troncal invalido para el interno ' + interno)
        rutas.append((interno, tid))

    # No pisar destinos directos preexistentes ajenos al bloque administrado.
    contexto = ''
    for linea in contenido.splitlines():
        mctx = re.match(r'^\s*\[([^]]+)\]', linea)
        if mctx:
            contexto = mctx.group(1).strip().lower()
            continue
        if contexto == 'from-did-direct':
            me = re.match(r'^\s*exten\s*=>\s*([^,\s]+)\s*,', linea, re.I)
            if me and me.group(1) in internos:
                raise ValueError('El interno ' + me.group(1) + ' ya tiene una regla directa en from-did-direct; no se sobrescribio.')

    lineas = [ini, '[from-did-direct]']
    for interno, tid in rutas:
        lineas.append('exten => {0},1,NoOp(VOIP Provisioner: DID entrante derivado por troncal IP Office)'.format(interno))
        lineas.append(' same => n,Macro(dialout-trunk,{0},${{EXTEN}},,off)'.format(tid))
        lineas.append(' same => n,Hangup()')
    lineas.append(fin)
    bloque = '\n'.join(lineas)
    return contenido.rstrip() + ('\n\n' if contenido.strip() else '') + bloque + '\n'


def _vp_ubicaciones_preparar_archivo_remoto(contenido, internos_ipoffice, troncales_internos=None):
    # Quitamos primero el bloque entrante propio para que la validacion original
    # de duplicados no lo confunda con reglas ajenas al Provisioner.
    contenido_limpio = _vp_ubicaciones_entrantes_incorporar(contenido, [], {})
    resultado = _vp_ubicaciones_preparar_archivo_remoto_base(contenido_limpio, internos_ipoffice, troncales_internos)
    if isinstance(resultado, tuple):
        partes = list(resultado)
        candidatos = []
        for i, valor in enumerate(partes):
            if isinstance(valor, str) and ('[from-internal-custom]' in valor or
                                           '; BEGIN VOIP-PROVISIONER UBICACIONES INTERNOS' in valor):
                candidatos.append(i)
        if len(candidatos) == 1:
            i = candidatos[0]
            partes[i] = _vp_ubicaciones_entrantes_incorporar(partes[i], internos_ipoffice, troncales_internos)
            return tuple(partes)
        # Si el preparador devolvio un error/resultado distinto, se conserva tal cual.
        return resultado
    if isinstance(resultado, str):
        return _vp_ubicaciones_entrantes_incorporar(resultado, internos_ipoffice, troncales_internos)
    return resultado


def _vp_ubicaciones_aplicar_dialplan(perfil, internos_ipoffice, troncales_internos=None):
    import re

    host = perfil.get('host', '127.0.0.1')
    puerto = int(perfil.get('puerto', 22))
    usuario = perfil.get('usuario', 'root')
    password = perfil.get('password', '')

    sftp = None
    transporte = None
    instalado = False
    respaldo = None
    temporal_remoto = None

    marca_tiempo = time.strftime('%Y%m%d_%H%M%S')
    respaldo = (
        '/etc/asterisk/extensions_custom.conf.'
        'bak_voip_provisioner_' + marca_tiempo
    )
    temporal_remoto = (
        '/etc/asterisk/.extensions_custom.conf.vp_tmp_' + marca_tiempo
    )

    try:
        sftp, transporte = conectar_sftp(
            host, puerto, usuario, password,
            ruta_remota='/etc/asterisk'
        )
        if sftp is None:
            raise RuntimeError(
                'No se pudo abrir SFTP en el servidor del perfil.'
            )

        ruta = _VP_UBICACIONES_ARCHIVO_REMOTO
        with sftp.file(ruta, 'rb') as remoto:
            contenido_bytes = remoto.read()
        if not isinstance(contenido_bytes, str):
            contenido = contenido_bytes.decode('utf-8')
        else:
            contenido = contenido_bytes

        nuevo_contenido = _vp_ubicaciones_preparar_archivo_remoto(
            contenido, internos_ipoffice, troncales_internos
        )

        with sftp.file(temporal_remoto, 'wb') as remoto:
            remoto.write(nuevo_contenido.encode('utf-8'))
            remoto.flush()

        comando_instalar = (
            'test -f "{ruta}" && '
            'cp -p "{ruta}" "{backup}" && '
            'chown --reference="{ruta}" "{tmp}" && '
            'chmod --reference="{ruta}" "{tmp}" && '
            'mv -f "{tmp}" "{ruta}"'
        ).format(
            ruta=ruta,
            backup=respaldo,
            tmp=temporal_remoto
        )

        salida, error, estado = ejecutar_comando_ssh(
            perfil, comando_instalar, timeout=30
        )
        if estado != 0:
            raise RuntimeError(
                'No se pudo instalar el archivo remoto. '
                'Salida: {} Error: {}'.format(salida, error)
            )

        instalado = True

        salida, error, estado = ejecutar_comando_ssh(
            perfil, 'asterisk -rx "dialplan reload"', timeout=30
        )
        if estado != 0:
            raise RuntimeError(
                'Falló la recarga del dialplan. '
                'Salida: {} Error: {}'.format(salida, error)
            )

        salida, error, estado = ejecutar_comando_ssh(
            perfil, 'asterisk -rx "dialplan show from-internal-custom"',
            timeout=30
        )
        if estado != 0:
            raise RuntimeError(
                'No se pudo verificar el contexto personalizado.'
            )

        for interno in internos_ipoffice:
            if "'{}' =>".format(interno) not in salida:
                raise RuntimeError(
                    'No se encontró el interno {} en el dialplan activo. '
                    'Se restaurará el archivo anterior.'.format(interno)
                )

        return respaldo

    except Exception:
        if instalado and respaldo:
            try:
                _vp_ubicaciones_restaurar_remoto(perfil, respaldo)
                logger.warning(
                    'Se restauró extensions_custom.conf desde %s', respaldo
                )
            except Exception:
                logger.exception(
                    'ATENCIÓN: no se pudo restaurar automáticamente '
                    'extensions_custom.conf.'
                )
        raise

    finally:
        if sftp is not None:
            try:
                sftp.close()
            except Exception:
                pass
        if transporte is not None:
            try:
                transporte.close()
            except Exception:
                pass
        if temporal_remoto:
            try:
                salida, error, estado = ejecutar_comando_ssh(
                    perfil,
                    'rm -f "{}"'.format(temporal_remoto),
                    timeout=10
                )
            except Exception:
                pass


@main_bp.route(
    '/api/monitoreo/internos/ubicaciones',
    methods=['GET']
)
def api_monitoreo_internos_ubicaciones():
    perfil_nombre = str(request.args.get('perfil_nombre', '')).strip()
    perfil = _vp_grupo_captura_perfil(perfil_nombre)
    if not perfil:
        return jsonify({'ok': False, 'mensaje': 'Perfil no encontrado.'}), 404
    try:
        with _VP_UBICACIONES_INTERNOS_LOCK:
            datos = _vp_ubicaciones_internos_leer()
        perfil_datos = datos.get(perfil_nombre, {})
        if not isinstance(perfil_datos, dict):
            perfil_datos = {}
        ubicaciones = {}
        for interno, ubicacion in perfil_datos.items():
            if (re.match(r'^[0-9]+$', str(interno)) and
                    (ubicacion in _VP_UBICACIONES_INTERNOS_VALIDAS or
                     _vp_ubicaciones_troncal_id(ubicacion))):
                ubicaciones[str(interno)] = ubicacion
        try:
            troncales = _vp_ubicaciones_troncales_leer(perfil)
        except Exception:
            logger.exception(
                'No se pudieron consultar las troncales OOH323 del perfil.'
            )
            troncales = []

        return jsonify({
            'ok': True,
            'ubicaciones': ubicaciones,
            'troncales': troncales
        })
    except Exception:
        logger.exception('Error leyendo ubicaciones administrativas de internos')
        return jsonify({
            'ok': False, 'mensaje': 'No se pudo leer el archivo de ubicaciones.'
        }), 500


@main_bp.route(
    '/api/monitoreo/internos/ubicaciones',
    methods=['POST']
)
def api_monitoreo_internos_ubicaciones_guardar():
    data = request.get_json(silent=True) or {}
    if not _vp_grupo_captura_password_valida(data):
        return jsonify({
            'ok': False, 'mensaje': 'Contraseña de administrador incorrecta.'
        }), 403

    perfil_nombre = str(data.get('perfil_nombre', '')).strip()
    perfil = _vp_grupo_captura_perfil(perfil_nombre)
    if not perfil:
        return jsonify({'ok': False, 'mensaje': 'Perfil no encontrado.'}), 404

    internos_recibidos = data.get('internos')
    if not isinstance(internos_recibidos, list) or not internos_recibidos:
        return jsonify({
            'ok': False, 'mensaje': 'No se recibieron internos para guardar.'
        }), 400
    if len(internos_recibidos) > 500:
        return jsonify({
            'ok': False, 'mensaje': 'El máximo permitido es 500 internos por operación.'
        }), 400

    internos = []
    for valor in internos_recibidos:
        interno = _vp_grupo_captura_interno(valor)
        if not interno:
            return jsonify({
                'ok': False, 'mensaje': 'La lista contiene un interno inválido.'
            }), 400
        if interno not in internos:
            internos.append(interno)

    ubicacion = str(data.get('ubicacion') or '').strip()
    if (ubicacion and
            ubicacion not in _VP_UBICACIONES_INTERNOS_VALIDAS and
            not _vp_ubicaciones_troncal_id(ubicacion)):
        return jsonify({
            'ok': False,
            'mensaje': 'Elegí Issabel o una troncal OOH323 disponible.'
        }), 400

    if ubicacion.startswith('IP Office:'):
        try:
            disponibles = _vp_ubicaciones_troncales_leer(perfil)
            ids = set(str(t['id']) for t in disponibles)
            if _vp_ubicaciones_troncal_id(ubicacion) not in ids:
                return jsonify({
                    'ok': False,
                    'mensaje': 'La troncal seleccionada no está disponible en Issabel.'
                }), 400
        except Exception as exc:
            logger.exception('Error validando troncal OOH323 seleccionada')
            return jsonify({
                'ok': False,
                'mensaje': 'No se pudo validar la troncal: ' + str(exc)
            }), 500

    respaldo_remoto = None
    try:
        with _VP_UBICACIONES_INTERNOS_LOCK:
            datos = _vp_ubicaciones_internos_leer()
            datos_originales = json.loads(json.dumps(datos))

            perfil_datos = datos.get(perfil_nombre, {})
            if not isinstance(perfil_datos, dict):
                perfil_datos = {}
            else:
                perfil_datos = dict(perfil_datos)

            for interno in internos:
                if ubicacion:
                    perfil_datos[interno] = ubicacion
                else:
                    perfil_datos.pop(interno, None)

            if perfil_datos:
                datos[perfil_nombre] = perfil_datos
            else:
                datos.pop(perfil_nombre, None)

            troncales_internos = {}
            for interno, destino in perfil_datos.items():
                if re.match(r'^[0-9]+$', str(interno)):
                    tid = _vp_ubicaciones_troncal_id(destino)
                    if tid:
                        troncales_internos[str(interno)] = tid

            ubicaciones_ipoffice = sorted(
                troncales_internos.keys(),
                key=lambda valor: int(valor)
            )

            respaldo_remoto = _vp_ubicaciones_aplicar_dialplan(
                perfil, ubicaciones_ipoffice, troncales_internos
            )

            try:
                _vp_ubicaciones_internos_guardar(datos)
            except Exception:
                logger.exception(
                    'No se pudo guardar el JSON; se intentará revertir '
                    'el dialplan remoto.'
                )
                try:
                    _vp_ubicaciones_restaurar_remoto(
                        perfil, respaldo_remoto
                    )
                except Exception:
                    logger.exception(
                        'ATENCIÓN: falló la reversión remota posterior '
                        'al error de escritura del JSON.'
                    )
                _vp_ubicaciones_internos_guardar(datos_originales)
                raise

        return jsonify({
            'ok': True,
            'mensaje': 'Ubicación guardada y dialplan actualizado.',
            'ubicacion': ubicacion,
            'internos': internos
        })

    except Exception as exc:
        logger.exception(
            'Error guardando ubicaciones y actualizando el dialplan'
        )
        return jsonify({
            'ok': False,
            'mensaje': 'No se pudo aplicar la ubicación: {}'.format(exc)
        }), 500


@main_bp.route(
    '/api/monitoreo/internos/grupos-captura',
    methods=['GET']
)
def api_monitoreo_internos_grupos_captura():
    """
    Consulta los grupos de captura reales de Asterisk.

    chan_sip:
      namedcallgroup
      namedpickupgroup

    PJSIP:
      callgroup
      pickupgroup
    """
    perfil_nombre = str(
        request.args.get('perfil_nombre', '')
    ).strip()

    perfil = _vp_grupo_captura_perfil(
        perfil_nombre
    )

    if not perfil:
        return jsonify({
            'ok': False,
            'mensaje': 'Perfil no encontrado.'
        }), 404

    sql = (
        "SELECT id, keyword, data "
        "FROM asterisk.sip "
        "WHERE keyword IN "
        "('namedcallgroup','namedpickupgroup',"
        "'named_call_group','named_pickup_group',"
        "'callgroup','pickupgroup',"
        "'call_group','pickup_group') "
        "AND data <> '' "
        "ORDER BY data, id;"
    )

    out, err, rc = _vp_grupo_captura_mysql(
        perfil,
        sql
    )

    if rc != 0:
        return jsonify({
            'ok': False,
            'mensaje': (
                err or out or
                'No se pudieron consultar los grupos de captura.'
            )
        }), 500

    grupos = {}
    nombres = set()
    uso = {}

    # Prioridad para obtener el grupo visible:
    #
    # 1) pickup del formato nativo
    # 2) call del formato nativo
    #
    # Se incluyen ambos formatos porque el mismo servidor
    # puede contener chan_sip y PJSIP.

    prioridad = {
        'namedpickupgroup': 100,
        'named_pickup_group': 100,
        'pickupgroup': 100,
        'pickup_group': 100,

        'namedcallgroup': 50,
        'named_call_group': 50,
        'callgroup': 50,
        'call_group': 50
    }

    prioridad_actual = {}

    for linea in str(out or '').splitlines():

        partes = linea.split('\t')

        if len(partes) < 3:
            partes = re.split(
                r'\s+',
                linea.strip(),
                maxsplit=2
            )

        if len(partes) < 3:
            continue

        interno = str(partes[0]).strip()
        keyword = str(partes[1]).strip()
        data = str(partes[2]).strip()

        if not interno or not data:
            continue

        if keyword not in prioridad:
            continue

        valores = [
            x.strip()
            for x in data.split(',')
            if x.strip()
        ]

        if not valores:
            continue

        for grupo in valores:
            nombres.add(grupo)
            uso.setdefault(grupo, set()).add(interno)

        p = prioridad.get(keyword, 0)
        p_actual = prioridad_actual.get(interno, -1)

        if p >= p_actual:
            grupos[interno] = valores[0]
            prioridad_actual[interno] = p

    # ========================================================
    # PERSISTENCIA POR SERVIDOR
    #
    # Los grupos viven en el Issabel remoto.
    # Si grupos.json no existe, el helper lo crea.
    #
    # Los grupos ya existentes en Asterisk se incorporan
    # automáticamente al archivo.
    # ========================================================

    grupos_json = _vp_grupos_catalogo(
        perfil
    )

    # Si no se pudo leer el catalogo, no se reescribe
    # (evita borrar grupos vacios ante un fallo transitorio de SSH).
    catalogo_legible = grupos_json is not None

    if grupos_json is None:
        grupos_json = []

    for grupo_json in grupos_json:

        nombre_json = str(
            grupo_json or ""
        ).strip()

        if nombre_json:
            nombres.add(
                nombre_json
            )

    # Guardamos la unión de:
    #
    #   Asterisk + grupos.json
    #
    # De esta forma el archivo se va construyendo
    # automáticamente aunque originalmente no exista.

    if catalogo_legible and set(grupos_json) != set(nombres):
        _vp_grupos_json_remoto_guardar(
            perfil,
            sorted(
                nombres,
                key=lambda x: x.lower()
            )
        )

    nombres_ordenados = sorted(nombres, key=lambda x: x.lower())
    numeros_existentes = _vp_grupos_numeros_leer(perfil)
    numeros = {}

    if numeros_existentes is not None:
        numeros = _vp_grupos_numeros_normalizar(
            nombres_ordenados, numeros_existentes
        )
        if numeros != numeros_existentes:
            if not _vp_grupos_json_remoto_guardar(
                perfil, nombres_ordenados, numeros
            ):
                logger.warning("No se pudo persistir la numeracion de grupos.")

    return jsonify({
        'ok': True,
        'grupos': grupos,
        'uso': {k: len(v) for k, v in uso.items()},
        'nombres': nombres_ordenados,
        'numeros': numeros
    })

@main_bp.route(
    '/api/monitoreo/interno/grupo-captura',
    methods=['POST']
)
def api_monitoreo_interno_grupo_captura():
    """
    Asigna un grupo de captura REAL de Asterisk.

    chan_sip:
      namedcallgroup
      namedpickupgroup

    PJSIP:
      callgroup
      pickupgroup

    La tecnologia se determina por el interno real.
    """
    data = request.get_json(silent=True) or {}

    if not _vp_grupo_captura_password_valida(data):
        return jsonify({
            'ok': False,
            'mensaje': 'Contraseña de administrador incorrecta.'
        }), 403

    perfil_nombre = str(
        data.get('perfil_nombre', '')
    ).strip()

    perfil = _vp_grupo_captura_perfil(
        perfil_nombre
    )

    if not perfil:
        return jsonify({
            'ok': False,
            'mensaje': 'Perfil no encontrado.'
        }), 404

    interno = _vp_grupo_captura_interno(
        data.get('interno')
    )

    if not interno:
        return jsonify({
            'ok': False,
            'mensaje': 'Interno inválido.'
        }), 400

    grupo = _vp_grupo_captura_nombre(
        data.get('grupo')
    )

    if str(data.get('grupo') or '').strip() and not grupo:
        return jsonify({
            'ok': False,
            'mensaje': (
                'Nombre de grupo inválido. Usá letras, números, '
                'espacios, punto, guion o guion bajo (máximo 64).'
            )
        }), 400

    grupo_sql = str(grupo or '').replace("'", "''")

    # ========================================================
    # DETERMINAR TECNOLOGIA REAL DEL INTERNO
    # Se consulta devices directamente para soportar tanto SIP como PJSIP.
    interno_sql = str(interno or "").strip().replace("'", "''")

    sql_tecnologia = (
        "SELECT tech "
        "FROM asterisk.devices "
        "WHERE id='{}' "
        "LIMIT 1;"
    ).format(interno_sql)

    out_tech, err_tech, rc_tech = _vp_grupo_captura_mysql(
        perfil, sql_tecnologia
    )

    tecnologia = ""
    if rc_tech == 0 and out_tech:
        tecnologia = str(out_tech).strip().lower().splitlines()[0].strip()

    if tecnologia == "chan_sip":
        tecnologia = "sip"

    if tecnologia not in ("sip", "pjsip"):
        logger.warning(
            "Grupo de captura: no se pudo determinar tecnologia para interno %s",
            interno
        )
        return jsonify({"ok": False, "mensaje": "No se pudo determinar la tecnología del interno (no figura en asterisk.devices)."}), 404  # VP GRUPOS MGR V2

    # ========================================================
    # CLAVES SOPORTADAS
    #
    # Eliminamos ambas familias primero.
    #
    # Esto evita que quede una configuracion anterior incorrecta
    # si un interno cambio de tecnologia.
    # ========================================================

    claves = (
        "'namedcallgroup',"
        "'namedpickupgroup',"
        "'named_call_group',"
        "'named_pickup_group',"
        "'callgroup',"
        "'pickupgroup',"
        "'call_group',"
        "'pickup_group'"
    )

    sql = (
        "DELETE FROM asterisk.sip "
        "WHERE id='{}' "
        "AND keyword IN ({}); "
    ).format(
        str(interno),
        claves
    )

    # ========================================================
    # ELIMINAR GRUPO
    # ========================================================

    if grupo:

        if tecnologia == 'pjsip':
            # PJSIP requiere nombres de opciones con guiones bajos.
            keyword_call = 'named_call_group'
            keyword_pickup = 'named_pickup_group'
        else:
            # chan_sip utiliza las claves sin guiones bajos.
            keyword_call = 'namedcallgroup'
            keyword_pickup = 'namedpickupgroup'

        sql += (
            "INSERT INTO asterisk.sip "
            "(id, keyword, data, flags) "
            "VALUES "
            "('{}','{}','{}',0); "
        ).format(
            str(interno),
            keyword_call,
            grupo_sql
        )

        sql += (
            "INSERT INTO asterisk.sip "
            "(id, keyword, data, flags) "
            "VALUES "
            "('{}','{}','{}',0);"
        ).format(
            str(interno),
            keyword_pickup,
            grupo_sql
        )


    # ========================================================
    # GUARDAR EN MYSQL
    # ========================================================

    out, err, rc = _vp_grupo_captura_mysql(
        perfil,
        sql
    )

    if rc != 0:
        return jsonify({
            'ok': False,
            'mensaje': (
                err or out or
                'No se pudo guardar el grupo en Asterisk.'
            )
        }), 500

    # ========================================================
    # retrieve_conf
    # ========================================================

    r_out, r_err, r_rc = ejecutar_comando_ssh(
        perfil,
        "/var/lib/asterisk/bin/retrieve_conf 2>&1"
    )

    if r_rc != 0:
        return jsonify({
            'ok': False,
            'mensaje': (
                'El grupo se guardo en MySQL/Asterisk, '
                'pero retrieve_conf fallo: ' +
                str(r_err or r_out or '')[:500]
            )
        }), 500

    # ========================================================
    # RELOAD
    # ========================================================

    if tecnologia == 'pjsip':
        reload_cmd = "asterisk -rx 'pjsip reload' 2>&1"
    else:
        reload_cmd = "asterisk -rx 'sip reload' 2>&1"

    rr_out, rr_err, rr_rc = ejecutar_comando_ssh(
        perfil, reload_cmd
    )

    if rr_rc != 0:
        return jsonify({
            'ok': False,
            'mensaje': (
                'El grupo fue guardado y retrieve_conf termino, '
                'pero fallo la recarga: ' +
                str(rr_err or rr_out or '')[:500]
            )
        }), 500

    # ========================================================
    # PERSISTENCIA DEL NUEVO GRUPO EN EL ISSABEL REMOTO
    #
    # El grupo ya quedó aplicado en Asterisk.
    # Ahora lo incorporamos al grupos.json del mismo servidor.
    # ========================================================

    if grupo:

        json_ok = _vp_grupos_json_remoto_actualizar(
            perfil,
            [grupo]
        )

        if not json_ok:

            logger.warning(
                "El grupo %s quedó aplicado en Asterisk, "
                "pero no pudo persistirse en grupos.json.",
                grupo
            )

    return jsonify({
        'ok': True,
        'interno': interno,
        'grupo': grupo or '',
        'tecnologia': tecnologia
    })


# === VP GRUPOS MGR V1 - GESTION DEL CATALOGO DE GRUPOS DE CAPTURA ===
#
# Alta, renombrado y baja de grupos de captura por servidor.
# El catalogo vive en cada Issabel (grupos.json). El renombrado y la
# baja tambien se propagan a asterisk.sip (internos que usan el grupo).

_VP_GRUPOS_KEYWORDS = (
    'namedcallgroup', 'namedpickupgroup',
    'named_call_group', 'named_pickup_group',
    'callgroup', 'pickupgroup',
    'call_group', 'pickup_group'
)


def _vp_grupo_sql_str(valor):
    return str(valor).replace('\\', '\\\\').replace("'", "''")


def _vp_grupos_catalogo(perfil):
    """
    Devuelve la lista de grupos de grupos.json del servidor,
    o None si NO se pudo leer (para no pisar el archivo ante un
    fallo transitorio de SSH).
    """
    import json
    import shlex

    cmd = "cat " + shlex.quote(_VP_GRUPOS_JSON_REMOTO) + " 2>/dev/null"

    try:
        out, err, rc = ejecutar_comando_ssh(perfil, cmd)

        if rc != 0:
            # Si el archivo no existe, el helper lo crea.
            _vp_grupos_json_remoto_leer(perfil)
            out, err, rc = ejecutar_comando_ssh(perfil, cmd)

            if rc != 0:
                return None

        datos = json.loads(str(out or "").strip() or "{}")

        if not isinstance(datos, dict):
            return None

        lista = datos.get("grupos", [])

        if not isinstance(lista, list):
            return None

        resultado = []

        for g in lista:
            nombre = str(g or "").strip()

            if nombre and nombre not in resultado:
                resultado.append(nombre)

        return resultado

    except Exception:
        logger.exception("Error leyendo catalogo de grupos")
        return None


def _vp_grupos_filas_asterisk(perfil):
    """
    Devuelve [(id, keyword, data), ...] de asterisk.sip con grupos
    de captura, o None si falla la consulta.
    """
    claves = ",".join("'%s'" % k for k in _VP_GRUPOS_KEYWORDS)

    sql = (
        "SELECT id, keyword, data FROM asterisk.sip "
        "WHERE keyword IN ({}) AND data <> '';"
    ).format(claves)

    out, err, rc = _vp_grupo_captura_mysql(perfil, sql)

    if rc != 0:
        return None

    filas = []

    for linea in str(out or "").splitlines():
        partes = linea.split("\t")

        if len(partes) < 3:
            continue

        filas.append((
            partes[0].strip(),
            partes[1].strip(),
            partes[2].strip()
        ))

    return filas


def _vp_grupos_transformar(filas, nombre, nuevo=None):
    """
    Calcula las sentencias SQL para quitar (nuevo=None) o renombrar
    (nuevo=str) un grupo en las filas de asterisk.sip.
    Respeta listas separadas por coma.
    Devuelve (sentencias, internos_afectados).
    """
    sentencias = []
    internos = set()

    for interno, keyword, data in filas:
        valores = [v.strip() for v in data.split(",") if v.strip()]

        if nombre not in valores:
            continue

        if nuevo is None:
            restantes = [v for v in valores if v != nombre]
        else:
            restantes = []

            for v in valores:
                v2 = nuevo if v == nombre else v

                if v2 not in restantes:
                    restantes.append(v2)

        internos.add(interno)

        donde = "id='{}' AND keyword='{}'".format(
            _vp_grupo_sql_str(interno),
            _vp_grupo_sql_str(keyword)
        )

        if restantes:
            sentencias.append(
                "UPDATE asterisk.sip SET data='{}' WHERE {};".format(
                    _vp_grupo_sql_str(",".join(restantes)),
                    donde
                )
            )
        else:
            sentencias.append(
                "DELETE FROM asterisk.sip WHERE {};".format(donde)
            )

    return sentencias, internos


def _vp_grupos_ejecutar_sql(perfil, sentencias, lote=100):
    for i in range(0, len(sentencias), lote):
        out, err, rc = _vp_grupo_captura_mysql(
            perfil,
            " ".join(sentencias[i:i + lote])
        )

        if rc != 0:
            return False, str(err or out or "Error de MySQL")[:500]

    return True, ""


def _vp_grupos_aplicar_asterisk(perfil):
    # Misma secuencia que usa la asignacion de grupo a un interno.
    r_out, r_err, r_rc = ejecutar_comando_ssh(
        perfil,
        "/var/lib/asterisk/bin/retrieve_conf 2>&1"
    )

    if r_rc != 0:
        return False, (
            "Los cambios se guardaron en MySQL, pero retrieve_conf fallo: "
            + str(r_err or r_out or "")[:500]
        )

    ejecutar_comando_ssh(
        perfil,
        "asterisk -rx 'pjsip reload' >/dev/null 2>&1 || true; "
        "asterisk -rx 'sip reload' >/dev/null 2>&1 || true"
    )

    return True, ""


@main_bp.route(
    '/api/monitoreo/grupos-captura/gestionar',
    methods=['POST']
)
def api_monitoreo_grupos_captura_gestionar():
    """
    Crear, renombrar o eliminar un grupo de captura del servidor.

    Body: perfil_nombre, accion (crear|renombrar|eliminar),
          grupo, nuevo (solo renombrar), password
    """
    data = request.get_json(silent=True) or {}

    if not _vp_grupo_captura_password_valida(data):
        return jsonify({
            'ok': False,
            'mensaje': 'Contraseña de administrador incorrecta.'
        }), 403

    perfil = _vp_grupo_captura_perfil(
        str(data.get('perfil_nombre', '')).strip()
    )

    if not perfil:
        return jsonify({
            'ok': False,
            'mensaje': 'Perfil no encontrado.'
        }), 404

    accion = str(data.get('accion', '')).strip().lower()

    if accion not in ('crear', 'renombrar', 'eliminar'):
        return jsonify({
            'ok': False,
            'mensaje': 'Acción inválida.'
        }), 400

    msg_nombre = (
        'Nombre de grupo inválido. Usá letras, números, espacios, '
        'punto, guion o guion bajo (máximo 64 caracteres).'
    )

    nombre = _vp_grupo_captura_nombre(data.get('grupo'))

    if not nombre:
        return jsonify({'ok': False, 'mensaje': msg_nombre}), 400

    catalogo = _vp_grupos_catalogo(perfil)

    if catalogo is None:
        return jsonify({
            'ok': False,
            'mensaje': 'No se pudo leer el catálogo de grupos del servidor.'
        }), 500

    numeros_existentes = _vp_grupos_numeros_leer(perfil)
    if numeros_existentes is None:
        return jsonify({
            'ok': False,
            'mensaje': 'No se pudo leer la numeracion actual de los grupos.'
        }), 500

    numeros_catalogo = _vp_grupos_numeros_normalizar(
        catalogo, numeros_existentes
    )

    # ---------------- CREAR ----------------
    if accion == 'crear':

        if any(g.lower() == nombre.lower() for g in catalogo):
            return jsonify({
                'ok': False,
                'mensaje': 'El grupo "{}" ya existe en este servidor.'.format(nombre)
            }), 409

        if not _vp_grupos_json_remoto_guardar(
            perfil, catalogo + [nombre], numeros_catalogo
        ):
            return jsonify({
                'ok': False,
                'mensaje': 'No se pudo guardar el grupo en el servidor.'
            }), 500

        return jsonify({
            'ok': True,
            'accion': accion,
            'grupo': nombre,
            'mensaje': 'Grupo "{}" creado en este servidor.'.format(nombre)
        })

    # ------------- RENOMBRAR / ELIMINAR -------------
    nuevo = None

    if accion == 'renombrar':
        nuevo = _vp_grupo_captura_nombre(data.get('nuevo'))

        if not nuevo:
            return jsonify({'ok': False, 'mensaje': msg_nombre}), 400

        if nuevo == nombre:
            return jsonify({
                'ok': False,
                'mensaje': 'El nuevo nombre es igual al actual.'
            }), 400

        if any(
            g.lower() == nuevo.lower() and g.lower() != nombre.lower()
            for g in catalogo
        ):
            return jsonify({
                'ok': False,
                'mensaje': 'Ya existe un grupo llamado "{}".'.format(nuevo)
            }), 409

    filas = _vp_grupos_filas_asterisk(perfil)

    if filas is None:
        return jsonify({
            'ok': False,
            'mensaje': 'No se pudieron consultar los grupos en Asterisk.'
        }), 500

    sentencias, internos = _vp_grupos_transformar(filas, nombre, nuevo)

    if nombre not in catalogo and not sentencias:
        return jsonify({
            'ok': False,
            'mensaje': 'El grupo "{}" no existe en este servidor.'.format(nombre)
        }), 404

    if sentencias:
        ok_sql, msg_sql = _vp_grupos_ejecutar_sql(perfil, sentencias)

        if not ok_sql:
            return jsonify({
                'ok': False,
                'mensaje': 'No se pudo actualizar Asterisk: ' + msg_sql
            }), 500

    if accion == 'eliminar':
        catalogo_nuevo = [g for g in catalogo if g != nombre]
    else:
        catalogo_nuevo = []

        for g in catalogo:
            g2 = nuevo if g == nombre else g

            if g2 not in catalogo_nuevo:
                catalogo_nuevo.append(g2)

        if nuevo not in catalogo_nuevo:
            catalogo_nuevo.append(nuevo)

    if accion == 'eliminar':
        numeros_nuevos = dict(numeros_catalogo)
        numeros_nuevos.pop(nombre, None)
    else:
        numeros_nuevos = dict(numeros_catalogo)
        numero_anterior = numeros_nuevos.pop(nombre, None)
        if numero_anterior is not None:
            numeros_nuevos[nuevo] = numero_anterior

    if not _vp_grupos_json_remoto_guardar(
        perfil, catalogo_nuevo, numeros_nuevos
    ):
        return jsonify({
            'ok': False,
            'mensaje': (
                'Asterisk se actualizó, pero no se pudo guardar el '
                'catálogo de grupos del servidor.'
            )
        }), 500

    if sentencias:
        ok_ap, msg_ap = _vp_grupos_aplicar_asterisk(perfil)

        if not ok_ap:
            return jsonify({'ok': False, 'mensaje': msg_ap}), 500

    cant = len(internos)
    detalle = (
        'sin internos asignados' if cant == 0
        else '1 interno actualizado' if cant == 1
        else '{} internos actualizados'.format(cant)
    )

    if accion == 'eliminar':
        mensaje = 'Grupo "{}" eliminado ({}).'.format(nombre, detalle)
    else:
        mensaje = 'Grupo "{}" renombrado a "{}" ({}).'.format(
            nombre, nuevo, detalle
        )

    return jsonify({
        'ok': True,
        'accion': accion,
        'grupo': nombre,
        'nuevo': nuevo or '',
        'internos_afectados': cant,
        'mensaje': mensaje
    })

# === FIN VP GRUPOS MGR V1 ===


# VP_RUTEO_EXTERNO_API_V1
@main_bp.route('/api/monitoreo/ruteo-externo', methods=['GET'])
def api_monitoreo_ruteo_externo():
    """Consulta de solo lectura de rutas, destinos y horarios de Issabel."""
    try:
        perfil_nombre = str(request.args.get('perfil_nombre') or '').strip()
        if not perfil_nombre:
            return jsonify({'ok': False, 'mensaje': 'Seleccioná un servidor.'}), 400

        perfil = buscar_perfil(perfil_nombre, cargar_perfiles())
        if not perfil:
            return jsonify({'ok': False, 'mensaje': 'No se encontró el servidor seleccionado.'}), 404

        def _vp_rut_query(sql):
            ok, salida, error = _vp_trunk_run_sql(perfil, sql, timeout=35)
            if not ok:
                detalle = str(error or salida or 'Error consultando Issabel.')
                raise RuntimeError(detalle[:700])
            return [x for x in str(salida or '').splitlines() if x.strip()]

        # Descubrir primero las tablas y columnas disponibles: Issabel/FreePBX
        # puede variar según su versión. Todas las consultas posteriores son SELECT.
        esquema = _vp_rut_query(
            "SELECT table_name,column_name "
            "FROM information_schema.columns "
            "WHERE table_schema=DATABASE() AND table_name IN ("
            "'outbound_routes','outbound_route_sequence','outbound_route_patterns',"
            "'outbound_route_trunks','trunks','incoming','timeconditions',"
            "'timegroups_details','timegroups_groups') "
            "ORDER BY table_name,ordinal_position;"
        )
        columnas = {}
        for linea in esquema:
            partes = linea.split('\t')
            if len(partes) >= 2:
                columnas.setdefault(partes[0].lower(), set()).add(partes[1].lower())

        def _vp_rut_has(tabla, *nombres):
            disponibles = columnas.get(tabla.lower(), set())
            return all(str(n).lower() in disponibles for n in nombres)

        def _vp_rut_valor(fila, nombres, indice):
            resultado = {}
            for i, nombre in enumerate(nombres):
                resultado[nombre] = fila[i] if i < len(fila) else ''
            return resultado

        salientes = []
        if _vp_rut_has('outbound_routes', 'route_id', 'name'):
            expr = [
                "COALESCE(r.route_id,'')",
                "COALESCE(r.name,'')",
                "COALESCE(r.outcid,'')" if _vp_rut_has('outbound_routes', 'outcid') else "''",
                "COALESCE(r.time_group_id,'0')" if _vp_rut_has('outbound_routes', 'time_group_id') else "'0'",
            ]
            base_rows = _vp_rut_query(
                'SELECT ' + ','.join(expr) + ' FROM outbound_routes r ORDER BY r.route_id;'
            )
            por_id = {}
            for linea in base_rows:
                v = linea.split('\t')
                if len(v) < 4:
                    continue
                rid = str(v[0]).strip()
                por_id[rid] = {
                    'route_id': rid, 'nombre': v[1], 'caller_id': v[2],
                    'time_group_id': v[3], 'orden': '', 'patrones': [], 'troncales': []
                }
            if por_id and _vp_rut_has('outbound_route_sequence', 'route_id', 'seq'):
                for linea in _vp_rut_query(
                    'SELECT route_id,seq FROM outbound_route_sequence ORDER BY seq,route_id;'
                ):
                    v = linea.split('\t')
                    if len(v) >= 2 and v[0] in por_id:
                        por_id[v[0]]['orden'] = v[1]

            if por_id and _vp_rut_has('outbound_route_patterns', 'route_id'):
                campos = [
                    ('match_pattern_prefix', 'prefijo'),
                    ('match_pattern_pass', 'patron'),
                    ('match_cid', 'cid'),
                    ('prepend_digits', 'anteponer'),
                ]
                select_pat = ['route_id']
                usados_pat = []
                for col, alias in campos:
                    if _vp_rut_has('outbound_route_patterns', col):
                        select_pat.append("COALESCE(`%s`,'')" % col)
                        usados_pat.append(alias)
                    else:
                        select_pat.append("''")
                        usados_pat.append(alias)
                for linea in _vp_rut_query(
                    'SELECT ' + ','.join(select_pat) +
                    ' FROM outbound_route_patterns ORDER BY route_id;'
                ):
                    v = linea.split('\t')
                    if not v or v[0] not in por_id:
                        continue
                    item = {'prefijo': '', 'patron': '', 'cid': '', 'anteponer': ''}
                    for ix, alias in enumerate(usados_pat, 1):
                        item[alias] = v[ix] if ix < len(v) else ''
                    if any(str(x or '').strip() for x in item.values()):
                        por_id[v[0]]['patrones'].append(item)

            if por_id and _vp_rut_has('outbound_route_trunks', 'route_id', 'trunk_id'):
                tiene_seq = _vp_rut_has('outbound_route_trunks', 'seq')
                tiene_trunks = _vp_rut_has('trunks', 'trunkid')
                if tiene_trunks:
                    nombre_expr = "COALESCE(t.name,'')" if _vp_rut_has('trunks', 'name') else "''"
                    tech_expr = "COALESCE(t.tech,'')" if _vp_rut_has('trunks', 'tech') else "''"
                    chan_expr = "COALESCE(t.channelid,'')" if _vp_rut_has('trunks', 'channelid') else "''"
                    join = ' LEFT JOIN trunks t ON t.trunkid=ot.trunk_id'
                else:
                    nombre_expr, tech_expr, chan_expr = "''", "''", "''"
                    join = ''
                seq_expr = "COALESCE(ot.seq,'')" if tiene_seq else "''"
                sql = (
                    'SELECT COALESCE(ot.route_id,\'\'),' + seq_expr +
                    ",COALESCE(ot.trunk_id,'')," + nombre_expr + ',' + tech_expr + ',' + chan_expr +
                    ' FROM outbound_route_trunks ot' + join +
                    ' ORDER BY ot.route_id' + (',ot.seq' if tiene_seq else ',ot.trunk_id') + ';'
                )
                for linea in _vp_rut_query(sql):
                    v = linea.split('\t')
                    if len(v) < 6 or v[0] not in por_id:
                        continue
                    por_id[v[0]]['troncales'].append({
                        'seq': v[1], 'trunk_id': v[2], 'nombre': v[3],
                        'tech': v[4], 'channelid': v[5]
                    })
            salientes = list(por_id.values())
            salientes.sort(key=lambda x: (int(x['orden']) if str(x['orden']).isdigit() else 999999,
                                          int(x['route_id']) if str(x['route_id']).isdigit() else 999999))

        entrantes = []
        if _vp_rut_has('incoming', 'destination'):
            in_campos = ['extension', 'cidnum', 'description', 'destination']
            disponibles = [c for c in in_campos if _vp_rut_has('incoming', c)]
            if 'destination' not in disponibles:
                disponibles.append('destination')
            select_in = []
            for c in disponibles:
                select_in.append("COALESCE(`%s`,'')" % c)
            orden_in = 'extension' if 'extension' in disponibles else 'destination'
            for linea in _vp_rut_query(
                'SELECT ' + ','.join(select_in) + ' FROM incoming ORDER BY `' + orden_in + '`;'
            ):
                v = linea.split('\t')
                fila = {c: (v[i] if i < len(v) else '') for i, c in enumerate(disponibles)}
                destino = str(fila.get('destination') or '').strip()
                trozos = destino.split(',')
                entrantes.append({
                    'did': str(fila.get('extension') or '').strip(),
                    'caller_id': str(fila.get('cidnum') or '').strip(),
                    'descripcion': str(fila.get('description') or '').strip(),
                    'destino': destino,
                    'contexto_destino': trozos[0].strip() if trozos else '',
                    'extension_destino': trozos[1].strip() if len(trozos) > 1 else '',
                    'prioridad_destino': trozos[2].strip() if len(trozos) > 2 else ''
                })

        condiciones = []
        if _vp_rut_has('timeconditions', 'timeconditions_id'):
            tabla = 'timeconditions'
            campos = [
                ('timeconditions_id', 'id'), ('displayname', 'nombre'),
                ('description', 'descripcion'), ('destmatch', 'destino_si'),
                ('destnomatch', 'destino_no'), ('truegoto', 'truegoto'),
                ('falsegoto', 'falsegoto'), ('time', 'horario'),
                ('mode', 'modo'), ('enabled', 'habilitada'), ('timezone', 'zona_horaria')
            ]
            select_tc, usados_tc = [], []
            for col, alias in campos:
                select_tc.append("COALESCE(`%s`,'')" % col if _vp_rut_has(tabla, col) else "''")
                usados_tc.append(alias)
            for linea in _vp_rut_query(
                'SELECT ' + ','.join(select_tc) + ' FROM timeconditions ORDER BY timeconditions_id;'
            ):
                v = linea.split('\t')
                item_tc = {alias: (v[i] if i < len(v) else '') for i, alias in enumerate(usados_tc)}
                if item_tc.get('truegoto'):
                    item_tc['destino_si'] = item_tc.get('truegoto')
                if item_tc.get('falsegoto'):
                    item_tc['destino_no'] = item_tc.get('falsegoto')
                item_tc['grupo_horario_id'] = item_tc.get('horario', '')
                condiciones.append(item_tc)

        grupos = []
        tabla_grupos = 'timegroups_groups'
        clave_grupo = 'id' if _vp_rut_has(tabla_grupos, 'id') else ''
        desc_grupo = 'description' if _vp_rut_has(tabla_grupos, 'description') else ''
        if clave_grupo:
            sel = "COALESCE(`id`,'')" + (",COALESCE(`description`,'')" if desc_grupo else ",''")
            for linea in _vp_rut_query('SELECT ' + sel + ' FROM timegroups_groups ORDER BY id;'):
                v = linea.split('\t')
                grupos.append({'id': v[0] if v else '', 'descripcion': v[1] if len(v) > 1 else ''})

        franjas = []
        if _vp_rut_has('timegroups_details', 'timegroupid'):
            tabla = 'timegroups_details'
            campos = [
                ('timegroupid', 'grupo_id'), ('description', 'descripcion'),
                ('time', 'horario'), ('wday', 'dias_semana'),
                ('mday', 'dias_mes'), ('month', 'meses')
            ]
            select_tg, usados_tg = [], []
            for col, alias in campos:
                select_tg.append("COALESCE(`%s`,'')" % col if _vp_rut_has(tabla, col) else "''")
                usados_tg.append(alias)
            for linea in _vp_rut_query(
                'SELECT ' + ','.join(select_tg) + ' FROM timegroups_details ORDER BY timegroupid;'
            ):
                v = linea.split('\t')
                item_tg = {alias: (v[i] if i < len(v) else '') for i, alias in enumerate(usados_tg)}
                # En esquemas modernos, timegroups_details.time guarda hora|días|día-mes|mes.
                hora_compuesta = str(item_tg.get('horario') or '')
                if '|' in hora_compuesta:
                    partes_hora = hora_compuesta.split('|')
                    item_tg['horario'] = partes_hora[0] if len(partes_hora) > 0 else ''
                    if not item_tg.get('dias_semana') and len(partes_hora) > 1:
                        item_tg['dias_semana'] = partes_hora[1]
                    if not item_tg.get('dias_mes') and len(partes_hora) > 2:
                        item_tg['dias_mes'] = partes_hora[2]
                    if not item_tg.get('meses') and len(partes_hora) > 3:
                        item_tg['meses'] = partes_hora[3]
                franjas.append(item_tg)

        troncales_catalogo = []
        if _vp_rut_has('trunks', 'trunkid'):
            nombre_expr = "COALESCE(name,'')" if _vp_rut_has('trunks', 'name') else "''"
            tech_expr = "COALESCE(tech,'')" if _vp_rut_has('trunks', 'tech') else "''"
            chan_expr = "COALESCE(channelid,'')" if _vp_rut_has('trunks', 'channelid') else "''"
            for linea in _vp_rut_query("SELECT COALESCE(trunkid,'')," + nombre_expr + ',' + tech_expr + ',' + chan_expr + ' FROM trunks ORDER BY name,trunkid;'):
                v = linea.split('\t')
                if len(v) >= 4 and str(v[0]).strip().isdigit():
                    troncales_catalogo.append({'trunkid': v[0], 'nombre': v[1], 'tech': v[2], 'channelid': v[3]})

        tablas = [t for t in ('incoming', 'outbound_routes', 'outbound_route_patterns',
                              'outbound_route_trunks', 'timeconditions',
                              'timegroups_groups', 'timegroups_details') if t in columnas]
        return jsonify({
            'ok': True,
            'perfil': perfil_nombre,
            'entrantes': entrantes,
            'salientes': salientes,
            'condiciones_horarias': condiciones,
            'grupos_horarios': grupos,
            'franjas_horarias': franjas,
            'troncales_catalogo': troncales_catalogo,
            'tablas_detectadas': tablas,
            'solo_lectura': False
        })
    except Exception as e:
        logger.exception('[RUTEO EXTERNO] Error consultando configuración')
        return jsonify({'ok': False, 'mensaje': 'No se pudo consultar el ruteo externo: ' + str(e)[:500]}), 500


# VP_RUTEO_EXTERNO_EDITAR_API_V2
@main_bp.route('/api/monitoreo/ruteo-externo/editar', methods=['POST'])
def api_monitoreo_ruteo_externo_editar():
    # Edita objetos existentes de rutas Issabel con autenticacion, backup y validacion.
    backup_path = ''
    try:
        data = request.get_json(silent=True) or {}
        if not _vp_trunk_admin_password(data):
            return jsonify({'ok': False, 'mensaje': 'Contraseña de administrador incorrecta.'}), 401

        perfil_nombre = str(data.get('perfil_nombre') or '').strip()
        tipo = str(data.get('tipo') or '').strip().lower()
        perfil = buscar_perfil(perfil_nombre, cargar_perfiles()) if perfil_nombre else None
        if not perfil:
            return jsonify({'ok': False, 'mensaje': 'Seleccioná un servidor válido.'}), 400
        tipos_permitidos = ('entrada', 'salida', 'condicion', 'grupo')
        if tipo not in tipos_permitidos:
            return jsonify({'ok': False, 'mensaje': 'Tipo de configuración no permitido.'}), 400

        def _vp_re_sql(sql, timeout=45):
            ok, salida, error = _vp_trunk_run_sql(perfil, sql, timeout=timeout)
            salida = str(salida or '')
            if (not ok) or re.search(r'(^|\n)ERROR\s+\d+|Unknown column|doesn.t exist|SQL syntax', salida, re.I):
                raise RuntimeError(str(error or salida or 'Error ejecutando consulta en Issabel.')[:800])
            return [x for x in salida.splitlines() if x.strip()]

        def _vp_re_esquema():
            lineas = _vp_re_sql(
                "SELECT table_name,column_name FROM information_schema.columns "
                "WHERE table_schema=DATABASE() AND table_name IN "
                "('incoming','outbound_routes','outbound_route_sequence','outbound_route_patterns',"
                "'outbound_route_trunks','trunks','timeconditions','timegroups_groups','timegroups_details') "
                "ORDER BY table_name,ordinal_position;"
            )
            salida = {}
            for linea in lineas:
                p = linea.split('\t')
                if len(p) >= 2:
                    salida.setdefault(p[0].lower(), set()).add(p[1].lower())
            return salida

        esquema = _vp_re_esquema()
        def _vp_re_has(tabla, *cols):
            disponibles = esquema.get(tabla.lower(), set())
            return all(str(c).lower() in disponibles for c in cols)

        def _vp_re_q(v):
            return _vp_trunk_sql_q(str(v if v is not None else ''))

        def _vp_re_txt(v, nombre, maxlen=160, allow_empty=True, patron=None):
            v = str(v if v is not None else '').strip()
            if (not allow_empty) and not v:
                raise ValueError('El campo ' + nombre + ' es obligatorio.')
            if len(v) > maxlen or any(ord(c) < 32 for c in v):
                raise ValueError('El campo ' + nombre + ' tiene una longitud o caracteres no permitidos.')
            if patron and v and not re.match(patron, v):
                raise ValueError('El formato del campo ' + nombre + ' no es válido.')
            return v

        def _vp_re_backup(tablas):
            permitidas = set(['incoming','outbound_routes','outbound_route_sequence',
                              'outbound_route_patterns','outbound_route_trunks',
                              'timeconditions','timegroups_groups','timegroups_details'])
            tablas = sorted(set(tablas))
            if not tablas or any(t not in permitidas or t not in esquema for t in tablas):
                raise RuntimeError('No se pudo comprobar la existencia de todas las tablas a respaldar; no se modificó nada.')
            sello = time.strftime('%Y%m%d_%H%M%S')
            destino = '/var/backups/voip-provisioner/ruteo-externo/ruteo_' + sello + '.sql'
            lista = ' '.join(tablas)
            comando = r'''set -eu
AMPDBUSER="$(grep '^AMPDBUSER=' /etc/amportal.conf | head -1 | cut -d= -f2-)"
AMPDBPASS="$(grep '^AMPDBPASS=' /etc/amportal.conf | head -1 | cut -d= -f2-)"
AMPDBHOST="$(grep '^AMPDBHOST=' /etc/amportal.conf | head -1 | cut -d= -f2-)"
AMPDBNAME="$(grep '^AMPDBNAME=' /etc/amportal.conf | head -1 | cut -d= -f2-)"
[ -n "$AMPDBUSER" ] && [ -n "$AMPDBPASS" ] && [ -n "$AMPDBNAME" ] || { echo 'No se pudieron leer credenciales AMPDB'; exit 20; }
command -v mysqldump >/dev/null 2>&1 || { echo 'No se encuentra mysqldump'; exit 21; }
mkdir -p /var/backups/voip-provisioner/ruteo-externo
chmod 700 /var/backups/voip-provisioner /var/backups/voip-provisioner/ruteo-externo
TMP_CNF="$(mktemp)"
trap 'rm -f "$TMP_CNF"' EXIT
printf '[client]\nuser=%s\npassword=%s\n' "$AMPDBUSER" "$AMPDBPASS" > "$TMP_CNF"
[ -z "$AMPDBHOST" ] || printf 'host=%s\n' "$AMPDBHOST" >> "$TMP_CNF"
chmod 600 "$TMP_CNF"
DEST="''' + destino + r'''"
TABLES="''' + lista + r'''"
mysqldump --defaults-extra-file="$TMP_CNF" --add-drop-table --skip-lock-tables "$AMPDBNAME" $TABLES > "$DEST"
[ -s "$DEST" ] || { echo 'El respaldo salió vacío'; rm -f "$DEST"; exit 22; }
chmod 600 "$DEST"
echo "VP_RUTEO_BACKUP_OK:$DEST"
'''
            out, err, code = _vp_trunk_exec(perfil, comando, timeout=120)
            out = str(out or '')
            if code != 0 or 'VP_RUTEO_BACKUP_OK:' not in out:
                raise RuntimeError('No se pudo crear el respaldo previo. ' + str(err or out or '')[:500])
            m = re.search(r'VP_RUTEO_BACKUP_OK:(/[^\r\n]+)', out)
            if not m:
                raise RuntimeError('El respaldo se creó, pero no se pudo confirmar su ruta.')
            return m.group(1).strip()

        def _vp_re_restore():
            if not re.match(r'^/var/backups/voip-provisioner/ruteo-externo/ruteo_\d{8}_\d{6}\.sql$', str(backup_path or '')):
                raise RuntimeError('La ruta del respaldo no pasó la validación de seguridad.')
            comando = r'''set -eu
AMPDBUSER="$(grep '^AMPDBUSER=' /etc/amportal.conf | head -1 | cut -d= -f2-)"
AMPDBPASS="$(grep '^AMPDBPASS=' /etc/amportal.conf | head -1 | cut -d= -f2-)"
AMPDBHOST="$(grep '^AMPDBHOST=' /etc/amportal.conf | head -1 | cut -d= -f2-)"
AMPDBNAME="$(grep '^AMPDBNAME=' /etc/amportal.conf | head -1 | cut -d= -f2-)"
TMP_CNF="$(mktemp)"
trap 'rm -f "$TMP_CNF"' EXIT
printf '[client]\nuser=%s\npassword=%s\n' "$AMPDBUSER" "$AMPDBPASS" > "$TMP_CNF"
[ -z "$AMPDBHOST" ] || printf 'host=%s\n' "$AMPDBHOST" >> "$TMP_CNF"
chmod 600 "$TMP_CNF"
BACKUP="''' + backup_path + r'''"
[ -s "$BACKUP" ] || { echo 'No se encuentra el respaldo SQL'; exit 41; }
mysql --defaults-extra-file="$TMP_CNF" "$AMPDBNAME" < "$BACKUP"
echo VP_RUTEO_RESTORE_OK
'''
            out, err, code = _vp_trunk_exec(perfil, comando, timeout=120)
            out = str(out or '')
            if code != 0 or 'VP_RUTEO_RESTORE_OK' not in out:
                raise RuntimeError('No se pudo restaurar automáticamente la copia SQL. Respaldo disponible en ' + backup_path + '. Detalle: ' + str(err or out or '')[-500:])
            return out

        def _vp_re_apply():
            comando = (
                "/var/lib/asterisk/bin/retrieve_conf 2>&1 || { echo VP_RUTEO_RETRIEVE_ERROR; exit 31; }; "
                "asterisk -rx 'dialplan reload' 2>&1 || { echo VP_RUTEO_DIALPLAN_ERROR; exit 32; }; "
                "echo VP_RUTEO_APPLY_OK"
            )
            out, err, code = _vp_trunk_exec(perfil, comando, timeout=120)
            out = str(out or '')
            if code != 0 or 'VP_RUTEO_APPLY_OK' not in out:
                raise RuntimeError('Issabel no confirmó la regeneración/recarga. Respaldo: ' + backup_path + '. Detalle: ' + str(err or out or '')[-500:])
            return out

        tablas_backup = []
        sql = ''
        if tipo == 'entrada':
            tabla = 'incoming'
            if not _vp_re_has(tabla, 'extension', 'cidnum', 'destination'):
                raise RuntimeError('Esta versión de Issabel no tiene la estructura esperada de rutas entrantes.')
            old_did = _vp_re_txt(data.get('extension_original'), 'DID actual', 80, True, r'^[A-Za-z0-9_*#+.\-]*$')
            old_cid = _vp_re_txt(data.get('cidnum_original'), 'CID actual', 80, True, r'^[A-Za-z0-9_*#+.\-]*$')
            did = _vp_re_txt(data.get('did'), 'DID', 80, True, r'^[A-Za-z0-9_*#+.\-]*$')
            cid = _vp_re_txt(data.get('caller_id'), 'identificador de origen', 80, True, r'^[A-Za-z0-9_*#+.\-]*$')
            descripcion = _vp_re_txt(data.get('descripcion'), 'descripción', 120, True)
            destino = _vp_re_txt(data.get('destino'), 'destino', 180, False, r'^[A-Za-z0-9_+*#!$.,\-]+$')
            encontrados = _vp_re_sql("SELECT COUNT(*) FROM incoming WHERE COALESCE(extension,'')=" + _vp_re_q(old_did) + " AND COALESCE(cidnum,'')=" + _vp_re_q(old_cid) + ";")
            if not encontrados or not encontrados[-1].isdigit() or int(encontrados[-1]) != 1:
                raise RuntimeError('La ruta entrante cambió o no es única. Actualizá la consulta y volvé a intentarlo.')
            duplicados = _vp_re_sql("SELECT COUNT(*) FROM incoming WHERE COALESCE(extension,'')=" + _vp_re_q(did) + " AND COALESCE(cidnum,'')=" + _vp_re_q(cid) + " AND NOT (COALESCE(extension,'')=" + _vp_re_q(old_did) + " AND COALESCE(cidnum,'')=" + _vp_re_q(old_cid) + ");")
            if duplicados and duplicados[-1].isdigit() and int(duplicados[-1]) > 0:
                raise ValueError('Ya existe otra ruta entrante con ese DID e identificador de origen.')
            updates = []
            for col, val in [('extension', did), ('cidnum', cid), ('description', descripcion), ('destination', destino)]:
                if col in esquema.get(tabla, set()):
                    updates.append('`' + col + '`=' + _vp_re_q(val))
                elif col in ('extension', 'cidnum', 'destination'):
                    raise RuntimeError('La tabla incoming no permite modificar el campo ' + col + '.')
            sql = 'UPDATE incoming SET ' + ','.join(updates) + ' WHERE COALESCE(extension,\'\')=' + _vp_re_q(old_did) + ' AND COALESCE(cidnum,\'\')=' + _vp_re_q(old_cid) + '; SELECT ROW_COUNT();'
            tablas_backup = ['incoming']

        elif tipo == 'condicion':
            tabla = 'timeconditions'
            if not _vp_re_has(tabla, 'timeconditions_id', 'time'):
                raise RuntimeError('No se detectó la tabla de condiciones horarias compatible.')
            ident = str(data.get('id') or '').strip()
            if not re.match(r'^\d+$', ident):
                raise ValueError('ID de condición horaria inválido.')
            group_id = _vp_re_txt(data.get('horario'), 'grupo horario', 12, False, r'^\d+$')
            if int(group_id) <= 0:
                raise ValueError('Seleccioná un grupo horario válido.')
            if 'timegroups_groups' not in esquema or 'id' not in esquema.get('timegroups_groups', set()):
                raise RuntimeError('No se puede comprobar el grupo horario.')
            group_exists = _vp_re_sql('SELECT COUNT(*) FROM timegroups_groups WHERE id=' + group_id + ';')
            if not group_exists or not group_exists[-1].isdigit() or int(group_exists[-1]) != 1:
                raise ValueError('El grupo horario seleccionado no existe.')
            goto_match = _vp_re_txt(data.get('destino_si'), 'destino cuando coincide', 180, False, r'^[A-Za-z0-9_+*#!$.,\-]+$')
            goto_no = _vp_re_txt(data.get('destino_no'), 'destino cuando no coincide', 180, False, r'^[A-Za-z0-9_+*#!$.,\-]+$')
            nombre = _vp_re_txt(data.get('nombre'), 'nombre', 50, True)
            descripcion = _vp_re_txt(data.get('descripcion'), 'descripción', 180, False)
            modo = _vp_re_txt(data.get('modo') or 'time-group', 'modo', 30, True, r'^[A-Za-z0-9_\-]+$')
            zona = _vp_re_txt(data.get('zona_horaria') or 'America/Argentina/Buenos_Aires', 'zona horaria', 80, False, r'^[A-Za-z0-9_+./\-]*$')
            enabled = _vp_re_txt(data.get('habilitada') or '1', 'estado habilitada', 8, True, r'^(1|0|true|false|yes|no|on|off)$').lower()
            vals = {'displayname': nombre, 'description': descripcion, 'time': group_id,
                    'truegoto': goto_match, 'falsegoto': goto_no,
                    'destmatch': goto_match, 'destnomatch': goto_no,
                    'mode': modo, 'timezone': zona, 'enabled': enabled}
            disponibles = esquema.get(tabla, set())
            updates = ['`' + k + '`=' + _vp_re_q(v) for k, v in vals.items() if k in disponibles]
            if not (('truegoto' in disponibles and 'falsegoto' in disponibles) or ('destmatch' in disponibles and 'destnomatch' in disponibles)):
                raise RuntimeError('No se detectaron columnas de destino compatibles para esta condición horaria.')
            count = _vp_re_sql('SELECT COUNT(*) FROM timeconditions WHERE timeconditions_id=' + ident + ';')
            if not count or not count[-1].isdigit() or int(count[-1]) != 1:
                raise RuntimeError('La condición horaria ya no existe o no es única.')
            sql = 'UPDATE timeconditions SET ' + ','.join(updates) + ' WHERE timeconditions_id=' + ident + '; SELECT ROW_COUNT();'
            tablas_backup = ['timeconditions']
        elif tipo == 'grupo':
            tabla_g = 'timegroups_groups'
            tabla_d = 'timegroups_details'
            if not _vp_re_has(tabla_g, 'id', 'description') or not _vp_re_has(tabla_d, 'timegroupid', 'time'):
                raise RuntimeError('Esta versión de Issabel no tiene una estructura compatible de grupos horarios.')
            ident = str(data.get('id') or '').strip()
            if not re.match(r'^\d+$', ident):
                raise ValueError('ID de grupo horario inválido.')
            nombre = _vp_re_txt(data.get('descripcion'), 'nombre del grupo horario', 50, True)
            existe = _vp_re_sql('SELECT COUNT(*) FROM timegroups_groups WHERE id=' + ident + ';')
            if not existe or not existe[-1].isdigit() or int(existe[-1]) != 1:
                raise RuntimeError('El grupo horario ya no existe.')
            franjas = data.get('franjas')
            if not isinstance(franjas, list) or not franjas or len(franjas) > 80:
                raise ValueError('Ingresá al menos una franja horaria válida (máximo 80).')
            validas = []
            for i, fr in enumerate(franjas):
                if not isinstance(fr, dict):
                    raise ValueError('Formato de franja horaria inválido.')
                desc = _vp_re_txt(fr.get('descripcion'), 'descripción de franja', 50, False)
                hora = _vp_re_txt(fr.get('horario'), 'horario de franja', 100, True, r'^[0-9:*.,&+|\- ]+$')
                wday = _vp_re_txt(fr.get('dias_semana') or '*', 'días de semana', 100, True, r'^[A-Za-z0-9_*.,\-]+$')
                mday = _vp_re_txt(fr.get('dias_mes') or '*', 'días del mes', 100, True, r'^[A-Za-z0-9_*.,\-]+$')
                month = _vp_re_txt(fr.get('meses') or '*', 'meses', 100, True, r'^[A-Za-z0-9_*.,\-]+$')
                validas.append((desc, hora, wday, mday, month))
            sql = 'UPDATE timegroups_groups SET description=' + _vp_re_q(nombre) + ' WHERE id=' + ident + '; '
            sql += 'DELETE FROM timegroups_details WHERE timegroupid=' + ident + '; '
            esquema_det = esquema.get(tabla_d, set())
            moderno = not all(c in esquema_det for c in ('wday','mday','month'))
            for desc, hora, wday, mday, month in validas:
                if moderno:
                    full = hora + '|' + wday + '|' + mday + '|' + month
                    dcols = ['timegroupid','time']
                    dvals = [ident, _vp_re_q(full)]
                    sql += 'INSERT INTO timegroups_details (' + ','.join('`' + c + '`' for c in dcols) + ') VALUES (' + ident + ',' + _vp_re_q(full) + '); '
                else:
                    dcols = ['timegroupid','time','wday','mday','month']
                    if 'description' in esquema_det:
                        dcols.append('description')
                    vmap = {'timegroupid': ident, 'time': hora, 'wday': wday, 'mday': mday, 'month': month, 'description': desc}
                    sql += 'INSERT INTO timegroups_details (' + ','.join('`' + c + '`' for c in dcols) + ') VALUES (' + ','.join(_vp_re_q(vmap[c]) for c in dcols) + '); '
            sql += 'SELECT ROW_COUNT();'
            tablas_backup = [tabla_g, tabla_d]
        elif tipo == 'salida':
            tabla = 'outbound_routes'
            ident = str(data.get('id') or '').strip()
            if not re.match(r'^\d+$', ident):
                raise ValueError('ID de ruta saliente inválido.')
            if not _vp_re_has(tabla, 'route_id', 'name'):
                raise RuntimeError('No se detectó una tabla de rutas salientes compatible.')
            nombre = _vp_re_txt(data.get('nombre'), 'nombre de ruta', 80, False, r'^[A-Za-z0-9 _().+\-]+$')
            callerid = _vp_re_txt(data.get('caller_id'), 'CID saliente', 80, True, r'^[A-Za-z0-9_*#+.\-]*$')
            time_group_id = str(data.get('time_group_id') or '0').strip() or '0'
            if not re.match(r'^\d+$', time_group_id):
                raise ValueError('El ID del grupo horario debe ser numérico (0 = sin grupo).')
            posicion = str(data.get('orden') or '').strip()
            if not re.match(r'^\d+$', posicion) or int(posicion) < 1:
                raise ValueError('La posición de la ruta debe ser un número mayor a cero.')
            existe = _vp_re_sql('SELECT COUNT(*) FROM outbound_routes WHERE route_id=' + ident + ';')
            if not existe or not existe[-1].isdigit() or int(existe[-1]) != 1:
                raise RuntimeError('La ruta saliente ya no existe.')
            duplicada = _vp_re_sql('SELECT COUNT(*) FROM outbound_routes WHERE name=' + _vp_re_q(nombre) + ' AND route_id<>' + ident + ';')
            if duplicada and duplicada[-1].isdigit() and int(duplicada[-1]) > 0:
                raise ValueError('Ya existe otra ruta con ese nombre.')
            if 'time_group_id' in esquema.get(tabla, set()) and int(time_group_id) > 0:
                if 'timegroups_groups' not in esquema:
                    raise ValueError('No se detectan grupos horarios para asociar.')
                group_exists = _vp_re_sql('SELECT COUNT(*) FROM timegroups_groups WHERE id=' + time_group_id + ';')
                if not group_exists or not group_exists[-1].isdigit() or int(group_exists[-1]) != 1:
                    raise ValueError('El grupo horario seleccionado no existe.')

            patterns_text = str(data.get('patrones_text') or '')
            patrones = []
            for raw in patterns_text.splitlines():
                raw = raw.strip()
                if not raw:
                    continue
                p = raw.split('|')
                if len(p) != 4:
                    raise ValueError('Cada patrón debe tener cuatro campos separados por |: prefijo|patrón|CID|anteponer.')
                prefijo = _vp_re_txt(p[0], 'prefijo de patrón', 40, True, r'^[A-Za-z0-9_*#+.\[\]\-]*$')
                patron = _vp_re_txt(p[1], 'patrón de marcado', 80, False, r'^[A-Za-z0-9_*#+.\[\]\-]+$')
                cidpat = _vp_re_txt(p[2], 'CID del patrón', 80, True, r'^[A-Za-z0-9_*#+.\-]*$')
                anteponer = _vp_re_txt(p[3], 'dígitos a anteponer', 40, True, r'^[0-9*#+.\-]*$')
                patrones.append((prefijo, patron, cidpat, anteponer))
            if not patrones or len(patrones) > 200:
                raise ValueError('La ruta debe conservar al menos un patrón de marcado y no superar 200 patrones.')
            if not _vp_re_has('outbound_route_patterns', 'route_id', 'match_pattern_pass'):
                raise RuntimeError('No se detectó una tabla compatible de patrones de marcado.')
            if not _vp_re_has('outbound_route_trunks', 'route_id', 'trunk_id'):
                raise RuntimeError('No se detectó una tabla compatible de troncales salientes.')
            tiene_seq_trunk = _vp_re_has('outbound_route_trunks', 'seq')
            trunks_text = str(data.get('troncales_text') or '').strip()
            trunk_ids = []
            for tid in trunks_text.split(','):
                tid = tid.strip()
                if not tid:
                    continue
                if not re.match(r'^\d+$', tid):
                    raise ValueError('Las troncales se indican mediante sus IDs numéricos separados por coma.')
                if tid not in trunk_ids:
                    trunk_ids.append(tid)
            if not trunk_ids or len(trunk_ids) > 100:
                raise ValueError('Seleccioná al menos una troncal para la ruta saliente.')
            if 'trunks' not in esquema or not _vp_re_has('trunks', 'trunkid'):
                raise RuntimeError('No se pudo consultar el catálogo de troncales de Issabel.')
            ids_sql = ','.join(trunk_ids)
            disponibles = _vp_re_sql('SELECT trunkid FROM trunks WHERE trunkid IN (' + ids_sql + ');')
            ids_disponibles = set([x.strip() for x in disponibles if re.match(r'^\d+$', x.strip())])
            if set(trunk_ids) != ids_disponibles:
                raise ValueError('Una o más troncales indicadas ya no existen en Issabel.')

            route_updates = ['`name`=' + _vp_re_q(nombre)]
            if 'outcid' in esquema.get(tabla, set()):
                route_updates.append('`outcid`=' + _vp_re_q(callerid))
            if 'time_group_id' in esquema.get(tabla, set()):
                route_updates.append('`time_group_id`=' + time_group_id)
            sql = 'UPDATE outbound_routes SET ' + ','.join(route_updates) + ' WHERE route_id=' + ident + '; '

            patcols = ['route_id']
            for col in ('match_pattern_prefix','match_pattern_pass','match_cid','prepend_digits'):
                if col in esquema.get('outbound_route_patterns', set()):
                    patcols.append(col)
            sql += 'DELETE FROM outbound_route_patterns WHERE route_id=' + ident + '; '
            for prefijo, patron, cidpat, anteponer in patrones:
                values = {'route_id': ident, 'match_pattern_prefix': prefijo, 'match_pattern_pass': patron, 'match_cid': cidpat, 'prepend_digits': anteponer}
                sql += 'INSERT INTO outbound_route_patterns (' + ','.join('`' + c + '`' for c in patcols) + ') VALUES (' + ','.join(_vp_re_q(values[c]) for c in patcols) + '); '

            sql += 'DELETE FROM outbound_route_trunks WHERE route_id=' + ident + '; '
            for idx, tid in enumerate(trunk_ids):
                tc = ['route_id','trunk_id'] + (['seq'] if tiene_seq_trunk else [])
                tv = {'route_id': ident, 'trunk_id': tid, 'seq': str(idx)}
                sql += 'INSERT INTO outbound_route_trunks (' + ','.join('`' + c + '`' for c in tc) + ') VALUES (' + ','.join(_vp_re_q(tv[c]) for c in tc) + '); '

            if _vp_re_has('outbound_route_sequence', 'route_id', 'seq'):
                seq_rows = _vp_re_sql('SELECT route_id,seq FROM outbound_route_sequence ORDER BY seq,route_id;')
                route_ids = []
                for line in seq_rows:
                    pp = line.split('\t')
                    if pp and pp[0].isdigit() and pp[0] not in route_ids:
                        route_ids.append(pp[0])
                if ident not in route_ids:
                    route_ids.append(ident)
                route_ids = [x for x in route_ids if x != ident]
                target = max(0, min(int(posicion) - 1, len(route_ids)))
                route_ids.insert(target, ident)
                # Valores temporales únicos negativos evitan choques durante reordenación.
                for rid in route_ids:
                    sql += 'UPDATE outbound_route_sequence SET seq=' + str(-100000-int(rid)) + ' WHERE route_id=' + rid + '; '
                for idx, rid in enumerate(route_ids):
                    count_seq = _vp_re_sql('SELECT COUNT(*) FROM outbound_route_sequence WHERE route_id=' + rid + ';')
                    if count_seq and count_seq[-1].isdigit() and int(count_seq[-1]) > 0:
                        sql += 'UPDATE outbound_route_sequence SET seq=' + str(idx) + ' WHERE route_id=' + rid + '; '
                    else:
                        sql += 'INSERT INTO outbound_route_sequence (route_id,seq) VALUES (' + rid + ',' + str(idx) + '); '
            elif posicion:
                raise RuntimeError('Esta versión no admite modificar el orden de las rutas.')
            sql += 'SELECT ROW_COUNT();'
            tablas_backup = ['outbound_routes','outbound_route_patterns','outbound_route_trunks']
            if 'outbound_route_sequence' in esquema:
                tablas_backup.append('outbound_route_sequence')

        # El respaldo sucede ANTES de cualquier INSERT/UPDATE/DELETE.
        backup_path = _vp_re_backup(tablas_backup)
        ok, salida, error = _vp_trunk_run_sql(perfil, sql, timeout=90)
        salida = str(salida or '')
        if (not ok) or re.search(r'(^|\n)ERROR\s+\d+|Unknown column|doesn.t exist|SQL syntax', salida, re.I):
            detalle = str(error or salida or '')[-500:]
            try:
                _vp_re_restore()
                try:
                    _vp_re_apply()
                    detalle += ' Se restauró automáticamente el respaldo y se recargó la configuración anterior.'
                except Exception as recarga_exc:
                    detalle += ' La base se restauró; falló la recarga de la configuración anterior: ' + str(recarga_exc)[:300]
                detalle += ' Respaldo: ' + backup_path
            except Exception as restore_exc:
                detalle += ' ADVERTENCIA: falló la restauración automática: ' + str(restore_exc)[:400] + ' Respaldo manual: ' + backup_path
            raise RuntimeError('Falló la actualización de base de datos. ' + detalle)
        try:
            _vp_re_apply()
        except Exception as exc:
            try:
                _vp_re_restore()
                try:
                    _vp_re_apply()
                    return jsonify({'ok': False, 'mensaje': 'Issabel rechazó la regeneración; se restauró la configuración anterior. Detalle: ' + str(exc)[:300], 'respaldo': backup_path, 'restaurado': True}), 500
                except Exception as recarga_exc:
                    return jsonify({'ok': False, 'mensaje': 'La base se restauró al respaldo, pero falló la recarga posterior. Detalle: ' + str(recarga_exc)[:300], 'respaldo': backup_path, 'restaurado': True}), 500
            except Exception as restore_exc:
                return jsonify({'ok': False, 'mensaje': str(exc) + ' ADVERTENCIA: también falló la restauración automática: ' + str(restore_exc)[:350], 'respaldo': backup_path, 'base_modificada': True}), 500
        return jsonify({'ok': True, 'mensaje': 'Cambios guardados y configuración de Issabel regenerada correctamente.', 'respaldo': backup_path, 'tipo': tipo})
    except ValueError as e:
        return jsonify({'ok': False, 'mensaje': str(e)}), 400
    except Exception as e:
        logger.exception('[RUTEO EXTERNO] Error editando configuración')
        msg = str(e)[:900]
        if backup_path:
            msg += ' Respaldo: ' + backup_path
        return jsonify({'ok': False, 'mensaje': msg, 'respaldo': backup_path}), 500


# VP_RUTEO_EXTERNO_CREAR_API_V3
@main_bp.route('/api/monitoreo/ruteo-externo/crear', methods=['POST'])
def api_monitoreo_ruteo_externo_crear():
    'Crea una ruta entrante, saliente, grupo horario o condición horaria con backup previo.'
    backup_path = ''
    perfil = None
    try:
        data = request.get_json(silent=True) or {}
        if not _vp_trunk_admin_password(data):
            return jsonify({'ok': False, 'mensaje': 'Contraseña de administrador incorrecta.'}), 401
        perfil_nombre = str(data.get('perfil_nombre') or '').strip()
        perfil = buscar_perfil(perfil_nombre, cargar_perfiles()) if perfil_nombre else None
        if not perfil:
            return jsonify({'ok': False, 'mensaje': 'Seleccioná un servidor válido.'}), 400
        tipo = str(data.get('tipo') or '').strip().lower()
        if tipo not in ('entrada', 'salida', 'grupo', 'condicion'):
            return jsonify({'ok': False, 'mensaje': 'Tipo de alta no permitido.'}), 400

        def q(v):
            return _vp_trunk_sql_q(str(v if v is not None else ''))

        def txt(v, nombre, maxlen=160, obligatorio=False, patron=None):
            v = str(v if v is not None else '').strip()
            if obligatorio and not v:
                raise ValueError('Completá el campo ' + nombre + '.')
            if len(v) > maxlen or any(ord(c) < 32 for c in v):
                raise ValueError('El campo ' + nombre + ' tiene una longitud o caracteres inválidos.')
            if patron and v and not re.match(patron, v):
                raise ValueError('El formato del campo ' + nombre + ' no es válido.')
            return v

        def sqlrun(sql, timeout=90):
            ok, out, err = _vp_trunk_run_sql(perfil, sql, timeout=timeout)
            out = str(out or '')
            if (not ok) or re.search(r'(^|\n)ERROR\s+\d+|Unknown column|doesn.t exist|SQL syntax|Duplicate entry', out, re.I):
                raise RuntimeError(str(err or out or 'Error SQL en Issabel.')[-700:])
            return [x for x in out.splitlines() if x.strip()]

        meta_rows = sqlrun(
            "SELECT table_name,column_name,is_nullable,COALESCE(column_default,'__VP_NULL__'),extra "
            "FROM information_schema.columns WHERE table_schema=DATABASE() AND table_name IN "
            "('incoming','outbound_routes','outbound_route_sequence','outbound_route_patterns',"
            "'outbound_route_trunks','trunks','timeconditions','timegroups_groups','timegroups_details') "
            "ORDER BY table_name,ordinal_position;"
        )
        cols = {}
        meta = {}
        for line in meta_rows:
            p = line.split('\t')
            if len(p) >= 5:
                table = p[0].lower(); col = p[1].lower()
                cols.setdefault(table, set()).add(col)
                meta[(table, col)] = {'nullable': p[2].upper(), 'default': p[3], 'extra': p[4].lower()}

        def has(table, *names):
            available = cols.get(table.lower(), set())
            return all(n.lower() in available for n in names)

        def insert_sql(table, values):
            available = cols.get(table, set())
            selected = [(k, v) for k, v in values.items() if k.lower() in available]
            if not selected:
                raise RuntimeError('No se detectaron columnas compatibles para insertar en ' + table + '.')
            provided = set(k.lower() for k, v in selected)
            missing = []
            for col in available:
                info = meta.get((table, col), {})
                if info.get('nullable') == 'NO' and info.get('default') == '__VP_NULL__' and 'auto_increment' not in info.get('extra', '') and col not in provided:
                    missing.append(col)
            if missing:
                raise RuntimeError('La tabla ' + table + ' exige campos adicionales sin valor predeterminado: ' + ', '.join(sorted(missing)) + '. No se insertó nada.')
            return 'INSERT INTO `' + table + '` (' + ','.join('`' + k + '`' for k, v in selected) + ') VALUES (' + ','.join(v if isinstance(v, _VP_SQL_RAW) else q(v) for k, v in selected) + ');'

        class _VP_SQL_RAW(str):
            pass

        def raw(v):
            return _VP_SQL_RAW(str(v))

        def backup(tables):
            allowed = set(['incoming','outbound_routes','outbound_route_sequence','outbound_route_patterns',
                           'outbound_route_trunks','timeconditions','timegroups_groups','timegroups_details'])
            tables = sorted(set(tables))
            if not tables or any(t not in allowed or t not in cols for t in tables):
                raise RuntimeError('No se pudo verificar todas las tablas a respaldar; no se modificó nada.')
            dest = '/var/backups/voip-provisioner/ruteo-externo/ruteo_' + time.strftime('%Y%m%d_%H%M%S') + '.sql'
            command = r'''set -eu
AMPDBUSER="$(grep '^AMPDBUSER=' /etc/amportal.conf | head -1 | cut -d= -f2-)"
AMPDBPASS="$(grep '^AMPDBPASS=' /etc/amportal.conf | head -1 | cut -d= -f2-)"
AMPDBHOST="$(grep '^AMPDBHOST=' /etc/amportal.conf | head -1 | cut -d= -f2-)"
AMPDBNAME="$(grep '^AMPDBNAME=' /etc/amportal.conf | head -1 | cut -d= -f2-)"
[ -n "$AMPDBUSER" ] && [ -n "$AMPDBPASS" ] && [ -n "$AMPDBNAME" ] || { echo 'No se pudieron leer credenciales AMPDB'; exit 20; }
command -v mysqldump >/dev/null 2>&1 || { echo 'No se encuentra mysqldump'; exit 21; }
mkdir -p /var/backups/voip-provisioner/ruteo-externo
chmod 700 /var/backups/voip-provisioner /var/backups/voip-provisioner/ruteo-externo
TMP_CNF="$(mktemp)"
trap 'rm -f "$TMP_CNF"' EXIT
printf '[client]\nuser=%s\npassword=%s\n' "$AMPDBUSER" "$AMPDBPASS" > "$TMP_CNF"
[ -z "$AMPDBHOST" ] || printf 'host=%s\n' "$AMPDBHOST" >> "$TMP_CNF"
chmod 600 "$TMP_CNF"
DEST="''' + dest + r'''"
TABLES="''' + ' '.join(tables) + r'''"
mysqldump --defaults-extra-file="$TMP_CNF" --add-drop-table --skip-lock-tables "$AMPDBNAME" $TABLES > "$DEST"
[ -s "$DEST" ] || { echo 'El respaldo SQL quedó vacío'; rm -f "$DEST"; exit 22; }
chmod 600 "$DEST"
echo "VP_RUTEO_BACKUP_OK:$DEST"
'''
            out, err, code = _vp_trunk_exec(perfil, command, timeout=120)
            out = str(out or '')
            m = re.search(r'VP_RUTEO_BACKUP_OK:(/[^\r\n]+)', out)
            if code != 0 or not m:
                raise RuntimeError('No se pudo generar el respaldo previo. ' + str(err or out or '')[-500:])
            return m.group(1).strip()

        def restore():
            if not re.match(r'^/var/backups/voip-provisioner/ruteo-externo/ruteo_\d{8}_\d{6}\.sql$', backup_path):
                raise RuntimeError('La ruta del respaldo no pasó la validación de seguridad.')
            command = r'''set -eu
AMPDBUSER="$(grep '^AMPDBUSER=' /etc/amportal.conf | head -1 | cut -d= -f2-)"
AMPDBPASS="$(grep '^AMPDBPASS=' /etc/amportal.conf | head -1 | cut -d= -f2-)"
AMPDBHOST="$(grep '^AMPDBHOST=' /etc/amportal.conf | head -1 | cut -d= -f2-)"
AMPDBNAME="$(grep '^AMPDBNAME=' /etc/amportal.conf | head -1 | cut -d= -f2-)"
TMP_CNF="$(mktemp)"
trap 'rm -f "$TMP_CNF"' EXIT
printf '[client]\nuser=%s\npassword=%s\n' "$AMPDBUSER" "$AMPDBPASS" > "$TMP_CNF"
[ -z "$AMPDBHOST" ] || printf 'host=%s\n' "$AMPDBHOST" >> "$TMP_CNF"
chmod 600 "$TMP_CNF"
mysql --defaults-extra-file="$TMP_CNF" "$AMPDBNAME" < "''' + backup_path + r'''"
echo VP_RUTEO_RESTORE_OK
'''
            out, err, code = _vp_trunk_exec(perfil, command, timeout=120)
            if code != 0 or 'VP_RUTEO_RESTORE_OK' not in str(out or ''):
                raise RuntimeError('No se pudo restaurar el respaldo: ' + str(err or out or '')[-400:])

        def apply_and_restore_on_failure():
            command = "/var/lib/asterisk/bin/retrieve_conf 2>&1 || { echo VP_RUTEO_RETRIEVE_ERROR; exit 31; }; asterisk -rx 'dialplan reload' 2>&1 || { echo VP_RUTEO_DIALPLAN_ERROR; exit 32; }; echo VP_RUTEO_APPLY_OK"
            out, err, code = _vp_trunk_exec(perfil, command, timeout=120)
            if code != 0 or 'VP_RUTEO_APPLY_OK' not in str(out or ''):
                detail = str(err or out or '')[-400:]
                try:
                    restore()
                    out2, err2, code2 = _vp_trunk_exec(perfil, command, timeout=120)
                    if code2 == 0 and 'VP_RUTEO_APPLY_OK' in str(out2 or ''):
                        raise RuntimeError('Issabel rechazó la aplicación. Se restauró la configuración anterior. Detalle: ' + detail)
                    raise RuntimeError('La base se restauró, pero falló la recarga posterior. Respaldo: ' + backup_path + '. ' + str(err2 or out2 or '')[-300:])
                except RuntimeError as restore_error:
                    if 'Se restauró la configuración anterior' in str(restore_error) or 'La base se restauró' in str(restore_error):
                        raise
                    raise RuntimeError('Falló la aplicación y también la restauración automática. Respaldo manual: ' + backup_path + '. ' + str(restore_error)[:300])
                except Exception as restore_error:
                    raise RuntimeError('Falló la aplicación y también la restauración automática. Respaldo manual: ' + backup_path + '. ' + str(restore_error)[:300])

        sql = ''
        tables_backup = []
        if tipo == 'entrada':
            if not has('incoming', 'extension', 'cidnum', 'destination'):
                raise RuntimeError('No se detectó la estructura habitual de rutas entrantes en esta versión de Issabel.')
            did = txt(data.get('did'), 'DID', 80, False, r'^[A-Za-z0-9_*#+.\-]*$')
            cid = txt(data.get('caller_id'), 'identificador de origen', 80, False, r'^[A-Za-z0-9_*#+.\-]*$')
            description = txt(data.get('descripcion'), 'descripción', 120, False)
            destination = txt(data.get('destino'), 'destino', 180, True, r'^[A-Za-z0-9_+*#!$.,\-]+$')
            dup = sqlrun('SELECT COUNT(*) FROM incoming WHERE COALESCE(extension,\'\')=' + q(did) + ' AND COALESCE(cidnum,\'\')=' + q(cid) + ';')
            if dup and dup[-1].isdigit() and int(dup[-1]) > 0:
                raise ValueError('Ya existe una ruta entrante con ese DID e identificador de origen.')
            vals = {'extension': did, 'cidnum': cid, 'description': description, 'destination': destination}
            sql = insert_sql('incoming', vals)
            sql += ' SELECT \'VP_RUTEO_ENTRADA_CREADA\';'
            tables_backup = ['incoming']

        elif tipo == 'salida':
            if not has('outbound_routes', 'route_id', 'name') or not has('outbound_route_patterns', 'route_id', 'match_pattern_pass') or not has('outbound_route_trunks', 'route_id', 'trunk_id') or not has('trunks', 'trunkid'):
                raise RuntimeError('Faltan tablas necesarias para crear una ruta saliente.')
            nombre = txt(data.get('nombre'), 'nombre de ruta', 80, True, r'^[A-Za-z0-9 _().+\-]+$')
            cidout = txt(data.get('caller_id'), 'CID saliente', 80, False, r'^[A-Za-z0-9_*#+.\-]*$')
            tg = str(data.get('time_group_id') or '0').strip() or '0'
            if not re.match(r'^\d+$', tg):
                raise ValueError('Seleccioná un grupo horario válido.')
            if int(tg) > 0:
                if not has('timegroups_groups', 'id'):
                    raise RuntimeError('Esta versión no permite asociar grupos horarios desde el panel.')
                ver = sqlrun('SELECT COUNT(*) FROM timegroups_groups WHERE id=' + tg + ';')
                if not ver or not ver[-1].isdigit() or int(ver[-1]) != 1:
                    raise ValueError('El grupo horario seleccionado no existe.')
            dupe = sqlrun('SELECT COUNT(*) FROM outbound_routes WHERE name=' + q(nombre) + ';')
            if dupe and dupe[-1].isdigit() and int(dupe[-1]) > 0:
                raise ValueError('Ya existe una ruta saliente con ese nombre.')
            patterns = data.get('patrones')
            if not isinstance(patterns, list) or not patterns or len(patterns) > 100:
                raise ValueError('Agregá al menos un patrón de marcado.')
            clean_patterns = []
            for idx, pat in enumerate(patterns):
                if not isinstance(pat, dict):
                    raise ValueError('Hay un patrón de marcado inválido.')
                pref = txt(pat.get('prefijo'), 'prefijo', 40, False, r'^[A-Za-z0-9_*#+.\[\]\-]*$')
                match = txt(pat.get('patron'), 'patrón', 80, True, r'^[A-Za-z0-9_*#+.\[\]\-]+$')
                matchcid = txt(pat.get('cid'), 'CID del patrón', 80, False, r'^[A-Za-z0-9_*#+.\-]*$')
                prepend = txt(pat.get('anteponer'), 'dígitos a anteponer', 40, False, r'^[0-9*#+.\-]*$')
                clean_patterns.append((pref, match, matchcid, prepend))
            trunks = data.get('troncales')
            if not isinstance(trunks, list) or not trunks or len(trunks) > 50:
                raise ValueError('Seleccioná al menos una troncal de salida.')
            trunk_ids = []
            for val in trunks:
                val = str(val or '').strip()
                if not re.match(r'^\d+$', val):
                    raise ValueError('La lista de troncales contiene un ID inválido.')
                if val not in trunk_ids:
                    trunk_ids.append(val)
            if not trunk_ids:
                raise ValueError('Seleccioná al menos una troncal de salida.')
            found = sqlrun('SELECT trunkid FROM trunks WHERE trunkid IN (' + ','.join(trunk_ids) + ');')
            available = set(x.strip() for x in found if x.strip().isdigit())
            if available != set(trunk_ids):
                raise ValueError('Alguna troncal seleccionada ya no existe. Actualizá el panel.')
            route_values = {
                'name': nombre, 'outcid': cidout, 'outcid_mode': '', 'password': '',
                'emergency_route': 'NO', 'intracompany_route': 'NO', 'mohclass': '', 'time_group_id': tg
            }
            sql = insert_sql('outbound_routes', route_values)
            sql += ' SET @VP_RUTEO_NEW_ROUTE=LAST_INSERT_ID(); '
            pcols = ['route_id']
            for col in ('match_pattern_prefix','match_pattern_pass','match_cid','prepend_digits'):
                if col in cols.get('outbound_route_patterns', set()):
                    pcols.append(col)
            for pref, match, matchcid, prepend in clean_patterns:
                pvals = {'route_id': raw('@VP_RUTEO_NEW_ROUTE'), 'match_pattern_prefix': pref, 'match_pattern_pass': match, 'match_cid': matchcid, 'prepend_digits': prepend}
                selected = [(c, pvals[c]) for c in pcols]
                sql += ' INSERT INTO outbound_route_patterns (' + ','.join('`' + c + '`' for c, v in selected) + ') VALUES (' + ','.join(v if isinstance(v, _VP_SQL_RAW) else q(v) for c, v in selected) + '); '
            has_trunk_seq = 'seq' in cols.get('outbound_route_trunks', set())
            for i, tid in enumerate(trunk_ids):
                sql += ' INSERT INTO outbound_route_trunks (`route_id`,`trunk_id`' + (',`seq`' if has_trunk_seq else '') + ') VALUES (@VP_RUTEO_NEW_ROUTE,' + tid + (',' + str(i) if has_trunk_seq else '') + '); '
            if has('outbound_route_sequence', 'route_id', 'seq'):
                sql += ' INSERT INTO outbound_route_sequence (`route_id`,`seq`) SELECT @VP_RUTEO_NEW_ROUTE,COALESCE(MAX(`seq`),-1)+1 FROM outbound_route_sequence; '
            sql += " SELECT CONCAT('VP_RUTEO_SALIDA_CREADA|',@VP_RUTEO_NEW_ROUTE);"
            tables_backup = ['outbound_routes','outbound_route_patterns','outbound_route_trunks']
            if 'outbound_route_sequence' in cols:
                tables_backup.append('outbound_route_sequence')

        elif tipo == 'grupo':
            if not has('timegroups_groups', 'id', 'description') or not has('timegroups_details', 'timegroupid', 'time'):
                raise RuntimeError('No se detectaron las tablas habituales de grupos horarios.')
            nombre = txt(data.get('descripcion'), 'nombre del horario', 50, True)
            dupe = sqlrun('SELECT COUNT(*) FROM timegroups_groups WHERE description=' + q(nombre) + ';')
            if dupe and dupe[-1].isdigit() and int(dupe[-1]) > 0:
                raise ValueError('Ya existe un horario con ese nombre.')
            slots = data.get('franjas')
            if not isinstance(slots, list) or not slots or len(slots) > 50:
                raise ValueError('Agregá al menos una franja horaria.')
            clean_slots = []
            for slot in slots:
                if not isinstance(slot, dict):
                    raise ValueError('Hay una franja horaria inválida.')
                desc = txt(slot.get('descripcion'), 'descripción de franja', 50, False)
                horario = txt(slot.get('horario'), 'horario', 100, True, r'^[0-9:*.,&+|\- ]+$')
                wday = txt(slot.get('dias_semana') or '*', 'días de semana', 100, True, r'^[A-Za-z0-9_*.,\-]+$')
                mday = txt(slot.get('dias_mes') or '*', 'días del mes', 100, True, r'^[A-Za-z0-9_*.,\-]+$')
                month = txt(slot.get('meses') or '*', 'meses', 100, True, r'^[A-Za-z0-9_*.,\-]+$')
                clean_slots.append((desc, horario, wday, mday, month))
            old_schema = has('timegroups_details', 'wday', 'mday', 'month')
            values = {'description': nombre}
            sql = insert_sql('timegroups_groups', values)
            sql += ' SET @VP_RUTEO_NEW_GROUP=LAST_INSERT_ID(); '
            for desc, horario, wday, mday, month in clean_slots:
                if old_schema:
                    detvals = {'timegroupid': raw('@VP_RUTEO_NEW_GROUP'), 'time': horario, 'wday': wday, 'mday': mday, 'month': month, 'description': desc}
                    sql += insert_sql('timegroups_details', detvals)
                else:
                    # Esquema moderno: FreePBX guarda horario|días-semana|días-mes|meses en una sola columna.
                    full_time = horario + '|' + wday + '|' + mday + '|' + month
                    detvals = {'timegroupid': raw('@VP_RUTEO_NEW_GROUP'), 'time': full_time}
                    sql += insert_sql('timegroups_details', detvals)
            sql += " SELECT CONCAT('VP_RUTEO_GRUPO_CREADO|',@VP_RUTEO_NEW_GROUP);"
            tables_backup = ['timegroups_groups','timegroups_details']

        elif tipo == 'condicion':
            if not has('timeconditions', 'timeconditions_id', 'displayname', 'time'):
                raise RuntimeError('No se detectó la tabla de condiciones horarias compatible.')
            if not (has('timeconditions', 'truegoto', 'falsegoto') or has('timeconditions', 'destmatch', 'destnomatch')):
                raise RuntimeError('La versión de Issabel no expone los dos destinos de la condición horaria; no se modificó nada.')
            nombre = txt(data.get('nombre'), 'nombre de la condición', 50, True)
            if 'displayname' in cols.get('timeconditions', set()):
                dupe_condition = sqlrun('SELECT COUNT(*) FROM timeconditions WHERE displayname=' + q(nombre) + ';')
                if dupe_condition and dupe_condition[-1].isdigit() and int(dupe_condition[-1]) > 0:
                    raise ValueError('Ya existe una condición horaria con ese nombre.')
            group_id = str(data.get('time_group_id') or '').strip()
            if not re.match(r'^\d+$', group_id) or int(group_id) <= 0:
                raise ValueError('Seleccioná un grupo horario existente para la condición.')
            if not has('timegroups_groups', 'id'):
                raise RuntimeError('No se puede comprobar el grupo horario en esta versión.')
            group_check = sqlrun('SELECT COUNT(*) FROM timegroups_groups WHERE id=' + group_id + ';')
            if not group_check or not group_check[-1].isdigit() or int(group_check[-1]) != 1:
                raise ValueError('El grupo horario seleccionado no existe.')
            goto_match = txt(data.get('destino_si'), 'destino cuando coincide', 180, True, r'^[A-Za-z0-9_+*#!$.,\-]+$')
            goto_no = txt(data.get('destino_no'), 'destino cuando no coincide', 180, True, r'^[A-Za-z0-9_+*#!$.,\-]+$')
            zone = txt(data.get('zona_horaria') or 'America/Argentina/Buenos_Aires', 'zona horaria', 80, False, r'^[A-Za-z0-9_+./\-]*$')
            values = {'displayname': nombre, 'time': group_id, 'truegoto': goto_match, 'falsegoto': goto_no,
                      'destmatch': goto_match, 'destnomatch': goto_no, 'mode': 'time-group', 'timezone': zone,
                      'description': nombre, 'enabled': '1', 'generate_hint': '0', 'priority': '0'}
            # Para este módulo, el campo time es el ID del grupo horario, no la expresión del horario.
            sql = insert_sql('timeconditions', values)
            sql += " SELECT CONCAT('VP_RUTEO_CONDICION_CREADA|',LAST_INSERT_ID());"
            tables_backup = ['timeconditions']

        # Respaldo garantizado antes del primer INSERT.
        backup_path = backup(tables_backup)
        output = sqlrun(sql, timeout=100)
        if not any(('VP_RUTEO_ENTRADA_CREADA' in x or 'VP_RUTEO_SALIDA_CREADA|' in x or 'VP_RUTEO_GRUPO_CREADO|' in x or 'VP_RUTEO_CONDICION_CREADA|' in x) for x in output):
            raise RuntimeError('Issabel no confirmó la creación. Se intentará restaurar el respaldo.')
        try:
            apply_and_restore_on_failure()
        except Exception:
            raise
        created_id = ''
        for line in output:
            m = re.search(r'VP_RUTEO_(?:SALIDA|GRUPO|CONDICION)_CREADA\|(\d+)', line)
            if m:
                created_id = m.group(1)
        return jsonify({'ok': True, 'mensaje': 'Configuración creada y aplicada en Issabel.', 'tipo': tipo, 'id': created_id, 'respaldo': backup_path})
    except ValueError as exc:
        return jsonify({'ok': False, 'mensaje': str(exc), 'respaldo': backup_path}), 400
    except Exception as exc:
        logger.exception('[RUTEO EXTERNO] Error creando configuración')
        detail = str(exc)[:900]
        if backup_path:
            # Si falló SQL luego del respaldo, restauramos solo si se obtuvo backup.
            try:
                if perfil:
                    allowed = re.match(r'^/var/backups/voip-provisioner/ruteo-externo/ruteo_\d{8}_\d{6}\.sql$', backup_path)
                    if allowed:
                        command = r'''set -eu
AMPDBUSER="$(grep '^AMPDBUSER=' /etc/amportal.conf | head -1 | cut -d= -f2-)"
AMPDBPASS="$(grep '^AMPDBPASS=' /etc/amportal.conf | head -1 | cut -d= -f2-)"
AMPDBHOST="$(grep '^AMPDBHOST=' /etc/amportal.conf | head -1 | cut -d= -f2-)"
AMPDBNAME="$(grep '^AMPDBNAME=' /etc/amportal.conf | head -1 | cut -d= -f2-)"
TMP_CNF="$(mktemp)"; trap 'rm -f "$TMP_CNF"' EXIT
printf '[client]\nuser=%s\npassword=%s\n' "$AMPDBUSER" "$AMPDBPASS" > "$TMP_CNF"
[ -z "$AMPDBHOST" ] || printf 'host=%s\n' "$AMPDBHOST" >> "$TMP_CNF"
chmod 600 "$TMP_CNF"
mysql --defaults-extra-file="$TMP_CNF" "$AMPDBNAME" < "''' + backup_path + r'''"
echo VP_RUTEO_RESTORE_OK
'''
                        out, err, code = _vp_trunk_exec(perfil, command, timeout=120)
                        if code == 0 and 'VP_RUTEO_RESTORE_OK' in str(out or ''):
                            detail += ' Se restauró el respaldo SQL previo.'
                        else:
                            detail += ' ADVERTENCIA: no se pudo restaurar automáticamente; respaldo manual: ' + backup_path + '. ' + str(err or out or '')[-300:]
            except Exception as restore_exc:
                detail += ' ADVERTENCIA: error restaurando el respaldo: ' + str(restore_exc)[:300]
            detail += ' Respaldo: ' + backup_path
        return jsonify({'ok': False, 'mensaje': detail, 'respaldo': backup_path}), 500


# VP_RUTEO_EXTERNO_DID_MASIVO_V5

def _vp_ruteo_v5_sqlrun(perfil, sql, timeout=90):
    ok, salida, error = _vp_trunk_run_sql(perfil, sql, timeout=timeout)
    salida = str(salida or '')
    if (not ok) or re.search(r'(^|\n)ERROR\s+\d+|Unknown column|doesn.t exist|SQL syntax|Duplicate entry|Access denied', salida, re.I):
        raise RuntimeError(str(error or salida or 'Error SQL en Issabel.')[-800:])
    return [x for x in salida.splitlines() if x.strip()]


def _vp_ruteo_v5_schema(perfil):
    rows = _vp_ruteo_v5_sqlrun(
        perfil,
        "SELECT column_name,is_nullable,COALESCE(column_default,'__VP_NULL__'),extra "
        "FROM information_schema.columns "
        "WHERE table_schema=DATABASE() AND table_name='incoming' "
        "ORDER BY ordinal_position;"
    )
    cols = set()
    meta = {}
    for line in rows:
        p = line.split('\t')
        if len(p) >= 4:
            col = p[0].lower()
            cols.add(col)
            meta[col] = {'nullable': p[1].upper(), 'default': p[2], 'extra': p[3].lower()}
    if not set(['extension', 'cidnum', 'destination']).issubset(cols):
        raise RuntimeError('La tabla incoming no tiene la estructura necesaria (extension, cidnum, destination).')
    return cols, meta


def _vp_ruteo_v5_backup(perfil):
    destino = '/var/backups/voip-provisioner/ruteo-externo/ruteo_v5_' + time.strftime('%Y%m%d_%H%M%S') + '_' + str(os.getpid()) + '.sql'
    comando = r'''set -eu
AMPDBUSER="$(grep '^AMPDBUSER=' /etc/amportal.conf | head -1 | cut -d= -f2-)"
AMPDBPASS="$(grep '^AMPDBPASS=' /etc/amportal.conf | head -1 | cut -d= -f2-)"
AMPDBHOST="$(grep '^AMPDBHOST=' /etc/amportal.conf | head -1 | cut -d= -f2-)"
AMPDBNAME="$(grep '^AMPDBNAME=' /etc/amportal.conf | head -1 | cut -d= -f2-)"
[ -n "$AMPDBUSER" ] && [ -n "$AMPDBPASS" ] && [ -n "$AMPDBNAME" ] || { echo 'No se pudieron leer credenciales AMPDB'; exit 20; }
command -v mysqldump >/dev/null 2>&1 || { echo 'No se encuentra mysqldump'; exit 21; }
mkdir -p /var/backups/voip-provisioner/ruteo-externo
chmod 700 /var/backups/voip-provisioner /var/backups/voip-provisioner/ruteo-externo
TMP_CNF="$(mktemp)"
trap 'rm -f "$TMP_CNF"' EXIT
printf '[client]\nuser=%s\npassword=%s\n' "$AMPDBUSER" "$AMPDBPASS" > "$TMP_CNF"
[ -z "$AMPDBHOST" ] || printf 'host=%s\n' "$AMPDBHOST" >> "$TMP_CNF"
chmod 600 "$TMP_CNF"
DEST="''' + destino + r'''"
mysqldump --defaults-extra-file="$TMP_CNF" --add-drop-table --skip-lock-tables "$AMPDBNAME" incoming > "$DEST"
[ -s "$DEST" ] || { echo 'El respaldo SQL quedó vacío'; rm -f "$DEST"; exit 22; }
chmod 600 "$DEST"
echo "VP_RUTEO_V5_BACKUP_OK:$DEST"
'''
    out, err, code = _vp_trunk_exec(perfil, comando, timeout=120)
    out = str(out or '')
    match = re.search(r'VP_RUTEO_V5_BACKUP_OK:(/[^\r\n]+)', out)
    if code != 0 or not match:
        raise RuntimeError('No se pudo generar el respaldo SQL previo. ' + str(err or out or '')[-500:])
    return match.group(1).strip()


def _vp_ruteo_v5_restore(perfil, backup_path):
    if not re.match(r'^/var/backups/voip-provisioner/ruteo-externo/ruteo_v5_\d{8}_\d{6}_\d+\.sql$', str(backup_path or '')):
        raise RuntimeError('La ruta del respaldo no pasó la validación de seguridad.')
    comando = r'''set -eu
AMPDBUSER="$(grep '^AMPDBUSER=' /etc/amportal.conf | head -1 | cut -d= -f2-)"
AMPDBPASS="$(grep '^AMPDBPASS=' /etc/amportal.conf | head -1 | cut -d= -f2-)"
AMPDBHOST="$(grep '^AMPDBHOST=' /etc/amportal.conf | head -1 | cut -d= -f2-)"
AMPDBNAME="$(grep '^AMPDBNAME=' /etc/amportal.conf | head -1 | cut -d= -f2-)"
TMP_CNF="$(mktemp)"
trap 'rm -f "$TMP_CNF"' EXIT
printf '[client]\nuser=%s\npassword=%s\n' "$AMPDBUSER" "$AMPDBPASS" > "$TMP_CNF"
[ -z "$AMPDBHOST" ] || printf 'host=%s\n' "$AMPDBHOST" >> "$TMP_CNF"
chmod 600 "$TMP_CNF"
BACKUP="''' + backup_path + r'''"
[ -s "$BACKUP" ] || { echo 'No se encuentra el respaldo SQL'; exit 41; }
mysql --defaults-extra-file="$TMP_CNF" "$AMPDBNAME" < "$BACKUP"
echo VP_RUTEO_V5_RESTORE_OK
'''
    out, err, code = _vp_trunk_exec(perfil, comando, timeout=120)
    if code != 0 or 'VP_RUTEO_V5_RESTORE_OK' not in str(out or ''):
        raise RuntimeError('No se pudo restaurar el respaldo. ' + str(err or out or '')[-500:])


def _vp_ruteo_v5_apply(perfil):
    comando = "/var/lib/asterisk/bin/retrieve_conf 2>&1 || { echo VP_RUTEO_V5_RETRIEVE_ERROR; exit 31; }; asterisk -rx 'dialplan reload' 2>&1 || { echo VP_RUTEO_V5_DIALPLAN_ERROR; exit 32; }; echo VP_RUTEO_V5_APPLY_OK"
    out, err, code = _vp_trunk_exec(perfil, comando, timeout=120)
    if code != 0 or 'VP_RUTEO_V5_APPLY_OK' not in str(out or ''):
        raise RuntimeError('Issabel no confirmó retrieve_conf y dialplan reload. ' + str(err or out or '')[-500:])


def _vp_ruteo_v5_validar_txt(valor, nombre, maxlen, requerido, patron):
    valor = str(valor if valor is not None else '').strip()
    if requerido and not valor:
        raise ValueError('Completá el campo ' + nombre + '.')
    if len(valor) > maxlen or any(ord(c) < 32 for c in valor):
        raise ValueError('El campo ' + nombre + ' tiene una longitud o caracteres inválidos.')
    if patron and valor and not re.match(patron, valor):
        raise ValueError('El formato del campo ' + nombre + ' no es válido.')
    return valor


def _vp_ruteo_v5_where_entrada(item, cols):
    q = _vp_trunk_sql_q
    partes = ["COALESCE(`extension`,'')=" + q(item.get('did', '')),
              "COALESCE(`cidnum`,'')=" + q(item.get('caller_id', ''))]
    if 'description' in cols:
        partes.append("COALESCE(`description`,'')=" + q(item.get('descripcion', '')))
    if 'destination' in cols:
        partes.append("COALESCE(`destination`,'')=" + q(item.get('destino', '')))
    return ' AND '.join(partes)


def _vp_ruteo_v5_restaurar_si_falla(perfil, backup_path, detalle):
    try:
        _vp_ruteo_v5_restore(perfil, backup_path)
        _vp_ruteo_v5_apply(perfil)
        return str(detalle) + ' Se restauró la tabla incoming y se recargó la configuración anterior.'
    except Exception as exc:
        return str(detalle) + ' ATENCIÓN: falló la restauración automática: ' + str(exc)[:350] + '. Respaldo manual: ' + str(backup_path)


@main_bp.route('/api/monitoreo/ruteo-externo/entradas/crear-lote', methods=['POST'])
def api_monitoreo_ruteo_externo_entradas_crear_lote_v5():
    backup_path = ''
    perfil = None
    try:
        data = request.get_json(silent=True) or {}
        if not _vp_trunk_admin_password(data):
            return jsonify({'ok': False, 'mensaje': 'Contraseña de administrador incorrecta.'}), 401
        perfil_nombre = str(data.get('perfil_nombre') or '').strip()
        perfil = buscar_perfil(perfil_nombre, cargar_perfiles()) if perfil_nombre else None
        if not perfil:
            return jsonify({'ok': False, 'mensaje': 'Seleccioná un servidor válido.'}), 400
        rutas = data.get('rutas')
        if not isinstance(rutas, list) or not rutas:
            raise ValueError('Ingresá al menos un DID.')
        if len(rutas) > 200:
            raise ValueError('Podés crear hasta 200 DID por operación. Dividí la lista en varios lotes.')

        cols, meta = _vp_ruteo_v5_schema(perfil)
        clean = []
        vistos = set()
        q = _vp_trunk_sql_q
        for idx, row in enumerate(rutas):
            if not isinstance(row, dict):
                raise ValueError('La línea ' + str(idx + 1) + ' no es válida.')
            did = _vp_ruteo_v5_validar_txt(row.get('did'), 'DID de la línea ' + str(idx + 1), 80, True, r'^[A-Za-z0-9_*#+.\-]+$')
            cid = _vp_ruteo_v5_validar_txt(row.get('caller_id'), 'CID de la línea ' + str(idx + 1), 80, False, r'^[A-Za-z0-9_*#+.\-]*$')
            desc = _vp_ruteo_v5_validar_txt(row.get('descripcion'), 'descripción de la línea ' + str(idx + 1), 120, False, None)
            dest = _vp_ruteo_v5_validar_txt(row.get('destino'), 'destino de la línea ' + str(idx + 1), 180, True, r'^[A-Za-z0-9_+*#!$.,\-]+$')
            key = (did, cid)
            if key in vistos:
                raise ValueError('La lista repite el DID ' + did + ' con el mismo CID ' + (cid or '(cualquiera)') + '.')
            vistos.add(key)
            clean.append({'did': did, 'caller_id': cid, 'descripcion': desc, 'destino': dest})

        for row in clean:
            dup = _vp_ruteo_v5_sqlrun(perfil,
                "SELECT COUNT(*) FROM incoming WHERE COALESCE(extension,'')=" + q(row['did']) +
                " AND COALESCE(cidnum,'')=" + q(row['caller_id']) + ";")
            if not dup or not dup[-1].isdigit():
                raise RuntimeError('No se pudo comprobar si ya existe el DID ' + row['did'] + '.')
            if int(dup[-1]) > 0:
                raise ValueError('Ya existe el DID ' + row['did'] + ' con CID ' + (row['caller_id'] or '(cualquiera)') + '. No se creó ningún elemento del lote.')

        values = {'extension': '', 'cidnum': '', 'description': '', 'destination': ''}
        selected = [c for c in ('extension', 'cidnum', 'description', 'destination') if c in cols]
        provided = set(selected)
        missing = []
        for col in cols:
            info = meta.get(col, {})
            if info.get('nullable') == 'NO' and info.get('default') == '__VP_NULL__' and 'auto_increment' not in info.get('extra', '') and col not in provided:
                missing.append(col)
        if missing:
            raise RuntimeError('La tabla incoming exige campos que no tienen valor predeterminado: ' + ', '.join(sorted(missing)) + '. No se modificó nada.')
        col_sql = ','.join('`' + c + '`' for c in selected)
        tuples = []
        for row in clean:
            vals = {'extension': row['did'], 'cidnum': row['caller_id'], 'description': row['descripcion'], 'destination': row['destino']}
            tuples.append('(' + ','.join(q(vals[c]) for c in selected) + ')')
        sql = 'INSERT INTO `incoming` (' + col_sql + ') VALUES ' + ','.join(tuples) + '; SELECT CONCAT(\'VP_RUTEO_V5_DID_CREADOS|\',ROW_COUNT());'

        backup_path = _vp_ruteo_v5_backup(perfil)
        rows_out = _vp_ruteo_v5_sqlrun(perfil, sql, timeout=120)
        confirmados = []
        for line in rows_out:
            m = re.search(r'VP_RUTEO_V5_DID_CREADOS\|(\d+)', line)
            if m:
                confirmados.append(int(m.group(1)))
        if not confirmados or confirmados[-1] != len(clean):
            detalle = 'Issabel no confirmó la creación de todos los DID. '
            detalle = _vp_ruteo_v5_restaurar_si_falla(perfil, backup_path, detalle)
            backup_path = ''
            raise RuntimeError(detalle)
        try:
            _vp_ruteo_v5_apply(perfil)
        except Exception as exc:
            detalle = _vp_ruteo_v5_restaurar_si_falla(perfil, backup_path, 'Falló la aplicación del lote: ' + str(exc)[:400])
            backup_path = ''
            raise RuntimeError(detalle)
        return jsonify({'ok': True, 'mensaje': 'Se crearon ' + str(len(clean)) + ' rutas entrantes y se aplicó la configuración.', 'creadas': len(clean), 'respaldo': backup_path})
    except ValueError as exc:
        return jsonify({'ok': False, 'mensaje': str(exc), 'respaldo': backup_path}), 400
    except Exception as exc:
        logger.exception('[RUTEO EXTERNO V5] Error creando DID masivos')
        detalle = str(exc)[:900]
        if backup_path and perfil:
            # Un INSERT multi-fila es una sola sentencia; ante una respuesta incierta se restaura el estado previo.
            detalle = _vp_ruteo_v5_restaurar_si_falla(perfil, backup_path, detalle)
        return jsonify({'ok': False, 'mensaje': detalle, 'respaldo': backup_path}), 500


@main_bp.route('/api/monitoreo/ruteo-externo/entradas/eliminar-lote', methods=['POST'])
def api_monitoreo_ruteo_externo_entradas_eliminar_lote_v5():
    backup_path = ''
    perfil = None
    try:
        data = request.get_json(silent=True) or {}
        if not _vp_trunk_admin_password(data):
            return jsonify({'ok': False, 'mensaje': 'Contraseña de administrador incorrecta.'}), 401
        perfil_nombre = str(data.get('perfil_nombre') or '').strip()
        perfil = buscar_perfil(perfil_nombre, cargar_perfiles()) if perfil_nombre else None
        if not perfil:
            return jsonify({'ok': False, 'mensaje': 'Seleccioná un servidor válido.'}), 400
        rutas = data.get('rutas')
        if not isinstance(rutas, list) or not rutas:
            raise ValueError('Seleccioná al menos una ruta DID para eliminar.')
        if len(rutas) > 200:
            raise ValueError('Podés eliminar hasta 200 rutas por operación.')
        cols, meta = _vp_ruteo_v5_schema(perfil)
        clean = []
        vistos = set()
        for idx, row in enumerate(rutas):
            if not isinstance(row, dict):
                raise ValueError('La selección ' + str(idx + 1) + ' no es válida.')
            item = {
                'did': _vp_ruteo_v5_validar_txt(row.get('did'), 'DID', 80, False, r'^[A-Za-z0-9_*#+.\-]*$'),
                'caller_id': _vp_ruteo_v5_validar_txt(row.get('caller_id'), 'CID', 80, False, r'^[A-Za-z0-9_*#+.\-]*$'),
                'descripcion': _vp_ruteo_v5_validar_txt(row.get('descripcion'), 'descripción', 120, False, None),
                'destino': _vp_ruteo_v5_validar_txt(row.get('destino'), 'destino', 180, False, r'^[A-Za-z0-9_+*#!$.,\-]*$')
            }
            key = (item['did'], item['caller_id'], item['descripcion'], item['destino'])
            if key in vistos:
                raise ValueError('La selección contiene una ruta repetida.')
            vistos.add(key)
            clean.append(item)

        where_list = []
        for item in clean:
            where = _vp_ruteo_v5_where_entrada(item, cols)
            count = _vp_ruteo_v5_sqlrun(perfil, 'SELECT COUNT(*) FROM incoming WHERE ' + where + ';')
            if not count or not count[-1].isdigit() or int(count[-1]) != 1:
                raise ValueError('No se pudo identificar una única ruta para el DID ' + (item['did'] or '(DID vacío)') + '. Actualizá la lista y volvé a seleccionar; no se eliminó nada.')
            where_list.append('(' + where + ')')

        backup_path = _vp_ruteo_v5_backup(perfil)
        sql = 'DELETE FROM incoming WHERE ' + ' OR '.join(where_list) + '; SELECT CONCAT(\'VP_RUTEO_V5_DID_ELIMINADOS|\',ROW_COUNT());'
        out = _vp_ruteo_v5_sqlrun(perfil, sql, timeout=120)
        deleted = None
        for line in out:
            m = re.search(r'VP_RUTEO_V5_DID_ELIMINADOS\|(\d+)', line)
            if m:
                deleted = int(m.group(1))
        if deleted != len(clean):
            detalle = 'La cantidad eliminada no coincide con la selección (' + str(deleted) + ' de ' + str(len(clean)) + '). '
            detalle = _vp_ruteo_v5_restaurar_si_falla(perfil, backup_path, detalle)
            backup_path = ''
            raise RuntimeError(detalle)
        try:
            _vp_ruteo_v5_apply(perfil)
        except Exception as exc:
            detalle = _vp_ruteo_v5_restaurar_si_falla(perfil, backup_path, 'Falló la aplicación de la eliminación: ' + str(exc)[:400])
            backup_path = ''
            raise RuntimeError(detalle)
        return jsonify({'ok': True, 'mensaje': 'Se eliminaron ' + str(deleted) + ' rutas entrantes y se aplicó la configuración.', 'eliminadas': deleted, 'respaldo': backup_path})
    except ValueError as exc:
        return jsonify({'ok': False, 'mensaje': str(exc), 'respaldo': backup_path}), 400
    except Exception as exc:
        logger.exception('[RUTEO EXTERNO V5] Error eliminando DID masivos')
        detalle = str(exc)[:900]
        if backup_path and perfil:
            detalle = _vp_ruteo_v5_restaurar_si_falla(perfil, backup_path, detalle)
        return jsonify({'ok': False, 'mensaje': detalle, 'respaldo': backup_path}), 500


# VP_RUTEO_EXTERNO_RANGOS_COS_V6

def _vp_ruteo_v6_sqlrun(perfil, sql, timeout=90):
    ok, salida, error = _vp_trunk_run_sql(perfil, sql, timeout=timeout)
    salida = str(salida or '')
    if (not ok) or re.search(r'(^|\n)ERROR\s+\d+|Unknown column|doesn.t exist|SQL syntax|Duplicate entry|Access denied|Table .* doesn.t exist', salida, re.I):
        raise RuntimeError(str(error or salida or 'Error SQL en Issabel.')[-900:])
    return [x for x in salida.splitlines() if x.strip()]


def _vp_ruteo_v6_schema(perfil):
    rows = _vp_ruteo_v6_sqlrun(
        perfil,
        "SELECT table_name,column_name,is_nullable,COALESCE(column_default,'__VP_NULL__'),extra "
        "FROM information_schema.columns WHERE table_schema=DATABASE() AND table_name IN "
        "('incoming','outbound_routes','outbound_route_sequence','customcontexts_contexts',"
        "'customcontexts_includes','customcontexts_includes_list','outbound_route_patterns','outbound_route_trunks','trunks','devices','sip','iax','zap','dahdi') ORDER BY table_name,ordinal_position;"
    )
    cols = {}
    meta = {}
    for line in rows:
        p = line.split('\t')
        if len(p) >= 5:
            table = p[0].lower()
            col = p[1].lower()
            cols.setdefault(table, set()).add(col)
            meta[(table, col)] = {'nullable': p[2].upper(), 'default': p[3], 'extra': p[4].lower()}
    return cols, meta


def _vp_ruteo_v6_backup(perfil, tablas):
    permitidas = set(['incoming', 'customcontexts_contexts', 'customcontexts_includes', 'devices', 'sip', 'iax', 'zap', 'dahdi'])
    tablas = sorted(set(str(t) for t in tablas))
    if not tablas or any(t not in permitidas for t in tablas):
        raise RuntimeError('Lista de tablas de respaldo no permitida.')
    destino = '/var/backups/voip-provisioner/ruteo-externo/ruteo_v6_' + time.strftime('%Y%m%d_%H%M%S') + '_' + str(os.getpid()) + '.sql'
    comando = r'''set -eu
AMPDBUSER="$(grep '^AMPDBUSER=' /etc/amportal.conf | head -1 | cut -d= -f2-)"
AMPDBPASS="$(grep '^AMPDBPASS=' /etc/amportal.conf | head -1 | cut -d= -f2-)"
AMPDBHOST="$(grep '^AMPDBHOST=' /etc/amportal.conf | head -1 | cut -d= -f2-)"
AMPDBNAME="$(grep '^AMPDBNAME=' /etc/amportal.conf | head -1 | cut -d= -f2-)"
[ -n "$AMPDBUSER" ] && [ -n "$AMPDBPASS" ] && [ -n "$AMPDBNAME" ] || { echo 'No se pudieron leer credenciales AMPDB'; exit 20; }
command -v mysqldump >/dev/null 2>&1 || { echo 'No se encuentra mysqldump'; exit 21; }
mkdir -p /var/backups/voip-provisioner/ruteo-externo
chmod 700 /var/backups/voip-provisioner /var/backups/voip-provisioner/ruteo-externo
TMP_CNF="$(mktemp)"
trap 'rm -f "$TMP_CNF"' EXIT
printf '[client]\nuser=%s\npassword=%s\n' "$AMPDBUSER" "$AMPDBPASS" > "$TMP_CNF"
[ -z "$AMPDBHOST" ] || printf 'host=%s\n' "$AMPDBHOST" >> "$TMP_CNF"
chmod 600 "$TMP_CNF"
DEST="''' + destino + r'''"
TABLES="''' + ' '.join(tablas) + r'''"
mysqldump --defaults-extra-file="$TMP_CNF" --add-drop-table --skip-lock-tables "$AMPDBNAME" $TABLES > "$DEST"
[ -s "$DEST" ] || { echo 'El respaldo SQL quedó vacío'; rm -f "$DEST"; exit 22; }
chmod 600 "$DEST"
echo "VP_RUTEO_V6_BACKUP_OK:$DEST"
'''
    out, err, code = _vp_trunk_exec(perfil, comando, timeout=120)
    out = str(out or '')
    match = re.search(r'VP_RUTEO_V6_BACKUP_OK:(/[^\r\n]+)', out)
    if code != 0 or not match:
        raise RuntimeError('No se pudo generar el respaldo SQL previo. ' + str(err or out or '')[-600:])
    return match.group(1).strip()


def _vp_ruteo_v6_restore(perfil, backup_path):
    if not re.match(r'^/var/backups/voip-provisioner/ruteo-externo/ruteo_v6_\d{8}_\d{6}_\d+\.sql$', str(backup_path or '')):
        raise RuntimeError('La ruta del respaldo V6 no pasó la validación de seguridad.')
    comando = r'''set -eu
AMPDBUSER="$(grep '^AMPDBUSER=' /etc/amportal.conf | head -1 | cut -d= -f2-)"
AMPDBPASS="$(grep '^AMPDBPASS=' /etc/amportal.conf | head -1 | cut -d= -f2-)"
AMPDBHOST="$(grep '^AMPDBHOST=' /etc/amportal.conf | head -1 | cut -d= -f2-)"
AMPDBNAME="$(grep '^AMPDBNAME=' /etc/amportal.conf | head -1 | cut -d= -f2-)"
TMP_CNF="$(mktemp)"; trap 'rm -f "$TMP_CNF"' EXIT
printf '[client]\nuser=%s\npassword=%s\n' "$AMPDBUSER" "$AMPDBPASS" > "$TMP_CNF"
[ -z "$AMPDBHOST" ] || printf 'host=%s\n' "$AMPDBHOST" >> "$TMP_CNF"
chmod 600 "$TMP_CNF"
BACKUP="''' + backup_path + r'''"
[ -s "$BACKUP" ] || { echo 'No se encuentra el respaldo SQL'; exit 41; }
mysql --defaults-extra-file="$TMP_CNF" "$AMPDBNAME" < "$BACKUP"
echo VP_RUTEO_V6_RESTORE_OK
'''
    out, err, code = _vp_trunk_exec(perfil, comando, timeout=120)
    if code != 0 or 'VP_RUTEO_V6_RESTORE_OK' not in str(out or ''):
        raise RuntimeError('No se pudo restaurar el respaldo V6. ' + str(err or out or '')[-500:])


def _vp_ruteo_v6_apply(perfil):
    comando = "/var/lib/asterisk/bin/retrieve_conf 2>&1 || { echo VP_RUTEO_V6_RETRIEVE_ERROR; exit 31; }; asterisk -rx 'dialplan reload' 2>&1 || { echo VP_RUTEO_V6_DIALPLAN_ERROR; exit 32; }; echo VP_RUTEO_V6_APPLY_OK"
    out, err, code = _vp_trunk_exec(perfil, comando, timeout=120)
    if code != 0 or 'VP_RUTEO_V6_APPLY_OK' not in str(out or ''):
        raise RuntimeError('Issabel no confirmó retrieve_conf y dialplan reload. ' + str(err or out or '')[-600:])


def _vp_ruteo_v6_restaurar_aplicar(perfil, backup_path, motivo):
    try:
        _vp_ruteo_v6_restore(perfil, backup_path)
        _vp_ruteo_v6_apply(perfil)
        return str(motivo) + ' Se restauró el respaldo SQL y se recargó la configuración anterior.'
    except Exception as exc:
        return str(motivo) + ' ATENCIÓN: falló la restauración automática: ' + str(exc)[:400] + '. Respaldo manual: ' + str(backup_path)


def _vp_ruteo_v6_cos_requerido(cols):
    tablas_contexto = ('sip', 'iax', 'zap', 'dahdi')
    tiene_contexto = any(set(['id', 'keyword', 'data']).issubset(cols.get(t, set())) for t in tablas_contexto)
    return (
        set(['context', 'description']).issubset(cols.get('customcontexts_contexts', set())) and
        set(['context', 'include']).issubset(cols.get('customcontexts_includes', set())) and
        set(['id', 'tech']).issubset(cols.get('devices', set())) and
        set(['route_id', 'name']).issubset(cols.get('outbound_routes', set())) and
        tiene_contexto
    )


@main_bp.route('/api/monitoreo/ruteo-externo/entradas/crear-rango', methods=['POST'])
def api_monitoreo_ruteo_externo_entradas_crear_rango_v6():
    'Crea DID consecutivos y los asigna a extensiones consecutivas, uno a uno.'
    backup_path = ''
    perfil = None
    try:
        data = request.get_json(silent=True) or {}
        if not _vp_trunk_admin_password(data):
            return jsonify({'ok': False, 'mensaje': 'Contraseña de administrador incorrecta.'}), 401
        perfil_nombre = str(data.get('perfil_nombre') or '').strip()
        perfil = buscar_perfil(perfil_nombre, cargar_perfiles()) if perfil_nombre else None
        if not perfil:
            return jsonify({'ok': False, 'mensaje': 'Seleccioná un servidor válido.'}), 400

        did_ini = str(data.get('did_desde') or '').strip()
        did_fin = str(data.get('did_hasta') or '').strip()
        int_ini = str(data.get('interno_desde') or '').strip()
        int_fin = str(data.get('interno_hasta') or '').strip()
        prefijo = str(data.get('prefijo_did') or '').strip()
        cid = str(data.get('caller_id') or '').strip()
        descripcion = str(data.get('descripcion') or '').strip()
        if not re.match(r'^\d{1,20}$', did_ini) or not re.match(r'^\d{1,20}$', did_fin):
            raise ValueError('Los dos extremos del rango DID deben ser numéricos y tener entre 1 y 20 dígitos.')
        if len(did_ini) != len(did_fin):
            raise ValueError('El DID inicial y final deben tener el mismo ancho para conservar ceros a la izquierda.')
        if int(did_fin) < int(did_ini):
            raise ValueError('El DID final debe ser mayor o igual al inicial.')
        if not re.match(r'^\d{1,20}$', int_ini) or not re.match(r'^\d{1,20}$', int_fin):
            raise ValueError('Los dos extremos del rango de internos deben ser numéricos.')
        if len(int_ini) != len(int_fin):
            raise ValueError('El interno inicial y final deben tener el mismo ancho.')
        if int(int_fin) < int(int_ini):
            raise ValueError('El interno final debe ser mayor o igual al inicial.')
        total_did = int(did_fin) - int(did_ini) + 1
        total_int = int(int_fin) - int(int_ini) + 1
        if total_did != total_int:
            raise ValueError('Los dos rangos deben tener la misma cantidad de números. Ejemplo: 0500–0599 mapea 100 números a 1000–1099.')
        if total_did > 500:
            raise ValueError('El límite es de 500 entradas por operación. Dividí el rango en lotes menores.')
        if prefijo and not re.match(r'^\d{1,30}$', prefijo):
            raise ValueError('El prefijo del DID debe contener solo dígitos (máximo 30).')
        if len(prefijo) + len(did_ini) > 80:
            raise ValueError('El DID completo supera el largo admitido.')
        if cid and (len(cid) > 80 or not re.match(r'^[0-9A-Za-z_*#+.\-]+$', cid)):
            raise ValueError('El CID de origen contiene caracteres no admitidos.')
        if len(descripcion) > 100 or any(ord(c) < 32 for c in descripcion):
            raise ValueError('La descripción no puede superar 100 caracteres ni contener controles.')

        schema_cols, schema_meta = _vp_ruteo_v6_schema(perfil)
        cols = schema_cols.get('incoming', set())
        meta = {col: schema_meta.get(('incoming', col), {}) for col in cols}
        q = _vp_trunk_sql_q
        rows = []
        for offset in range(total_did):
            did_suffix = str(int(did_ini) + offset).zfill(len(did_ini))
            did = prefijo + did_suffix
            interno = str(int(int_ini) + offset).zfill(len(int_ini))
            desc = (descripcion + ' ' if descripcion else 'DID ') + did_suffix + ' -> interno ' + interno
            rows.append({'did': did, 'caller_id': cid, 'descripcion': desc, 'destino': 'ext-local,' + interno + ',1'})

        if 'extension' not in cols or 'cidnum' not in cols or 'destination' not in cols:
            raise RuntimeError('No se detectan las columnas necesarias de la tabla incoming.')
        ext_in = ','.join(q(x['did']) for x in rows)
        dup = _vp_ruteo_v6_sqlrun(perfil,
            "SELECT extension FROM incoming WHERE extension IN (" + ext_in + ") AND COALESCE(cidnum,'')=" + q(cid) + " LIMIT 1;")
        if dup:
            raise ValueError('Ya existe una entrada con el DID ' + str(dup[0].split('\t')[0]) + ' y el mismo CID. No se creó nada.')

        vals_map = {'extension': '', 'cidnum': '', 'description': '', 'destination': ''}
        selected = [c for c in ('extension', 'cidnum', 'description', 'destination') if c in cols]
        provided = set(selected)
        missing = []
        for col in cols:
            info = meta.get(col, {})
            if info.get('nullable') == 'NO' and info.get('default') == '__VP_NULL__' and 'auto_increment' not in info.get('extra', '') and col not in provided:
                missing.append(col)
        if missing:
            raise RuntimeError('La tabla incoming exige campos adicionales sin valor predeterminado: ' + ', '.join(sorted(missing)) + '. No se modificó nada.')
        col_sql = ','.join('`' + c + '`' for c in selected)
        tuples = []
        for item in rows:
            row_vals = {'extension': item['did'], 'cidnum': item['caller_id'], 'description': item['descripcion'], 'destination': item['destino']}
            tuples.append('(' + ','.join(q(row_vals[c]) for c in selected) + ')')

        backup_path = _vp_ruteo_v6_backup(perfil, ['incoming'])
        sql = 'INSERT INTO `incoming` (' + col_sql + ') VALUES ' + ','.join(tuples) + "; SELECT CONCAT('VP_RUTEO_V6_RANGO_CREADO|',ROW_COUNT());"
        out = _vp_ruteo_v6_sqlrun(perfil, sql, timeout=120)
        confirmed = []
        for line in out:
            match = re.search(r'VP_RUTEO_V6_RANGO_CREADO\|(\d+)', line)
            if match:
                confirmed.append(int(match.group(1)))
        if not confirmed or confirmed[-1] != len(rows):
            msg = _vp_ruteo_v6_restaurar_aplicar(perfil, backup_path, 'Issabel no confirmó todas las entradas del rango.')
            backup_path = ''
            raise RuntimeError(msg)
        try:
            _vp_ruteo_v6_apply(perfil)
        except Exception as exc:
            msg = _vp_ruteo_v6_restaurar_aplicar(perfil, backup_path, 'Falló la aplicación del rango: ' + str(exc)[:300])
            backup_path = ''
            raise RuntimeError(msg)
        return jsonify({'ok': True, 'mensaje': 'Se crearon ' + str(len(rows)) + ' rutas DID y se mapearon a los internos ' + int_ini + '–' + int_fin + '.', 'creadas': len(rows), 'respaldo': backup_path})
    except ValueError as exc:
        return jsonify({'ok': False, 'mensaje': str(exc), 'respaldo': backup_path}), 400
    except Exception as exc:
        logger.exception('[RUTEO EXTERNO V6] Error creando rango DID')
        detalle = str(exc)[:900]
        if backup_path and perfil:
            detalle = _vp_ruteo_v6_restaurar_aplicar(perfil, backup_path, detalle)
        return jsonify({'ok': False, 'mensaje': detalle, 'respaldo': backup_path}), 500


# VP_RUTEO_EXTERNO_PJSIP_INTL_V8

def _vp_ruteo_v8_remote(perfil, payload, timeout=90):
    """Lee o actualiza archivos personalizados remotos sin mostrar secretos."""
    import base64 as _b64
    import json as _json
    remote = r"""
import os, re, sys, json, base64, shutil, tempfile, time, subprocess
P = json.loads(base64.b64decode('__VP_RUTEO_PAYLOAD_B64__').decode('utf-8'))
BASE = '/etc/asterisk'
PJSIP_BEGIN = '; BEGIN VP_RUTEO_COS_PJSIP_V8'
PJSIP_END = '; END VP_RUTEO_COS_PJSIP_V8'
INTL_BEGIN = '; BEGIN VP_RUTEO_INTL_V8 '
INTL_END = '; END VP_RUTEO_INTL_V8 '

def emit(obj, code=0):
    raw = json.dumps(obj, ensure_ascii=True, separators=(',', ':')).encode('utf-8')
    print('VP_RUTEO_V8_RESULT:' + base64.b64encode(raw).decode('ascii'))
    sys.exit(code)

def read_text(path):
    try:
        with open(path, 'r') as f:
            return f.read()
    except IOError:
        return ''

def config_includes(path):
    found = []
    lines = read_text(path).splitlines()
    for n, line in enumerate(lines):
        m = re.match(r'^\s*#(?:try)?include\s+["<]?([^">\s]+)', line, re.I)
        if not m:
            continue
        target = m.group(1).strip()
        if not target.startswith('/'):
            target = os.path.join(BASE, target)
        target = os.path.normpath(target)
        found.append((n, os.path.basename(target), target))
    return found

def discover_post_file():
    conf = os.path.join(BASE, 'pjsip.conf')
    if not os.path.isfile(conf):
        return '', 'No existe /etc/asterisk/pjsip.conf.'
    inc = config_includes(conf)
    anchors = [n for n, name, path in inc if name in ('pjsip_additional.conf', 'pjsip.endpoint_additional.conf', 'pjsip.endpoint.conf')]
    if not anchors:
        return '', 'No pude confirmar el include de la configuración PJSIP generada.'
    anchor = max(anchors)
    candidates = [(n, path) for n, name, path in inc
                  if n > anchor and name in ('pjsip.endpoint_custom_post.conf', 'pjsip_custom_post.conf')]
    if not candidates:
        return '', 'No hay un archivo custom-post incluido después de la configuración PJSIP generada. No se asignó contexto PJSIP.'
    n, path = candidates[-1]
    real = os.path.realpath(path)
    if not real.startswith(BASE + os.sep) or os.path.dirname(os.path.normpath(path)) != BASE:
        return '', 'El archivo PJSIP custom-post debe estar directamente en /etc/asterisk; se canceló por seguridad.'
    return path, ''

def strip_global_block(content, begin, end):
    lines = content.splitlines(True)
    begins = [i for i, x in enumerate(lines) if x.strip() == begin]
    ends = [i for i, x in enumerate(lines) if x.strip() == end]
    if len(begins) != len(ends) or len(begins) > 1:
        raise RuntimeError('Los marcadores administrados V8 están incompletos o repetidos en un archivo de configuración.')
    entries = {}
    if not begins:
        return content, entries
    a, b = begins[0], ends[0]
    if b < a:
        raise RuntimeError('El orden de los marcadores administrados V8 es inválido.')
    for line in lines[a + 1:b]:
        m = re.match(r'^\s*;\s*VP-RUTEO-PJSIP\|([0-9]+)\|([A-Za-z0-9_.-]{1,80})\|([A-Za-z0-9_.-]{1,80})\s*$', line)
        if m:
            entries[m.group(1)] = {'original_context': m.group(2), 'contexto': m.group(3)}
    outside = ''.join(lines[:a] + lines[b + 1:])
    return outside, entries

def strip_intl_blocks(content):
    lines = content.splitlines(True)
    kept = []
    contexts = set()
    i = 0
    while i < len(lines):
        m = re.match(r'^\s*; BEGIN VP_RUTEO_INTL_V8 ([A-Za-z][A-Za-z0-9_]{1,39})\s*$', lines[i].rstrip('\r\n'))
        if not m:
            kept.append(lines[i]); i += 1; continue
        ctx = m.group(1)
        endmark = INTL_END + ctx
        j = i + 1
        while j < len(lines) and lines[j].strip() != endmark:
            j += 1
        if j >= len(lines):
            raise RuntimeError('Falta el cierre de un bloque de bloqueo internacional V8 para ' + ctx + '.')
        contexts.add(ctx)
        i = j + 1
    return ''.join(kept), contexts

def current_context(ext):
    try:
        p = subprocess.Popen(['asterisk', '-rx', 'pjsip show endpoint ' + ext], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        out, err = p.communicate(timeout=15)
        if not isinstance(out, str): out = out.decode('utf-8', 'replace')
        m = re.search(r'^\s*context\s*:\s*([A-Za-z0-9_.-]{1,80})\s*$', out, re.I | re.M)
        if m and 'Unable to find object' not in out and 'No objects found' not in out:
            return m.group(1)
    except Exception:
        pass
    return ''

def context_has_intl_rule(ctx):
    if not re.match(r'^[A-Za-z][A-Za-z0-9_]{1,39}$', ctx):
        return False
    try:
        p = subprocess.Popen(['asterisk', '-rx', 'dialplan show ' + ctx], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        out, err = p.communicate(timeout=15)
        if not isinstance(out, str): out = out.decode('utf-8', 'replace')
        return bool(re.search(r"['_\"]_900(?:\.|[XZN\[])", out, re.I))
    except Exception:
        return False

def pjsip_include_custom_loaded():
    conf = read_text(os.path.join(BASE, 'extensions.conf'))
    return bool(re.search(r'^\s*#(?:try)?include\s+["<]?[^">\s]*extensions_custom\.conf', conf, re.I | re.M))

def probe():
    post_file, post_error = discover_post_file()
    post_text = read_text(post_file) if post_file else ''
    post_base, pmap = strip_global_block(post_text, PJSIP_BEGIN, PJSIP_END)
    ext_file = os.path.join(BASE, 'extensions_custom.conf')
    ext_text = read_text(ext_file)
    ext_base, blocked = strip_intl_blocks(ext_text)
    exts = []
    for x in P.get('extensiones', []):
        x = str(x)
        if re.match(r'^\d{1,20}$', x):
            exts.append(x)
    contexts = dict((x, current_context(x)) for x in exts)
    checks = []
    for x in P.get('verificar_bloqueo', []):
        x = str(x)
        if re.match(r'^[A-Za-z][A-Za-z0-9_]{1,39}$', x):
            checks.append(x)
    active_blocked = {}
    for x in sorted(set(list(blocked) + checks)):
        active_blocked[x] = context_has_intl_rule(x)
    return {'ok': True, 'pjsip_supported': bool(post_file), 'pjsip_post_file': post_file,
            'pjsip_error': post_error, 'pjsip_map': pmap, 'contextos_activos': contexts,
            'clases_bloqueo': sorted(blocked), 'bloqueo_activo': active_blocked,
            'extensions_custom_cargado': pjsip_include_custom_loaded(),
            'extensions_custom_existe': os.path.isfile(ext_file)}

def atomic_write(path, content):
    folder = os.path.dirname(path)
    oldst = None
    if os.path.exists(path): oldst = os.stat(path)
    fd, tmp = tempfile.mkstemp(prefix='.vp_ruteo_v8_', dir=folder)
    try:
        with os.fdopen(fd, 'w') as f:
            f.write(content)
            f.flush()
            os.fsync(f.fileno())
        if oldst:
            os.chmod(tmp, oldst.st_mode & 0o7777)
            try: os.chown(tmp, oldst.st_uid, oldst.st_gid)
            except Exception: pass
        else:
            os.chmod(tmp, 0o640)
            try: os.chown(tmp, 0, 0)
            except Exception: pass
        os.rename(tmp, path)
    finally:
        if os.path.exists(tmp):
            try: os.unlink(tmp)
            except Exception: pass

def make_backup(path, existed):
    if not existed:
        return {'path': path, 'existed': False, 'backup': ''}
    backup = path + '.bak_vp_ruteo_cos_v8_' + time.strftime('%Y%m%d_%H%M%S') + '_' + str(os.getpid())
    shutil.copy2(path, backup)
    return {'path': path, 'existed': True, 'backup': backup}

def validate_paths(path):
    real = os.path.realpath(path)
    if not real.startswith(BASE + os.sep):
        raise RuntimeError('Se rechazó una ruta de configuración fuera de /etc/asterisk.')

def apply_files():
    post_file, post_error = discover_post_file()
    if not post_file and (P.get('pjsip_map') or P.get('pjsip_map_actual')):
        raise RuntimeError(post_error or 'No se encontró el archivo PJSIP custom-post cargado.')
    ext_file = os.path.join(BASE, 'extensions_custom.conf')
    if P.get('clases_bloqueo') and not pjsip_include_custom_loaded():
        raise RuntimeError('extensions.conf no incluye extensions_custom.conf; no se instaló el bloqueo internacional.')
    if not os.path.isfile(ext_file) and P.get('clases_bloqueo'):
        raise RuntimeError('No existe extensions_custom.conf; no se creó un archivo no incluido por Asterisk.')

    post_text = read_text(post_file) if post_file else ''
    post_base, old_map = strip_global_block(post_text, PJSIP_BEGIN, PJSIP_END)
    new_map = {}
    for item in P.get('pjsip_map', []):
        if not isinstance(item, dict):
            raise RuntimeError('Formato inválido de asignación PJSIP.')
        ext = str(item.get('interno') or '')
        orig = str(item.get('original_context') or '')
        ctx = str(item.get('contexto') or '')
        if not re.match(r'^\d{1,20}$', ext) or not re.match(r'^[A-Za-z0-9_.-]{1,80}$', orig) or not re.match(r'^[A-Za-z][A-Za-z0-9_]{1,39}$', ctx):
            raise RuntimeError('Una asignación PJSIP no pasó la validación de seguridad.')
        new_map[ext] = {'original_context': orig, 'contexto': ctx}
    if new_map and not post_file:
        raise RuntimeError(post_error or 'No se puede guardar el contexto PJSIP de forma persistente.')

    ext_text = read_text(ext_file)
    ext_base, old_blocked = strip_intl_blocks(ext_text)
    wanted_blocked = set()
    for x in P.get('clases_bloqueo', []):
        x = str(x)
        if not re.match(r'^[A-Za-z][A-Za-z0-9_]{1,39}$', x):
            raise RuntimeError('Nombre de clase no válido para el bloqueo internacional.')
        wanted_blocked.add(x)
    for ctx in sorted(wanted_blocked - old_blocked):
        shown = ''
        try:
            p = subprocess.Popen(['asterisk', '-rx', 'dialplan show ' + ctx], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            out, err = p.communicate(timeout=15)
            if not isinstance(out, str): out = out.decode('utf-8', 'replace')
            shown = out
        except Exception:
            shown = ''
        if re.search(r"['_\"]_900(?:\.|[XZN\[])", shown, re.I):
            raise RuntimeError('El contexto ' + ctx + ' ya contiene una regla que coincide con 900. No se duplicó ni sobrescribió esa regla.')

    def render_map(base, mapping):
        if not mapping:
            return base
        block = [PJSIP_BEGIN]
        for ext in sorted(mapping, key=lambda k: (len(k), k)):
            v = mapping[ext]
            block.append('; VP-RUTEO-PJSIP|' + ext + '|' + v['original_context'] + '|' + v['contexto'])
            block.append('[' + ext + '](+)')
            block.append('context=' + v['contexto'])
        block.append(PJSIP_END)
        result = base
        if result and not result.endswith('\n'): result += '\n'
        if result: result += '\n'
        return result + '\n'.join(block) + '\n'

    def render_blocks(base, contexts):
        if not contexts:
            return base
        result = base
        if result and not result.endswith('\n'): result += '\n'
        if result: result += '\n'
        blocks = []
        for ctx in sorted(contexts):
            blocks.extend([INTL_BEGIN + ctx, '[' + ctx + '](+)',
                           'exten => _900.,1,NoOp(VoIP Provisioner: bloqueo internacional para ' + ctx + ')',
                           ' same => n,Playback(ss-noservice)',
                           ' same => n,Hangup(21)', INTL_END + ctx, ''])
        return result + '\n'.join(blocks)

    new_post_text = render_map(post_base, new_map) if post_file else post_text
    new_ext_text = render_blocks(ext_base, wanted_blocked)
    changes = []
    if post_file and new_post_text != post_text:
        validate_paths(post_file)
        changes.append((post_file, new_post_text))
    if new_ext_text != ext_text:
        validate_paths(ext_file)
        changes.append((ext_file, new_ext_text))

    backups = []
    snapshots = []
    try:
        for path, content in changes:
            existed = os.path.isfile(path)
            b = make_backup(path, existed)
            backups.append(b)
            snapshots.append((path, existed, b.get('backup', '')))
            atomic_write(path, content)
    except Exception:
        for path, existed, backup in reversed(snapshots):
            try:
                if existed and backup and os.path.isfile(backup): shutil.copy2(backup, path)
                elif not existed and os.path.exists(path): os.unlink(path)
            except Exception:
                pass
        raise
    return {'ok': True, 'file_backups': backups, 'pjsip_post_file': post_file,
            'pjsip_map_changed': bool(post_file and new_post_text != post_text),
            'bloqueo_clases': sorted(wanted_blocked)}

def restore_files():
    backups = P.get('file_backups') or []
    for item in reversed(backups):
        if not isinstance(item, dict): raise RuntimeError('Formato de respaldo de archivo inválido.')
        path = str(item.get('path') or '')
        existed = bool(item.get('existed'))
        backup = str(item.get('backup') or '')
        if path not in (os.path.join(BASE, 'extensions_custom.conf'),
                        os.path.join(BASE, 'pjsip.endpoint_custom_post.conf'),
                        os.path.join(BASE, 'pjsip_custom_post.conf')):
            raise RuntimeError('Se rechazó restaurar una ruta de archivo no permitida.')
        validate_paths(path)
        if existed:
            if backup != path + '.bak_vp_ruteo_cos_v8_' + os.path.basename(backup).split('_')[-2] + '_' + os.path.basename(backup).split('_')[-1] and not backup.startswith(path + '.bak_vp_ruteo_cos_v8_'):
                raise RuntimeError('La ruta del respaldo no pasó la validación.')
            if not backup.startswith(path + '.bak_vp_ruteo_cos_v8_') or not os.path.isfile(backup):
                raise RuntimeError('No se encuentra el respaldo de ' + path + '.')
            shutil.copy2(backup, path)
        else:
            if os.path.exists(path): os.unlink(path)
    return {'ok': True, 'restaurados': len(backups)}

try:
    mode = str(P.get('mode') or 'probe')
    if mode == 'probe': emit(probe())
    elif mode == 'apply': emit(apply_files())
    elif mode == 'restore': emit(restore_files())
    else: emit({'ok': False, 'mensaje': 'Modo remoto no permitido.'}, 2)
except Exception as e:
    emit({'ok': False, 'mensaje': str(e)[:800]}, 2)
"""
    payload_b64 = _b64.b64encode(_json.dumps(payload, separators=(',', ':')).encode('utf-8')).decode('ascii')
    remote = remote.replace('__VP_RUTEO_PAYLOAD_B64__', payload_b64)
    code_b64 = _b64.b64encode(remote.encode('utf-8')).decode('ascii')
    command = "python3 -c \"import base64; exec(compile(base64.b64decode('%s'), '<vp-ruteo-v8>', 'exec'))\"" % code_b64
    out, err, rc = _vp_trunk_exec(perfil, command, timeout=timeout)
    out = str(out or '')
    match = re.search(r'VP_RUTEO_V8_RESULT:([A-Za-z0-9+/=]+)', out)
    if not match:
        raise RuntimeError('No se pudo consultar/actualizar el soporte PJSIP/COS remoto. ' + str(err or out or '')[-650:])
    try:
        result = _json.loads(_b64.b64decode(match.group(1)).decode('utf-8'))
    except Exception:
        raise RuntimeError('La respuesta del diagnóstico remoto no es JSON válido.')
    if not result.get('ok'):
        raise RuntimeError(str(result.get('mensaje') or 'Falló la operación remota.')[:800])
    return result


# VP_TRONCALES_EXISTENTES_API_V13
@main_bp.route('/api/monitoreo/ruteo-externo/troncales-configuradas', methods=['GET'])
def api_monitoreo_troncales_configuradas_v13():
    """Lectura de todas las tecnologías de troncal de Issabel; no devuelve secretos."""
    try:
        perfil_nombre = str(request.args.get('perfil_nombre') or '').strip()
        perfil = buscar_perfil(perfil_nombre, cargar_perfiles()) if perfil_nombre else None
        if not perfil:
            return jsonify({'ok': False, 'mensaje': 'Seleccioná un servidor válido.'}), 400

        meta_rows = _vp_ruteo_v6_sqlrun(
            perfil,
            "SELECT table_name,column_name FROM information_schema.columns "
            "WHERE table_schema=DATABASE() AND table_name IN "
            "('trunks','outbound_routes','outbound_route_trunks','outbound_route_patterns') "
            "ORDER BY table_name,ordinal_position;"
        )
        cols = {}
        for line in meta_rows:
            p = line.split('\t')
            if len(p) >= 2:
                cols.setdefault(p[0].strip().lower(), set()).add(p[1].strip().lower())

        trunk_cols = cols.get('trunks', set())
        if not set(['trunkid', 'name', 'tech']).issubset(trunk_cols):
            return jsonify({
                'ok': False,
                'mensaje': 'La tabla trunks no contiene las columnas obligatorias trunkid, name y tech.'
            }), 409

        def col(alias, name):
            if name in trunk_cols:
                return "COALESCE(CAST(" + alias + ".`" + name + "` AS CHAR),'')"
            return "''"

        sql = (
            "SELECT " + col('t', 'trunkid') + "," + col('t', 'name') + "," +
            col('t', 'tech') + "," + col('t', 'channelid') + "," +
            col('t', 'disabled') + "," + col('t', 'usercontext') +
            " FROM `trunks` t ORDER BY t.`trunkid`;"
        )
        trunk_rows = _vp_ruteo_v6_sqlrun(perfil, sql)

        labels = {
            'iax': 'IAX2', 'iax2': 'IAX2', 'sip': 'SIP', 'pjsip': 'PJSIP',
            'dahdi': 'DAHDI', 'zap': 'ZAP/DAHDI', 'ooh323': 'OOH323',
            'h323': 'H.323', 'custom': 'CUSTOM', 'dongle': 'DONGLE/GSM',
            'gsm': 'GSM', 'khomp': 'KHOMP', 'misdn': 'MISDN',
            'dundi': 'DUNDi', 'enum': 'ENUM'
        }

        def enabled(value):
            return str(value or '').strip().lower() in ('1', 'on', 'yes', 'true', 'enabled', 'si', 'sí')

        inventory = []
        by_id = {}
        disabled_known = 'disabled' in trunk_cols
        for line in trunk_rows:
            p = line.split('\t')
            if len(p) < 3 or not p[0].strip().isdigit():
                continue
            tid = p[0].strip()
            tech = p[2].strip().lower()
            channel = p[3].strip() if len(p) > 3 else ''
            raw_disabled = p[4].strip() if len(p) > 4 else ''
            # Issabel guarda OOH323 como tech=custom; distinguirlo por channelid.
            if tech == 'custom' and channel.upper().startswith('OOH323/'):
                label = 'OOH323'
            else:
                label = labels.get(tech, tech.upper() if tech else 'NO INFORMADA')
            item = {
                'id': tid,
                'nombre': p[1].strip() or ('Troncal ' + tid),
                'tecnologia': label,
                'tecnologia_original': tech,
                'canal': channel,
                'contexto_entrada': p[5].strip() if len(p) > 5 else '',
                'deshabilitada': enabled(raw_disabled) if disabled_known else None,
                'estado_configuracion': ('Deshabilitada' if enabled(raw_disabled) else 'Habilitada') if disabled_known else 'No informado',
                'rutas': [],
                'patrones': []
            }
            inventory.append(item)
            by_id[tid] = item

        route_cols = cols.get('outbound_routes', set())
        route_names = {}
        if set(['route_id', 'name']).issubset(route_cols):
            route_sql = "SELECT CAST(`route_id` AS CHAR),COALESCE(`name`,'') FROM `outbound_routes` ORDER BY `route_id`;"
            for line in _vp_ruteo_v6_sqlrun(perfil, route_sql):
                p = line.split('\t')
                if len(p) >= 2 and p[0].strip().isdigit():
                    route_names[p[0].strip()] = p[1].strip()

        pattern_cols = cols.get('outbound_route_patterns', set())
        pattern_supported = set(['route_id', 'match_pattern_prefix', 'match_pattern_pass']).issubset(pattern_cols)
        patterns_by_route = {}
        if pattern_supported:
            cid_expr = "COALESCE(`match_cid`,'')" if 'match_cid' in pattern_cols else "''"
            prepend_expr = "COALESCE(`prepend_digits`,'')" if 'prepend_digits' in pattern_cols else "''"
            pat_sql = "SELECT CAST(`route_id` AS CHAR),COALESCE(`match_pattern_prefix`,''),COALESCE(`match_pattern_pass`,'')," + cid_expr + "," + prepend_expr + " FROM `outbound_route_patterns` ORDER BY `route_id`;"
            for line in _vp_ruteo_v6_sqlrun(perfil, pat_sql):
                p = line.split('\t')
                if len(p) >= 3 and p[0].strip().isdigit():
                    rid = p[0].strip()
                    value = (p[1].strip() + p[2].strip()) or '(sin patrón)'
                    if len(p) > 3 and p[3].strip():
                        value += ' · CID ' + p[3].strip()
                    if len(p) > 4 and p[4].strip():
                        value += ' · antepone ' + p[4].strip()
                    patterns_by_route.setdefault(rid, []).append(value)

        assoc_cols = cols.get('outbound_route_trunks', set())
        assoc_supported = set(['route_id', 'trunk_id']).issubset(assoc_cols)
        if assoc_supported and route_names:
            seq_expr = "COALESCE(CAST(ort.`seq` AS CHAR),'')" if 'seq' in assoc_cols else "''"
            order_seq = 'ort.`seq`,' if 'seq' in assoc_cols else ''
            assoc_sql = "SELECT CAST(ort.`route_id` AS CHAR),CAST(ort.`trunk_id` AS CHAR)," + seq_expr + " FROM `outbound_route_trunks` ort ORDER BY ort.`route_id`," + order_seq + "ort.`trunk_id`;"
            for line in _vp_ruteo_v6_sqlrun(perfil, assoc_sql):
                p = line.split('\t')
                if len(p) < 2 or not p[0].strip().isdigit() or not p[1].strip().isdigit():
                    continue
                rid, tid = p[0].strip(), p[1].strip()
                item = by_id.get(tid)
                if item is None:
                    continue
                name = route_names.get(rid, '') or ('Ruta ' + rid)
                seq = p[2].strip() if len(p) > 2 else ''
                label = name + ' (ID ' + rid + (', orden ' + seq if seq else '') + ')'
                if label not in item['rutas']:
                    item['rutas'].append(label)
                for pat in patterns_by_route.get(rid, []):
                    tagged = name + ': ' + pat
                    if tagged not in item['patrones']:
                        item['patrones'].append(tagged)

        warnings = []
        if not assoc_supported:
            warnings.append('No está disponible la asociación ruta-troncal; se muestran igualmente todas las troncales configuradas.')
        if not pattern_supported:
            warnings.append('No se pudieron consultar los patrones de salida.')
        return jsonify({
            'ok': True,
            'perfil': perfil_nombre,
            'total': len(inventory),
            'troncales': inventory,
            'asociaciones_soportadas': assoc_supported and bool(route_names),
            'patrones_soportados': pattern_supported,
            'mensaje': ' '.join(warnings)
        })
    except Exception as exc:
        logger.exception('[TRONCALES V13] Error consultando catálogo completo')
        return jsonify({'ok': False, 'mensaje': 'No se pudo consultar la configuración de troncales: ' + str(exc)[:500]}), 500
# FIN VP_TRONCALES_EXISTENTES_API_V13


# VP_CALL_RESTRICTIONS_MANAGER_V1
# VP_CALL_RESTRICTIONS_HOOK_INTEGRATION_V1
# Administrador de bloqueos por prefijo, interno y troncal/ruta, independiente de COS.
def _vp_call_restrictions_catalog(perfil):
    cols, meta = _vp_ruteo_v6_schema(perfil)
    routes = []
    route_names = {}
    for line in _vp_ruteo_v6_sqlrun(perfil, "SELECT COALESCE(CAST(`route_id` AS CHAR),''),COALESCE(`name`,'') FROM `outbound_routes` ORDER BY `route_id`;"):
        p = line.split('\t')
        if len(p) >= 1 and p[0].strip().isdigit():
            rid = p[0].strip()
            name = p[1].strip() if len(p) > 1 else ''
            routes.append({'id': rid, 'nombre': name or ('Ruta ' + rid)})
            route_names[rid] = name or ('Ruta ' + rid)

    trunks = []
    trunk_ids = set()
    trunk_cols = cols.get('trunks', set())
    if set(['trunkid', 'name', 'tech']).issubset(trunk_cols):
        for line in _vp_ruteo_v6_sqlrun(perfil, "SELECT COALESCE(CAST(`trunkid` AS CHAR),''),COALESCE(`name`,''),LOWER(COALESCE(`tech`,'')) FROM `trunks` ORDER BY `trunkid`;"):
            p = line.split('\t')
            if len(p) >= 3 and p[0].strip().isdigit():
                tid = p[0].strip()
                trunk_ids.add(tid)
                trunks.append({'id': tid, 'nombre': p[1].strip() or ('Troncal ' + tid), 'tecnologia': p[2].strip()})

    route_trunk_cols = cols.get('outbound_route_trunks', set())
    association_supported = bool(set(['route_id', 'trunk_id']).issubset(route_trunk_cols) and trunk_ids)
    route_trunks = dict((rid, []) for rid in route_names)
    trunk_routes = dict((tid, []) for tid in trunk_ids)
    if association_supported:
        rows = _vp_ruteo_v6_sqlrun(perfil, "SELECT COALESCE(CAST(`route_id` AS CHAR),''),COALESCE(CAST(`trunk_id` AS CHAR),'') FROM `outbound_route_trunks` ORDER BY `route_id`,`trunk_id`;")
        for line in rows:
            p = line.split('\t')
            if len(p) < 2:
                continue
            rid, tid = p[0].strip(), p[1].strip()
            if rid in route_trunks and tid in trunk_ids:
                if tid not in route_trunks[rid]:
                    route_trunks[rid].append(tid)
                if rid not in trunk_routes.setdefault(tid, []):
                    trunk_routes[tid].append(rid)

    for route in routes:
        route['troncales_ids'] = route_trunks.get(route['id'], [])
        route['troncales_compartidas'] = sorted(set(
            tid for tid in route['troncales_ids']
            if any(other != route['id'] for other in trunk_routes.get(tid, []))
        ))
    return {
        'rutas': routes,
        'troncales': trunks,
        'route_names': route_names,
        'route_trunks': route_trunks,
        'trunk_routes': trunk_routes,
        'trunk_ids': trunk_ids,
        'association_supported': association_supported
    }


def _vp_call_restrictions_remote(perfil, payload, timeout=90):
    import base64 as _b64
    import json as _json
    remote = r'''
import os, re, sys, json, base64, shutil, tempfile, time, subprocess
P = json.loads(base64.b64decode('__VP_CALL_RESTRICTIONS_PAYLOAD__').decode('utf-8'))
BASE = '/etc/asterisk'
PATH = os.path.join(BASE, 'extensions_custom.conf')
BEGIN = '; BEGIN VP_CALL_RESTRICTIONS_V1'
END = '; END VP_CALL_RESTRICTIONS_V1'
DATA_PREFIX = '; VP_CALL_RESTRICTIONS_DATA '

def emit(obj, code=0):
    raw = json.dumps(obj, ensure_ascii=True, separators=(',', ':')).encode('utf-8')
    print('VP_CALL_RESTRICTIONS_RESULT:' + base64.b64encode(raw).decode('ascii'))
    sys.exit(code)

def read_text(path):
    try:
        with open(path, 'r') as f:
            return f.read()
    except IOError:
        return ''

def strip_block(content):
    lines = content.splitlines(True)
    begins = [i for i, x in enumerate(lines) if x.strip() == BEGIN]
    ends = [i for i, x in enumerate(lines) if x.strip() == END]
    if len(begins) != len(ends) or len(begins) > 1:
        raise RuntimeError('Los marcadores de restricciones están incompletos o repetidos.')
    if not begins:
        return content, '', False
    a, b = begins[0], ends[0]
    if b < a:
        raise RuntimeError('El orden de los marcadores de restricciones es inválido.')
    block = ''.join(lines[a:b + 1])
    base = ''.join(lines[:a] + lines[b + 1:])
    return base, block, True

def decode_rules(block):
    if not block:
        return []
    chunks = []
    for line in block.splitlines():
        line = line.strip()
        if line.startswith(DATA_PREFIX):
            chunks.append(line[len(DATA_PREFIX):].strip())
    if not chunks:
        return []
    raw = base64.b64decode(''.join(chunks).encode('ascii'))
    rules = json.loads(raw.decode('utf-8'))
    if not isinstance(rules, list):
        raise RuntimeError('El catálogo guardado de restricciones no es una lista válida.')
    return rules

def custom_include_ok():
    text = read_text(os.path.join(BASE, 'extensions.conf'))
    return bool(re.search(r'^\s*#(?:try)?include\s+["<]?[^">\s]*extensions_custom\.conf', text, re.I | re.M))

def hook_context_bounds(content):
    lines = content.splitlines(True)
    start = None
    end = len(lines)
    for i, line in enumerate(lines):
        m = re.match(r'^\s*\[([^\]]+)\]', line)
        if not m:
            continue
        name = m.group(1).strip().lower()
        if start is not None:
            end = i
            break
        if name == 'macro-dialout-trunk-predial-hook':
            start = i
    if start is None:
        return None, None, None
    anchor = None
    for i in range(start + 1, end):
        if re.match(r'^\s*exten\s*=>\s*s\s*,\s*1\s*,', lines[i], re.I):
            action = lines[i].split(',', 2)[-1]
            # No insertar después de una acción que pueda salir del hook antes
            # de evaluar las restricciones.
            if re.search(r'\b(?:MacroExit|Hangup|Return|Goto)\b', action, re.I):
                return start, end, None
            anchor = i
            break
    return start, end, anchor

def hook_context_supported(content):
    start, end, anchor = hook_context_bounds(content)
    return start is not None and anchor is not None

def insert_managed_block(content, block):
    lines = content.splitlines(True)
    start, end, anchor = hook_context_bounds(content)
    if start is None or anchor is None:
        raise RuntimeError('Se encontró un hook manual, pero no pude integrarlo con seguridad. El contexto [macro-dialout-trunk-predial-hook] debe tener una línea inicial ext en s, prioridad 1, que no termine el macro ni salte a otro destino. No se modificó el dialplan.')
    prefix = lines[:anchor + 1]
    if prefix and not prefix[-1].endswith('\n'):
        prefix[-1] += '\n'
    return ''.join(prefix) + block + ''.join(lines[anchor + 1:])

def safe_rules(rules):
    if not isinstance(rules, list) or len(rules) > 500:
        raise RuntimeError('La lista de restricciones es inválida o supera 500 reglas.')
    ids = set()
    out = []
    for r in rules:
        if not isinstance(r, dict):
            raise RuntimeError('Hay una restricción con formato inválido.')
        rid = str(r.get('id') or '')
        if not re.match(r'^[a-f0-9]{8,32}$', rid) or rid in ids:
            raise RuntimeError('Hay un identificador de restricción inválido o repetido.')
        ids.add(rid)
        exts = r.get('internos') or []
        prefs = r.get('prefijos') or []
        tids = r.get('troncales_ids') or []
        mode = str(r.get('modo') or 'bloquear_prefijos')
        if mode not in ('bloquear_prefijos', 'permitir_troncales', 'clase_servicio'):
            raise RuntimeError('El modo de una restricción no es válido.')
        if not exts or len(exts) > 1000:
            raise RuntimeError('Cada regla requiere internos válidos.')
        if mode == 'bloquear_prefijos' and (not prefs or len(prefs) > 30):
            raise RuntimeError('Cada bloqueo requiere entre 1 y 30 prefijos.')
        if mode == 'permitir_troncales' and (not tids or len(tids) > 100 or prefs):
            raise RuntimeError('Cada permiso de troncales requiere una lista válida de troncales y no admite prefijos.')
        if mode == 'clase_servicio' and (not tids or len(tids) > 100 or not prefs or len(prefs) > 30):
            raise RuntimeError('Cada clase de servicio requiere troncales y entre 1 y 30 prefijos de destino permitidos.')
        for x in exts:
            if not re.match(r'^\d{1,20}$', str(x)):
                raise RuntimeError('Un interno de la restricción no pasó la validación.')
        for x in prefs:
            if not re.match(r'^\d{1,20}$', str(x)):
                raise RuntimeError('Un prefijo de la restricción no pasó la validación.')
        for x in tids:
            if not re.match(r'^\d+$', str(x)):
                raise RuntimeError('Una troncal de la restricción no pasó la validación.')
        if r.get('alcance') not in ('todas', 'troncal', 'ruta'):
            raise RuntimeError('El alcance de una restricción no es válido.')
        if r.get('alcance') == 'troncal' and (len(tids) != 1):
            raise RuntimeError('Una restricción por troncal requiere exactamente una troncal.')
        if r.get('alcance') == 'ruta' and not tids:
            raise RuntimeError('Una restricción por ruta no tiene troncales asociadas.')
        out.append(r)
    return out

def render_block(rules, include_context=True):
    if not rules:
        return ''
    encoded = base64.b64encode(json.dumps(rules, ensure_ascii=True, separators=(',', ':')).encode('utf-8')).decode('ascii')
    lines = [BEGIN]
    for offset in range(0, len(encoded), 120):
        lines.append(DATA_PREFIX + encoded[offset:offset + 120])
    if include_context:
        lines.extend(['[macro-dialout-trunk-predial-hook]',
                      'exten => s,1,NoOp(VoIP Provisioner call restrictions V1)',
                      ' same => n,Set(VP_RSTR_CALLER=${AMPUSER})',
                      ' same => n,ExecIf($["${VP_RSTR_CALLER}" = ""]?Set(VP_RSTR_CALLER=${CALLERID(num)}))',
                      ' same => n,Set(VP_RSTR_DEST=${DIAL_NUMBER})'])
    else:
        # El hook ya existe: agregar sólo prioridades administradas, sin
        # redefinir el contexto ni ejecutar MacroExit antes del código ajeno.
        lines.extend([' same => n,NoOp(VoIP Provisioner call restrictions V1 - hook existente)',
                      ' same => n,Set(VP_RSTR_CALLER=${AMPUSER})',
                      ' same => n,ExecIf($["${VP_RSTR_CALLER}" = ""]?Set(VP_RSTR_CALLER=${CALLERID(num)}))',
                      ' same => n,Set(VP_RSTR_DEST=${DIAL_NUMBER})'])
    for ix, r in enumerate(rules, 1):
        if not r.get('activo', True):
            continue
        label = 'vpcr_%03d' % ix
        exts = sorted(set(str(x) for x in (r.get('internos') or [])))
        caller_expr = ' | '.join('"${VP_RSTR_CALLER}" = "%s"' % x for x in exts)
        lines.append(' same => n,GotoIf($[' + caller_expr + ']?' + label + '_p0:' + label + '_next)')
        mode = str(r.get('modo') or 'bloquear_prefijos')
        if mode == 'permitir_troncales':
            tids = sorted(set(str(x) for x in (r.get('troncales_ids') or [])), key=lambda x: int(x))
            trunk_expr = ' | '.join('"${DIAL_TRUNK}" = "%s"' % x for x in tids)
            lines.append(' same => n(' + label + '_p0),GotoIf($[' + trunk_expr + ']?' + label + '_next:' + label + '_block)')
        elif mode == 'clase_servicio':
            tids = sorted(set(str(x) for x in (r.get('troncales_ids') or [])), key=lambda x: int(x))
            trunk_expr = ' | '.join('"${DIAL_TRUNK}" = "%s"' % x for x in tids)
            lines.append(' same => n(' + label + '_p0),GotoIf($[' + trunk_expr + ']?' + label + '_d0:' + label + '_block)')
            prefs = sorted(set(str(x) for x in (r.get('prefijos') or [])), key=lambda x: (len(x), x))
            for pi, pref in enumerate(prefs):
                fallback = label + ('_d%d' % (pi + 1) if pi + 1 < len(prefs) else '_block')
                lines.append(' same => n(' + label + '_d%d),GotoIf($["${VP_RSTR_DEST:0:%d}" = "%s"]?%s:%s)' % (pi, len(pref), pref, label + '_next', fallback))
        else:
            prefs = sorted(set(str(x) for x in (r.get('prefijos') or [])), key=lambda x: (len(x), x))
            for pi, pref in enumerate(prefs):
                following = label + ('_p%d' % (pi + 1) if pi + 1 < len(prefs) else '_next')
                scope = label + ('_scope' if r.get('alcance') != 'todas' else '_block')
                lines.append(' same => n(' + label + '_p%d),GotoIf($["${VP_RSTR_DEST:0:%d}" = "%s"]?%s:%s)' %
                             (pi, len(pref), pref, scope, following))
            if r.get('alcance') != 'todas':
                tids = sorted(set(str(x) for x in (r.get('troncales_ids') or [])), key=lambda x: int(x))
                trunk_expr = ' | '.join('"${DIAL_TRUNK}" = "%s"' % x for x in tids)
                lines.append(' same => n(' + label + '_scope),GotoIf($[' + trunk_expr + ']?' + label + '_block:' + label + '_next)')
        lines.append(' same => n(' + label + '_block),NoOp(VoIP Provisioner restriction matched)')
        lines.append(' same => n,Playback(ss-noservice)')
        lines.append(' same => n,Hangup(21)')
        lines.append(' same => n(' + label + '_next),NoOp(Continuar evaluación de restricciones)')
    if include_context:
        lines.extend([' same => n,MacroExit()', END, ''])
    else:
        lines.extend([END, ''])
    return '\n'.join(lines)

def atomic_write(path, content):
    folder = os.path.dirname(path)
    oldst = os.stat(path) if os.path.exists(path) else None
    fd, tmp = tempfile.mkstemp(prefix='.vp_call_restrictions_', dir=folder)
    try:
        with os.fdopen(fd, 'w') as f:
            f.write(content)
            f.flush()
            os.fsync(f.fileno())
        if oldst:
            os.chmod(tmp, oldst.st_mode & 0o7777)
            try: os.chown(tmp, oldst.st_uid, oldst.st_gid)
            except Exception: pass
        else:
            os.chmod(tmp, 0o640)
            try: os.chown(tmp, 0, 0)
            except Exception: pass
        os.rename(tmp, path)
    finally:
        if os.path.exists(tmp):
            try: os.unlink(tmp)
            except Exception: pass

def run_ast(command):
    p = subprocess.Popen(['asterisk', '-rx', command], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    out, err = p.communicate(timeout=30)
    if not isinstance(out, str): out = out.decode('utf-8', 'replace')
    if not isinstance(err, str): err = err.decode('utf-8', 'replace')
    return out, err, p.returncode

def main():
    mode = str(P.get('mode') or 'probe')
    if not os.path.isfile(PATH):
        raise RuntimeError('No existe /etc/asterisk/extensions_custom.conf; no se creó un archivo no incluido por Issabel.')
    if not custom_include_ok():
        raise RuntimeError('extensions.conf no incluye extensions_custom.conf. No se modificó el dialplan.')
    original = read_text(PATH)
    base, old_block, had_block = strip_block(original)
    old_rules = decode_rules(old_block)
    file_hook = bool(re.search(r'^\s*\[macro-dialout-trunk-predial-hook\]', base, re.I | re.M))
    hook_integrable = bool(file_hook and hook_context_supported(base))
    if mode == 'probe':
        loaded_hook = False
        if not had_block and not file_hook:
            shown, serr, src = run_ast('dialplan show macro-dialout-trunk-predial-hook')
            loaded_hook = bool(re.search(r"macro-dialout-trunk-predial-hook.*created by", shown, re.I))
        emit({'ok': True, 'reglas': old_rules, 'hay_bloque': had_block,
              'hook_conflictivo': bool((file_hook and not hook_integrable) or (not file_hook and loaded_hook)),
              'hook_existente_integrable': hook_integrable})
    if mode != 'apply':
        raise RuntimeError('Modo remoto no permitido.')
    rules = safe_rules(P.get('reglas') or [])
    if rules and file_hook and not hook_integrable:
        raise RuntimeError('Se encontró un hook manual, pero su estructura no permite integrarlo con seguridad. No se modificó el dialplan. Revisá [macro-dialout-trunk-predial-hook] en extensions_custom.conf: debe tener una línea inicial ext en s, prioridad 1, que no termine el macro ni salte a otro destino.')
    if rules and not had_block and not file_hook:
        shown, serr, src = run_ast('dialplan show macro-dialout-trunk-predial-hook')
        if re.search(r"macro-dialout-trunk-predial-hook.*created by", shown, re.I):
            raise RuntimeError('Asterisk ya tiene un hook cargado desde otra configuración y no aparece como contexto integrable en extensions_custom.conf. No se modificó el dialplan.')
    block = render_block(rules, include_context=(not file_hook))
    new_text = base
    if block:
        if file_hook:
            new_text = insert_managed_block(base, block)
        else:
            if new_text and not new_text.endswith('\n'):
                new_text += '\n'
            if new_text:
                new_text += '\n'
            new_text += block
    if new_text == original:
        emit({'ok': True, 'reglas': rules, 'backup': '', 'sin_cambios': True})
    stamp = time.strftime('%Y%m%d_%H%M%S') + '_' + str(os.getpid())
    backup = PATH + '.bak_vp_call_restrictions_' + stamp
    shutil.copy2(PATH, backup)
    try:
        atomic_write(PATH, new_text)
        out, err, rc = run_ast('dialplan reload')
        if rc != 0 or re.search(r'Unable to connect|No such command|Could not reload', out + '\n' + err, re.I):
            raise RuntimeError('Asterisk no confirmó dialplan reload: ' + (err or out)[-400:])
        if rules:
            shown, serr, src = run_ast('dialplan show macro-dialout-trunk-predial-hook')
            if 'VoIP Provisioner call restrictions V1' not in shown:
                raise RuntimeError('Asterisk no confirmó la carga del hook de restricciones: ' + (serr or shown)[-400:])
    except Exception:
        try:
            shutil.copy2(backup, PATH)
            run_ast('dialplan reload')
        except Exception:
            pass
        raise
    emit({'ok': True, 'reglas': rules, 'backup': backup, 'sin_cambios': False})

try:
    main()
except Exception as e:
    emit({'ok': False, 'mensaje': str(e)[:900]}, 2)
'''
    payload_b64 = _b64.b64encode(_json.dumps(payload, separators=(',', ':')).encode('utf-8')).decode('ascii')
    remote = remote.replace('__VP_CALL_RESTRICTIONS_PAYLOAD__', payload_b64)
    code_b64 = _b64.b64encode(remote.encode('utf-8')).decode('ascii')
    command = "python3 -c \"import base64; exec(compile(base64.b64decode('%s'), '<vp-call-restrictions>', 'exec'))\"" % code_b64
    out, err, rc = _vp_trunk_exec(perfil, command, timeout=timeout)
    out = str(out or '')
    match = re.search(r'VP_CALL_RESTRICTIONS_RESULT:([A-Za-z0-9+/=]+)', out)
    if not match:
        raise RuntimeError('No se pudo consultar/aplicar el administrador de restricciones remoto. ' + str(err or out or '')[-650:])
    try:
        result = _json.loads(_b64.b64decode(match.group(1)).decode('utf-8'))
    except Exception:
        raise RuntimeError('La respuesta del administrador de restricciones no es JSON válido.')
    if not result.get('ok'):
        raise RuntimeError(str(result.get('mensaje') or 'Falló la operación remota.')[:900])
    return result

# VP_CALL_RESTRICTIONS_API_V2
@main_bp.route('/api/monitoreo/ruteo-externo/restricciones-llamadas/catalogo', methods=['GET'])
def api_vp_call_restrictions_catalog_v2():
    try:
        perfil_nombre = str(request.args.get('perfil_nombre') or '').strip()
        perfil = buscar_perfil(perfil_nombre, cargar_perfiles()) if perfil_nombre else None
        if not perfil:
            return jsonify({'ok': False, 'mensaje': 'Seleccioná un servidor válido.'}), 400
        catalog = _vp_call_restrictions_catalog(perfil)
        cols, meta = _vp_ruteo_v6_schema(perfil)
        device_cols = cols.get('devices', set())
        if 'id' not in device_cols:
            raise RuntimeError('No se pudo detectar el catálogo de internos en Issabel.')
        tech_expr = "COALESCE(CAST(`tech` AS CHAR),'')" if 'tech' in device_cols else "''"
        sql = "SELECT COALESCE(CAST(`id` AS CHAR),'')," + tech_expr + " FROM `devices` WHERE CAST(`id` AS CHAR) REGEXP '^[0-9]+$' ORDER BY CAST(`id` AS UNSIGNED),`id`;"
        extensiones = []
        for line in _vp_ruteo_v6_sqlrun(perfil, sql):
            p = line.split('\t')
            if p and p[0].strip().isdigit():
                ext = p[0].strip()
                tech = p[1].strip().lower() if len(p) > 1 else ''
                if tech == 'chan_sip':
                    tech = 'sip'
                extensiones.append({'interno': ext, 'tech': tech.upper() if tech else 'No informada'})
        rutas = []
        for r in catalog.get('rutas', []):
            rutas.append({'id': str(r.get('id') or ''), 'nombre': str(r.get('nombre') or ''),
                          'troncales_ids': list(r.get('troncales_ids') or []),
                          'troncales_compartidas': list(r.get('troncales_compartidas') or [])})
        return jsonify({'ok': True, 'perfil': perfil_nombre, 'extensiones': extensiones,
                        'rutas_salientes': rutas, 'troncales': catalog.get('troncales', []),
                        'asociacion_ruta_troncal': bool(catalog.get('association_supported'))})
    except Exception as exc:
        logger.exception('[RESTRICCIONES V2] Error consultando catálogo')
        return jsonify({'ok': False, 'mensaje': 'No se pudo consultar el catálogo de Issabel: ' + str(exc)[:700]}), 500


@main_bp.route('/api/monitoreo/ruteo-externo/restricciones-llamadas', methods=['GET'])
def api_vp_call_restrictions_list_v2():
    try:
        perfil_nombre = str(request.args.get('perfil_nombre') or '').strip()
        perfil = buscar_perfil(perfil_nombre, cargar_perfiles()) if perfil_nombre else None
        if not perfil:
            return jsonify({'ok': False, 'mensaje': 'Seleccioná un servidor válido.'}), 400
        result = _vp_call_restrictions_remote(perfil, {'mode': 'probe'}, timeout=90)
        rules = [r for r in (result.get('reglas', []) or [])
                 if str(r.get('modo') or 'bloquear_prefijos') != 'permitir_troncales']
        return jsonify({'ok': True, 'reglas': rules,
                        'hook_conflictivo': result.get('hook_conflictivo', False)})
    except Exception as exc:
        logger.exception('[RESTRICCIONES V2] Error leyendo reglas')
        return jsonify({'ok': False, 'mensaje': 'No se pudieron leer las restricciones: ' + str(exc)[:700]}), 500


@main_bp.route('/api/monitoreo/ruteo-externo/restricciones-llamadas/guardar', methods=['POST'])
def api_vp_call_restrictions_save_v2():
    backup_path = ''
    try:
        data = request.get_json(silent=True) or {}
        if not _vp_trunk_admin_password(data):
            return jsonify({'ok': False, 'mensaje': 'Contraseña de administrador incorrecta.'}), 401
        perfil_nombre = str(data.get('perfil_nombre') or '').strip()
        perfil = buscar_perfil(perfil_nombre, cargar_perfiles()) if perfil_nombre else None
        if not perfil:
            return jsonify({'ok': False, 'mensaje': 'Seleccioná un servidor válido.'}), 400
        nombre = str(data.get('nombre') or '').strip()
        if not nombre or len(nombre) > 100 or any(ord(c) < 32 for c in nombre):
            raise ValueError('Ingresá un nombre de regla de hasta 100 caracteres.')
        service_class = str(data.get('modo') or '').strip().lower() == 'clase_servicio'
        raw_prefixes = data.get('prefijos_permitidos', []) if service_class else data.get('prefijos', [])
        if isinstance(raw_prefixes, str):
            raw_prefixes = re.split(r'[\s,;]+', raw_prefixes.strip())
        if not isinstance(raw_prefixes, list):
            raise ValueError('Ingresá uno o varios prefijos numéricos.')
        prefixes = []
        for value in raw_prefixes:
            value = str(value or '').strip()
            if not value:
                continue
            if not re.match(r'^\d{1,20}$', value):
                raise ValueError('Cada prefijo debe contener sólo dígitos (1–20). Ejemplos: 15, 00, 011.')
            if value not in prefixes:
                prefixes.append(value)
        if not prefixes or len(prefixes) > 30:
            raise ValueError('Definí entre 1 y 30 prefijos para la regla.')
        raw_exts = data.get('internos', [])
        if not isinstance(raw_exts, list):
            raise ValueError('Seleccioná uno o varios internos.')
        extensions = []
        for value in raw_exts:
            value = str(value or '').strip()
            if not re.match(r'^\d{1,20}$', value):
                raise ValueError('La selección contiene un interno inválido.')
            if value not in extensions:
                extensions.append(value)
        if not extensions or len(extensions) > 1000:
            raise ValueError('Seleccioná entre 1 y 1000 internos.')
        known_rows = _vp_ruteo_v6_sqlrun(perfil, "SELECT COALESCE(CAST(`id` AS CHAR),'') FROM `devices` WHERE CAST(`id` AS CHAR) REGEXP '^[0-9]+$';")
        known_exts = set(x.split('\t')[0].strip() for x in known_rows if x.split('\t')[0].strip().isdigit())
        if not set(extensions).issubset(known_exts):
            raise ValueError('Uno o más internos ya no existen en Issabel. Actualizá el catálogo y volvé a guardar.')
        alcance = str(data.get('alcance') or 'todas').strip().lower()
        catalog = _vp_call_restrictions_catalog(perfil)
        target_id = ''
        target_name = 'Todas las troncales'
        trunk_ids = []
        if service_class:
            raw_tids = data.get('troncales_ids', [])
            if not isinstance(raw_tids, list) or not raw_tids or len(raw_tids) > 100:
                raise ValueError('Seleccioná al menos una troncal permitida para esta clase.')
            for value in raw_tids:
                value = str(value or '').strip()
                if not value.isdigit() or len(value) > 12:
                    raise ValueError('La selección contiene una troncal inválida.')
                if value not in trunk_ids:
                    trunk_ids.append(value)
            if not set(trunk_ids).issubset(catalog.get('trunk_ids', set())):
                raise ValueError('Una o más troncales seleccionadas ya no existen en Issabel.')
            target_name = 'Troncales permitidas'
            alcance = 'todas'
        else:
            if alcance not in ('todas', 'troncal', 'ruta'):
                raise ValueError('Elegí Todas las troncales, Una troncal o Una ruta saliente.')
            if alcance == 'troncal':
                target_id = str(data.get('destino_id') or '').strip()
                if not target_id.isdigit() or target_id not in catalog['trunk_ids']:
                    raise ValueError('La troncal seleccionada no existe o no se pudo verificar.')
                target_name = next((x.get('nombre') for x in catalog['troncales'] if x.get('id') == target_id), 'Troncal ' + target_id)
                trunk_ids = [target_id]
            elif alcance == 'ruta':
                target_id = str(data.get('destino_id') or '').strip()
                if not target_id.isdigit() or target_id not in catalog['route_names']:
                    raise ValueError('La ruta saliente seleccionada no existe o no se pudo verificar.')
                if not catalog['association_supported']:
                    raise ValueError('Esta versión no permite verificar la relación ruta–troncal. No se guardó nada para evitar afectar otras rutas.')
                trunk_ids = sorted(set(catalog['route_trunks'].get(target_id, [])), key=lambda x: int(x))
                if not trunk_ids:
                    raise ValueError('La ruta no tiene troncales asociadas en Issabel.')
                shared = []
                for tid in trunk_ids:
                    others = [rid for rid in catalog['trunk_routes'].get(tid, []) if rid != target_id]
                    if others:
                        tname = next((x.get('nombre') for x in catalog['troncales'] if x.get('id') == tid), 'Troncal ' + tid)
                        shared.append(tname + ' (también usada por ' + ', '.join(catalog['route_names'].get(rid, 'Ruta ' + rid) for rid in others) + ')')
                if shared:
                    raise ValueError('No guardé la regla por ruta porque comparte troncales con otras rutas: ' + '; '.join(shared[:5]) + '. Elegí una troncal específica o Todas las troncales para evitar bloquear llamadas ajenas.')
                target_name = catalog['route_names'].get(target_id, 'Ruta ' + target_id)
        active = data.get('activo', True)
        if not isinstance(active, bool):
            raise ValueError('El estado activo debe ser verdadero o falso.')
        rule_id = str(data.get('id') or '').strip().lower()
        if rule_id and not re.match(r'^[a-f0-9]{8,32}$', rule_id):
            raise ValueError('El identificador de regla no es válido.')
        result = _vp_call_restrictions_remote(perfil, {'mode': 'probe'}, timeout=90)
        rules = result.get('reglas', []) or []
        found = None
        if rule_id:
            for ix, old in enumerate(rules):
                if str(old.get('id') or '') == rule_id:
                    found = ix
                    break
            if found is None:
                raise ValueError('La regla a editar ya no existe. Actualizá la pantalla.')
        else:
            rule_id = uuid.uuid4().hex
        was_edit = found is not None
        if service_class:
            # Una extensión sólo debe pertenecer a una clase combinada. Al asignarla,
            # quitarla de clases previas y de permisos de troncal individuales.
            filtered_rules = []
            for old_rule in rules:
                old_mode = str(old_rule.get('modo') or 'bloquear_prefijos')
                if str(old_rule.get('id') or '') == rule_id:
                    continue
                if old_mode not in ('clase_servicio', 'permitir_troncales'):
                    filtered_rules.append(old_rule)
                    continue
                remaining = [str(x) for x in (old_rule.get('internos') or []) if str(x) not in extensions]
                if remaining:
                    copy_rule = dict(old_rule)
                    copy_rule['internos'] = remaining
                    filtered_rules.append(copy_rule)
            rules = filtered_rules
            found = None
        rule = {'id': rule_id, 'modo': ('clase_servicio' if service_class else 'bloquear_prefijos'),
                'nombre': nombre, 'prefijos': prefixes,
                'internos': sorted(extensions, key=lambda x: (len(x), x)),
                'alcance': alcance, 'destino_id': target_id, 'destino_nombre': target_name,
                'troncales_ids': trunk_ids, 'activo': active}
        if found is None:
            rules.append(rule)
        else:
            rules[found] = rule
        applied = _vp_call_restrictions_remote(perfil, {'mode': 'apply', 'reglas': rules}, timeout=90)
        backup_path = applied.get('backup', '') or ''
        verb = 'Se actualizó' if was_edit else 'Se creó'
        return jsonify({'ok': True, 'id': rule_id, 'mensaje': verb + ' la regla "' + nombre + '". Prefijos: ' + ', '.join(prefixes) + '. Internos afectados: ' + str(len(extensions)) + '.', 'respaldo': backup_path})
    except ValueError as exc:
        return jsonify({'ok': False, 'mensaje': str(exc), 'respaldo': backup_path}), 400
    except Exception as exc:
        logger.exception('[RESTRICCIONES V2] Error guardando regla')
        return jsonify({'ok': False, 'mensaje': 'No se pudo guardar la regla: ' + str(exc)[:800], 'respaldo': backup_path}), 500


@main_bp.route('/api/monitoreo/ruteo-externo/restricciones-llamadas/eliminar', methods=['POST'])
def api_vp_call_restrictions_delete_v2():
    backup_path = ''
    try:
        data = request.get_json(silent=True) or {}
        if not _vp_trunk_admin_password(data):
            return jsonify({'ok': False, 'mensaje': 'Contraseña de administrador incorrecta.'}), 401
        perfil_nombre = str(data.get('perfil_nombre') or '').strip()
        perfil = buscar_perfil(perfil_nombre, cargar_perfiles()) if perfil_nombre else None
        if not perfil:
            return jsonify({'ok': False, 'mensaje': 'Seleccioná un servidor válido.'}), 400
        rule_id = str(data.get('id') or '').strip().lower()
        if not re.match(r'^[a-f0-9]{8,32}$', rule_id):
            raise ValueError('El identificador de regla no es válido.')
        result = _vp_call_restrictions_remote(perfil, {'mode': 'probe'}, timeout=90)
        rules = result.get('reglas', []) or []
        new_rules = [x for x in rules if str(x.get('id') or '') != rule_id]
        if len(new_rules) == len(rules):
            raise ValueError('La regla ya no existe. Actualizá la pantalla.')
        applied = _vp_call_restrictions_remote(perfil, {'mode': 'apply', 'reglas': new_rules}, timeout=90)
        backup_path = applied.get('backup', '') or ''
        return jsonify({'ok': True, 'mensaje': 'Se eliminó la regla de restricción.', 'respaldo': backup_path})
    except ValueError as exc:
        return jsonify({'ok': False, 'mensaje': str(exc), 'respaldo': backup_path}), 400
    except Exception as exc:
        logger.exception('[RESTRICCIONES V2] Error eliminando regla')
        return jsonify({'ok': False, 'mensaje': 'No se pudo eliminar la regla: ' + str(exc)[:800], 'respaldo': backup_path}), 500
# VP_INTERNAL_TRUNK_PERMISSIONS_API_V1
@main_bp.route('/api/monitoreo/internos/permisos-troncales', methods=['GET'])
def api_vp_internos_permisos_troncales_v1():
    try:
        perfil_nombre = str(request.args.get('perfil_nombre') or '').strip()
        perfil = buscar_perfil(perfil_nombre, cargar_perfiles()) if perfil_nombre else None
        if not perfil:
            return jsonify({'ok': False, 'mensaje': 'Seleccioná un servidor válido.'}), 400
        catalog = _vp_call_restrictions_catalog(perfil)
        rows = _vp_ruteo_v6_sqlrun(perfil, "SELECT COALESCE(CAST(id AS CHAR),''),COALESCE(CAST(tech AS CHAR),'') FROM devices WHERE CAST(id AS CHAR) REGEXP '^[0-9]+$' ORDER BY CAST(id AS UNSIGNED),id;")
        extensions = []
        for line in rows:
            p = line.split('\t')
            if p and p[0].strip().isdigit():
                extensions.append({'interno': p[0].strip(), 'tecnologia': (p[1].strip().upper() if len(p) > 1 else '')})
        remote = _vp_call_restrictions_remote(perfil, {'mode': 'probe'}, timeout=90)
        assignments = {}
        for rule in (remote.get('reglas', []) or []):
            if str(rule.get('modo') or 'bloquear_prefijos') not in ('permitir_troncales', 'clase_servicio'):
                continue
            ids = sorted(set(str(x) for x in (rule.get('troncales_ids') or [])), key=lambda x: int(x))
            for ext in (rule.get('internos') or []):
                assignments[str(ext)] = ids
        return jsonify({'ok': True, 'perfil': perfil_nombre, 'extensiones': extensions,
                        'troncales': catalog.get('troncales', []), 'asignaciones': assignments,
                        'hook_conflictivo': bool(remote.get('hook_conflictivo'))})
    except Exception as exc:
        logger.exception('[PERMISOS TRONCALES] Error consultando permisos por interno')
        return jsonify({'ok': False, 'mensaje': 'No se pudieron consultar los permisos: ' + str(exc)[:700]}), 500


@main_bp.route('/api/monitoreo/interno/permisos-troncales', methods=['POST'])
def api_vp_interno_permisos_troncales_v1():
    backup_path = ''
    try:
        data = request.get_json(silent=True) or {}
        if not _vp_trunk_admin_password(data):
            return jsonify({'ok': False, 'mensaje': 'Contraseña de administrador incorrecta.'}), 401
        perfil_nombre = str(data.get('perfil_nombre') or '').strip()
        perfil = buscar_perfil(perfil_nombre, cargar_perfiles()) if perfil_nombre else None
        if not perfil:
            return jsonify({'ok': False, 'mensaje': 'Seleccioná un servidor válido.'}), 400
        raw_exts = data.get('internos', [])
        if not isinstance(raw_exts, list) or not raw_exts or len(raw_exts) > 1000:
            raise ValueError('Seleccioná uno o varios internos válidos.')
        extensions = []
        for value in raw_exts:
            value = str(value or '').strip()
            if not re.match(r'^\d{1,20}$', value):
                raise ValueError('La selección contiene un interno inválido.')
            if value not in extensions:
                extensions.append(value)
        raw_tids = data.get('troncales_ids', [])
        if not isinstance(raw_tids, list) or len(raw_tids) > 100:
            raise ValueError('La lista de troncales permitidas no es válida.')
        trunk_ids = []
        for value in raw_tids:
            value = str(value or '').strip()
            if not re.match(r'^\d{1,12}$', value):
                raise ValueError('La selección contiene una troncal inválida.')
            if value not in trunk_ids:
                trunk_ids.append(value)
        known_rows = _vp_ruteo_v6_sqlrun(perfil, "SELECT COALESCE(CAST(id AS CHAR),'') FROM devices WHERE CAST(id AS CHAR) REGEXP '^[0-9]+$';")
        known_exts = set(x.split('\t')[0].strip() for x in known_rows if x.split('\t')[0].strip().isdigit())
        if not set(extensions).issubset(known_exts):
            raise ValueError('Uno o más internos ya no existen en Issabel. Actualizá el listado.')
        catalog = _vp_call_restrictions_catalog(perfil)
        if not set(trunk_ids).issubset(catalog.get('trunk_ids', set())):
            raise ValueError('Una o más troncales ya no existen en el servidor seleccionado.')
        remote = _vp_call_restrictions_remote(perfil, {'mode': 'probe'}, timeout=90)
        rules = remote.get('reglas', []) or []
        updated = []
        for rule in rules:
            if str(rule.get('modo') or 'bloquear_prefijos') not in ('permitir_troncales', 'clase_servicio'):
                updated.append(rule)
                continue
            remaining = [str(x) for x in (rule.get('internos') or []) if str(x) not in extensions]
            if remaining:
                copy_rule = dict(rule)
                copy_rule['internos'] = remaining
                updated.append(copy_rule)
        if trunk_ids:
            for ext in extensions:
                updated.append({'id': uuid.uuid4().hex, 'modo': 'permitir_troncales',
                                'nombre': 'Permisos de troncales - ' + ext, 'prefijos': [],
                                'internos': [ext], 'alcance': 'todas', 'destino_id': '',
                                'destino_nombre': 'Troncales permitidas', 'troncales_ids': trunk_ids,
                                'activo': True})
        applied = _vp_call_restrictions_remote(perfil, {'mode': 'apply', 'reglas': updated}, timeout=90)
        backup_path = applied.get('backup', '') or ''
        message = ('Se asignaron ' + str(len(trunk_ids)) + ' troncal(es) permitida(s) a ' + str(len(extensions)) + ' interno(s).') if trunk_ids else ('Se quitaron las restricciones de troncal de ' + str(len(extensions)) + ' interno(s).')
        return jsonify({'ok': True, 'mensaje': message, 'respaldo': backup_path,
                        'internos': extensions, 'troncales_ids': trunk_ids})
    except ValueError as exc:
        return jsonify({'ok': False, 'mensaje': str(exc), 'respaldo': backup_path}), 400
    except Exception as exc:
        logger.exception('[PERMISOS TRONCALES] Error guardando permisos por interno')
        return jsonify({'ok': False, 'mensaje': 'No se pudieron guardar los permisos: ' + str(exc)[:800], 'respaldo': backup_path}), 500
# FIN VP_INTERNAL_TRUNK_PERMISSIONS_API_V1


# FIN VP_CALL_RESTRICTIONS_API_V2

# VP_RUTEO_EXTERNO_ELIMINAR_API_V1
@main_bp.route('/api/monitoreo/ruteo-externo/eliminar', methods=['POST'])
def api_monitoreo_ruteo_externo_eliminar_v1():
    'Elimina una ruta saliente, condición horaria o grupo horario con respaldo y validación.'
    backup_path = ''
    sql_started = False
    perfil = None
    try:
        data = request.get_json(silent=True) or {}
        if not _vp_trunk_admin_password(data):
            return jsonify({'ok': False, 'mensaje': 'Contraseña de administrador incorrecta.'}), 401
        perfil_nombre = str(data.get('perfil_nombre') or '').strip()
        perfil = buscar_perfil(perfil_nombre, cargar_perfiles()) if perfil_nombre else None
        if not perfil:
            return jsonify({'ok': False, 'mensaje': 'Seleccioná un servidor válido.'}), 400

        tipo = str(data.get('tipo') or '').strip().lower()
        ident = str(data.get('id') or '').strip()
        if tipo not in ('salida', 'condicion', 'grupo'):
            raise ValueError('Solo se pueden eliminar rutas salientes, condiciones horarias o grupos horarios desde esta opción.')
        if not re.match(r'^\d{1,12}$', ident) or int(ident) <= 0:
            raise ValueError('El ID del elemento no es válido.')

        schema_rows = _vp_ruteo_v6_sqlrun(
            perfil,
            "SELECT table_name,column_name FROM information_schema.columns "
            "WHERE table_schema=DATABASE() AND table_name IN "
            "('outbound_routes','outbound_route_sequence','outbound_route_patterns',"
            "'outbound_route_trunks','timeconditions','timegroups_groups','timegroups_details','incoming') "
            "ORDER BY table_name,ordinal_position;"
        )
        cols = {}
        for line in schema_rows:
            p = line.split('\t')
            if len(p) >= 2:
                cols.setdefault(p[0].strip().lower(), set()).add(p[1].strip().lower())

        def has(table, *names):
            return all(str(n).lower() in cols.get(table.lower(), set()) for n in names)

        def count(sql, error_text):
            rows = _vp_ruteo_v6_sqlrun(perfil, sql, timeout=60)
            if not rows or not str(rows[-1]).strip().isdigit():
                raise RuntimeError(error_text + ' No se pudo verificar el registro en la base de datos.')
            return int(str(rows[-1]).strip())

        def q(value):
            return _vp_trunk_sql_q(str(value if value is not None else ''))

        if tipo == 'salida':
            if not has('outbound_routes', 'route_id', 'name'):
                raise RuntimeError('No se detectó la tabla de rutas salientes compatible.')
            if count('SELECT COUNT(*) FROM outbound_routes WHERE route_id=' + ident + ';', 'Ruta saliente:') != 1:
                raise ValueError('La ruta saliente ya no existe. Actualizá Monitoreo.')
            rows = _vp_ruteo_v6_sqlrun(perfil, "SELECT COALESCE(name,'') FROM outbound_routes WHERE route_id=" + ident + ' LIMIT 1;')
            label = str(rows[0]).strip() if rows else ('Ruta ' + ident)
            tables = ['outbound_routes']
            for table in ('outbound_route_patterns', 'outbound_route_trunks', 'outbound_route_sequence'):
                if table in cols and 'route_id' in cols[table]:
                    tables.append(table)
            statements = []
            for table in ('outbound_route_patterns', 'outbound_route_trunks', 'outbound_route_sequence'):
                if table in tables:
                    statements.append('DELETE FROM `' + table + '` WHERE `route_id`=' + ident + ';')
            statements.append('DELETE FROM `outbound_routes` WHERE `route_id`=' + ident + ';')
            backup_label = 'ruta saliente "' + label + '"'
        elif tipo == 'condicion':
            if not has('timeconditions', 'timeconditions_id'):
                raise RuntimeError('No se detectó la tabla de condiciones horarias compatible.')
            if count('SELECT COUNT(*) FROM timeconditions WHERE timeconditions_id=' + ident + ';', 'Condición horaria:') != 1:
                raise ValueError('La condición horaria ya no existe. Actualizá Monitoreo.')
            name_col = 'displayname' if 'displayname' in cols.get('timeconditions', set()) else ('description' if 'description' in cols.get('timeconditions', set()) else '')
            rows = _vp_ruteo_v6_sqlrun(perfil, "SELECT COALESCE(`" + name_col + "`,'') FROM timeconditions WHERE timeconditions_id=" + ident + ' LIMIT 1;') if name_col else []
            label = str(rows[0]).strip() if rows else ('Condición ' + ident)
            pattern = q('%timeconditions,' + ident + ',%')
            references = []
            if has('incoming', 'destination'):
                references.append(('incoming', "SELECT COUNT(*) FROM incoming WHERE COALESCE(destination,'') LIKE " + pattern + ';'))
            for col in ('truegoto', 'falsegoto', 'destmatch', 'destnomatch'):
                if col in cols.get('timeconditions', set()):
                    references.append(('timeconditions', 'SELECT COUNT(*) FROM timeconditions WHERE timeconditions_id<>' + ident + ' AND COALESCE(`' + col + "`,'') LIKE " + pattern + ';'))
            for table, sqlref in references:
                if count(sqlref, 'No se pudo comprobar si la condición horaria está en uso:') > 0:
                    raise ValueError('No se eliminó la condición "' + label + '": otra regla entrante u horaria la utiliza como destino. Quitá primero esa referencia y volvé a intentar.')
            tables = ['timeconditions']
            statements = ['DELETE FROM `timeconditions` WHERE `timeconditions_id`=' + ident + ';']
            backup_label = 'condición horaria "' + label + '"'
        else:
            if not has('timegroups_groups', 'id', 'description') or not has('timegroups_details', 'timegroupid'):
                raise RuntimeError('No se detectaron las tablas compatibles de grupos horarios.')
            if count('SELECT COUNT(*) FROM timegroups_groups WHERE id=' + ident + ';', 'Grupo horario:') != 1:
                raise ValueError('El grupo horario ya no existe. Actualizá Monitoreo.')
            rows = _vp_ruteo_v6_sqlrun(perfil, "SELECT COALESCE(description,'') FROM timegroups_groups WHERE id=" + ident + ' LIMIT 1;')
            label = str(rows[0]).strip() if rows else ('Grupo horario ' + ident)
            if has('outbound_routes', 'time_group_id'):
                if count('SELECT COUNT(*) FROM outbound_routes WHERE time_group_id=' + ident + ';', 'No se pudo comprobar el uso del grupo horario:') > 0:
                    raise ValueError('No se eliminó el grupo horario "' + label + '": está asociado a una o más rutas salientes. Quitá primero esa asociación.')
            if has('timeconditions', 'time'):
                if count('SELECT COUNT(*) FROM timeconditions WHERE CAST(`time` AS CHAR)=' + q(ident) + ';', 'No se pudo comprobar el uso del grupo horario:') > 0:
                    raise ValueError('No se eliminó el grupo horario "' + label + '": está asociado a una o más condiciones horarias. Quitá primero esa asociación.')
            tables = ['timegroups_groups', 'timegroups_details']
            statements = ['DELETE FROM `timegroups_details` WHERE `timegroupid`=' + ident + ';',
                          'DELETE FROM `timegroups_groups` WHERE `id`=' + ident + ';']
            backup_label = 'grupo horario "' + label + '"'

        allowed_tables = set(['outbound_routes','outbound_route_patterns','outbound_route_trunks',
                              'outbound_route_sequence','timeconditions','timegroups_groups','timegroups_details'])
        tables = sorted(set(t for t in tables if t in allowed_tables and t in cols))
        if not tables:
            raise RuntimeError('No se pudo determinar qué tablas respaldar. No se eliminó nada.')

        stamp = time.strftime('%Y%m%d_%H%M%S') + '_' + str(os.getpid())
        backup_path = '/var/backups/voip-provisioner/ruteo-externo/ruteo_del_' + stamp + '.sql'
        comando_backup = r'''set -eu
AMPDBUSER="$(grep '^AMPDBUSER=' /etc/amportal.conf | head -1 | cut -d= -f2-)"
AMPDBPASS="$(grep '^AMPDBPASS=' /etc/amportal.conf | head -1 | cut -d= -f2-)"
AMPDBHOST="$(grep '^AMPDBHOST=' /etc/amportal.conf | head -1 | cut -d= -f2-)"
AMPDBNAME="$(grep '^AMPDBNAME=' /etc/amportal.conf | head -1 | cut -d= -f2-)"
[ -n "$AMPDBUSER" ] && [ -n "$AMPDBPASS" ] && [ -n "$AMPDBNAME" ] || { echo 'No se pudieron leer credenciales AMPDB'; exit 20; }
command -v mysqldump >/dev/null 2>&1 || { echo 'No se encuentra mysqldump'; exit 21; }
mkdir -p /var/backups/voip-provisioner/ruteo-externo
chmod 700 /var/backups/voip-provisioner /var/backups/voip-provisioner/ruteo-externo
TMP_CNF="$(mktemp)"
trap 'rm -f "$TMP_CNF"' EXIT
printf '[client]\nuser=%s\npassword=%s\n' "$AMPDBUSER" "$AMPDBPASS" > "$TMP_CNF"
[ -z "$AMPDBHOST" ] || printf 'host=%s\n' "$AMPDBHOST" >> "$TMP_CNF"
chmod 600 "$TMP_CNF"
DEST="''' + backup_path + r'''"
TABLES="''' + ' '.join(tables) + r'''"
mysqldump --defaults-extra-file="$TMP_CNF" --add-drop-table --skip-lock-tables "$AMPDBNAME" $TABLES > "$DEST"
[ -s "$DEST" ] || { echo 'El respaldo SQL quedó vacío'; rm -f "$DEST"; exit 22; }
chmod 600 "$DEST"
echo "VP_RUTEO_DELETE_BACKUP_OK:$DEST"
'''
        out, err, code = _vp_trunk_exec(perfil, comando_backup, timeout=120)
        out = str(out or '')
        if code != 0 or 'VP_RUTEO_DELETE_BACKUP_OK:' not in out or not re.search(r'VP_RUTEO_DELETE_BACKUP_OK:' + re.escape(backup_path), out):
            backup_path = ''
            raise RuntimeError('No se pudo crear el respaldo SQL previo; no se eliminó nada. ' + str(err or out or '')[-500:])

        sql_started = True
        sql = ' '.join(statements) + " SELECT CONCAT('VP_RUTEO_DELETE_OK|'," + q(ident) + ");"
        output = _vp_ruteo_v6_sqlrun(perfil, sql, timeout=100)
        if not any('VP_RUTEO_DELETE_OK|' + ident in str(line) for line in output):
            raise RuntimeError('Issabel no confirmó la eliminación de ' + backup_label + '.')
        try:
            _vp_ruteo_v6_apply(perfil)
        except Exception as apply_error:
            raise RuntimeError('La eliminación SQL se ejecutó, pero Issabel no confirmó la regeneración/recarga. ' + str(apply_error)[:400])
        return jsonify({'ok': True, 'mensaje': 'Se eliminó ' + backup_label + ' y se aplicó la configuración de Issabel.',
                        'tipo': tipo, 'id': ident, 'respaldo': backup_path})
    except ValueError as exc:
        return jsonify({'ok': False, 'mensaje': str(exc), 'respaldo': backup_path}), 400
    except Exception as exc:
        detail = str(exc)[:850]
        if backup_path and sql_started and perfil:
            try:
                if not re.match(r'^/var/backups/voip-provisioner/ruteo-externo/ruteo_del_\d{8}_\d{6}_\d+\.sql$', backup_path):
                    raise RuntimeError('La ruta del respaldo no pasó la validación.')
                command = r'''set -eu
AMPDBUSER="$(grep '^AMPDBUSER=' /etc/amportal.conf | head -1 | cut -d= -f2-)"
AMPDBPASS="$(grep '^AMPDBPASS=' /etc/amportal.conf | head -1 | cut -d= -f2-)"
AMPDBHOST="$(grep '^AMPDBHOST=' /etc/amportal.conf | head -1 | cut -d= -f2-)"
AMPDBNAME="$(grep '^AMPDBNAME=' /etc/amportal.conf | head -1 | cut -d= -f2-)"
TMP_CNF="$(mktemp)"; trap 'rm -f "$TMP_CNF"' EXIT
printf '[client]\nuser=%s\npassword=%s\n' "$AMPDBUSER" "$AMPDBPASS" > "$TMP_CNF"
[ -z "$AMPDBHOST" ] || printf 'host=%s\n' "$AMPDBHOST" >> "$TMP_CNF"
chmod 600 "$TMP_CNF"
mysql --defaults-extra-file="$TMP_CNF" "$AMPDBNAME" < "''' + backup_path + r'''"
echo VP_RUTEO_DELETE_RESTORE_OK
'''
                out2, err2, code2 = _vp_trunk_exec(perfil, command, timeout=120)
                if code2 != 0 or 'VP_RUTEO_DELETE_RESTORE_OK' not in str(out2 or ''):
                    raise RuntimeError(str(err2 or out2 or 'No se confirmó la restauración.')[-350:])
                _vp_ruteo_v6_apply(perfil)
                detail += ' Se restauró el respaldo SQL y se recargó la configuración anterior.'
            except Exception as restore_exc:
                detail += ' ATENCIÓN: no se confirmó la restauración automática; respaldo manual: ' + backup_path + '. ' + str(restore_exc)[:350]
            detail += ' Respaldo: ' + backup_path
        logger.exception('[RUTEO EXTERNO ELIMINAR] Error eliminando configuración')
        return jsonify({'ok': False, 'mensaje': 'No se pudo eliminar la configuración: ' + detail,
                        'respaldo': backup_path}), 500
# FIN VP_RUTEO_EXTERNO_ELIMINAR_API_V1

# VP_ISSABEL_NATIVE_COS_READ_API_V1
@main_bp.route('/api/monitoreo/ruteo-externo/clases-issabel', methods=['GET'])
def api_vp_issabel_native_cos_read_v1():
    """Lee las clases nativas de Issabel/Custom Contexts; consulta solamente, no modifica nada."""
    try:
        perfil_nombre = str(request.args.get('perfil_nombre') or '').strip()
        perfil = buscar_perfil(perfil_nombre, cargar_perfiles()) if perfil_nombre else None
        if not perfil:
            return jsonify({'ok': False, 'mensaje': 'Seleccioná un servidor válido.'}), 400

        schema_rows = _vp_ruteo_v6_sqlrun(
            perfil,
            "SELECT table_name,column_name FROM information_schema.columns "
            "WHERE table_schema=DATABASE() AND table_name IN "
            "('customcontexts_contexts','customcontexts_contexts_list','customcontexts_includes',"
            "'customcontexts_includes_list','sip','iax','pjsip','ps_endpoints') "
            "ORDER BY table_name,ordinal_position;"
        )
        cols = {}
        for line in schema_rows:
            p = line.split('\t')
            if len(p) >= 2:
                cols.setdefault(p[0].strip().lower(), set()).add(p[1].strip().lower())

        cc = cols.get('customcontexts_contexts', set())
        cc_list = cols.get('customcontexts_contexts_list', set())
        if not (set(['context', 'description']).issubset(cc) or set(['context', 'description']).issubset(cc_list)):
            return jsonify({
                'ok': True, 'disponible': False, 'clases': [],
                'mensaje': 'No se encontraron las tablas nativas customcontexts_contexts / customcontexts_contexts_list en la base de datos de Issabel. No se modificó ninguna configuración.'
            })

        def expr(columns, name):
            if name in columns:
                return "COALESCE(CAST(`" + name + "` AS CHAR),'')"
            return "''"

        classes = {}
        source = 'customcontexts_contexts' if set(['context', 'description']).issubset(cc) else 'customcontexts_contexts_list'
        source_cols = cols.get(source, set())
        wanted = ['context', 'description', 'dialrules', 'faildestination', 'failpin',
                  'failpincdr', 'featurefaildestination', 'featurefailpin', 'featurefailpincdr']
        select_cols = [expr(source_cols, name) for name in wanted]
        if 'context' not in source_cols or 'description' not in source_cols:
            raise RuntimeError('La tabla de clases nativas no contiene context y description.')
        sql = 'SELECT ' + ','.join(select_cols) + ' FROM `' + source + '` ORDER BY `context`;'
        for line in _vp_ruteo_v6_sqlrun(perfil, sql):
            p = line.split('\t')
            if not p or not p[0].strip():
                continue
            row = dict((wanted[i], p[i].strip() if i < len(p) else '') for i in range(len(wanted)))
            context = row.get('context', '')
            if not context:
                continue
            classes[context] = {
                'context': context,
                'nombre': row.get('description') or context,
                'descripcion': row.get('description') or '',
                'dialrules': row.get('dialrules') or '',
                'destino_fallo': row.get('faildestination') or '',
                'pin_fallo': bool(row.get('failpin')),
                'destino_fallo_funciones': row.get('featurefaildestination') or '',
                'inclusiones': [],
                'internos': [],
                'bloqueada': False
            }

        # Algunas versiones conservan metadatos/estado de bloqueo en la tabla _list.
        if 'customcontexts_contexts_list' in cols and set(['context', 'description']).issubset(cc_list):
            lock_expr = expr(cc_list, 'locked')
            for line in _vp_ruteo_v6_sqlrun(
                    perfil,
                    'SELECT ' + expr(cc_list, 'context') + ',' + expr(cc_list, 'description') + ',' + lock_expr +
                    ' FROM `customcontexts_contexts_list` ORDER BY `context`;'):
                p = line.split('\t')
                if len(p) < 2 or not p[0].strip():
                    continue
                ctx = p[0].strip()
                if ctx not in classes:
                    classes[ctx] = {'context': ctx, 'nombre': p[1].strip() or ctx, 'descripcion': p[1].strip(),
                                    'dialrules': '', 'destino_fallo': '', 'pin_fallo': False,
                                    'destino_fallo_funciones': '', 'inclusiones': [], 'internos': [],
                                    'bloqueada': False}
                if len(p) > 2:
                    classes[ctx]['bloqueada'] = str(p[2]).strip().lower() in ('1', 'yes', 'true', 'on')

        # Mostrar los contextos/reglas incluidos en cada clase nativa, sin alterarlos.
        inc_cols = cols.get('customcontexts_includes', set())
        if set(['context', 'include']).issubset(inc_cols):
            inc_sql = 'SELECT ' + ','.join([
                expr(inc_cols, 'context'), expr(inc_cols, 'include'), expr(inc_cols, 'sort'),
                expr(inc_cols, 'timegroupid'), expr(inc_cols, 'userules')
            ]) + ' FROM `customcontexts_includes` ORDER BY ' + expr(inc_cols, 'context') + ',' + expr(inc_cols, 'sort') + ';'
            for line in _vp_ruteo_v6_sqlrun(perfil, inc_sql):
                p = line.split('\t')
                if len(p) < 2:
                    continue
                ctx, included = p[0].strip(), p[1].strip()
                if ctx not in classes or not included:
                    continue
                detail = {'context': included,
                          'orden': p[2].strip() if len(p) > 2 else '',
                          'grupo_horario': p[3].strip() if len(p) > 3 else '',
                          'usar_reglas': p[4].strip() if len(p) > 4 else ''}
                classes[ctx]['inclusiones'].append(detail)

        # Detectar internos cuyo contexto coincide con una clase; solo se consultan campos de contexto.
        assignments = []
        for table in ('sip', 'iax', 'pjsip'):
            c = cols.get(table, set())
            if set(['id', 'keyword', 'data']).issubset(c):
                q = ("SELECT COALESCE(CAST(`id` AS CHAR),''),COALESCE(CAST(`data` AS CHAR),'') FROM `" + table +
                     "` WHERE LOWER(`keyword`)='context' AND COALESCE(`data`,'')<>'';")
                for line in _vp_ruteo_v6_sqlrun(perfil, q):
                    p = line.split('\t')
                    if len(p) >= 2 and p[0].strip().isdigit() and p[1].strip() in classes:
                        assignments.append((p[1].strip(), p[0].strip()))
        for table in ('ps_endpoints',):
            c = cols.get(table, set())
            if set(['id', 'context']).issubset(c):
                q = "SELECT COALESCE(CAST(`id` AS CHAR),''),COALESCE(CAST(`context` AS CHAR),'') FROM `" + table + "` WHERE COALESCE(`context`,'')<>'';"
                for line in _vp_ruteo_v6_sqlrun(perfil, q):
                    p = line.split('\t')
                    if len(p) >= 2 and p[0].strip().isdigit() and p[1].strip() in classes:
                        assignments.append((p[1].strip(), p[0].strip()))
        for ctx, ext in assignments:
            if ext not in classes[ctx]['internos']:
                classes[ctx]['internos'].append(ext)

        result = sorted(classes.values(), key=lambda x: (str(x.get('nombre') or '').lower(), str(x.get('context') or '').lower()))
        return jsonify({
            'ok': True, 'disponible': True, 'perfil': perfil_nombre,
            'total': len(result), 'clases': result,
            'mensaje': 'Leídas directamente de las tablas nativas de clases de Issabel. Consulta de solo lectura.'
        })
    except Exception as exc:
        logger.exception('[COS NATIVO ISSABEL] Error leyendo clases')
        return jsonify({'ok': False, 'mensaje': 'No se pudieron leer las clases nativas de Issabel: ' + str(exc)[:700]}), 500
# FIN VP_ISSABEL_NATIVE_COS_READ_API_V1
