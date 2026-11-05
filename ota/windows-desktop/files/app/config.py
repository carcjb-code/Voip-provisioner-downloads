import os
import sys
import json
import logging

logger = logging.getLogger(__name__)

COLUMNAS_EXCEL = ["Edificio", "Piso", "Dependencia", "Nombre Mostrar", "Nombre Saliente", "Interno", "MAC", "IP", "Numero Serie", "Modelo Telefono"]

# Catálogo exclusivo de plantillas válidas. El nombre y la extensión deben coincidir.
PLANTILLAS_PERMITIDAS = {
    "GrandstreamGXP1628": "xml",
    "Fanvil V61G": "cfg",
    "Denwa DW-320G": "cfg",
}

def get_data_folder():
    if sys.platform == 'win32':
        base = os.environ.get('APPDATA', os.path.expanduser('~'))
        folder = os.path.join(base, 'VoIP_Provisioner')
    else:
        base = os.path.expanduser('~')
        folder = os.path.join(base, '.config', 'VoIP_Provisioner')
    os.makedirs(folder, exist_ok=True)
    return folder

# Los clientes instalables mantienen sus datos en el perfil local del usuario.
if os.environ.get("VOIP_PROVISIONER_DATA"):
    DATA_FOLDER = os.environ["VOIP_PROVISIONER_DATA"]
elif sys.platform == "win32" or os.environ.get("VOIP_PROVISIONER_DESKTOP") == "1":
    DATA_FOLDER = os.path.join(get_data_folder(), "data")
else:
    DATA_FOLDER = "/opt/voip-provisioner/data"
CONFIG_FILE = os.path.join(DATA_FOLDER, "config.json")
SFTP_SERVERS_FILE = os.path.join(DATA_FOLDER, "sftp_servers.json")
CARPETA_CONFIGS = os.path.join(DATA_FOLDER, "configs")
PLANTILLAS_FOLDER = os.path.join(DATA_FOLDER, "plantillas")
PLANTILLAS_ESTADO_FILE = os.path.join(DATA_FOLDER, "plantillas_estado.json")
UBICACIONES_PERFILES_FILE = os.path.join(DATA_FOLDER, "ubicaciones_perfiles.json")
os.makedirs(CARPETA_CONFIGS, exist_ok=True)
os.makedirs(PLANTILLAS_FOLDER, exist_ok=True)

DEFAULT_CONFIG = {
    "empresa": {"nombre": "Mi Empresa", "logo": "", "color_primario": "#4da6ff", "soporte_email": "soporte@empresa.com", "prefijo_password": "emp*"},
    "seguridad": {"admin_password": "admin123"},
    "aplicacion": {"titulo": "VoIP Provisioner", "puerto": 8080, "usar_ssl": False, "modulos": {"softphone": False, "monitoreo": True}}
}

def cargar_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                cfg = json.load(f)
                for k, v in DEFAULT_CONFIG.items():
                    if k not in cfg: cfg[k] = v
                    elif isinstance(v, dict):
                        for sk, sv in v.items():
                            if sk not in cfg[k]: cfg[k][sk] = sv
                return cfg
        except Exception as e:
            logger.error(f"Error config: {e}")
    guardar_config(DEFAULT_CONFIG)
    return DEFAULT_CONFIG.copy()

def guardar_config(cfg):
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(cfg, f, indent=2, ensure_ascii=False)
    except Exception as e:
        logger.error(f"Error guardando config: {e}")

def resource_path(relative_path):
    try: return os.path.join(sys._MEIPASS, relative_path)
    except Exception: return os.path.join(os.path.abspath("."), relative_path)

def tools_path(filename):
    if getattr(sys, 'frozen', False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.abspath(".")
    for ruta in [os.path.join(base, filename), os.path.join(base, "tools", filename)]:
        if os.path.exists(ruta): return ruta
    return filename

def cargar_plantillas_eliminadas():
    """Devuelve las plantillas que el usuario decidió quitar del catálogo."""
    try:
        with open(PLANTILLAS_ESTADO_FILE, 'r', encoding='utf-8') as fh:
            data = json.load(fh)
        if isinstance(data, dict):
            data = data.get('eliminadas', [])
        if not isinstance(data, list):
            return set()
        return set(str(nombre) for nombre in data if str(nombre) in PLANTILLAS_PERMITIDAS)
    except FileNotFoundError:
        return set()
    except Exception as e:
        logger.error("Error leyendo el estado de plantillas: %s", e)
        return set()


def guardar_plantillas_eliminadas(nombres):
    """Guarda de forma atómica las bajas de plantillas para no reponerlas al reiniciar."""
    temporal = PLANTILLAS_ESTADO_FILE + '.tmp'
    try:
        os.makedirs(DATA_FOLDER, exist_ok=True)
        permitidas = sorted(set(str(nombre) for nombre in nombres if str(nombre) in PLANTILLAS_PERMITIDAS))
        with open(temporal, 'w', encoding='utf-8') as fh:
            json.dump({'eliminadas': permitidas}, fh, indent=2, ensure_ascii=False)
        os.replace(temporal, PLANTILLAS_ESTADO_FILE)
        return True
    except Exception as e:
        logger.error("Error guardando el estado de plantillas: %s", e)
        try:
            if os.path.exists(temporal):
                os.remove(temporal)
        except Exception:
            pass
        return False

def cargar_plantillas():
    """Carga plantillas válidas sin revivir las que el usuario eliminó."""
    plantillas = {}
    eliminadas = cargar_plantillas_eliminadas()

    def leer_carpeta(carpeta, sobrescribir=False):
        if not carpeta or not os.path.isdir(carpeta):
            return
        try:
            nombres = sorted(os.listdir(carpeta), key=lambda x: x.lower())
        except Exception:
            return
        for archivo in nombres:
            extension = os.path.splitext(archivo)[1].lower()
            if extension not in ('.cfg', '.xml'):
                continue
            nombre = os.path.splitext(archivo)[0]
            # Solo ofrecer los tres modelos/formatos que están probados.
            # Un borrado queda guardado en el estado local y prevalece sobre
            # los archivos incluidos dentro del ejecutable/APK.
            if PLANTILLAS_PERMITIDAS.get(nombre) != extension[1:] or nombre in eliminadas:
                continue
            ruta = os.path.join(carpeta, archivo)
            try:
                with open(ruta, 'r', encoding='utf-8') as fh:
                    contenido = fh.read()
                if sobrescribir or nombre not in plantillas:
                    plantillas[nombre] = contenido
            except Exception as e:
                logger.warning("No se pudo leer la plantilla %s: %s", ruta, e)

    # Las plantillas administradas por el usuario tienen prioridad.
    leer_carpeta(PLANTILLAS_FOLDER, sobrescribir=True)
    candidatos = [
        os.environ.get("VOIP_PROVISIONER_BUNDLED_PLANTILLAS", ""),
        os.environ.get("VOIP_PROVISIONER_BUNDLED_PLANTILLAS_EXTRA", ""),
    ]
    try:
        candidatos.extend([resource_path(os.path.join("data", "plantillas")), resource_path("plantillas")])
    except Exception:
        pass
    raiz_proyecto = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    candidatos.extend([os.path.join(raiz_proyecto, "data", "plantillas"), os.path.join(raiz_proyecto, "plantillas")])
    vistos = set()
    for carpeta in candidatos:
        if not carpeta:
            continue
        carpeta = os.path.abspath(carpeta)
        if carpeta in vistos or carpeta == os.path.abspath(PLANTILLAS_FOLDER):
            continue
        vistos.add(carpeta)
        leer_carpeta(carpeta, sobrescribir=False)

    # Devolver exactamente los modelos permitidos y en el orden acordado.
    return {nombre: plantillas[nombre] for nombre in PLANTILLAS_PERMITIDAS if nombre in plantillas}


def cargar_perfiles():
    if os.path.exists(SFTP_SERVERS_FILE):
        try: return json.load(open(SFTP_SERVERS_FILE, 'r', encoding='utf-8'))
        except Exception: pass
    return []

def guardar_perfiles(perfiles):
    try: json.dump(perfiles, open(SFTP_SERVERS_FILE, 'w', encoding='utf-8'), indent=2, ensure_ascii=False)
    except Exception: pass


def _normalizar_ubicaciones(data):
    """Normaliza catálogos antiguos/nuevos sin usar el Excel."""
    if not isinstance(data, dict):
        return {}
    out = {}
    for perfil, valor in data.items():
        if not isinstance(valor, dict):
            continue
        # Compatibilidad con formato antiguo: edificio/piso/dependencia
        # y formato plural: edificios/pisos/dependencias.
        perfil_out = {}
        for singular, plural in (("edificio", "edificios"), ("piso", "pisos"), ("dependencia", "dependencias")):
            vals = valor.get(singular, valor.get(plural, []))
            if not isinstance(vals, list):
                vals = []
            limpio = []
            for v in vals:
                v = str(v).strip()
                if v and v not in limpio:
                    limpio.append(v)
            perfil_out[singular] = sorted(limpio, key=lambda x: x.casefold())
        out[str(perfil)] = perfil_out
    return out

def cargar_ubicaciones_perfiles():
    if os.path.exists(UBICACIONES_PERFILES_FILE):
        try:
            with open(UBICACIONES_PERFILES_FILE, 'r', encoding='utf-8') as f:
                return _normalizar_ubicaciones(json.load(f))
        except Exception as e:
            logger.error(f"Error cargando ubicaciones por perfil: {e}")
    return {}

def guardar_ubicaciones_perfiles(data):
    try:
        os.makedirs(DATA_FOLDER, exist_ok=True)
        normalizado = _normalizar_ubicaciones(data)
        tmp = UBICACIONES_PERFILES_FILE + '.tmp'
        with open(tmp, 'w', encoding='utf-8') as f:
            json.dump(normalizado, f, indent=2, ensure_ascii=False)
        os.replace(tmp, UBICACIONES_PERFILES_FILE)
        return True
    except Exception as e:
        logger.error(f"Error guardando ubicaciones por perfil: {e}")
        return False


def buscar_perfil(nombre, perfiles=None):
    if perfiles is None: perfiles = cargar_perfiles()
    if not nombre: return perfiles[0] if perfiles else None
    return next((p for p in perfiles if p['nombre'] == nombre), None)

def _excel_basename(nombre):
    if not nombre:
        return "Registro_default.xlsx"
    p = buscar_perfil(nombre)
    if p and 'host' in p:
        host = str(p['host']).replace('.', '').replace(':', '')
    else:
        host = str(nombre).replace(' ', '').replace('.', '').replace(':', '')
    return "Registro_{}.xlsx".format(host)


def get_excel_file_legacy(nombre):
    """Ruta usada por versiones anteriores, solo para migrar el Excel existente."""
    return os.path.join(DATA_FOLDER, _excel_basename(nombre))


def get_excel_file(nombre):
    # En los clientes instalables el archivo de trabajo es una caché temporal.
    # El maestro es el archivo del servidor remoto configurado por SFTP.
    # En Issabel se conserva DATA_FOLDER (/opt/voip-provisioner/data).
    if os.environ.get("VOIP_PROVISIONER_REMOTE_EXCEL") == "1":
        base_dir = os.environ.get(
            "VOIP_PROVISIONER_EXCEL_WORKDIR",
            os.path.join(DATA_FOLDER, "excel-cache")
        )
    else:
        base_dir = DATA_FOLDER
    os.makedirs(base_dir, exist_ok=True)
    return os.path.join(base_dir, _excel_basename(nombre))

ALARMAS_PERFILES = {}