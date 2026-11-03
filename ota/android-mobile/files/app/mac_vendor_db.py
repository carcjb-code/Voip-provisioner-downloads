# -*- coding: utf-8 -*-
"""Consulta offline de fabricante por MAC usando asignaciones IEEE MA-L/MA-M/MA-S."""
from __future__ import print_function
import csv
import datetime
import io
import json
import os
import re
import shutil
import tempfile
import threading
import urllib.request

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BUNDLED_DATABASE_PATH = os.path.join(PROJECT_ROOT, 'data', 'mac_vendor_db.json')
_DATA_FOLDER = os.environ.get('VOIP_PROVISIONER_DATA', '').strip()
DATABASE_PATH = (
    os.path.join(os.path.abspath(_DATA_FOLDER), 'mac_vendor_db.json')
    if _DATA_FOLDER else
    os.environ.get('VOIP_PROVISIONER_MAC_VENDOR_DB_PATH', BUNDLED_DATABASE_PATH)
)
REGISTRIES = (
    ('MA-L', 'https://standards-oui.ieee.org/oui/oui.csv', 6),
    ('MA-M', 'https://standards-oui.ieee.org/oui28/mam.csv', 7),
    ('MA-S', 'https://standards-oui.ieee.org/oui36/oui36.csv', 9),
)
_CACHE_LOCK = threading.RLock()
_CACHE_STAMP = None
_CACHE_VENDORS = {}


def _instalar_base_incluida(reemplazar=False):
    """Copia la base IEEE incluida en el paquete al directorio privado escribible."""
    try:
        if not reemplazar and os.path.isfile(DATABASE_PATH) and os.path.getsize(DATABASE_PATH) > 1024:
            return False
    except OSError:
        pass

    candidatos = [
        os.environ.get('VOIP_PROVISIONER_MAC_VENDOR_DB_SOURCE', '').strip(),
        BUNDLED_DATABASE_PATH,
    ]
    destino_abs = os.path.normcase(os.path.abspath(DATABASE_PATH))
    for candidato in candidatos:
        if not candidato:
            continue
        try:
            origen_abs = os.path.normcase(os.path.abspath(candidato))
            if origen_abs == destino_abs or not os.path.isfile(candidato):
                continue
            if os.path.getsize(candidato) < 1024:
                continue
            carpeta = os.path.dirname(os.path.abspath(DATABASE_PATH)) or '.'
            if not os.path.isdir(carpeta):
                os.makedirs(carpeta, exist_ok=True)
            fd, temporal = tempfile.mkstemp(prefix='.mac_vendor_db_seed_', suffix='.json', dir=carpeta)
            try:
                with os.fdopen(fd, 'wb') as salida:
                    with open(candidato, 'rb') as entrada:
                        shutil.copyfileobj(entrada, salida, 256 * 1024)
                os.replace(temporal, DATABASE_PATH)
                return True
            finally:
                if os.path.isfile(temporal):
                    try:
                        os.remove(temporal)
                    except OSError:
                        pass
        except Exception:
            continue
    return False


def _normalizar_hex(valor):
    return re.sub(r'[^0-9a-f]', '', str(valor or '').lower())


def _organizacion_legible(nombre):
    nombre = str(nombre or '').strip().strip(',')
    if not nombre:
        return ''
    if nombre.upper() == 'IEEE REGISTRATION AUTHORITY':
        return ''
    bajo = nombre.lower()
    # Estas equivalencias mejoran el nombre mostrado; no intervienen en el
    # reconocimiento de la MAC ni en la selección/generación de plantillas.
    marcas = (
        ('grandstream', 'Grandstream'),
        ('fanvil', 'Fanvil'),
        ('yealink', 'Yealink'),
        ('avaya', 'Avaya'),
        ('cisco', 'Cisco'),
        ('polycom', 'Poly'),
        ('plantronics', 'Poly'),
        ('poly, inc', 'Poly'),
        ('snom', 'Snom'),
        ('gigaset', 'Gigaset'),
        ('mitel', 'Mitel'),
        ('panasonic', 'Panasonic'),
        ('audiocodes', 'AudioCodes'),
        ('sangoma', 'Sangoma'),
        ('digium', 'Digium'),
        ('unify', 'Unify'),
        ('alcatel', 'Alcatel'),
        ('escene', 'Escene'),
        ('htek', 'Htek'),
        ('vtech', 'VTech'),
        ('denwa', 'Denwa'),
        ('d-link', 'D-Link'),
        ('dlink', 'D-Link'),
        ('jabra', 'Jabra'),
        ('hewlett packard', 'HP'),
        ('nec corporation', 'NEC'),
        ('nec platforms', 'NEC'),
        ('siemens', 'Siemens'),
        ('broadsoft', 'BroadSoft'),
    )
    for fragmento, marca in marcas:
        if fragmento in bajo:
            return marca
    # IEEE publica bastantes nombres societarios en mayúsculas.
    if nombre == nombre.upper():
        return nombre.title()
    return nombre


def _cargar_vendedores():
    global _CACHE_STAMP, _CACHE_VENDORS
    try:
        st = os.stat(DATABASE_PATH)
    except OSError:
        _instalar_base_incluida()
        try:
            st = os.stat(DATABASE_PATH)
        except OSError:
            return {}
    stamp = (st.st_mtime, st.st_size)
    with _CACHE_LOCK:
        if _CACHE_STAMP == stamp:
            return _CACHE_VENDORS
        for intento in range(2):
            try:
                with open(DATABASE_PATH, 'r', encoding='utf-8') as fh:
                    data = json.load(fh)
                vendors = data.get('vendors', {}) if isinstance(data, dict) else {}
                if not isinstance(vendors, dict) or len(vendors) < 1000:
                    raise ValueError('La base IEEE no tiene un número suficiente de asignaciones.')
                _CACHE_VENDORS = vendors
                st = os.stat(DATABASE_PATH)
                _CACHE_STAMP = (st.st_mtime, st.st_size)
                return _CACHE_VENDORS
            except Exception:
                if intento == 0 and _instalar_base_incluida(reemplazar=True):
                    continue
                logger = __import__('logging').getLogger(__name__)
                logger.warning('No se pudo cargar la base local de fabricantes MAC desde %s.', DATABASE_PATH, exc_info=True)
                _CACHE_VENDORS = {}
                _CACHE_STAMP = None
                return {}


def lookup_mac_vendor(mac):
    """Devuelve la marca/fabricante reconocible; '' si no hay asignación."""
    mac_hex = _normalizar_hex(mac).upper()
    if len(mac_hex) != 12:
        return ''
    vendors = _cargar_vendedores()
    # MA-S (36 bits), luego MA-M (28 bits), por último MA-L (24 bits).
    for length in (9, 7, 6):
        record = vendors.get(mac_hex[:length])
        if not record:
            continue
        if isinstance(record, str):
            organization = record
        else:
            organization = record.get('organization') or record.get('manufacturer') or ''
        # La asignación genérica de IEEE indica que hay que buscar un bloque
        # más específico; no se debe mostrar IEEE como si fuera la marca.
        return _organizacion_legible(organization)
    return ''


def _descargar_registro(url, registro_esperado, longitud):
    request = urllib.request.Request(
        url,
        headers={
            'User-Agent': 'VoIP-Provisioner-MAC-Registry/1.0',
            'Accept': 'text/csv,text/plain,*/*',
        },
    )
    with urllib.request.urlopen(request, timeout=30) as response:
        raw = response.read(25 * 1024 * 1024 + 1)
    if not raw or len(raw) > 25 * 1024 * 1024:
        raise RuntimeError('El registro {} está vacío o excede el tamaño esperado.'.format(registro_esperado))
    texto = raw.decode('utf-8-sig', 'replace')
    if '<html' in texto[:1000].lower() or '<!doctype' in texto[:1000].lower():
        raise RuntimeError('IEEE devolvió HTML en lugar de CSV para {}.'.format(registro_esperado))
    reader = csv.DictReader(io.StringIO(texto))
    if not reader.fieldnames:
        raise RuntimeError('No se reconocieron los encabezados CSV de {}.'.format(registro_esperado))
    campos = {}
    for field in reader.fieldnames:
        campos[re.sub(r'[^a-z]', '', str(field or '').lower())] = field
    asignacion_key = campos.get('assignment')
    organizacion_key = campos.get('organizationname') or campos.get('organization') or campos.get('manufacturer')
    registro_key = campos.get('registry')
    direccion_key = campos.get('organizationaddress') or campos.get('address')
    if not asignacion_key or not organizacion_key:
        raise RuntimeError('El CSV de {} no tiene las columnas Assignment / Organization Name.'.format(registro_esperado))
    records = {}
    for row in reader:
        registry = str(row.get(registro_key, '') if registro_key else registro_esperado).strip().upper()
        if registry and registry != registro_esperado:
            continue
        prefix = _normalizar_hex(row.get(asignacion_key, ''))
        organization = str(row.get(organizacion_key, '') or '').strip()
        if len(prefix) != longitud or not organization:
            continue
        records[prefix.upper()] = {
            'registry': registro_esperado,
            'organization': organization,
            'address': str(row.get(direccion_key, '') or '').strip() if direccion_key else '',
        }
    # Evita reemplazar la base con una respuesta parcial, página de error o
    # cambio inesperado del formato de IEEE.
    if len(records) < 100:
        raise RuntimeError('El CSV de {} solo contiene {} asignaciones válidas; se conserva la base anterior.'.format(registro_esperado, len(records)))
    return records



def refresh_ieee_database_from_files(files_by_registry):
    """Construye y reemplaza la base usando CSV IEEE cargados desde una PC.

    files_by_registry: dict MA-L/MA-M/MA-S -> bytes CSV.
    Nunca reemplaza la base hasta que los tres registros pasan validacion.
    """
    specs = (
        ('MA-L', 6),
        ('MA-M', 7),
        ('MA-S', 9),
    )
    all_vendors = {}
    counts = {}
    for registry, prefix_len in specs:
        raw = files_by_registry.get(registry)
        if not raw:
            raise RuntimeError('Falta el archivo CSV del registro {}.'.format(registry))
        if isinstance(raw, str):
            raw = raw.encode('utf-8')
        if len(raw) > 25 * 1024 * 1024:
            raise RuntimeError('El CSV {} excede el tamaño máximo permitido.'.format(registry))
        texto = raw.decode('utf-8-sig', 'replace')
        if '<html' in texto[:1000].lower() or '<!doctype' in texto[:1000].lower():
            raise RuntimeError('El archivo de {} parece HTML, no un CSV IEEE.'.format(registry))
        reader = csv.DictReader(io.StringIO(texto))
        if not reader.fieldnames:
            raise RuntimeError('No se reconocieron encabezados CSV para {}.'.format(registry))
        campos = {}
        for field in reader.fieldnames:
            campos[re.sub(r'[^a-z]', '', str(field or '').lower())] = field
        assignment_key = campos.get('assignment')
        organization_key = (campos.get('organizationname') or
                            campos.get('organization') or
                            campos.get('manufacturer'))
        registry_key = campos.get('registry')
        address_key = campos.get('organizationaddress') or campos.get('address')
        if not assignment_key or not organization_key:
            raise RuntimeError('El CSV {} no tiene las columnas Assignment / Organization Name.'.format(registry))
        records = {}
        for row in reader:
            registry_value = str(row.get(registry_key, '') if registry_key else registry).strip().upper()
            if registry_value and registry_value != registry:
                continue
            prefix = _normalizar_hex(row.get(assignment_key, ''))
            organization = str(row.get(organization_key, '') or '').strip()
            if len(prefix) != prefix_len or not organization:
                continue
            records[prefix.upper()] = {
                'registry': registry,
                'organization': organization,
                'address': str(row.get(address_key, '') or '').strip() if address_key else '',
            }
        if len(records) < 100:
            raise RuntimeError('El CSV {} solo contiene {} asignaciones válidas; se conserva la base anterior.'.format(registry, len(records)))
        counts[registry] = len(records)
        all_vendors.update(records)

    if len(all_vendors) < 1000:
        raise RuntimeError('La base consolidada es demasiado pequeña; se conserva la base anterior.')

    now = datetime.datetime.now().astimezone().isoformat()
    urls = [
        'https://standards-oui.ieee.org/oui/oui.csv',
        'https://standards-oui.ieee.org/oui28/mam.csv',
        'https://standards-oui.ieee.org/oui36/oui36.csv',
    ]
    data = {
        'schema_version': 1,
        'source': 'IEEE Registration Authority - registros MA-L, MA-M y MA-S (cargados desde el navegador)',
        'source_urls': urls,
        'updated_at': now,
        'counts': counts,
        'total_entries': len(all_vendors),
        'vendors': all_vendors,
    }
    folder = os.path.dirname(DATABASE_PATH)
    if not os.path.isdir(folder):
        os.makedirs(folder, exist_ok=True)
    fd, temp_path = tempfile.mkstemp(prefix='.mac_vendor_db_upload_', suffix='.json', dir=folder)
    try:
        with os.fdopen(fd, 'w', encoding='utf-8') as fh:
            json.dump(data, fh, ensure_ascii=False, separators=(',', ':'))
            fh.write('\n')
        with open(temp_path, 'r', encoding='utf-8') as fh:
            check = json.load(fh)
        if len(check.get('vendors', {})) < 1000:
            raise RuntimeError('La base temporal no pasó la validación.')
        os.replace(temp_path, DATABASE_PATH)
        temp_path = None
    finally:
        if temp_path and os.path.isfile(temp_path):
            try:
                os.remove(temp_path)
            except OSError:
                pass
    global _CACHE_STAMP, _CACHE_VENDORS
    with _CACHE_LOCK:
        _CACHE_STAMP = None
        _CACHE_VENDORS = {}
    return data


def refresh_ieee_database():
    """Descarga los tres CSV IEEE y sustituye el JSON local solo si todos sirven."""
    all_vendors = {}
    counts = {}
    for registry, url, prefix_len in REGISTRIES:
        records = _descargar_registro(url, registry, prefix_len)
        counts[registry] = len(records)
        all_vendors.update(records)
    if len(all_vendors) < 1000:
        raise RuntimeError('La base consolidada es demasiado pequeña; se conserva la versión anterior.')

    now = datetime.datetime.now().astimezone().isoformat()
    data = {
        'schema_version': 1,
        'source': 'IEEE Registration Authority - registros MA-L, MA-M y MA-S',
        'source_urls': [item[1] for item in REGISTRIES],
        'updated_at': now,
        'counts': counts,
        'total_entries': len(all_vendors),
        'vendors': all_vendors,
    }
    folder = os.path.dirname(DATABASE_PATH)
    if not os.path.isdir(folder):
        os.makedirs(folder, exist_ok=True)
    fd, temp_path = tempfile.mkstemp(prefix='.mac_vendor_db_', suffix='.json', dir=folder)
    try:
        with os.fdopen(fd, 'w', encoding='utf-8') as fh:
            json.dump(data, fh, ensure_ascii=False, separators=(',', ':'))
            fh.write('\n')
        # Comprueba que el fichero temporal sea JSON legible y tenga registros.
        with open(temp_path, 'r', encoding='utf-8') as fh:
            check = json.load(fh)
        if len(check.get('vendors', {})) < 1000:
            raise RuntimeError('La base temporal no pasó la validación.')
        os.replace(temp_path, DATABASE_PATH)
        temp_path = None
    finally:
        if temp_path and os.path.isfile(temp_path):
            try:
                os.remove(temp_path)
            except OSError:
                pass
    # Obliga a recargar la base en el siguiente lookup.
    global _CACHE_STAMP, _CACHE_VENDORS
    with _CACHE_LOCK:
        _CACHE_STAMP = None
        _CACHE_VENDORS = {}
    return data
