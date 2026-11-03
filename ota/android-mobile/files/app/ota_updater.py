# -*- coding: utf-8 -*-
"""Actualizador OTA desde el repositorio publico de descargas.

Descarga solo archivos Python/interfaz/plantillas permitidos, valida SHA-256
y prepara una carpeta reversible dentro del almacenamiento privado de la app.
Nunca actualiza perfiles SSH, contrasenas, Excel ni datos del usuario.
Compatible con Python 3.6+.
"""
from __future__ import print_function

import hashlib
import json
import logging
import os
import posixpath
import shutil
import time
import urllib.error
import urllib.parse
import urllib.request

logger = logging.getLogger(__name__)

PUBLIC_REPOSITORY = 'carcjb-code/Voip-provisioner-downloads'
PUBLIC_BRANCH = 'main'
RAW_ROOT = 'https://raw.githubusercontent.com/{}/{}/'.format(
    PUBLIC_REPOSITORY, PUBLIC_BRANCH
)
MAX_FILE_BYTES = 2 * 1024 * 1024
MAX_TOTAL_BYTES = 20 * 1024 * 1024
MAX_FILES_PER_UPDATE = 250
_PACKAGE_INIT = (
    'from pkgutil import extend_path\n'
    '__path__ = extend_path(__path__, __name__)\n'
)
_STATIC_SUFFIXES = (
    '.html', '.css', '.js', '.svg', '.png', '.webp', '.ico',
    '.jpg', '.jpeg', '.gif', '.woff', '.woff2', '.ttf'
)


def _platform():
    if os.environ.get('VOIP_PROVISIONER_ANDROID') == '1':
        return 'android-mobile'
    if os.environ.get('VOIP_PROVISIONER_WINDOWS') == '1':
        return 'windows-desktop'
    # Previous Windows builds already set DESKTOP=1. Android sets ANDROID=1
    # before this fallback, so older desktop builds remain recognizable.
    if os.environ.get('VOIP_PROVISIONER_DESKTOP') == '1':
        return 'windows-desktop'
    return ''


def updater_enabled():
    return bool(_platform())


def _update_root():
    configured = os.environ.get('VOIP_PROVISIONER_UPDATE_DIR', '').strip()
    if configured:
        root = os.path.abspath(configured)
    else:
        data_root = os.environ.get('VOIP_PROVISIONER_DATA', '').strip()
        if data_root:
            root = os.path.join(os.path.abspath(data_root), 'updates')
        else:
            root = os.path.join(os.path.abspath(os.getcwd()), '.app-updates')
    if not os.path.isdir(root):
        os.makedirs(root, exist_ok=True)
    try:
        os.chmod(root, 0o700)
    except Exception:
        pass
    return root


def _build_info():
    info_path = os.environ.get('VOIP_PROVISIONER_BUILD_INFO', '').strip()
    if not info_path:
        return {}
    try:
        with open(info_path, 'r', encoding='utf-8') as fh:
            data = json.load(fh)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def _build_commit():
    info = _build_info()
    sha = str(info.get('commit') or info.get('source_commit') or '').strip()
    if len(sha) == 40:
        return sha
    sha = str(os.environ.get('VOIP_PROVISIONER_BUILD_SHA', '') or '').strip()
    return sha if len(sha) == 40 else ''


def _active_manifest():
    path = os.path.join(_update_root(), 'current', 'manifest.json')
    try:
        with open(path, 'r', encoding='utf-8') as fh:
            data = json.load(fh)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def _active_commit():
    sha = str(_active_manifest().get('source_commit', '') or '').strip()
    return sha if len(sha) == 40 else _build_commit()


def _installed_hashes():
    hashes = {}
    # Fallback for Android installations that predate build-time file hashes:
    # fingerprint the app's bundled files directly so compatible old versions
    # can still update without reinstalling.
    info_path = os.environ.get('VOIP_PROVISIONER_BUILD_INFO', '').strip()
    if info_path:
        site_root = os.path.dirname(os.path.dirname(os.path.abspath(info_path)))
        for folder_name in ('app', 'templates', 'static', os.path.join('data', 'plantillas')):
            folder = os.path.join(site_root, folder_name)
            if not os.path.isdir(folder):
                continue
            for base, dirs, names in os.walk(folder):
                dirs[:] = [d for d in dirs if d not in ('__pycache__', '.git')]
                for name in names:
                    full = os.path.join(base, name)
                    rel = os.path.relpath(full, site_root).replace(os.sep, '/')
                    if not _updateable_path(rel):
                        continue
                    try:
                        digest = hashlib.sha256()
                        with open(full, 'rb') as fh:
                            for chunk in iter(lambda: fh.read(65536), b''):
                                digest.update(chunk)
                        hashes[rel] = digest.hexdigest()
                    except Exception:
                        logger.debug('No se pudo calcular el hash local de %s', rel, exc_info=True)

    built = _build_info().get('file_hashes', {})
    if isinstance(built, dict):
        hashes.update(built)
    active = _active_manifest().get('file_hashes', {})
    if isinstance(active, dict):
        hashes.update(active)
    cleaned = {}
    for path, digest in hashes.items():
        path = _normalize_path(path)
        digest = str(digest or '').lower()
        if path and _updateable_path(path) and len(digest) == 64:
            cleaned[path] = digest
    return cleaned


def _manifest_url():
    platform = _platform()
    if not platform:
        raise RuntimeError('El actualizador no está habilitado para esta plataforma.')
    # A fresh URL prevents stale intermediary caches from hiding new manifests.
    return RAW_ROOT + 'ota/{}/manifest.json?ts={}'.format(platform, int(time.time()))


def _request_bytes(url, timeout=30, limit=MAX_TOTAL_BYTES):
    headers = {
        'Accept': 'application/octet-stream',
        'User-Agent': 'VoIP-Provisioner-Public-Updater/2.0',
        'Cache-Control': 'no-cache',
    }
    req = urllib.request.Request(url, headers=headers)
    try:
        with urllib.request.urlopen(req, timeout=timeout) as response:
            raw = response.read(limit + 1)
            status = getattr(response, 'status', 200)
    except urllib.error.HTTPError as exc:
        if exc.code == 404:
            raise RuntimeError('Todavía no hay una publicación OTA para esta plataforma en el repositorio público.')
        raise RuntimeError('GitHub respondió HTTP {} al descargar la actualización.'.format(exc.code))
    except Exception as exc:
        raise RuntimeError('No se pudo conectar con el repositorio público: {}'.format(str(exc)[:220]))
    if status < 200 or status >= 300:
        raise RuntimeError('El repositorio público respondió HTTP {}.'.format(status))
    if len(raw) > limit:
        raise RuntimeError('La descarga supera el límite seguro de {} bytes.'.format(limit))
    return raw


def _load_public_manifest():
    raw = _request_bytes(_manifest_url(), limit=2 * 1024 * 1024)
    try:
        manifest = json.loads(raw.decode('utf-8'))
    except Exception:
        raise RuntimeError('El manifiesto público de actualización no es JSON válido.')
    if not isinstance(manifest, dict) or manifest.get('schema_version') != 1:
        raise RuntimeError('El manifiesto público tiene una versión no compatible.')
    if manifest.get('platform') != _platform():
        raise RuntimeError('El manifiesto público no corresponde a esta plataforma.')
    commit = str(manifest.get('source_commit') or '').strip()
    if len(commit) != 40:
        raise RuntimeError('El manifiesto público no identifica una versión válida.')
    files = manifest.get('files')
    if not isinstance(files, list) or len(files) > MAX_FILES_PER_UPDATE:
        raise RuntimeError('El manifiesto público contiene una lista de archivos no válida.')
    return manifest


def _normalize_path(path):
    path = str(path or '').replace('\\', '/').strip()
    path = posixpath.normpath(path)
    if not path or path == '.' or path.startswith('/') or path == '..' or path.startswith('../'):
        return ''
    if any(part in ('', '.', '..') for part in path.split('/')):
        return ''
    return path


def _site_file_allowed(path):
    return path.startswith(('templates/', 'static/', 'data/plantillas/'))


def _updateable_path(path):
    path = _normalize_path(path)
    if not path:
        return False
    lower = path.lower()
    if path.startswith('app/') and lower.endswith('.py') and path != 'app/__init__.py':
        return True
    if path.startswith('templates/') and lower.endswith(('.html', '.css', '.js')):
        return True
    if path.startswith('static/') and lower.endswith(_STATIC_SUFFIXES):
        return True
    if path.startswith('data/plantillas/') and lower.endswith(('.cfg', '.xml')):
        return True
    return False


def _git_blob_sha(content):
    if isinstance(content, str):
        content = content.encode('utf-8')
    header = b'blob ' + str(len(content)).encode('ascii') + b'\0'
    return hashlib.sha1(header + content).hexdigest()


def _sha256(content):
    return hashlib.sha256(content).hexdigest()


def _compare():
    if not updater_enabled():
        raise RuntimeError('El actualizador público está disponible solo en Android y Windows.')
    manifest = _load_public_manifest()
    latest = manifest['source_commit']
    installed = _active_commit()
    local_hashes = _installed_hashes()

    history = manifest.get('history', [])
    if not isinstance(history, list):
        history = []
    installed_index = -1
    for index, item in enumerate(history):
        if str((item or {}).get('source_commit', '') or '') == installed:
            installed_index = index
            break

    full_update_paths = []
    if installed != latest:
        if not installed or installed_index < 0:
            full_update_paths.append('Esta instalación todavía no está registrada en el historial OTA público.')
        else:
            for item in history[installed_index + 1:]:
                if (item or {}).get('requires_full_update'):
                    full_update_paths.extend((item or {}).get('full_update_paths', []) or [])
                    if not full_update_paths:
                        full_update_paths.append('La versión intermedia requiere una actualización completa.')

    remote_files = {}
    for item in manifest.get('files', []):
        if not isinstance(item, dict):
            continue
        path = _normalize_path(item.get('path', ''))
        digest = str(item.get('sha256', '') or '').lower()
        size = int(item.get('size', 0) or 0)
        if not path or not _updateable_path(path) or len(digest) != 64 or size < 0:
            raise RuntimeError('El manifiesto contiene una ruta o hash no permitido.')
        if size > MAX_FILE_BYTES:
            full_update_paths.append('El archivo supera el límite OTA: ' + path)
            continue
        remote_files[path] = {'path': path, 'sha256': digest, 'size': size}

    changed = []
    for path, item in remote_files.items():
        if local_hashes.get(path) != item['sha256']:
            changed.append(item)

    deleted_paths = sorted(set(local_hashes.keys()) - set(remote_files.keys()))
    unsafe_deletions = [p for p in deleted_paths if not _site_file_allowed(p)]
    if unsafe_deletions:
        full_update_paths.extend(unsafe_deletions)

    if len(changed) > MAX_FILES_PER_UPDATE:
        full_update_paths.append('La cantidad de archivos supera el límite de actualización parcial.')

    return {
        'status': 'ok',
        'enabled': True,
        'platform': _platform(),
        'installed_commit': installed,
        'latest_commit': latest,
        'update_available': bool(changed or deleted_paths),
        'files': changed,
        'deleted_paths': deleted_paths,
        'requires_full_update': bool(full_update_paths),
        'requires_full_update_paths': sorted(set(full_update_paths)),
        'manifest': manifest,
        'message': (
            'Hay cambios compatibles que se pueden aplicar desde el repositorio público.'
            if (changed or deleted_paths) else
            'La aplicación ya tiene los mismos archivos que la publicación pública.'
        ),
    }


def check_for_updates():
    result = _compare()
    files = result.get('files', [])
    deleted = result.get('deleted_paths', [])
    required = result.get('requires_full_update_paths', [])
    result['files'] = [item['path'] for item in files[:50]]
    result['file_count'] = len(files)
    result['deleted_count'] = len(deleted)
    result['deleted_paths'] = deleted[:50]
    result['requires_full_update_count'] = len(required)
    result['requires_full_update_paths'] = required[:30]
    # Compatibility with the previous Android screen until all clients are upgraded.
    result['requires_apk_count'] = len(required)
    result['requires_apk'] = required[:30]
    result.pop('manifest', None)
    return result


def _fetch_file(path, platform, expected_sha):
    quoted = urllib.parse.quote(path, safe='/')
    url = RAW_ROOT + 'ota/{}/files/{}?sha256={}'.format(
        platform, quoted, urllib.parse.quote(expected_sha, safe='')
    )
    raw = _request_bytes(url, limit=MAX_FILE_BYTES)
    if _sha256(raw) != expected_sha:
        raise RuntimeError('La verificación SHA-256 no coincide para el archivo {}.'.format(path))
    if path.lower().endswith('.py'):
        try:
            compile(raw.decode('utf-8'), path, 'exec')
        except Exception as exc:
            raise RuntimeError('El archivo Python {} no pasó la validación de sintaxis: {}'.format(path, str(exc)[:180]))
    return raw


def _make_tree_writable(root):
    """Normaliza permisos de la copia temporal sin tocar la version activa."""
    if not os.path.isdir(root):
        os.makedirs(root, exist_ok=True)
    for base, dirs, names in os.walk(root):
        try:
            os.chmod(base, 0o700)
        except Exception:
            pass
        for name in names:
            full = os.path.join(base, name)
            try:
                if not os.path.islink(full):
                    os.chmod(full, 0o600)
            except Exception:
                pass
    try:
        os.chmod(root, 0o700)
    except Exception:
        pass


def _write_checked(stage, path, content):
    path = _normalize_path(path)
    if not path or not _updateable_path(path):
        raise RuntimeError('Ruta no permitida para actualizar: ' + str(path))
    destination = os.path.abspath(os.path.join(stage, *path.split('/')))
    stage_abs = os.path.abspath(stage) + os.sep
    if not destination.startswith(stage_abs):
        raise RuntimeError('Ruta fuera del directorio privado de actualizaciones.')
    parent = os.path.dirname(destination)
    if not os.path.isdir(parent):
        os.makedirs(parent, exist_ok=True)
    try:
        os.chmod(parent, 0o700)
    except Exception:
        pass
    with open(destination + '.tmp', 'wb') as fh:
        fh.write(content)
    os.replace(destination + '.tmp', destination)


def apply_updates():
    """Descarga archivos diferenciales y promueve el conjunto validado de forma atomica."""
    preview = _compare()
    if preview.get('requires_full_update'):
        examples = ', '.join(preview.get('requires_full_update_paths', [])[:4])
        raise RuntimeError(
            'Esta versión necesita una actualización completa mediante APK/instalador. '
            'No se aplicó ningún archivo parcialmente. Motivo: ' + examples
        )
    if not preview.get('update_available'):
        return {
            'status': 'ok',
            'update_available': False,
            'message': preview.get('message', 'La aplicación ya está al día.'),
            'latest_commit': preview.get('latest_commit', ''),
            'downloaded_count': 0,
            'profiles_preserved': True,
        }

    root = _update_root()
    current = os.path.join(root, 'current')
    previous = os.path.join(root, 'previous')
    stage = os.path.join(root, '.staging-{}-{}'.format(os.getpid(), int(time.time() * 1000)))
    manifest = preview['manifest']
    platform = _platform()
    local_hashes = _installed_hashes()
    remote_files = {item['path']: item for item in manifest.get('files', [])}
    deleted_paths = sorted(set(preview.get('deleted_paths', [])))
    total_bytes = 0
    downloaded = []

    try:
        if os.path.isdir(current):
            shutil.copytree(current, stage)
        else:
            os.makedirs(stage, exist_ok=True)
        # copytree conserva permisos del arbol anterior; si una version dejo
        # carpetas/archivos de solo lectura, la nueva etapa no podia escribirse.
        _make_tree_writable(stage)

        for path in deleted_paths:
            existing = os.path.abspath(os.path.join(stage, *path.split('/')))
            if existing.startswith(os.path.abspath(stage) + os.sep) and os.path.isfile(existing):
                os.remove(existing)

        for item in preview.get('files', []):
            path = item['path']
            if local_hashes.get(path) == item['sha256']:
                continue
            content = _fetch_file(path, platform, item['sha256'])
            total_bytes += len(content)
            if total_bytes > MAX_TOTAL_BYTES:
                raise RuntimeError('La actualización supera el límite seguro de 20 MB.')
            _write_checked(stage, path, content)
            downloaded.append(path)

        app_folder = os.path.join(stage, 'app')
        os.makedirs(app_folder, exist_ok=True)
        init_path = os.path.join(app_folder, '__init__.py')
        with open(init_path, 'w', encoding='utf-8') as fh:
            fh.write(_PACKAGE_INIT)

        manifest_files = sorted(remote_files.keys())
        local_manifest = {
            'schema_version': 1,
            'source_commit': manifest['source_commit'],
            'platform': platform,
            'updated_at': int(time.time()),
            'files': manifest_files,
            'file_hashes': dict((p, remote_files[p]['sha256']) for p in manifest_files),
            'deleted_paths': deleted_paths,
            'downloaded_this_update': downloaded,
        }
        manifest_path = os.path.join(stage, 'manifest.json')
        with open(manifest_path + '.tmp', 'w', encoding='utf-8') as fh:
            json.dump(local_manifest, fh, ensure_ascii=False, indent=2)
            fh.write('\n')
        os.replace(manifest_path + '.tmp', manifest_path)

        if os.path.isdir(previous):
            shutil.rmtree(previous)
        if os.path.isdir(current):
            os.replace(current, previous)
        try:
            os.replace(stage, current)
        except Exception:
            if os.path.isdir(previous) and not os.path.exists(current):
                os.replace(previous, current)
            raise

        return {
            'status': 'ok',
            'update_available': True,
            'latest_commit': manifest['source_commit'],
            'downloaded_count': len(downloaded),
            'downloaded_files': downloaded[:50],
            'bytes_downloaded': total_bytes,
            'restart_required': True,
            'profiles_preserved': True,
            'message': (
                'Actualización descargada y verificada desde el repositorio público. '
                'No se modificaron servidores, contraseñas ni datos del usuario. '
                'Cerrá completamente VoIP Provisioner y volvé a abrirlo para activarla.'
            ),
        }
    except Exception:
        logger.exception('No se pudo preparar la actualización OTA pública.')
        if os.path.isdir(stage):
            try:
                shutil.rmtree(stage)
            except Exception:
                pass
        raise


def updater_status():
    enabled = updater_enabled()
    active_sha = _active_commit()
    return {
        'enabled': enabled,
        'platform': _platform() or 'not-supported',
        'installed_commit': active_sha,
        'repository': PUBLIC_REPOSITORY,
        'message': (
            'Actualizador incremental público habilitado.'
            if enabled else 'El actualizador público está disponible en Android y Windows.'
        ),
    }
