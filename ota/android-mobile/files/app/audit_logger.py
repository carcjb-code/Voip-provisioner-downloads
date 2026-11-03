import os
import json
import logging
from datetime import datetime
from threading import Lock

class AuditLogger:
    def __init__(self, max_lines=10000):
        if os.name == 'nt':
            base = os.path.join(os.environ.get('APPDATA', os.path.expanduser('~')), 'VoIP_Provisioner')
        else:
            base = os.path.join(os.path.expanduser('~'), '.config', 'VoIP_Provisioner')
        
        self.log_dir = os.path.join(base, 'logs')
        self.audit_file = os.path.join(self.log_dir, 'audit.log')
        self.error_file = os.path.join(self.log_dir, 'error.log')
        self.max_lines = max_lines
        self.lock = Lock()
        
        os.makedirs(self.log_dir, exist_ok=True)
        
        self.console_logger = logging.getLogger('voip_audit')
        if not self.console_logger.handlers:
            h = logging.StreamHandler()
            h.setFormatter(logging.Formatter('%(asctime)s [%(levelname)s] %(message)s'))
            self.console_logger.addHandler(h)
            self.console_logger.setLevel(logging.INFO)

    def log_audit(self, accion, perfil='', detalles='', estado='success', user='system', tipo='info'):
        entry = {
            'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'accion': accion,
            'perfil': perfil,
            'detalles': detalles,
            'estado': estado,
            'user': user,
            'tipo': tipo
        }
        self._write_to_file(self.audit_file, entry)
        self.console_logger.info(f"[AUDIT] [{tipo}] {accion} | {perfil} | {detalles} ({estado})")

    def log_error(self, error_type, context='', message='', traceback=''):
        entry = {
            'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'error_type': error_type,
            'context': context,
            'message': message,
            'traceback': traceback[:500] if traceback else '',
            'tipo': 'error'
        }
        self._write_to_file(self.error_file, entry)
        self.console_logger.error(f"[ERROR] [{error_type}] {context}: {message}")

    def _write_to_file(self, filepath, entry):
        line = json.dumps(entry, ensure_ascii=False) + '\n'
        with self.lock:
            if os.path.exists(filepath):
                with open(filepath, 'r', encoding='utf-8') as f:
                    if sum(1 for _ in f) >= self.max_lines:
                        self._rotate(filepath)
            with open(filepath, 'a', encoding='utf-8') as f:
                f.write(line)

    def _rotate(self, filepath):
        old = f"{filepath}.old"
        if os.path.exists(old):
            os.remove(old)
        if os.path.exists(filepath):
            os.rename(filepath, old)

    def get_audit_logs(self, limit=50, offset=0, action=None, profile=None, status=None, search=None, tipo=None, fecha_desde=None, fecha_hasta=None):
        """
        Obtiene logs de auditoría con filtros avanzados
        
        Args:
            limit: Cantidad máxima de registros a retornar
            offset: Número de registros a saltar (para paginación)
            action: Filtrar por acción específica
            profile: Filtrar por perfil
            status: Filtrar por estado (success/error)
            search: Búsqueda de texto libre
            tipo: Filtrar por tipo (create/delete/edit/sync/error)
            fecha_desde: Filtrar logs desde esta fecha (YYYY-MM-DD)
            fecha_hasta: Filtrar logs hasta esta fecha (YYYY-MM-DD)
        
        Returns:
            dict con logs, total, limit, offset
        """
        logs, total = self._read_logs_with_pagination(
            self.audit_file, limit, offset, {
                'action': action, 'profile': profile, 'status': status, 
                'search': search, 'tipo': tipo,
                'fecha_desde': fecha_desde, 'fecha_hasta': fecha_hasta
            }
        )
        return {
            'logs': logs,
            'total': total,
            'limit': limit,
            'offset': offset,
            'pagina_actual': (offset // limit) + 1 if limit > 0 else 1,
            'total_paginas': (total // limit) + (1 if total % limit > 0 else 0) if limit > 0 else 1
        }

    def get_error_logs(self, limit=50, offset=0, error_type=None, context=None, search=None, fecha_desde=None, fecha_hasta=None):
        logs, total = self._read_logs_with_pagination(
            self.error_file, limit, offset, {
                'error_type': error_type, 'context': context, 'search': search,
                'fecha_desde': fecha_desde, 'fecha_hasta': fecha_hasta
            }
        )
        return {
            'logs': logs,
            'total': total,
            'limit': limit,
            'offset': offset,
            'pagina_actual': (offset // limit) + 1 if limit > 0 else 1,
            'total_paginas': (total // limit) + (1 if total % limit > 0 else 0) if limit > 0 else 1
        }

    def _read_logs_with_pagination(self, filepath, limit, offset, filters):
        """Lee logs aplicando filtros y retorna página específica"""
        all_logs = []
        files = [filepath]
        old = f"{filepath}.old"
        if os.path.exists(old):
            files.append(old)
        
        for f in files:
            try:
                with open(f, 'r', encoding='utf-8') as fh:
                    for line in fh:
                        if not line.strip():
                            continue
                        try:
                            entry = json.loads(line.strip())
                            
                            # Aplicar filtros
                            if filters.get('action') and entry.get('accion') != filters['action']:
                                continue
                            if filters.get('profile') and entry.get('perfil') != filters['profile']:
                                continue
                            if filters.get('status') and entry.get('estado') != filters['status']:
                                continue
                            if filters.get('error_type') and entry.get('error_type') != filters['error_type']:
                                continue
                            if filters.get('context') and entry.get('context') != filters['context']:
                                continue
                            if filters.get('tipo') and entry.get('tipo') != filters['tipo']:
                                continue
                            
                            # Filtro de búsqueda de texto
                            if filters.get('search'):
                                search_lower = filters['search'].lower()
                                if not any(search_lower in str(v).lower() for v in entry.values()):
                                    continue
                            
                            # Filtros de fecha
                            if filters.get('fecha_desde') or filters.get('fecha_hasta'):
                                entry_date = entry.get('timestamp', '')[:10]  # YYYY-MM-DD
                                if filters.get('fecha_desde') and entry_date < filters['fecha_desde']:
                                    continue
                                if filters.get('fecha_hasta') and entry_date > filters['fecha_hasta']:
                                    continue
                            
                            all_logs.append(entry)
                        except json.JSONDecodeError:
                            continue
            except Exception:
                continue
        
        # Ordenar por fecha descendente
        all_logs.sort(key=lambda x: x.get('timestamp', ''), reverse=True)
        
        total = len(all_logs)
        
        # Aplicar paginación
        paginated_logs = all_logs[offset:offset + limit]
        
        return paginated_logs, total

    def clear_logs(self, log_type='all'):
        with self.lock:
            if log_type in ['audit', 'all']:
                if os.path.exists(self.audit_file):
                    os.remove(self.audit_file)
                old = f"{self.audit_file}.old"
                if os.path.exists(old):
                    os.remove(old)
            if log_type in ['error', 'all']:
                if os.path.exists(self.error_file):
                    os.remove(self.error_file)
                old = f"{self.error_file}.old"
                if os.path.exists(old):
                    os.remove(old)
        self.console_logger.info(f"[LOGS] Logs {log_type} limpiados")

    def get_stats(self):
        stats = {'audit': 0, 'error': 0, 'audit_errors': 0}
        for f in [self.audit_file, self.error_file]:
            if os.path.exists(f):
                try:
                    with open(f, 'r', encoding='utf-8') as fh:
                        count = sum(1 for _ in fh)
                        if 'audit' in f:
                            stats['audit'] = count
                        else:
                            stats['error'] = count
                except Exception:
                    pass
        try:
            if os.path.exists(self.audit_file):
                with open(self.audit_file, 'r', encoding='utf-8') as fh:
                    for line in fh:
                        if '"estado": "error"' in line:
                            stats['audit_errors'] += 1
        except Exception:
            pass
        return stats

audit_logger = AuditLogger()