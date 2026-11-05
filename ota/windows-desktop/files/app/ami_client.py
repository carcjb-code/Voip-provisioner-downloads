import socket
import re
import logging
import threading
import time
from typing import Dict, List, Optional, Tuple

logger = logging.getLogger(__name__)

class AMIClient:
    """Cliente AMI ligero para monitoreo de Asterisk."""
    def __init__(self, host: str, username: str, secret: str, port: int = 5038, timeout: float = 5.0):
        self.host = host
        self.port = port
        self.username = username
        self.secret = secret
        self.timeout = timeout
        self._lock = threading.Lock()

    def _connect(self) -> Optional[socket.socket]:
        try:
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.settimeout(self.timeout)
            sock.connect((self.host, self.port))
            banner = sock.recv(1024).decode('utf-8', errors='ignore')
            if 'Asterisk Call Manager' not in banner:
                logger.warning(f"Banner AMI inesperado de {self.host}: {banner[:50]}")
                sock.close()
                return None
            return sock
        except Exception as e:
            logger.error(f"Error conectando AMI a {self.host}:{self.port}: {e}")
            return None

    def _login(self, sock: socket.socket) -> bool:
        try:
            action = f"Action: Login\r\nUsername: {self.username}\r\nSecret: {self.secret}\r\nEvents: off\r\n\r\n"
            sock.sendall(action.encode('utf-8'))
            response = self._read_response(sock)
            return response.get('Response', '').lower() == 'success'
        except Exception as e:
            logger.error(f"Error en login AMI: {e}")
            return False

    def _read_response(self, sock: socket.socket) -> Dict[str, str]:
        buffer = b""
        try:
            while True:
                chunk = sock.recv(4096)
                if not chunk: break
                buffer += chunk
                if b"\r\n\r\n" in buffer: break
        except socket.timeout:
            pass
        
        response = {}
        text = buffer.decode('utf-8', errors='ignore')
        for line in text.split('\r\n'):
            if ':' in line:
                key, _, value = line.partition(':')
                response[key.strip()] = value.strip()
        return response

    def _send_action(self, sock: socket.socket, action: str, **params) -> List[Dict[str, str]]:
        import random
        action_id = f"{int(time.time())}-{random.randint(1000, 9999)}"
        cmd = f"Action: {action}\r\n"
        for k, v in params.items():
            cmd += f"{k}: {v}\r\n"
        cmd += f"ActionID: {action_id}\r\n\r\n"
        
        sock.sendall(cmd.encode('utf-8'))
        responses = []
        buffer = b""
        
        try:
            while True:
                chunk = sock.recv(8192)
                if not chunk: break
                buffer += chunk
                while b"\r\n\r\n" in buffer:
                    msg, buffer = buffer.split(b"\r\n\r\n", 1)
                    text = msg.decode('utf-8', errors='ignore')
                    event = {}
                    for line in text.split('\r\n'):
                        if ':' in line:
                            k, _, v = line.partition(':')
                            event[k.strip()] = v.strip()
                    if event:
                        responses.append(event)
                        if event.get('Event') == 'CommandComplete':
                            return responses
                        if event.get('Response') in ('Error',) and event.get('ActionID') == action_id:
                            return responses
        except socket.timeout:
            pass
        return responses

    def _logout(self, sock: socket.socket):
        try:
            sock.sendall(b"Action: Logoff\r\n\r\n")
            sock.close()
        except Exception:
            pass

    def execute(self, command: str) -> str:
        sock = self._connect()
        if not sock: return ""
        try:
            if not self._login(sock): return ""
            responses = self._send_action(sock, "Command", Command=command)
            output_lines = []
            capture = False
            for r in responses:
                if r.get('Response') == 'Follows':
                    capture = True
                    if 'Output' in r: output_lines.append(r['Output'])
                elif capture:
                    if 'Output' in r: output_lines.append(r['Output'])
                    if r.get('Event') == 'CommandComplete': break
            return "\n".join(output_lines)
        finally:
            self._logout(sock)

    def test_connection(self) -> Tuple[bool, str]:
        sock = self._connect()
        if not sock: return False, f"No se pudo conectar a {self.host}:{self.port}"
        try:
            if self._login(sock): return True, "Conexión AMI exitosa"
            return False, "Autenticación AMI fallida"
        finally:
            self._logout(sock)