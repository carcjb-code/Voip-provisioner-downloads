import paramiko
import logging

logger = logging.getLogger(__name__)

def ejecutar_comando_ssh(perfil, comando, timeout=15):
    host = perfil.get('host', '127.0.0.1')
    port = int(perfil.get('puerto', 22))
    usuario = perfil.get('usuario', 'root')
    password = perfil.get('password', '')

    try:
        client = paramiko.SSHClient()
        client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        client.connect(hostname=host, port=port, username=usuario, password=password, timeout=10, allow_agent=False, look_for_keys=False)

        # Ejecutar sin PTY para evitar prompts interactivos de Issabel
        stdin, stdout, stderr = client.exec_command(comando, timeout=timeout)
        out = stdout.read().decode('utf-8', errors='replace').strip()
        err = stderr.read().decode('utf-8', errors='replace').strip()
        exit_status = stdout.channel.recv_exit_status()

        client.close()
        return out, err, exit_status
    except paramiko.AuthenticationException:
        return "", "Authentication failed", 255
    except paramiko.SSHException as e:
        return "", str(e), 1
    except Exception as e:
        return "", str(e), 2

def conectar_sftp(host, puerto, usuario, password, ruta_remota='/tftpboot', compat=False):
    try:
        transport = paramiko.Transport((host, int(puerto)))
        transport.connect(username=usuario, password=password)
        sftp = paramiko.SFTPClient.from_transport(transport)
        return sftp, transport
    except Exception as e:
        logger.error(f"Error SFTP: {e}")
        return None, None