import os
import subprocess
import logging

logger = logging.getLogger(__name__)

TFTP_DIR = '/tftpboot'

def escribir_archivo_provision_local(mac, contenido_cfg):
    try:
        if not os.path.exists(TFTP_DIR):
            os.makedirs(TFTP_DIR, exist_ok=True)
            
        mac_clean = mac.replace(":", "").replace("-", "").lower()
        filepath = os.path.join(TFTP_DIR, f"{mac_clean}.cfg")
        
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(contenido_cfg)
            
        return True, filepath
    except Exception as e:
        logger.error(f"Error escribiendo archivo local en {TFTP_DIR}: {e}")
        return False, str(e)

def eliminar_archivo_provision_local(mac):
    try:
        mac_clean = mac.replace(":", "").replace("-", "").lower()
        filepath = os.path.join(TFTP_DIR, f"{mac_clean}.cfg")
        if os.path.exists(filepath):
            os.remove(filepath)
        return True
    except Exception as e:
        logger.error(f"Error eliminando archivo {mac}: {e}")
        return False

def recargar_asterisk_local():
    try:
        subprocess.run(["asterisk", "-rx", "core reload"], check=True)
        return True
    except Exception as e:
        logger.error(f"Error recargando Asterisk: {e}")
        return False
