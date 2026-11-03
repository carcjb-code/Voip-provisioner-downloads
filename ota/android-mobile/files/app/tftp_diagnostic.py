# -*- coding: utf-8 -*-

import os
import re
import time
import hashlib
import shutil
import subprocess
import threading
from datetime import datetime

TFTP_ROOT = "/tftpboot"

DEFAULT_PHONE = "12.1.2.100"
DEFAULT_SERVER = "10.12.9.5"

BACKUP_ROOT = "/opt/backup/voip-provisioner/tftp"

_capture_lock = threading.Lock()
_capture_jobs = {}


def _run(cmd, timeout=10):
    try:
        p = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )

        out, err = p.communicate(timeout=timeout)

        return {
            "code": p.returncode,
            "stdout": out.decode("utf-8", "replace"),
            "stderr": err.decode("utf-8", "replace")
        }

    except Exception as e:
        return {
            "code": 999,
            "stdout": "",
            "stderr": str(e)
        }


def _md5(path):
    try:
        h = hashlib.md5()

        with open(path, "rb") as f:
            for chunk in iter(
                lambda: f.read(65536),
                b""
            ):
                h.update(chunk)

        return h.hexdigest()

    except Exception:
        return ""


def _safe_filename(filename):
    filename = os.path.basename(
        str(filename or "")
    )

    if not re.match(
        r'^[A-Za-z0-9_.-]+$',
        filename
    ):
        return ""

    return filename


def _file_info(filename, root=TFTP_ROOT):

    filename = _safe_filename(filename)

    if not filename:
        return {
            "exists": False,
            "filename": "",
            "size": 0,
            "md5": "",
            "path": ""
        }

    path = os.path.join(root, filename)

    if not os.path.isfile(path):

        return {
            "exists": False,
            "filename": filename,
            "size": 0,
            "md5": "",
            "path": path
        }

    try:
        size = os.path.getsize(path)
    except Exception:
        size = 0

    return {
        "exists": True,
        "filename": filename,
        "size": size,
        "md5": _md5(path),
        "path": path
    }


def get_file(filename):

    info = _file_info(filename)

    if not info["exists"]:

        return {
            "status": "error",
            "message": "Archivo no encontrado",
            "file": info
        }

    try:

        with open(
            info["path"],
            "r",
            encoding="utf-8",
            errors="replace"
        ) as f:

            content = f.read()

        return {
            "status": "ok",
            "file": info,
            "content": content
        }

    except Exception as e:

        return {
            "status": "error",
            "message": str(e)
        }


def status():

    result = {
        "timestamp": datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        ),

        "tftp": {
            "service": "xinetd",
            "status": "unknown",
            "port": 69,
            "directory": TFTP_ROOT
        },

        "systemd": {
            "socket": "unknown",
            "service": "unknown"
        },

        "socket": "",
        "firewall": "",
        "sample_file": _file_info(
            "cfg000b82fcdcec.xml"
        )
    }

    x = _run(
        ["systemctl", "is-active", "xinetd"],
        5
    )

    result["tftp"]["status"] = (
        "active"
        if x["stdout"].strip() == "active"
        else x["stdout"].strip() or "inactive"
    )

    s = _run(
        ["systemctl", "is-active", "tftp.socket"],
        5
    )

    result["systemd"]["socket"] = (
        s["stdout"].strip() or "inactive"
    )

    s = _run(
        ["systemctl", "is-active", "tftp.service"],
        5
    )

    result["systemd"]["service"] = (
        s["stdout"].strip() or "inactive"
    )

    ss = _run(
        ["ss", "-lunp"],
        5
    )

    lines = []

    for line in ss["stdout"].splitlines():

        if ":69 " in line or ":69\t" in line:
            lines.append(line.strip())

    result["socket"] = "\n".join(lines)

    fw = _run(
        ["firewall-cmd", "--list-services"],
        5
    )

    if fw["code"] == 0:
        result["firewall"] = fw["stdout"].strip()

    return {
        "status": "ok",
        "data": result
    }


def local_test(filename):

    filename = _safe_filename(
        filename or
        "cfg000b82fcdcec.xml"
    )

    if not filename:

        return {
            "status": "error",
            "success": False,
            "message": "Nombre de archivo inválido"
        }

    target = (
        "/tmp/tftp_admin_test_%s"
        % int(time.time())
    )

    started = time.time()

    try:

        cmd = [
            "bash",
            "-c",
            "(sleep 1; "
            "echo 'get %s %s'; "
            "sleep 1; "
            "echo 'quit') | tftp 127.0.0.1"
            % (filename, target)
        ]

        p = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )

        out, err = p.communicate(
            timeout=15
        )

        elapsed = round(
            time.time() - started,
            3
        )

        ok = os.path.isfile(target)

        result = {
            "status": "ok" if ok else "error",
            "success": ok,
            "filename": filename,
            "elapsed": elapsed,
            "output":
                out.decode(
                    "utf-8",
                    "replace"
                )
                +
                err.decode(
                    "utf-8",
                    "replace"
                )
        }

        if ok:

            result["size"] = os.path.getsize(
                target
            )

            result["md5"] = _md5(target)

        try:
            os.remove(target)
        except Exception:
            pass

        return result

    except Exception as e:

        return {
            "status": "error",
            "success": False,
            "message": str(e)
        }


def backup_tftp():

    stamp = datetime.now().strftime(
        "%Y%m%d_%H%M%S"
    )

    dest = os.path.join(
        BACKUP_ROOT,
        stamp
    )

    try:

        os.makedirs(
            dest,
            exist_ok=True
        )

        files = [
            (
                "/etc/xinetd.d/tftp",
                "tftp.xinetd"
            ),
            (
                "/usr/lib/systemd/system/tftp.socket",
                "tftp.socket"
            ),
            (
                "/usr/lib/systemd/system/tftp.service",
                "tftp.service"
            )
        ]

        copied = []

        for src, name in files:

            if os.path.isfile(src):

                shutil.copy2(
                    src,
                    os.path.join(
                        dest,
                        name
                    )
                )

                copied.append(name)

        return {
            "status": "ok",
            "path": dest,
            "files": copied
        }

    except Exception as e:

        return {
            "status": "error",
            "message": str(e)
        }


def list_backups():

    result = []

    if not os.path.isdir(
        BACKUP_ROOT
    ):
        return result

    for name in sorted(
        os.listdir(BACKUP_ROOT),
        reverse=True
    ):

        path = os.path.join(
            BACKUP_ROOT,
            name
        )

        if os.path.isdir(path):

            result.append({
                "name": name,
                "path": path
            })

    return result[:30]


# ============================================================
# SERVIDOR REMOTO
# ============================================================

def _remote_ssh(server):

    """
    El servidor se recibe desde el perfil.
    Se utiliza la infraestructura SSH existente
    del proyecto.
    """

    return server


def server_status(server):

    if not server:

        return {
            "status": "error",
            "message": "No se recibió servidor"
        }

    return {
        "status": "ok",
        "server": server,
        "note":
            "El diagnóstico debe ejecutarse "
            "sobre el servidor seleccionado."
    }


def server_local_test(server, filename):

    """
    Esta función prepara la información para que
    routes.py utilice el SSH existente del proyecto.
    """

    if not server:

        return {
            "status": "error",
            "message": "Servidor no seleccionado"
        }

    filename = _safe_filename(filename)

    if not filename:

        return {
            "status": "error",
            "message": "Archivo inválido"
        }

    return {
        "status": "ready",
        "server": server,
        "filename": filename
    }


def get_server_file(server, filename):

    if not server:

        return {
            "status": "error",
            "message": "Servidor no seleccionado"
        }

    filename = _safe_filename(filename)

    if not filename:

        return {
            "status": "error",
            "message": "Archivo inválido"
        }

    return {
        "status": "ready",
        "server": server,
        "filename": filename
    }


def resend_file(server, filename, phone_ip):

    if not server:

        return {
            "status": "error",
            "message": "Servidor no seleccionado"
        }

    filename = _safe_filename(filename)

    if not filename:

        return {
            "status": "error",
            "message": "Archivo inválido"
        }

    return {
        "status": "ready",
        "server": server,
        "filename": filename,
        "phone_ip": phone_ip
    }


# ============================================================
# CAPTURA
# ============================================================

def start_capture(
    phone_ip,
    seconds=60
):

    phone_ip = str(
        phone_ip or
        DEFAULT_PHONE
    ).strip()

    if not re.match(
        r"^[0-9a-fA-F:.]+$",
        phone_ip
    ):

        return {
            "status": "error",
            "message": "IP inválida"
        }

    try:
        seconds = int(seconds)
    except Exception:
        seconds = 60

    seconds = max(
        10,
        min(seconds, 180)
    )

    with _capture_lock:

        running = [
            x
            for x in _capture_jobs.values()
            if x.get("running")
        ]

        if running:

            return {
                "status": "error",
                "message":
                    "Ya hay una captura en curso",
                "job_id":
                    running[0]["id"]
            }

        job_id = datetime.now().strftime(
            "%Y%m%d_%H%M%S"
        )

        job = {
            "id": job_id,
            "running": True,
            "started":
                datetime.now().strftime(
                    "%Y-%m-%d %H:%M:%S"
                ),
            "phone_ip": phone_ip,
            "seconds": seconds,
            "elapsed": 0,
            "rrq": 0,
            "data": 0,
            "ack": 0,
            "error": 0,
            "lines": [],
            "result":
                "CAPTURANDO..."
        }

        _capture_jobs[job_id] = job

    t = threading.Thread(
        target=_capture_worker,
        args=(job_id,),
        daemon=True
    )

    t.start()

    return {
        "status": "ok",
        "job_id": job_id
    }


def _capture_worker(job_id):

    job = _capture_jobs[job_id]

    phone_ip = job["phone_ip"]
    seconds = job["seconds"]

    try:

        cmd = [
            "tcpdump",
            "-ni",
            "any",
            "-vv",
            "-s0",
            "host",
            phone_ip,
            "and",
            "udp"
        ]

        p = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT
        )

        start = time.time()

        while (
            time.time() - start <
            seconds
        ):

            if p.poll() is not None:
                break

            time.sleep(1)

            job["elapsed"] = int(
                time.time() - start
            )

            try:

                data = p.stdout.read1(
                    65536
                )

                out = data.decode(
                    "utf-8",
                    "replace"
                )

            except Exception:
                out = ""

            if out:

                for line in out.splitlines():

                    line = line.strip()

                    if not line:
                        continue

                    job["lines"].append(
                        line
                    )

                    if " RRQ " in line:
                        job["rrq"] += 1

                    if " DATA " in line:
                        job["data"] += 1

                    if " ACK " in line:
                        job["ack"] += 1

                    if " ERROR " in line:
                        job["error"] += 1

                job["lines"] = \
                    job["lines"][-100:]

        try:
            p.terminate()
            p.wait(timeout=3)
        except Exception:
            try:
                p.kill()
            except Exception:
                pass

        if (
            job["rrq"] > 0 and
            job["data"] > 0 and
            job["ack"] > 0
        ):

            job["result"] = \
                "PROVISIONADO - RRQ → DATA → ACK"

        elif job["rrq"] > 0:

            job["result"] = \
                "ERROR: TELEFONO SOLICITA PERO TFTP NO COMPLETA"

        else:

            job["result"] = \
                "NO SE DETECTARON PEDIDOS TFTP"

    except Exception as e:

        job["result"] = \
            "ERROR: %s" % e

    finally:

        job["running"] = False

        job["elapsed"] = seconds


def capture_status(job_id):

    job = _capture_jobs.get(
        job_id
    )

    if not job:

        return {
            "status": "error",
            "message":
                "Captura no encontrada"
        }

    return {
        "status": "ok",
        "job": dict(job)
    }
