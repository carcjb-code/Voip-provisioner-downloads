# -*- coding: utf-8 -*-

import time
import logging
import threading
import re
import queue

from app.asterisk import (
    ejecutar_comando_ssh,
    ejecutar_comando_local,
    es_host_local
)

logger = logging.getLogger(__name__)

profile_monitors = {}
monitor_lock = threading.Lock()

FAST_INTERVAL = 1.0
SLOW_INTERVAL = 5.0


def _ejecutar_comando(perfil, comando):
    """
    Ejecuta comandos contra Asterisk.

    Si el perfil corresponde al propio servidor:
        -> ejecución local

    Si corresponde a un servidor remoto:
        -> SSH
    """
    try:
        host = str(perfil.get("host", "") or "").strip()

        # Servidor local: NO usar SSH
        if es_host_local(host):
            try:
                resultado = ejecutar_comando_local(comando)

                if isinstance(resultado, tuple):
                    if len(resultado) >= 3:
                        return (
                            resultado[0] or "",
                            resultado[1] or "",
                            resultado[2]
                        )
                    if len(resultado) == 2:
                        return (
                            resultado[0] or "",
                            resultado[1] or "",
                            0
                        )
                    if len(resultado) == 1:
                        return (
                            resultado[0] or "",
                            "",
                            0
                        )

                if resultado is None:
                    return "", "", 1

                return str(resultado), "", 0

            except Exception as e:
                logger.error(
                    "[MONITOR] Error comando local: %s",
                    e
                )
                return "", str(e), 1

        # Servidor remoto: SSH
        try:
            resultado = ejecutar_comando_ssh(
                perfil,
                comando
            )

            if isinstance(resultado, tuple):
                if len(resultado) >= 3:
                    return (
                        resultado[0] or "",
                        resultado[1] or "",
                        resultado[2]
                    )
                if len(resultado) == 2:
                    return (
                        resultado[0] or "",
                        resultado[1] or "",
                        0
                    )
                if len(resultado) == 1:
                    return (
                        resultado[0] or "",
                        "",
                        0
                    )

            if resultado is None:
                return "", "", 1

            return str(resultado), "", 0

        except Exception as e:
            logger.error(
                "[MONITOR] Error SSH: %s",
                e
            )
            return "", str(e), 1

    except Exception as e:
        logger.error(
            "[MONITOR] Error ejecutando comando: %s",
            e
        )
        return "", str(e), 1


def _ips_desde_db(perfil):
    cmd = (
        "mysql -N -B -u root "
        "-e \"SELECT name,data "
        "FROM asterisk.sip "
        "WHERE keyword='ip' "
        "AND data NOT IN "
        "('dynamic','','0.0.0.0')\" "
        "asterisk 2>/dev/null"
    )

    out, err, code = _ejecutar_comando(perfil, cmd)

    if code != 0 or not out:
        return {}

    resultado = {}

    for linea in out.splitlines():
        partes = linea.split("\t")

        if len(partes) < 2:
            continue

        nombre = partes[0].strip()
        ip = partes[1].strip()

        if nombre and ip:
            resultado[nombre] = ip

    return resultado


def _parse_peers(perfil, ips_dict):
    """
    Obtiene internos tanto CHAN_SIP como PJSIP.

    CHAN_SIP:
        sip show peers

    PJSIP:
        pjsip show endpoints
        pjsip show contacts

    Los dos resultados se unifican por nombre de interno.
    Si un interno existe en ambas tecnologías se evita
    mostrarlo dos veces y se conserva preferentemente
    la información PJSIP.
    """

    peers = {}

    # =========================================================
    # 1. CHAN_SIP
    # =========================================================

    out, err, code = _ejecutar_comando(
        perfil,
        "asterisk -rx 'sip show peers' 2>/dev/null"
    )

    if code == 0 and out:

        for raw in out.splitlines():

            linea = raw.strip()

            if not linea:
                continue

            if linea.startswith("Name"):
                continue

            if "---" in linea:
                continue

            if "sip peers" in linea.lower():
                continue

            partes = linea.split()

            if len(partes) < 2:
                continue

            nombre = (
                partes[0]
                .split("/")[0]
                .strip()
            )

            # Solamente internos numéricos.
            if not re.match(
                r'^\d{2,8}$',
                nombre
            ):
                continue

            ip = ips_dict.get(
                nombre,
                ""
            )

            if not ip:

                for parte in partes[1:]:

                    if re.match(
                        r'^\d+\.\d+\.\d+\.\d+(?::\d+)?$',
                        parte
                    ):
                        ip = parte.split(":")[0]
                        break

            linea_low = linea.lower()

            status_raw = "Unknown"

            if "unmonitored" in linea_low:
                status_raw = "Unmonitored"

            elif "unreachable" in linea_low:
                status_raw = "Unreachable"

            elif "lagged" in linea_low:
                status_raw = "LAGGED"

            elif re.search(
                r'\bOK\b',
                linea,
                re.I
            ):
                status_raw = "OK"

            online = (
                bool(ip)
                or status_raw == "OK"
            )

            latencia = "-"

            match = re.search(
                r'(\d+(?:\.\d+)?)\s*ms',
                linea
            )

            if match:
                latencia = match.group(1)

            peers[nombre] = {
                "nombre": nombre,
                "ip": ip or "-",
                "status": (
                    "Online"
                    if online
                    else "Offline"
                ),
                "status_raw": status_raw,
                "latencia": latencia,
                "en_llamada": False,
                "tecnologia": "SIP"
            }

    # =========================================================
    # 2. PJSIP - DESCUBRIR ENDPOINTS
    # =========================================================
    #
    # NO usamos pjsip_additional.conf para descubrir internos.
    #
    # Asterisk nos dice directamente qué endpoints PJSIP tiene
    # cargados mediante:
    #
    #   pjsip show endpoints
    #
    # Esto permite detectar un interno aunque esté OFFLINE.
    #
    # Ejemplo habitual:
    #
    # Endpoint: 1993/1993
    #
    # =========================================================

    pjsip_endpoints = set()

    out_ep, err_ep, code_ep = _ejecutar_comando(
        perfil,
        "asterisk -rx 'pjsip show endpoints' 2>/dev/null"
    )

    if code_ep == 0 and out_ep:

        for raw in out_ep.splitlines():

            linea = raw.strip()

            if not linea:
                continue

            # Formato habitual:
            #
            # Endpoint: 1993/1993 ...
            #
            match = re.search(
                r'^\s*Endpoint:\s*(\d{2,8})(?:/|\s|$)',
                linea,
                re.I
            )

            if match:

                nombre = match.group(1).strip()

                if nombre:
                    pjsip_endpoints.add(nombre)

            # Algunas versiones/salidas pueden presentar:
            #
            # 1993/1993
            #
            # como columna de endpoint.
            #
            if not match:

                match = re.match(
                    r'^(\d{2,8})/\d{2,8}\b',
                    linea
                )

                if match:
                    pjsip_endpoints.add(
                        match.group(1)
                    )

    # =========================================================
    # 3. PJSIP - CONTACTS
    # =========================================================
    #
    # Esta consulta sirve para conocer:
    #
    #   - qué interno está registrado
    #   - IP
    #   - latencia
    #
    # Un endpoint puede existir sin tener Contact.
    # En ese caso queda Offline.
    #
    # =========================================================

    pjsip_contacts = {}

    out_contacts, err_contacts, code_contacts = _ejecutar_comando(
        perfil,
        "asterisk -rx 'pjsip show contacts' 2>/dev/null"
    )

    if code_contacts == 0 and out_contacts:

        for raw in out_contacts.splitlines():

            linea = raw.strip()

            if not linea:
                continue

            if (
                linea.startswith("Contact:")
                or re.search(
                    r'^\s*Contact:',
                    linea,
                    re.I
                )
            ):

                # -------------------------------------------------
                # Ejemplo:
                #
                # Contact: 1993/sip:1993@10.1.1.20:5060
                #          1993/sip:1993@10.1.1.20:5060;...
                # -------------------------------------------------

                match_nombre = re.search(
                    r'Contact:\s*(\d{2,8})/',
                    linea,
                    re.I
                )

                if not match_nombre:
                    continue

                nombre = match_nombre.group(1)

                ip = ""

                # IPv4
                match_ip = re.search(
                    r'@(\d+\.\d+\.\d+\.\d+)(?::\d+)?',
                    linea
                )

                if match_ip:
                    ip = match_ip.group(1)

                # IPv6 entre []
                if not ip:

                    match_ip6 = re.search(
                        r'@\[([0-9a-fA-F:]+)\](?::\d+)?',
                        linea
                    )

                    if match_ip6:
                        ip = match_ip6.group(1)

                # Latencia, cuando Asterisk la muestra.
                latencia = "-"

                match_ms = re.search(
                    r'(\d+(?:\.\d+)?)\s*ms',
                    linea,
                    re.I
                )

                if match_ms:
                    latencia = match_ms.group(1)

                pjsip_contacts[(nombre, ip or "-")] = {
                    "ip": ip or "-",
                    "latencia": latencia
                }

                continue

            # -------------------------------------------------
            # Algunas versiones muestran Contact sin el formato
            # anterior o pueden mostrar una URI en una línea.
            #
            # Si encontramos un sip:usuario@IP y el usuario es
            # numérico, también lo reconocemos.
            # -------------------------------------------------

            match_uri = re.search(
                r'sip:(\d{2,8})@(\d+\.\d+\.\d+\.\d+)(?::\d+)?',
                linea,
                re.I
            )

            if match_uri:

                nombre = match_uri.group(1)
                ip = match_uri.group(2)

                latencia = "-"

                match_ms = re.search(
                    r'(\d+(?:\.\d+)?)\s*ms',
                    linea,
                    re.I
                )

                if match_ms:
                    latencia = match_ms.group(1)

                pjsip_contacts[(nombre, ip)] = {
                    "ip": ip,
                    "latencia": latencia
                }

    # =========================================================
    # 4. SI HAY CONTACTOS QUE NO APARECIERON EN ENDPOINTS,
    #    TAMBIÉN LOS AGREGAMOS.
    # =========================================================

    for nombre, _ip in pjsip_contacts.keys():
        pjsip_endpoints.add(nombre)

    # =========================================================
    # 5. CONSTRUIR INTERNOS PJSIP
    # =========================================================

    for nombre in sorted(pjsip_endpoints, key=lambda x: str(x)):
        contactos = [
            (ip_contacto, datos)
            for (nombre_contacto, ip_contacto), datos
            in pjsip_contacts.items()
            if nombre_contacto == nombre
        ]

        if contactos:
            # Si el interno es PJSIP, prevalece sobre CHAN_SIP,
            # pero se conservan todos sus teléfonos/IP distintos.
            peers.pop(nombre, None)

            for ip_contacto, contacto in contactos:
                ip = str(contacto.get("ip") or "").strip()
                latencia = str(contacto.get("latencia") or "-").strip()
                registrado = ip not in ("", "-", "0.0.0.0")

                clave_peer = nombre + "|" + ip_contacto
                peers[clave_peer] = {
                    "nombre": nombre,
                    "ip": ip or "-",
                    "status": "Online" if registrado else "Offline",
                    "status_raw": "OK" if registrado else "Unavailable",
                    "latencia": latencia,
                    "en_llamada": False,
                    "tecnologia": "PJSIP"
                }
        else:
            # El endpoint existe, pero no tiene ningún contacto registrado.
            peers[nombre] = {
                "nombre": nombre,
                "ip": "-",
                "status": "Offline",
                "status_raw": "Unavailable",
                "latencia": "-",
                "en_llamada": False,
                "tecnologia": "PJSIP"
            }

    # =========================================================
    # 6. ORDEN FINAL
    # =========================================================

    resultado = list(
        peers.values()
    )

    resultado.sort(
        key=lambda x: (
            str(x.get("nombre", "")),
            str(x.get("ip", ""))
        )
    )

    return resultado

def colgar_canales(perfil, canales):
    """
    Finaliza una llamada activa utilizando los canales reales
    informados por Asterisk.

    Se ejecuta sobre cada pata de la llamada para evitar que
    quede un canal conectado después de cortar el otro.
    """
    if not isinstance(canales, (list, tuple)):
        return {
            "ok": False,
            "mensaje": "Lista de canales inválida",
            "resultados": []
        }

    resultados = []

    for canal in canales:
        canal = str(canal or "").strip()

        if not canal:
            continue

        # Evitamos aceptar comandos arbitrarios.
        # Un canal Asterisk debe comenzar por una tecnología
        # conocida y no puede contener comillas ni saltos de línea.
        if (
            "\n" in canal
            or "\r" in canal
            or '"' in canal
            or "'" in canal
            or ";" in canal
            or "|" in canal
            or "`" in canal
            or "$(" in canal
        ):
            resultados.append({
                "canal": canal,
                "ok": False,
                "mensaje": "Canal inválido"
            })
            continue

        if not re.match(
            r"^(SIP|PJSIP|IAX2|OOH323|DAHDI|Local|CUSTOM|DONGLE|ZAP)/[^\s]+$",
            canal,
            re.I
        ):
            resultados.append({
                "canal": canal,
                "ok": False,
                "mensaje": "Formato de canal no permitido"
            })
            continue

        comando = (
            'asterisk -rx "channel request hangup {0}"'
        ).format(canal)

        out, err, code = _ejecutar_comando(
            perfil,
            comando
        )

        resultados.append({
            "canal": canal,
            "ok": code == 0,
            "salida": str(out or "").strip(),
            "error": str(err or "").strip(),
            "codigo": code
        })

    exitos = [
        r for r in resultados
        if r.get("ok") is True
    ]

    return {
        "ok": bool(exitos),
        "mensaje": (
            "Solicitud de corte enviada"
            if exitos
            else "No se pudo cortar ningún canal"
        ),
        "resultados": resultados
    }

def _parse_channels(perfil):
    """
    Obtiene los canales activos de Asterisk y los agrupa por llamada lógica.

    Asterisk puede mostrar dos patas de una misma llamada:
        SIP/57320-xxxx  -> Dial -> SIP/57364
        SIP/57364-yyyy  -> AppDial

    El monitor debe mostrar eso como una sola llamada:
        57320 -> 57364
    """
    import re

    out, err, code = _ejecutar_comando(
        perfil,
        "asterisk -rx 'core show channels concise' 2>/dev/null"
    )

    if code != 0 or not out:
        return []

    registros = []

    def endpoint_desde_canal(canal):
        canal = str(canal or "").strip()

        m = re.match(
            r'^(?:SIP|PJSIP|IAX2|OOH323|DAHDI|Local|CUSTOM)/([^-\s;]+)',
            canal,
            re.I
        )
        if m:
            return m.group(1).strip()

        return ""

    def extraer_destinos(texto):
        """
        Extrae posibles destinos desde Dial/AppDial data.
        """
        texto = str(texto or "")
        encontrados = []

        patrones = [
            r'(?:SIP|PJSIP|IAX2|OOH323|DAHDI|Local|CUSTOM)/([^,\s;&]+)',
            r'^(?:SIP|PJSIP|IAX2|OOH323|DAHDI|Local|CUSTOM)/([^,\s;&]+)',
        ]

        for patron in patrones:
            for m in re.finditer(patron, texto, re.I):
                valor = m.group(1).strip()
                if valor and valor not in encontrados:
                    encontrados.append(valor)

        return encontrados

    def origen_desde_callerid(callerid, canal):
        callerid = str(callerid or "").strip()

        # Formatos habituales: 57320, "57320" <57320>, etc.
        m = re.search(r'<([^>]+)>', callerid)
        if m:
            valor = m.group(1).strip()
            if valor:
                return valor

        m = re.search(r'([0-9*#+]+)', callerid)
        if m:
            return m.group(1)

        # No usamos el nombre del endpoint como CallerID.
        #
        # Ejemplo Synway:
        #   canal = PJSIP/synway-e1-0000001c
        #
        # Si Asterisk no recibió CallerID, el endpoint NO es el
        # número del llamante. El monitor lo tratará más abajo
        # como origen externo/troncal.
        return ""

    # ---------------------------------------------------------
    # Parseo de canales físicos de Asterisk
    # ---------------------------------------------------------
    for linea in out.splitlines():
        linea = linea.strip()

        if not linea:
            continue

        partes = linea.split("!")

        if len(partes) < 7:
            continue

        canal = partes[0].strip()
        contexto = partes[1].strip() if len(partes) > 1 else ""
        extension = partes[2].strip() if len(partes) > 2 else ""
        prioridad = partes[3].strip() if len(partes) > 3 else ""
        estado = partes[4].strip() if len(partes) > 4 else ""
        aplicacion = partes[5].strip() if len(partes) > 5 else ""
        datos = partes[6].strip() if len(partes) > 6 else ""
        callerid = partes[7].strip() if len(partes) > 7 else ""
        duracion = partes[9].strip() if len(partes) > 9 else "-"

        endpoint = endpoint_desde_canal(canal)

        origen = origen_desde_callerid(callerid, canal)

        # -----------------------------------------------------
        # Fallback de origen externo.
        #
        # IMPORTANTE:
        # No inventamos un número cuando Asterisk no recibió
        # CallerID. Esto es exactamente lo que ocurre hoy con
        # las llamadas históricas del Synway E1: clid/src/cid_num
        # llegan vacíos.
        #
        # Cuando la E1 esté conectada y Asterisk reciba el CallerID
        # por SIP, origen_desde_callerid() devolverá el número real
        # automáticamente y este fallback NO se utilizará.
        # -----------------------------------------------------
        contexto_low = str(contexto or "").strip().lower()
        canal_low = str(canal or "").strip().lower()

        if not origen:
            if "synway-e1" in canal_low:
                origen = "EXTERNO / SYNWAY E1"
            elif any(
                x in contexto_low
                for x in (
                    "from-pstn",
                    "from-trunk",
                    "from-external",
                    "from-ipoffice"
                )
            ):
                origen = "EXTERNO / TRONCAL"

        destinos = []

        # Primero el campo extension.
        if extension and extension not in ("-", "s", "unknown"):
            destinos.append(extension)

        # Después lo que aparezca en Dial/AppDial.
        for valor in extraer_destinos(datos):
            if valor not in destinos:
                destinos.append(valor)

        # Si no encontramos destino por datos, usar extension.
        destino = destinos[0] if destinos else "-"

        registros.append({
            "canal": canal,
            "endpoint": endpoint,
            "origen": origen or "-",
            "destino": destino or "-",
            "destinos": destinos,
            "estado": estado or "-",
            "duracion": duracion or "-",
            "contexto": contexto,
            "aplicacion": aplicacion,
            "datos": datos,
            "callerid": callerid,
        })

    # ---------------------------------------------------------
    # Agrupar las dos patas de una misma llamada
    # ---------------------------------------------------------
    usados = set()
    llamadas = []

    def es_appdial(r):
        app = str(r.get("aplicacion", "")).lower()
        return app in ("appdial", "dial")

    def destinos_de(r):
        valores = list(r.get("destinos") or [])

        d = str(r.get("destino") or "").strip()
        if d and d != "-" and d not in valores:
            valores.append(d)

        return valores

    def determinar_direccion(*registros_llamada):
        """
        Clasificación de la llamada para la interfaz.

        Valores posibles:

        entrante_externa
        saliente_externa
        entrante_interna
        """

        registros_validos = [
            r for r in registros_llamada
            if isinstance(r, dict)
        ]

        contextos = []
        canales = []
        endpoints = []
        origenes = []
        destinos = []

        for r in registros_validos:

            contexto = str(
                r.get("contexto") or ""
            ).strip().lower()

            canal = str(
                r.get("canal") or ""
            ).strip().lower()

            endpoint = str(
                r.get("endpoint") or ""
            ).strip().lower()

            origen = str(
                r.get("origen") or ""
            ).strip()

            destino = str(
                r.get("destino") or ""
            ).strip()

            if contexto:
                contextos.append(contexto)

            if canal:
                canales.append(canal)

            if endpoint:
                endpoints.append(endpoint)

            if origen and origen != "-":
                origenes.append(origen)

            if destino and destino != "-":
                destinos.append(destino)

        texto = " ".join(
            contextos +
            canales +
            endpoints
        )

        # -----------------------------------------------------
        # Determinar si existe una pata externa
        # -----------------------------------------------------

        externa = False

        # H.323 / Avaya / IP Office
        if any(
            ep.startswith("ooh323")
            or "ipoffice" in ep
            for ep in endpoints
        ):
            externa = True

        # DAHDI / ZAP / CUSTOM / DONGLE
        if any(
            ep.startswith("dahdi")
            or ep.startswith("zap")
            or ep.startswith("custom")
            or ep.startswith("dongle")
            for ep in endpoints
        ):
            externa = True

        # Contextos claramente externos
        if any(
            "from-ipoffice" in c
            or "from-trunk" in c
            or "from-pstn" in c
            or "from-external" in c
            for c in contextos
        ):
            externa = True

        # -----------------------------------------------------
        # Números externos
        #
        # Los internos del sistema tienen hasta 5 dígitos.
        # Un número de más de 5 dígitos se considera externo.
        # -----------------------------------------------------

        todos_los_numeros = (
            origenes +
            destinos
        )

        for numero in todos_los_numeros:

            limpio = re.sub(
                r"[^0-9]",
                "",
                str(numero)
            )

            if len(limpio) > 5:
                externa = True
                break

        # -----------------------------------------------------
        # IAX2 no se considera llamada externa solamente
        # por ser un canal distinto.
        # -----------------------------------------------------

        canales_no_iax = [
            c for c in canales
            if not c.startswith("iax2/")
        ]

        tiene_iax = bool(canales) and not bool(canales_no_iax)

        # Si solamente hay IAX2 y no hay número externo,
        # lo consideramos interno.
        if tiene_iax and not externa:
            return "entrante_interna"

        # -----------------------------------------------------
        # DIRECCION
        # -----------------------------------------------------

        # Desde IP Office / troncal hacia este Asterisk
        if any(
            "from-ipoffice" in c
            or "from-trunk" in c
            or "from-pstn" in c
            or "from-external" in c
            for c in contextos
        ):
            if externa:
                return "entrante_externa"

        # Desde un canal H323 hacia Asterisk
        if any(
            ep.startswith("ooh323")
            or "ipoffice" in ep
            for ep in endpoints
        ):
            if any(
                "from-" in c and
                "dialout" not in c
                for c in contextos
            ):
                return "entrante_externa"

        # Salida mediante macro-dialout-trunk
        if any(
            "macro-dialout-trunk" in c
            or "outbound" in c
            or "from-internal" in c
            for c in contextos
        ):
            if externa:
                return "saliente_externa"

        # Si tenemos un número de más de 5 dígitos
        # y la llamada parte desde un interno,
        # se considera salida externa.
        if externa:
            if any(
                c == "from-internal"
                or c.startswith("from-internal-")
                or "macro-dialout-trunk" in c
                for c in contextos
            ):
                return "saliente_externa"

        # -----------------------------------------------------
        # Si no es externa, es una llamada entre internos.
        # -----------------------------------------------------

        return "entrante_interna"


    for i, a in enumerate(registros):
        if i in usados:
            continue

        mejor_j = None
        mejor_score = 0

        a_ep = str(a.get("endpoint") or "").strip()
        a_origen = str(a.get("origen") or "").strip()
        a_dests = destinos_de(a)

        for j in range(i + 1, len(registros)):
            if j in usados:
                continue

            b = registros[j]

            b_ep = str(b.get("endpoint") or "").strip()
            b_origen = str(b.get("origen") or "").strip()
            b_dests = destinos_de(b)

            score = 0

            # Caso principal:
            # SIP/57320 -> Dial -> SIP/57364
            # SIP/57364 -> AppDial
            if a_ep and a_ep in b_dests:
                score += 100

            if b_ep and b_ep in a_dests:
                score += 100

            # Caso troncal:
            #
            # SIP/57320 -> Dial -> OOH323/ipoffice/6200
            # OOH323/ipoffice -> AppDial -> 6200
            #
            # En la primera pata el destino puede venir como:
            #   ipoffice/6200
            # mientras que el endpoint de la segunda es:
            #   ipoffice
            #
            # Consideramos ambos como la misma pata.
            if a_ep and any(
                str(d).strip().lower() == a_ep.lower()
                or str(d).strip().lower().startswith(
                    a_ep.lower() + "/"
                )
                for d in b_dests
            ):
                score += 120

            if b_ep and any(
                str(d).strip().lower() == b_ep.lower()
                or str(d).strip().lower().startswith(
                    b_ep.lower() + "/"
                )
                for d in a_dests
            ):
                score += 120

            # Una pata tiene como origen el endpoint de la otra.
            if a_ep and b_origen and a_ep == b_origen:
                score += 60

            if b_ep and a_origen and b_ep == a_origen:
                score += 60

            # Dial/AppDial casi siempre representa dos patas.
            if (
                str(a.get("aplicacion", "")).lower() == "dial"
                and str(b.get("aplicacion", "")).lower() == "appdial"
            ):
                score += 80

            if (
                str(b.get("aplicacion", "")).lower() == "dial"
                and str(a.get("aplicacion", "")).lower() == "appdial"
            ):
                score += 80

            if score > mejor_score:
                mejor_score = score
                mejor_j = j

        if mejor_j is None or mejor_score < 100:
            # No hay segunda pata claramente asociada.
            llamadas.append({
                "canal": a.get("canal", ""),
                "origen": a.get("origen", "-"),
                "destino": a.get("destino", "-"),
                "estado": a.get("estado", "-"),
                "duracion": a.get("duracion", "-"),
                "contexto": a.get("contexto", ""),
                "aplicacion": a.get("aplicacion", ""),
                "datos": a.get("datos", ""),
                "direccion": determinar_direccion(a),
            })
            usados.add(i)
            continue

        b = registros[mejor_j]

        # -----------------------------------------------------
        # Determinar origen/destino de la llamada fusionada
        # -----------------------------------------------------
        origen = a.get("origen", "-")
        destino = a.get("destino", "-")

        a_ep = str(a.get("endpoint") or "").strip()
        b_ep = str(b.get("endpoint") or "").strip()

        a_dests = destinos_de(a)
        b_dests = destinos_de(b)

        # Preferimos una pata Dial con destino explícito.
        if str(a.get("aplicacion", "")).lower() == "dial":
            origen = a.get("origen") or a_ep or "-"

            # Si la otra pata tiene un destino más limpio,
            # por ejemplo:
            #   SIP/57320 -> OOH323/ipoffice/6200
            #   OOH323/ipoffice -> 6200
            # mostramos 6200 y no ipoffice/6200.
            if b_dests:
                if b_ep and any(
                    str(d).strip().lower() == b_ep.lower()
                    or str(d).strip().lower().startswith(
                        b_ep.lower() + "/"
                    )
                    for d in a_dests
                ):
                    destino = b_dests[0]
                elif a_dests:
                    destino = a_dests[0]
            elif a_dests:
                destino = a_dests[0]

        elif str(b.get("aplicacion", "")).lower() == "dial":
            origen = b.get("origen") or b_ep or "-"

            if a_dests:
                if a_ep and any(
                    str(d).strip().lower() == a_ep.lower()
                    or str(d).strip().lower().startswith(
                        a_ep.lower() + "/"
                    )
                    for d in b_dests
                ):
                    destino = a_dests[0]
                elif b_dests:
                    destino = b_dests[0]
            elif b_dests:
                destino = b_dests[0]

        # Si el destino todavía coincide con el origen,
        # usar el endpoint de la segunda pata.
        if (
            str(origen).strip()
            and str(destino).strip() == str(origen).strip()
        ):
            if a_ep and a_ep != str(origen).strip():
                destino = a_ep
            elif b_ep and b_ep != str(origen).strip():
                destino = b_ep

        # Duración: tomar la mayor disponible.
        duraciones = []
        for r in (a, b):
            try:
                v = str(r.get("duracion", "")).strip()
                if v not in ("", "-"):
                    duraciones.append(int(v))
            except Exception:
                pass

        duracion = str(max(duraciones)) if duraciones else "-"

        # Si una pata está Up, la llamada está Up.
        estados = [
            str(a.get("estado") or "").strip(),
            str(b.get("estado") or "").strip()
        ]

        if any(e.lower() == "up" for e in estados):
            estado = "Up"
        elif any(e.lower() == "ringing" for e in estados):
            estado = "Ringing"
        elif any(e.lower() == "ring" for e in estados):
            estado = "Ring"
        else:
            estado = next((e for e in estados if e), "-")

        # Elegimos como canal principal la pata Dial,
        # porque es la que contiene el origen y destino real.
        principal = a
        secundaria = b

        if str(b.get("aplicacion", "")).lower() == "dial":
            principal = b
            secundaria = a

        # Conservamos todos los canales reales de Asterisk
        # para poder actuar posteriormente sobre la llamada completa.
        canales_llamada = []

        for r in (principal, secundaria):
            canal_real = str(r.get("canal") or "").strip()

            if canal_real and canal_real not in canales_llamada:
                canales_llamada.append(canal_real)

        callerid_final = str(
            principal.get("callerid") or
            secundaria.get("callerid") or
            ""
        ).strip()

        canal_principal = str(
            principal.get("canal") or ""
        ).strip()

        canal_principal_low = canal_principal.lower()

        if "synway-e1" in canal_principal_low:
            troncal = "synway-e1"
        else:
            troncal = ""

        origen_fue_callerid = bool(
            callerid_final and
            origen and
            not str(origen).startswith("EXTERNO /")
        )

        llamadas.append({
            "canal": principal.get("canal", ""),
            "canales": canales_llamada,
            "origen": origen or "-",
            "destino": destino or "-",
            "estado": estado,
            "duracion": duracion,
            "contexto": principal.get("contexto", ""),
            "aplicacion": principal.get("aplicacion", ""),
            "datos": principal.get("datos", ""),
            "callerid": callerid_final,
            "origen_fuente": (
                "callerid"
                if origen_fue_callerid
                else ("troncal" if troncal else "desconocido")
            ),
            "origen_disponible": bool(
                origen_fue_callerid
            ),
            "troncal": troncal,
        })

        usados.add(i)
        usados.add(mejor_j)

    return llamadas


# Estado real de la troncal OOH323: sondeo TCP al puerto H.225 del peer,
# hecho desde el propio Asterisk. Cacheado para no frenar el ciclo del monitor.
_h225_cache = {}
H225_TTL = 30.0


def _sondear_h225(perfil, host_port):
    """
    True  -> el puerto H.225 responde
    False -> no responde (cerrado o filtrado)
    None  -> no se pudo sondear (no se penaliza la troncal)
    """
    try:
        ip, _, puerto = str(host_port or "").partition(":")
        puerto = puerto or "1720"

        if (
            not re.match(r'^\d+\.\d+\.\d+\.\d+$', ip)
            or not puerto.isdigit()
        ):
            return None

        clave = (str(perfil.get("host", "")), ip, puerto)
        ahora = time.time()
        previo = _h225_cache.get(clave)

        if previo and ahora - previo[0] < H225_TTL:
            return previo[1]

        out, err, code = _ejecutar_comando(
            perfil,
            "timeout 2 bash -c '</dev/tcp/%s/%s' >/dev/null 2>&1; echo $?"
            % (ip, puerto)
        )

        resultado = None

        if code == 0 and out:
            rc = out.strip().splitlines()[-1].strip()

            if rc == "0":
                resultado = True
            elif rc in ("1", "124"):
                resultado = False

        _h225_cache[clave] = (ahora, resultado)
        return resultado

    except Exception as e:
        logger.error("[MONITOR] Error sondeando H.225: %s", e)
        return None




def _vp_custom_trunk_names(perfil, tipo):
    """
    Obtiene nombres de peers custom que tienen estructura de
    troncal administrada por Asterisk/Issabel.

    No considera cualquier peer no numérico como troncal.

    Se consideran candidatos solamente las secciones que tienen:
        type=peer
        context=from-trunk

    Esto permite detectar configuraciones que existen realmente
    en *_custom.conf aunque no tengan fila en asterisk.trunks.
    """

    tipo = str(tipo or "").upper()

    if tipo == "SIP":
        archivo = "/etc/asterisk/sip_custom.conf"
    elif tipo == "IAX2":
        archivo = "/etc/asterisk/iax_custom.conf"
    else:
        return set()

    comando = r"""
FILE='""" + archivo + r"""'

if [ ! -f "$FILE" ]; then
    exit 0
fi

awk '
function limpiar(s) {
    gsub(/^[[:space:]]+/, "", s)
    gsub(/[[:space:]]+$/, "", s)
    return s
}

BEGIN {
    nombre=""
    type_peer=0
    context_trunk=0
}

/^[[:space:]]*\[[^]]+\][[:space:]]*$/ {
    if (nombre != "" && type_peer && context_trunk) {
        print nombre
    }

    nombre=$0
    sub(/^[[:space:]]*\[/, "", nombre)
    sub(/\][[:space:]]*$/, "", nombre)
    nombre=limpiar(nombre)

    type_peer=0
    context_trunk=0
    next
}

/^[[:space:]]*type[[:space:]]*=[[:space:]]*peer([[:space:]]|$)/ {
    type_peer=1
    next
}

/^[[:space:]]*context[[:space:]]*=[[:space:]]*from-trunk([[:space:]]|$)/ {
    context_trunk=1
    next
}

END {
    if (nombre != "" && type_peer && context_trunk) {
        print nombre
    }
}
' "$FILE"
"""

    try:
        resultado = _ejecutar_comando(
            perfil,
            comando
        )
    except Exception as exc:
        logger.warning(
            "[MONITOR] No se pudieron leer troncales custom %s: %s",
            tipo,
            exc
        )
        return set()

    if isinstance(resultado, tuple):
        out = resultado[0]
        code = resultado[2] if len(resultado) > 2 else 0
    else:
        out = resultado
        code = 0

    if code not in (0, None):
        return set()

    nombres = set()

    for linea in str(out or "").splitlines():
        nombre = linea.strip()

        if nombre:
            nombres.add(nombre.lower())

    return nombres


def _parse_trunks(perfil):
    # VP_IP_TRONCALES_SIMPLIFICADAS_V17

    """Lista troncales desde la configuracion y devuelve su host configurado.

    Importante: no ejecuta comandos para sondear el estado operativo. La
    columna de Monitoreo representa la direccion configurada, no disponibilidad.
    """
    import base64

    def consulta_sql(sql):
        sql64 = base64.b64encode(sql.encode('utf-8')).decode('ascii')
        comando = r"""
AMPDBUSER="$(awk -F= '$1=="AMPDBUSER"{print substr($0,index($0,"=")+1)}' /etc/amportal.conf)"
AMPDBPASS="$(awk -F= '$1=="AMPDBPASS"{print substr($0,index($0,"=")+1)}' /etc/amportal.conf)"
AMPDBHOST="$(awk -F= '$1=="AMPDBHOST"{print substr($0,index($0,"=")+1)}' /etc/amportal.conf)"
AMPDBNAME="$(awk -F= '$1=="AMPDBNAME"{print substr($0,index($0,"=")+1)}' /etc/amportal.conf)"
VP_SQL_B64='""" + sql64 + r"""'
VP_SQL="$(printf '%s' "$VP_SQL_B64" | base64 -d)"
if [ -z "$AMPDBUSER" ] || [ -z "$AMPDBNAME" ]; then exit 2; fi
if [ -n "$AMPDBHOST" ]; then
  mysql -N -B -u"$AMPDBUSER" -p"$AMPDBPASS" -h"$AMPDBHOST" "$AMPDBNAME" -e "$VP_SQL" 2>/dev/null
else
  mysql -N -B -u"$AMPDBUSER" -p"$AMPDBPASS" "$AMPDBNAME" -e "$VP_SQL" 2>/dev/null
fi
"""
        out, err, rc = _ejecutar_comando(perfil, comando)
        return (str(out or ''), str(err or ''), rc)

    def host_de_valor(valor):
        """Extrae el host configurado de host=, sip:host o URI equivalente."""
        texto = str(valor or '').strip().strip('"').strip("'")
        if not texto:
            return ''
        bajo = texto.lower()
        if bajo in ('dynamic', ' dinamico', 'dyn', 'dynamic/auto'):
            return 'Local'
        if bajo in ('no', 'none', 'unknown', '0.0.0.0', 'auto'):
            return ''
        # Quitar scheme y credenciales, sin conservar información de usuario.
        texto = re.sub(r'^(?:sip|sips|iax2?|h323)\s*:', '', texto, flags=re.I)
        if '@' in texto:
            texto = texto.rsplit('@', 1)[1]
        texto = re.split(r'[;?,/\s]', texto, 1)[0].strip()
        if texto.startswith('[') and ']' in texto:
            ip = texto[1:texto.index(']')]
            resto = texto[texto.index(']') + 1:]
            return ip + resto if resto.startswith(':') else ip
        # Mostrar la IPv4 sola, sin puerto SIP ni máscara CIDR.
        m_ip = re.match(r'^([0-9]{1,3}(?:\.[0-9]{1,3}){3})(?::[0-9]{1,5})?(?:/[0-9]{1,2})?$', texto)
        if m_ip:
            octetos = m_ip.group(1).split('.')
            if all(0 <= int(o) <= 255 for o in octetos):
                return m_ip.group(1)
        # Conservar IPv6/hostname; la interfaz solo enlazará direcciones IP literales.
        if re.match(r'^[A-Za-z0-9_.:-]+(?:/[0-9]{1,3})?$', texto):
            return texto
        m = re.search(r'([0-9]{1,3}(?:\.[0-9]{1,3}){3}(?::[0-9]{1,5})?(?:/[0-9]{1,2})?)', texto)
        if m:
            return m.group(1)
        return ''

    # Los únicos campos leídos de archivos son parámetros de destino/host;
    # nunca se emiten claves, passwords, secrets o datos de autenticación.
    comando_cfg = r"""
for f in /etc/asterisk/sip*.conf /etc/asterisk/iax*.conf /etc/asterisk/pjsip*.conf /etc/asterisk/ooh323*.conf /etc/asterisk/chan_dahdi*.conf /etc/asterisk/dahdi*.conf; do
  [ -f "$f" ] || continue
  awk '
  function trim(s){gsub(/^[[:space:]]+/,"",s);gsub(/[[:space:]]+$/, "",s);return s}
  /^[[:space:]]*\[[^]]+\][[:space:]]*$/ {sec=$0;sub(/^[[:space:]]*\[/,"",sec);sub(/\][[:space:]]*$/,"",sec);sec=trim(sec);next}
  /^[[:space:]]*[;#]/ {next}
  {
    line=$0; sub(/[;#].*/,"",line); eq=index(line,"="); if(!eq)next;
    key=tolower(trim(substr(line,1,eq-1))); val=trim(substr(line,eq+1));
    if(key=="host" || key=="ip" || key=="ipaddr" || key=="server" || key=="defaultip" || key=="fromdomain" || key=="outboundproxy" || key=="outbound_proxy" || key=="server_uri" || key=="contact" || key=="aors" || key=="endpoint" || key=="match" || key=="type")
      print FILENAME "\t" sec "\t" key "\t" val;
  }' "$f"
done
"""
    cfg_out, cfg_err, cfg_rc = _ejecutar_comando(perfil, comando_cfg)
    secciones = {}
    if cfg_rc == 0 and cfg_out:
        for line in str(cfg_out).splitlines():
            p = line.split('\t', 3)
            if len(p) < 4:
                continue
            archivo, seccion, clave, valor = p[0].strip(), p[1].strip(), p[2].strip().lower(), p[3].strip()
            if not seccion or not clave:
                continue
            key = (archivo, seccion.lower())
            secciones.setdefault(key, {'nombre': seccion, 'valores': {}})['valores'].setdefault(clave, []).append(valor)

    # Leer primero la tabla de troncales: es el inventario configurado de Issabel.
    sql_trunks = (
        "SELECT CAST(trunkid AS CHAR),COALESCE(name,''),LOWER(COALESCE(tech,'')),"
        "COALESCE(channelid,'') FROM trunks ORDER BY trunkid;"
    )
    out, err, rc = consulta_sql(sql_trunks)
    if rc != 0:
        logger.warning('[MONITOR] No se pudo leer la lista de troncales configuradas: %s', err or out or rc)
        return []

    # Fallback de parametros de host guardados en las tablas antiguas SIP/IAX.
    met_out, met_err, met_rc = consulta_sql(
        "SELECT table_name,column_name FROM information_schema.columns "
        "WHERE table_schema=DATABASE() AND table_name IN ('sip','iax','pjsip') "
        "AND column_name IN ('id','keyword','data') ORDER BY table_name,ordinal_position;"
    )
    columnas = {}
    if met_rc == 0:
        for line in str(met_out or '').splitlines():
            p = line.split('\t')
            if len(p) >= 2:
                columnas.setdefault(p[0].strip().lower(), set()).add(p[1].strip().lower())

    cfg_db = {}
    for tabla in ('sip', 'iax', 'pjsip'):
        if not set(('id', 'keyword', 'data')).issubset(columnas.get(tabla, set())):
            continue
        q = (
            "SELECT CAST(id AS CHAR),LOWER(keyword),COALESCE(data,'') FROM `" + tabla + "` "
            "WHERE LOWER(keyword) IN ('host','ip','ipaddr','server','defaultip','fromdomain',"
            "'outboundproxy','outbound_proxy','server_uri','contact','match','aors','endpoint') ORDER BY id,keyword;"
        )
        detail_out, detail_err, detail_rc = consulta_sql(q)
        if detail_rc != 0:
            continue
        for line in str(detail_out or '').splitlines():
            p = line.split('\t', 2)
            if len(p) < 3:
                continue
            cfg_db.setdefault(tabla, {}).setdefault(p[0].strip().lower(), {}).setdefault(p[1].strip().lower(), []).append(p[2].strip())

    def seccion_para(candidatos, tabla_preferida=None):
        candidatos = [str(x or '').strip().lower() for x in candidatos if str(x or '').strip()]
        # Buscar primero en archivos de tecnología correspondiente.
        items = list(secciones.items())
        if tabla_preferida:
            items.sort(key=lambda pair: (0 if tabla_preferida in pair[0][0].lower() else 1, pair[0][0]))
        for candidato in candidatos:
            for (archivo, nombre_sec), info in items:
                if nombre_sec == candidato:
                    return info['valores']
        return {}

    def valor_seccion(valores, claves):
        for clave in claves:
            for value in valores.get(clave, []):
                parsed = host_de_valor(value)
                if parsed:
                    return parsed
        return ''

    def uri_a_host(valor):
        return host_de_valor(valor)

    resultado = []
    for line in str(out or '').splitlines():
        p = line.split('\t')
        if len(p) < 4 or not p[0].strip().isdigit():
            continue
        trunk_id = p[0].strip()
        nombre = p[1].strip() or ('Troncal ' + trunk_id)
        tech = p[2].strip().lower()
        channel = p[3].strip()
        upper_channel = channel.upper()

        if tech in ('iax', 'iax2'):
            label, tabla, prefijos = 'IAX2', 'iax', ('IAX2/', 'IAX/')
        elif tech == 'sip':
            label, tabla, prefijos = 'SIP', 'sip', ('SIP/',)
        elif tech == 'pjsip':
            label, tabla, prefijos = 'PJSIP', 'pjsip', ('PJSIP/',)
        elif tech == 'custom' and upper_channel.startswith('OOH323/'):
            label, tabla, prefijos = 'OOH323', '', ('OOH323/',)
        elif tech in ('dahdi', 'zap') or upper_channel.startswith(('DAHDI/', 'ZAP/')):
            label, tabla, prefijos = ('DAHDI' if tech != 'zap' else 'ZAP/DAHDI'), '', ('DAHDI/', 'ZAP/')
        elif tech == 'custom':
            label, tabla, prefijos = 'CUSTOM', '', ('CUSTOM/',)
        else:
            label, tabla, prefijos = tech.upper() if tech else 'NO INFORMADA', '', ()

        peer = ''
        for prefix in prefijos:
            if upper_channel.startswith(prefix):
                peer = channel[len(prefix):].split('/')[0].strip()
                break
        if not peer:
            match_peer = re.match(r'^(?:SIP|PJSIP|IAX2|IAX|OOH323|DAHDI|ZAP|CUSTOM)/([^/]+)', channel, re.I)
            if match_peer:
                peer = match_peer.group(1).strip()
        candidatos = [peer, nombre, trunk_id]
        host = ''

        # 1) Parámetros de configuración en archivos: sip/iax/ooh323/pjsip.
        valores = seccion_para(candidatos, tabla)
        if label == 'PJSIP':
            # Endpoint -> aor(s) -> contact, sin consultar estado de Asterisk.
            host = valor_seccion(valores, ('host', 'ip', 'ipaddr', 'server', 'outbound_proxy', 'outboundproxy', 'server_uri', 'contact', 'match'))
            if not host:
                aors = []
                for x in valores.get('aors', []):
                    aors.extend([v.strip() for v in x.split(',') if v.strip()])
                for aor in aors:
                    av = seccion_para([aor], 'pjsip')
                    host = valor_seccion(av, ('contact', 'server_uri', 'outbound_proxy', 'outboundproxy', 'host', 'ip', 'ipaddr', 'match'))
                    if host:
                        break
                if not host:
                    # Identify contiene endpoint=<nombre> y match=<IP/red>.
                    for info in secciones.values():
                        v = info['valores']
                        eps = [x.strip().lower() for x in v.get('endpoint', [])]
                        if peer.lower() in eps or nombre.lower() in eps:
                            host = valor_seccion(v, ('match',))
                            if host:
                                break
        else:
            host = valor_seccion(valores, ('host', 'ip', 'ipaddr', 'server', 'defaultip', 'fromdomain', 'outbound_proxy', 'outboundproxy', 'server_uri', 'contact', 'match'))

        # 2) Configuración en tablas SIP/IAX/PJSIP; ids pueden ser trunkid, nombre o peer.
        if not host and tabla and tabla in cfg_db:
            for cand in (trunk_id, nombre.lower(), peer.lower()):
                v = cfg_db[tabla].get(str(cand or '').lower(), {})
                host = valor_seccion(v, ('host', 'ip', 'ipaddr', 'server', 'defaultip', 'fromdomain', 'outbound_proxy', 'outboundproxy', 'server_uri', 'contact', 'match'))
                if host:
                    break
            if not host and tabla in ('sip', 'iax'):
                # Si el canal contiene directamente una dirección, usarla sin hacer sondeo.
                host = uri_a_host(peer)

        if label in ('DAHDI', 'ZAP/DAHDI'):
            host = 'No aplica (troncal física; sin IP)'
        elif not host and label == 'OOH323':
            host = 'Local'
        elif not host and label == 'CUSTOM':
            host = 'Local'
        elif not host:
            host = 'Local'

        resultado.append({
            'nombre': nombre,
            'tipo': label,
            'host': host,
            'ip_configurada': host,
            'canal': channel,
            'estado': 'Configurada',
            'estado_raw': '',
            'online': None
        })

    return resultado


def _parse_cdr(perfil):
    """
    Obtiene los últimos CDR de Issabel usando las credenciales
    definidas en /etc/amportal.conf.
    Funciona tanto en perfiles locales como remotos.
    """

    # Obtener credenciales desde amportal.conf
    cmd = (
        "AMPDBUSER=$(grep '^AMPDBUSER=' /etc/amportal.conf | cut -d= -f2-); "
        "AMPDBPASS=$(grep '^AMPDBPASS=' /etc/amportal.conf | cut -d= -f2-); "
        "mysql -N -B "
        "-u \"$AMPDBUSER\" "
        "-p\"$AMPDBPASS\" "
        "asteriskcdrdb "
        "-e \"SELECT DATE_FORMAT(calldate,'%H:%i:%s'), "
        "src,dst,duration,disposition "
        "FROM cdr "
        "ORDER BY calldate DESC LIMIT 10\" "
        "2>/dev/null"
    )

    out, err, code = _ejecutar_comando(perfil, cmd)

    if code != 0 or not out:
        logger.warning(
            "[MONITOR] No se pudo obtener CDR: %s",
            err or "sin datos"
        )
        return []

    historial = []

    for linea in out.splitlines():
        partes = linea.split("\t")

        if len(partes) < 5:
            continue

        hora = partes[0].strip()
        origen = partes[1].strip()
        destino = partes[2].strip()
        duracion = partes[3].strip()
        estado = partes[4].strip()

        historial.append({
            "hora": hora,
            "origen": origen,
            "destino": destino,
            "duracion": duracion + "s",
            "estado": estado,
            "badge": (
                "badge-success"
                if estado == "ANSWERED"
                else "badge-danger"
            )
        })

    return historial


def obtener_estado_completo(
    perfil,
    perfil_nombre,
    cache=None,
    force_slow=True
):
    if cache is None:
        cache = {}

    try:

        if (
            force_slow
            or not cache.get("slow_ready")
        ):

            ips = _ips_desde_db(perfil)

            cache["peers"] = _parse_peers(
                perfil,
                ips
            )

            cache["troncales"] = _parse_trunks(
                perfil
            )

            cache["historial"] = _parse_cdr(
                perfil
            )

            cache["slow_ready"] = True

        canales = _parse_channels(perfil)

        activos = set()

        for llamada in canales:

            for campo in (
                "origen",
                "destino"
            ):

                valor = str(
                    llamada.get(campo) or ""
                )

                numeros = re.findall(
                    r'\b\d{2,8}\b',
                    valor
                )

                for numero in numeros:
                    activos.add(numero)

        peers = []

        for peer in cache.get(
            "peers",
            []
        ):

            item = dict(peer)

            nombre = str(
                item.get("nombre") or ""
            )

            item["en_llamada"] = (
                nombre in activos
            )

            # Si conocemos la IP del teléfono,
            # se considera disponible.
            if item.get("ip") not in (
                None,
                "",
                "-",
                "0.0.0.0"
            ):
                item["status"] = "Online"

            if item["en_llamada"]:
                item["status"] = "Online"

            peers.append(item)

        return {
            "peers": peers,
            "canales": canales,
            "troncales": list(
                cache.get(
                    "troncales",
                    []
                )
            ),
            "historial": list(
                cache.get(
                    "historial",
                    []
                )
            ),
            "colas": []
        }

    except Exception as e:

        logger.exception(
            "[MONITOR] Error obteniendo estado"
        )

        return {
            "peers": cache.get(
                "peers",
                []
            ),
            "canales": [],
            "troncales": cache.get(
                "troncales",
                []
            ),
            "historial": cache.get(
                "historial",
                []
            ),
            "colas": []
        }


def monitor_asterisk(
    perfil_nombre,
    es_demo=False
):
    from app.config import (
        buscar_perfil,
        cargar_perfiles
    )

    cache = {
        "slow_ready": False
    }

    ultimo_slow = 0

    logger.info(
        "[MONITOR] Iniciando: %s",
        perfil_nombre
    )

    while True:

        try:

            with monitor_lock:

                if (
                    perfil_nombre
                    not in profile_monitors
                ):
                    break

                data = profile_monitors[
                    perfil_nombre
                ]

                perfil = buscar_perfil(
                    perfil_nombre,
                    cargar_perfiles()
                )

            if not perfil:

                logger.error(
                    "[MONITOR] Perfil no encontrado: %s",
                    perfil_nombre
                )

                time.sleep(
                    FAST_INTERVAL
                )

                continue

            ahora = time.time()

            force_slow = (
                ahora - ultimo_slow
                >= SLOW_INTERVAL
                or not cache.get(
                    "slow_ready"
                )
            )

            estado = obtener_estado_completo(
                perfil,
                perfil_nombre,
                cache,
                force_slow=force_slow
            )

            if force_slow:
                ultimo_slow = ahora

            payload = {
                "timestamp": ahora,
                "peers": estado["peers"],
                "canales": estado["canales"],
                "troncales": estado["troncales"],
                "historial": estado["historial"],
                "colas": estado["colas"]
            }

            with monitor_lock:

                data = profile_monitors.get(
                    perfil_nombre
                )

                if not data:
                    break

                data["ultimo_estado"] = payload

                for q in list(
                    data.get(
                        "queues",
                        []
                    )
                ):

                    try:
                        q.put_nowait(
                            payload
                        )

                    except queue.Full:
                        pass

        except Exception as e:

            logger.exception(
                "[MONITOR] Error ciclo: %s",
                e
            )

        time.sleep(
            FAST_INTERVAL
        )
