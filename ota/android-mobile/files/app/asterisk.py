import re
import time
import logging
import subprocess
import os
from app.ssh_client import ejecutar_comando_ssh

logger = logging.getLogger(__name__)

def es_host_local(host):
    """Detecta Issabel local; los clientes instalables siempre trabajan por SSH."""
    if os.environ.get('VOIP_PROVISIONER_DESKTOP') == '1':
        return False
    host = str(host or '').strip().lower().strip('[]')
    return host in ('127.0.0.1', 'localhost', '10.12.9.5', '10.5.0.80', '::1')

def ejecutar_comando_local(cmd):
    try:
        result = subprocess.run(cmd, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=30)
        stdout = result.stdout.decode('utf-8', errors='ignore').strip() if result.stdout else ''
        stderr = result.stderr.decode('utf-8', errors='ignore').strip() if result.stderr else ''
        return stdout, stderr, result.returncode
    except Exception as e:
        return '', str(e), 1

def obtener_credenciales_mysql(perfil):
    """
    Obtiene las credenciales MariaDB/MySQL del Issabel
    correspondiente al perfil.

    Prioridad:

      AMPDBUSER + AMPDBPASS
      DBUSER + DBPASS
      mysqlrootpwd

    Funciona para perfiles locales y remotos.
    """

    cmd = (
        "grep -hE "
        "'^(AMPDBUSER|AMPDBPASS|DBUSER|DBPASS|mysqlrootpwd)=' "
        "/etc/amportal.conf "
        "/etc/issabel.conf "
        "/etc/elastix.conf "
        "2>/dev/null"
    )

    host = str(
        perfil.get("host", "")
    ).strip()

    if es_host_local(host):

        out, err, rc = ejecutar_comando_local(
            cmd
        )

    else:

        out, err, rc = ejecutar_comando_ssh(
            perfil,
            cmd
        )

    valores = {}

    if out:

        for linea in out.splitlines():

            linea = linea.strip()

            if "=" not in linea:
                continue

            clave, valor = linea.split(
                "=",
                1
            )

            clave = clave.strip().upper()
            valor = valor.strip()

            # Quitar comillas externas.
            if (
                len(valor) >= 2
                and (
                    (
                        valor[0] == "'"
                        and valor[-1] == "'"
                    )
                    or
                    (
                        valor[0] == '"'
                        and valor[-1] == '"'
                    )
                )
            ):
                valor = valor[1:-1]

            if valor:
                valores[clave] = valor

    usuario = (
        valores.get("AMPDBUSER")
        or valores.get("DBUSER")
        or "asteriskuser"
    )

    password = (
        valores.get("AMPDBPASS")
        or valores.get("DBPASS")
        or valores.get("MYSQLROOTPWD")
        or "..voip26"
    )

    return usuario, password


def obtener_password_interno_issabel(perfil, interno):
    """
    Consulta silenciosamente un interno en el Issabel seleccionado.

    IMPORTANTE:
    - Primero determina la tecnologia mediante CLI:
        sip show peer
        pjsip show endpoint
    - La contraseña NUNCA se devuelve al navegador.
    - La contraseña solamente se utiliza internamente para la
      creacion individual.
    """

    try:
        interno = str(interno or '').strip()

        if not interno:
            return {
                'existe': False,
                'tecnologia': '',
                'password': '',
                'tiene_password': False
            }

        host = str(perfil.get('host', '') or '').strip()

        def ejecutar(cmd):
            if es_host_local(host):
                return ejecutar_comando_local(cmd)
            return ejecutar_comando_ssh(perfil, cmd)

        # ----------------------------------------------------
        # 1. Primero probar CHAN_SIP
        # ----------------------------------------------------
        out_sip, err_sip, rc_sip = ejecutar(
            'asterisk -rx "sip show peer {}" 2>&1'.format(interno)
        )

        texto_sip = "{}\n{}".format(out_sip or '', err_sip or '')

        existe_sip = False

        if re.search(
            r'Object name:\s*{}'.format(re.escape(interno)),
            texto_sip,
            re.IGNORECASE
        ):
            existe_sip = True

        if re.search(
            r'^\s*{}[/\s]'.format(re.escape(interno)),
            texto_sip,
            re.IGNORECASE | re.MULTILINE
        ):
            existe_sip = True

        if 'Object name not found' in texto_sip:
            existe_sip = False

        # ----------------------------------------------------
        # 2. Si no es SIP, probar PJSIP
        # ----------------------------------------------------
        tecnologia = ''

        if existe_sip:
            tecnologia = 'sip'
        else:
            out_pj, err_pj, rc_pj = ejecutar(
                'asterisk -rx "pjsip show endpoint {}" 2>&1'.format(interno)
            )

            texto_pj = "{}\n{}".format(out_pj or '', err_pj or '')

            existe_pjsip = False

            if re.search(
                r'^\s*Endpoint:\s*{}/'.format(re.escape(interno)),
                texto_pj,
                re.IGNORECASE | re.MULTILINE
            ):
                existe_pjsip = True

            if re.search(
                r'^\s*Endpoint:\s*{}'.format(re.escape(interno)),
                texto_pj,
                re.IGNORECASE | re.MULTILINE
            ):
                existe_pjsip = True

            if 'Endpoint not found' in texto_pj:
                existe_pjsip = False

            if existe_pjsip:
                tecnologia = 'pjsip'

        # ----------------------------------------------------
        # No existe en ninguno de los dos
        # ----------------------------------------------------
        if not tecnologia:
            return {
                'existe': False,
                'tecnologia': '',
                'password': '',
                'tiene_password': False
            }

        # ----------------------------------------------------
        # Obtener credenciales MySQL DEL SERVIDOR SELECCIONADO
        # ----------------------------------------------------
        db_user, db_pass = obtener_credenciales_mysql(perfil)

        # ----------------------------------------------------
        # Ejecutar SQL local/remoto
        # ----------------------------------------------------
        def mysql_consulta(sql):
            import base64

            sql_b64 = base64.b64encode(
                sql.encode('utf-8')
            ).decode('ascii')

            cmd = (
                "echo {} | base64 -d | "
                "mysql -N -B -u{} -p'{}'"
            ).format(
                sql_b64,
                db_user,
                db_pass.replace("'", "'\\''")
            )

            cmd += " 2>/dev/null"

            out, err, rc = ejecutar(cmd)

            if rc != 0:
                raise Exception(
                    "Error consultando MySQL en {}: {}".format(
                        host,
                        err or out or 'error desconocido'
                    )
                )

            return out.strip()

        # ----------------------------------------------------
        # La fuente principal para SIP/PJSIP es asterisk.sip
        # keyword=secret.
        #
        # No devolvemos este valor al frontend.
        # ----------------------------------------------------
        sql = (
            "SELECT data "
            "FROM asterisk.sip "
            "WHERE id='{}' "
            "AND keyword='secret' "
            "ORDER BY id "
            "LIMIT 1"
        ).format(
            interno.replace("'", "''")
        )

        password = mysql_consulta(sql).strip()

        # Fallback a asterisk.users si por alguna razon
        # no existe secret en asterisk.sip.
        if not password:
            sql_users = (
                "SELECT password "
                "FROM asterisk.users "
                "WHERE extension='{}' "
                "LIMIT 1"
            ).format(
                interno.replace("'", "''")
            )

            password = mysql_consulta(sql_users).strip()

        return {
            'existe': True,
            'tecnologia': tecnologia,
            'password': password,
            'tiene_password': bool(password)
        }

    except Exception as e:
        logger.exception(
            "Error consultando interno %s en Issabel: %s",
            interno,
            e
        )

        return {
            'existe': False,
            'tecnologia': '',
            'password': '',
            'tiene_password': False,
            'error': str(e)
        }


def interno_existe_mysql(perfil, interno):
    try:
        db_user, db_pass = obtener_credenciales_mysql(perfil)
        query = "SELECT COUNT(*) FROM asterisk.sip WHERE id='" + interno + "' AND keyword='account';"
        cmd = "/usr/bin/mysql -N -B -u " + db_user + " -p'" + db_pass + "' -e \"" + query + "\" asterisk 2>/dev/null"
        if es_host_local(perfil['host']):
            out, _, c = ejecutar_comando_local(cmd)
        else:
            out, _, c = ejecutar_comando_ssh(perfil, cmd)
        return c == 0 and out and int(out.strip()) > 0
    except Exception as e:
        logger.error("Error interno_existe: " + str(e))
        return False

def sincronizar_internos_rapido(perfil):
    """
    Sincroniza internos desde Issabel.

    Compatibilidad:
      - chan_sip: sip_additional.conf + sip show peers
      - PJSIP: pjsip_additional.conf + pjsip show aor

    No modifica la configuración del Issabel.
    Solamente lee configuración/estado.
    """

    try:
        host = str(perfil.get('host', '')).strip()
        es_local = es_host_local(host)

        def ejecutar(cmd, timeout=30):
            if es_local:
                return ejecutar_comando_local(cmd)

            return ejecutar_comando_ssh(
                perfil,
                cmd,
                timeout=timeout
            )

        # =====================================================
        # 1. INTENTAR CHAN_SIP
        # =====================================================

        cmd_conf = (
            "cat /etc/asterisk/sip_additional.conf "
            "2>/dev/null"
        )

        out_conf, err_conf, c_conf = ejecutar(cmd_conf)

        if c_conf == 0 and out_conf:

            extensiones = {}
            seccion_actual = None

            for linea in out_conf.splitlines():

                linea = linea.strip()

                if (
                    not linea
                    or linea.startswith(';')
                    or linea.startswith('#')
                ):
                    continue

                m = re.match(
                    r'^\[(\d+)\]',
                    linea
                )

                if m:
                    seccion_actual = m.group(1)

                    extensiones[seccion_actual] = {
                        'ip': '',
                        'mac': '',
                        'display_name': seccion_actual,
                        'tecnologia': 'SIP'
                    }

                    continue

                if (
                    seccion_actual
                    and '=' in linea
                ):

                    clave, _, valor = linea.partition('=')

                    clave = clave.strip().lower()
                    valor = valor.strip()

                    if clave == 'callerid':

                        nm = re.match(
                            r'^(.*?)\s*<\d+>$',
                            valor
                        )

                        if nm:
                            nombre = (
                                nm.group(1)
                                .strip()
                                .strip('"')
                                .strip("'")
                            )

                            if (
                                nombre
                                and nombre.lower() != 'device'
                            ):
                                extensiones[
                                    seccion_actual
                                ]['display_name'] = nombre

            if extensiones:

                # IP / estado SIP
                cmd_peers = (
                    "/usr/sbin/asterisk -rx "
                    "'sip show peers' 2>/dev/null"
                )

                out_peers, _, c_peers = ejecutar(cmd_peers)

                if (
                    c_peers == 0
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
                        )

                        peer_host = parts[1]

                        if (
                            peer_name in extensiones
                            and peer_host
                            and '.' in peer_host
                            and 'Unspecified' not in peer_host
                        ):
                            extensiones[
                                peer_name
                            ]['ip'] = peer_host

                # MAC desde AstDB
                cmd_db = (
                    "/usr/sbin/asterisk -rx "
                    "'database show DEVICE' 2>/dev/null"
                )

                out_db, _, c_db = ejecutar(cmd_db)

                if (
                    out_db
                    and c_db == 0
                ):

                    for line in out_db.splitlines():

                        m = re.match(
                            r'^/DEVICE/(\d+)/mac\s*:\s*(\S+)',
                            line.strip()
                        )

                        if (
                            m
                            and m.group(1) in extensiones
                        ):
                            extensiones[
                                m.group(1)
                            ]['mac'] = m.group(2)

                logger.info(
                    "Sync SIP: %s internos obtenidos",
                    len(extensiones)
                )

                return extensiones, ""

        # =====================================================
        # 2. PJSIP
        # =====================================================

        logger.info(
            "Sync: intentando PJSIP para %s",
            host
        )

        cmd_pjsip_conf = (
            "cat /etc/asterisk/pjsip_additional.conf "
            "2>/dev/null"
        )

        out_pjsip, err_pjsip, c_pjsip = ejecutar(
            cmd_pjsip_conf
        )

        if (
            c_pjsip != 0
            or not out_pjsip
        ):
            return (
                {},
                "No se encontró configuración SIP ni PJSIP"
            )

        extensiones = {}

        seccion_actual = None
        seccion_tipo = None

        # -----------------------------------------------------
        # Detectar endpoints numéricos en pjsip_additional.conf
        # -----------------------------------------------------

        for linea in out_pjsip.splitlines():

            linea = linea.strip()

            if (
                not linea
                or linea.startswith(';')
                or linea.startswith('#')
            ):
                continue

            m = re.match(
                r'^\[(\d+)\]',
                linea
            )

            if m:

                nombre = m.group(1)

                seccion_actual = nombre
                seccion_tipo = None

                extensiones.setdefault(
                    nombre,
                    {
                        'ip': '',
                        'mac': '',
                        'display_name': nombre,
                        'tecnologia': 'PJSIP'
                    }
                )

                continue

            if (
                seccion_actual
                and '=' in linea
            ):

                clave, _, valor = linea.partition('=')

                clave = clave.strip().lower()
                valor = valor.strip()

                if clave == 'type':
                    seccion_tipo = valor.lower()

                elif (
                    clave == 'callerid'
                    and seccion_actual in extensiones
                ):

                    nm = re.match(
                        r'^(.*?)\s*<[^>]+>$',
                        valor
                    )

                    if nm:
                        nombre = (
                            nm.group(1)
                            .strip()
                            .strip('"')
                            .strip("'")
                        )

                        if nombre:
                            extensiones[
                                seccion_actual
                            ]['display_name'] = nombre

        # -----------------------------------------------------
        # Solo endpoints numéricos.
        #
        # pjsip_additional.conf puede tener:
        #   [1993] endpoint
        #   [1993] aor
        #   [auth1993] auth
        #
        # El mismo nombre numérico puede aparecer más de una vez.
        # Eso es normal.
        # -----------------------------------------------------

        if not extensiones:

            return (
                {},
                "pjsip_additional.conf no contiene internos numéricos"
            )

        # -----------------------------------------------------
        # Obtener estado/contacto desde AOR.
        #
        # No usamos pjsip show endpoints para decidir Online.
        # La fuente de verdad de registro es el AOR/contact.
        # -----------------------------------------------------

        for interno in sorted(
            extensiones.keys()
        ):

            cmd_aor = (
                "/usr/sbin/asterisk -rx "
                "'pjsip show aor {}' 2>/dev/null"
            ).format(interno)

            out_aor, _, c_aor = ejecutar(
                cmd_aor
            )

            if (
                c_aor != 0
                or not out_aor
            ):
                continue

            contact_ip = ""

            for linea in out_aor.splitlines():

                linea = linea.strip()

                # Ejemplo:
                # Contact: 1993/sip:1993@10.0.0.10:5060
                #
                # También soportamos:
                # sip:1993@10.0.0.10:5060

                m = re.search(
                    r'sip:[^@]+@'
                    r'(\d+\.\d+\.\d+\.\d+)'
                    r'(?::\d+)?',
                    linea,
                    re.I
                )

                if m:
                    contact_ip = m.group(1)
                    break

            if contact_ip:
                extensiones[
                    interno
                ]['ip'] = contact_ip

        logger.info(
            "Sync PJSIP: %s internos obtenidos",
            len(extensiones)
        )

        return extensiones, ""

    except Exception as e:

        logger.error(
            "Error sync SIP/PJSIP: %s",
            e,
            exc_info=True
        )

        return {}, str(e)

def recargar_asterisk(perfil):
    if es_host_local(perfil['host']):
        ejecutar_comando_local("/var/lib/asterisk/bin/retrieve_conf")
        time.sleep(1.5)
        ejecutar_comando_local("/usr/sbin/asterisk -rx 'sip reload'")
    else:
        ejecutar_comando_ssh(perfil, "/var/lib/asterisk/bin/retrieve_conf")
        time.sleep(1.5)
        ejecutar_comando_ssh(perfil, "/usr/sbin/asterisk -rx 'sip reload'")
    time.sleep(1)


def poblar_astdb(perfil, interno, password, nombre, tecnologia='sip'):
    """Puebla AstDB siguiendo el formato de Issabel para SIP/PJSIP."""
    try:
        tecnologia = str(tecnologia or 'sip').strip().lower()
        if tecnologia not in ('sip', 'pjsip'):
            tecnologia = 'sip'

        es_pjsip = tecnologia == 'pjsip'
        dial = ('PJSIP/' if es_pjsip else 'SIP/') + interno

        entries = [
            ("AMPUSER/" + interno, "cidname", nombre),
            ("AMPUSER/" + interno, "cidnum", interno),
            ("AMPUSER/" + interno, "device", interno),
            ("AMPUSER/" + interno, "password", password),
            ("AMPUSER/" + interno, "voicemail", "novm"),
            ("AMPUSER/" + interno, "recording", "none"),
            ("AMPUSER/" + interno, "recording/in/external", "dontcare"),
            ("AMPUSER/" + interno, "recording/in/internal", "dontcare"),
            ("AMPUSER/" + interno, "recording/out/external", "dontcare"),
            ("AMPUSER/" + interno, "recording/out/internal", "dontcare"),
            ("AMPUSER/" + interno, "recording/ondemand", "disabled"),
            ("AMPUSER/" + interno, "recording/priority", "10"),
            ("AMPUSER/" + interno, "ringtimer", "0"),
            ("AMPUSER/" + interno, "cfringtimer", "0"),
            ("AMPUSER/" + interno, "concurrency_limit", "0"),
            ("AMPUSER/" + interno, "noanswer", "none"),
            ("AMPUSER/" + interno, "outboundcid", '"' + nombre + '" <' + interno + '>'),
            ("AMPUSER/" + interno, "queues/qnostate", "usestate"),
            ("AMPUSER/" + interno, "language", "es"),
            ("AMPUSER/" + interno, "answermode", "disabled"),
            ("DEVICE/" + interno, "dial", dial),
            ("DEVICE/" + interno, "type", "fixed"),
            ("DEVICE/" + interno, "user", interno),
            ("DEVICE/" + interno, "default_user", interno),
        ]

        commands = []

        for familia, clave, valor in entries:
            commands.append(
                "database put {} {} {}".format(
                    familia,
                    clave,
                    valor
                )
            )

        full_cmd = " && ".join(
            ["/usr/sbin/asterisk -rx '{}'".format(c) for c in commands]
        )

        if es_host_local(perfil['host']):
            out, err, c = ejecutar_comando_local(full_cmd + " 2>&1")
        else:
            out, err, c = ejecutar_comando_ssh(
                perfil,
                full_cmd + " 2>&1"
            )

        if c != 0:
            logger.error(
                "ASTDB: error poblando {}: {}".format(
                    interno,
                    out[:300] if out else err[:300]
                )
            )
            return False, "Error poblando AstDB"

        verify_cmd = (
            "/usr/sbin/asterisk -rx "
            "'database show DEVICE/{}' 2>/dev/null | wc -l"
        ).format(interno)

        if es_host_local(perfil['host']):
            v_out, _, _ = ejecutar_comando_local(verify_cmd)
        else:
            v_out, _, _ = ejecutar_comando_ssh(perfil, verify_cmd)

        count = (
            int(v_out.strip())
            if v_out and v_out.strip().isdigit()
            else 0
        )

        logger.info(
            "ASTDB: {} tecnologia={} - {} entradas DEVICE verificadas".format(
                interno,
                tecnologia,
                count
            )
        )

        if count == 0:
            logger.error(
                "ASTDB: FALLÓ poblar DEVICE/{}".format(interno)
            )
            return False, "AstDB vacío después de poblar"

        return True, "OK ({} entries DEVICE)".format(count)

    except Exception as e:
        logger.error("ASTDB error: {}".format(e), exc_info=True)
        return False, str(e)


def crear_peer_en_asterisk(perfil, interno, password,
                            contexto='from-internal',
                            nombre_mostrar='',
                            nombre_saliente=''):
    """
    Crea o actualiza un interno en Issabel usando SIP o PJSIP.

    La tecnologia se obtiene del perfil recibido.
    Mantiene las tablas users/devices/sip utilizadas por Issabel.
    """

    import subprocess

    try:
        db_user, db_pass = obtener_credenciales_mysql(perfil)

        interno = str(interno or '').strip()
        password = str(password or '')
        nm = str(nombre_mostrar or interno).strip() or interno
        ns = str(nombre_saliente or interno).strip() or interno

        if not interno:
            return False, 'Interno vacio'

        tecnologia = str(
            perfil.get('tecnologia', 'sip') or 'sip'
        ).strip().lower()

        if tecnologia not in ('sip', 'pjsip'):
            return False, (
                'Tecnologia invalida para {}: {}'.format(
                    interno,
                    tecnologia
                )
            )

        es_pjsip = tecnologia == 'pjsip'

        def esc(valor):
            return str(valor).replace("'", "''")

        tech = 'pjsip' if es_pjsip else 'sip'
        dial = ('PJSIP/' if es_pjsip else 'SIP/') + interno
        callerid = '"{}" <{}>'.format(nm, interno)

        # ============================================================
        # USERS
        # ============================================================

        sql = (
            "REPLACE INTO asterisk.users "
            "(extension, name, password, voicemail, ringtimer, "
            "noanswer, recording, outboundcid, sipname, mohclass) "
            "VALUES ("
            "'{}', ".format(esc(interno)) +
            "'{}', ".format(esc(nm)) +
            "'{}', ".format(esc(password)) +
            "'novm', "
            "0, '', '', "
            "'{}', ".format(esc(callerid)) +
            "'{}', ".format(esc(interno)) +
            "NULL"
            "); "
        )

        # ============================================================
        # DEVICES
        # ============================================================

        sql += (
            "REPLACE INTO asterisk.devices "
            "(id, tech, dial, devicetype, user, description, emergency_cid) "
            "VALUES ("
            "'{}', ".format(esc(interno)) +
            "'{}', ".format(tech) +
            "'{}', ".format(esc(dial)) +
            "'fixed', "
            "'{}', ".format(esc(interno)) +
            "'{}', ".format(esc(nm)) +
            "'{}'"
            "); "
        )

        # ============================================================
        # ASTERISK.SIP
        #
        # En este Issabel los internos PJSIP existentes tambien
        # tienen sus parametros en esta tabla.
        # ============================================================

        filas = [
            ('account', interno),
            ('accountcode', ''),
            ('allow_subscribe', 'yes'),
            ('authenticate_qualify', 'no'),
            ('callerid', callerid),
            ('context', contexto),
            ('direct_media', 'no'),
            ('disallow', ''),
            ('allow', ''),
            ('dtls_ca_file', ''),
            ('dtls_cert_file', ''),
            ('dtls_private_key', ''),
            ('dtls_setup', 'actpass'),
            ('dtls_verify', 'no'),
            ('dtmfmode', 'rfc2833'),
            ('force_rport', 'yes'),
            ('host', 'dynamic'),
            ('ice_support', 'no'),
            ('mailbox', interno + '@device'),
            ('max_contacts', '1'),
            ('media_encryption', 'no'),
            ('media_use_received_transport', 'no'),
            ('message_context', ''),
            ('nat', 'no'),
            ('permit', '0.0.0.0/0.0.0.0'),
            ('qualify', 'yes'),
            ('qualifyfreq', '60'),
            ('rewrite_contact', 'yes'),
            ('rtcp_mux', 'no'),
            ('rtp_symmetric', 'yes'),
            ('secret', password),
            ('sendrpid', 'no'),
            ('stir_shaken', 'off'),
            ('subscribe_context', ''),
            ('transport', 'transport-udp'),
            ('trustrpid', 'yes'),
            ('trust_id_inbound', 'yes'),
            ('type', 'friend'),
            ('use_avpf', 'no')
        ]

        if not es_pjsip:
            filas.append(('canreinvite', 'no'))

        sql += (
            "DELETE FROM asterisk.sip "
            "WHERE id='{}'; ".format(esc(interno))
        )

        for keyword, data in filas:
            sql += (
                "INSERT INTO asterisk.sip "
                "(id, keyword, data, flags) "
                "VALUES ('{}', '{}', '{}', ''); ".format(
                    esc(interno),
                    esc(keyword),
                    esc(data)
                )
            )

        # ============================================================
        # EJECUTAR MYSQL
        # ============================================================

        comando = [
            'mysql',
            '-u' + str(db_user),
            '-p' + str(db_pass),
            '-e',
            sql
        ]

        resultado = subprocess.run(
            comando,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )

        if resultado.returncode != 0:
            error = resultado.stderr.decode(
                'utf-8',
                'replace'
            ).strip()

            return False, (
                'Error SQL creando {} {}: {}'.format(
                    tecnologia.upper(),
                    interno,
                    error
                )
            )

        # ============================================================
        # VERIFICAR USERS
        # ============================================================

        ver_users = subprocess.run(
            [
                'mysql',
                '-u' + str(db_user),
                '-p' + str(db_pass),
                '-N',
                '-B',
                '-e',
                (
                    "SELECT COUNT(*) "
                    "FROM asterisk.users "
                    "WHERE extension='{}';"
                ).format(esc(interno))
            ],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )

        if ver_users.returncode != 0:
            return False, (
                'No se pudo verificar users/{}: {}'.format(
                    interno,
                    ver_users.stderr.decode(
                        'utf-8',
                        'replace'
                    ).strip()
                )
            )

        if ver_users.stdout.decode(
            'utf-8',
            'replace'
        ).strip() != '1':
            return False, (
                'El interno {} no quedo creado en asterisk.users'.format(
                    interno
                )
            )

        # ============================================================
        # VERIFICAR DEVICES
        # ============================================================

        ver_devices = subprocess.run(
            [
                'mysql',
                '-u' + str(db_user),
                '-p' + str(db_pass),
                '-N',
                '-B',
                '-e',
                (
                    "SELECT COUNT(*) "
                    "FROM asterisk.devices "
                    "WHERE id='{}' AND tech='{}';"
                ).format(
                    esc(interno),
                    tech
                )
            ],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )

        if ver_devices.returncode != 0:
            return False, (
                'No se pudo verificar devices/{}: {}'.format(
                    interno,
                    ver_devices.stderr.decode(
                        'utf-8',
                        'replace'
                    ).strip()
                )
            )

        if ver_devices.stdout.decode(
            'utf-8',
            'replace'
        ).strip() != '1':
            return False, (
                'El interno {} no quedo creado como {}'.format(
                    interno,
                    tech
                )
            )

        # ============================================================
        # ASTDB
        #
        # poblar_astdb() ya existe en este proyecto y es la funcion
        # que mantiene AMPUSER/DEVICE como los internos existentes.
        # ============================================================

        try:
            poblar_astdb(
                perfil,
                interno,
                password,
                nm,
                tecnologia
            )
        except Exception as e:
            return False, (
                'Interno {} creado en MySQL pero fallo poblar_astdb: {}'.format(
                    interno,
                    e
                )
            )

        # ============================================================
        # RETRIEVE_CONF Y RECARGA
        # ============================================================

        host = str(
            perfil.get('host', '') or ''
        ).strip()

        es_host_local_flag = (
            os.environ.get('VOIP_PROVISIONER_DESKTOP') != '1'
            and host.lower().strip('[]') in ('127.0.0.1', 'localhost', '10.12.9.5', '10.5.0.80', '::1')
        )

        if es_host_local_flag:

            r = subprocess.run(
                ['asterisk', '-rx', 'retrieve_conf'],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE
            )

            if r.returncode != 0:
                return False, (
                    'Interno {} creado pero retrieve_conf fallo: {}'.format(
                        interno,
                        r.stderr.decode(
                            'utf-8',
                            'replace'
                        ).strip()
                    )
                )

            if es_pjsip:
                r = subprocess.run(
                    ['asterisk', '-rx', 'pjsip reload'],
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE
                )

                if r.returncode != 0:
                    return False, (
                        'Interno {} creado pero pjsip reload fallo: {}'.format(
                            interno,
                            r.stderr.decode(
                                'utf-8',
                                'replace'
                            ).strip()
                        )
                    )

        return True, ''

    except Exception as e:
        return False, (
            'Excepcion creando interno {}: {}'.format(
                interno if 'interno' in locals() else '',
                e
            )
        )




def crear_interno_individual_en_asterisk(
        perfil,
        interno,
        password,
        contexto='from-internal',
        nombre_mostrar='',
        nombre_saliente='',
        tecnologia='sip', max_contactos=None):
    """
    Crea un unico interno en Issabel respetando SIP/PJSIP.

    IMPORTANTE:
      - Si el perfil apunta al servidor local, ejecuta MySQL localmente.
      - Si el perfil apunta a un servidor remoto, ejecuta MySQL por SSH.
      - La carga masiva NO utiliza esta funcion.
    """

    try:
        import base64
        import shlex

        interno = str(interno or '').strip()
        password = str(password or '').strip()

        if not interno:
            return False, "Interno vacio"

        tecnologia = str(
            tecnologia or perfil.get('tecnologia', 'sip')
        ).strip().lower()

        if tecnologia not in ('sip', 'pjsip'):
            tecnologia = 'sip'

        if not password:
            password = "pjn{}".format(interno)

        nm = str(nombre_mostrar or interno).strip() or interno
        ns = str(nombre_saliente or interno).strip() or interno

        es_pjsip = tecnologia == 'pjsip'
        tech = 'pjsip' if es_pjsip else 'sip'
        dial = ('PJSIP/' if es_pjsip else 'SIP/') + interno

        logger.info(
            "INDIVIDUAL: interno=%s tecnologia=%s tech=%s dial=%s host=%s",
            interno,
            tecnologia,
            tech,
            dial,
            perfil.get('host', '')
        )

        perfil_individual = dict(perfil)
        perfil_individual['tecnologia'] = tecnologia

        db_user, db_pass = obtener_credenciales_mysql(
            perfil_individual
        )

        def esc(v):
            return str(v).replace("'", "''")

        # ----------------------------------------------------
        # TODA ejecucion local/remota pasa por este wrapper.
        # ----------------------------------------------------
        def ejecutar(cmd):
            if es_host_local(perfil.get('host', '')):
                return ejecutar_comando_local(cmd)

            return ejecutar_comando_ssh(
                perfil,
                cmd
            )

        # ----------------------------------------------------
        # Ejecutar SQL de forma segura mediante base64.
        # Esto evita problemas con comillas, saltos de linea
        # y caracteres especiales del SQL.
        # ----------------------------------------------------
        def mysql_script(sql):
            sql_b64 = base64.b64encode(
                sql.encode('utf-8')
            ).decode('ascii')

            cmd = (
                "SQL_B64={}; "
                "printf '%s' \"$SQL_B64\" | base64 -d | "
                "/usr/bin/mysql -A -u {} -p{} asterisk 2>&1"
            ).format(
                shlex.quote(sql_b64),
                shlex.quote(db_user),
                shlex.quote(db_pass)
            )

            return ejecutar(cmd)

        def mysql_consulta(sql):
            sql_b64 = base64.b64encode(
                sql.encode('utf-8')
            ).decode('ascii')

            cmd = (
                "SQL_B64={}; "
                "SQL=\"$(printf '%s' \"$SQL_B64\" | base64 -d)\"; "
                "/usr/bin/mysql -N -B -u {} -p{} "
                "-e \"$SQL\" asterisk 2>/dev/null"
            ).format(
                shlex.quote(sql_b64),
                shlex.quote(db_user),
                shlex.quote(db_pass)
            )

            return ejecutar(cmd)

        # ====================================================
        # CONTACTOS SIMULTANEOS PJSIP
        # Si no se envia una opcion, conservar la configuracion
        # existente. Los valores permitidos son 1 o 2.
        # ====================================================
        if es_pjsip:
            if max_contactos is None or str(max_contactos).strip() == '':
                q_max = (
                    "SELECT data FROM asterisk.sip "
                    "WHERE id='{}' AND keyword='max_contacts' "
                    "LIMIT 1;"
                ).format(esc(interno))
                max_out, max_err, max_rc = mysql_consulta(q_max)
                max_contactos = (
                    str(max_out or '').strip() if max_rc == 0 else ''
                )

            max_contactos = str(max_contactos or '1').strip()
            if max_contactos not in ('1', '2'):
                max_contactos = '1'
        else:
            max_contactos = '1'

        # ====================================================
        # USERS
        # ====================================================

        sql = (
            "REPLACE INTO asterisk.users "
            "(extension,name,password,voicemail,ringtimer,"
            "noanswer,recording,outboundcid,sipname,mohclass) "
            "VALUES ('{}','{}','{}','novm',0,'','','{}','{}',NULL); "
        ).format(
            esc(interno),
            esc(nm),
            esc(password),
            esc('"' + ns + '" <' + interno + '>'),
            esc(interno)
        )

        # ====================================================
        # DEVICES
        # ====================================================

        sql += (
            "REPLACE INTO asterisk.devices "
            "(id,tech,dial,devicetype,user,description,"
            "emergency_cid) "
            "VALUES ('{}','{}','{}','fixed','{}','{}','{}'); "
        ).format(
            esc(interno),
            esc(tech),
            esc(dial),
            esc(interno),
            esc(nm),
            esc(interno)
        )

        # ====================================================
        # CONFIGURACION BASE
        # ====================================================

        filas = [
            ('account', interno),
            ('accountcode', ''),
            ('callerid', '{} <{}>'.format(nm, interno)),
            ('context', contexto),
            ('dtmfmode', 'rfc2833'),
            ('host', 'dynamic'),
            ('nat', 'no'),
            ('qualify', 'yes'),
            ('qualifyfreq', '60'),
            ('qualify_timeout', '3'),
            ('trustrpid', 'yes'),
            ('sendrpid', 'no'),
            ('type', 'friend'),
        ]

        # ====================================================
        # PJSIP
        # ====================================================

        if es_pjsip:

            filas.extend([
                ('allow_subscribe', 'yes'),
                ('authenticate_qualify', 'no'),
                ('direct_media', 'no'),

                ('dtls_ca_file', ''),
                ('dtls_cert_file', ''),
                ('dtls_private_key', ''),
                ('dtls_setup', 'actpass'),
                ('dtls_verify', 'no'),

                ('force_rport', 'yes'),
                ('ice_support', 'no'),

                ('mailbox', interno + '@device'),
                ('max_contacts', str(max_contactos)),

                ('media_encryption', 'no'),
                ('media_use_received_transport', 'no'),
                ('message_context', ''),

                ('permit', '0.0.0.0/0.0.0.0'),
                ('rewrite_contact', 'yes'),

                ('rtcp_mux', 'no'),
                ('rtp_symmetric', 'yes'),

                ('stir_shaken', 'off'),
                ('subscribe_context', ''),

                # IMPORTANTE:
                # En este Issabel PJSIP escucha por transport-udp
                # en 0.0.0.0:5066.
                ('transport', 'transport-udp'),

                ('trust_id_inbound', 'yes'),
                ('use_avpf', 'no'),

                ('secret', password),

                # Una sola entrada de cada una.
                # Evita el error Duplicate entry ...-disallow.
                ('disallow', 'all'),
                ('allow', 'ilbc,ulaw,alaw,gsm,opus,g729'),

                ('deny', '0.0.0.0/0.0.0.0'),
                ('pickupgroup', ''),
                ('callgroup', ''),
            ])

        # ====================================================
        # CHAN_SIP
        # ====================================================

        else:

            filas.extend([
                ('secret', password),
                ('disallow', ''),
                ('allow', ''),
                ('canreinvite', 'no'),
            ])

        # Preservar solo las claves validas para cada tecnologia.
        if es_pjsip:
            claves_grupos_preservar = (
                'callgroup', 'pickupgroup',
                'named_call_group', 'named_pickup_group'
            )
        else:
            claves_grupos_preservar = (
                'callgroup', 'pickupgroup',
                'namedcallgroup', 'namedpickupgroup'
            )

        sql_claves_grupos = ",".join(
            "'{}'".format(k) for k in claves_grupos_preservar
        )
        q_grupos = (
            "SELECT keyword, data FROM asterisk.sip "
            "WHERE id='{}' AND keyword IN ({}) "
            "AND data <> '';"
        ).format(esc(interno), sql_claves_grupos)

        grupos_out, grupos_err, grupos_rc = mysql_consulta(q_grupos)

        if grupos_rc != 0:
            return False, (
                "No se pudieron consultar los grupos de captura "
                "del interno {}: {}".format(
                    interno, grupos_err or grupos_out or ''
                )
            )

        for linea in str(grupos_out or '').splitlines():
            partes = linea.split('\t', 1)

            if len(partes) != 2:
                continue

            keyword = partes[0].strip()
            valor = partes[1].strip()

            if keyword not in claves_grupos_preservar or not valor:
                continue

            for indice, fila in enumerate(filas):
                if fila[0] == keyword:
                    filas[indice] = (keyword, valor)
                    break
            else:
                filas.append((keyword, valor))

        # ====================================================
        # LIMPIAR CONFIGURACION ANTERIOR
        # ====================================================

        sql += (
            "DELETE FROM asterisk.sip "
            "WHERE id='{}'; "
        ).format(
            esc(interno)
        )

        # ====================================================
        # INSERTAR CONFIGURACION
        #
        # REPLACE evita conflictos con PRIMARY (id,keyword).
        # flags es INTEGER.
        # ====================================================

        for keyword, data in filas:

            sql += (
                "REPLACE INTO asterisk.sip "
                "(id,keyword,data,flags) "
                "VALUES ('{}','{}','{}',0); "
            ).format(
                esc(interno),
                esc(keyword),
                esc(data)
            )

        # ====================================================
        # EJECUTAR MYSQL
        #
        # LOCAL  -> .5 directamente
        # REMOTO -> SSH al servidor del perfil
        # ====================================================

        stdout_sql, stderr_sql, rc_sql = mysql_script(sql)

        if rc_sql != 0:

            logger.error(
                "INDIVIDUAL SQL ERROR %s host=%s: %s",
                interno,
                perfil.get('host', ''),
                stderr_sql or stdout_sql
            )

            return False, (
                "Error SQL: " +
                (stderr_sql or stdout_sql)[:700]
            )

        # ====================================================
        # VERIFICAR USERS
        # ====================================================

        q_users = (
            "SELECT extension,name,password "
            "FROM asterisk.users "
            "WHERE extension='{}';"
        ).format(
            esc(interno)
        )

        users_out, users_err, users_rc = mysql_consulta(q_users)

        if users_rc != 0 or not users_out:

            return False, (
                "No se pudo verificar asterisk.users para " +
                interno
            )

        # ====================================================
        # VERIFICAR DEVICES
        # ====================================================

        q_device = (
            "SELECT id,tech,dial,devicetype,user "
            "FROM asterisk.devices "
            "WHERE id='{}';"
        ).format(
            esc(interno)
        )

        device_out, device_err, device_rc = mysql_consulta(q_device)

        if device_rc != 0 or not device_out:

            return False, (
                "No se pudo verificar asterisk.devices para " +
                interno
            )

        # ====================================================
        # VERIFICAR SIP
        # ====================================================

        q_sip = (
            "SELECT COUNT(*) "
            "FROM asterisk.sip "
            "WHERE id='{}';"
        ).format(
            esc(interno)
        )

        sip_out, sip_err, sip_rc = mysql_consulta(q_sip)

        sip_out = str(sip_out or '0').strip()

        if sip_rc != 0 or sip_out == '0':

            return False, (
                "No se pudo verificar asterisk.sip para " +
                interno
            )

        # ====================================================
        # ASTDB
        #
        # poblar_astdb() ya conoce la diferencia entre host
        # local y remoto y utiliza SSH cuando corresponde.
        # ====================================================

        try:

            poblar_astdb(
                perfil_individual,
                interno,
                password,
                nm,
                tecnologia
            )

        except Exception as e:

            logger.error(
                "INDIVIDUAL ASTDB ERROR %s host=%s: %s",
                interno,
                perfil.get('host', ''),
                e
            )

            return False, (
                "Error poblando AstDB: {}".format(e)
            )

        # ====================================================
        # RECARGAR CONFIGURACION EN EL SERVIDOR SELECCIONADO
        # ====================================================

        comandos_reload = [
            '/var/lib/asterisk/bin/retrieve_conf',
            'asterisk -rx "dialplan reload"',
            'asterisk -rx "sip reload"',
        ]

        if es_pjsip:
            comandos_reload.append(
                'asterisk -rx "pjsip reload"'
            )

        comandos_reload.append(
            'asterisk -rx "core reload"'
        )

        for cmd in comandos_reload:
            rr_out, rr_err, rr_rc = ejecutar(cmd)

            if rr_rc != 0:
                detalle = str(rr_err or rr_out or '').strip()

                logger.error(
                    "INDIVIDUAL reload fallo host=%s cmd=%s: %s",
                    perfil.get('host', ''),
                    cmd,
                    detalle
                )

                return False, (
                    "Interno {} guardado, pero fallo el comando "
                    "{}: {}".format(
                        interno,
                        cmd,
                        detalle[:500] or 'sin detalle'
                    )
                )

        logger.info(
            "INDIVIDUAL OK: interno=%s tecnologia=%s tech=%s host=%s",
            interno,
            tecnologia,
            tech,
            perfil.get('host', '')
        )

        return True, ""

    except Exception as e:

        logger.exception(
            "Error en crear_interno_individual_en_asterisk: %s",
            e
        )

        return False, str(e)


def habilitar_paralelo_pjsip(perfil, interno, maximo=2):
    """
    Habilita un segundo telefono (paralelo) para un interno PJSIP existente.

    SOLO sube max_contacts en asterisk.sip. No toca users, devices, astdb,
    secret ni ningun otro parametro, y no elimina contactos registrados:
    el telefono original queda exactamente como estaba.
    La recarga la aplica quien llama (recargar_issabel_completo).
    """
    try:
        import base64
        import shlex

        interno = str(interno or '').strip()
        if not interno:
            return False, "Interno vacio"

        p = dict(perfil)
        p['tecnologia'] = 'pjsip'
        db_user, db_pass = obtener_credenciales_mysql(p)

        def esc(v):
            return str(v).replace("'", "''")

        def ejecutar(cmd):
            if es_host_local(perfil.get('host', '')):
                return ejecutar_comando_local(cmd)
            return ejecutar_comando_ssh(perfil, cmd)

        def mysql(sql):
            sql_b64 = base64.b64encode(sql.encode('utf-8')).decode('ascii')
            cmd = (
                "SQL_B64={}; "
                "SQL=\"$(printf '%s' \"$SQL_B64\" | base64 -d)\"; "
                "/usr/bin/mysql -N -B -u {} -p{} "
                "-e \"$SQL\" asterisk 2>/dev/null"
            ).format(
                shlex.quote(sql_b64),
                shlex.quote(db_user),
                shlex.quote(db_pass)
            )
            return ejecutar(cmd)

        out, err, rc = mysql(
            "SELECT tech FROM asterisk.devices WHERE id='{}' LIMIT 1;".format(
                esc(interno))
        )
        if rc != 0 or str(out or '').strip().lower() != 'pjsip':
            return False, "El interno {} no es PJSIP en Issabel".format(interno)

        out, err, rc = mysql(
            "SELECT data FROM asterisk.sip WHERE id='{}' "
            "AND keyword='max_contacts' LIMIT 1;".format(esc(interno))
        )
        actual = str(out or '').strip() if rc == 0 else ''
        if actual.isdigit() and int(actual) >= int(maximo):
            return True, ""

        # REPLACE sobre la clave (id,keyword): solo max_contacts.
        out, err, rc = mysql(
            "REPLACE INTO asterisk.sip (id,keyword,data,flags) "
            "VALUES ('{}','max_contacts','{}',0);".format(
                esc(interno), int(maximo))
        )
        if rc != 0:
            return False, "No se pudo actualizar max_contacts de {}".format(interno)

        out, err, rc = mysql(
            "SELECT data FROM asterisk.sip WHERE id='{}' "
            "AND keyword='max_contacts' LIMIT 1;".format(esc(interno))
        )
        if rc != 0 or str(out or '').strip() != str(int(maximo)):
            return False, "No se pudo verificar max_contacts de {}".format(interno)

        logger.info(
            "PARALELO: max_contacts=%s interno=%s host=%s",
            maximo, interno, perfil.get('host', '')
        )
        return True, ""

    except Exception as e:
        logger.exception("Error en habilitar_paralelo_pjsip: %s", e)
        return False, str(e)


def crear_internos_lote_en_asterisk(perfil, internos, obtener_password,
                                    contexto='from-internal', tecnologia='sip',
                                    progreso=None, tanda_sql=20, tanda_astdb=5, idioma='es'):
    # Crea varios internos en Issabel por tandas.
    #
    # - Un interno que ya existe en users, devices o sip NO se toca.
    # - Cada comando viaja como UN argumento de shell, que en Linux no puede pasar
    #   de 128 KB: por eso el SQL va en tandas de 'tanda_sql' internos y AstDB en
    #   subtandas de 'tanda_astdb'.
    # - No recarga Asterisk: eso lo hace quien llama, una sola vez, al final.
    # - Devuelve un dict: creados, existentes, fallidos, no_procesados, error.
    import base64
    import re as _re

    resultado = {'creados': [], 'existentes': [], 'fallidos': [],
                 'no_procesados': [], 'error': ''}

    try:
        internos = [str(x) for x in internos]
        if not internos:
            return resultado

        tecnologia = str(tecnologia or perfil.get('tecnologia', 'sip')).strip().lower()
        if tecnologia not in ('sip', 'pjsip'):
            tecnologia = 'sip'
        es_pjsip = tecnologia == 'pjsip'

        db_user, db_pass = obtener_credenciales_mysql(perfil)

        def esc(v):
            return str(v).replace("'", "''")

        idioma = str(idioma or '')
        if not _re.match(r'^[a-z]{2,3}(_[A-Za-z]{2})?$', idioma):
            idioma = ''

        def limpiar(txt):
            # Saca claves del texto antes de loguearlo o mostrarlo.
            txt = str(txt or '')
            secretos = set([db_pass] + [obtener_password(i) for i in internos])
            for secreto in secretos:
                if secreto:
                    txt = txt.replace(str(secreto), '***')
            lineas = [l for l in txt.splitlines() if 'Using a password' not in l]
            return ' | '.join(lineas)[:300]

        def ejecutar(cmd):
            if es_host_local(perfil['host']):
                return ejecutar_comando_local(cmd)
            return ejecutar_comando_ssh(perfil, cmd)

        def b64(txt):
            return base64.b64encode(txt.encode('utf-8')).decode('ascii')

        def mysql_consulta(sql):
            cmd = (
                "SQL_B64='{}'; "
                "SQL=\"$(printf '%s' \"$SQL_B64\" | base64 -d)\"; "
                "/usr/bin/mysql -N -B -u \"{}\" -p'{}' -e \"$SQL\" asterisk 2>/dev/null"
            ).format(b64(sql), db_user, db_pass)
            return ejecutar(cmd)

        def mysql_script(sql):
            cmd = (
                "SQL_B64='{}'; "
                "printf '%s' \"$SQL_B64\" | base64 -d | "
                "/usr/bin/mysql -A -u \"{}\" -p'{}' asterisk 2>&1"
            ).format(b64(sql), db_user, db_pass)
            return ejecutar(cmd)

        def ids_en(lista):
            return ",".join("'" + esc(x) + "'" for x in lista)

        # -----------------------------------------------------
        # Defensa: solo caracteres seguros en las claves
        # -----------------------------------------------------
        pw_ok = _re.compile(r'^[A-Za-z0-9_.@#%+=:-]{1,64}$')
        malos = [i for i in internos if not pw_ok.match(str(obtener_password(i)))]
        if malos:
            resultado['fallidos'] = list(internos)
            resultado['error'] = "Clave con caracteres no permitidos (interno {})".format(malos[0])
            return resultado

        # -----------------------------------------------------
        # Existentes: users, devices o sip
        # -----------------------------------------------------
        lista = ids_en(internos)
        consulta = (
            "SELECT extension FROM asterisk.users WHERE extension IN ({0}) "
            "UNION SELECT id FROM asterisk.devices WHERE id IN ({0}) "
            "UNION SELECT id FROM asterisk.sip WHERE keyword='account' AND id IN ({0});"
        ).format(lista)

        out, err, rc = mysql_consulta(consulta)
        if rc != 0:
            resultado['fallidos'] = list(internos)
            resultado['error'] = "No se pudo consultar los internos existentes en Issabel"
            return resultado

        existentes = set(l.strip() for l in (out or '').splitlines() if l.strip())
        resultado['existentes'] = [i for i in internos if i in existentes]
        nuevos = [i for i in internos if i not in existentes]

        logger.info("RANGO LOTE: %s solicitados / %s existentes / %s nuevos",
                    len(internos), len(resultado['existentes']), len(nuevos))

        if not nuevos:
            return resultado

        # -----------------------------------------------------
        # SQL de un interno (misma estructura que crear_peer_en_asterisk)
        # -----------------------------------------------------
        tech = 'pjsip' if es_pjsip else 'sip'

        def sql_interno(interno):
            password = obtener_password(interno)
            nm = interno
            ns = interno
            dial = ('PJSIP/' if es_pjsip else 'SIP/') + interno

            partes = []
            partes.append(
                "REPLACE INTO asterisk.users "
                "(extension, name, password, voicemail, ringtimer, noanswer, recording, "
                "outboundcid, sipname, mohclass) "
                "VALUES ('{}', '{}', '{}', 'novm', 0, '', '', '{}', '', NULL);".format(
                    esc(interno), esc(nm), esc(password), esc(ns)))
            partes.append(
                "REPLACE INTO asterisk.devices "
                "(id, tech, dial, devicetype, user, description, emergency_cid) "
                "VALUES ('{}', '{}', '{}', 'fixed', '{}', '{}', '');".format(
                    esc(interno), tech, esc(dial), esc(interno), esc(nm)))

            filas = [
                ('account', interno),
                ('callerid', 'device <{}>'.format(interno)),
                ('secret', password),
                ('type', 'friend'),
                ('host', 'dynamic'),
                ('context', contexto),
                ('dtmfmode', 'rfc2833'),
                ('nat', 'no'),
                ('qualify', 'yes'),
                ('qualifyfreq', '60'),
                ('trustrpid', 'yes'),
                ('sendrpid', 'no'),
                ('disallow', ''),
                ('allow', ''),
            ]
            if es_pjsip:
                filas.extend([
                    ('accountcode', ''),
                    ('allow_subscribe', 'yes'),
                    ('authenticate_qualify', 'no'),
                    ('direct_media', 'no'),
                    ('dtls_ca_file', ''),
                    ('dtls_cert_file', ''),
                    ('dtls_private_key', ''),
                    ('dtls_setup', 'actpass'),
                    ('dtls_verify', 'no'),
                    ('force_rport', 'yes'),
                    ('ice_support', 'no'),
                    ('mailbox', interno + '@device'),
                    ('max_contacts', '1'),
                    ('media_encryption', 'no'),
                    ('media_use_received_transport', 'no'),
                    ('message_context', ''),
                    ('permit', '0.0.0.0/0.0.0.0'),
                    ('rewrite_contact', 'yes'),
                    ('rtcp_mux', 'no'),
                    ('rtp_symmetric', 'yes'),
                    ('stir_shaken', 'off'),
                    ('subscribe_context', ''),
                    ('transport', 'transport-udp'),
                    ('trust_id_inbound', 'yes'),
                    ('use_avpf', 'no'),
                    # Campos de normalizacion PJSIP (copiados de la version anterior
                    # del lote: comparar con crear_peer_en_asterisk).
                    ('dial', 'PJSIP/{}'.format(interno)),
                    ('deny', '0.0.0.0/0.0.0.0'),
                    ('callgroup', ''),
                    ('pickupgroup', ''),
                    ('qualify_timeout', '3'),
                ])
            else:
                filas.append(('canreinvite', 'no'))

            valores = ", ".join(
                "('{}', '{}', '{}', 2)".format(esc(interno), esc(k), esc(v))
                for k, v in filas)
            partes.append(
                "REPLACE INTO asterisk.sip (id, keyword, data, flags) VALUES {};".format(valores))
            return "\n".join(partes)

        # -----------------------------------------------------
        # AstDB de un interno (mismos datos que poblar_astdb)
        # -----------------------------------------------------
        def astdb_interno(interno):
            password = obtener_password(interno)
            nm = interno
            dial = ('PJSIP/' if es_pjsip else 'SIP/') + interno
            u = "AMPUSER/{}".format(interno)
            rec = "AMPUSER/{}/recording".format(interno)
            dev = "DEVICE/{}".format(interno)
            entradas = [
                (u, "cidname", nm), (u, "cidnum", interno), (u, "device", interno),
                (u, "password", password), (u, "voicemail", "novm"), (u, "recording", "none"),
                (rec, "in/external", "dontcare"), (rec, "in/internal", "dontcare"),
                (rec, "out/external", "dontcare"), (rec, "out/internal", "dontcare"),
                (rec, "ondemand", "disabled"), (rec, "priority", "10"),
                (u, "ringtimer", "0"), (u, "cfringtimer", "0"),
                (u, "concurrency_limit", "0"), (u, "noanswer", "none"),
                (u, "outboundcid", interno),
                ("AMPUSER/{}/queues".format(interno), "qnostate", "usestate"),
                (u, "language", idioma), (u, "answermode", "disabled"),
                (dev, "dial", dial), (dev, "type", "fixed"),
                (dev, "user", interno), (dev, "default_user", interno),
            ]
            return ["/usr/sbin/asterisk -rx 'database put {} {} {}'".format(f, k, v)
                    for f, k, v in entradas if v != '']

        # -----------------------------------------------------
        # Proceso por tandas
        # -----------------------------------------------------
        total = len(nuevos)
        hechos_ok = []
        for ini in range(0, total, tanda_sql):
            tanda = nuevos[ini:ini + tanda_sql]
            resto = nuevos[ini + len(tanda):]

            if progreso:
                progreso(20 + int(65 * ini / total),
                         'Creando internos {}-{} de {}...'.format(ini + 1, ini + len(tanda), total))

            sql = "\n".join(
                ["SET autocommit=0;", "START TRANSACTION;"]
                + [sql_interno(i) for i in tanda]
                + ["COMMIT;"])
            out, err, rc = mysql_script(sql)
            if rc != 0:
                resultado['fallidos'].extend(tanda)
                resultado['no_procesados'].extend(resto)
                resultado['error'] = "Error SQL: {}".format(limpiar(out or err))
                logger.error("RANGO LOTE SQL: %s", resultado['error'])
                break

            corto = False
            for j in range(0, len(tanda), tanda_astdb):
                sub = tanda[j:j + tanda_astdb]
                cmds = []
                for i in sub:
                    cmds.extend(astdb_interno(i))
                out, err, rc = ejecutar("( " + " && ".join(cmds) + " ) 2>&1")
                if rc != 0 or 'Failed to update entry' in (out or ''):
                    pendientes = tanda[j:]
                    resultado['fallidos'].extend(pendientes)
                    resultado['no_procesados'].extend(resto)
                    resultado['error'] = (
                        "El SQL se aplico pero fallo AstDB desde el interno {}: {}"
                    ).format(sub[0], limpiar(out or err))
                    logger.error("RANGO LOTE ASTDB: %s", resultado['error'])
                    corto = True
                    break
                hechos_ok.extend(sub)
            if corto:
                break

        # -----------------------------------------------------
        # Verificacion final en SQL
        # -----------------------------------------------------
        if progreso:
            progreso(86, 'Verificando internos creados...')
        verificados = set()
        if hechos_ok:
            out, err, rc = mysql_consulta(
                "SELECT id FROM asterisk.sip WHERE keyword='account' AND id IN ({});".format(
                    ids_en(hechos_ok)))
            if rc == 0:
                verificados = set(l.strip() for l in (out or '').splitlines() if l.strip())

        resultado['creados'] = [i for i in hechos_ok if i in verificados]
        no_verif = [i for i in hechos_ok if i not in verificados]
        if no_verif:
            resultado['fallidos'].extend(no_verif)
            if not resultado['error']:
                resultado['error'] = "No se pudieron verificar en SQL: {}".format(", ".join(no_verif[:10]))

        logger.info("RANGO LOTE FINAL: creados=%s existentes=%s fallidos=%s no_procesados=%s",
                    len(resultado['creados']), len(resultado['existentes']),
                    len(resultado['fallidos']), len(resultado['no_procesados']))
        return resultado

    except Exception as e:
        logger.error("RANGO LOTE exception: %s", e, exc_info=True)
        resultado['error'] = "Excepcion en el lote: {}".format(str(e)[:150])
        hechos = set(resultado['creados']) | set(resultado['existentes']) | set(resultado['fallidos']) | set(resultado['no_procesados'])
        resultado['fallidos'].extend([i for i in internos if i not in hechos])
        return resultado



def eliminar_interno_asterisk(perfil, interno):
    try:
        db_user, db_pass = obtener_credenciales_mysql(perfil)

        tecnologia = str(
            perfil.get('tecnologia', 'sip')
        ).strip().lower()

        if tecnologia not in ('sip', 'pjsip'):
            tecnologia = 'sip'

        sql = (
            "DELETE FROM asterisk.sip "
            "WHERE id='" + interno + "'; "
        )

        sql += (
            "DELETE FROM asterisk.users "
            "WHERE extension='" + interno + "'; "
        )

        sql += (
            "DELETE FROM asterisk.devices "
            "WHERE id='" + interno + "';"
        )

        cmd = (
            "/usr/bin/mysql -u " + db_user +
            " -p'" + db_pass +
            "' -e \"" + sql + "\" asterisk 2>&1"
        )

        if es_host_local(perfil['host']):
            out, err, c = ejecutar_comando_local(cmd)
        else:
            out, err, c = ejecutar_comando_ssh(perfil, cmd)

        if c != 0:
            return False, "Error SQL: " + (err or out or "")[:100]

        if es_host_local(perfil['host']):

            ejecutar_comando_local(
                "/var/lib/asterisk/bin/retrieve_conf "
                "> /dev/null 2>&1"
            )

            ejecutar_comando_local(
                "/usr/sbin/asterisk -rx 'core reload' "
                "> /dev/null 2>&1"
            )

            if tecnologia == 'pjsip':
                ejecutar_comando_local(
                    "/usr/sbin/asterisk -rx 'pjsip reload' "
                    "> /dev/null 2>&1"
                )
            else:
                ejecutar_comando_local(
                    "/usr/sbin/asterisk -rx 'sip reload' "
                    "> /dev/null 2>&1"
                )

        else:

            if tecnologia == 'pjsip':
                reload_cmd = (
                    "/var/lib/asterisk/bin/retrieve_conf "
                    "> /dev/null 2>&1 && "
                    "asterisk -rx 'core reload' && "
                    "asterisk -rx 'pjsip reload'"
                )
            else:
                reload_cmd = (
                    "/var/lib/asterisk/bin/retrieve_conf "
                    "> /dev/null 2>&1 && "
                    "asterisk -rx 'core reload' && "
                    "asterisk -rx 'sip reload'"
                )

            ejecutar_comando_ssh(perfil, reload_cmd)

        # (parche) purga AstDB: poblar_astdb() crea AMPUSER/<ext> y DEVICE/<ext>,
        # y el telefono registrado deja contactos/registro. Si quedan, el peer
        # sigue apareciendo en Monitoreo y el index (con IP) tras eliminarlo.
        try:
            if re.match(r'^[A-Za-z0-9_.-]+$', str(interno)):
                cmd_purga = (
                    "A=/usr/sbin/asterisk; EXT='{ext}'; "
                    "$A -rx \"database deltree AMPUSER $EXT\" >/dev/null 2>&1; "
                    "$A -rx \"database deltree DEVICE $EXT\" >/dev/null 2>&1; "
                    "$A -rx \"database del SIP/Registry $EXT\" >/dev/null 2>&1; "
                    "$A -rx 'database show registrar/contact' 2>/dev/null "
                    "| grep -F \"/registrar/contact/${{EXT}};@\" "
                    "| awk '{{print $1}}' | sed 's/:$//' "
                    "| while read k; do "
                    "$A -rx \"database del registrar/contact ${{k#/registrar/contact/}}\" "
                    ">/dev/null 2>&1; done; true"
                ).format(ext=str(interno))

                if es_host_local(perfil['host']):
                    ejecutar_comando_local(cmd_purga)
                else:
                    ejecutar_comando_ssh(perfil, cmd_purga)
        except Exception as e:
            logger.warning("Purga AstDB del interno %s: %s", interno, e)

        return True, "Eliminado"

    except Exception as e:
        return False, str(e)


def asegurar_dispositivo_mysql(pf, interno):
    """Asegura que el interno tenga registro correcto en asterisk.devices (con devicetype y user)"""
    try:
        db_user, db_pass = obtener_credenciales_mysql(pf)

        # 1. Insertar en devices con devicetype='fixed' y user=interno (OBLIGATORIO en Issabel)
        cmd_device = ("/usr/bin/mysql -u " + db_user + " -p'" + db_pass + "' asterisk -e "
                      "\"REPLACE INTO asterisk.devices (id, tech, dial, devicetype, user, description, emergency_cid) "
                      "VALUES ('" + interno + "', 'sip', CONCAT('SIP/', '" + interno + "'), 'fixed', '" + interno + "', "
                      "'Interno " + interno + "', '');\" 2>&1")

        # 2. Asegurar ringtimer en users
        cmd_ring = ("/usr/bin/mysql -u " + db_user + " -p'" + db_pass + "' asterisk -e "
                    "\"UPDATE asterisk.users SET ringtimer = '15' WHERE extension = '" + interno + "';\" 2>&1")

        if es_host_local(pf['host']):
            out1, _, c1 = ejecutar_comando_local(cmd_device)
            out2, _, c2 = ejecutar_comando_local(cmd_ring)
        else:
            out1, _, c1 = ejecutar_comando_ssh(pf, cmd_device)
            out2, _, c2 = ejecutar_comando_ssh(pf, cmd_ring)

        if c1 != 0 or c2 != 0:
            logger.warning("asegurar_dispositivo_mysql: error SQL - devices:{} ({}) users:{} ({})".format(c1, out1[:150] if out1 else '', c2, out2[:150] if out2 else ''))
            return False, "Error SQL: " + (out1 or out2)[:200]

        return True, "OK"
    except Exception as e:
        logger.error("asegurar_dispositivo_mysql exception: " + str(e), exc_info=True)
        return False, str(e)

def limpiar_dispositivo_mysql(pf, interno):
    """Limpia el registro de la tabla devices al eliminar un interno"""
    try:
        db_user, db_pass = obtener_credenciales_mysql(pf)
        cmd = ("/usr/bin/mysql -u " + db_user + " -p'" + db_pass + "' asterisk -e "
               "\"DELETE FROM asterisk.devices WHERE id='" + interno + "';\" 2>&1")

        if es_host_local(pf['host']):
            out, _, c = ejecutar_comando_local(cmd)
        else:
            out, _, c = ejecutar_comando_ssh(pf, cmd)

        if c != 0:
            logger.warning("limpiar_dispositivo_mysql: error SQL - {}".format(out[:150] if out else ''))
            return False, "Error SQL: " + (out or '')[:200]

        return True, "OK"
    except Exception as e:
        logger.error("limpiar_dispositivo_mysql exception: " + str(e), exc_info=True)
        return False, str(e)

# VP_EDITAR_NOMBRES_SIP_PJSIP_V2

def actualizar_nombres_interno_asterisk(perfil, interno, nombre_mostrar='',
                                         nombre_saliente='',
                                         actualizar_mostrar=True,
                                         actualizar_saliente=True):
    """Actualiza solo los nombres de un interno SIP/PJSIP existente.

    No reconstruye el interno ni modifica secret, max_contacts, codecs,
    transport, grupos de captura u otros parámetros.
    """
    try:
        import base64
        import shlex
        import re

        interno = str(interno or '').strip()
        if not re.match(r'^[0-9]{1,10}$', interno):
            return False, 'Número de interno inválido.'

        actualizar_mostrar = bool(actualizar_mostrar)
        actualizar_saliente = bool(actualizar_saliente)
        if not actualizar_mostrar and not actualizar_saliente:
            return True, 'No hay nombres para actualizar.'

        nm = str(nombre_mostrar or '').strip() or interno
        ns = str(nombre_saliente or '').strip() or interno
        callerid = '"{}" <{}>'.format(nm, interno)
        outboundcid = '"{}" <{}>'.format(ns, interno)

        perfil_host = str(perfil.get('host', '') or '').strip()

        def ejecutar(cmd):
            if es_host_local(perfil_host):
                return ejecutar_comando_local(cmd)
            return ejecutar_comando_ssh(perfil, cmd)

        db_user, db_pass = obtener_credenciales_mysql(perfil)

        def sql_run(sql):
            sql_b64 = base64.b64encode(sql.encode('utf-8')).decode('ascii')
            cmd = (
                "SQL_B64={}; printf '%s' \"$SQL_B64\" | base64 -d | "
                "/usr/bin/mysql -N -B -u {} -p{} asterisk 2>&1"
            ).format(
                shlex.quote(sql_b64),
                shlex.quote(str(db_user)),
                shlex.quote(str(db_pass))
            )
            return ejecutar(cmd)

        def esc(value):
            return str(value).replace('\\', '\\\\').replace("'", "''")

        ext = esc(interno)
        out, err, rc = sql_run(
            "SELECT LOWER(COALESCE(d.tech,'')), "
            "IF(u.extension IS NULL,'0','1') "
            "FROM asterisk.devices d "
            "LEFT JOIN asterisk.users u ON u.extension=d.id "
            "WHERE d.id='{}' LIMIT 1;".format(ext)
        )
        if rc != 0:
            return False, 'No se pudo consultar el interno en Issabel: {}'.format((err or out or 'error SQL')[:300])
        lineas = [x.strip() for x in str(out or '').splitlines() if x.strip()]
        if not lineas:
            return False, 'El interno {} no existe en asterisk.devices de Issabel.'.format(interno)
        partes = lineas[-1].split('\t')
        tecnologia = partes[0].strip().lower() if partes else ''
        tiene_usuario = len(partes) > 1 and partes[1].strip() == '1'
        if tecnologia not in ('sip', 'pjsip'):
            return False, 'Issabel no identifica al interno {} como SIP o PJSIP (tecnología: {}).'.format(interno, tecnologia or 'desconocida')
        if not tiene_usuario:
            return False, 'El interno {} existe en devices pero falta en asterisk.users; no se modificó.'.format(interno)

        # Issabel genera pjsip_additional.conf desde asterisk.sip.
        # Editar nombres no debe cambiar trust_id_inbound ni el transporte.
        tabla_config = 'sip'

        statements = ['START TRANSACTION;']
        condiciones = ["u.extension='{}'".format(ext), "d.id='{}'".format(ext)]
        if actualizar_mostrar:
            nm_sql = esc(nm)
            callerid_sql = esc(callerid)
            statements.extend([
                "UPDATE asterisk.users SET name='{}' WHERE extension='{}';".format(nm_sql, ext),
                "UPDATE asterisk.devices SET description='{}' WHERE id='{}';".format(nm_sql, ext),
                "DELETE FROM asterisk.{} WHERE id='{}' AND keyword='callerid';".format(tabla_config, ext),
                "INSERT INTO asterisk.{} (id,keyword,data,flags) VALUES ('{}','callerid','{}',0);".format(tabla_config, ext, callerid_sql)
            ])
            condiciones.extend([
                "u.name='{}'".format(nm_sql),
                "d.description='{}'".format(nm_sql),
                "(SELECT COUNT(*) FROM asterisk.{} s WHERE s.id='{}' AND s.keyword='callerid' AND s.data='{}')=1".format(tabla_config, ext, callerid_sql)
            ])
        if actualizar_saliente:
            cid_sql = esc(outboundcid)
            statements.append(
                "UPDATE asterisk.users SET outboundcid='{}' WHERE extension='{}';".format(cid_sql, ext)
            )
            condiciones.append("u.outboundcid='{}'".format(cid_sql))
        statements.append('COMMIT;')

        out, err, rc = sql_run('\n'.join(statements))
        if rc != 0:
            return False, 'Falló la actualización de nombres en la base de Issabel: {}'.format((err or out or 'error SQL')[:350])

        ver_sql = (
            "SELECT COUNT(*) FROM asterisk.users u "
            "JOIN asterisk.devices d ON d.id=u.extension WHERE " +
            ' AND '.join(condiciones) + ';'
        )
        v_out, v_err, v_rc = sql_run(ver_sql)
        if v_rc != 0 or not [x for x in str(v_out or '').splitlines() if x.strip()] or str(v_out).strip().splitlines()[-1].strip() != '1':
            return False, 'Issabel no confirmó todos los nombres guardados. Revisá el interno antes de volver a editarlo.'

        def cli_quoted(value):
            value = str(value).replace('\\', '\\\\').replace('"', '\\"').replace('\r', ' ').replace('\n', ' ')
            return '"{}"'.format(value)

        def astdb_put(key, value):
            cli = 'database put AMPUSER/{} {} {}'.format(interno, key, cli_quoted(value))
            cmd = '/usr/sbin/asterisk -rx {}'.format(shlex.quote(cli))
            a_out, a_err, a_rc = ejecutar(cmd)
            combined = '{}\n{}'.format(a_out or '', a_err or '').lower()
            if a_rc != 0 or 'no such command' in combined or 'usage:' in combined or 'database put failed' in combined:
                raise RuntimeError('No se pudo actualizar AstDB ({})'.format(key))

        comandos = [
            '/var/lib/asterisk/bin/retrieve_conf',
            "/usr/sbin/asterisk -rx 'dialplan reload'",
            "/usr/sbin/asterisk -rx 'sip reload'"
        ]
        if tecnologia == 'pjsip':
            comandos.append("/usr/sbin/asterisk -rx 'pjsip reload'")

        for cmd in comandos:
            rr_out, rr_err, rr_rc = ejecutar(cmd)
            if rr_rc != 0:
                return False, 'Los datos de nombres se guardaron, pero falló la recarga de Issabel ({}): {}'.format(
                    cmd.split()[0], (rr_err or rr_out or 'sin detalle')[:250])

        # V1.2: establecer AstDB al final; retrieve_conf puede regenerar sus valores.
        # Verificar el valor real leído desde Asterisk para no informar éxito falso.
        # VP_EDITAR_NOMBRES_ASTDB_FINAL_VERIFY_V12
        def astdb_get(key):
            cli = 'database get AMPUSER/{} {}'.format(interno, key)
            cmd = '/usr/sbin/asterisk -rx {}'.format(shlex.quote(cli))
            g_out, g_err, g_rc = ejecutar(cmd)
            combinado = '{}\n{}'.format(g_out or '', g_err or '')
            if g_rc != 0:
                raise RuntimeError('No se pudo verificar AstDB ({})'.format(key))
            import re
            m = re.search(r'(?m)^\s*Value:\s?(.*?)\s*$', combinado)
            if not m:
                raise RuntimeError('Asterisk no devolvió el valor de AstDB ({})'.format(key))
            return m.group(1)

        if actualizar_mostrar:
            astdb_put('cidname', nm)
            comprobado = astdb_get('cidname')
            if comprobado != nm:
                raise RuntimeError('AstDB cidname no coincide. Esperado {!r}; leído {!r}.'.format(nm, comprobado))
        if actualizar_saliente:
            astdb_put('outboundcid', outboundcid)
            comprobado = astdb_get('outboundcid')
            if comprobado != outboundcid:
                raise RuntimeError('AstDB outboundcid no coincide. Esperado {!r}; leído {!r}.'.format(outboundcid, comprobado))

        if tecnologia == 'pjsip' and actualizar_mostrar:
            cli = 'pjsip show endpoint {}'.format(interno)
            cmd = '/usr/sbin/asterisk -rx {}'.format(shlex.quote(cli))
            p_out, p_err, p_rc = ejecutar(cmd)
            p_text = '{}\n{}'.format(p_out or '', p_err or '')
            if p_rc != 0:
                raise RuntimeError('No se pudo verificar el endpoint PJSIP.')
            m_cid = re.search(r'(?im)^\s*callerid\s*:\s*(.*?)\s*$', p_text)
            if not m_cid or nm.lower() not in m_cid.group(1).lower():
                actual = m_cid.group(1) if m_cid else 'callerid no informado'
                raise RuntimeError(
                    'PJSIP sigue informando callerid={!r}; se esperaba {!r}.'.format(actual, nm)
                )

        return True, 'Nombres actualizados y verificados en Issabel ({}) y AstDB.'.format(tecnologia.upper())
    except Exception as e:
        try:
            logger.exception('Error actualizando nombres del interno %s en Issabel: %s', interno, e)
        except Exception:
            pass
        return False, 'Error actualizando nombres en Issabel: {}'.format(str(e)[:350])

