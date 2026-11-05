# app/excel_utils.py - Utilidades de manejo de Excel
import os
import logging
from openpyxl import Workbook, load_workbook
from app.config import COLUMNAS_EXCEL, get_excel_file, cargar_perfiles, buscar_perfil
from app.ssh_client import conectar_sftp

logger = logging.getLogger(__name__)


def asegurar_encabezados_excel(excel_path):
    if not os.path.exists(excel_path):
        wb = Workbook()
        ws = wb.active
        ws.title = "Provisionamiento"
        ws.append(COLUMNAS_EXCEL)
        wb.save(excel_path)
        wb.close()
    else:
        wb = load_workbook(excel_path)
        ws = wb.active
        actual = [cell.value for cell in ws[1]]
        if any(v is None for v in actual) or [str(v).strip() for v in actual] != COLUMNAS_EXCEL:
            ws.insert_rows(1)
            for i, titulo in enumerate(COLUMNAS_EXCEL, start=1):
                ws.cell(row=1, column=i, value=titulo)
            wb.save(excel_path)
        wb.close()


def verificar_duplicados(excel_path, interno, mac):
    if not os.path.exists(excel_path):
        return False, ""
    wb = load_workbook(excel_path)
    ws = wb.active
    for row in ws.iter_rows(min_row=2, values_only=True):
        if len(row) >= 4:
            if str(row[3]) == str(interno):
                wb.close()
                return True, f"Interno '{interno}' ya existe"
            if mac and str(row[4]) == str(mac):
                wb.close()
                return True, f"MAC '{mac}' ya existe"
    wb.close()
    return False, ""


def actualizar_excel(perfil_nombre, edificio, piso, dependencia, interno, mac, ip=''):
    excel_path = get_excel_file(perfil_nombre)
    asegurar_encabezados_excel(excel_path)
    wb = load_workbook(excel_path)
    ws = wb.active
    ws.append([edificio, piso, dependencia, interno, mac, ip])
    wb.save(excel_path)
    wb.close()


def subir_excel_remoto(perfil_nombre, ruta_remota):
    if not perfil_nombre:
        return
    try:
        perfiles = cargar_perfiles()
        perfil = buscar_perfil(perfil_nombre, perfiles)
        if not perfil:
            logger.warning(f"Perfil {perfil_nombre} no encontrado")
            return
        host = perfil['host']
        puerto = int(perfil.get('puerto', 22))
        nombre_remoto = f"{host}.xlsx"
        destino = f"{ruta_remota}/{nombre_remoto}"
        sftp, _ = conectar_sftp(
            host, puerto, perfil['usuario'], perfil['password'],
            ruta_remota, perfil.get('ssh_compat', False)
        )
        excel_local = get_excel_file(perfil_nombre)
        with sftp:
            sftp.put(excel_local, destino)
        logger.info(f"Excel subido a {destino}")
    except Exception as e:
        logger.error(f"No se pudo subir el Excel: {e}")