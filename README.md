# VoIP Provisioner — Descargas oficiales

Este repositorio contiene exclusivamente los instaladores públicos de VoIP Provisioner. El código fuente y la configuración interna permanecen en un repositorio privado.

## Descargas

Las versiones se publican como **versiones de prueba (pre-release)**. Los enlaces permanecen estables y apuntan siempre a la publicación pública más reciente.

| Plataforma | Descarga |
|---|---|
| Android 32 bits (ARM) | [Descargar APK](https://github.com/carcjb-code/Voip-provisioner-downloads/releases/download/preview-current/VoIPProvisioner-Android-32bit.apk) |
| Android 64 bits (ARM) | [Descargar APK](https://github.com/carcjb-code/Voip-provisioner-downloads/releases/download/preview-current/VoIPProvisioner-Android-64bit.apk) |
| Windows 32 bits — Instalador | [Descargar Setup EXE](https://github.com/carcjb-code/Voip-provisioner-downloads/releases/download/preview-current/VoIPProvisioner-Setup-Windows-32-bit.exe) |
| Windows 64 bits — Instalador | [Descargar Setup EXE](https://github.com/carcjb-code/Voip-provisioner-downloads/releases/download/preview-current/VoIPProvisioner-Setup-Windows-64-bit.exe) |
| Windows 32 bits — Portable | [Descargar ZIP](https://github.com/carcjb-code/Voip-provisioner-downloads/releases/download/preview-current/VoIPProvisioner-Windows-32-bit.zip) |
| Windows 64 bits — Portable | [Descargar ZIP](https://github.com/carcjb-code/Voip-provisioner-downloads/releases/download/preview-current/VoIPProvisioner-Windows-64-bit.zip) |
| Verificación de integridad | [SHA256SUMS.txt](https://github.com/carcjb-code/Voip-provisioner-downloads/releases/download/preview-current/SHA256SUMS.txt) |

## Instalación Windows

1. Descargá el instalador que corresponda a tu arquitectura (32 o 64 bits).
2. Ejecutalo y leé/aceptá los Términos y Condiciones.
3. La instalación conserva los datos de usuario bajo `%APPDATA%\VoIP_Provisioner`, incluso al actualizar o desinstalar.
4. El programa usa una interfaz local; no abre el puerto 8080 como servicio de red.

Los paquetes ZIP son versiones portables: extraé su contenido y ejecutá `VoIPProvisioner.exe`.

## Actualizaciones y advertencias

- Las descargas son públicas; el código fuente y la configuración interna no se publican.
- En las compilaciones Android, el actualizador incremental de Mantenimiento puede depender de servicios del repositorio privado; esa función podría no estar disponible para usuarios externos.
- No utilices compilaciones de prueba en centrales de producción sin validarlas primero en un entorno de laboratorio.
- Antes de reinstalar o cambiar una versión, exportá los perfiles de servidor y comprobá que puedan restaurarse.

## Publicaciones

[Ver versiones, notas y archivos](https://github.com/carcjb-code/Voip-provisioner-downloads/releases).
