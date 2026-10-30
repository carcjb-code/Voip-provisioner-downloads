# Borrador de publicación — Windows

Este archivo es la guía de carga para la primera publicación pública. No es un instalador.

## Datos de la publicación

- **Título:** VoIP Provisioner para Windows — compilación 112 (prueba)
- **Etiqueta (tag):** `windows-preview-2026-10-29`
- **Destino:** rama `main`
- **Marcar:** Pre-release
- **No marcar:** Latest release

## Archivos que deben adjuntarse

Descargar ambos ZIP desde la publicación privada de origen y adjuntarlos aquí con estos nombres exactos:

- `VoIPProvisioner-Windows-32-bit.zip` (x86)
- `VoIPProvisioner-Windows-64-bit.zip` (x64)

Publicación privada de origen: https://github.com/carcjb-code/Voip-provisioner/releases/tag/windows-android-build-112-1

## Texto sugerido para las notas

Compilación 112 de VoIP Provisioner para Windows, distribuida en paquetes portables separados para 32 y 64 bits.

**Versión de prueba:** validar apertura, navegación entre las cuatro secciones, conexión SSH, sincronización, lectura y escritura de plantillas, y persistencia de los perfiles guardados antes de utilizarla en producción.

Los ZIP son paquetes portables; extraer cada paquete en una carpeta propia y ejecutar el programa incluido. No mezclar archivos de las versiones x86 y x64.

Antes de reemplazar una instalación anterior, realizar una exportación de los perfiles de servidor y verificar que se puedan importar en la nueva versión.

Esta publicación no debe considerarse una versión estable.
