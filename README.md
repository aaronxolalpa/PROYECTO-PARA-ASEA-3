# PROYECTO-PARA-ASEA-3
Renombrador y Actualizador de Oficios Escaneados desde PDF
Este proyecto en Python realizado en mi servicio social en ASEA (Agencia de Seguridad Energia y Medio Ambiente) proporciona una herramienta gráfica para procesar archivos PDF que contienen imágenes de guías con códigos de barras. Permite extraer el número de guía desde los códigos escaneados, obtener automáticamente el número de oficio desde una base de datos Access, renombrar los archivos, y actualizar las fechas de notificación según el tipo de entrega (CC o DEV).

Funcionalidades
Carga masiva de archivos PDF desde una carpeta seleccionada.

Extracción automática de imágenes de cada página del PDF.

Lectura de códigos de barras con Pyzbar (basado en ZBar).

Búsqueda del número de oficio relacionado en una base de datos Access mediante el número de guía.

Renombrado automático del archivo PDF con formato NUMERO_OFICIO_tipo.pdf.

Opción para seleccionar tipo de entrega (CC o DEV).

Actualización de la fecha de notificación en la base de datos para los casos tipo CC.

Visualización del PDF con el visor por defecto del sistema.

Cierre automático del visor al avanzar al siguiente archivo.

Interfaz gráfica con progreso, contador y campos de edición.

Gestión y reporte de archivos no-PDF (omitidos).

Requisitos
Python 3.8+

Microsoft Access instalado (o controlador ODBC compatible)
