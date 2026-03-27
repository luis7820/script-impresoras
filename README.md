Zebra Config (ZBRN 2)

Zebra Config es una herramienta avanzada de automatización desarrollada en PowerShell diseñada para gestionar, configurar e instalar impresoras de etiquetas Zebra de forma rápida y visual.

Características Principales

Interfaz Gráfica Moderna: Utiliza un componente WebBrowser para mostrar una interfaz basada en HTML/CSS, ofreciendo una experiencia de usuario superior a la consola estándar.

Instalación de Perfiles: Configuración automática de drivers para diferentes formatos de etiquetas (Nissan Corto, Nissan Largo, MQ, etc.).

Gestor de Spooler: Botones directos para detener e iniciar el servicio de cola de impresión (Spooler) de Windows.

Limpieza Automática: Función integrada para purgar archivos temporales de la cola de impresión que suelen bloquear las impresoras.

Selección de DPI: Permite alternar fácilmente entre resoluciones de impresión (203 DPI / 300 DPI) para asegurar la calidad de la etiqueta.

Ejecución Segura: El script se eleva automáticamente a privilegios de Administrador para poder gestionar servicios y drivers del sistema.

Requisitos

Sistema Operativo: Windows 10 o Windows 11.

Permisos: El usuario debe tener capacidad de ejecución de scripts de PowerShell (ExecutionPolicy Bypass).

Drivers: Requiere que los archivos de drivers de Zebra estén ubicados en la ruta predefinida en el script (C:\Drivers\Zebra).

Instrucciones de Uso

Descarga el script: Asegúrate de tener el archivo ZBRN 2.ps1 en tu equipo.

Ejecución:

Haz clic derecho sobre el archivo y selecciona "Ejecutar con PowerShell".

El script solicitará permisos de administrador; acéptalos para continuar.

Interfaz:

Selecciona el DPI correspondiente a tu impresora en la parte superior.

Utiliza los botones verdes (Nissan, MQ, etc.) para instalar la configuración deseada.

Si la impresora se bloquea, utiliza "Stop Spooler", luego "Limpiar Cola" y finalmente "Start Spooler".

Estructura Técnica

El script utiliza un puente COM (ZebraBridge) en C# para permitir que la interfaz HTML se comunique directamente con las funciones de PowerShell:

SetDPI(string dpi): Cambia la resolución activa.

Instalar(string opcion): Dispara la lógica de instalación de drivers mediante comandos rundll32 printui.dll.

StopSpooler() / StartSpooler(): Gestiona el servicio del sistema mediante Stop-Service y Start-Service.

Notas Importantes

Rutas de Archivos: Verifica que los archivos .dat y drivers mencionados en el script existan en C:.

Compatibilidad: Diseñado específicamente para modelos Zebra que utilicen el driver de Windows estándar.

Desarrollado para optimizar la logística y el etiquetado industrial.
