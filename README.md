🏗️ ACERO_SCRIPT
Automatización del cálculo y colocación de aceros en planos de prelosas mediante Python y ezdxf
📋 Descripción
ACERO_SCRIPT es una herramienta especializada para automatizar el cálculo y la colocación de especificaciones de acero en planos de prelosas de concreto. Desarrollada para ingenieros civiles y estructurales, esta aplicación procesa archivos DXF y extrae información sobre aceros longitudinales y transversales, calculando automáticamente los valores adecuados según normas estructurales.
🚀 Características

✅ Procesamiento automático de planos DXF de prelosas

✅ Detección inteligente de textos de especificación de aceros

✅ Cálculo de valores estructurales mediante integración con Excel

✅ Soporte para múltiples tipos de prelosas (macizas, aligeradas, bidireccionales)

✅ Interfaz gráfica intuitiva desarrollada con tkinter

✅ Generación automática de bloques con atributos de acero en planos

✅ Configuración personalizable de valores predeterminados

🛠️ Componentes Principales

main.py - Script principal para iniciar la aplicación
gui.py - Interfaz gráfica desarrollada con tkinter
dxf_processor.py - Funciones para procesar archivos DXF
acero_calculator.py - Módulo de cálculo de especificaciones de acero
excel_integration.py - Integración con Excel para cálculos estructurales
templates/ - Directorio con plantillas de bloques DXF
README.md - Documentación del proyecto

🔧 Requisitos

Python 3.7 o superior
Bibliotecas requeridas:
pip install ezdxf pandas openpyxl xlwings tkinter pyinstaller shapely

Microsoft Excel (para la integración de cálculos)
Archivos DXF generados con AutoCAD 2018 o superior

📦 Instalación

Clona este repositorio:
git clone https://github.com/yourusername/ACERO_SCRIPT.git

Navega al directorio del proyecto:
cd ACERO_SCRIPT

Instala las dependencias:
pip install -r requirements.txt

Ejecuta la aplicación:
python main.py


🚦 Uso

Inicia la aplicación ejecutando main.py
En la interfaz gráfica, configura los tipos de prelosas y sus valores predeterminados
Selecciona el archivo DXF que contiene las prelosas a procesar
Elige la carpeta de destino para guardar el archivo procesado
Haz clic en "PROCESAR PRELOSAS" para iniciar el procesamiento
Revisa los resultados en la consola de registro y en el archivo DXF generado

📊 Proceso de Cálculo
El proceso de cálculo sigue estos pasos:

Análisis del archivo DXF para identificar polilíneas que representan prelosas
Detección de textos dentro de las polilíneas que especifican aceros
Procesamiento de textos para extraer diámetros y espaciamientos
Integración con Excel para realizar cálculos estructurales
Generación de bloques con las especificaciones de acero calculadas
Inserción de los bloques en las ubicaciones correctas dentro del plano

💡 Ejemplos de Casos de Uso
Procesamiento de Prelosas Aligeradas
pythonfrom acero_calculator import procesar_prelosa_aligerada

# Ejemplo de procesamiento de una prelosa aligerada
resultado = procesar_prelosa_aligerada(
    textos_longitudinal=["1Ø3/8\"@.20"],
    textos_transversal=["1Ø6mm@.50"],
    tipo_prelosa="PRELOSA ALIGERADA 20",
    default_valores={"espaciamiento": 0.20, "acero": "3/8\""}
)

# Insertar bloque con los resultados
insertar_bloque_acero(
    msp=modelspace,
    definicion_bloque=bloque_info,
    centro=(100, 100),
    as_long=resultado["as_long"],
    as_tra1=resultado["as_tra1"],
    as_tra2=resultado["as_tra2"]
)
Detección de Textos en Polilíneas
pythonfrom dxf_processor import obtener_textos_dentro_de_polilinea

# Detectar especificaciones de acero dentro de una polilínea
textos = obtener_textos_dentro_de_polilinea(
    polilinea=vertices,
    textos=todos_los_textos,
    capa_polilinea="ACERO LONGITUDINAL"
)
🤝 Contribuciones
Las contribuciones son bienvenidas. Por favor, siente libre de:

Hacer fork del proyecto
Crear una rama para tu característica (git checkout -b feature/MejorDeteccionTextos)
Hacer commit de tus cambios (git commit -m 'Mejorar detección de textos en planos')
Push a la rama (git push origin feature/MejorDeteccionTextos)
Abrir un Pull Request

📄 Licencia
Este proyecto está licenciado bajo la Licencia MIT - ver el archivo LICENSE para más detalles.
📞 Contacto
EderJair - @GitHub
Link del proyecto: https://github.com/EderJair/AutocadAUT

⌨️ con ❤️ por EderJair
