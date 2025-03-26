import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import os
import sys
import traceback
import time
import subprocess
import random
import re
import platform
import importlib.util

def import_module_from_path():
    """
    Importar script.py dinámicamente
    """
    try:
        # Primero, intenta encontrar el script en la ubicación del ejecutable
        if getattr(sys, 'frozen', False):
            # Si es un ejecutable compilado
            script_path = os.path.join(sys._MEIPASS, 'script.py')
        else:
            # Si se ejecuta como script normal
            script_path = os.path.join(os.path.dirname(__file__), 'script.py')
        
        # Verificar si el archivo existe
        if not os.path.exists(script_path):
            print(f"Error: No se encontró script.py en {script_path}")
            return None

        # Importar el módulo
        spec = importlib.util.spec_from_file_location("script", script_path)
        script = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(script)
        return script
    except Exception as e:
        print(f"Error importando script: {e}")
        traceback.print_exc()
        input("Presione Enter para continuar...")
        return None

def generar_numero_random(digitos=2):
    """
    Genera un número aleatorio con el número de dígitos especificado.
    """
    min_valor = 10 ** (digitos - 1)
    max_valor = (10 ** digitos) - 1
    return random.randint(min_valor, max_valor)

def imprimir_banner_script():
    """
    Imprime un banner ASCII decorativo e impresionante con el texto 'ACERO SCRIPT'
    """
    banner = """
╔════════════════════════════════════════════════════════════════════════════════════════════════╗
║                                                                                                ║
║     █████╗  ██████╗███████╗██████╗  ██████╗     ███████╗ ██████╗██████╗ ██╗██████╗████████╗    ║
║    ██╔══██╗██╔════╝██╔════╝██╔══██╗██╔═══██╗    ██╔════╝██╔════╝██╔══██╗██║██╔══██╗╚══██╔══╝   ║
║    ███████║██║     █████╗  ██████╔╝██║   ██║    ███████╗██║     ██████╔╝██║██████╔╝   ██║      ║
║    ██╔══██║██║     ██╔══╝  ██╔══██╗██║   ██║    ╚════██║██║     ██╔══██╗██║██╔═══╝    ██║      ║
║    ██║  ██║╚██████╗███████╗██║  ██║╚██████╔╝    ███████║╚██████╗██║  ██║██║██║        ██║      ║
║    ╚═╝  ╚═╝ ╚═════╝╚══════╝╚═╝  ╚═╝ ╚═════╝     ╚══════╝ ╚═════╝╚═╝  ╚═╝╚═╝╚═╝        ╚═╝      ║
║                                                                                                ║
║                    Herramienta para Automatización de Aceros en Prelosas                       ║
║                                     by DODOD SOLUTIONS                                         ║
║                                       v1.0.0 (2025)                                            ║
║                                                                                                ║
╚════════════════════════════════════════════════════════════════════════════════════════════════╝
    """
    print(banner)

def proceso_dxf(dxf_path, excel_path, output_path, valores_predeterminados=None):
    """
    Función para ejecutar el procesamiento de DXF con valores personalizados.
    """
    try:
        # Intentar abrir consola para Windows
        if platform.system() == "Windows":
            import ctypes
            ctypes.windll.kernel32.AllocConsole()
        
        # Redirigir stdout y stderr
        sys.stdout = open('CONOUT$', 'w') if platform.system() == "Windows" else sys.stdout
        sys.stderr = open('CONOUT$', 'w') if platform.system() == "Windows" else sys.stderr

        print("Iniciando procesamiento de DXF...")
        print(imprimir_banner_script())
        print("=" * 80)
        print(f"Archivo de entrada: {dxf_path}")
        print(f"Archivo Excel: {excel_path}")
        print(f"Archivo de salida: {output_path}")

        # Importar script dinámicamente
        script_module = import_module_from_path()
        
        if script_module is None:
            print("No se pudo importar el script")
            time.sleep(0.3)
            return
        print("=" * 80)
        print("Valores predeterminados recibidos:")
        print(valores_predeterminados)
        print("=" * 80)
        
        # Ejecutar procesamiento con valores predeterminados
        total = script_module.procesar_prelosas_con_bloques(
            dxf_path, 
            excel_path,
            output_path,
            valores_predeterminados
        )
        print(f"Procesamiento completado. Bloques insertados: {total}")
        
        # Abrir exactamente la carpeta donde se guardó el archivo
        output_folder = os.path.dirname(output_path)
        try:
            if platform.system() == "Windows":
                print(f"Abriendo carpeta: {output_folder}")
                subprocess.Popen(f'explorer "{output_folder}"')
        except Exception as e:
            print(f"No se pudo abrir el directorio de salida: {e}")
        
        # Esperar en lugar de usar input()
        print("El proceso ha finalizado. Esta ventana se cerrará en 5 segundos...")
        time.sleep(0.3)
        
    except Exception as e:
        print(f"Error de procesamiento: {e}")
        print(traceback.format_exc())
        print("La aplicación se cerrará en 10 segundos...")
        time.sleep(0.3)
    finally:
        # Restaurar streams de sistema
        sys.stdout = sys.__stdout__
        sys.stderr = sys.__stderr__

class DXFProcessorApp:
    def __init__(self, master):
        self.master = master
        master.title("Procesador de Prelosas - DODOD SOLUTIONS")
        master.geometry("900x950")
        
        # Configurar estilo personalizado
        self.configurar_estilo()
        
        # Directorio del script
        self.script_dir = os.path.dirname(os.path.abspath(__file__))
        
        # Ruta fija del archivo Excel
        self.excel_path = os.path.join(self.script_dir, "CONVERTIDOR.xlsx")
        
        # Crear contenedor principal con efecto de elevación
        self.main_container = tk.Frame(master, bg="#f5f5f5", padx=15, pady=15)
        self.main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Cabecera con banner/logo
        self.crear_cabecera()
        
        # Crear los diferentes paneles
        self.crear_panel_archivos()
        self.crear_panel_configuracion()
        self.crear_panel_accion()
        self.crear_panel_registro()
        
        # Información del pie de página
        self.crear_pie_pagina()
    
    def configurar_estilo(self):
        """Configurar estilos personalizados para la interfaz"""
        self.master.configure(bg="#e0e0e0")  # Fondo gris claro
        
        # Configurar estilo para ttk
        style = ttk.Style()
        style.theme_use('clam')  # Usar tema clam que es más personalizable
        
        # Colores principales
        primary_color = "#1976D2"  # Azul
        secondary_color = "#388E3C"  # Verde
        bg_color = "#f5f5f5"  # Gris muy claro
        
        # Configurar estilos de widgets
        style.configure("TFrame", background=bg_color)
        style.configure("TLabel", background=bg_color, font=('Segoe UI', 10))
        style.configure("TButton", font=('Segoe UI', 10, 'bold'))
        style.configure("Header.TLabel", font=('Segoe UI', 14, 'bold'), foreground=primary_color)
        style.configure("Subheader.TLabel", font=('Segoe UI', 12, 'bold'), foreground=primary_color)
        
        # Botón principal (azul)
        style.configure("Primary.TButton", 
                       background=primary_color, 
                       foreground="white",
                       font=('Segoe UI', 12, 'bold'),
                       padding=10)
        style.map("Primary.TButton",
                 background=[('active', '#1565C0'), ('pressed', '#0D47A1')])
        
        # Botón secundario (verde)
        style.configure("Secondary.TButton", 
                       background=secondary_color, 
                       foreground="white")
        style.map("Secondary.TButton",
                 background=[('active', '#2E7D32'), ('pressed', '#1B5E20')])
        
        # Marco con borde
        style.configure("Card.TFrame", 
                       background=bg_color,
                       relief="raised",
                       borderwidth=1)
        
        # Marco de configuración
        style.configure("Config.TLabelframe", 
                       background=bg_color,
                       font=('Segoe UI', 11, 'bold'),
                       foreground=primary_color)
        style.configure("Config.TLabelframe.Label", 
                       background=bg_color,
                       font=('Segoe UI', 11, 'bold'),
                       foreground=primary_color)
    
    def crear_cabecera(self):
        """Crear sección de cabecera con logo o título"""
        header_frame = ttk.Frame(self.main_container, style="Card.TFrame")
        header_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Título principal
        title_text = "ACERO SCRIPT - Automatización de Prelosas"
        title_label = ttk.Label(
            header_frame, 
            text=title_text,
            style="Header.TLabel",
            padding=(20, 15)
        )
        title_label.pack()
        
        # Subtítulo
        subtitle_text = "Herramienta para procesamiento automático de aceros "
        subtitle_label = ttk.Label(
            header_frame,
            text=subtitle_text,
            style="Subheader.TLabel",
            padding=(0, 0, 0, 15)
        )
        subtitle_label.pack()
        
        # Línea separadora
        separator = ttk.Separator(header_frame, orient='horizontal')
        separator.pack(fill=tk.X, padx=20)
        
        # Información de la empresa
        company_label = ttk.Label(
            header_frame,
            text="Desarrollado por DODOD SOLUTIONS © 2025",
            padding=(0, 15, 0, 5)
        )
        company_label.pack()
    
    def crear_panel_archivos(self):
        """Crear panel de selección de archivos"""
        file_panel = ttk.LabelFrame(
            self.main_container, 
            text="Selección de Archivos",
            style="Config.TLabelframe",
            padding=15
        )
        file_panel.pack(fill=tk.X, pady=(0, 15))
        
        # Grilla para los campos
        file_frame = ttk.Frame(file_panel)
        file_frame.pack(fill=tk.X)
        
        # DXF File
        ttk.Label(file_frame, text="Archivo DXF:").grid(row=0, column=0, sticky='w', padx=(0,10), pady=8)
        self.dxf_path = tk.StringVar()
        self.dxf_entry = ttk.Entry(file_frame, textvariable=self.dxf_path, width=60)
        self.dxf_entry.grid(row=0, column=1, padx=5, pady=8, sticky='ew')
        ttk.Button(file_frame, text="Buscar", command=self.select_dxf_file).grid(row=0, column=2, padx=5, pady=8)
        
        # Excel File (Fixed)
        ttk.Label(file_frame, text="Archivo Excel:").grid(row=1, column=0, sticky='w', padx=(0,10), pady=8)
        self.excel_entry = ttk.Entry(file_frame, width=60, state='readonly')
        self.excel_entry.grid(row=1, column=1, padx=5, pady=8, sticky='ew')
        self.excel_entry.insert(0, "CONVERTIDOR.xlsx (Predeterminado)")
        
        # Output Directory
        ttk.Label(file_frame, text="Directorio de Salida:").grid(row=2, column=0, sticky='w', padx=(0,10), pady=8)
        self.output_path = tk.StringVar()
        self.output_entry = ttk.Entry(file_frame, textvariable=self.output_path, width=60)
        self.output_entry.grid(row=2, column=1, padx=5, pady=8, sticky='ew')
        ttk.Button(file_frame, text="Seleccionar", command=self.select_output_directory).grid(row=2, column=2, padx=5, pady=8)
        
        # Configurar columna expansible
        file_frame.columnconfigure(1, weight=1)
    
    def crear_panel_configuracion(self):
        """Crear panel de configuración de valores predeterminados"""
        config_panel = ttk.LabelFrame(
            self.main_container, 
            text="Configuración de Valores Predeterminados",
            style="Config.TLabelframe",
            padding=15
        )
        config_panel.pack(fill=tk.X, pady=(0, 15))
        
        # Variables para almacenar los valores predeterminados
        self.default_values = {
            'PRELOSA MACIZA': {
                'espaciamiento': tk.StringVar(value='0.20')
            },
            'PRELOSA ALIGERADA 20': {
                'espaciamiento': tk.StringVar(value='0.605')
            },
            'PRELOSA ALIGERADA 20 - 2 SENT': {
                'espaciamiento': tk.StringVar(value='0.605')
            }
        }
        
        # Crear una tabla para los valores predeterminados
        config_frame = ttk.Frame(config_panel)
        config_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Cabeceras de tabla
        ttk.Label(config_frame, text="Tipo de Prelosa", font=('Segoe UI', 10, 'bold')).grid(
            row=0, column=0, sticky='w', padx=10, pady=(0, 10))
        ttk.Label(config_frame, text="Espaciamiento (m)", font=('Segoe UI', 10, 'bold')).grid(
            row=0, column=1, sticky='w', padx=10, pady=(0, 10))
        
        # Separador horizontal
        separator = ttk.Separator(config_frame, orient='horizontal')
        separator.grid(row=1, column=0, columnspan=3, sticky='ew', pady=(0, 10))
        
        # Tipos de prelosa
        tipos_prelosa = [
            'PRELOSA MACIZA', 
            'PRELOSA ALIGERADA 20', 
            'PRELOSA ALIGERADA 20 - 2 SENT'
        ]
        
        # Crear campos para cada tipo con diseño de tabla
        for idx, tipo in enumerate(tipos_prelosa):
            # Etiqueta del tipo de prelosa
            ttk.Label(config_frame, text=tipo).grid(
                row=idx+2, column=0, sticky='w', padx=10, pady=8)
            
            # Campo de entrada para espaciamiento
            entry = ttk.Entry(
                config_frame, 
                textvariable=self.default_values[tipo]['espaciamiento'], 
                width=15
            )
            entry.grid(row=idx+2, column=1, sticky='w', padx=10, pady=8)
            
            # Agregar un pequeño descriptivo
            descripcion = ""
            if tipo == "PRELOSA MACIZA":
                descripcion = "Prefabricada sólida, mayor capacidad estructural"
            elif tipo == "PRELOSA ALIGERADA 20":
                descripcion = "Prefabricada con nervios, peso reducido"
            else:  # PRELOSA ALIGERADA 20 - 2 SENT
                descripcion = "Bidireccional, para mayores luces"
            
            ttk.Label(config_frame, text=descripcion, foreground="#666666").grid(
                row=idx+2, column=2, sticky='w', padx=10, pady=8)
    
    def crear_panel_accion(self):
        """Crear panel de botones de acción"""
        action_panel = ttk.Frame(self.main_container, style="Card.TFrame")
        action_panel.pack(fill=tk.X, pady=(0, 15))
        
        # Contenedor centrado para el botón
        button_container = ttk.Frame(action_panel)
        button_container.pack(pady=15)
        
        # Botón de procesamiento grande y destacado
        self.process_button = ttk.Button(
            button_container, 
            text="🚀 PROCESAR PRELOSAS",
            style="Primary.TButton",
            command=self.process_dxf
        )
        self.process_button.pack(pady=5)
        
        # Texto de ayuda bajo el botón
        hint_label = ttk.Label(
            button_container,
            text="Haga clic para procesar el archivo DXF con los parámetros configurados",
            foreground="#666666"
        )
        hint_label.pack(pady=(5, 0))
    
    def crear_panel_registro(self):
        """Crear panel de registro de procesamiento"""
        log_panel = ttk.LabelFrame(
            self.main_container, 
            text="Registro de Procesamiento",
            style="Config.TLabelframe",
            padding=15
        )
        log_panel.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # Área de texto para el registro
        self.log_area = scrolledtext.ScrolledText(
            log_panel, 
            wrap=tk.WORD, 
            width=80, 
            height=15, 
            font=('Consolas', 10),
            bg="#f8f8f8",
            fg="#333333"
        )
        self.log_area.pack(fill=tk.BOTH, expand=True)
        
        # Insertar texto inicial en el área de registro
        self.log_area.insert(tk.END, "Aplicación iniciada. Configure los parámetros y haga clic en PROCESAR PRELOSAS.\n")
        self.log_area.insert(tk.END, "Los resultados del procesamiento se mostrarán aquí.\n")
        self.log_area.config(state=tk.DISABLED)  # Hacer el área de solo lectura inicialmente
    
    def crear_pie_pagina(self):
        """Crear pie de página con información adicional"""
        footer_frame = ttk.Frame(self.main_container)
        footer_frame.pack(fill=tk.X, pady=(0, 5))
        
        # Línea separadora
        separator = ttk.Separator(footer_frame, orient='horizontal')
        separator.pack(fill=tk.X, pady=(0, 5))
        
        # Información de versión y copyright
        version_label = ttk.Label(
            footer_frame,
            text="ACERO SCRIPT v1.0.0 | © 2025 DODOD SOLUTIONS. Todos los derechos reservados.",
            foreground="#888888",
            font=('Segoe UI', 8)
        )
        version_label.pack(side=tk.RIGHT, padx=5)
    
    def select_dxf_file(self):
        """Seleccionar archivo DXF"""
        filename = filedialog.askopenfilename(
            title="Seleccionar Archivo DXF",
            filetypes=[("Archivos DXF", "*.dxf")]
        )
        if filename:
            self.dxf_path.set(filename)
            # Sugerir carpeta de salida
            default_output = os.path.join(os.path.dirname(filename), "Procesados")
            os.makedirs(default_output, exist_ok=True)
            self.output_path.set(default_output)
    
    def select_output_directory(self):
        """Seleccionar directorio de salida"""
        directory = filedialog.askdirectory(
            title="Seleccionar Carpeta de Salida"
        )
        if directory:
            self.output_path.set(directory)
    
    def process_dxf(self):
        """Procesar el archivo DXF"""
        # Validaciones
        if not self.dxf_path.get():
            messagebox.showerror("Error", "Debe seleccionar un archivo DXF")
            return
        
        if not os.path.exists(self.excel_path):
            messagebox.showerror("Error", f"No se encontró {self.excel_path}")
            return
        
        if not self.output_path.get():
            messagebox.showerror("Error", "Debe seleccionar un directorio de salida")
            return
        
        # Crear directorio de salida si no existe
        os.makedirs(self.output_path.get(), exist_ok=True)
        
        # Generar nombre de archivo de salida con número aleatorio
        nombre_archivo = os.path.splitext(os.path.basename(self.dxf_path.get()))[0]
        numero_random = generar_numero_random()
        output_dxf_path = os.path.join(
            self.output_path.get(), 
            f"{nombre_archivo}_{numero_random}.dxf"
        )
        
        # Preparar valores predeterminados
        valores_predeterminados = {
            tipo: {
                'espaciamiento': valores['espaciamiento'].get()
            } 
            for tipo, valores in self.default_values.items()
        }
        
        # Llamar directamente a la función de procesamiento
        proceso_dxf(
            self.dxf_path.get(), 
            self.excel_path,
            output_dxf_path,
            valores_predeterminados
        )
        
        # Mostrar mensaje de inicio
        messagebox.showinfo(
            "Procesamiento Iniciado", 
            f"Procesando archivo: {self.dxf_path.get()}\n"
            f"Archivo de salida: {output_dxf_path}"
        )

def main():
    # Modo GUI
    root = tk.Tk()
    app = DXFProcessorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()