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
    ╔══════════════════════════════════════════════════════════════════════════════════════════════════════════╗
    ║                                                                                                          ║
    ║       ▄▄▄       ▄████▄  ▓█████  ██▀███   ▒█████      ██████  ▄████▄   ███▀███   ██▓ ██▓███  ▄▄▄█████▓    ║
    ║      ▒████▄    ▒██▀ ▀█  ▓█   ▀ ▓██ ▒ ██▒▒██▒  ██▒   ▒██    ▒ ▒██▀ ▀█  ▓██ ▒ ██▒▓██▒▓██░  ██▒▓  ██▒ ▓▒    ║
    ║      ▒██  ▀█▄  ▒▓█    ▄ ▒███   ▓██ ░▄█ ▒▒██░  ██▒   ░ ▓██▄   ▒▓█    ▄ ▓██ ░▄█ ▒▒██▒▓██░ ██▓▒▒ ▓██░ ▒░    ║
    ║      ░██▄▄▄▄██ ▒▓▓▄ ▄██▒▒▓█  ▄ ▒██▀▀█▄  ▒██   ██░     ▒   ██▒▒▓▓▄ ▄██▒▒██▀▀█▄  ░██░▒██▄█▓▒ ▒░ ▓██▓ ░     ║
    ║       ▓█   ▓██▒▒ ▓███▀ ░░▒████▒░██▓ ▒██▒░ ████▓▒░   ▒██████▒▒▒ ▓███▀ ░░██▓ ▒██▒░██░▒██▒ ░  ░  ▒██▒ ░     ║
    ║       ▒▒   ▓▒█░░ ░▒ ▒  ░░░ ▒░ ░░ ▒▓ ░▒▓░░ ▒░▒░▒░    ▒ ▒▓▒ ▒ ░░ ░▒ ▒  ░░ ▒▓ ░▒▓░░▓  ▒▓▒░ ░  ░  ▒ ░░       ║
    ║        ▒   ▒▒ ░  ░  ▒    ░ ░  ░  ░▒ ░ ▒░  ░ ▒ ▒░    ░ ░▒  ░ ░  ░  ▒     ░▒ ░ ▒░ ▒ ░░▒ ░         ░        ║
    ║        ░   ▒   ░           ░     ░░   ░ ░ ░ ░ ▒     ░  ░  ░  ░          ░░   ░  ▒ ░░░         ░          ║
    ║            ░  ░░ ░         ░  ░   ░         ░ ░           ░  ░ ░         ░      ░                        ║
    ║                ░                                              ░                                          ║
    ║                                                                                                          ║
    ║                        Herramienta para Automatización de Aceros en Prelosas                             ║
    ║                                      by DODOD SOLUTIONS                                                  ║
    ║                                                                                                          ║
    ╚══════════════════════════════════════════════════════════════════════════════════════════════════════════╝
    """
    print(banner)



def proceso_dxf(dxf_path, excel_path, output_path, valores_predeterminados=None):
    """
    Función para ejecutar el procesamiento de DXF con valores personalizados.
    """
    # Abrir consola para mostrar mensajes
    import ctypes
    ctypes.windll.kernel32.AllocConsole()
    sys.stdout = open('CONOUT$', 'w')
    sys.stderr = open('CONOUT$', 'w')

    print("Iniciando procesamiento de DXF...")
    print(imprimir_banner_script())
    print(f"Archivo de entrada: {dxf_path}")
    print(f"Archivo Excel: {excel_path}")
    print(f"Archivo de salida: {output_path}")

    try:
        # Importar script dinámicamente
        script_module = import_module_from_path()
        
        if script_module is None:
            print("No se pudo importar el script")
            input("Presione Enter para continuar...")
            return

        print("Valores predeterminados recibidos:")
        print(valores_predeterminados)
        
        # Ejecutar procesamiento con valores predeterminados
        total = script_module.procesar_prelosas_con_bloques(
            dxf_path, 
            excel_path,
            output_path,
            valores_predeterminados
        )
        print(f"Procesamiento completado. Bloques insertados: {total}")
        
        # Abrir directorio de salida
        try:
            if platform.system() == "Windows":
                subprocess.Popen(f'explorer "{os.path.dirname(output_path)}"')
        except Exception as e:
            print(f"No se pudo abrir el directorio de salida: {e}")
        
        input("Presione Enter para cerrar...")
        
    except Exception as e:
        print(f"Error de procesamiento: {e}")
        traceback.print_exc()
        input("Presione Enter para cerrar...")

class DXFProcessorApp:
    def __init__(self, master):
        self.master = master
        master.title("Procesador de Prelosas - DODOD SOLUTIONS")
        master.geometry("800x900")
        master.configure(bg='#f0f0f0')
        
        # Directorio del script
        self.script_dir = os.path.dirname(os.path.abspath(__file__))
        
        # Ruta fija del archivo Excel
        self.excel_path = os.path.join(self.script_dir, "CONVERTIDOR.xlsx")
        
        # Frame principal
        self.main_frame = ttk.Frame(master, padding="10 10 10 10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Logo o título
        self.title_label = ttk.Label(
            self.main_frame, 
            text="DODOD SOLUTIONS - Automatización de Aceros en Prelosas", 
            font=('Arial', 12, 'bold')
        )
        self.title_label.pack(pady=(0,10))
        
        # Frame de selección de archivos
        self.file_frame = ttk.Frame(self.main_frame)
        self.file_frame.pack(fill=tk.X, pady=5)
        
        # DXF File
        ttk.Label(self.file_frame, text="Archivo DXF:").grid(row=0, column=0, sticky='w', padx=(0,5))
        self.dxf_path = tk.StringVar()
        self.dxf_entry = ttk.Entry(self.file_frame, textvariable=self.dxf_path, width=50)
        self.dxf_entry.grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(self.file_frame, text="Buscar", command=self.select_dxf_file).grid(row=0, column=2, padx=5)
        
        # Excel File (Fixed)
        ttk.Label(self.file_frame, text="Archivo Excel:").grid(row=1, column=0, sticky='w', padx=(0,5))
        self.excel_entry = ttk.Entry(self.file_frame, width=50, state='readonly')
        self.excel_entry.grid(row=1, column=1, padx=5, pady=5)
        self.excel_entry.insert(0, "CONVERTIDOR.xlsx (Fijo)")
        
        # Output Directory
        ttk.Label(self.file_frame, text="Directorio de Salida:").grid(row=2, column=0, sticky='w', padx=(0,5))
        self.output_path = tk.StringVar()
        self.output_entry = ttk.Entry(self.file_frame, textvariable=self.output_path, width=50)
        self.output_entry.grid(row=2, column=1, padx=5, pady=5)
        ttk.Button(self.file_frame, text="Seleccionar", command=self.select_output_directory).grid(row=2, column=2, padx=5)
        
        # Sección de Configuración de Valores Predeterminados
        self.configuracion_frame = ttk.LabelFrame(self.main_frame, text="Configuración de Valores Predeterminados")
        self.configuracion_frame.pack(fill=tk.X, pady=10, padx=10)

        # Variables para almacenar los valores predeterminados
        self.default_values = {
            'PRELOSA MACIZA': {
                'espaciamiento': tk.StringVar(value='0.20')
            },
            'PRELOSA ALIGERADA 20': {
                'espaciamiento': tk.StringVar(value='0.20')
            },
            'PRELOSA ALIGERADA 20 - 2 SENT': {
                'espaciamiento': tk.StringVar(value='0.605')
            }
        }

        # Crear campos para cada tipo de prelosa
        tipos_prelosa = [
            'PRELOSA MACIZA', 
            'PRELOSA ALIGERADA 20', 
            'PRELOSA ALIGERADA 20 - 2 SENT'
        ]

        for idx, tipo in enumerate(tipos_prelosa):
            # Frame para cada tipo de prelosa
            tipo_frame = ttk.Frame(self.configuracion_frame)
            tipo_frame.pack(fill=tk.X, pady=5)

            # Etiqueta del tipo de prelosa
            ttk.Label(tipo_frame, text=tipo, font=('Arial', 10, 'bold')).grid(row=0, column=0, sticky='w')

            # Solo mostrar el campo de espaciamiento
            ttk.Label(tipo_frame, text="Espaciamiento:").grid(row=0, column=1)
            ttk.Entry(tipo_frame, textvariable=self.default_values[tipo]['espaciamiento'], width=10).grid(row=0, column=2)
        
        # Botón de Procesamiento
        self.process_button = ttk.Button(
            self.main_frame, 
            text="🚀 Procesar Prelosas", 
            command=self.process_dxf
        )
        self.process_button.pack(pady=10)
        
        # Área de Registro
        ttk.Label(self.main_frame, text="Registro de Procesamiento:").pack()
        self.log_area = scrolledtext.ScrolledText(
            self.main_frame, 
            wrap=tk.WORD, 
            width=80, 
            height=20, 
            font=('Consolas', 10)
        )
        self.log_area.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
    
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