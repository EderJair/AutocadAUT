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
import threading
from PIL import Image, ImageTk  # Necesitarás instalar pillow: pip install pillow

class DXFProcessorApp:
    def __init__(self, master):
        self.master = master
        master.title("ACERO SCRIPT - DODOD SOLUTIONS v1.0.0")
        master.geometry("1000x950")
        master.minsize(900, 750)
        
        # Variables de estado
        self.is_dark_mode = tk.BooleanVar(value=False)
        self.processing = False
        
        # Directorio del script
        self.script_dir = os.path.dirname(os.path.abspath(__file__))
        
        # Cargar imágenes e íconos
        self.load_icons()
        
        # Configurar estilo personalizado
        self.configurar_estilo()
        
        # Ruta fija del archivo Excel
        self.excel_path = os.path.join(self.script_dir, "CONVERTIDOR.xlsx")
        
        # Frame principal
        self.main_frame = ttk.Frame(master)
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Crear menú
        self.crear_menu()
        
        # Sidebar y contenido principal
        self.crear_layout()
        
        # Información de estado en barra inferior
        self.crear_statusbar()
        
        # Mostrar mensaje de bienvenida
        self.mostrar_mensaje_bienvenida()
    
    def load_icons(self):
        """Cargar íconos para la interfaz"""
        # Aquí se cargarían los íconos, pero como no tenemos acceso a archivos
        # usaremos emojis unicode como sustitutos
        self.icon_file = "📄"
        self.icon_folder = "📁"
        self.icon_process = "🚀"
        self.icon_settings = "⚙️"
        self.icon_info = "ℹ️"
        self.icon_theme = "🌓"
        self.icon_help = "❓"
    
    def configurar_estilo(self):
        """Configurar estilos personalizados para la interfaz"""
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        self.update_theme()
    
    def update_theme(self):
        """Actualizar el tema claro/oscuro"""
        # Definir colores según el tema
        if self.is_dark_mode.get():
            # Tema oscuro
            self.colors = {
                'bg': '#2d2d2d',
                'fg': '#e0e0e0',
                'accent': '#03a9f4',
                'accent_dark': '#0288d1',
                'widget_bg': '#3d3d3d',
                'sidebar_bg': '#252525',
                'card_bg': '#3d3d3d',
                'input_bg': '#454545',
                'input_fg': '#e0e0e0',
                'header_fg': '#4fc3f7',
                'subheader_fg': '#81d4fa',
                'muted_fg': '#9e9e9e',
                'success': '#4caf50',
                'warning': '#ff9800',
                'error': '#f44336',
                'info': '#2196f3'
            }
        else:
            # Tema claro
            self.colors = {
                'bg': '#f5f5f5',
                'fg': '#212121',
                'accent': '#1976d2',
                'accent_dark': '#1565c0',
                'widget_bg': '#ffffff',
                'sidebar_bg': '#e0e0e0',
                'card_bg': '#ffffff',
                'input_bg': '#ffffff',
                'input_fg': '#212121',
                'header_fg': '#1976d2',
                'subheader_fg': '#2196f3',
                'muted_fg': '#757575',
                'success': '#4caf50',
                'warning': '#ff9800',
                'error': '#f44336',
                'info': '#2196f3'
            }
        
        # Aplicar colores a los estilos
        self.style.configure("TFrame", background=self.colors['bg'])
        self.style.configure("TLabel", background=self.colors['bg'], foreground=self.colors['fg'])
        self.style.configure("TButton", font=('Segoe UI', 10), padding=5)
        
        # Etiquetas
        self.style.configure("Header.TLabel", 
                            background=self.colors['bg'],
                            foreground=self.colors['header_fg'],
                            font=('Segoe UI', 16, 'bold'))
        
        self.style.configure("Subheader.TLabel", 
                            background=self.colors['bg'],
                            foreground=self.colors['subheader_fg'],
                            font=('Segoe UI', 12, 'bold'))
        
        self.style.configure("Muted.TLabel", 
                            background=self.colors['bg'],
                            foreground=self.colors['muted_fg'],
                            font=('Segoe UI', 9))
        
        # Botones
        self.style.configure("Accent.TButton", 
                           background=self.colors['accent'],
                           foreground="white",
                           font=('Segoe UI', 10, 'bold'))
        
        self.style.map("Accent.TButton",
                     background=[('active', self.colors['accent_dark']), 
                                ('pressed', self.colors['accent_dark'])])
        
        self.style.configure("Primary.TButton", 
                           background=self.colors['accent'],
                           foreground="white",
                           padding=10,
                           font=('Segoe UI', 12, 'bold'))
        
        self.style.map("Primary.TButton",
                     background=[('active', self.colors['accent_dark']), 
                                ('pressed', self.colors['accent_dark'])])
        
        # Sidebar
        self.style.configure("Sidebar.TFrame", 
                           background=self.colors['sidebar_bg'])
        
        self.style.configure("Sidebar.TLabel", 
                           background=self.colors['sidebar_bg'],
                           foreground=self.colors['fg'],
                           font=('Segoe UI', 10, 'bold'))
        
        # Tarjetas
        self.style.configure("Card.TFrame", 
                           background=self.colors['card_bg'],
                           relief="raised",
                           borderwidth=1)
        
        # LabelFrame
        self.style.configure("TLabelframe", 
                           background=self.colors['bg'],
                           foreground=self.colors['fg'])
        
        self.style.configure("TLabelframe.Label", 
                           background=self.colors['bg'],
                           foreground=self.colors['header_fg'],
                           font=('Segoe UI', 11, 'bold'))
        
        # Entradas
        self.style.configure("TEntry", 
                           fieldbackground=self.colors['input_bg'],
                           foreground=self.colors['input_fg'])
        
        # Combobox
        self.style.configure("TCombobox", 
                           fieldbackground=self.colors['input_bg'],
                           background=self.colors['widget_bg'],
                           foreground=self.colors['input_fg'])
        
        # Progressbar
        self.style.configure("Horizontal.TProgressbar", 
                           background=self.colors['accent'],
                           troughcolor=self.colors['widget_bg'])
        
        # Actualizar colores de fondo de la ventana principal
        self.master.configure(background=self.colors['bg'])
        
        # Actualizar widgets si ya existen
        if hasattr(self, 'main_frame'):
            self.main_frame.configure(style="TFrame")
            self.update_all_widgets()
    
    def update_all_widgets(self):
        """Actualizar todos los widgets con el nuevo tema"""
        # Aquí actualizaríamos manualmente los widgets importantes
        # que necesiten cambios específicos para el tema
        
        # Actualizar el área de log
        if hasattr(self, 'log_area'):
            bg_color = self.colors['input_bg']
            fg_color = self.colors['input_fg']
            self.log_area.config(bg=bg_color, fg=fg_color)
    
    def crear_menu(self):
        """Crear menú superior"""
        menu_bar = tk.Menu(self.master)
        self.master.config(menu=menu_bar)
        
        # Menú Archivo
        file_menu = tk.Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label="Archivo", menu=file_menu)
        file_menu.add_command(label="Seleccionar DXF", command=self.select_dxf_file)
        file_menu.add_command(label="Seleccionar Carpeta de Salida", command=self.select_output_directory)
        file_menu.add_separator()
        file_menu.add_command(label="Salir", command=self.master.quit)
        
        # Menú Procesar
        process_menu = tk.Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label="Procesar", menu=process_menu)
        process_menu.add_command(label="Procesar Prelosas", command=self.process_dxf)
        
        # Menú Ver
        view_menu = tk.Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label="Ver", menu=view_menu)
        view_menu.add_checkbutton(label="Modo Oscuro", variable=self.is_dark_mode, 
                                command=self.update_theme)
        
        # Menú Ayuda
        help_menu = tk.Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label="Ayuda", menu=help_menu)
        help_menu.add_command(label="Acerca de", command=self.show_about)
        help_menu.add_command(label="Documentación", command=self.show_documentation)
    
    def crear_layout(self):
        """Crear layout principal con sidebar y contenido"""
        # Panel principal con dos columnas
        self.paned_window = ttk.PanedWindow(self.main_frame, orient=tk.HORIZONTAL)
        self.paned_window.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Sidebar izquierdo
        self.sidebar_frame = ttk.Frame(self.paned_window, style="Sidebar.TFrame", width=200)
        self.paned_window.add(self.sidebar_frame, weight=1)
        
        # Contenido principal
        self.content_frame = ttk.Frame(self.paned_window)
        self.paned_window.add(self.content_frame, weight=4)
        
        # Llenar el sidebar
        self.crear_sidebar()
        
        # Llenar el contenido
        self.crear_contenido()
    
    def crear_sidebar(self):
        """Crear sidebar con enlaces rápidos"""
        # Logo o título en la parte superior
        header_frame = ttk.Frame(self.sidebar_frame, style="Sidebar.TFrame")
        header_frame.pack(fill=tk.X, pady=(15, 20))
        
        logo_label = ttk.Label(header_frame, 
                             text="ACERO SCRIPT", 
                             style="Header.TLabel",
                             background=self.colors['sidebar_bg'],
                             foreground=self.colors['header_fg'])
        logo_label.pack(pady=(5, 0))
        
        # Separador
        separator = ttk.Separator(self.sidebar_frame, orient='horizontal')
        separator.pack(fill=tk.X, padx=15, pady=5)
        
        # Lista de acciones rápidas
        actions_frame = ttk.Frame(self.sidebar_frame, style="Sidebar.TFrame")
        actions_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # Estilo para botones de sidebar
        self.style.configure("Sidebar.TButton", 
                           font=('Segoe UI', 10),
                           padding=10,
                           background=self.colors['sidebar_bg'])
        
        ttk.Button(actions_frame, 
                  text=f"{self.icon_file} Seleccionar DXF", 
                  style="Sidebar.TButton", 
                  command=self.select_dxf_file).pack(fill=tk.X, pady=3)
        
        ttk.Button(actions_frame, 
                  text=f"{self.icon_folder} Carpeta de Salida", 
                  style="Sidebar.TButton", 
                  command=self.select_output_directory).pack(fill=tk.X, pady=3)
        
        ttk.Button(actions_frame, 
                  text=f"{self.icon_process} Procesar Prelosas", 
                  style="Sidebar.TButton", 
                  command=self.process_dxf).pack(fill=tk.X, pady=3)
        
        ttk.Button(actions_frame, 
                  text=f"{self.icon_theme} Cambiar Tema", 
                  style="Sidebar.TButton", 
                  command=self.toggle_theme).pack(fill=tk.X, pady=3)
        
        ttk.Button(actions_frame, 
                  text=f"{self.icon_help} Ayuda", 
                  style="Sidebar.TButton", 
                  command=self.show_documentation).pack(fill=tk.X, pady=3)
        
        # Información del software en la parte inferior
        info_frame = ttk.Frame(self.sidebar_frame, style="Sidebar.TFrame")
        info_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=15)
        
        version_label = ttk.Label(info_frame, 
                                text="Versión 1.0.0", 
                                style="Muted.TLabel",
                                background=self.colors['sidebar_bg'])
        version_label.pack()
        
        company_label = ttk.Label(info_frame, 
                                text="DODOD SOLUTIONS", 
                                style="Muted.TLabel",
                                background=self.colors['sidebar_bg'])
        company_label.pack()
    
    def crear_contenido(self):
        """Crear contenido principal"""
        # Frame para el contenido con padding
        content_inner = ttk.Frame(self.content_frame)
        content_inner.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)
        
        # Cabecera
        header_frame = ttk.Frame(content_inner, style="Card.TFrame")
        header_frame.pack(fill=tk.X, pady=(0, 15))
        
        header_label = ttk.Label(header_frame, 
                               text="Automatización de Aceros en Prelosas", 
                               style="Header.TLabel",
                               padding=(15, 10))
        header_label.pack()
        
        # Crear un frame con columnas para la disposición de paneles
        panels_frame = ttk.Frame(content_inner)
        panels_frame.pack(fill=tk.BOTH, expand=True)
        
        # Columna izquierda para archivos y configuración
        left_column = ttk.Frame(panels_frame)
        left_column.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        # Columna derecha para progreso, acción y registro
        right_column = ttk.Frame(panels_frame)
        right_column.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))
        
        # Panel de archivos (columna izquierda)
        self.crear_panel_archivos(left_column)
        
        # Panel de configuración (columna izquierda)
        self.crear_panel_configuracion(left_column)
        
        # Barra de progreso (columna derecha)
        self.crear_barra_progreso(right_column)
        
        # Panel de acción (columna derecha)
        self.crear_panel_accion(right_column)
        
        # Panel de registro (columna derecha)
        self.crear_panel_registro(right_column)
    
    def crear_panel_archivos(self, parent):
        """Crear panel de selección de archivos"""
        file_panel = ttk.LabelFrame(parent, text="Selección de Archivos", padding=15)
        file_panel.pack(fill=tk.X, pady=(0, 15))
        
        # Grid para los campos
        file_frame = ttk.Frame(file_panel)
        file_frame.pack(fill=tk.X)
        
        # DXF File
        ttk.Label(file_frame, text=f"{self.icon_file} Archivo DXF:").grid(
            row=0, column=0, sticky='w', padx=(0,10), pady=8)
        
        self.dxf_path = tk.StringVar()
        self.dxf_entry = ttk.Entry(file_frame, textvariable=self.dxf_path, width=60)
        self.dxf_entry.grid(row=0, column=1, padx=5, pady=8, sticky='ew')
        
        ttk.Button(file_frame, text="Buscar", 
                  style="Accent.TButton",
                  command=self.select_dxf_file).grid(row=0, column=2, padx=5, pady=8)
        
        # Excel File (Fixed)
        ttk.Label(file_frame, text=f"{self.icon_file} Archivo Excel:").grid(
            row=1, column=0, sticky='w', padx=(0,10), pady=8)
        
        self.excel_entry = ttk.Entry(file_frame, width=60, state='readonly')
        self.excel_entry.grid(row=1, column=1, padx=5, pady=8, sticky='ew')
        self.excel_entry.insert(0, "CONVERTIDOR.xlsx (Predeterminado)")
        
        # Output Directory
        ttk.Label(file_frame, text=f"{self.icon_folder} Directorio de Salida:").grid(
            row=2, column=0, sticky='w', padx=(0,10), pady=8)
        
        self.output_path = tk.StringVar()
        self.output_entry = ttk.Entry(file_frame, textvariable=self.output_path, width=60)
        self.output_entry.grid(row=2, column=1, padx=5, pady=8, sticky='ew')
        
        ttk.Button(file_frame, text="Seleccionar", 
                  style="Accent.TButton",
                  command=self.select_output_directory).grid(row=2, column=2, padx=5, pady=8)
        
        # Expandir columna central
        file_frame.columnconfigure(1, weight=1)
    
    def crear_panel_configuracion(self, parent):
        """Crear panel de configuración de valores predeterminados"""
        config_panel = ttk.LabelFrame(parent, text="Configuración de Valores Predeterminados", padding=15)
        config_panel.pack(fill=tk.X, pady=(0, 15))
        
        # Variables para almacenar los valores predeterminados
        self.default_values = {
            'PRELOSA MACIZA': {
                'espaciamiento': tk.StringVar(value='0.20'),
                'acero': tk.StringVar(value='3/8"')
            },
            'PRELOSA MACIZA 15': {
                'espaciamiento': tk.StringVar(value='0.15'),
                'acero': tk.StringVar(value='3/8"')
            },
            'PRELOSA MACIZA TIPO 3': {
                'espaciamiento': tk.StringVar(value='0.20'),
                'acero': tk.StringVar(value='3/8"')
            },
            'PRELOSA MACIZA TIPO 4': {
                'espaciamiento': tk.StringVar(value='0.20'),
                'acero': tk.StringVar(value='3/8"')
            },
            'PRELOSA ALIGERADA 20': {
                'espaciamiento': tk.StringVar(value='0.20'),
                'acero': tk.StringVar(value='3/8"')
            },
            'PRELOSA ALIGERADA 20 - 2 SENT': {
                'espaciamiento': tk.StringVar(value='0.20'),
                'acero': tk.StringVar(value='3/8"')
            },
            'PRELOSA ALIGERADA 25': {
                'espaciamiento': tk.StringVar(value='0.20'),
                'acero': tk.StringVar(value='3/8"')
            }
            # ... resto de tipos predefinidos ...
        }
        
        # Lista dinámica para almacenar todos los tipos de prelosa (predefinidos + personalizados)
        self.tipos_prelosa = [
            'PRELOSA MACIZA', 
            'PRELOSA MACIZA 15',
            'PRELOSA MACIZA TIPO 3',
            'PRELOSA MACIZA TIPO 4',
            'PRELOSA ALIGERADA 20', 
            'PRELOSA ALIGERADA 20 - 2 SENT',
            'PRELOSA ALIGERADA 25',
            'PRELOSA ALIGERADA 25 - 2 SENT',
            'PRELOSA ALIGERADA 30',
            'PRELOSA ALIGERADA 30 - 2 SENT'
        ]
        
        # Opciones de acero disponibles
        self.acero_opciones = ['6mm', '8mm', '3/8"', '12mm', '1/2"', '5/8"', '3/4"', '1']
        
        # Crear notebook para las configuraciones
        notebook = ttk.Notebook(config_panel)
        notebook.pack(fill=tk.X, padx=5, pady=5)
        
        # Pestaña general
        general_tab = ttk.Frame(notebook)
        notebook.add(general_tab, text="General")
        
        # Añadir scrollable frame para la tabla de configuración
        # Primero, un canvas para el scrolling
        canvas = tk.Canvas(general_tab, height=250)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Scrollbar vertical
        vsb = ttk.Scrollbar(general_tab, orient="vertical", command=canvas.yview)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.configure(yscrollcommand=vsb.set)
        
        # Frame que contiene los widgets de configuración
        self.config_frame = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=self.config_frame, anchor=tk.NW)
        
        # Cabeceras de tabla
        ttk.Label(self.config_frame, text="Tipo de Prelosa", font=('Segoe UI', 10, 'bold')).grid(
            row=0, column=0, sticky='w', padx=10, pady=(0, 10))
        
        ttk.Label(self.config_frame, text="Espaciamiento", font=('Segoe UI', 10, 'bold')).grid(
            row=0, column=1, sticky='w', padx=10, pady=(0, 10))
        
        ttk.Label(self.config_frame, text="Acero", font=('Segoe UI', 10, 'bold')).grid(
            row=0, column=2, sticky='w', padx=10, pady=(0, 10))
        
        # Columna para botón Eliminar (para tipos personalizados)
        ttk.Label(self.config_frame, text="Acciones", font=('Segoe UI', 10, 'bold')).grid(
            row=0, column=3, sticky='w', padx=10, pady=(0, 10))
        
        # Separador horizontal
        separator = ttk.Separator(self.config_frame, orient='horizontal')
        separator.grid(row=1, column=0, columnspan=4, sticky='ew', pady=(0, 10))
        
        # Llenar la tabla con los tipos predefinidos
        self.render_prelosa_table()
        
        # Botón para agregar nuevo tipo de prelosa
        add_frame = ttk.Frame(self.config_frame)
        add_frame.grid(row=len(self.tipos_prelosa)+2, column=0, columnspan=4, sticky='ew', pady=10)
        
        add_button = ttk.Button(
            add_frame, 
            text="+ AGREGAR TIPO DE PRELOSA",
            style="Accent.TButton",
            command=self.add_new_prelosa_type
        )
        add_button.pack(pady=5)
        
        # Actualizar tamaño del canvas basado en su contenido
        def _configure_canvas(event):
            # Actualizar el scrollregion para que incluya todo el contenido
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas_width = event.width
            canvas.itemconfig(canvas_window, width=canvas_width)
        
        canvas_window = canvas.create_window((0, 0), window=self.config_frame, anchor="nw")
        self.config_frame.bind("<Configure>", _configure_canvas)
        canvas.bind("<Configure>", _configure_canvas)
        
        # Asegurar que el canvas tenga un tamaño razonable
        canvas.update_idletasks()
        canvas.config(width=500)

    def render_prelosa_table(self):
        """Renderiza la tabla de tipos de prelosa con sus valores"""
        # Limpiar widgets existentes a partir de la fila 2
        for widget in self.config_frame.grid_slaves():
            if int(widget.grid_info()["row"]) >= 2:
                widget.grid_forget()
        
        # Crear campos para cada tipo con diseño de tabla
        for idx, tipo in enumerate(self.tipos_prelosa):
            # Si es un tipo personalizado, puede que no esté en default_values
            if tipo not in self.default_values:
                self.default_values[tipo] = {
                    'espaciamiento': tk.StringVar(value='0.20'),
                    'acero': tk.StringVar(value='3/8"')
                }
            
            # Etiqueta del tipo de prelosa
            ttk.Label(self.config_frame, text=tipo).grid(
                row=idx+2, column=0, sticky='w', padx=10, pady=8)
            
            # Campo de entrada para espaciamiento
            entry = ttk.Entry(
                self.config_frame, 
                textvariable=self.default_values[tipo]['espaciamiento'], 
                width=15
            )
            entry.grid(row=idx+2, column=1, sticky='w', padx=10, pady=8)
            
            # ComboBox para selección de acero
            acero_combo = ttk.Combobox(
                self.config_frame,
                textvariable=self.default_values[tipo]['acero'],
                values=self.acero_opciones,
                width=10,
                state="readonly"
            )
            acero_combo.grid(row=idx+2, column=2, sticky='w', padx=10, pady=8)
            
            # Botón de eliminar (solo para tipos personalizados)
            if tipo not in ['PRELOSA MACIZA', 'PRELOSA MACIZA 15', 'PRELOSA MACIZA TIPO 3', 'PRELOSA MACIZA TIPO 4',
                            'PRELOSA ALIGERADA 20', 'PRELOSA ALIGERADA 20 - 2 SENT', 'PRELOSA ALIGERADA 25',
                            'PRELOSA ALIGERADA 25 - 2 SENT', 'PRELOSA ALIGERADA 30', 'PRELOSA ALIGERADA 30 - 2 SENT']:
                delete_button = ttk.Button(
                    self.config_frame,
                    text="Eliminar",
                    command=lambda t=tipo: self.delete_prelosa_type(t)
                )
                delete_button.grid(row=idx+2, column=3, padx=10, pady=8)
        
        # Botón para agregar nuevo tipo de prelosa
        add_frame = ttk.Frame(self.config_frame)
        add_frame.grid(row=len(self.tipos_prelosa)+2, column=0, columnspan=4, sticky='ew', pady=10)
        
        add_button = ttk.Button(
            add_frame, 
            text="+ AGREGAR TIPO DE PRELOSA",
            style="Accent.TButton",
            command=self.add_new_prelosa_type
        )
        add_button.pack(pady=5)

    def add_new_prelosa_type(self):
        """Muestra un diálogo para agregar un nuevo tipo de prelosa"""
        # Crear ventana de diálogo
        dialog = tk.Toplevel(self.master)
        dialog.title("Agregar Nuevo Tipo de Prelosa")
        dialog.geometry("400x200")
        dialog.resizable(False, False)
        dialog.transient(self.master)
        dialog.grab_set()
        
        # Configurar como modal
        dialog.focus_set()
        
        # Contenido
        content_frame = ttk.Frame(dialog, padding=20)
        content_frame.pack(fill=tk.BOTH, expand=True)
        
        # Nombre del tipo
        ttk.Label(content_frame, text="Nombre del tipo de prelosa:").grid(
            row=0, column=0, sticky='w', pady=(0, 10))
        
        nombre_var = tk.StringVar()
        nombre_entry = ttk.Entry(content_frame, textvariable=nombre_var, width=30)
        nombre_entry.grid(row=0, column=1, sticky='w', padx=(10, 0), pady=(0, 10))
        nombre_entry.focus()
        
        # Espaciamiento
        ttk.Label(content_frame, text="Espaciamiento predeterminado:").grid(
            row=1, column=0, sticky='w', pady=(0, 10))
        
        espaciamiento_var = tk.StringVar(value="0.20")
        espaciamiento_entry = ttk.Entry(content_frame, textvariable=espaciamiento_var, width=10)
        espaciamiento_entry.grid(row=1, column=1, sticky='w', padx=(10, 0), pady=(0, 10))
        
        # Acero
        ttk.Label(content_frame, text="Acero predeterminado:").grid(
            row=2, column=0, sticky='w', pady=(0, 10))
        
        acero_var = tk.StringVar(value="3/8\"")
        acero_combo = ttk.Combobox(
            content_frame,
            textvariable=acero_var,
            values=self.acero_opciones,
            width=10,
            state="readonly"
        )
        acero_combo.grid(row=2, column=1, sticky='w', padx=(10, 0), pady=(0, 10))
        
        # Botones
        button_frame = ttk.Frame(content_frame)
        button_frame.grid(row=3, column=0, columnspan=2, pady=(10, 0))
        
        def add_type():
            nombre = nombre_var.get().strip().upper()
            if not nombre:
                messagebox.showerror("Error", "Debe ingresar un nombre para el tipo de prelosa")
                return
            
            # Verificar si ya existe
            if nombre in self.tipos_prelosa:
                messagebox.showerror("Error", "Ya existe un tipo de prelosa con ese nombre")
                return
            
            # Agregar nuevo tipo
            self.tipos_prelosa.append(nombre)
            self.default_values[nombre] = {
                'espaciamiento': tk.StringVar(value=espaciamiento_var.get()),
                'acero': tk.StringVar(value=acero_var.get())
            }
            
            # Actualizar tabla
            self.render_prelosa_table()
            
            # Cerrar diálogo
            dialog.destroy()
            
            # Mensaje de éxito
            self.add_to_log(f"Se agregó el tipo de prelosa: {nombre}", "success")
        
        ttk.Button(
            button_frame,
            text="Agregar",
            style="Accent.TButton",
            command=add_type
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            button_frame,
            text="Cancelar",
            command=dialog.destroy
        ).pack(side=tk.LEFT, padx=5)

    def delete_prelosa_type(self, tipo):
        """Elimina un tipo de prelosa personalizado"""
        # Confirmar eliminación
        if messagebox.askyesno("Confirmar eliminación", 
                            f"¿Está seguro de eliminar el tipo de prelosa '{tipo}'?"):
            # Eliminar de la lista y del diccionario
            self.tipos_prelosa.remove(tipo)
            if tipo in self.default_values:
                del self.default_values[tipo]
            
            # Actualizar tabla
            self.render_prelosa_table()
            
            # Mensaje de éxito
            self.add_to_log(f"Se eliminó el tipo de prelosa: {tipo}", "info")
    
    def crear_barra_progreso(self, parent):
        """Crear barra de progreso"""
        progress_frame = ttk.Frame(parent)
        progress_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            progress_frame, 
            orient=tk.HORIZONTAL, 
            length=100, 
            mode='determinate',
            variable=self.progress_var
        )
        self.progress_bar.pack(fill=tk.X)
        
        # Etiqueta de estado
        self.status_label = ttk.Label(
            progress_frame, 
            text="Listo para procesar", 
            anchor=tk.CENTER
        )
        self.status_label.pack(fill=tk.X, pady=(5, 0))
    
    def crear_panel_accion(self, parent):
        """Crear panel de botones de acción"""
        action_panel = ttk.Frame(parent, style="Card.TFrame")
        action_panel.pack(fill=tk.X, pady=(0, 15))
        
        # Contenedor centrado para los botones
        button_container = ttk.Frame(action_panel)
        button_container.pack(pady=15)
        
        # Botón principal
        self.process_button = ttk.Button(
            button_container, 
            text=f"{self.icon_process} PROCESAR PRELOSAS",
            style="Primary.TButton",
            command=self.process_dxf
        )
        self.process_button.pack(pady=5)
        
        # Texto de ayuda
        hint_label = ttk.Label(
            button_container,
            text="Haga clic para procesar el archivo DXF con los parámetros configurados",
            foreground=self.colors['muted_fg']
        )
        hint_label.pack(pady=(5, 0))
    
    def crear_panel_registro(self, parent):
        """Crear panel de registro con formato mejorado"""
        log_panel = ttk.LabelFrame(parent, text="Registro de Procesamiento", padding=15)
        log_panel.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # Barra de herramientas para el log
        log_toolbar = ttk.Frame(log_panel)
        log_toolbar.pack(fill=tk.X, pady=(0, 5))
        
        # Botón para limpiar el log
        clear_button = ttk.Button(
            log_toolbar,
            text="Limpiar",
            style="Accent.TButton",
            command=self.clear_log
        )
        clear_button.pack(side=tk.RIGHT)
        
        # Área de texto para el registro con etiquetas para formateo
        self.log_area = tk.Text(
            log_panel, 
            wrap=tk.WORD, 
            width=80, 
            height=15, 
            font=('Consolas', 10),
            bg=self.colors['input_bg'],
            fg=self.colors['input_fg']
        )
        
        # Configurar etiquetas para colorear diferentes tipos de mensajes
        self.log_area.tag_configure("info", foreground=self.colors['info'])
        self.log_area.tag_configure("success", foreground=self.colors['success'])
        self.log_area.tag_configure("warning", foreground=self.colors['warning'])
        self.log_area.tag_configure("error", foreground=self.colors['error'])
        self.log_area.tag_configure("bold", font=('Consolas', 10, 'bold'))
        self.log_area.tag_configure("muted", foreground=self.colors['muted_fg'])
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(log_panel, orient="vertical", command=self.log_area.yview)
        self.log_area.configure(yscrollcommand=scrollbar.set)
        
        # Posicionar widgets
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_area.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    
    def crear_statusbar(self):
        """Crear barra de estado inferior"""
        self.status_bar = ttk.Frame(self.master, relief=tk.SUNKEN, style="TFrame")
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Separador
        separator = ttk.Separator(self.status_bar, orient='horizontal')
        separator.pack(fill=tk.X)
        
        # Contenido de la barra de estado
        status_frame = ttk.Frame(self.status_bar)
        status_frame.pack(fill=tk.X, padx=5, pady=2)
        
        # Información de versión
        version_label = ttk.Label(
            status_frame,
            text="ACERO SCRIPT v1.0.0",
            style="Muted.TLabel"
        )
        version_label.pack(side=tk.LEFT)
        
        # Información de copyright
        copyright_label = ttk.Label(
            status_frame,
            text="© 2025 DODOD SOLUTIONS",
            style="Muted.TLabel"
        )
        copyright_label.pack(side=tk.RIGHT)
    
    def mostrar_mensaje_bienvenida(self):
        """Mostrar mensaje de bienvenida en el área de log"""
        self.log_area.config(state=tk.NORMAL)
        self.log_area.delete(1.0, tk.END)
        
        self.log_area.insert(tk.END, "=== ACERO SCRIPT - Sistema de Automatización de Prelosas ===\n\n", "bold")
        self.log_area.insert(tk.END, "Bienvenido al sistema de procesamiento de prelosas.\n\n", "info")
        self.log_area.insert(tk.END, "Para comenzar:\n", "bold")
        self.log_area.insert(tk.END, "1. Seleccione un archivo DXF\n")
        self.log_area.insert(tk.END, "2. Elija una carpeta de salida\n")
        self.log_area.insert(tk.END, "3. Configure los valores predeterminados según sus necesidades\n")
        self.log_area.insert(tk.END, "4. Haga clic en PROCESAR PRELOSAS para iniciar el procesamiento\n\n")
        
        self.log_area.insert(tk.END, "Los resultados del procesamiento se mostrarán en esta área.\n", "info")
        self.log_area.config(state=tk.DISABLED)
    
    def select_dxf_file(self):
        """Seleccionar archivo DXF"""
        filename = filedialog.askopenfilename(
            title="Seleccionar Archivo DXF",
            filetypes=[("Archivos DXF", "*.dxf")]
        )
        if filename:
            self.dxf_path.set(filename)
            
            # Añadir entrada en el log
            self.add_to_log(f"Archivo DXF seleccionado: {filename}", "info")
            
            # Sugerir carpeta de salida
            default_output = os.path.join(os.path.dirname(filename), "Procesados")
            os.makedirs(default_output, exist_ok=True)
            self.output_path.set(default_output)
            
            # Actualizar estado
            self.status_label.config(text=f"Archivo seleccionado: {os.path.basename(filename)}")
    
    def select_output_directory(self):
        """Seleccionar directorio de salida"""
        directory = filedialog.askdirectory(
            title="Seleccionar Carpeta de Salida"
        )
        if directory:
            self.output_path.set(directory)
            self.add_to_log(f"Carpeta de salida seleccionada: {directory}", "info")
    
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
            
            # Evitar procesamiento múltiple
            if self.processing:
                messagebox.showinfo("Procesando", "Ya hay un proceso en ejecución")
                return
            
            # Crear directorio de salida si no existe
            os.makedirs(self.output_path.get(), exist_ok=True)
            
            # Generar nombre de archivo de salida con número aleatorio
            nombre_archivo = os.path.splitext(os.path.basename(self.dxf_path.get()))[0]
            numero_random = self.generar_numero_random()
            output_dxf_path = os.path.join(
                self.output_path.get(), 
                f"{nombre_archivo}_{numero_random}.dxf"
            )
            
            # Preparar valores predeterminados
            valores_predeterminados = {
                tipo: {
                    'espaciamiento': valores['espaciamiento'].get(),
                    'acero': valores['acero'].get()  # Añadir valor de acero
                } 
                for tipo, valores in self.default_values.items()
            }

            tipos_personalizados = [
                tipo for tipo in self.tipos_prelosa 
                if tipo not in ['PRELOSA MACIZA', 'PRELOSA MACIZA 15', 'PRELOSA MACIZA TIPO 3', 'PRELOSA MACIZA TIPO 4',
                                'PRELOSA ALIGERADA 20', 'PRELOSA ALIGERADA 20 - 2 SENT', 'PRELOSA ALIGERADA 25',
                                'PRELOSA ALIGERADA 25 - 2 SENT', 'PRELOSA ALIGERADA 30', 'PRELOSA ALIGERADA 30 - 2 SENT']
            ]
                    
            # Limpiar log y mostrar información inicial
            self.clear_log()
            self.add_to_log("Iniciando procesamiento de DXF...", "bold")
            self.add_to_log(f"Archivo de entrada: {self.dxf_path.get()}", "info")
            self.add_to_log(f"Archivo Excel: {self.excel_path}", "info")
            self.add_to_log(f"Archivo de salida: {output_dxf_path}", "info")
            self.add_to_log("Valores predeterminados:", "bold")
            
            for tipo, valores in valores_predeterminados.items():
                self.add_to_log(f"  {tipo}: espaciamiento = {valores['espaciamiento']}, acero = {valores['acero']}")
            
            # Configurar interfaz para procesamiento
            self.process_button.config(state=tk.DISABLED)
            self.status_label.config(text="Procesando...")
            self.progress_var.set(0)
            self.processing = True
            
            # Iniciar procesamiento en hilo separado
            self.processing_thread = threading.Thread(
                target=self.run_processing,
                args=(self.dxf_path.get(), self.excel_path, output_dxf_path, valores_predeterminados)
            )
            self.processing_thread.daemon = True
            self.processing_thread.start()
            
            # Iniciar actualización de progreso
            self.master.after(100, self.update_progress)
    
    def run_processing(self, dxf_path, excel_path, output_path, valores_predeterminados):
        """Ejecutar el procesamiento en un hilo separado"""
        try:
            # Importar script dinámicamente
            script_module = self.import_module_from_path()
            
            if script_module is None:
                self.add_to_log("Error: No se pudo importar el script", "error")
                return
            
            # Ejecutar procesamiento
            total = script_module.procesar_prelosas_con_bloques(
                dxf_path, 
                excel_path,
                output_path,
                valores_predeterminados
            )
            
            # Actualizar log con el resultado
            self.add_to_log(f"Procesamiento completado. Bloques insertados: {total}", "success")
            
            # Preguntar si quiere abrir la carpeta
            if messagebox.askyesno("Procesamiento completado", 
                                 f"Se han insertado {total} bloques.\n¿Desea abrir la carpeta de destino?"):
                self.open_output_folder(os.path.dirname(output_path))
        
        except Exception as e:
            self.add_to_log(f"Error durante el procesamiento: {str(e)}", "error")
            self.add_to_log(traceback.format_exc(), "error")
            messagebox.showerror("Error", f"Error durante el procesamiento: {str(e)}")
        
        finally:
            # Restaurar interfaz
            self.master.after(0, self.restore_interface)
    
    def update_progress(self):
        """Actualizar barra de progreso durante el procesamiento"""
        if not self.processing:
            return
        
        # Incrementar progreso (simulado)
        current = self.progress_var.get()
        if current < 100:
            # Incremento variable para simular progreso
            increment = min(2, 100 - current)
            self.progress_var.set(current + increment)
        
        # Verificar si el hilo sigue activo
        if self.processing_thread.is_alive():
            self.master.after(100, self.update_progress)
        else:
            self.progress_var.set(100)
            self.status_label.config(text="Procesamiento completado")
    
    def restore_interface(self):
        """Restaurar interfaz después del procesamiento"""
        self.process_button.config(state=tk.NORMAL)
        self.processing = False
    
    def add_to_log(self, message, tag=None):
        """Añadir mensaje al área de log con formato opcional"""
        self.log_area.config(state=tk.NORMAL)
        
        # Añadir timestamp
        timestamp = time.strftime("%H:%M:%S")
        self.log_area.insert(tk.END, f"[{timestamp}] ", "muted")
        
        # Añadir mensaje con formato opcional
        if tag:
            self.log_area.insert(tk.END, f"{message}\n", tag)
        else:
            self.log_area.insert(tk.END, f"{message}\n")
        
        # Desplazar al final
        self.log_area.see(tk.END)
        self.log_area.config(state=tk.DISABLED)
    
    def clear_log(self):
        """Limpiar área de log"""
        self.log_area.config(state=tk.NORMAL)
        self.log_area.delete(1.0, tk.END)
        self.log_area.config(state=tk.DISABLED)
    
    def toggle_theme(self):
        """Alternar entre tema claro y oscuro"""
        self.is_dark_mode.set(not self.is_dark_mode.get())
        self.update_theme()
    
    def show_about(self):
        """Mostrar ventana 'Acerca de'"""
        about_window = tk.Toplevel(self.master)
        about_window.title("Acerca de ACERO SCRIPT")
        about_window.geometry("400x300")
        about_window.resizable(False, False)
        about_window.transient(self.master)
        about_window.grab_set()
        
        # Configurar como modal
        about_window.focus_set()
        
        # Contenido
        content_frame = ttk.Frame(about_window, padding=20)
        content_frame.pack(fill=tk.BOTH, expand=True)
        
        # Título
        ttk.Label(
            content_frame, 
            text="ACERO SCRIPT", 
            font=('Segoe UI', 16, 'bold'),
            foreground=self.colors['header_fg']
        ).pack(pady=(0, 10))
        
        # Versión
        ttk.Label(
            content_frame,
            text="Versión 1.0.0",
            font=('Segoe UI', 10)
        ).pack()
        
        # Descripción
        ttk.Label(
            content_frame,
            text="Sistema de Automatización para Procesamiento de Prelosas",
            wraplength=350,
            justify="center"
        ).pack(pady=10)
        
        # Copyright
        ttk.Label(
            content_frame,
            text="© 2025 DODOD SOLUTIONS\nTodos los derechos reservados.",
            justify="center"
        ).pack(pady=10)
        
        # Botón de cerrar
        ttk.Button(
            content_frame,
            text="Cerrar",
            command=about_window.destroy,
            style="Accent.TButton"
        ).pack(pady=10)
    
    def show_documentation(self):
        """Mostrar documentación o ayuda"""
        # Aquí se podría abrir un archivo PDF o HTML con la documentación
        messagebox.showinfo(
            "Documentación",
            "La documentación completa está disponible en el manual de usuario.\n\n"
            "Para más información, contacte con soporte@dododsolutions.com"
        )
    
    def open_output_folder(self, folder_path):
        """Abrir la carpeta de salida"""
        try:
            if platform.system() == "Windows":
                os.startfile(folder_path)
            elif platform.system() == "Darwin":  # macOS
                subprocess.Popen(["open", folder_path])
            else:  # Linux
                subprocess.Popen(["xdg-open", folder_path])
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir la carpeta: {str(e)}")
    
    def create_tooltip(self, widget, text):
        """Crear tooltip para un widget"""
        def enter(event):
            x, y, _, _ = widget.bbox("insert")
            x += widget.winfo_rootx() + 25
            y += widget.winfo_rooty() + 25
            
            # Crear ventana de tooltip
            self.tooltip = tk.Toplevel(widget)
            self.tooltip.wm_overrideredirect(True)
            self.tooltip.wm_geometry(f"+{x}+{y}")
            
            label = ttk.Label(
                self.tooltip, 
                text=text, 
                background=self.colors['info'],
                foreground="white",
                relief="solid", 
                borderwidth=1,
                padding=5
            )
            label.pack()
        
        def leave(event):
            if hasattr(self, 'tooltip'):
                self.tooltip.destroy()
        
        widget.bind("<Enter>", enter)
        widget.bind("<Leave>", leave)
    
    def generar_numero_random(self, digitos=2):
        """Genera un número aleatorio con el número de dígitos especificado"""
        min_valor = 10 ** (digitos - 1)
        max_valor = (10 ** digitos) - 1
        return random.randint(min_valor, max_valor)
    
    def import_module_from_path(self):
        """Importar script.py dinámicamente"""
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
                self.add_to_log(f"Error: No se encontró script.py en {script_path}", "error")
                return None
            
            # Importar el módulo
            spec = importlib.util.spec_from_file_location("script", script_path)
            script = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(script)
            return script
        except Exception as e:
            self.add_to_log(f"Error importando script: {str(e)}", "error")
            self.add_to_log(traceback.format_exc(), "error")
            return None

def main():
    """Función principal"""
    # Configurar tema de alta DPI para Windows
    if platform.system() == "Windows":
        try:
            from ctypes import windll
            windll.shcore.SetProcessDpiAwareness(1)
        except:
            pass
    
    # Iniciar GUI
    root = tk.Tk()
    root.title("ACERO SCRIPT - DODOD SOLUTIONS")
    
    # Configurar icono si existe
    try:
        icon_path = os.path.join(os.path.dirname(__file__), "icon.ico")
        if os.path.exists(icon_path):
            root.iconbitmap(icon_path)
    except:
        pass
    
    # Crear aplicación
    app = DXFProcessorApp(root)
    
    # Centrar ventana
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')
    
    # Iniciar bucle principal
    root.mainloop()

if __name__ == "__main__":
    main()