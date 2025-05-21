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
from PIL import Image, ImageTk  # You'll need to install pillow: pip install pillow

class DXFProcessorApp:
    def __init__(self, master):
        self.master = master
        master.title("ACERO SCRIPT v1.0.0")
        master.geometry("1000x700")
        master.minsize(900, 650)
        
        # State variables
        self.is_dark_mode = tk.BooleanVar(value=False)
        self.processing = False
        
        # Script directory
        self.script_dir = os.path.dirname(os.path.abspath(__file__))
        
        # Fixed Excel file path
        self.excel_path = os.path.join(self.script_dir, "CONVERTIDOR.xlsx")
        
        # Configure style and colors
        self.configure_style()
        
        # Create the sidebar and main content area
        self.create_layout()
        
        # Show welcome message
        self.show_welcome_message()
    
    def configure_style(self):
        """Configure custom styles for the interface"""
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # Define colors based on theme
        self.update_colors()
        
        # Configure common styles
        self.style.configure("TFrame", background=self.colors['bg'])
        self.style.configure("TLabel", background=self.colors['bg'], foreground=self.colors['fg'])
        
        # Sidebar styles
        self.style.configure("Sidebar.TFrame", 
                           background=self.colors['sidebar_bg'],
                           relief="flat")
        
        self.style.configure("Sidebar.TLabel", 
                           background=self.colors['sidebar_bg'],
                           foreground=self.colors['sidebar_fg'],
                           font=('Arial', 10))
        
        self.style.configure("SidebarTitle.TLabel", 
                           background=self.colors['sidebar_bg'],
                           foreground=self.colors['accent'],
                           font=('Arial', 14, 'bold'))
        
        # Tab buttons in sidebar
        self.style.configure("Tab.TButton", 
                           font=('Arial', 11),
                           padding=10,
                           width=20,
                           anchor="w")
        
        self.style.map("Tab.TButton",
                     background=[('active', self.colors['sidebar_active']), 
                                 ('selected', self.colors['sidebar_active']),
                                 ('!active', self.colors['sidebar_bg'])],
                     foreground=[('active', self.colors['sidebar_active_fg']), 
                                 ('selected', self.colors['sidebar_active_fg']),
                                 ('!active', self.colors['sidebar_fg'])])
        
        # Button styles
        self.style.configure("TButton", 
                           padding=8,
                           font=('Arial', 10))
        
        self.style.configure("Primary.TButton", 
                           font=('Arial', 11, 'bold'),
                           padding=10,
                           background=self.colors['accent'],
                           foreground="#ffffff")
        
        self.style.map("Primary.TButton",
                     background=[('active', self.colors['accent_dark']), ('!active', self.colors['accent'])],
                     foreground=[('active', '#ffffff'), ('!active', '#ffffff')])
        
        # Content frame
        self.style.configure("Content.TFrame", 
                           background=self.colors['content_bg'])
        
        # Card frame
        self.style.configure("Card.TFrame", 
                           background=self.colors['card_bg'],
                           relief="solid",
                           borderwidth=0)
        
        # LabelFrame
        self.style.configure("TLabelframe", 
                           background=self.colors['card_bg'],
                           foreground=self.colors['fg'],
                           borderwidth=0,
                           relief="solid")
        
        self.style.configure("TLabelframe.Label", 
                           background=self.colors['card_bg'],
                           foreground=self.colors['accent'],
                           font=('Arial', 12, 'bold'))
        
        # Section headers
        self.style.configure("Section.TLabel", 
                           background=self.colors['content_bg'],
                           foreground=self.colors['accent'],
                           font=('Arial', 14, 'bold'))
        
        self.style.configure("Subsection.TLabel", 
                           background=self.colors['card_bg'],
                           foreground=self.colors['accent'],
                           font=('Arial', 12, 'bold'))
        
        # Entry
        self.style.configure("TEntry", 
                           padding=8,
                           fieldbackground=self.colors['input_bg'])
        
        # Combobox
        self.style.configure("TCombobox", 
                           padding=8,
                           fieldbackground=self.colors['input_bg'])
        
        # Footer
        self.style.configure("Footer.TFrame", 
                           background=self.colors['sidebar_bg'])
        
        self.style.configure("Footer.TLabel", 
                           background=self.colors['sidebar_bg'],
                           foreground=self.colors['muted_fg'],
                           font=('Arial', 9))
        
        # Progress bar
        self.style.configure("Horizontal.TProgressbar", 
                           background=self.colors['accent'],
                           troughcolor=self.colors['widget_bg'])
        
        # Icon buttons
        self.style.configure("Icon.TButton", 
                          padding=4,
                          font=('Arial', 12),
                          background=self.colors['sidebar_bg'],
                          foreground=self.colors['sidebar_fg'])
        
        self.style.map("Icon.TButton",
                    background=[('active', self.colors['sidebar_active']), ('!active', self.colors['sidebar_bg'])],
                    foreground=[('active', self.colors['sidebar_active_fg']), ('!active', self.colors['sidebar_fg'])])
    
    def update_colors(self):
        """Update color scheme based on theme"""
        if self.is_dark_mode.get():
            # Dark theme
            self.colors = {
                'bg': '#1e1e1e',
                'fg': '#f0f0f0',
                'accent': '#3498db',
                'accent_dark': '#2980b9',
                'sidebar_bg': '#252526',
                'sidebar_fg': '#cccccc',
                'sidebar_active': '#37373d',
                'sidebar_active_fg': '#ffffff',
                'content_bg': '#1e1e1e',
                'widget_bg': '#333333',
                'card_bg': '#2d2d2d',
                'input_bg': '#3c3c3c',
                'input_fg': '#f0f0f0',
                'muted_fg': '#888888',
                'success': '#27ae60',
                'warning': '#f39c12',
                'error': '#e74c3c',
                'info': '#3498db',
                'border': '#3d3d3d'
            }
        else:
            # Light theme
            self.colors = {
                'bg': '#f5f5f5',
                'fg': '#333333',
                'accent': '#2980b9',
                'accent_dark': '#1c6ca1',
                'sidebar_bg': '#2c3e50',
                'sidebar_fg': '#ecf0f1',
                'sidebar_active': '#34495e',
                'sidebar_active_fg': '#ffffff',
                'content_bg': '#f5f5f5',
                'widget_bg': '#ffffff',
                'card_bg': '#ffffff',
                'input_bg': '#ffffff',
                'input_fg': '#333333',
                'muted_fg': '#95a5a6',
                'success': '#27ae60',
                'warning': '#f39c12',
                'error': '#e74c3c',
                'info': '#3498db',
                'border': '#e0e0e0'
            }
    
    def create_layout(self):
        """Create the main layout with sidebar and content area"""
        # Main container with paned window
        self.main_container = tk.PanedWindow(self.master, orient=tk.HORIZONTAL, 
                                            background=self.colors['bg'],
                                            sashwidth=4, sashpad=0, 
                                            sashrelief=tk.FLAT)
        self.main_container.pack(fill=tk.BOTH, expand=True)
        
        # Create sidebar
        self.sidebar = ttk.Frame(self.main_container, style="Sidebar.TFrame")
        self.main_container.add(self.sidebar, width=220)
        
        # Create content area
        self.content_area = ttk.Frame(self.main_container, style="Content.TFrame", padding=15)
        self.main_container.add(self.content_area, width=700)
        
        # Fill sidebar with navigation
        self.create_sidebar()
        
        # Create content frames for each section but only show the first one
        self.create_content_frames()
        self.show_frame('procesamiento')
    
    def create_sidebar(self):
        """Create the sidebar with navigation and branding"""
        # Branding header
        header_frame = ttk.Frame(self.sidebar, style="Sidebar.TFrame")
        header_frame.pack(fill=tk.X, padx=15, pady=20)
        
        # Title
        title_label = ttk.Label(
            header_frame, 
            text="ACERO SCRIPT", 
            style="SidebarTitle.TLabel"
        )
        title_label.pack(anchor=tk.W)
        
        subtitle_label = ttk.Label(
            header_frame, 
            text="Automatización de Prelosas", 
            style="Sidebar.TLabel"
        )
        subtitle_label.pack(anchor=tk.W, pady=(5, 0))
        
        # Separator
        separator = ttk.Separator(self.sidebar, orient='horizontal')
        separator.pack(fill=tk.X, padx=15, pady=15)
        
        # Navigation buttons
        nav_frame = ttk.Frame(self.sidebar, style="Sidebar.TFrame")
        nav_frame.pack(fill=tk.X, padx=5, pady=10)
        
        # Navigation buttons using TButton with Tab style
        self.selected_tab = tk.StringVar(value="procesamiento")
        
        procesamiento_btn = ttk.Button(
            nav_frame,
            text="📄 Procesamiento",
            command=lambda: self.show_frame('procesamiento'),
            style="Tab.TButton"
        )
        procesamiento_btn.pack(fill=tk.X, pady=2)
        
        configuracion_btn = ttk.Button(
            nav_frame,
            text="⚙️ Configuración",
            command=lambda: self.show_frame('configuracion'),
            style="Tab.TButton"
        )
        configuracion_btn.pack(fill=tk.X, pady=2)
        
        log_btn = ttk.Button(
            nav_frame,
            text="📋 Registro",
            command=lambda: self.show_frame('log'),
            style="Tab.TButton"
        )
        log_btn.pack(fill=tk.X, pady=2)
        
        # Store references to buttons for handling selection state
        self.nav_buttons = {
            'procesamiento': procesamiento_btn,
            'configuracion': configuracion_btn,
            'log': log_btn
        }
        
        # Actions section in sidebar - Create a small frame for toolbar buttons
        actions_frame = ttk.Frame(self.sidebar, style="Sidebar.TFrame")
        actions_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=15, pady=15)
        
        # Theme toggle button
        self.theme_text = tk.StringVar(value="🌙" if not self.is_dark_mode.get() else "☀️")
        theme_button = ttk.Button(
            actions_frame,
            textvariable=self.theme_text,
            command=self.toggle_theme,
            style="Icon.TButton",
            width=3
        )
        theme_button.pack(side=tk.LEFT, padx=5)
        
        # Help button
        help_button = ttk.Button(
            actions_frame,
            text="❓",
            command=self.show_documentation,
            style="Icon.TButton",
            width=3
        )
        help_button.pack(side=tk.LEFT, padx=5)
        
        # About button
        about_button = ttk.Button(
            actions_frame,
            text="ℹ️",
            command=self.show_about,
            style="Icon.TButton",
            width=3
        )
        about_button.pack(side=tk.LEFT, padx=5)
        
        # Version info at the bottom
        version_label = ttk.Label(
            self.sidebar,
            text="v1.0.0 • DODOD SOLUTIONS",
            style="Footer.TLabel"
        )
        version_label.pack(side=tk.BOTTOM, fill=tk.X, padx=15, pady=10)
    
    def create_content_frames(self):
        """Create frames for each content section"""
        self.frames = {}
        
        # Procesamiento content
        procesamiento_frame = ttk.Frame(self.content_area, style="Content.TFrame")
        self.create_procesamiento_content(procesamiento_frame)
        self.frames['procesamiento'] = procesamiento_frame
        
        # Configuracion content
        configuracion_frame = ttk.Frame(self.content_area, style="Content.TFrame")
        self.create_configuracion_content(configuracion_frame)
        self.frames['configuracion'] = configuracion_frame
        
        # Log content
        log_frame = ttk.Frame(self.content_area, style="Content.TFrame")
        self.create_log_content(log_frame)
        self.frames['log'] = log_frame
    
    def show_frame(self, frame_id):
        """Show the selected content frame and hide others"""
        # Update button states
        for key, button in self.nav_buttons.items():
            if key == frame_id:
                button.state(['selected'])
            else:
                button.state(['!selected'])
        
        # Hide all frames
        for frame in self.frames.values():
            frame.pack_forget()
        
        # Show selected frame
        self.frames[frame_id].pack(fill=tk.BOTH, expand=True)
        
        # Update selected tab variable
        self.selected_tab.set(frame_id)
    
    def create_procesamiento_content(self, parent):
        """Create the processing content"""
        # Section title
        ttk.Label(
            parent,
            text="Procesamiento de Archivos",
            style="Section.TLabel"
        ).pack(anchor=tk.W, pady=(0, 15))
        
        # File selection card
        file_card = ttk.Frame(parent, style="Card.TFrame", padding=20)
        file_card.pack(fill=tk.X, pady=(0, 15))
        
        # Card title
        ttk.Label(
            file_card,
            text="Selección de Archivos",
            style="Subsection.TLabel"
        ).pack(anchor=tk.W, pady=(0, 15))
        
        # DXF File Row
        file_frame = ttk.Frame(file_card, style="Card.TFrame")
        file_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(file_frame, text="Archivo DXF:", width=15).pack(side=tk.LEFT)
        
        self.dxf_path = tk.StringVar()
        self.dxf_entry = ttk.Entry(file_frame, textvariable=self.dxf_path)
        self.dxf_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        browse_dxf_button = ttk.Button(
            file_frame, 
            text="Explorar", 
            command=self.select_dxf_file
        )
        browse_dxf_button.pack(side=tk.LEFT, padx=(5, 0))
        
        # Excel File Row (fixed)
        excel_frame = ttk.Frame(file_card, style="Card.TFrame")
        excel_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(excel_frame, text="Archivo Excel:", width=15).pack(side=tk.LEFT)
        
        self.excel_entry = ttk.Entry(excel_frame, state='readonly')
        self.excel_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.excel_entry.insert(0, "CONVERTIDOR.xlsx (Predeterminado)")
        
        # Output Directory Row
        output_frame = ttk.Frame(file_card, style="Card.TFrame")
        output_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(output_frame, text="Directorio Salida:", width=15).pack(side=tk.LEFT)
        
        self.output_path = tk.StringVar()
        self.output_entry = ttk.Entry(output_frame, textvariable=self.output_path)
        self.output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        browse_output_button = ttk.Button(
            output_frame, 
            text="Explorar", 
            command=self.select_output_directory
        )
        browse_output_button.pack(side=tk.LEFT, padx=(5, 0))
        
        # Processing card
        process_card = ttk.Frame(parent, style="Card.TFrame", padding=20)
        process_card.pack(fill=tk.X, pady=15)
        
        # Card title
        ttk.Label(
            process_card,
            text="Estado del Procesamiento",
            style="Subsection.TLabel"
        ).pack(anchor=tk.W, pady=(0, 15))
        
        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            process_card, 
            orient=tk.HORIZONTAL, 
            length=100, 
            mode='determinate',
            variable=self.progress_var
        )
        self.progress_bar.pack(fill=tk.X)
        
        # Status label
        self.status_label = ttk.Label(
            process_card, 
            text="Listo para procesar", 
            anchor=tk.CENTER,
            background=self.colors['card_bg']
        )
        self.status_label.pack(fill=tk.X, pady=(5, 15))
        
        # Process button
        self.process_button = ttk.Button(
            process_card, 
            text="PROCESAR PRELOSAS",
            style="Primary.TButton",
            command=self.process_dxf
        )
        self.process_button.pack(pady=(0, 5))
    
    def create_configuracion_content(self, parent):
        """Create the configuration content"""
        # Section title
        ttk.Label(
            parent,
            text="Configuración de Tipos de Prelosa",
            style="Section.TLabel"
        ).pack(anchor=tk.W, pady=(0, 15))
        
        # Description
        ttk.Label(
            parent,
            text="Configure los espaciamientos y tipos de acero predeterminados para cada tipo de prelosa",
            background=self.colors['content_bg']
        ).pack(anchor=tk.W, pady=(0, 15))
        
        # Variables to store default values
        self.default_values = {
            'PRELOSA MACIZA': {
                'espaciamiento_long': tk.StringVar(value='0.20'),
                'espaciamiento_trans': tk.StringVar(value='0.20'),
                'acero': tk.StringVar(value='3/8"')
            },
            'PRELOSA MACIZA 15': {
                'espaciamiento_long': tk.StringVar(value='0.15'),
                'espaciamiento_trans': tk.StringVar(value='0.15'),
                'acero': tk.StringVar(value='3/8"')
            },
            'PRELOSA ALIGERADA 20': {
                'espaciamiento_long': tk.StringVar(value='0.20'),
                'espaciamiento_trans': tk.StringVar(value='0.20'),
                'acero': tk.StringVar(value='3/8"')
            },
            'PRELOSA ALIGERADA 20 - 2 SENT': {
                'espaciamiento_long': tk.StringVar(value='0.20'),
                'espaciamiento_trans': tk.StringVar(value='0.20'),
                'acero': tk.StringVar(value='3/8"')
            }
        }
        
        # List to store all prelosa types (predefined + custom)
        self.tipos_prelosa = [
            'PRELOSA MACIZA', 
            'PRELOSA MACIZA 15',
            'PRELOSA ALIGERADA 20', 
            'PRELOSA ALIGERADA 20 - 2 SENT'
        ]
        
        # Available steel options
        self.acero_opciones = ['6mm', '8mm', '3/8"', '12mm', '1/2"', '5/8"', '3/4"', '1']
        
        # Table container (card style)
        self.config_card = ttk.Frame(parent, style="Card.TFrame", padding=20)
        self.config_card.pack(fill=tk.BOTH, expand=True)
        
        # Headers
        # Headers
        headers_frame = ttk.Frame(self.config_card, style="Card.TFrame")
        headers_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(
            headers_frame, 
            text="Tipo de Prelosa",
            font=('Arial', 11, 'bold'),
            background=self.colors['card_bg'],
            foreground=self.colors['accent'],
            width=30
        ).grid(row=0, column=0, padx=(0, 15), sticky='w')

        ttk.Label(
            headers_frame, 
            text="Esp. Longitudinal",
            font=('Arial', 11, 'bold'),
            background=self.colors['card_bg'],
            foreground=self.colors['accent'],
            width=15
        ).grid(row=0, column=1, padx=5, sticky='w')

        ttk.Label(
            headers_frame, 
            text="Esp. Transversal",
            font=('Arial', 11, 'bold'),
            background=self.colors['card_bg'],
            foreground=self.colors['accent'],
            width=15
        ).grid(row=0, column=2, padx=5, sticky='w')

        ttk.Label(
            headers_frame, 
            text="Acero",
            font=('Arial', 11, 'bold'),
            background=self.colors['card_bg'],
            foreground=self.colors['accent'],
            width=10
        ).grid(row=0, column=3, padx=5, sticky='w')

        ttk.Label(
            headers_frame, 
            text="Acciones",
            font=('Arial', 11, 'bold'),
            background=self.colors['card_bg'],
            foreground=self.colors['accent'],
            width=10
        ).grid(row=0, column=4, padx=5, sticky='w')
        
        # Separator
        separator = ttk.Separator(self.config_card, orient='horizontal')
        separator.pack(fill=tk.X, pady=(0, 10))
        
        # Scrollable container for types
        container_frame = ttk.Frame(self.config_card, style="Card.TFrame")
        container_frame.pack(fill=tk.BOTH, expand=True)
        
        # Canvas for scrolling
        self.canvas = tk.Canvas(container_frame, background=self.colors['card_bg'], 
                               highlightthickness=0, height=300)
        scrollbar = ttk.Scrollbar(container_frame, orient="vertical", command=self.canvas.yview)
        
        # Create a frame inside the canvas to hold the types
        self.tipos_frame = ttk.Frame(self.canvas, style="Card.TFrame")
        
        # Configure canvas scrolling
        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Create window inside canvas
        canvas_window = self.canvas.create_window((0, 0), window=self.tipos_frame, anchor=tk.NW)
        
        # Update scroll region when frame changes
        def update_scroll_region(event):
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
            # Adjust the width of the window to fill the canvas
            self.canvas.itemconfig(canvas_window, width=self.canvas.winfo_width())
        
        self.tipos_frame.bind("<Configure>", update_scroll_region)
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfig(canvas_window, width=e.width))
        
        # Render initial types
        self.render_prelosa_types()
        
        # Add button
        add_button = ttk.Button(
            self.config_card, 
            text="+ AGREGAR TIPO DE PRELOSA",
            style="Primary.TButton",
            command=self.add_new_prelosa_type
        )
        add_button.pack(pady=15)
    
    def render_prelosa_types(self):
        """Render the prelosa types in the configuration tab"""
        # Clear all existing widgets
        for widget in self.tipos_frame.winfo_children():
            widget.destroy()
        
        # Create a row for each type
        for idx, tipo in enumerate(self.tipos_prelosa):
            # Create a row frame
            row_frame = ttk.Frame(self.tipos_frame, style="Card.TFrame")
            row_frame.pack(fill=tk.X, pady=5)
            
            # Background color alternating
            bg_color = self.colors['card_bg']
            
            # Type name
            ttk.Label(
                row_frame, 
                text=tipo,
                background=bg_color,
                width=30
            ).grid(row=0, column=0, padx=(0, 15), sticky='w')
            
            # Longitudinal spacing entry
            spacing_long_entry = ttk.Entry(
                row_frame, 
                textvariable=self.default_values[tipo]['espaciamiento_long'],
                width=15
            )
            spacing_long_entry.grid(row=0, column=1, padx=5, sticky='w')
            
            # Transversal spacing entry
            spacing_trans_entry = ttk.Entry(
                row_frame, 
                textvariable=self.default_values[tipo]['espaciamiento_trans'],
                width=15
            )
            spacing_trans_entry.grid(row=0, column=2, padx=5, sticky='w')
            
            # Steel combobox
            steel_combo = ttk.Combobox(
                row_frame,
                textvariable=self.default_values[tipo]['acero'],
                values=self.acero_opciones,
                width=10,
                state="readonly"
            )
            steel_combo.grid(row=0, column=3, padx=5, sticky='w')
            
            # Delete button (only for custom types)
            predefined_types = [
                'PRELOSA MACIZA', 'PRELOSA MACIZA 15',
                'PRELOSA ALIGERADA 20', 'PRELOSA ALIGERADA 20 - 2 SENT'
            ]
            
            if tipo not in predefined_types:
                delete_button = ttk.Button(
                    row_frame,
                    text="Eliminar",
                    command=lambda t=tipo: self.delete_prelosa_type(t)
                )
                delete_button.grid(row=0, column=4, padx=5, sticky='w')
            
            # Add separator after each row (except last)
            if idx < len(self.tipos_prelosa) - 1:
                separator_frame = ttk.Frame(self.tipos_frame, height=1, style="Card.TFrame")
                separator_frame.pack(fill=tk.X, pady=5)
                separator = ttk.Separator(separator_frame, orient='horizontal')
            separator.pack(fill=tk.X)

    def create_log_content(self, parent):
        """Create the log content"""
        # Section title
        ttk.Label(
            parent,
            text="Registro de Procesamiento",
            style="Section.TLabel"
        ).pack(anchor=tk.W, pady=(0, 15))
        
        # Log card
        log_card = ttk.Frame(parent, style="Card.TFrame", padding=20)
        log_card.pack(fill=tk.BOTH, expand=True)
        
        # Toolbar
        toolbar_frame = ttk.Frame(log_card, style="Card.TFrame")
        toolbar_frame.pack(fill=tk.X, pady=(0, 10))
        
        clear_button = ttk.Button(
            toolbar_frame,
            text="Limpiar Registro",
            command=self.clear_log
        )
        clear_button.pack(side=tk.RIGHT)
        
        # Separator
        separator = ttk.Separator(log_card, orient='horizontal')
        separator.pack(fill=tk.X, pady=(0, 10))
        
        # Log area with scrollbar
        log_container = ttk.Frame(log_card, style="Card.TFrame")
        log_container.pack(fill=tk.BOTH, expand=True)
        
        # Create text area for log
        self.log_area = tk.Text(
            log_container, 
            wrap=tk.WORD, 
            font=('Consolas', 10),
            bg=self.colors['input_bg'],
            fg=self.colors['input_fg'],
            borderwidth=0,
            padx=10,
            pady=10
        )
        
        # Configure tags for coloring different message types
        self.log_area.tag_configure("info", foreground=self.colors['info'])
        self.log_area.tag_configure("success", foreground=self.colors['success'])
        self.log_area.tag_configure("warning", foreground=self.colors['warning'])
        self.log_area.tag_configure("error", foreground=self.colors['error'])
        self.log_area.tag_configure("bold", font=('Consolas', 10, 'bold'))
        self.log_area.tag_configure("muted", foreground=self.colors['muted_fg'])
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(log_container, orient="vertical", command=self.log_area.yview)
        self.log_area.configure(yscrollcommand=scrollbar.set)
        
        # Position widgets
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_area.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    
    def update_theme(self):
        """Toggle between light and dark theme"""
        self.is_dark_mode.set(not self.is_dark_mode.get())
        self.update_colors()
        self.configure_style()
        
        # Update background of main window
        self.master.configure(background=self.colors['bg'])
        
        # Update log area colors
        if hasattr(self, 'log_area'):
            self.log_area.config(
                bg=self.colors['input_bg'],
                fg=self.colors['input_fg']
            )
            
            # Update tags
            self.log_area.tag_configure("info", foreground=self.colors['info'])
            self.log_area.tag_configure("success", foreground=self.colors['success'])
            self.log_area.tag_configure("warning", foreground=self.colors['warning'])
            self.log_area.tag_configure("error", foreground=self.colors['error'])
        
        # Update canvas background in configuration tab
        if hasattr(self, 'canvas'):
            self.canvas.configure(background=self.colors['card_bg'])
        
        # Refresh the tab frames
        if hasattr(self, 'frames'):
            # First store which frame is currently active
            current_tab = self.selected_tab.get()
            
            # Recreate the frames
            for frame in self.frames.values():
                frame.destroy()
            
            self.create_content_frames()
            
            # Show the previously active frame
            self.show_frame(current_tab)
    
    def toggle_theme(self):
        """Toggle between light and dark theme"""
        self.is_dark_mode.set(not self.is_dark_mode.get())
        self.theme_text.set("☀️" if self.is_dark_mode.get() else "🌙")
        self.update_theme()
    
    def select_dxf_file(self):
        """Select DXF file"""
        filename = filedialog.askopenfilename(
            title="Seleccionar Archivo DXF",
            filetypes=[("Archivos DXF", "*.dxf")]
        )
        if filename:
            self.dxf_path.set(filename)
            
            # Add entry to log
            self.add_to_log(f"Archivo DXF seleccionado: {filename}", "info")
            
            # Suggest output folder
            default_output = os.path.join(os.path.dirname(filename), "Procesados")
            os.makedirs(default_output, exist_ok=True)
            self.output_path.set(default_output)
            
            # Update status
            self.status_label.config(text=f"Archivo seleccionado: {os.path.basename(filename)}")
    
    def select_output_directory(self):
        """Select output directory"""
        directory = filedialog.askdirectory(
            title="Seleccionar Carpeta de Salida"
        )
        if directory:
            self.output_path.set(directory)
            self.add_to_log(f"Carpeta de salida seleccionada: {directory}", "info")
    
    def add_new_prelosa_type(self):
        """Show dialog to add a new prelosa type"""
        # Crear una ventana de diálogo simple sin dependencias complejas
        dialog = tk.Toplevel(self.master)
        dialog.title("Agregar Nuevo Tipo de Prelosa")
        dialog.geometry("400x350")  # Aumentar altura para el nuevo campo
        dialog.resizable(False, False)
        dialog.transient(self.master)
        dialog.grab_set()
        
        # Aplicar color de fondo
        dialog.configure(bg=self.colors['bg'])
        
        # Título
        title_label = tk.Label(
            dialog, 
            text="Agregar Nuevo Tipo de Prelosa",
            font=('Arial', 14, 'bold'),
            bg=self.colors['bg'],
            fg=self.colors['accent']
        )
        title_label.pack(pady=(20, 20), padx=20, anchor='w')
        
        # Marco de formulario
        form_frame = tk.Frame(dialog, bg=self.colors['bg'])
        form_frame.pack(fill='x', padx=20)
        
        # Campo Nombre
        nombre_label = tk.Label(form_frame, text="Nombre:", bg=self.colors['bg'], fg=self.colors['fg'])
        nombre_label.grid(row=0, column=0, sticky='w', pady=(0, 15))
        
        nombre_var = tk.StringVar()
        nombre_entry = ttk.Entry(form_frame, textvariable=nombre_var, width=30)
        nombre_entry.grid(row=0, column=1, sticky='w', padx=(15, 0), pady=(0, 15))
        nombre_entry.focus()
        
        # Campo Espaciamiento Longitudinal
        espac_long_label = tk.Label(form_frame, text="Espaciamiento Long:", bg=self.colors['bg'], fg=self.colors['fg'])
        espac_long_label.grid(row=1, column=0, sticky='w', pady=(0, 15))
        
        espaciamiento_long_var = tk.StringVar(value="0.20")
        espaciamiento_long_entry = ttk.Entry(form_frame, textvariable=espaciamiento_long_var, width=10)
        espaciamiento_long_entry.grid(row=1, column=1, sticky='w', padx=(15, 0), pady=(0, 15))
        
        # Campo Espaciamiento Transversal
        espac_trans_label = tk.Label(form_frame, text="Espaciamiento Trans:", bg=self.colors['bg'], fg=self.colors['fg'])
        espac_trans_label.grid(row=2, column=0, sticky='w', pady=(0, 15))
        
        espaciamiento_trans_var = tk.StringVar(value="0.20")
        espaciamiento_trans_entry = ttk.Entry(form_frame, textvariable=espaciamiento_trans_var, width=10)
        espaciamiento_trans_entry.grid(row=2, column=1, sticky='w', padx=(15, 0), pady=(0, 15))
        
        # Campo Acero
        acero_label = tk.Label(form_frame, text="Acero:", bg=self.colors['bg'], fg=self.colors['fg'])
        acero_label.grid(row=3, column=0, sticky='w', pady=(0, 15))
        
        acero_var = tk.StringVar(value="3/8\"")
        acero_combo = ttk.Combobox(
            form_frame,
            textvariable=acero_var,
            values=self.acero_opciones,
            width=10,
            state="readonly"
        )
        acero_combo.grid(row=3, column=1, sticky='w', padx=(15, 0), pady=(0, 15))
        
        # Marco de botones
        button_frame = tk.Frame(dialog, bg=self.colors['bg'])
        button_frame.pack(side='bottom', fill='x', padx=20, pady=20)
        
        def add_type():
            nombre = nombre_var.get().strip().upper()
            if not nombre:
                messagebox.showerror("Error", "Debe ingresar un nombre para el tipo de prelosa")
                return
            
            # Check if it already exists
            if nombre in self.tipos_prelosa:
                messagebox.showerror("Error", "Ya existe un tipo de prelosa con ese nombre")
                return
            
            # Add new type
            self.tipos_prelosa.append(nombre)
            self.default_values[nombre] = {
                'espaciamiento_long': tk.StringVar(value=espaciamiento_long_var.get()),
                'espaciamiento_trans': tk.StringVar(value=espaciamiento_trans_var.get()),
                'acero': tk.StringVar(value=acero_var.get())
            }
            
            # Update table
            self.render_prelosa_types()
            
            # Close dialog
            dialog.destroy()
            
            # Success message
            self.add_to_log(f"Se agregó el tipo de prelosa: {nombre}", "success")
        
        # Botones: cancelar y agregar
        cancel_button = tk.Button(
            button_frame,
            text="Cancelar",
            width=10,
            bg=self.colors['widget_bg'],
            fg=self.colors['fg'],
            relief='solid',
            bd=1,
            command=dialog.destroy
        )
        cancel_button.pack(side='right', padx=5)
        
        add_button = tk.Button(
            button_frame,
            text="Agregar",
            width=10,
            bg=self.colors['accent'],
            fg='white',
            relief='solid',
            bd=1,
            command=add_type
        )
        add_button.pack(side='right', padx=5)

    def delete_prelosa_type(self, tipo):
        """Delete a custom prelosa type"""
        # Confirm deletion
        if messagebox.askyesno(
            "Confirmar eliminación", 
            f"¿Está seguro de eliminar el tipo de prelosa '{tipo}'?"
        ):
            # Remove from list and dictionary
            self.tipos_prelosa.remove(tipo)
            if tipo in self.default_values:
                del self.default_values[tipo]
            
            # Update table
            self.render_prelosa_types()
            
            # Success message
            self.add_to_log(f"Se eliminó el tipo de prelosa: {tipo}", "info")
    
    def process_dxf(self):
        """Process the DXF file"""
        # Validations
        if not self.dxf_path.get():
            messagebox.showerror("Error", "Debe seleccionar un archivo DXF")
            return
        
        if not os.path.exists(self.excel_path):
            messagebox.showerror("Error", f"No se encontró {self.excel_path}")
            return
        
        if not self.output_path.get():
            messagebox.showerror("Error", "Debe seleccionar un directorio de salida")
            return
        
        # Avoid multiple processing
        if self.processing:
            messagebox.showinfo("Procesando", "Ya hay un proceso en ejecución")
            return
        
        # Create output directory if it doesn't exist
        os.makedirs(self.output_path.get(), exist_ok=True)
        
        # Generate output filename with random number
        file_name = os.path.splitext(os.path.basename(self.dxf_path.get()))[0]
        random_number = self.generate_random_number()
        output_dxf_path = os.path.join(
            self.output_path.get(), 
            f"{file_name}_{random_number}.dxf"
        )
        
        # Prepare default values
        # Prepare default values
        default_values = {
            tipo: {
                'espaciamiento_long': valores['espaciamiento_long'].get(),
                'espaciamiento_trans': valores['espaciamiento_trans'].get(),
                'acero': valores['acero'].get()
            } 
            for tipo, valores in self.default_values.items()
        }
        
        custom_types = [
            tipo for tipo in self.tipos_prelosa 
            if tipo not in ['PRELOSA MACIZA', 'PRELOSA MACIZA 15',
                           'PRELOSA ALIGERADA 20', 'PRELOSA ALIGERADA 20 - 2 SENT']
        ]
        
        # Clear log and show initial information
        self.clear_log()
        self.add_to_log("Iniciando procesamiento de DXF...", "bold")
        self.add_to_log(f"Archivo de entrada: {self.dxf_path.get()}", "info")
        self.add_to_log(f"Archivo Excel: {self.excel_path}", "info")
        self.add_to_log(f"Archivo de salida: {output_dxf_path}", "info")
        self.add_to_log("Valores predeterminados:", "bold")
        
# Add to log
        for tipo, valores in default_values.items():
            self.add_to_log(f"  {tipo}: esp. long = {valores['espaciamiento_long']}, esp. trans = {valores['espaciamiento_trans']}, acero = {valores['acero']}")
        
        # Configure interface for processing
        self.process_button.config(state=tk.DISABLED)
        self.status_label.config(text="Procesando...")
        self.progress_var.set(0)
        self.processing = True
        
        # Start processing in separate thread
        self.processing_thread = threading.Thread(
            target=self.run_processing,
            args=(self.dxf_path.get(), self.excel_path, output_dxf_path, default_values)
        )
        self.processing_thread.daemon = True
        self.processing_thread.start()
        
        # Start progress update
        self.master.after(100, self.update_progress)
    
    def run_processing(self, dxf_path, excel_path, output_path, default_values):
        """Run processing in separate thread"""
        try:
            # Import script dynamically
            script_module = self.import_module_from_path()
            
            if script_module is None:
                self.add_to_log("Error: No se pudo importar el script", "error")
                return
            
            # Run processing
            total = script_module.procesar_prelosas_con_bloques(
                dxf_path, 
                excel_path,
                output_path,
                default_values
            )
            
            # Update log with result
            self.add_to_log(f"Procesamiento completado. Bloques insertados: {total}", "success")
            
            # Ask if user wants to open folder
            if messagebox.askyesno(
                "Procesamiento completado", 
                f"Se han insertado {total} bloques.\n¿Desea abrir la carpeta de destino?"
            ):
                self.open_output_folder(os.path.dirname(output_path))
        
        except Exception as e:
            self.add_to_log(f"Error durante el procesamiento: {str(e)}", "error")
            self.add_to_log(traceback.format_exc(), "error")
            messagebox.showerror("Error", f"Error durante el procesamiento: {str(e)}")
        
        finally:
            # Restore interface
            self.master.after(0, self.restore_interface)
    
    def update_progress(self):
        """Update progress bar during processing"""
        if not self.processing:
            return
        
        # Increment progress (simulated)
        current = self.progress_var.get()
        if current < 100:
            # Variable increment to simulate progress
            increment = min(2, 100 - current)
            self.progress_var.set(current + increment)
        
        # Check if thread is still active
        if self.processing_thread.is_alive():
            self.master.after(100, self.update_progress)
        else:
            self.progress_var.set(100)
            self.status_label.config(text="Procesamiento completado")
    
    def restore_interface(self):
        """Restore interface after processing"""
        self.process_button.config(state=tk.NORMAL)
        self.processing = False
    
    def add_to_log(self, message, tag=None):
        """Add message to log area with optional formatting"""
        # If log tab doesn't exist, create it
        if not hasattr(self, 'log_area'):
            return
            
        self.log_area.config(state=tk.NORMAL)
        
        # Add timestamp
        timestamp = time.strftime("%H:%M:%S")
        self.log_area.insert(tk.END, f"[{timestamp}] ", "muted")
        
        # Add message with optional formatting
        if tag:
            self.log_area.insert(tk.END, f"{message}\n", tag)
        else:
            self.log_area.insert(tk.END, f"{message}\n")
        
        # Scroll to end
        self.log_area.see(tk.END)
        self.log_area.config(state=tk.DISABLED)
        
        # If log tab is not visible, show a small indication on the log button
        if self.selected_tab.get() != 'log':
            self.nav_buttons['log'].configure(text="📋 Registro (•)")
    
    def clear_log(self):
        """Clear log area"""
        self.log_area.config(state=tk.NORMAL)
        self.log_area.delete(1.0, tk.END)
        self.log_area.config(state=tk.DISABLED)
        
        # Reset the log button text
        self.nav_buttons['log'].configure(text="📋 Registro")
    
    def show_welcome_message(self):
        """Display welcome message in the log area"""
        if not hasattr(self, 'log_area'):
            # Create log area first by showing the log tab
            self.show_frame('log')
            # Then switch back to the processing tab
            self.show_frame('procesamiento')
        
        self.clear_log()
        
        self.log_area.config(state=tk.NORMAL)
        self.log_area.insert(tk.END, "=== ACERO SCRIPT - Sistema de Automatización de Prelosas ===\n\n", "bold")
        self.log_area.insert(tk.END, "Bienvenido al sistema de procesamiento de prelosas.\n\n", "info")
        self.log_area.insert(tk.END, "Pasos para comenzar:\n", "bold")
        self.log_area.insert(tk.END, "1. Seleccione un archivo DXF\n")
        self.log_area.insert(tk.END, "2. Elija una carpeta de salida\n")
        self.log_area.insert(tk.END, "3. Configure los valores predeterminados según sus necesidades\n")
        self.log_area.insert(tk.END, "4. Haga clic en PROCESAR PRELOSAS para iniciar el procesamiento\n\n")
        
        self.log_area.insert(tk.END, "Los resultados del procesamiento se mostrarán en esta área.\n", "info")
        self.log_area.config(state=tk.DISABLED)
    
    def show_about(self):
        """Show About window"""
        about_window = tk.Toplevel(self.master)
        about_window.title("Acerca de ACERO SCRIPT")
        about_window.geometry("450x350")
        about_window.resizable(False, False)
        about_window.transient(self.master)
        about_window.grab_set()
        about_window.focus_set()
        
        # Configure background
        about_window.configure(background=self.colors['bg'])
        
        # Content
        content_frame = ttk.Frame(about_window, style="Card.TFrame", padding=30)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Logo (placeholder)
        logo_label = ttk.Label(
            content_frame, 
            text="🏢", 
            font=('Arial', 48),
            foreground=self.colors['accent'],
            background=self.colors['card_bg']
        )
        logo_label.pack(pady=(0, 15))
        
        # Title
        ttk.Label(
            content_frame, 
            text="ACERO SCRIPT", 
            font=('Arial', 18, 'bold'),
            foreground=self.colors['accent'],
            background=self.colors['card_bg']
        ).pack(pady=(0, 5))
        
        # Version
        ttk.Label(
            content_frame,
            text="Versión 1.0.0",
            font=('Arial', 12),
            background=self.colors['card_bg']
        ).pack()
        
        # Description
        ttk.Label(
            content_frame,
            text="Sistema de Automatización para Procesamiento de Prelosas",
            wraplength=350,
            justify="center",
            background=self.colors['card_bg'],
            font=('Arial', 11)
        ).pack(pady=15)
        
        # Copyright
        ttk.Label(
            content_frame,
            text="© 2025 DODOD SOLUTIONS\nTodos los derechos reservados.",
            justify="center",
            background=self.colors['card_bg'],
            font=('Arial', 10)
        ).pack(pady=15)
        
        # Close button
        ttk.Button(
            content_frame,
            text="Cerrar",
            command=about_window.destroy,
            style="Primary.TButton"
        ).pack(pady=10)
    
    def show_documentation(self):
        """Show documentation or help"""
        help_window = tk.Toplevel(self.master)
        help_window.title("Ayuda - ACERO SCRIPT")
        help_window.geometry("600x500")
        help_window.transient(self.master)
        help_window.grab_set()
        help_window.focus_set()
        
        # Configure background
        help_window.configure(background=self.colors['bg'])
        
        # Content frame
        main_frame = ttk.Frame(help_window, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        ttk.Label(
            main_frame,
            text="Documentación de ACERO SCRIPT",
            style="Section.TLabel"
        ).pack(anchor=tk.W, pady=(0, 15))
        
        # Help content in a card
        help_card = ttk.Frame(main_frame, style="Card.TFrame", padding=20)
        help_card.pack(fill=tk.BOTH, expand=True)
        
        # Create text area for help content
        help_text = tk.Text(
            help_card,
            wrap=tk.WORD,
            width=70,
            height=20,
            font=('Arial', 11),
            bg=self.colors['card_bg'],
            fg=self.colors['fg'],
            borderwidth=0,
            padx=10,
            pady=10
        )
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(help_card, orient="vertical", command=help_text.yview)
        help_text.configure(yscrollcommand=scrollbar.set)
        
        # Position widgets
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        help_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Add help content
        help_text.insert(tk.END, "Guía de Usuario - ACERO SCRIPT\n\n", ("title",))
        help_text.insert(tk.END, "ACERO SCRIPT es una aplicación para automatizar el procesamiento de archivos DXF de prelosas. A continuación se detallan los pasos para utilizar la aplicación:\n\n")
        
        help_text.insert(tk.END, "1. Pestaña de Procesamiento\n", ("section",))
        help_text.insert(tk.END, "   - Seleccione un archivo DXF utilizando el botón 'Explorar'\n")
        help_text.insert(tk.END, "   - El archivo Excel CONVERTIDOR.xlsx se utiliza automáticamente\n")
        help_text.insert(tk.END, "   - Seleccione una carpeta de destino para los archivos procesados\n")
        help_text.insert(tk.END, "   - Haga clic en 'PROCESAR PRELOSAS' para iniciar el procesamiento\n\n")
        
        help_text.insert(tk.END, "2. Pestaña de Configuración\n", ("section",))
        help_text.insert(tk.END, "   - Configure los espaciamientos y tipos de acero para cada tipo de prelosa\n")
        help_text.insert(tk.END, "   - Puede agregar nuevos tipos de prelosa utilizando el botón correspondiente\n")
        help_text.insert(tk.END, "   - Los tipos personalizados pueden ser eliminados si ya no se necesitan\n\n")
        
        help_text.insert(tk.END, "3. Pestaña de Registro\n", ("section",))
        help_text.insert(tk.END, "   - Muestra un registro detallado del procesamiento\n")
        help_text.insert(tk.END, "   - Puede limpiar el registro con el botón 'Limpiar Registro'\n\n")
        
        help_text.insert(tk.END, "Para más información o soporte técnico, contacte a soporte@dododsolutions.com\n\n")
        
        # Configure text tags
        help_text.tag_configure("title", font=('Arial', 14, 'bold'), foreground=self.colors['accent'])
        help_text.tag_configure("section", font=('Arial', 12, 'bold'))
        
        # Make text read-only
        help_text.config(state=tk.DISABLED)
        
        # Button frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(15, 0))
        
        # Close button
        ttk.Button(
            button_frame,
            text="Cerrar",
            command=help_window.destroy,
            style="Primary.TButton"
        ).pack(side=tk.RIGHT)
    
    def open_output_folder(self, folder_path):
        """Open output folder"""
        try:
            if platform.system() == "Windows":
                os.startfile(folder_path)
            elif platform.system() == "Darwin":  # macOS
                subprocess.Popen(["open", folder_path])
            else:  # Linux
                subprocess.Popen(["xdg-open", folder_path])
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir la carpeta: {str(e)}")
    
    def generate_random_number(self, digits=2):
        """Generate random number with specified number of digits"""
        min_value = 10 ** (digits - 1)
        max_value = (10 ** digits) - 1
        return random.randint(min_value, max_value)
    
    def import_module_from_path(self):
        """Import script.py dynamically"""
        try:
            # First, try to find the script in the executable location
            if getattr(sys, 'frozen', False):
                # If it's a compiled executable
                script_path = os.path.join(sys._MEIPASS, 'script.py')
            else:
                # If running as a normal script
                script_path = os.path.join(os.path.dirname(__file__), 'script.py')
            
            # Check if file exists
            if not os.path.exists(script_path):
                self.add_to_log(f"Error: No se encontró script.py en {script_path}", "error")
                return None
            
            # Import module
            spec = importlib.util.spec_from_file_location("script", script_path)
            script = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(script)
            return script
        except Exception as e:
            self.add_to_log(f"Error importando script: {str(e)}", "error")
            self.add_to_log(traceback.format_exc(), "error")
            return None


def main():
    """Main function"""
    # Configure high DPI theme for Windows
    if platform.system() == "Windows":
        try:
            from ctypes import windll
            windll.shcore.SetProcessDpiAwareness(1)
        except:
            pass
    
    # Start GUI
    root = tk.Tk()
    root.title("ACERO SCRIPT - DODOD SOLUTIONS")
    
    # Configure icon if it exists
    try:
        icon_path = os.path.join(os.path.dirname(__file__), "icon.ico")
        if os.path.exists(icon_path):
            root.iconbitmap(icon_path)
    except:
        pass
    
    # Create application
    app = DXFProcessorApp(root)
    
    # Center window
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')
    
    # Start main loop
    root.mainloop()


if __name__ == "__main__":
    main()