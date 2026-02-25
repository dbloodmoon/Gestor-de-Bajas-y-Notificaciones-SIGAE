import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
import sys
import threading
import os
import time
import webbrowser
import json
from datetime import datetime

from config import (
    VERSION_ACTUAL, APP_NOMBRE, URL_VERSION, URL_DESCARGA,
    SIGAE_URL, ARCHIVO_RECUPERACION, ARCHIVO_CONFIG, carpeta_con_fecha
)

# Imports pesados diferidos (se cargan después del splash)
cifrar_texto = None
descifrar_texto = None
AuditorSIGAE = None
plt = None
FigureCanvasTkAgg = None
pd = None

# Servicios (se importan después del splash)
update_service = None
word_service = None
bot_service = None

def _importar_dependencias():
    """Carga los módulos pesados. Se llama después de mostrar el splash."""
    global cifrar_texto, descifrar_texto, AuditorSIGAE
    global plt, FigureCanvasTkAgg, pd
    global update_service, word_service, bot_service

    import pandas
    pd = pandas

    from seguridad import cifrar_texto as _ct, descifrar_texto as _dt
    cifrar_texto = _ct
    descifrar_texto = _dt

    from auditoria import AuditorSIGAE as _Aud
    AuditorSIGAE = _Aud

    import matplotlib.pyplot as _plt
    plt = _plt
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg as _FCA
    FigureCanvasTkAgg = _FCA

    from services import update_service as _us, word_service as _ws, bot_service as _bs
    update_service = _us
    word_service = _ws
    bot_service = _bs

class PrintRedirector:
    """Redirige print() al widget de texto de manera segura para hilos."""
    def __init__(self, text_widget, root):
        self.text_widget = text_widget
        self.root = root
        self.log_file = "registro_proceso.log"
        self.tag_config()

    def tag_config(self):
        try:
            self.text_widget.tag_config("INFO", foreground="#00ff00")
            self.text_widget.tag_config("ERROR", foreground="#ff5555")
            self.text_widget.tag_config("NORMAL", foreground="white")
            self.text_widget.tag_config("WARNING", foreground="orange")
        except:
            pass

    def write(self, string):
        if not string: return
        try:
            if self.root.winfo_exists():
                self.root.after(0, lambda: self._append_text(string))
        except:
            pass
            
        try:
            with open(self.log_file, "a", encoding="utf-8") as f:
                f.write(string)
        except:
            pass 

    def _append_text(self, string):
        try:
            if not self.root.winfo_exists(): return
            self.text_widget.configure(state='normal')
            tag = "NORMAL"
            if "Error" in string or "Fallo" in string or "incorrectas" in string: tag = "ERROR"
            elif "EXITO" in string or "✓" in string or "correctamente" in string: tag = "INFO"
            elif "Interrumpido" in string or "Detenido" in string: tag = "WARNING"
            
            self.text_widget.insert('end', string, tag)
            self.text_widget.see('end')
            self.text_widget.configure(state='disabled')
        except:
            pass

    def flush(self):
        pass

class SigaeApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_NOMBRE)
        self.root.state('zoomed')

        self.is_closing = False
        self.driver = None
        self.sesion_valida = False
               
        self._configurar_estilos()
        
        self.usuario_var = tk.StringVar()
        self.clave_var = tk.StringVar()
        self.archivo_excel_bot_var = tk.StringVar()
        self.archivo_excel_word_var = tk.StringVar()
        self.plantilla_word_var = tk.StringVar(value="plantilla_bajas.docx")
        self.plantilla_bot_var = tk.StringVar(value="plantilla_bajas.docx")
        self.headless_var = tk.BooleanVar(value=False)
        self.tipo_programa_var = tk.StringVar(value="pnf")
        self.archivo_auditoria_var = tk.StringVar()

        # --- RASTREADOR PARA CAMBIAR NOMBRE DE PLANTILLA ---
        self.tipo_programa_var.trace_add("write", self._actualizar_nombres_plantillas)
        
        self.stop_event = threading.Event()
        self.stop_word_event = threading.Event()
        
        self.crear_carpetas()
        self.cargar_credenciales_config()
        
        self.crear_interfaz()
        sys.stdout = PrintRedirector(self.console_text, self.root)
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        self.verificar_actualizacion()

    # --- MANEJO DE CONFIGURACIÓN (JSON) ---
    def cargar_credenciales_config(self):
        """Carga los datos desde el archivo JSON si existe."""
        if os.path.exists(ARCHIVO_CONFIG):
            try:
                with open(ARCHIVO_CONFIG, "r", encoding="utf-8") as f:
                    datos = json.load(f)
                    self.usuario_var.set(datos.get('usuario', ''))
                    # Intentar descifrar la clave
                    clave_cifrada = datos.get('clave', '')
                    if clave_cifrada:
                        self.clave_var.set(descifrar_texto(clave_cifrada))
            except Exception as e:
                print(f"Nota: No se pudo cargar config previa: {e}")

    def guardar_credenciales_config(self):
        """Guarda los datos en un archivo JSON independiente del .exe."""
        usuario = self.usuario_var.get()
        clave_plana = self.clave_var.get()
        clave_cifrada = cifrar_texto(clave_plana)
        
        datos = {
            "usuario": usuario,
            "clave": clave_cifrada
        }
        try:
            with open(ARCHIVO_CONFIG, "w", encoding="utf-8") as f:
                json.dump(datos, f, indent=4)
        except Exception as e:
            self.safe_messagebox("error", f"Error guardando configuración local: {e}")

    def verificar_actualizacion(self):
        print("Buscando actualizaciones...")
        threading.Thread(target=self._thread_verificar_update).start()

    def _thread_verificar_update(self):
        hay_update, version_remota = update_service.verificar_actualizacion(VERSION_ACTUAL, URL_VERSION)
        if hay_update:
            print(f"¡Actualización disponible! ({version_remota})")
            self.safe_ui_update(lambda: self.mostrar_aviso_update(version_remota))
        elif version_remota:
            print("El sistema está actualizado.")

    def mostrar_aviso_update(self, nueva_version):
        if messagebox.askyesno("Actualización Disponible", 
                               f"Hay una nueva versión disponible ({nueva_version}).\n"
                               f"Tienes la versión {VERSION_ACTUAL}.\n\n"
                               "¿Deseas ir a la página de descarga ahora?"):
            webbrowser.open(URL_DESCARGA)
            self.on_closing() 

    def _configurar_estilos(self):
        style = ttk.Style()
        style.theme_use('clam')
        bg_color = "#f0f0f0"
        self.root.configure(bg=bg_color)
        style.configure('TButton', font=('Segoe UI', 10), padding=6)
        style.configure('Action.TButton', font=('Segoe UI', 11, 'bold'), foreground="white", background="#28a745")
        style.map('Action.TButton', background=[('active', '#218838')])
        style.configure('Danger.TButton', font=('Segoe UI', 11, 'bold'), foreground="white", background="#dc3545")
        style.map('Danger.TButton', background=[('active', '#c82333')])
        style.configure('TLabel', background=bg_color, font=('Segoe UI', 10))
        style.configure('Header.TLabel', background=bg_color, font=('Segoe UI', 14, 'bold'), foreground="#333")
        style.configure('TCheckbutton', background=bg_color, font=('Segoe UI', 10))
        style.configure('TFrame', background=bg_color)
        style.configure('TLabelframe', background=bg_color)
        style.configure('TLabelframe.Label', font=('Segoe UI', 10, 'bold'), background=bg_color, foreground="#555")

    def crear_carpetas(self):
        for c in ["Reportes", "Notificaciones"]:
            if not os.path.exists(c): os.makedirs(c)

    def _carpeta_reportes(self):
        """Devuelve la ruta Reportes/YYYY/MM - Mes/ y la crea si no existe."""
        return carpeta_con_fecha("Reportes")

    def _actualizar_nombres_plantillas(self, *args):
        """Cambia el texto de la plantilla por defecto según el programa seleccionado."""
        if self.tipo_programa_var.get() == "pnfa":
            self.plantilla_word_var.set("plantilla_bajas_pnfa.docx")
            self.plantilla_bot_var.set("plantilla_bajas_pnfa.docx")
        else:
            self.plantilla_word_var.set("plantilla_bajas.docx")
            self.plantilla_bot_var.set("plantilla_bajas.docx")

    def safe_messagebox(self, type, title, message=None):
        if self.is_closing: return
        if message is None: message = title
        if type == "info":
            self.root.after(0, lambda: messagebox.showinfo(title, message))
        elif type == "error":
            self.root.after(0, lambda: messagebox.showerror(title, message))
        elif type == "warning":
            self.root.after(0, lambda: messagebox.showwarning(title, message))

    def safe_ui_update(self, func, *args):
        if not self.is_closing:
            self.root.after(0, lambda: func(*args))

    def on_closing(self):
        if messagebox.askokcancel("Salir", "¿Desea salir de la aplicación?\nSi hay un proceso activo, se detendrá y guardará el respaldo."):
            self.is_closing = True
            print("\n=== CERRANDO APLICACIÓN... GUARDANDO DATOS ===")
            self.stop_event.set()
            self.stop_word_event.set()
            
            if self.driver:
                try:
                    self.driver.quit()
                except:
                    pass
            
            def forzar_cierre():
                self.root.destroy()
                os._exit(0)  
                
            self.root.after(1000, forzar_cierre)

    def crear_interfaz(self):
        # 1. Cabecera superior (Header)
        header_frame = tk.Frame(self.root, bg="#0078d7", height=60)
        header_frame.pack(fill='x')
        tk.Label(header_frame, text="Sistema de Automatización SIGAE", 
                 bg="#0078d7", fg="white", font=("Segoe UI", 16, "bold")).pack(pady=15)

        # 2. Contenedor Principal (Cuerpo dividido en 2 columnas)
        main_container = tk.Frame(self.root, bg="#f0f0f0")
        main_container.pack(fill='both', expand=True, padx=20, pady=15)
        
        # Configurar proporciones: Columna 0 (Izquierda) = 66%, Columna 1 (Derecha) = 33%
        main_container.columnconfigure(0, weight=3)
        main_container.columnconfigure(1, weight=1)
        main_container.rowconfigure(0, weight=1)

        # 3. Lado Izquierdo: Pestañas de Control (Notebook)
        self.notebook = ttk.Notebook(main_container)
        self.notebook.grid(row=0, column=0, sticky='nsew', padx=(0, 10))

        self.tab_login = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_login, text=" 🔐 Acceso ")
        self._construir_login(self.tab_login)

        self.tab_bot = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_bot, text=" 🚀 Ejecutar Bot de Bajas ")
        self._construir_bot(self.tab_bot)

        self.tab_word = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_word, text=" 📄 Generador de Bajas Word ")
        self._construir_word(self.tab_word)

        self.tab_auditoria = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_auditoria, text=" 📊 Auditoría ")
        self._construir_auditoria(self.tab_auditoria)

        self.notebook.tab(1, state='disabled') 

        # 4. Lado Derecho: Consola de Registro (Log)
        frame_console = ttk.LabelFrame(main_container, text="Registro de Eventos (Log)", padding=10)
        frame_console.grid(row=0, column=1, sticky='nsew')
        
        # Quitamos el parámetro 'height' para que se expanda verticalmente de forma natural
        self.console_text = scrolledtext.ScrolledText(
            frame_console, state='disabled', width=50,
            bg="#1e1e1e", fg="#d4d4d4", font=("Consolas", 9), insertbackground="white"
        )
        self.console_text.pack(fill='both', expand=True)

    def _construir_login(self, parent):
        container = ttk.Frame(parent)
        container.pack(expand=True, fill='both', padx=50, pady=30)
        card = ttk.LabelFrame(container, text="Credenciales del Sistema", padding=20)
        card.pack(fill='x', pady=10)
        
        ttk.Label(card, text="Nombre de Usuario:").pack(anchor='w', pady=(0, 5))
        ttk.Entry(card, textvariable=self.usuario_var, font=('Segoe UI', 11)).pack(fill='x', pady=(0, 15))
        
        ttk.Label(card, text="Contraseña:").pack(anchor='w', pady=(0, 5))
        ttk.Entry(card, textvariable=self.clave_var, show="•", font=('Segoe UI', 11)).pack(fill='x', pady=(0, 20))
        
        self.btn_login = ttk.Button(card, text="Verificar y Guardar", command=self.verificar_login, style='Action.TButton')
        self.btn_login.pack(fill='x', pady=5)
        self.lbl_status = ttk.Label(card, text="Esperando verificación...", foreground="#666", justify="center")
        self.lbl_status.pack(pady=10)

    def _construir_word(self, parent):
        container = ttk.Frame(parent, padding=20)
        container.pack(fill='both')
        
        lf_files = ttk.LabelFrame(container, text="Selección de Archivos", padding=15)
        lf_files.pack(fill='x', pady=10)
        
        ttk.Label(lf_files, text="📂 Archivo Excel con Datos:").pack(anchor='w')
        f1 = ttk.Frame(lf_files); f1.pack(fill='x', pady=(5, 15))
        ttk.Entry(f1, textvariable=self.archivo_excel_word_var).pack(side='left', fill='x', expand=True)
        ttk.Button(f1, text="Examinar", command=lambda: self.sel_archivo(self.archivo_excel_word_var)).pack(side='right', padx=(5,0))
        
        ttk.Label(lf_files, text="📝 Plantilla Word (.docx):").pack(anchor='w')
        f2 = ttk.Frame(lf_files); f2.pack(fill='x', pady=5)
        ttk.Entry(f2, textvariable=self.plantilla_word_var).pack(side='left', fill='x', expand=True)
        ttk.Button(f2, text="Examinar", command=lambda: self.sel_archivo(self.plantilla_word_var, "*.docx")).pack(side='right', padx=(5,0))
        
        ttk.Label(lf_files, text="🎓 Tipo de Programa:").pack(anchor='w')
        f_prog = ttk.Frame(lf_files); f_prog.pack(fill='x', pady=(0, 10))
        ttk.Radiobutton(f_prog, text="PNF (Pregrado)", variable=self.tipo_programa_var, value="pnf").pack(side='left', padx=(0, 20))
        ttk.Radiobutton(f_prog, text="PNFA (Postgrado)", variable=self.tipo_programa_var, value="pnfa").pack(side='left')

        lf_action = ttk.LabelFrame(container, text="Acciones", padding=15)
        lf_action.pack(fill='x', pady=10)
        
        f_btns = ttk.Frame(lf_action)
        f_btns.pack(fill='x', pady=5)
        
        self.btn_run_word = ttk.Button(f_btns, text="▶ GENERAR DOCUMENTOS", command=self.ejecutar_word, style='Action.TButton')
        self.btn_run_word.pack(side='left', fill='x', expand=True, padx=(0,5))
        
        self.btn_stop_word = ttk.Button(f_btns, text="⏹ DETENER", command=self.detener_word, style='Danger.TButton', state='disabled')
        self.btn_stop_word.pack(side='right', fill='x', expand=True, padx=(5,0))

    def _construir_bot(self, parent):
        container = ttk.Frame(parent, padding=20)
        container.pack(fill='both')
        
        lf_config = ttk.LabelFrame(container, text="Configuración de la Tarea", padding=15)
        lf_config.pack(fill='x', pady=10)
        
        ttk.Label(lf_config, text="📊 Excel de Origen:").pack(anchor='w')
        f_bot = ttk.Frame(lf_config); f_bot.pack(fill='x', pady=(5, 10))
        ttk.Entry(f_bot, textvariable=self.archivo_excel_bot_var).pack(side='left', fill='x', expand=True)
        ttk.Button(f_bot, text="Examinar", command=lambda: self.sel_archivo(self.archivo_excel_bot_var)).pack(side='right', padx=5)

        ttk.Label(lf_config, text="📝 Plantilla (Opcional):").pack(anchor='w')
        f_plb = ttk.Frame(lf_config); f_plb.pack(fill='x', pady=(5, 10))
        ttk.Entry(f_plb, textvariable=self.plantilla_bot_var).pack(side='left', fill='x', expand=True)
        ttk.Button(f_plb, text="Examinar", command=lambda: self.sel_archivo(self.plantilla_bot_var, "*.docx")).pack(side='right', padx=5)

        ttk.Label(lf_config, text="🎓 Tipo de Programa:").pack(anchor='w')
        f_prog = ttk.Frame(lf_config); f_prog.pack(fill='x', pady=(0, 10))
        ttk.Radiobutton(f_prog, text="PNF (Pregrado)", variable=self.tipo_programa_var, value="pnf").pack(side='left', padx=(0, 20))
        ttk.Radiobutton(f_prog, text="PNFA (Postgrado)", variable=self.tipo_programa_var, value="pnfa").pack(side='left')
        ttk.Checkbutton(lf_config, text="Modo Silencioso (Ocultar Navegador)", variable=self.headless_var).pack(anchor='w', pady=5)
        
        lf_control = ttk.Frame(container, padding=10)
        lf_control.pack(fill='x', pady=10)
        
        self.btn_run_bot = ttk.Button(lf_control, text="▶ INICIAR AUTOMATIZACIÓN", command=self.ejecutar_bot, style='Action.TButton')
        self.btn_run_bot.pack(side='left', fill='x', expand=True, padx=(0, 5))
        
        self.btn_stop_bot = ttk.Button(lf_control, text="⏹ DETENER", command=self.detener_bot, style='Danger.TButton', state='disabled')
        self.btn_stop_bot.pack(side='right', fill='x', expand=True, padx=(5, 0))

    def sel_archivo(self, var, tipo="*.xlsx *.xls"):
        path = filedialog.askopenfilename(filetypes=[("Archivos", tipo)])
        if path: var.set(path)

    def verificar_login(self):
        self.btn_login.config(state='disabled')
        self.lbl_status.config(text="⏳ Conectando con SIGAE...", foreground="black")
        threading.Thread(target=self._thread_login).start()

    def _thread_login(self):
        driver_login = None
        try:
            ops = bot_service.Options(); ops.add_argument("--headless")
            servicio = bot_service.Service(bot_service.ChromeDriverManager().install())
            driver_login = bot_service.webdriver.Chrome(service=servicio, options=ops)
            driver_login.get(SIGAE_URL)
            bot = bot_service.SigaeBot(driver_login)
            if bot.login(self.usuario_var.get(), self.clave_var.get()):
                self.safe_ui_update(self.login_exitoso)
            else:
                self.safe_ui_update(lambda: self.lbl_status.config(text="❌ Usuario o clave incorrectos", foreground="red"))
                self.safe_ui_update(lambda: self.btn_login.config(state='normal'))
        except Exception as e:
            self.safe_ui_update(lambda: self.lbl_status.config(text=f"Error: {e}", foreground="red"))
            self.safe_ui_update(lambda: self.btn_login.config(state='normal'))
        finally:
            if driver_login: driver_login.quit()

    def login_exitoso(self):
        self.guardar_credenciales_config()
        self.sesion_valida = True
        self.lbl_status.config(text="✅ Conexión Exitosa", foreground="green")
        
        self.notebook.tab(1, state='normal')
        self.notebook.tab(2, state='normal')
        self.notebook.select(self.tab_bot)
        
        self.safe_messagebox("info", "Acceso Concedido", "Bienvenido. Las funciones han sido desbloqueadas.")

    def ejecutar_word(self):
        self.stop_word_event.clear()
        self.btn_run_word.config(state='disabled')
        self.btn_stop_word.config(state='normal')
        
        self.console_text.configure(state='normal')
        self.console_text.delete(1.0, tk.END)
        self.console_text.configure(state='disabled')

        threading.Thread(target=self._thread_word).start()

    def detener_word(self):
        if messagebox.askyesno("Detener", "¿Desea detener la generación de documentos?"):
            self.stop_word_event.set()
            print("\n!!! DETENIENDO GENERACIÓN WORD... !!!\n")

    def _thread_word(self):
        callbacks = {
            'messagebox': self.safe_messagebox,
            'ui_update': self.safe_ui_update,
        }
        try:
            word_service.generar_words_desde_excel(
                archivo=self.archivo_excel_word_var.get(),
                plantilla=self.plantilla_word_var.get(),
                tipo_programa=self.tipo_programa_var.get(),
                stop_event=self.stop_word_event,
                callbacks=callbacks,
            )
        finally:
            self.safe_ui_update(lambda: self.btn_run_word.config(state='normal'))
            self.safe_ui_update(lambda: self.btn_stop_word.config(state='disabled'))

    def ejecutar_bot(self):
        if not self.sesion_valida:
            self.safe_messagebox("error", "Acceso Denegado", "Debe iniciar sesión correctamente antes de ejecutar el bot.")
            self.notebook.select(self.tab_login)
            return

        archivo_seleccionado = self.archivo_excel_bot_var.get()
        usar_recuperacion = False
        archivo_a_usar = archivo_seleccionado
        
        if os.path.exists(ARCHIVO_RECUPERACION):
            respuesta = messagebox.askyesno("Recuperación Detectada", "Se encontró un proceso anterior interrumpido.\n¿Desea continuar con los pendientes?")
            if respuesta:
                archivo_a_usar = ARCHIVO_RECUPERACION
                usar_recuperacion = True
                print(f"-> Usando archivo de recuperación: {archivo_a_usar}")
            else:
                try:
                    backup = os.path.join(self._carpeta_reportes(), f"backup_descartado_{datetime.now().strftime('%M%S')}.xlsx")
                    os.rename(ARCHIVO_RECUPERACION, backup)
                    print(f"-> Recuperación descartada. Backup movido a {backup}")
                except: pass

        plantilla_actual = self.plantilla_bot_var.get()
        if plantilla_actual and not os.path.exists(plantilla_actual):
            messagebox.showerror("Error", f"La plantilla especificada no existe en la carpeta:\n{plantilla_actual}")
            return

        self.btn_run_bot.config(state='disabled')
        self.btn_stop_bot.config(state='normal')
        self.stop_event.clear()
        self.console_text.configure(state='normal')
        self.console_text.delete(1.0, tk.END)
        self.console_text.configure(state='disabled')
        
        threading.Thread(
            target=self._thread_bot, 
            args=(archivo_a_usar, self.plantilla_bot_var.get(), self.headless_var.get(), usar_recuperacion)
        ).start()

    def detener_bot(self):
        if messagebox.askyesno("Detener", "¿Seguro que desea detener el proceso?\nSe guardará el progreso actual."):
            self.stop_event.set()
            print("\n!!! DETENIENDO PROCESO... Por favor espere a que termine la tarea actual !!!\n")

    def _thread_bot(self, archivo, plantilla, headless, es_recuperacion):
        def set_driver(d):
            self.driver = d

        callbacks = {
            'messagebox': self.safe_messagebox,
            'set_driver': set_driver,
        }
        try:
            bot_service.ejecutar_proceso_bot(
                archivo=archivo,
                plantilla=plantilla,
                headless=headless,
                es_recuperacion=es_recuperacion,
                usuario=self.usuario_var.get(),
                clave=self.clave_var.get(),
                tipo_programa=self.tipo_programa_var.get(),
                stop_event=self.stop_event,
                callbacks=callbacks,
            )
        finally:
            self.safe_ui_update(lambda: self.btn_run_bot.config(state='normal'))
            self.safe_ui_update(lambda: self.btn_stop_bot.config(state='disabled'))

    def _construir_auditoria(self, parent):
        container = ttk.Frame(parent, padding=10)
        container.pack(fill='both', expand=True)

        # 1. Selector de Archivo
        lf_arch = ttk.LabelFrame(container, text="1. Seleccionar Reporte a Auditar", padding=10)
        lf_arch.pack(fill='x', pady=(0, 5))
        
        f_arch = ttk.Frame(lf_arch); f_arch.pack(fill='x', pady=5)
        ttk.Entry(f_arch, textvariable=self.archivo_auditoria_var).pack(side='left', fill='x', expand=True)
        ttk.Button(f_arch, text="Examinar", command=lambda: self.sel_archivo(self.archivo_auditoria_var, "*.xlsx")).pack(side='right', padx=5)

        self.btn_run_auditoria = ttk.Button(lf_arch, text="▶ GENERAR DASHBOARD", command=self.ejecutar_auditoria, style='Action.TButton')
        self.btn_run_auditoria.pack(fill='x', pady=5)

        # 2. Panel de Resultados (Sub-pestañas para Gráficos y Tablas)
        self.notebook_audit = ttk.Notebook(container)
        self.notebook_audit.pack(fill='both', expand=True, pady=5)

        # Pestaña del Gráfico
        self.tab_grafico = ttk.Frame(self.notebook_audit)
        self.notebook_audit.add(self.tab_grafico, text=" 📊 Gráfico de Rendimiento ")

        # Pestaña de Exitosos
        self.tab_exitosos = ttk.Frame(self.notebook_audit)
        self.notebook_audit.add(self.tab_exitosos, text=" ✅ Estudiantes Exitosos ")
        self.tree_exitosos = self.crear_treeview(self.tab_exitosos)

        # Pestaña de Fallidos
        self.tab_fallidos = ttk.Frame(self.notebook_audit)
        self.notebook_audit.add(self.tab_fallidos, text=" ❌ Estudiantes Fallidos ")
        self.tree_fallidos = self.crear_treeview(self.tab_fallidos)

    def crear_treeview(self, parent):
        """Crea una tabla bonita para mostrar estudiantes"""
        columnas = ('Cédula', 'Nombres', 'Nota del Sistema')
        tree = ttk.Treeview(parent, columns=columnas, show='headings', height=6)
        
        tree.heading('Cédula', text='Cédula')
        tree.heading('Nombres', text='Nombres')
        tree.heading('Nota del Sistema', text='Nota del Sistema')
        
        tree.column('Cédula', width=100, anchor='center')
        tree.column('Nombres', width=200)
        tree.column('Nota del Sistema', width=300)
        
        scroll = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scroll.set)
        
        tree.pack(side='left', fill='both', expand=True)
        scroll.pack(side='right', fill='y')
        return tree

    def dibujar_grafico(self, cant_exitos, cant_fallos):
        """Dibuja un gráfico de torta en la interfaz"""
        for widget in self.tab_grafico.winfo_children():
            widget.destroy() # Limpia gráfico anterior
            
        fig, ax = plt.subplots(figsize=(4, 3), facecolor='#f0f0f0')
        if cant_exitos == 0 and cant_fallos == 0:
            return

        etiquetas = ['Exitosos', 'Fallidos']
        valores = [cant_exitos, cant_fallos]
        colores = ['#28a745', '#dc3545'] # Verde y Rojo
        
        ax.pie(valores, labels=etiquetas, colors=colores, autopct='%1.1f%%', startangle=140, 
               textprops={'fontsize': 10, 'weight': 'bold'})
        ax.axis('equal') 
        
        canvas = FigureCanvasTkAgg(fig, master=self.tab_grafico)
        canvas.draw()
        canvas.get_tk_widget().pack(fill='both', expand=True)

    def ejecutar_auditoria(self):
        archivo = self.archivo_auditoria_var.get()
        if not archivo:
            messagebox.showerror("Error", "Debe seleccionar un reporte de la carpeta 'Reportes'.")
            return

        self.console_text.configure(state='normal')
        self.console_text.delete(1.0, tk.END)
        self.console_text.configure(state='disabled')
        
        print("=== GENERANDO DASHBOARD ANALÍTICO ===")
        auditor = AuditorSIGAE()
        
        # Como es lectura de Excel rápida, no necesitamos hilo, lo hacemos directo para evitar crasheos visuales
        exito, datos = auditor.generar_auditoria(archivo)
        
        if exito and datos:
            df_exito = datos['exitosos']
            df_fallo = datos['fallidos']
            
            # Poblar la tabla de Exitosos
            self.tree_exitosos.delete(*self.tree_exitosos.get_children())
            for _, row in df_exito.iterrows():
                nombre = str(row.get('NOMBRES', '')) + " " + str(row.get('APELLIDOS', ''))
                self.tree_exitosos.insert('', 'end', values=(row.get('CÉDULA', ''), nombre.strip(), row.get('NOTA_SISTEMA', '')))
                
            # Poblar la tabla de Fallidos
            self.tree_fallidos.delete(*self.tree_fallidos.get_children())
            for _, row in df_fallo.iterrows():
                nombre = str(row.get('NOMBRES', '')) + " " + str(row.get('APELLIDOS', ''))
                self.tree_fallidos.insert('', 'end', values=(row.get('CÉDULA', ''), nombre.strip(), row.get('NOTA_SISTEMA', '')))
                
            # Dibujar el gráfico!
            self.dibujar_grafico(len(df_exito), len(df_fallo))
            
            self.safe_messagebox("info", "Dashboard Listo", "Gráficos y tablas generadas con éxito.")

if __name__ == "__main__":
    # ── Splash de carga (se muestra al instante) ──
    splash = tk.Tk()
    splash.title("Cargando...")
    splash.overrideredirect(True)
    splash.configure(bg="#f0f0f0")
    sw, sh = 380, 140
    sx = (splash.winfo_screenwidth() - sw) // 2
    sy = (splash.winfo_screenheight() - sh) // 2
    splash.geometry(f"{sw}x{sh}+{sx}+{sy}")
    splash.attributes("-topmost", True)

    header_sp = tk.Frame(splash, bg="#0078d7", height=45)
    header_sp.pack(fill="x")
    header_sp.pack_propagate(False)
    tk.Label(header_sp, text="Gestor de Bajas y Notificaciones SIGAE",
             bg="#0078d7", fg="white", font=("Segoe UI", 11, "bold")).pack(expand=True)

    lbl_sp = tk.Label(splash, text="Cargando módulos, por favor espere...",
                      bg="#f0f0f0", fg="#666", font=("Segoe UI", 9))
    lbl_sp.pack(pady=(15, 8))

    style_sp = ttk.Style(splash)
    style_sp.theme_use("clam")
    style_sp.configure("Splash.Horizontal.TProgressbar",
                       troughcolor="#d0d0d0", background="#0078d7",
                       bordercolor="#ccc", thickness=10)
    pb = ttk.Progressbar(splash, mode="indeterminate",
                         style="Splash.Horizontal.TProgressbar", length=320)
    pb.pack(pady=(0, 10))
    pb.start(15)

    splash.update()

    # Cargar dependencias pesadas en hilo separado
    carga_ok = threading.Event()

    def _cargar():
        _importar_dependencias()
        carga_ok.set()

    threading.Thread(target=_cargar, daemon=True).start()

    # Esperar a que termine, manteniendo el splash vivo
    def _verificar_carga():
        if carga_ok.is_set():
            pb.stop()
            splash.destroy()
        else:
            splash.after(100, _verificar_carga)

    _verificar_carga()
    splash.mainloop()

    # ── Iniciar la aplicación principal ──
    root = tk.Tk()
    app = SigaeApp(root)
    root.mainloop()
