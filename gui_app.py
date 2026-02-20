import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
import sys
import threading
import os
import pandas as pd
import time
import urllib.request
import webbrowser
import ssl
import json
from packaging import version
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from seguridad import cifrar_texto, descifrar_texto
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# --- CONSTANTES GLOBALES ---
SIGAE_URL_GLOBAL = "http://sigae.ucs.gob.ve"
ARCHIVO_RECUPERACION = "pendientes_recuperacion.xlsx"
ARCHIVO_CONFIG = "config_sigae.json"

from sigae_bot import SigaeBot
from generar_notificacion import generar_notificacion_baja_word

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
        self.root.title("Gestor de Bajas y Notificaciones SIGAE v1.0.3")
        self.root.state('zoomed')
        
        # --- CONTROL DE VERSIONES ---
        self.VERSION_ACTUAL = "1.0.3"
        self.URL_VERSION = "https://raw.githubusercontent.com/dbloodmoon/Gestor-de-Bajas-y-Notificaciones-SIGAE/refs/heads/main/version.txt"
        self.URL_DESCARGA = "https://github.com/dbloodmoon/Gestor-de-Bajas-y-Notificaciones-SIGAE/releases/latest"

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
        try:
            contexto_ssl = ssl._create_unverified_context()
            with urllib.request.urlopen(self.URL_VERSION, context=contexto_ssl) as response:
                version_remota = response.read().decode('utf-8').strip()
                
            tupla_remota = tuple(map(int, version_remota.split('.')))
            tupla_actual = tuple(map(int, self.VERSION_ACTUAL.split('.')))
            
            if tupla_remota > tupla_actual:
                print(f"¡Actualización disponible! ({version_remota})")
                self.safe_ui_update(lambda: self.mostrar_aviso_update(version_remota))
            else:
                print("El sistema está actualizado.")
        except Exception as e:
            print(f"No se pudo verificar actualizaciones: {e}")

    def mostrar_aviso_update(self, nueva_version):
        if messagebox.askyesno("Actualización Disponible", 
                               f"Hay una nueva versión disponible ({nueva_version}).\n"
                               f"Tienes la versión {self.VERSION_ACTUAL}.\n\n"
                               "¿Deseas ir a la página de descarga ahora?"):
            webbrowser.open(self.URL_DESCARGA)
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
            self.root.after(1000, self.root.destroy)

    def crear_interfaz(self):
        header_frame = tk.Frame(self.root, bg="#0078d7", height=60)
        header_frame.pack(fill='x')
        tk.Label(header_frame, text="Sistema de Automatización SIGAE", 
                 bg="#0078d7", fg="white", font=("Segoe UI", 16, "bold")).pack(pady=15)

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(expand=True, fill='both', padx=20, pady=15)

        self.tab_login = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_login, text=" 🔐 Acceso ")
        self._construir_login(self.tab_login)

        self.tab_bot = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_bot, text=" 🚀 Ejecutar Bot de Bajas ")
        self._construir_bot(self.tab_bot)

        self.tab_word = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_word, text=" 📄 Generador de Bajas Word ")
        self._construir_word(self.tab_word)

        self.notebook.tab(1, state='disabled') 

        frame_console = ttk.LabelFrame(self.root, text="Registro de Eventos (Log)", padding=10)
        frame_console.pack(fill='both', expand=False, padx=20, pady=(0, 20))
        
        self.console_text = scrolledtext.ScrolledText(
            frame_console, height=12, state='disabled',
            bg="#1e1e1e", fg="#d4d4d4", font=("Consolas", 9), insertbackground="white"
        )
        self.console_text.pack(fill='both')

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
            ops = Options(); ops.add_argument("--headless")
            servicio = Service(ChromeDriverManager().install())
            driver_login = webdriver.Chrome(service=servicio, options=ops)
            driver_login.get(SIGAE_URL_GLOBAL)
            bot = SigaeBot(driver_login)
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
        if not self.sesion_valida:
            self.safe_messagebox("error", "Acceso Denegado", "Debe iniciar sesión primero.")
            return

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
        archivo = self.archivo_excel_word_var.get()
        plantilla = self.plantilla_word_var.get()
        
        if not archivo or not plantilla:
            self.safe_messagebox("error", "Faltan archivos", "Seleccione el Excel y la Plantilla.")
            self.safe_ui_update(lambda: self.btn_run_word.config(state='normal'))
            self.safe_ui_update(lambda: self.btn_stop_word.config(state='disabled'))
            return

        try:
            print("=== INICIANDO GENERADOR WORD ===")
            df = pd.read_excel(archivo)
            df.columns = df.columns.str.strip()
            total = len(df)
            print(f"Registros encontrados: {total}")
            
            cont_ok = 0
            
            for i, row in df.iterrows():
                if self.stop_word_event.is_set():
                    print(f"--- PROCESO INTERRUMPIDO POR USUARIO EN REGISTRO {i} ---")
                    break
                
                try:
                    datos = row.to_dict()
                    cedula = str(datos.get('CÉDULA', 'SN'))
                    datos['cedula'] = cedula
                    
                    print(f"[{i+1}/{total}] Generando doc para: {cedula}...")
                    generar_notificacion_baja_word(datos, plantilla)
                    cont_ok += 1
                    time.sleep(0.05) 
                    
                except Exception as e_row:
                    print(f"Error en fila {i}: {e_row}")

            if not self.stop_word_event.is_set():
                self.safe_messagebox("info", "Proceso terminado", f"Se generaron {cont_ok} documentos.")
                print(f"✓ Finalizado. {cont_ok} documentos creados.")
            else:
                self.safe_messagebox("warning", "Detenido", f"Proceso detenido. Se generaron {cont_ok} documentos.")

        except Exception as e:
            self.safe_messagebox("error", f"Error general: {e}")
            print(f"Error crítico: {e}")
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
                    backup = f"Reportes/backup_descartado_{datetime.now().strftime('%M%S')}.xlsx"
                    os.rename(ARCHIVO_RECUPERACION, backup)
                    print(f"-> Recuperación descartada. Backup movido a {backup}")
                except: pass

        if not archivo_a_usar:
            messagebox.showerror("Error", "Seleccione un archivo de Excel.")
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
        self.driver = None
        resultados = []
        cedulas_procesadas = [] 
        
        try:
            print("=== INICIANDO BOT ===")
            try:
                df = pd.read_excel(archivo, dtype={'CÉDULA': str})
                df.columns = df.columns.str.strip()
                if 'CÉDULA' in df.columns:
                    df = df.dropna(subset=['CÉDULA'])
                    df['CÉDULA'] = df['CÉDULA'].astype(str).str.strip()
                    df = df.drop_duplicates(subset=['CÉDULA'])
            except Exception as e:
                self.safe_messagebox("error", f"Error leyendo Excel: {e}")
                return

            total = len(df)
            print(f"Total registros a procesar: {total}")

            ops = Options()
            ops.add_argument("--start-maximized")
            if headless: ops.add_argument("--headless")
            
            servicio = Service(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=servicio, options=ops)
            bot = SigaeBot(self.driver)
            
            self.driver.get(SIGAE_URL_GLOBAL)
            if not bot.login(self.usuario_var.get(), self.clave_var.get()):
                print("Error de Login. Abortando.")
                return

            for i, row in df.iterrows():
                if self.stop_event.is_set():
                    print("--- PROCESO DETENIDO ---")
                    break
                
                cedula = str(row.get('CÉDULA', 'SN'))
                print(f"\n[{i+1}/{total}] Procesando: {cedula}")
                
                exito = False
                nota = ""
                
                try:
                    if bot.buscar_estudiante(cedula):
                        if bot.solicitar_baja_estudiante(cedula):
                            motivo = row.get('CAUSAL', 'Desconocido')
                            if bot.procesar_formulario_baja(motivo):
                                exito = True
                                nota = "Procesado correctamente"
                                if plantilla and os.path.exists(plantilla):
                                    try:
                                        d_word = row.to_dict()
                                        d_word['cedula'] = cedula
                                        d_word['fecha'] = datetime.now().strftime("%d/%m/%Y")
                                        generar_notificacion_baja_word(d_word, plantilla)
                                    except Exception as ew:
                                        print(f"Error Word: {ew}")
                                        nota = "Baja registrada en SIGAE, pero falló al generar el Word."
                            else:
                                nota = "No se pudo completar el formulario (Verifique si la causal es válida)."
                        else:
                            nota = "El estudiante existe, pero no se pudo hacer click en la opción de baja."
                    else:
                        nota = "Estudiante no encontrado. Verifique la cédula en SIGAE."
                except Exception as e_proc:
                    nota = f"Error Critico: {str(e_proc)[:50]}"
                    print(nota)
                    
                resultado_fila = row.to_dict()

                resultado_fila.update({
                    "ESTADO_BOT": "EXITO" if exito else "FALLO", 
                    "NOTA_SISTEMA": nota,
                    "FECHA_PROCESO": datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                })

                resultados.append(resultado_fila)
                cedulas_procesadas.append(cedula)
                time.sleep(1)

        except Exception as e:
            if "invalid session id" not in str(e).lower() and "chrome not reachable" not in str(e).lower():
                print(f"\nERROR GENERAL DEL HILO: {e}")
                self.safe_messagebox("error", f"Error fatal: {e}")
        
        finally:
            print("\n=== FINALIZANDO Y GUARDANDO ===")
            if self.driver: 
                try: self.driver.quit()
                except: pass
                self.driver = None 

            if resultados:
                try:
                    rep_name = f"Reportes/resultado_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                    pd.DataFrame(resultados).to_excel(rep_name, index=False)
                    print(f"✓ Reporte de sesión guardado: {rep_name}")
                except Exception as e:
                    print(f"Error guardando reporte final: {e}")

            if 'df' in locals():
                try:
                    pendientes = df[~df['CÉDULA'].isin(cedulas_procesadas)]
                    if not pendientes.empty:
                        pendientes.to_excel(ARCHIVO_RECUPERACION, index=False)
                        print(f"⚠ Quedan {len(pendientes)} pendientes. Guardados en: {ARCHIVO_RECUPERACION}")
                        self.safe_messagebox("warning", "Proceso Incompleto", f"Se guardó '{ARCHIVO_RECUPERACION}' con los pendientes.")
                    else:
                        if os.path.exists(ARCHIVO_RECUPERACION):
                            try: os.remove(ARCHIVO_RECUPERACION)
                            except: pass
                            print("✓ Proceso completado totalmente. Archivo de recuperación limpiado.")
                        self.safe_messagebox("info", "Finalizado", "Proceso completado con éxito.")
                except Exception as e:
                    print(f"Error gestionando archivo recuperación: {e}")

            self.safe_ui_update(lambda: self.btn_run_bot.config(state='normal'))
            self.safe_ui_update(lambda: self.btn_stop_bot.config(state='disabled'))

if __name__ == "__main__":
    root = tk.Tk()
    app = SigaeApp(root)
    root.mainloop()
