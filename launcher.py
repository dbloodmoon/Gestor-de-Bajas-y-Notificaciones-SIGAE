"""
launcher.py — Auto-Updater Launcher para Gestor de Bajas y Notificaciones SIGAE
Requisitos: solo biblioteca estándar de Python + tkinter + packaging
PROHIBIDO importar: selenium, pandas, matplotlib
"""

import os
import sys
import json
import threading
import subprocess
import urllib.request
import urllib.error
import zipfile
import tempfile
import shutil
import tkinter as tk
from tkinter import ttk
from packaging import version

# ─────────────────────────────────────────────
#  CONFIGURACIÓN CENTRAL
# ─────────────────────────────────────────────
APP_NAME        = "Gestor de Bajas y Notificaciones SIGAE"
EXE_NAME        = "Gestor de Bajas y Notificaciones SIGAE.exe"
VERSION_FILE    = "version.txt"
REPO            = "dbloodmoon/Gestor-de-Bajas-y-Notificaciones-SIGAE"

URL_VERSION_REMOTA  = (
    f"https://raw.githubusercontent.com/{REPO}/refs/heads/main/version.txt"
)
URL_API_RELEASE     = (
    f"https://api.github.com/repos/{REPO}/releases/latest"
)

TIMEOUT_HTTP = 10   # segundos


# ─────────────────────────────────────────────
#  HELPERS
# ─────────────────────────────────────────────

def _directorio_base() -> str:
    """Devuelve el directorio donde corre el launcher (compilado o script)."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def leer_version_local() -> str:
    """Lee version.txt local. Devuelve '0.0.0' si no existe."""
    ruta = os.path.join(_directorio_base(), VERSION_FILE)
    try:
        with open(ruta, "r", encoding="utf-8") as f:
            return f.read().strip()
    except FileNotFoundError:
        return "0.0.0"


def guardar_version_local(nueva_version: str) -> None:
    """Sobrescribe version.txt con la nueva versión."""
    ruta = os.path.join(_directorio_base(), VERSION_FILE)
    with open(ruta, "w", encoding="utf-8") as f:
        f.write(nueva_version)


def obtener_version_remota() -> str | None:
    """Descarga y retorna la versión remota, o None si falla."""
    try:
        req = urllib.request.Request(
            URL_VERSION_REMOTA,
            headers={"User-Agent": "SIGAE-Launcher/1.0"},
        )
        with urllib.request.urlopen(req, timeout=TIMEOUT_HTTP) as resp:
            return resp.read().decode("utf-8").strip()
    except Exception:
        return None


def obtener_url_descarga() -> tuple[str, str, str] | tuple[None, None, None]:
    """
    Consulta la API de GitHub para obtener el asset del último release.
    Prioridad: .exe > .zip > primero disponible.
    Retorna (url_descarga, tag_version, nombre_archivo) o (None, None, None) si falla.
    """
    try:
        req = urllib.request.Request(
            URL_API_RELEASE,
            headers={
                "User-Agent": "SIGAE-Launcher/1.0",
                "Accept": "application/vnd.github+json",
            },
        )
        with urllib.request.urlopen(req, timeout=TIMEOUT_HTTP) as resp:
            data = json.loads(resp.read().decode("utf-8"))

        tag = data.get("tag_name", "").lstrip("v")
        assets = data.get("assets", [])

        # 1) Preferir asset .exe directo
        for asset in assets:
            nombre = asset.get("name", "")
            if nombre.endswith(".exe"):
                return asset["browser_download_url"], tag, nombre

        # 2) Aceptar .zip (se extrae el .exe dentro)
        for asset in assets:
            nombre = asset.get("name", "")
            if nombre.endswith(".zip"):
                return asset["browser_download_url"], tag, nombre

        # 3) Primer asset disponible
        if assets:
            nombre = assets[0].get("name", "asset")
            return assets[0]["browser_download_url"], tag, nombre

        return None, None, None

    except Exception:
        return None, None, None


def lanzar_aplicacion() -> None:
    """Lanza el ejecutable principal con subprocess.Popen y cierra el launcher."""
    ruta_exe = os.path.join(_directorio_base(), EXE_NAME)
    if os.path.exists(ruta_exe):
        subprocess.Popen([ruta_exe], cwd=_directorio_base())
    else:
        # Fallback: buscar cualquier .exe en la carpeta que no sea el launcher
        for archivo in os.listdir(_directorio_base()):
            if archivo.endswith(".exe") and "launcher" not in archivo.lower():
                subprocess.Popen(
                    [os.path.join(_directorio_base(), archivo)],
                    cwd=_directorio_base(),
                )
                break


# ─────────────────────────────────────────────
#  INTERFAZ GRÁFICA
# ─────────────────────────────────────────────

class LauncherUI:
    COLOR_BG       = "#1a1a2e"
    COLOR_PANEL    = "#16213e"
    COLOR_ACENTO   = "#0f3460"
    COLOR_PRIMARY  = "#e94560"
    COLOR_TEXTO    = "#eaeaea"
    COLOR_SUBTEXTO = "#a0a0b0"
    FUENTE_TITULO  = ("Segoe UI", 13, "bold")
    FUENTE_ESTADO  = ("Segoe UI", 9)
    ANCHO          = 420
    ALTO           = 180

    def __init__(self, root: tk.Tk):
        self.root = root
        self._construir_ventana()
        self._construir_widgets()

    # ── Construcción ──────────────────────────

    def _construir_ventana(self) -> None:
        self.root.title(APP_NAME)
        self.root.configure(bg=self.COLOR_BG)
        self.root.resizable(False, False)
        # Sin botón de maximizar (Windows)
        self.root.attributes("-toolwindow", False)
        self.root.overrideredirect(False)

        # Centrar en pantalla
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth()  - self.ANCHO) // 2
        y = (self.root.winfo_screenheight() - self.ALTO)  // 2
        self.root.geometry(f"{self.ANCHO}x{self.ALTO}+{x}+{y}")

    def _construir_widgets(self) -> None:
        # Panel principal con padding
        panel = tk.Frame(self.root, bg=self.COLOR_PANEL, bd=0)
        panel.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)

        # Título
        tk.Label(
            panel,
            text=APP_NAME,
            font=self.FUENTE_TITULO,
            bg=self.COLOR_PANEL,
            fg=self.COLOR_PRIMARY,
            pady=12,
        ).pack()

        # Estado
        self._var_estado = tk.StringVar(value="Iniciando...")
        self._lbl_estado = tk.Label(
            panel,
            textvariable=self._var_estado,
            font=self.FUENTE_ESTADO,
            bg=self.COLOR_PANEL,
            fg=self.COLOR_SUBTEXTO,
        )
        self._lbl_estado.pack(pady=(0, 8))

        # Barra de progreso
        style = ttk.Style()
        style.theme_use("clam")
        style.configure(
            "SIGAE.Horizontal.TProgressbar",
            troughcolor=self.COLOR_ACENTO,
            background=self.COLOR_PRIMARY,
            bordercolor=self.COLOR_PANEL,
            lightcolor=self.COLOR_PRIMARY,
            darkcolor=self.COLOR_PRIMARY,
            thickness=10,
        )
        self._progreso_var = tk.DoubleVar(value=0)
        self._barra = ttk.Progressbar(
            panel,
            variable=self._progreso_var,
            maximum=100,
            mode="determinate",
            style="SIGAE.Horizontal.TProgressbar",
            length=360,
        )
        self._barra.pack(pady=(0, 10))

        # Versión local (pie)
        self._var_version = tk.StringVar(value=f"v{leer_version_local()}")
        tk.Label(
            panel,
            textvariable=self._var_version,
            font=("Segoe UI", 7),
            bg=self.COLOR_PANEL,
            fg=self.COLOR_ACENTO,
        ).pack(side=tk.BOTTOM, pady=4)

    # ── API pública (thread-safe vía after) ──

    def set_estado(self, texto: str) -> None:
        self.root.after(0, self._var_estado.set, texto)

    def set_progreso(self, valor: float) -> None:
        """Valor entre 0 y 100."""
        self.root.after(0, self._progreso_var.set, valor)

    def set_modo_indeterminado(self) -> None:
        self.root.after(0, lambda: self._barra.config(mode="indeterminate"))
        self.root.after(0, self._barra.start, 15)

    def set_modo_determinado(self) -> None:
        self.root.after(0, self._barra.stop)
        self.root.after(0, lambda: self._barra.config(mode="determinate"))

    def set_version(self, ver: str) -> None:
        self.root.after(0, self._var_version.set, f"v{ver}")

    def cerrar(self) -> None:
        self.root.after(0, self.root.destroy)


# ─────────────────────────────────────────────
#  LÓGICA PRINCIPAL (corre en hilo secundario)
# ─────────────────────────────────────────────

def flujo_actualizacion(ui: LauncherUI) -> None:
    """
    Hilo principal del launcher:
    1. Verifica versión remota
    2. Descarga si hay actualización
    3. Lanza la aplicación y cierra el launcher
    """
    try:
        # ── 1. Verificar versión ──────────────
        ui.set_modo_indeterminado()
        ui.set_estado("Buscando actualizaciones...")

        version_local  = leer_version_local()
        version_remota = obtener_version_remota()

        ui.set_modo_determinado()
        ui.set_progreso(10)

        if version_remota is None:
            # Sin internet → modo offline
            ui.set_estado("Sin conexión — iniciando en modo offline...")
            ui.set_progreso(100)
            import time; time.sleep(1.5)
            lanzar_aplicacion()
            ui.cerrar()
            return

        hay_actualizacion = version.parse(version_remota) > version.parse(version_local)

        if not hay_actualizacion:
            # Ya es la última versión
            ui.set_estado(f"Ya tienes la última versión ({version_local}) ✓")
            ui.set_progreso(100)
            import time; time.sleep(1.2)
            lanzar_aplicacion()
            ui.cerrar()
            return

        # ── 2. Descargar actualización ────────
        ui.set_estado(f"Descargando versión {version_remota}...")
        ui.set_progreso(15)

        url_descarga, tag_version, nombre_asset = obtener_url_descarga()

        if url_descarga is None:
            ui.set_estado("No se encontró release — iniciando versión actual...")
            ui.set_progreso(100)
            import time; time.sleep(1.5)
            lanzar_aplicacion()
            ui.cerrar()
            return

        ruta_destino = os.path.join(_directorio_base(), EXE_NAME)
        es_zip       = (nombre_asset or "").endswith(".zip")
        extension    = ".zip" if es_zip else ".tmp"
        ruta_tmp     = ruta_destino + extension

        def _reporthook(bloque, tam_bloque, tam_total):
            if tam_total > 0:
                porcentaje = 15 + (bloque * tam_bloque / tam_total) * 80
                ui.set_progreso(min(porcentaje, 95))
                mb_descargado = min(bloque * tam_bloque, tam_total) // (1024 * 1024)
                mb_total      = tam_total // (1024 * 1024)
                ui.set_estado(
                    f"Descargando {version_remota} — {mb_descargado} MB / {mb_total} MB"
                )

        urllib.request.urlretrieve(url_descarga, ruta_tmp, _reporthook)

        # Si es ZIP: extraer TODO el contenido al directorio de la app
        if es_zip:
            ui.set_estado("Descomprimiendo actualización...")
            dir_base = _directorio_base()
            nombre_launcher = os.path.basename(sys.executable).lower()

            with zipfile.ZipFile(ruta_tmp, "r") as zf:
                entradas = zf.infolist()
                total    = len(entradas)

                for i, entrada in enumerate(entradas, start=1):
                    # Nunca sobreescribir el propio launcher
                    nombre_entrada = os.path.basename(entrada.filename).lower()
                    if nombre_entrada and nombre_entrada == nombre_launcher:
                        continue

                    zf.extract(entrada, dir_base)

                    # Actualizar progreso durante la extracción (95→98)
                    progreso_extraccion = 95 + (i / total) * 3
                    ui.set_progreso(min(progreso_extraccion, 98))

            os.remove(ruta_tmp)
        else:
            # Asset directo .exe: reemplazar
            if os.path.exists(ruta_destino):
                os.replace(ruta_tmp, ruta_destino)
            else:
                os.rename(ruta_tmp, ruta_destino)

        # Actualizar version.txt local
        nueva_ver = tag_version or version_remota
        guardar_version_local(nueva_ver)
        ui.set_version(nueva_ver)

        # ── 3. Lanzar ─────────────────────────
        ui.set_estado(f"¡Actualizado a v{nueva_ver}! Iniciando...")
        ui.set_progreso(100)

        import time; time.sleep(1.0)
        lanzar_aplicacion()
        ui.cerrar()

    except Exception as exc:
        # Error inesperado → modo offline con mensaje
        try:
            ui.set_modo_determinado()
        except Exception:
            pass
        ui.set_estado(f"Error: {exc} — iniciando de todos modos...")
        ui.set_progreso(100)
        import time; time.sleep(2.0)
        lanzar_aplicacion()
        ui.cerrar()


# ─────────────────────────────────────────────
#  PUNTO DE ENTRADA
# ─────────────────────────────────────────────

def main() -> None:
    root = tk.Tk()
    ui   = LauncherUI(root)

    # Lanzar el flujo en un hilo daemon para no bloquear el mainloop
    hilo = threading.Thread(
        target=flujo_actualizacion,
        args=(ui,),
        daemon=True,
        name="hilo-actualizacion",
    )
    hilo.start()

    root.mainloop()


if __name__ == "__main__":
    main()
