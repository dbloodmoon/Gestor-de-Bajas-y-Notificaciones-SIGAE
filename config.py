"""Configuración centralizada del proyecto SIGAE."""
import os
from datetime import datetime

# --- Versión (se lee de version.txt para tener una sola fuente de verdad) ---
with open(os.path.join(os.path.dirname(__file__), "version.txt"), "r") as _f:
    VERSION_ACTUAL = _f.read().strip()
APP_NOMBRE = f"Gestor de Bajas y Notificaciones SIGAE v{VERSION_ACTUAL}"

# --- URLs ---
SIGAE_URL = "http://sigae.ucs.gob.ve"
URL_VERSION = "https://raw.githubusercontent.com/dbloodmoon/Gestor-de-Bajas-y-Notificaciones-SIGAE/refs/heads/main/version.txt"
URL_DESCARGA = "https://github.com/dbloodmoon/Gestor-de-Bajas-y-Notificaciones-SIGAE/releases/latest"

# --- Archivos ---
ARCHIVO_RECUPERACION = "pendientes_recuperacion.xlsx"
ARCHIVO_CONFIG = "config_sigae.json"

# --- Meses en español (reutilizable) ---
MESES_ES = {
    1: 'Enero', 2: 'Febrero', 3: 'Marzo', 4: 'Abril',
    5: 'Mayo', 6: 'Junio', 7: 'Julio', 8: 'Agosto',
    9: 'Septiembre', 10: 'Octubre', 11: 'Noviembre', 12: 'Diciembre'
}

def carpeta_con_fecha(base: str) -> str:
    """Genera y crea la ruta base/YYYY/MM - Mes/ según la fecha actual."""
    ahora = datetime.now()
    carpeta = os.path.join(
        base, str(ahora.year),
        f"{ahora.month:02d} - {MESES_ES[ahora.month]}"
    )
    os.makedirs(carpeta, exist_ok=True)
    return carpeta
