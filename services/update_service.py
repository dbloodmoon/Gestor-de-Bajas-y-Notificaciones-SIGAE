"""Servicio de verificación de actualizaciones."""
import ssl
import urllib.request


def verificar_actualizacion(version_actual: str, url_version: str):
    """Compara la versión local con la remota.

    Returns:
        tuple: (hay_update: bool, version_remota: str)
    """
    try:
        contexto_ssl = ssl._create_unverified_context()
        with urllib.request.urlopen(url_version, context=contexto_ssl) as response:
            version_remota = response.read().decode('utf-8').strip()

        tupla_remota = tuple(map(int, version_remota.split('.')))
        tupla_actual = tuple(map(int, version_actual.split('.')))

        if tupla_remota > tupla_actual:
            return True, version_remota
        return False, version_remota
    except Exception as e:
        print(f"No se pudo verificar actualizaciones: {e}")
        return False, None
