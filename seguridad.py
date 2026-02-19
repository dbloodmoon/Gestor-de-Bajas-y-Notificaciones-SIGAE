import os
from cryptography.fernet import Fernet

ARCHIVO_LLAVE = "secret.key"

def obtener_o_crear_llave():
    """Obtiene la llave existente o crea una nueva si no existe."""
    if not os.path.exists(ARCHIVO_LLAVE):
        llave = Fernet.generate_key()
        with open(ARCHIVO_LLAVE, "wb") as archivo_llave:
            archivo_llave.write(llave)
    else:
        with open(ARCHIVO_LLAVE, "rb") as archivo_llave:
            llave = archivo_llave.read()
    return llave

def cifrar_texto(texto):
    """Cifra un texto plano."""
    if not texto: return ""
    f = Fernet(obtener_o_crear_llave())
    return f.encrypt(texto.encode()).decode()

def descifrar_texto(texto_cifrado):
    """Descifra un texto cifrado. Retorna el mismo texto si falla."""
    if not texto_cifrado: return ""
    try:
        f = Fernet(obtener_o_crear_llave())
        return f.decrypt(texto_cifrado.encode()).decode()
    except Exception:
        # Si falla el descifrado (ej. era una contraseña plana antigua), retorna la original
        return texto_cifrado