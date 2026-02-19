# SIGAE Automation Tool 🚀

Herramienta de automatización de procesos administrativos desarrollada en Python. Este software permite gestionar masivamente la baja de estudiantes en la plataforma SIGAE (Sistema de Gestión Académica) y generar automáticamente los soportes documentales en Word.

![Python](https://img.shields.io/badge/Python-3.10+-blue.svg)
![Selenium](https://img.shields.io/badge/Selenium-Automation-green.svg)
![Tkinter](https://img.shields.io/badge/GUI-Tkinter-orange.svg)

## 📋 Características Principales

* **Automatización Web (RPA):** Bot inteligente que navega, busca estudiantes y procesa formularios en el sistema SIGAE usando Selenium WebDriver.
* **Gestión Automática de Drivers:** Integra `webdriver-manager`, por lo que **no es necesario descargar ni configurar ChromeDriver manualmente**; el sistema lo actualiza solo.
* **Interfaz Gráfica (GUI):** Aplicación de escritorio amigable construida con Tkinter, con pestañas de navegación, validación de sesión y consola de logs.
* **Seguridad:** Sistema de cifrado de credenciales locales utilizando `cryptography` (Fernet) para proteger el acceso del usuario.
* **Procesamiento Masivo:** Lectura de datos desde Excel (`pandas`) con capacidad de procesar cientos de registros automáticamente.
* **Generación de Documentos:** Creación automática de cartas de notificación en Word (`python-docx`) rellenando plantillas predefinidas.
* **Resiliencia:** Sistema de auto-recuperación ante fallos de internet o cierres inesperados (guarda el progreso y permite retomar).

## 🛠️ Tecnologías Utilizadas

* **Python 3**
* **Selenium & Webdriver-Manager**
* **Pandas & OpenPyXL**
* **Tkinter**
* **Python-Docx**
* **Cryptography**

## 🚀 Instalación y Uso (Código Fuente)

Si deseas ejecutar el script desde el código fuente en lugar del `.exe`:

1.  **Clonar el repositorio:**
    ```bash
    git clone [https://github.com/TU_USUARIO/sigae-automation-tool.git](https://github.com/TU_USUARIO/sigae-automation-tool.git)
    cd sigae-automation-tool
    ```

2.  **Instalar dependencias:**
    ```bash
    pip install -r requirements.txt
    ```
    *(Asegúrate de que `webdriver-manager` esté en tu requirements.txt)*

3.  **Requisitos:**
    * Solo necesitas tener el archivo `plantilla_bajas.docx` en la carpeta raíz.
    * No necesitas descargar el driver de Chrome, el script lo hará automáticamente al iniciar.

4.  **Ejecución:**
    ```bash
    python gui_app.py
    ```

## 📦 Compilación a Ejecutable (.exe)

Para generar un ejecutable portable que no requiera instalación de Python:

```bash
pyinstaller --noconfirm --onefile --windowed --name "GestorSIGAE" gui_app.py
