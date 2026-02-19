# SIGAE Automation Tool 🚀

Herramienta de automatización de procesos administrativos desarrollada en Python. Este software permite gestionar masivamente la baja de estudiantes en la plataforma SIGAE (Sistema de Gestión Académica) y generar automáticamente los soportes documentales en Word.

![Python](https://img.shields.io/badge/Python-3.10+-blue.svg)
![Selenium](https://img.shields.io/badge/Selenium-Automation-green.svg)
![Tkinter](https://img.shields.io/badge/GUI-Tkinter-orange.svg)

## 📋 Características Principales

* **Automatización Web (RPA):** Bot inteligente que navega, busca estudiantes y procesa formularios en el sistema SIGAE usando Selenium WebDriver.
* **Interfaz Gráfica (GUI):** Aplicación de escritorio amigable construida con Tkinter, con pestañas de navegación y consola de logs en tiempo real.
* **Seguridad:** Sistema de cifrado de credenciales locales utilizando `cryptography` (Fernet) para proteger el acceso del usuario.
* **Procesamiento Masivo:** Lectura de datos desde Excel (`pandas`) con capacidad de procesar cientos de registros automáticamente.
* **Generación de Documentos:** Creación automática de cartas de notificación en Word (`python-docx`) rellenando plantillas predefinidas.
* **Resiliencia:** Sistema de auto-recuperación ante fallos de internet o cierres inesperados (guarda el progreso y permite retomar).

## 🛠️ Tecnologías Utilizadas

* **Python 3**
* **Selenium:** Para la automatización del navegador.
* **Pandas & OpenPyXL:** Para manipulación de datos Excel.
* **Tkinter:** Para la interfaz gráfica de usuario.
* **Python-Docx:** Para la generación de reportes.
* **Cryptography:** Para el manejo seguro de contraseñas.
* **Threading:** Para evitar el congelamiento de la interfaz durante procesos largos.

## 🚀 Instalación y Uso

1.  **Clonar el repositorio:**
    ```bash
    git clone [https://github.com/TU_USUARIO/sigae-automation-tool.git](https://github.com/TU_USUARIO/sigae-automation-tool.git)
    cd sigae-automation-tool
    ```

2.  **Instalar dependencias:**
    ```bash
    pip install -r requirements.txt
    ```

3.  **Configuración:**
    * Asegúrate de tener `chromedriver.exe` en la carpeta raíz (o usa `webdriver-manager`).
    * Debes tener el archivo `plantilla_bajas.docx` en la carpeta.

4.  **Ejecución:**
    ```bash
    python gui_app.py
    ```

## ⚠️ Nota Legal y Responsabilidad

Esta herramienta fue desarrollada con fines de optimización administrativa y educativa. El uso de bots en plataformas de terceros debe realizarse bajo la supervisión y autorización correspondiente. El autor no se hace responsable por el mal uso de la herramienta.

## 📄 Licencia

Este proyecto está bajo la Licencia MIT - ver el archivo [LICENSE](LICENSE) para más detalles.
