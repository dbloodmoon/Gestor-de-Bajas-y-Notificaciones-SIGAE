"""Servicio de ejecución del bot de bajas SIGAE."""
import os
import time
import pandas as pd
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from sigae_bot import SigaeBot
from generar_notificacion import generar_notificacion_baja_word
from config import SIGAE_URL, ARCHIVO_RECUPERACION, carpeta_con_fecha


def ejecutar_proceso_bot(archivo, plantilla, headless, es_recuperacion,
                         usuario, clave, tipo_programa, stop_event, callbacks):
    """Ejecuta el proceso completo del bot de bajas.

    Args:
        archivo: Ruta al archivo Excel con las cédulas.
        plantilla: Ruta a la plantilla Word (o vacío si no se generan).
        headless: bool, ejecutar Chrome sin ventana.
        es_recuperacion: bool, si usa archivo de recuperación.
        usuario: Nombre de usuario SIGAE.
        clave: Contraseña SIGAE.
        tipo_programa: 'pnf' o 'pnfa'.
        stop_event: threading.Event para detener el proceso.
        callbacks: dict con funciones:
            - messagebox(type, title, message)
            - set_driver(driver)  para que la UI pueda cerrarlo

    Returns:
        dict: {'resultados': list, 'pendientes': int, 'reporte': str}
    """
    driver = None
    resultados = []
    cedulas_procesadas = []
    nombre_hoja = "BAJAS TOTALES" if tipo_programa == "pnf" else "BAJAS PNFA TOTALES"
    reporte_guardado = ""

    try:
        print("=== INICIANDO BOT ===")

        # Leer Excel
        try:
            print(f"    📄 Leyendo hoja: {nombre_hoja}...")
            df = pd.read_excel(archivo, sheet_name=nombre_hoja, dtype={'CÉDULA': str})
            df.columns = df.columns.str.strip()
            if 'CÉDULA' in df.columns:
                df = df.dropna(subset=['CÉDULA'])
                df['CÉDULA'] = df['CÉDULA'].astype(str).str.strip()
                df = df.drop_duplicates(subset=['CÉDULA'])
        except Exception as e:
            callbacks['messagebox']('error', f'Error leyendo Excel: {e}')
            return {'resultados': [], 'pendientes': 0, 'reporte': ''}

        total = len(df)
        print(f"Total registros a procesar: {total}")

        # Iniciar navegador
        ops = Options()
        ops.add_argument("--start-maximized")
        if headless:
            ops.add_argument("--headless")

        servicio = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=servicio, options=ops)
        callbacks['set_driver'](driver)
        bot = SigaeBot(driver)

        # Login
        driver.get(SIGAE_URL)
        if not bot.login(usuario, clave):
            print("Error de Login. Abortando.")
            return {'resultados': [], 'pendientes': total, 'reporte': ''}

        # Procesar cada estudiante
        for i, row in df.iterrows():
            if stop_event.is_set():
                print("--- PROCESO DETENIDO ---")
                break

            cedula = str(row.get('CÉDULA', 'SN'))
            print(f"\n[{i+1}/{total}] Procesando: {cedula}")

            exito = False
            nota = ""

            try:
                if bot.buscar_estudiante(cedula, tipo_programa):
                    if bot.solicitar_baja_estudiante(cedula):
                        motivo = str(row.get('CAUSAL', row.get('MOTIVO', 'Desconocido')))
                        if bot.procesar_formulario_baja(motivo):
                            exito = True
                            nota = "Procesado correctamente"
                            if plantilla and os.path.exists(plantilla):
                                try:
                                    d_word = row.to_dict()
                                    cedula_limpia = cedula
                                    if cedula_limpia.endswith('.0'):
                                        cedula_limpia = cedula_limpia[:-2]
                                    d_word['cedula'] = cedula_limpia
                                    d_word['fecha'] = str(row.get('FECHA', '')).strip()
                                    d_word['causal'] = motivo
                                    d_word['CAUSAL'] = motivo
                                    generar_notificacion_baja_word(d_word, plantilla)
                                except Exception as ew:
                                    print(f"Error Word: {ew}")
                                    nota = "Baja registrada en SIGAE, pero falló al generar el Word."
                        else:
                            nota = "No se pudo completar el formulario"
                    else:
                        nota = "Estudiante no encontrado. Verifique la cédula en SIGAE."
                else:
                    nota = "Estudiante no encontrado. Verifique la cédula en SIGAE."
            except Exception as e_proc:
                nota = f"Error Critico: {str(e_proc)[:50]}"
                print(nota)

            resultado_fila = row.to_dict()
            for columna, valor in resultado_fila.items():
                if pd.notna(valor):
                    if isinstance(valor, (datetime, pd.Timestamp)):
                        resultado_fila[columna] = valor.strftime("%d/%m/%Y")
                    elif str(valor).endswith(" 00:00:00"):
                        resultado_fila[columna] = str(valor).replace(" 00:00:00", "")

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
            callbacks['messagebox']('error', f'Error fatal: {e}')

    finally:
        print("\n=== FINALIZANDO Y GUARDANDO ===")
        if driver:
            try:
                driver.quit()
            except:
                pass
        callbacks['set_driver'](None)

        # Guardar reporte
        if resultados:
            try:
                rep_name = os.path.join(carpeta_con_fecha("Reportes"), f"resultado_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
                pd.DataFrame(resultados).to_excel(rep_name, index=False)
                print(f"✓ Reporte de sesión guardado: {rep_name}")
                reporte_guardado = rep_name
            except Exception as e:
                print(f"Error guardando reporte final: {e}")

        # Gestionar pendientes
        pendientes_count = 0
        if 'df' in locals():
            try:
                pendientes = df[~df['CÉDULA'].isin(cedulas_procesadas)]
                if not pendientes.empty:
                    pendientes_count = len(pendientes)
                    pendientes.to_excel(ARCHIVO_RECUPERACION, index=False, sheet_name=nombre_hoja)
                    print(f"⚠ Quedan {pendientes_count} pendientes. Guardados en: {ARCHIVO_RECUPERACION}")
                    callbacks['messagebox']('warning', 'Proceso Incompleto', f"Se guardó '{ARCHIVO_RECUPERACION}' con los pendientes.")
                else:
                    if os.path.exists(ARCHIVO_RECUPERACION):
                        try:
                            os.remove(ARCHIVO_RECUPERACION)
                        except:
                            pass
                        print("✓ Proceso completado totalmente. Archivo de recuperación limpiado.")
                    callbacks['messagebox']('info', 'Finalizado', 'Proceso completado con éxito.')
            except Exception as e:
                print(f"Error gestionando archivo recuperación: {e}")

    return {'resultados': resultados, 'pendientes': pendientes_count, 'reporte': reporte_guardado}
