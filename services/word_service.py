"""Servicio de generación masiva de documentos Word."""
import os
import time
import pandas as pd
from generar_notificacion import generar_notificacion_baja_word


def generar_words_desde_excel(archivo, plantilla, tipo_programa, stop_event, callbacks):
    """Genera documentos Word a partir de un archivo Excel.

    Args:
        archivo: Ruta al archivo Excel.
        plantilla: Ruta a la plantilla Word.
        tipo_programa: 'pnf' o 'pnfa'.
        stop_event: threading.Event para detener el proceso.
        callbacks: dict con funciones de la UI:
            - messagebox(type, title, message)
            - ui_update(func)

    Returns:
        tuple: (documentos_creados: int, fue_detenido: bool)
    """
    if not os.path.exists(plantilla):
        callbacks['messagebox']('error', 'Plantilla no encontrada', f'No se encontró el archivo:\n{plantilla}')
        return 0, False

    if not os.path.exists(archivo):
        callbacks['messagebox']('error', 'Excel no encontrado', f'No se encontró el archivo:\n{archivo}')
        return 0, False

    nombre_hoja = "BAJAS TOTALES" if tipo_programa == "pnf" else "BAJAS PNFA TOTALES"

    try:
        print("=== INICIANDO GENERADOR WORD ===")
        try:
            print(f"    📄 Leyendo hoja: {nombre_hoja}...")
            df = pd.read_excel(archivo, sheet_name=nombre_hoja, dtype={'CÉDULA': str})
        except Exception:
            callbacks['messagebox']('error', 'Error leyendo Excel', f"No se encontró la pestaña '{nombre_hoja}'.")
            return 0, False

        df.columns = df.columns.str.strip()
        total = len(df)
        print(f"Registros encontrados: {total}")

        cont_ok = 0

        for i, row in df.iterrows():
            if stop_event.is_set():
                print(f"--- PROCESO INTERRUMPIDO POR USUARIO EN REGISTRO {i} ---")
                break

            try:
                datos = row.to_dict()
                cedula = str(datos.get('CÉDULA', 'SN'))
                if cedula.endswith('.0'):
                    cedula = cedula[:-2]
                datos['cedula'] = cedula

                causal = str(datos.get('CAUSAL', datos.get('MOTIVO', 'Desconocido')))
                if causal.lower() == 'nan':
                    causal = 'DESINCORPORACION POR MOTIVOS PERSONALES'
                datos['causal'] = causal
                datos['CAUSAL'] = causal

                print(f"[{i+1}/{total}] Generando doc para: {cedula}...")
                generar_notificacion_baja_word(datos, plantilla)
                cont_ok += 1
                time.sleep(0.05)

            except Exception as e_row:
                print(f"Error en fila {i}: {e_row}")

        fue_detenido = stop_event.is_set()

        if not fue_detenido:
            callbacks['messagebox']('info', 'Proceso terminado', f'Se generaron {cont_ok} documentos.')
            print(f"✓ Finalizado. {cont_ok} documentos creados.")
        else:
            callbacks['messagebox']('warning', 'Detenido', f'Proceso detenido. Se generaron {cont_ok} documentos.')

        return cont_ok, fue_detenido

    except Exception as e:
        callbacks['messagebox']('error', f'Error general: {e}')
        print(f"Error crítico: {e}")
        return 0, False
