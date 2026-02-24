import pandas as pd
import os
from datetime import datetime

class AuditorSIGAE:
    def __init__(self):
        # 1. Nueva carpeta exclusiva para organización
        self.carpeta_salida = "Auditorias"
        if not os.path.exists(self.carpeta_salida):
            os.makedirs(self.carpeta_salida)

    def generar_auditoria(self, archivo_reporte):
        if not archivo_reporte or not os.path.exists(archivo_reporte):
            print(f"❌ Archivo no encontrado: {archivo_reporte}")
            return False, None

        try:
            print(f"📄 Analizando reporte: {os.path.basename(archivo_reporte)}")
            df = pd.read_excel(archivo_reporte, dtype={'CÉDULA': str})
            
            if 'ESTADO_BOT' not in df.columns:
                print("❌ El archivo no tiene el formato correcto (Falta ESTADO_BOT).")
                return False, None

            # --- 1. CLASIFICACIÓN Y CÁLCULOS ---
            exitosos = df[df['ESTADO_BOT'] == 'EXITO']
            fallidos = df[df['ESTADO_BOT'] == 'FALLO']
            
            total_proc = len(df)
            tasa_exito = f"{(len(exitosos) / total_proc * 100):.1f}%" if total_proc > 0 else "0%"
            fecha_audit = datetime.now().strftime("%d/%m/%Y %I:%M %p")
            nombre_origen = os.path.basename(archivo_reporte)

            resumen_errores = pd.DataFrame()
            if 'NOTA_SISTEMA' in fallidos.columns and not fallidos.empty:
                resumen_errores = fallidos['NOTA_SISTEMA'].value_counts().reset_index()
                resumen_errores.columns = ['Motivo del Fallo', 'Cantidad']

            # --- 2. GENERACIÓN DEL NOMBRE DE SALIDA ---
            nombre_base = nombre_origen.replace("resultado_", "Auditoria_")
            if not nombre_base.startswith("Auditoria_"):
                nombre_base = f"Auditoria_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            
            ruta_salida = os.path.join(self.carpeta_salida, nombre_base)
            
            # --- 3. CREACIÓN DEL EXCEL ENRIQUECIDO ---
            with pd.ExcelWriter(ruta_salida, engine='openpyxl') as writer:
                
                # Pestaña 1: Resumen General (Ahora con muchos más datos)
                pd.DataFrame({
                    'Métrica de Auditoría': [
                        'Documento Origen',
                        'Fecha de Evaluación',
                        'Total de Estudiantes Procesados', 
                        'Bajas Ejecutadas con Éxito', 
                        'Bajas Fallidas / No Encontradas',
                        'Tasa de Efectividad del Sistema'
                    ],
                    'Valor': [
                        nombre_origen,
                        fecha_audit,
                        total_proc, 
                        len(exitosos), 
                        len(fallidos),
                        tasa_exito
                    ]
                }).to_excel(writer, sheet_name='Resumen General', index=False)

                # Pestaña 2: Listado de Exitosos (Para el archivo de la secretaría)
                if not exitosos.empty:
                    cols_exito = [c for c in ['CÉDULA', 'NOMBRES', 'APELLIDOS', 'NOTA_SISTEMA'] if c in exitosos.columns]
                    exitosos[cols_exito].to_excel(writer, sheet_name='Procesados con Éxito', index=False)

                # Pestaña 3: Listado de Fallos (Para revisión manual)
                if not fallidos.empty:
                    cols_fallo = [c for c in ['CÉDULA', 'NOMBRES', 'APELLIDOS', 'NOTA_SISTEMA'] if c in fallidos.columns]
                    fallidos[cols_fallo].to_excel(writer, sheet_name='Requieren Revisión', index=False)
                
                # Pestaña 4: Agrupación de errores
                if not resumen_errores.empty:
                    resumen_errores.to_excel(writer, sheet_name='Desglose Errores', index=False)

            print(f"💾 Auditoría exportada en carpeta: {self.carpeta_salida}")
            
            datos = {'exitosos': exitosos, 'fallidos': fallidos}
            return True, datos

        except Exception as e:
            print(f"❌ Error al generar la auditoría: {e}")
            return False, None