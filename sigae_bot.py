from datetime import datetime
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException
import pandas as pd
import time
class SigaeBot:
    """Clase para automatizar procesos en el sistema SIGAE."""
    
    # --- CONSTANTES: LOCALIZADORES DE ELEMENTOS ---
    INPUT_USUARIO = (By.ID, "loginform-username")
    INPUT_CLAVE = (By.ID, "loginform-password")
    MENU_ESTUDIANTE = (By.CSS_SELECTOR, 'a[href="#estudiante"]')
    SUBMENU_LISTA_PNF = (By.XPATH, "//a[contains(@href, 'alumno-pnf')][.//span[contains(text(), 'Lista PNF')]]")
    SELECT_NACIONALIDAD = (By.NAME, "AlumnoSearch[nacionalidad]")
    INPUT_CEDULA = (By.NAME, "AlumnoSearch[cedula]")
    ELEMENTO_VACIO = (By.CSS_SELECTOR, ".empty")
    OPCION_SOLICITAR_BAJA = (By.CSS_SELECTOR, 'a[href*="solicitar-baja"]')
    SELECT_MOTIVO = (By.ID, "alumnobajaslicencias-id_estatus_academico")
    TEXTAREA_DESCRIPCION = (By.ID, "alumnobajaslicencias-descripcion_solicitud")
    BOTON_ENVIAR = (By.ID, "button-submit-inscripcion")
    
    # URL principal (debe estar en config.py)
    URL_PRINCIPAL = "http://sigae.ucs.gob.ve"

    def __init__(self, driver):
        """Inicializa la instancia con el driver de Selenium."""
        self.driver = driver
        self.wait = WebDriverWait(self.driver, 15)
        self._inicializar_mapeo_causales()
        self.tipo_prog = ""

    def _inicializar_mapeo_causales(self):
        """Inicializa el diccionario de mapeo de causales de baja."""
        # NOTA IMPORTANTE: Los valores deben coincidir EXACTAMENTE con los del HTML
        self.MAPEO_CAUSALES = {
            "SUSPENSION POR SOLICITUD PERSONAL": "5",  # "Solicitud personal por escrito"
            "SUSPENSION POR DESERCION": "6",           # "Deserción"
            "INSUFICIENCIA ACADÉMICA": "3",            # "Insuficiencia académica durante el proceso docente-educativo"
            "SUSPENSION TEMPORAL POR INASISTENCIA": "2",  # "Inasistencia"
            "APLICACIÓN DE MEDIDAS DISCIPLINARIAS": "4",  # "Aplicación de medidas disciplinarias contenidas en el Reglamento Disciplinario de la UCS-HCF"
            "BAJA DEFINITIVA": "9",                   # "Baja definitiva"
            "FALLECIMIENTO": "7",                     # "Fallecimiento"
            "PÉRDIDA DE REQUISITO": "8",              # "Pérdida de requisito"
            "INSUFICIENCIA ACADEMICA": "3",           # Variante sin tilde
            "PERDIDA DE REQUISITO": "8"               # Variante sin tilde
        }

    # --- MÉTODOS BÁSICOS ---
    def esperar_elemento(self, localizador, timeout=15, mensaje_error="Elemento no encontrado"):
        """Espera a que un elemento esté presente y visible."""
        try:
            return WebDriverWait(self.driver, timeout).until(
                EC.visibility_of_element_located(localizador)
            )
        except TimeoutException:
            print(f"    ⏱️  Timeout esperando: {mensaje_error}")
            return None
        
    def esperar_presencia_elemento(self, localizador, timeout=15, mensaje_error="Elemento no encontrado"):
        """Espera a que un elemento exista en el DOM, sin importar si es visible o no."""
        try:
            return WebDriverWait(self.driver, timeout).until(
                EC.presence_of_element_located(localizador)
            )
        except TimeoutException:
            print(f"    ⏱️  Timeout esperando presencia: {mensaje_error}")
            return None

    def escribir_en_campo(self, localizador, texto, timeout=10):
        """Escribe texto en un campo de formulario visible."""
        try:
            campo = self.esperar_elemento(localizador, timeout, f"Campo {localizador}")
            if campo:
                campo.clear()
                campo.send_keys(str(texto))
                return True
            return False
        except Exception as e:
            print(f"    ⚠ Error al escribir en campo {localizador}: {e}")
            return False

    def hacer_click(self, localizador, timeout=10, usar_javascript=False):
        """Hace click en un elemento."""
        try:
            elemento = self.esperar_elemento(localizador, timeout, f"Elemento para click {localizador}")
            if elemento:
                if usar_javascript:
                    self.driver.execute_script("arguments[0].click();", elemento)
                else:
                    elemento.click()
                return True
            return False
        except Exception as e:
            print(f"    ⚠ Error al hacer click en {localizador}: {e}")
            return False
        
    def esperar_url_contenga(self, texto_url, timeout=15):
        """Espera dinámicamente hasta que la URL cambie al texto esperado."""
        try:
            WebDriverWait(self.driver, timeout).until(
                EC.url_contains(texto_url)
            )
            return True
        except TimeoutException:
            return False

    def esperar_desaparicion(self, localizador, timeout=10):
        """Espera dinámicamente a que un elemento (ej. pantalla de carga) desaparezca."""
        try:
            WebDriverWait(self.driver, timeout).until(
                EC.invisibility_of_element_located(localizador)
            )
            return True
        except TimeoutException:
            return False

    def obtener_id_causal(self, texto_causal):
        """Convierte texto descriptivo al ID numérico usado por el sistema."""
        if pd.isna(texto_causal):
            return "5"
        
        texto_normalizado = str(texto_causal).strip().upper()
        
        # Intentar coincidencia exacta primero con el diccionario original
        id_exacto = self.MAPEO_CAUSALES.get(texto_normalizado)
        if id_exacto:
            return id_exacto
        
        # Búsqueda inteligente por palabras clave si escribieron mal en el Excel
        if "DESERCI" in texto_normalizado: 
            return "6"
        if "INASISTENCIA" in texto_normalizado: 
            return "2"
        if "INSUFICIENCIA" in texto_normalizado: 
            return "3"
        if "DISCIPLINARIA" in texto_normalizado: 
            return "4"
        if "DEFINITIVA" in texto_normalizado: 
            return "9"
        if "FALLECIMIENTO" in texto_normalizado: 
            return "7"
        if "REQUISITO" in texto_normalizado: 
            return "8"
        if "PERSONAL" in texto_normalizado or "VOLUNTARIA" in texto_normalizado: 
            return "5"
        
        print(f"    ⚠ Causal no reconocida en Excel: '{texto_causal}'. Usando por defecto (5).")
        return "5"

    # --- AUTENTICACIÓN ---
    def login(self, usuario, clave):
        """Inicia sesión en el sistema SIGAE con las credenciales proporcionadas."""
        try:
            print("    ↻ Iniciando sesión...")
            
            # Intentar escribir usuario
            if not self.escribir_en_campo(self.INPUT_USUARIO, usuario):
                print("    ✗ No se pudo escribir usuario")
                return False
            
            # Intentar escribir clave
            if not self.escribir_en_campo(self.INPUT_CLAVE, clave):
                print("    ✗ No se pudo escribir clave")
                return False
            
            # Presionar Enter para iniciar sesión
            campo_clave = self.driver.find_element(*self.INPUT_CLAVE)
            campo_clave.send_keys(Keys.RETURN)
            
            elementos_login = self.driver.find_elements(*self.INPUT_USUARIO)
            
            try:
                WebDriverWait(self.driver, 5).until(
                    EC.staleness_of(campo_clave)
                )
            except TimeoutException:
                pass # Si no se vuelve obsoleto, revisamos si hay error visible
            
            elementos_login = self.driver.find_elements(*self.INPUT_USUARIO)
            if len(elementos_login) > 0 and elementos_login[0].is_displayed():
                print("    ✗ Credenciales incorrectas.")
                return False
            
            print("    ✓ Sesión iniciada")
            return True
                
        except Exception as e:
            print(f"Error en login: {e}")
            return False

    # --- NAVEGACIÓN ---
    def navegar_a_listado(self, tipo_programa="pnf"):
        """Navega directamente a la lista de estudiantes PNF mediante URL."""
        tipo = str(tipo_programa).strip().lower()
        self.tipo_prog = tipo

        print("    ↻ Navegando directamente al listado PNF...")
        try:
            # Construir la URL exacta usando la base y la ruta que vimos en el HTML
            url_lista = f"{self.URL_PRINCIPAL}/index.php?r=estudiante%2Falumno-{tipo}"
            
            # Navegar directo, sin hacer clics en menús
            self.driver.get(url_lista)
            
            # Verificar que llegamos correctamente
            if self.esperar_url_contenga(f"alumno-{tipo}"):
                print("    ✓ Listado PNF cargado instantáneamente")
                return True
            else:
                print("    ✗ Error al cargar la URL directa del listado")
                return False
                
        except Exception as e:
            print(f"Error al navegar al listado PNF: {e}")
            return False

    # --- BÚSQUEDA ---
    def buscar_estudiante(self, cedula, tipo_programa="pnf", nacionalidad=""):
        """Busca un estudiante por cédula en el sistema."""
        tipo = str(tipo_programa).strip().lower()

        try:
            print(f"    🔍 Buscando estudiante {cedula}...")
            
            # Verificar si ya estamos en el listado correcto
            if f"alumno-{tipo}" not in self.driver.current_url:
                print("    ↻ No estamos en listado PNF, navegando...")
                if not self.navegar_a_listado(tipo): 
                    print("    ✗ No se pudo navegar al listado PNF")
                    return False
            
            # Limpiar filtros previos si existen
            try:
                campo_cedula = self.driver.find_element(*self.INPUT_CEDULA)
                if campo_cedula.get_attribute("value"):
                    campo_cedula.clear()
                    # Presionar Enter para limpiar búsqueda
                    campo_cedula.send_keys(Keys.RETURN)
                    time.sleep(1.5)
            except:
                pass
            
            # Configurar nacionalidad si es necesario
            try:
                selector_nacionalidad = self.driver.find_element(*self.SELECT_NACIONALIDAD)
                if selector_nacionalidad.get_attribute("value") != nacionalidad:
                    Select(selector_nacionalidad).select_by_value(nacionalidad)
                    time.sleep(0.5)
            except:
                print("    ⚠ No se pudo configurar nacionalidad, continuando...")
            
            # Escribir la cédula y buscar
            if not self.escribir_en_campo(self.INPUT_CEDULA, cedula):
                return False
            
            # Presionar Enter para buscar
            campo_cedula = self.driver.find_element(*self.INPUT_CEDULA)
            campo_cedula.send_keys(Keys.RETURN)
            
            # Esperar resultados
            time.sleep(2.5)
            
            # Verificar si hay resultados
            try:
                # Primero intentar encontrar el mensaje de "no hay resultados"
                elementos_vacios = self.driver.find_elements(*self.ELEMENTO_VACIO)
                if elementos_vacios:
                    for elemento in elementos_vacios:
                        if "no hay" in elemento.text.lower() or "vacio" in elemento.text.lower() or "empty" in elemento.text.lower():
                            print(f"    ✗ No hay resultados para {cedula}")
                            return False
                
                # Si no hay mensaje de vacío, verificar si hay una tabla con resultados
                # Buscar cualquier fila de tabla que no sea el encabezado
                filas_tabla = self.driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
                if len(filas_tabla) > 0:
                    print(f"    ✓ Estudiante {cedula} encontrado")
                    time.sleep(1)  # Pequeña pausa para estabilizar
                    return True
                else:
                    print(f"    ✗ Tabla vacía para {cedula}")
                    return False
                    
            except Exception as e:
                print(f"    ⚠ Error verificando resultados: {e}")
                # Si no podemos determinar, asumir que sí hay resultados
                print(f"    ⚠ Asumiendo que {cedula} fue encontrado")
                return True
                
        except Exception as e:
            print(f"Error al buscar estudiante {cedula}: {e}")
            return False

    # --- SOLICITUD DE BAJA ---
    def solicitar_baja_estudiante(self, cedula):
        """Abre el formulario de solicitud de baja para un estudiante específico."""
        try:
            print(f"    📝 Abriendo formulario para {cedula}...")
            
            # Construir XPath para encontrar la fila del estudiante
            xpath_fila = f"//tr[td[contains(text(), '{cedula}')]]"
            
            # Esperar a que la fila sea visible
            fila_estudiante = self.esperar_elemento((By.XPATH, xpath_fila),
                                                    timeout=1,
                                                    mensaje_error=f"Fila del estudiante {cedula}")
            if not fila_estudiante:
                print(f"    ✗ No se encontró la fila para {cedula}")
                return False

            try:
                opcion_baja_directa = fila_estudiante.find_element(By.CSS_SELECTOR, 'a[href*="solicitar-baja"]')
                if opcion_baja_directa.is_displayed():
                    print("    ↻ Clic en botón directo de baja (Sin menú)...")
                    self.driver.execute_script("arguments[0].click();", opcion_baja_directa)
                    time.sleep(0.5)
                    print(f"    ✓ Formulario abierto para {cedula}")
                    self.tipo_prog = "pnfa"
                    return True
            except:
                self.tipo_prog = "pnf"
                pass
            
            # Intentar diferentes selectores para el botón del menú
            selectores_menu = [
                "a.dropdown-toggle",
                "button.dropdown-toggle", 
                ".btn-group > button",
                ".dropdown > button",
                "a[data-toggle='dropdown']",
                "button[data-toggle='dropdown']",
                ".btn.dropdown-toggle"
            ]
            
            boton_menu = None
            for selector in selectores_menu:
                try:
                    elementos = fila_estudiante.find_elements(By.CSS_SELECTOR, selector)
                    for elemento in elementos:
                        if elemento.is_displayed():
                            boton_menu = elemento
                            break
                    if boton_menu:
                        break
                except:
                    continue
            
            if not boton_menu:
                print(f"    ✗ No se encontró el botón del menú para {cedula}")
                return False
            
            # Hacer click en el botón del menú
            print("    ↻ Abriendo menú desplegable...")
            self.driver.execute_script("arguments[0].click();", boton_menu)
            time.sleep(0.5)  # Esperar a que se abra el menú
            
            # Buscar y hacer click en la opción de baja
            print("    ↻ Seleccionando opción de baja...")
            opcion_baja = self.esperar_elemento(self.OPCION_SOLICITAR_BAJA, 
                                               mensaje_error="Opción 'Solicitar baja'")
            if not opcion_baja:
                print("    ✗ No se encontró la opción de baja")
                return False
            
            self.driver.execute_script("arguments[0].click();", opcion_baja)
            
            # Esperar a que cargue el formulario
            time.sleep(0.5)
            
            print(f"    ✓ Formulario abierto para {cedula}")
            return True
            
        except Exception as error:
            print(f"Error al abrir menú para {cedula}: {error}")
            return False

    # --- PROCESAMIENTO DE FORMULARIO ---
    def procesar_formulario_baja(self, causal_texto):
        """Completa y envía el formulario de baja con el motivo especificado."""

        try:
            print(f"    ✍️  Procesando formulario...")
            
            print("    ↻ Esperando carga del formulario...")
            
            if not self.esperar_presencia_elemento(self.SELECT_MOTIVO, timeout=5, 
                                         mensaje_error="Campo de motivo en formulario"):
                print("    ⚠ Formulario no se cargó completamente, pero continuamos...")
                        
            # Completar los campos del formulario en orden
            print("    ↻ Seleccionando motivo...")
            self._seleccionar_motivo_select2(causal_texto)
            time.sleep(0.25)  
            
            print("    ↻ Estableciendo fecha...")
            self._establecer_fecha_actual()
            time.sleep(0.25)
            
            print("    ↻ Escribiendo descripción...")
            self._escribir_descripcion(causal_texto)
            time.sleep(0.25)
            
            # Enviar el formulario
            print("    ↻ Enviando formulario...")
            
            # Guardamos la URL actual antes de hacer clic
            url_formulario = self.driver.current_url
            
            if self._enviar_formulario():
                print(f"    ✓ Formulario enviado: {causal_texto}")
                
                print("    ⏳ Verificando redirección del sistema...")
                try:
                    WebDriverWait(self.driver, 10).until(
                        EC.url_changes(url_formulario)
                    )
                except TimeoutException:
                    print("    ⚠ La URL no cambió rápido, pero forzaremos la salida.")

                print("    ↻ Evadiendo pop-up visual y volviendo al inicio...")
                
                url_lista_pnf = f"{self.URL_PRINCIPAL}/index.php?r=estudiante%2Falumno-{self.tipo_prog}"
                self.driver.get(url_lista_pnf)

                return True
            else:
                print(f"    ✗ Problema al enviar formulario")
                return False

        except Exception as error:
            print(f"Error al procesar formulario: {error}")
            return False

    def _seleccionar_motivo_select2(self, causal_texto):
        """Selecciona el motivo de baja en el formulario usando JavaScript para Select2."""
        id_causal = self.obtener_id_causal(causal_texto)
        
        # Script para interactuar con Select2
        script = f"""
        try {{
            // Método 1: Intentar usar jQuery si está disponible
            if (typeof jQuery !== 'undefined' && jQuery.fn.select2) {{
                var $select = jQuery('#alumnobajaslicencias-id_estatus_academico');
                if ($select.length) {{
                    // Establecer el valor
                    $select.val('{id_causal}');
                    // Disparar los eventos necesarios para Select2
                    $select.trigger('change.select2');
                    $select.trigger('change');
                    console.log('Select2 actualizado con valor: ' + '{id_causal}');
                    return true;
                }}
            }}
            
            // Método 2: JavaScript puro
            var select = document.getElementById('alumnobajaslicencias-id_estatus_academico');
            if (select) {{
                select.value = '{id_causal}';
                
                // Disparar eventos para activar la validación
                var event = new Event('change', {{ bubbles: true }});
                select.dispatchEvent(event);
                
                // También disparar input event
                var inputEvent = new Event('input', {{ bubbles: true }});
                select.dispatchEvent(inputEvent);
                
                console.log('Select actualizado con valor: ' + '{id_causal}');
                return true;
            }}
            
            console.log('No se pudo encontrar el elemento select');
            return false;
        }} catch(e) {{
            console.log('Error al seleccionar motivo: ' + e.message);
            return false;
        }}
        """
        
        try:
            # Ejecutar el script
            resultado = self.driver.execute_script(script)
            
            # Verificar visualmente que se seleccionó
            time.sleep(1)
            
            if resultado:
                print(f"    ✓ Motivo seleccionado: {id_causal} ({causal_texto})")
            else:
                print(f"    ⚠ No se pudo confirmar la selección del motivo {id_causal}")
                
        except Exception as e:
            print(f"    ⚠ Error ejecutando script para seleccionar motivo: {e}")

    def _establecer_fecha_actual(self):
        """Establece la fecha actual en el campo correspondiente."""
        fecha_actual = datetime.now().strftime("%d/%m/%Y")
        
        # Script para establecer la fecha
        script = f"""
        try {{
            // Método 1: Usar jQuery si está disponible
            if (typeof jQuery !== 'undefined') {{
                var $fecha = jQuery('#alumnobajaslicencias-fecha_inicio');
                if ($fecha.length) {{
                    // Establecer el valor
                    $fecha.val('{fecha_actual}');
                    
                    // Disparar eventos para el datepicker
                    $fecha.trigger('change');
                    $fecha.trigger('input');
                    
                    // Si hay un datepicker material, activarlo
                    if ($fecha.hasClass('krajee-datepicker')) {{
                        $fecha.trigger('dp.change');
                    }}
                    
                    console.log('Fecha establecida (jQuery): ' + '{fecha_actual}');
                    return true;
                }}
            }}
            
            // Método 2: JavaScript puro
            var fechaField = document.getElementById('alumnobajaslicencias-fecha_inicio');
            if (fechaField) {{
                fechaField.value = '{fecha_actual}';
                
                // Disparar eventos
                var changeEvent = new Event('change', {{ bubbles: true }});
                fechaField.dispatchEvent(changeEvent);
                
                var inputEvent = new Event('input', {{ bubbles: true }});
                fechaField.dispatchEvent(inputEvent);
                
                console.log('Fecha establecida (JS puro): ' + '{fecha_actual}');
                return true;
            }}
            
            console.log('No se pudo encontrar el campo de fecha');
            return false;
        }} catch(e) {{
            console.log('Error al establecer fecha: ' + e.message);
            return false;
        }}
        """
        
        try:
            resultado = self.driver.execute_script(script)
            if resultado:
                print(f"    ✓ Fecha establecida: {fecha_actual}")
            else:
                print(f"    ⚠ No se pudo establecer la fecha")
                
        except Exception as e:
            print(f"    ⚠ Error estableciendo fecha: {e}")

    def _escribir_descripcion(self, causal_texto):
        """Escribe la descripción del motivo en el formulario."""
        descripcion = f"Proceso automatizado - {causal_texto}"
        
        # Script para escribir en el textarea
        script = f"""
        try {{
            // Método 1: Usar jQuery
            if (typeof jQuery !== 'undefined') {{
                var $textarea = jQuery('#alumnobajaslicencias-descripcion_solicitud');
                if ($textarea.length) {{
                    $textarea.val('{descripcion}');
                    $textarea.trigger('change');
                    $textarea.trigger('input');
                    console.log('Descripción escrita (jQuery)');
                    return true;
                }}
            }}
            
            // Método 2: JavaScript puro
            var textarea = document.getElementById('alumnobajaslicencias-descripcion_solicitud');
            if (textarea) {{
                textarea.value = '{descripcion}';
                
                var changeEvent = new Event('change', {{ bubbles: true }});
                textarea.dispatchEvent(changeEvent);
                
                var inputEvent = new Event('input', {{ bubbles: true }});
                textarea.dispatchEvent(inputEvent);
                
                console.log('Descripción escrita (JS puro)');
                return true;
            }}
            
            console.log('No se pudo encontrar el textarea');
            return false;
        }} catch(e) {{
            console.log('Error al escribir descripción: ' + e.message);
            return false;
        }}
        """
        
        try:
            resultado = self.driver.execute_script(script)
            if resultado:
                print(f"    ✓ Descripción escrita: {descripcion[:50]}...")
            else:
                # Fallback: usar el método tradicional de Selenium
                if self.escribir_en_campo(self.TEXTAREA_DESCRIPCION, descripcion):
                    print(f"    ✓ Descripción escrita (fallback): {descripcion[:50]}...")
                else:
                    print("    ⚠ No se pudo escribir la descripción")
                    
        except Exception as e:
            print(f"    ⚠ Error escribiendo descripción: {e}")

    def _enviar_formulario(self):
        """Envía el formulario completado."""
        try:
            # Buscar el botón de enviar
            boton_enviar = self.esperar_elemento(self.BOTON_ENVIAR, 
                                                mensaje_error="Botón de enviar formulario")
            if not boton_enviar:
                # Intentar encontrar otro botón de submit
                try:
                    boton_enviar = self.driver.find_element(By.CSS_SELECTOR, "button[type='submit']")
                except:
                    print("    ✗ No se encontró botón de enviar")
                    return False
            
            # Hacer scroll para asegurar visibilidad
            self.driver.execute_script(
                "arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", 
                boton_enviar
            )
            time.sleep(0.5)
            
            # Primero intentar hacer click normal
            try:
                boton_enviar.click()
                print("    ✓ Clic normal en botón enviar")
            except:
                # Si falla, usar JavaScript
                self.driver.execute_script("arguments[0].click();", boton_enviar)
                print("    ✓ Clic con JavaScript en botón enviar")
            
            return True
            
        except Exception as error:
            print(f"Error al enviar formulario: {error}")
            return False

    # --- MÉTODOS DE UTILIDAD ---
    def verificar_conexion(self):
        """Verifica que el navegador esté conectado y funcionando."""
        try:
            self.driver.current_url
            return True
        except:
            return False
