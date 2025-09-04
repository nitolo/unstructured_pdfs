"""
Automatización para procesar cartas NDF desde Outlook
Versión mejorada con manejo de excepciones y mejores prácticas
"""
import win32com.client
import re
import logging
import time
import shutil
from datetime import datetime
from pathlib import Path
from typing import Optional, List, Tuple
from contextlib import contextmanager
import sys
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException, 
    NoSuchElementException, 
    WebDriverException
)


class NDFProcessor:
    """Procesador de cartas NDF desde Outlook"""
    
    def __init__(self, base_output_dir: str = None, downloads_dir: str = None, headless_mode: bool = True):
        """
        Inicializa el procesador
        
        Args:
            base_output_dir: Directorio base para guardar archivos
            downloads_dir: Directorio de descargas del navegador
            headless_mode: Si True, ejecuta Selenium en segundo plano sin ventana visible
        """
        self.today = datetime.now()
        self.year = self.today.strftime("%Y")
        self.month = self.today.strftime("%m")
        self.day = self.today.strftime("%d%m%y")
        
        # Configurar directorios
        self.base_dir = Path(base_output_dir or r"Z:\17. Reporting Automation\Cartas NDFs\Cartas sin firmas")
        self.output_dir = self.base_dir / self.year / self.month / self.day
        self.downloads_dir = Path(downloads_dir or r"C:\Users\ntorreslo\Downloads")
        
        # Configurar modo de ejecución
        self.headless_mode = headless_mode
        
        # Configurar logging
        self._setup_logging()
        
        # Expresiones regulares compiladas para mejor rendimiento
        self.safe_filename_pattern = re.compile(r'[^0-9a-zA-ZáéíóúÁÉÍÓÚñÑ\.\-_ ]+')
        self.url_pattern = re.compile(r"https://[^\s]+")
        self.code_pattern = re.compile(r"Verification Code:\s*([A-Z0-9]+)")
        
    def _setup_logging(self) -> None:
        """Configura el sistema de logging"""
        log_dir = self.base_dir / "logs"
        log_dir.mkdir(parents=True, exist_ok=True)
        
        log_file = log_dir / f"ndf_processing_{self.day}.log"
        
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_file, encoding='utf-8'),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)
        
    def _create_output_directory(self) -> bool:
        """
        Crea el directorio de salida
        
        Returns:
            bool: True si se creó exitosamente
        """
        try:
            self.output_dir.mkdir(parents=True, exist_ok=True)
            self.logger.info(f"Directorio de salida creado/verificado: {self.output_dir}")
            return True
        except Exception as e:
            self.logger.error(f"Error creando directorio de salida: {e}")
            return False
            
    def _get_outlook_connection(self) -> Optional[object]:
        """
        Establece conexión con Outlook
        
        Returns:
            object: Objeto inbox o None si falla
        """
        try:
            self.logger.info("Iniciando conexión con Outlook...")
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            
            # Usar el método más robusto que funcionó
            main_folder = outlook.Folders["Mercado de Capitales Colombia"]
            inbox = main_folder.Folders["Cartas NDF"]
            
            # Verificar acceso
            message_count = inbox.Items.Count
            self.logger.info(f"✓ Conexión exitosa. Mensajes en carpeta: {message_count}")
            
            return inbox
            
        except Exception as e:
            self.logger.error(f"Error conectando con Outlook: {e}")
            
            # Intentar método alternativo como backup
            try:
                self.logger.info("Intentando método alternativo...")
                outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
                inbox = outlook.Folders("Mercado de Capitales Colombia").Folders("Cartas NDF")
                
                message_count = inbox.Items.Count
                self.logger.info(f"✓ Método alternativo exitoso. Mensajes: {message_count}")
                
                return inbox
                
            except Exception as e2:
                self.logger.error(f"Método alternativo también falló: {e2}")
                return None
            
    def _get_filtered_messages(self, inbox: object) -> Tuple[List[object], datetime, datetime]:
        """
        Obtiene mensajes filtrados por fecha
        
        Args:
            inbox: Objeto inbox de Outlook
            
        Returns:
            tuple: (lista_mensajes_filtrados, fecha_inicio, fecha_fin)
        """
        try:
            messages = inbox.Items
            messages.Sort("[ReceivedTime]", True)  # Orden descendente
            
            self.logger.info(f"Total mensajes en carpeta: {messages.Count}")
            
            # Definir rango de fechas
            start = datetime(self.today.year, self.today.month, 1)
            end = self.today
            
            self.logger.info(f"Filtrando mensajes entre {start.date()} y {end.date()}")
            
            # Filtrar manualmente (más confiable que Restrict)
            filtered_messages = []
            
            for i in range(1, messages.Count + 1):
                try:
                    message = messages.Item(i)
                    received = message.ReceivedTime
                    
                    # Convertir a naive datetime si tiene zona horaria
                    if hasattr(received, "tzinfo") and received.tzinfo is not None:
                        received = received.replace(tzinfo=None)
                    
                    # Verificar si está en el rango de fechas
                    if start <= received <= end:
                        filtered_messages.append(message)
                        
                except Exception as e:
                    self.logger.warning(f"Error procesando mensaje {i}: {e}")
                    continue
            
            self.logger.info(f"Filtrados {len(filtered_messages)} mensajes del período")
            
            # Log de algunos mensajes encontrados para verificación
            if filtered_messages:
                self.logger.info("Primeros mensajes encontrados:")
                for i, msg in enumerate(filtered_messages[:3]):
                    try:
                        self.logger.info(f"  - {msg.Subject[:50]} ({msg.ReceivedTime})")
                    except:
                        self.logger.info(f"  - Mensaje {i+1} (error leyendo detalles)")
            
            return filtered_messages, start, end
            
        except Exception as e:
            self.logger.error(f"Error filtrando mensajes: {e}")
            return [], None, None
            
    def _sanitize_filename(self, filename: str) -> str:
        """
        Limpia el nombre de archivo para evitar caracteres problemáticos
        
        Args:
            filename: Nombre original del archivo
            
        Returns:
            str: Nombre sanitizado
        """
        return self.safe_filename_pattern.sub('', filename)
        
    def _save_attachment(self, attachment: object) -> bool:
        """
        Guarda un adjunto específico
        
        Args:
            attachment: Objeto attachment de Outlook
            
        Returns:
            bool: True si se guardó exitosamente
        """
        try:
            filename = attachment.FileName.lower()
            
            if not (filename.endswith(".pdf") or filename.endswith(".xlsx")):
                return False
                
            safe_filename = self._sanitize_filename(attachment.FileName)
            file_path = self.output_dir / safe_filename
            
            # Evitar sobrescribir archivos existentes
            if file_path.exists():
                timestamp = datetime.now().strftime("%H%M%S")
                name_parts = safe_filename.rsplit('.', 1)
                if len(name_parts) == 2:
                    safe_filename = f"{name_parts[0]}_{timestamp}.{name_parts[1]}"
                else:
                    safe_filename = f"{safe_filename}_{timestamp}"
                file_path = self.output_dir / safe_filename
            
            attachment.SaveAsFile(str(file_path))
            self.logger.info(f"Adjunto guardado: {safe_filename}")
            return True
            
        except Exception as e:
            self.logger.error(f"Error guardando adjunto {attachment.FileName}: {e}")
            return False
            
    def process_attachments(self, messages: List[object], start: datetime, end: datetime) -> int:
        """
        Procesa todos los adjuntos de los mensajes
        
        Args:
            messages: Lista de mensajes filtrados
            start: Fecha de inicio
            end: Fecha de fin
            
        Returns:
            int: Número de adjuntos procesados
        """
        if not messages or len(messages) == 0:
            self.logger.info("No hay correos con adjuntos para procesar")
            return 0
            
        processed_count = 0
        
        try:
            for message in messages:
                try:
                    received = message.ReceivedTime
                    # Manejo seguro de timezone
                    if hasattr(received, "tzinfo") and received.tzinfo is not None:
                        received = received.replace(tzinfo=None)
                        
                    if not (start <= received <= end):
                        continue
                        
                    attachments = message.Attachments
                    for attachment in attachments:
                        if self._save_attachment(attachment):
                            processed_count += 1
                            
                except Exception as e:
                    self.logger.error(f"Error procesando mensaje: {e}")
                    continue
                    
        except Exception as e:
            self.logger.error(f"Error en proceso de adjuntos: {e}")
            
        self.logger.info(f"Total de adjuntos procesados: {processed_count}")
        return processed_count
        
    @contextmanager
    def _get_webdriver(self, headless: bool = True):
        """
        Context manager para manejar el navegador web
        
        Args:
            headless: Si True, ejecuta en modo headless (segundo plano)
        
        Yields:
            webdriver: Instancia del navegador
        """
        driver = None
        try:
            # Configurar opciones del navegador
            chrome_options = Options()
            
            # MODO HEADLESS - Ejecuta en segundo plano sin ventana visible
            if headless:
                chrome_options.add_argument("--headless=new")  # Nuevo modo headless más estable
                chrome_options.add_argument("--disable-gpu")
                chrome_options.add_argument("--no-sandbox")
                chrome_options.add_argument("--disable-dev-shm-usage")
                chrome_options.add_argument("--disable-extensions")
                chrome_options.add_argument("--disable-logging")
                chrome_options.add_argument("--silent")
                # Configurar tamaño de ventana virtual
                chrome_options.add_argument("--window-size=1920,1080")
            else:
                # Modo visible pero minimizado
                chrome_options.add_argument("--start-minimized")
                chrome_options.add_argument("--no-sandbox")
                chrome_options.add_argument("--disable-dev-shm-usage")
            
            # Configurar directorio de descarga
            prefs = {
                "download.default_directory": str(self.downloads_dir),
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "safebrowsing.enabled": True,
                "profile.default_content_settings.popups": 0,
                "profile.default_content_setting_values.automatic_downloads": 1
            }
            chrome_options.add_experimental_option("prefs", prefs)
            
            driver = webdriver.Chrome(options=chrome_options)
            driver.set_page_load_timeout(30)
            
            yield driver
            
        except WebDriverException as e:
            self.logger.error(f"Error inicializando navegador: {e}")
            yield None
        finally:
            if driver:
                try:
                    driver.quit()
                except Exception as e:
                    self.logger.error(f"Error cerrando navegador: {e}")
                    
    def _extract_url_and_code(self, body: str) -> Tuple[Optional[str], Optional[str]]:
        """
        Extrae URL y código de verificación del cuerpo del correo
        
        Args:
            body: Cuerpo del correo
            
        Returns:
            tuple: (url, código) o (None, None) si no se encuentran
        """
        url_match = self.url_pattern.search(body)
        code_match = self.code_pattern.search(body)
        
        if url_match and code_match:
            return url_match.group(0), code_match.group(1)
        return None, None
        
    def _automate_download(self, url: str, code: str) -> bool:
        """
        Automatiza la descarga usando Selenium
        
        Args:
            url: URL de descarga
            code: Código de verificación
            
        Returns:
            bool: True si la descarga fue exitosa
        """
        with self._get_webdriver(headless=self.headless_mode) as driver:
            if not driver:
                return False
                
            try:
                self.logger.info(f"Accediendo a URL en modo {'headless' if self.headless_mode else 'visible'}: {url}")
                driver.get(url)
                wait = WebDriverWait(driver, 15)  # Aumentado timeout para modo headless
                
                # Esperar y encontrar el campo de código
                input_box = wait.until(
                    EC.presence_of_element_located((By.ID, "txtAccessCode"))
                )
                
                input_box.clear()
                input_box.send_keys(code)
                input_box.send_keys(Keys.RETURN)
                
                self.logger.info("Código de verificación ingresado")
                
                # Esperar y hacer clic en el enlace de descarga
                download_link = wait.until(
                    EC.element_to_be_clickable((By.CLASS_NAME, "Q416_downloadspecific"))
                )
                
                download_link.click()
                self.logger.info("Descarga iniciada en segundo plano")
                
                # Esperar más tiempo en modo headless para que complete la descarga
                wait_time = 5 if self.headless_mode else 3
                time.sleep(wait_time)
                return True
                
            except TimeoutException:
                self.logger.error("Timeout esperando elementos en la página")
                return False
            except NoSuchElementException as e:
                self.logger.error(f"Elemento no encontrado: {e}")
                return False
            except Exception as e:
                self.logger.error(f"Error durante automatización: {e}")
                return False
                
    def _move_downloaded_files(self) -> int:
        """
        Mueve archivos PDF descargados al directorio de salida
        
        Returns:
            int: Número de archivos movidos
        """
        moved_count = 0
        
        try:
            pattern = "Confirmation-AE*.pdf"
            today_date = self.today.date()
            
            for file_path in self.downloads_dir.glob(pattern):
                try:
                    if not file_path.is_file():
                        continue
                        
                    # Verificar que el archivo fue modificado hoy
                    modified_time = datetime.fromtimestamp(file_path.stat().st_mtime)
                    if modified_time.date() != today_date:
                        continue
                        
                    # Generar nombre único si el archivo ya existe
                    destination = self.output_dir / file_path.name
                    if destination.exists():
                        timestamp = datetime.now().strftime("%H%M%S")
                        name_parts = file_path.name.rsplit('.', 1)
                        new_name = f"{name_parts[0]}_{timestamp}.{name_parts[1]}"
                        destination = self.output_dir / new_name
                        
                    shutil.move(str(file_path), str(destination))
                    self.logger.info(f"Archivo movido: {file_path.name}")
                    moved_count += 1
                    
                except Exception as e:
                    self.logger.error(f"Error moviendo archivo {file_path.name}: {e}")
                    
        except Exception as e:
            self.logger.error(f"Error en proceso de movimiento de archivos: {e}")
            
        return moved_count
        
    def process_download_links(self, messages: List[object], start: datetime, end: datetime) -> int:
        """
        Procesa correos con enlaces de descarga JPM
        
        Args:
            messages: Lista de mensajes filtrados
            start: Fecha de inicio
            end: Fecha de fin
            
        Returns:
            int: Número de descargas procesadas
        """
        if not messages:
            return 0
            
        processed_count = 0
        
        try:
            for message in messages:
                try:
                    if not message.Subject.startswith("JPM Confirmation ID"):
                        continue
                        
                    received = message.ReceivedTime
                    # Manejo seguro de timezone
                    if hasattr(received, "tzinfo") and received.tzinfo is not None:
                        received = received.replace(tzinfo=None)
                        
                    if not (start <= received <= end):
                        continue
                        
                    url, code = self._extract_url_and_code(message.Body)
                    
                    if not url or not code:
                        self.logger.warning(f"No se encontró URL o código en mensaje: {message.Subject}")
                        continue
                        
                    if self._automate_download(url, code):
                        processed_count += 1
                        time.sleep(2)  # Pausa entre descargas
                        
                except Exception as e:
                    self.logger.error(f"Error procesando enlace de descarga: {e}")
                    continue
                    
        except Exception as e:
            self.logger.error(f"Error en proceso de enlaces de descarga: {e}")
            
        # Mover archivos descargados
        moved_files = self._move_downloaded_files()
        self.logger.info(f"Archivos descargados y movidos: {moved_files}")
        
        return processed_count
        
    def run(self) -> bool:
        """
        Ejecuta el proceso completo
        
        Returns:
            bool: True si el proceso se completó exitosamente
        """
        self.logger.info("=== Iniciando procesamiento de cartas NDF ===")
        
        # Crear directorio de salida
        if not self._create_output_directory():
            return False
            
        # Conectar con Outlook
        inbox = self._get_outlook_connection()
        if not inbox:
            return False
            
        # Obtener mensajes filtrados
        messages, start, end = self._get_filtered_messages(inbox)
        if not messages:
            return False
            
        try:
            # Procesar adjuntos
            attachments_processed = self.process_attachments(messages, start, end)
            
            # Procesar enlaces de descarga
            downloads_processed = self.process_download_links(messages, start, end)
            
            self.logger.info(f"=== Proceso completado ===")
            self.logger.info(f"Adjuntos procesados: {attachments_processed}")
            self.logger.info(f"Descargas procesadas: {downloads_processed}")
            
            return True
            
        except Exception as e:
            self.logger.error(f"Error durante el procesamiento: {e}")
            return False


def main():
    """Función principal"""
    try:
        # CONFIGURACIÓN DE MODO DE EJECUCIÓN
        # headless_mode=True  -> Ejecuta en segundo plano (recomendado)
        # headless_mode=False -> Muestra ventana del navegador
        
        processor = NDFProcessor(headless_mode=True)  # Cambiar a False si necesitas ver el navegador
        success = processor.run()
        
        if not success:
            print("El procesamiento no se completó exitosamente. Revisar logs.")
            return 1
            
        print("Procesamiento completado exitosamente en segundo plano.")
        return 0
        
    except Exception as e:
        print(f"Error crítico: {e}")
        return 1

if __name__ == "__main__":
    sys.exit(main())