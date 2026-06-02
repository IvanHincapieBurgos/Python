"""
Módulo simplificado para registrar acciones y errores en archivos CSV y subirlos a Google Drive.

Funciones principales:
- registrar_accion(): Para registrar operaciones exitosas
- registrar_error(): Para registrar errores

Ambas detectan automáticamente la función y línea desde donde se llaman.
"""
import os
import csv
import traceback
import subprocess
import inspect
from datetime import datetime
from pathlib import Path
from typing import Optional, Dict, Any

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

class _LoggerHelper:
    """Clase para gestionar el registro de acciones y errores."""
    
    def __init__(self, credentials_path: str = None):
        """
        Inicializa el helper de logging.
        
        Args:
            credentials_path: Ruta al archivo de credenciales de Google Drive
        """
        # Si no se especifica ruta, buscar en la carpeta registros/
        if credentials_path is None:
            current_dir = Path(__file__).parent
            credentials_path = current_dir / "credenciales_registros.json"
        
        self.credentials_path = str(credentials_path)
        self.scopes = [
            'https://www.googleapis.com/auth/drive',
            'https://www.googleapis.com/auth/spreadsheets'
        ]
        
        # Configurar las carpetas de Drive
        self.actions_folder_id = '1RA52JZQa3TN-0gC0sXwj4HNfhsp3yHIW'
        self.errors_folder_id = '1e0jViDXChLScTLkslC60V4OpoD38sBqT'
        # Directorio local temporal para CSVs
        self.temp_dir = Path("logs_temp")
        self.temp_dir.mkdir(exist_ok=True)

        self.fixed_fields = {
            "Equipo": "LOGÍSTICA",
            "Proceso": "Producción"
        }

        # Rastreo de errores por función
        self.errors_by_function = {}
        
        # Inicializar servicio de Drive
        self._init_drive_service()
    
    def _init_drive_service(self):
        """Inicializa el servicio de Google Drive API."""
        try:
            if not os.path.exists(self.credentials_path):
                self.drive_service = None
                return
            
            creds = service_account.Credentials.from_service_account_file(
                self.credentials_path,
                scopes=self.scopes
            )
            self.drive_service = build('drive', 'v3', credentials=creds)
            
        except Exception as e:
            print(f"✗ ERROR al inicializar Drive: {e}")
            self.drive_service = None
    
    def _get_git_version(self) -> str:
        """Obtiene la fecha de la última actualización del código desde Git."""
        try:
            result = subprocess.run(
                ['git', 'log', '-1', '--format=%cd', '--date=format:%Y-%m-%d %H:%M:%S'],
                capture_output=True,
                text=True,
                timeout=5
            )
            if result.returncode == 0:
                return result.stdout.strip()
        except Exception:
            pass
        
        return "N/A"
    
    def log_action(
        self,
        function_name: str,
        total_orders: int,
        agente: str,
        additional_data: Optional[Dict[str, Any]] = None
    ) -> bool:
        """Registra una acción exitosa en CSV y lo sube a Drive."""
        try:
            fecha = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            fecha_legible = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            csv_filename = f"action_{function_name}_{fecha}.csv"
            csv_path = self.temp_dir / csv_filename
            
            # Asegurar que el directorio temporal exista
            self.temp_dir.mkdir(parents=True, exist_ok=True)
            
            # Determinar el estado basado en si hubo errores
            if function_name in self.errors_by_function and self.errors_by_function[function_name] > 0:
                estado = "Acción con errores"
            else:
                estado = "Exitoso"
            
            # Datos a registrar
            data = {
                "Fecha": fecha_legible,
                "Función": function_name,
                "Total Órdenes": total_orders,
                "Agente": agente,
                "Versión Git": self._get_git_version(),
                "Estado": estado
            }
            data.update(self.fixed_fields)

            if additional_data:
                data.update(additional_data)
            
            # Crear CSV
            with open(csv_path, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.DictWriter(csvfile, fieldnames=data.keys())
                writer.writeheader()
                writer.writerow(data)
            
            print(f"✓ Acción registrada: {csv_filename}")
            
            # Subir a Drive
            if self.actions_folder_id:
                self._upload_to_drive(csv_path, csv_filename, self.actions_folder_id)
            
            # Limpiar el contador de errores para esta función
            if function_name in self.errors_by_function:
                self.errors_by_function[function_name] = 0
            
            return True
            
        except Exception as e:
            print(f"✗ Error al registrar acción: {e}")
            return False
    
    def log_error(
        self,
        function_name: str,
        error_message: str,
        line_number: Optional[int],
        order_id: Optional[str],
        agente: str,
        full_traceback: Optional[str] = None,
        additional_data: Optional[Dict[str, Any]] = None
    ) -> bool:
        """Registra un error en CSV y lo sube a Drive."""
        try:
            # Incrementar contador de errores para esta función
            if function_name not in self.errors_by_function:
                self.errors_by_function[function_name] = 0
            self.errors_by_function[function_name] += 1
            
            fecha = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            fecha_legible = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            csv_filename = f"error_{function_name}_{fecha}.csv"
            csv_path = self.temp_dir / csv_filename
            
            # Asegurar que el directorio temporal exista
            self.temp_dir.mkdir(parents=True, exist_ok=True)
            
            # Datos a registrar
            data = {
                "Fecha": fecha_legible,
                "Función": function_name,
                "Error": error_message,
                "Línea": line_number if line_number else "N/A",
                "Orden ID": order_id if order_id else "N/A",
                "Agente": agente,
                "Versión Git": self._get_git_version(),
                "Traceback": full_traceback if full_traceback else "N/A"
            }
            data.update(self.fixed_fields)

            if additional_data:
                data.update(additional_data)
            
            # Crear CSV
            with open(csv_path, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.DictWriter(csvfile, fieldnames=data.keys())
                writer.writeheader()
                writer.writerow(data)
            
            print(f"✗ Error registrado: {csv_filename}")
            
            # Subir a Drive
            if self.errors_folder_id:
                self._upload_to_drive(csv_path, csv_filename, self.errors_folder_id)
            
            return True
            
        except Exception as e:
            print(f"✗ Error al registrar error: {e}")
            return False
    
    def _upload_to_drive(self, file_path: Path, file_name: str, folder_id: str) -> bool:
        """Sube un archivo a Google Drive y lo elimina localmente."""
        try:
            if not self.drive_service:
                return False
            
            file_metadata = {
                'name': file_name,
                'parents': [folder_id]
            }
            
            media = MediaFileUpload(str(file_path), mimetype='text/csv', resumable=True)
            
            self.drive_service.files().create(
                body=file_metadata,
                media_body=media,
                fields='id,name,webViewLink'
            ).execute()
            
            # Cerrar el manejador y eliminar archivo local
            del media
            
            try:
                import time
                time.sleep(0.5)  # Pausa para liberar el archivo
                
                if file_path.exists():
                    file_path.unlink()
                    self._cleanup_temp_folder()
            except Exception:
                pass  # No crítico si falla la eliminación
            
            return True
            
        except Exception as e:
            print(f"✗ Error al subir a Drive: {e}")
            return False
    
    def _cleanup_temp_folder(self):
        """Elimina la carpeta temporal si está vacía."""
        try:
            if self.temp_dir.exists() and not any(self.temp_dir.iterdir()):
                self.temp_dir.rmdir()
        except Exception:
            pass  # No crítico si falla

# ============================================================================
# Instancia global del logger (se inicializa una sola vez)
_global_logger = None

def _get_logger():
    """Obtiene o crea la instancia global del logger."""
    global _global_logger
    if _global_logger is None:
        _global_logger = _LoggerHelper()
    return _global_logger

def _get_caller_info():
    """
    Obtiene información del contexto desde donde se llamó la función.
    Retorna: (nombre_funcion, numero_linea)
    """
    # Obtener el stack de llamadas
    stack = inspect.stack()
    
    # stack[0] es _get_caller_info
    # stack[1] es registrar_accion o registrar_error
    # stack[2] es la función que realmente nos interesa
    if len(stack) >= 3:
        caller_frame = stack[2]
        function_name = caller_frame.function
        line_number = caller_frame.lineno
        return function_name, line_number
    
    return "unknown_function", None

def registrar_accion(
    total_orders: int,
    agente: str,
    additional_data: Optional[Dict[str, Any]] = None
) -> bool:
    """Registra una acción exitosa en el sistema de logging.
    Ejemplo de uso:
        registrar_accion(
            total_orders=15,
            agente="Liz Molina"
        )
    """
    logger = _get_logger()
    function_name, _ = _get_caller_info()
    
    return logger.log_action(
        function_name=function_name,
        total_orders=total_orders,
        agente=agente,
        additional_data=additional_data
    )

def registrar_error(
    error_message: str,
    agente: str,
    order_id: Optional[str] = None,
    exception: Optional[Exception] = None,
    additional_data: Optional[Dict[str, Any]] = None
) -> bool:
    """Registra un error en el sistema de logging.
    Ejemplo de uso:
        try:
            # ... código que puede fallar ...
        except Exception as e:
            registrar_error(
                error_message=f"Error procesando orden: {str(e)}",
                agente="Liz Molina",
                order_id="SO-12345",
                exception=e
            )
    """
    logger = _get_logger()
    function_name, _ = _get_caller_info()
    
    # Obtener traceback y línea del error real
    full_traceback = None
    line_number = None
    
    if exception:
        full_traceback = traceback.format_exc()
        # Extraer el número de línea donde ocurrió el error
        tb = traceback.extract_tb(exception.__traceback__)
        if tb:
            # Obtener la última entrada del traceback (donde ocurrió el error)
            line_number = tb[-1].lineno
    
    return logger.log_error(
        function_name=function_name,
        error_message=error_message,
        line_number=line_number,
        order_id=order_id,
        agente=agente,
        full_traceback=full_traceback,
        additional_data=additional_data
    )