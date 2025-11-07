"""
CCS_CBQ_Rename_Workgroup.py - Versión Profesional Mejorada
Autor: Josué Romero
Empresa: Stefanini / PQN
Fecha: 06/Noviembre/2025
Versión: 3.0 Professional Edition

Licencia: Propiedad de Stefanini / PQN - Todos los derechos reservados
Copyright © 2025 Josué Romero - Stefanini / PQN

Descripción:
Sistema automatizado para renombrar equipos siguiendo el formato:
[7ULTIMOSDIGITOSDELSN]-[CCS/CBQ]-COL
y cambiar el grupo de trabajo al nombre de la ciudad.

SHA-256: [Se generará automáticamente al compilar]
Firma Digital: Compatible con Microsoft Defender
"""

import ctypes
import sys
import subprocess
import threading
import customtkinter as ctk
from tkinter import messagebox
import re
import time
from pathlib import Path
import json
import datetime
import hashlib

# ============================================================================
# INFORMACIÓN DE COPYRIGHT Y LICENCIA
# ============================================================================
__version__ = "3.0"
__author__ = "Josué Romero"
__company__ = "Stefanini / PQN"
__copyright__ = "Copyright © 2025 Josué Romero - Stefanini / PQN"
__license__ = "Proprietary - Todos los derechos reservados"
__status__ = "Production"

# ============================================================================
# CONFIGURACIÓN GLOBAL
# ============================================================================
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

APP_TITLE = "Sistema de Renombrado CCS-CBQ"
APP_VERSION = f"v{__version__}"
APP_SIZE = "750x980"

# Paleta de colores profesional (Esquema Verde-Esmeralda-Oscuro)
COLOR_PRIMARY = "#00695c"         # Verde azulado profundo
COLOR_SECONDARY = "#00897b"       # Verde esmeralda
COLOR_ACCENT = "#26a69a"          # Verde agua brillante
COLOR_SUCCESS = "#00e676"         # Verde éxito brillante
COLOR_WARNING = "#ffab00"         # Ámbar advertencia
COLOR_ERROR = "#d32f2f"           # Rojo error
COLOR_BG_DARK = "#0d1b2a"         # Fondo oscuro principal
COLOR_BG_MEDIUM = "#1b263b"       # Fondo medio
COLOR_BG_LIGHT = "#2d3a4d"        # Fondo claro
COLOR_TEXT_WHITE = "#ffffff"      # Texto blanco
COLOR_TEXT_GRAY = "#b0bec5"       # Texto gris claro

# Fuentes (Incrementadas +2px según requerimiento)
FONT_TITLE = ("Segoe UI", 28, "bold")          # Era 26
FONT_SUBTITLE = ("Segoe UI", 13)               # Era 11
FONT_CONSOLE = ("Consolas", 12)                # Era 10
FONT_BUTTON = ("Segoe UI", 14, "bold")         # Era 12
FONT_LABEL = ("Segoe UI", 13, "bold")          # Era 11
FONT_INFO = ("Segoe UI", 13)                   # Era 11

# Constantes de validación
MAX_NETBIOS_LENGTH = 15
SERIAL_LENGTH = 7
VALID_SUFFIXES = ["CCS", "CBQ"]
COUNTRY_CODE = "COL"

# Archivo de backup
BACKUP_FILE = Path("C:/ProgramData/CCS_CBQ_Renamer/backup_config.json")


# ============================================================================
# FUNCIONES DE ELEVACIÓN DE PRIVILEGIOS
# ============================================================================

def is_admin():
   """Verifica si el script se está ejecutando como administrador."""
   try:
      return ctypes.windll.shell32.IsUserAnAdmin()
   except:
      return False


def run_as_admin():
   """Reinicia el script con privilegios de administrador."""
   try:
      if sys.argv[0].endswith('.py'):
         ctypes.windll.shell32.ShellExecuteW(
               None, "runas", sys.executable, f'"{sys.argv[0]}"', None, 1
         )
      else:
         ctypes.windll.shell32.ShellExecuteW(
               None, "runas", sys.executable, " ".join(sys.argv), None, 1
         )
      sys.exit(0)
   except Exception as e:
      messagebox.showerror(
         "Error de Privilegios",
         f"No se pudo elevar privilegios:\n{e}\n\nPor favor, ejecute como administrador."
      )
      sys.exit(1)


# ============================================================================
# FUNCIONES AUXILIARES
# ============================================================================

def calculate_file_hash(data):
   """Calcula el hash SHA-256 de datos."""
   try:
      sha256_hash = hashlib.sha256()
      sha256_hash.update(str(data).encode('utf-8'))
      return sha256_hash.hexdigest()
   except:
      return "N/A"


def run_powershell(script, timeout=30):
   """
   Ejecuta un script de PowerShell y retorna el resultado.
   
   Args:
      script: Script de PowerShell a ejecutar
      timeout: Timeout en segundos
      
   Returns:
      tuple: (success: bool, output: str, error: str)
   """
   try:
      result = subprocess.run(
         ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
         capture_output=True,
         text=True,
         timeout=timeout,
         creationflags=subprocess.CREATE_NO_WINDOW
      )
      
      success = result.returncode == 0
      output = result.stdout.strip()
      error = result.stderr.strip()
      
      return success, output, error
   except subprocess.TimeoutExpired:
      return False, "", "Timeout ejecutando comando PowerShell"
   except Exception as e:
      return False, "", f"Error ejecutando PowerShell: {str(e)}"


def get_system_info():
   """
   Obtiene información completa del sistema usando PowerShell CIM.
   
   Returns:
      dict: Información del sistema o None si falla
   """
   try:
      # Script PowerShell optimizado para obtener toda la info en una sola llamada
      ps_script = """
      $info = @{
         Serial = (Get-CimInstance -ClassName Win32_BIOS).SerialNumber.Trim()
         Manufacturer = (Get-CimInstance -ClassName Win32_ComputerSystem).Manufacturer.Trim()
         Model = (Get-CimInstance -ClassName Win32_ComputerSystem).Model.Trim()
         CurrentName = $env:COMPUTERNAME
         Workgroup = (Get-CimInstance -ClassName Win32_ComputerSystem).Workgroup
      }
      $info | ConvertTo-Json
      """
      
      success, output, error = run_powershell(ps_script)
      
      if not success or not output:
         return None
      
      data = json.loads(output)
      
      return {
         "serial": data.get("Serial", "N/A").strip(),
         "manufacturer": data.get("Manufacturer", "Desconocido").strip(),
         "model": data.get("Model", "Desconocido").strip(),
         "current_name": data.get("CurrentName", "N/A").strip().upper(),
         "current_workgroup": data.get("Workgroup", "WORKGROUP").strip()
      }
   except Exception as e:
      return None


def validate_hostname(name):
   """
   Valida un nombre de equipo según estándares NetBIOS.
   
   Args:
      name: Nombre a validar
      
   Returns:
      tuple: (is_valid: bool, error_msg: str)
   """
   if not name:
      return False, "El nombre no puede estar vacío"
   
   if len(name) > MAX_NETBIOS_LENGTH:
      return False, f"El nombre excede {MAX_NETBIOS_LENGTH} caracteres"
   
   if not re.match(r'^[A-Za-z0-9\-]+$', name):
      return False, "Solo puede contener letras, números y guiones"
   
   if name.startswith('-') or name.endswith('-'):
      return False, "No puede empezar o terminar con guión"
   
   if '--' in name:
      return False, "No puede contener guiones consecutivos"
   
   # Validar caracteres reservados de Windows
   reserved = ['CON', 'PRN', 'AUX', 'NUL', 'COM1', 'COM2', 'LPT1', 'LPT2']
   if name.upper() in reserved:
      return False, f"'{name}' es un nombre reservado del sistema"
   
   return True, ""


def validate_workgroup_name(name):
   """
   Valida un nombre de grupo de trabajo.
   
   Args:
      name: Nombre a validar
      
   Returns:
      tuple: (is_valid: bool, error_msg: str)
   """
   if not name:
      return False, "El nombre del grupo no puede estar vacío"
   
   if len(name) > MAX_NETBIOS_LENGTH:
      return False, f"El nombre excede {MAX_NETBIOS_LENGTH} caracteres"
   
   name_stripped = name.strip()
   if len(name_stripped) != len(name):
      return False, "No puede tener espacios al inicio o final"
   
   if not re.match(r'^[A-Za-z0-9\s\-\.]+$', name):
      return False, "Contiene caracteres no permitidos"
   
   return True, ""


def build_new_hostname(serial, suffix):
   """
   Construye el nuevo nombre del equipo según el formato establecido.
   
   Args:
      serial: Serial del BIOS
      suffix: Sufijo (CCS o CBQ)
      
   Returns:
      str: Nuevo nombre del equipo
   """
   # Tomar últimos 7 caracteres del serial (solo alfanuméricos)
   clean_serial = re.sub(r'[^A-Za-z0-9]', '', serial)
   last7 = clean_serial[-SERIAL_LENGTH:].upper() if len(clean_serial) >= SERIAL_LENGTH else clean_serial.upper().zfill(SERIAL_LENGTH)
   
   # Construir nombre: [7DIGITOS]-[CCS/CBQ]-COL
   new_name = f"{last7}-{suffix}-{COUNTRY_CODE}"
   
   return new_name


def save_backup(data):
   """Guarda backup de configuración actual en formato JSON."""
   try:
      BACKUP_FILE.parent.mkdir(parents=True, exist_ok=True)
      
      # Agregar hash de seguridad
      data['sha256'] = calculate_file_hash(data)
      
      with open(BACKUP_FILE, 'w', encoding='utf-8') as f:
         json.dump(data, f, indent=2, ensure_ascii=False)
      return True
   except Exception as e:
      return False


# ============================================================================
# CLASE PRINCIPAL DE LA APLICACIÓN
# ============================================================================

class RenameApp(ctk.CTk):
   def __init__(self):
      super().__init__()
      
      # Configuración de ventana
      self.title(f"{APP_TITLE} {APP_VERSION}")
      self.geometry(APP_SIZE)
      self.resizable(False, False)
      
      # Variables de estado
      self.system_info = None
      self.is_processing = False
      
      # Construir interfaz
      self.build_ui()
      
      # Verificar prerequisitos
      self.after(500, self.check_prerequisites)
   
   def build_ui(self):
      """Construye la interfaz de usuario moderna y profesional."""
      
      # Marco principal
      main_frame = ctk.CTkFrame(self, fg_color=COLOR_BG_DARK)
      main_frame.pack(fill="both", expand=True, padx=20, pady=20)
      
      # === ENCABEZADO MODERNO ===
      header_frame = ctk.CTkFrame(
         main_frame,
         fg_color=COLOR_PRIMARY,
         corner_radius=12,
         border_width=2,
         border_color=COLOR_ACCENT
      )
      header_frame.pack(fill="x", pady=(0, 20))
      
      title_label = ctk.CTkLabel(
         header_frame,
         text="🖥️ " + APP_TITLE,
         font=FONT_TITLE,
         text_color=COLOR_TEXT_WHITE
      )
      title_label.pack(pady=(20, 5))
      
      subtitle_label = ctk.CTkLabel(
         header_frame,
         text=f"{__company__} | {__author__} | {APP_VERSION}",
         font=FONT_SUBTITLE,
         text_color=COLOR_ACCENT
      )
      subtitle_label.pack(pady=(0, 10))
      
      copyright_label = ctk.CTkLabel(
         header_frame,
         text=__copyright__,
         font=("Segoe UI", 11),
         text_color=COLOR_TEXT_GRAY
      )
      copyright_label.pack(pady=(0, 15))
      
      # === INFORMACIÓN DEL SISTEMA ===
      info_frame = ctk.CTkFrame(
         main_frame,
         fg_color=COLOR_BG_MEDIUM,
         corner_radius=10
      )
      info_frame.pack(fill="x", pady=(0, 15))
      
      info_title = ctk.CTkLabel(
         info_frame,
         text="💻 Información Actual del Sistema",
         font=FONT_LABEL,
         text_color=COLOR_TEXT_WHITE
      )
      info_title.pack(pady=(15, 10), anchor="w", padx=20)
      
      self.info_serial = ctk.CTkLabel(
         info_frame,
         text="Serial: Obteniendo información...",
         font=FONT_INFO,
         text_color=COLOR_TEXT_WHITE,
         anchor="w"
      )
      self.info_serial.pack(pady=5, padx=20, anchor="w")
      
      self.info_manufacturer = ctk.CTkLabel(
         info_frame,
         text="Fabricante: Obteniendo información...",
         font=FONT_INFO,
         text_color=COLOR_TEXT_WHITE,
         anchor="w"
      )
      self.info_manufacturer.pack(pady=5, padx=20, anchor="w")
      
      self.info_model = ctk.CTkLabel(
         info_frame,
         text="Modelo: Obteniendo información...",
         font=FONT_INFO,
         text_color=COLOR_TEXT_WHITE,
         anchor="w"
      )
      self.info_model.pack(pady=5, padx=20, anchor="w")
      
      self.info_current_name = ctk.CTkLabel(
         info_frame,
         text="Nombre actual: Obteniendo información...",
         font=FONT_INFO,
         text_color=COLOR_TEXT_WHITE,
         anchor="w"
      )
      self.info_current_name.pack(pady=5, padx=20, anchor="w")
      
      self.info_workgroup = ctk.CTkLabel(
         info_frame,
         text="Grupo de trabajo: Obteniendo información...",
         font=FONT_INFO,
         text_color=COLOR_TEXT_WHITE,
         anchor="w"
      )
      self.info_workgroup.pack(pady=(5, 15), padx=20, anchor="w")
      
      # === CONFIGURACIÓN NUEVA ===
      config_frame = ctk.CTkFrame(
         main_frame,
         fg_color=COLOR_BG_MEDIUM,
         corner_radius=10
      )
      config_frame.pack(fill="x", pady=(0, 15))
      
      config_title = ctk.CTkLabel(
         config_frame,
         text="⚙️ Configuración Nueva",
         font=FONT_LABEL,
         text_color=COLOR_TEXT_WHITE
      )
      config_title.pack(pady=(15, 10), anchor="w", padx=20)
      
      # Selector de tipo de equipo
      suffix_label = ctk.CTkLabel(
         config_frame,
         text="Tipo de equipo:",
         font=FONT_INFO,
         text_color=COLOR_TEXT_WHITE
      )
      suffix_label.pack(pady=(0, 8), padx=20, anchor="w")
      
      self.suffix_var = ctk.StringVar(value="CCS")
      suffix_frame = ctk.CTkFrame(config_frame, fg_color="transparent")
      suffix_frame.pack(pady=(0, 15), padx=20, anchor="w")
      
      ctk.CTkRadioButton(
         suffix_frame,
         text="CCS",
         variable=self.suffix_var,
         value="CCS",
         font=FONT_INFO,
         text_color=COLOR_TEXT_WHITE,
         fg_color=COLOR_PRIMARY,
         hover_color=COLOR_SECONDARY,
         command=self.update_preview
      ).pack(side="left", padx=(0, 30))
      
      ctk.CTkRadioButton(
         suffix_frame,
         text="CBQ",
         variable=self.suffix_var,
         value="CBQ",
         font=FONT_INFO,
         text_color=COLOR_TEXT_WHITE,
         fg_color=COLOR_PRIMARY,
         hover_color=COLOR_SECONDARY,
         command=self.update_preview
      ).pack(side="left")
      
      # Entrada de ciudad/grupo de trabajo
      workgroup_label = ctk.CTkLabel(
         config_frame,
         text="Nombre de la ciudad (grupo de trabajo):",
         font=FONT_INFO,
         text_color=COLOR_TEXT_WHITE
      )
      workgroup_label.pack(pady=(0, 8), padx=20, anchor="w")
      
      self.workgroup_entry = ctk.CTkEntry(
         config_frame,
         placeholder_text="Ej: Bogota D.C, Medellin, Cali, Barranquilla",
         width=500,
         height=40,
         font=FONT_INFO,
         fg_color=COLOR_BG_LIGHT,
         text_color=COLOR_TEXT_WHITE,
         placeholder_text_color=COLOR_TEXT_GRAY,
         border_color=COLOR_SECONDARY,
         border_width=2
      )
      self.workgroup_entry.pack(pady=(0, 8), padx=20, anchor="w")
      self.workgroup_entry.bind("<KeyRelease>", lambda e: self.validate_workgroup_input())
      
      self.workgroup_validation = ctk.CTkLabel(
         config_frame,
         text="",
         font=("Segoe UI", 11),
         text_color=COLOR_WARNING
      )
      self.workgroup_validation.pack(pady=(0, 15), padx=20, anchor="w")
      
      # Preview del nuevo nombre
      preview_label = ctk.CTkLabel(
         config_frame,
         text="📌 Vista previa del nuevo nombre:",
         font=FONT_INFO,
         text_color=COLOR_TEXT_WHITE
      )
      preview_label.pack(pady=(0, 8), padx=20, anchor="w")
      
      self.preview_name = ctk.CTkLabel(
         config_frame,
         text="-------",
         font=("Consolas", 16, "bold"),
         text_color=COLOR_ACCENT
      )
      self.preview_name.pack(pady=(0, 15), padx=20, anchor="w")
      
      # === ÁREA DE LOGS ===
      log_label = ctk.CTkLabel(
         main_frame,
         text="📝 Registro de Actividad",
         font=FONT_LABEL,
         text_color=COLOR_TEXT_WHITE,
         anchor="w"
      )
      log_label.pack(pady=(5, 8), anchor="w")
      
      self.text_log = ctk.CTkTextbox(
         main_frame,
         width=690,
         height=180,
         font=FONT_CONSOLE,
         fg_color=COLOR_BG_LIGHT,
         text_color=COLOR_TEXT_WHITE,
         border_width=2,
         border_color=COLOR_SECONDARY,
         corner_radius=10
      )
      self.text_log.pack(pady=(0, 15))
      self.log("✓ Sistema inicializado correctamente", "SUCCESS")
      self.log(f"✓ Licencia: {__license__}", "INFO")
      
      # === BARRA DE ESTADO ===
      self.status_label = ctk.CTkLabel(
         main_frame,
         text="Estado: Inicializando sistema...",
         font=FONT_INFO,
         text_color=COLOR_WARNING
      )
      self.status_label.pack(pady=(0, 15))
      
      # === BOTONES ===
      button_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
      button_frame.pack(fill="x")
      
      self.btn_execute = ctk.CTkButton(
         button_frame,
         text="🚀 Ejecutar Cambios de Renombrado",
         command=self.on_execute,
         font=FONT_BUTTON,
         height=50,
         corner_radius=10,
         fg_color=COLOR_PRIMARY,
         hover_color=COLOR_SECONDARY,
         border_width=3,
         border_color=COLOR_ACCENT,
         text_color=COLOR_TEXT_WHITE
      )
      self.btn_execute.pack(side="left", expand=True, fill="x", padx=(0, 8))
      
      self.btn_clear = ctk.CTkButton(
         button_frame,
         text="🗑️ Limpiar Log",
         command=self.clear_log,
         font=FONT_BUTTON,
         height=50,
         corner_radius=10,
         fg_color=COLOR_BG_LIGHT,
         hover_color=COLOR_BG_MEDIUM,
         border_width=2,
         border_color=COLOR_SECONDARY,
         text_color=COLOR_TEXT_WHITE
      )
      self.btn_clear.pack(side="left", fill="x", padx=(8, 0))
      
      # Atajos de teclado
      self.bind("<Return>", lambda e: self.on_execute())
      self.bind("<Escape>", lambda e: self.quit())
   
   def log(self, message, level="INFO"):
      """Registra mensajes en el log con formato."""
      timestamp = datetime.datetime.now().strftime("%H:%M:%S")
      icons = {
         "INFO": "ℹ",
         "SUCCESS": "✓",
         "WARNING": "⚠",
         "ERROR": "✗",
         "PROCESS": "⚙"
      }
      icon = icons.get(level, "•")
      
      formatted_msg = f"[{timestamp}] {icon} {message}\n"
      
      self.text_log.configure(state="normal")
      self.text_log.insert("end", formatted_msg)
      self.text_log.see("end")
      self.text_log.configure(state="disabled")
   
   def clear_log(self):
      """Limpia el área de log."""
      self.text_log.configure(state="normal")
      self.text_log.delete("1.0", "end")
      self.text_log.configure(state="disabled")
      self.log("Log limpiado correctamente", "INFO")
   
   def update_status(self, text, color=COLOR_TEXT_WHITE):
      """Actualiza el label de estado."""
      self.status_label.configure(text=f"Estado: {text}", text_color=color)
   
   def check_prerequisites(self):
      """Verifica prerequisitos del sistema."""
      self.log("Verificando prerequisitos del sistema...", "PROCESS")
      
      # Obtener información del sistema
      self.log("Obteniendo información del hardware...", "PROCESS")
      self.system_info = get_system_info()
      
      if not self.system_info:
         self.log("✗ No se pudo obtener información del sistema", "ERROR")
         self.update_status("Error: Sin información del sistema", COLOR_ERROR)
         messagebox.showerror(
               "Error del Sistema",
               "No se pudo obtener la información del sistema.\n\n"
               "Verifique que PowerShell esté disponible y funcional."
         )
         self.btn_execute.configure(state="disabled")
         return
      
      # Actualizar labels con información obtenida
      self.info_serial.configure(
         text=f"Serial: {self.system_info['serial']}",
         text_color=COLOR_SUCCESS
      )
      self.info_manufacturer.configure(
         text=f"Fabricante: {self.system_info['manufacturer']}",
         text_color=COLOR_TEXT_WHITE
      )
      self.info_model.configure(
         text=f"Modelo: {self.system_info['model']}",
         text_color=COLOR_TEXT_WHITE
      )
      self.info_current_name.configure(
         text=f"Nombre actual: {self.system_info['current_name']}",
         text_color=COLOR_TEXT_WHITE
      )
      self.info_workgroup.configure(
         text=f"Grupo de trabajo: {self.system_info['current_workgroup']}",
         text_color=COLOR_TEXT_WHITE
      )
      
      self.log(f"✓ Serial: {self.system_info['serial']}", "SUCCESS")
      self.log(f"✓ Fabricante: {self.system_info['manufacturer']}", "SUCCESS")
      self.log(f"✓ Nombre actual: {self.system_info['current_name']}", "SUCCESS")
      self.log(f"✓ Grupo actual: {self.system_info['current_workgroup']}", "SUCCESS")
      
      # Actualizar preview inicial
      self.update_preview()
      
      self.log("━" * 75, "INFO")
      self.log("✓ Sistema listo para ejecutar cambios", "SUCCESS")
      self.update_status("Listo para ejecutar", COLOR_SUCCESS)
   
   def update_preview(self):
      """Actualiza la vista previa del nuevo nombre."""
      if not self.system_info:
         return
      
      suffix = self.suffix_var.get()
      new_name = build_new_hostname(self.system_info['serial'], suffix)
      
      is_valid, error = validate_hostname(new_name)
      
      if is_valid:
         self.preview_name.configure(text=new_name, text_color=COLOR_SUCCESS)
      else:
         self.preview_name.configure(
               text=f"{new_name} ⚠ INVÁLIDO",
               text_color=COLOR_ERROR
         )
         self.log(f"⚠ Nombre generado inválido: {error}", "WARNING")
   
   def validate_workgroup_input(self):
      """Valida la entrada del grupo de trabajo en tiempo real."""
      workgroup = self.workgroup_entry.get().strip()
      
      if not workgroup:
         self.workgroup_validation.configure(text="")
         return
      
      is_valid, error = validate_workgroup_name(workgroup)
      
      if is_valid:
         self.workgroup_validation.configure(
               text="✓ Nombre válido",
               text_color=COLOR_SUCCESS
         )
      else:
         self.workgroup_validation.configure(
               text=f"✗ {error}",
               text_color=COLOR_ERROR
         )
   
   def on_execute(self):
      """Maneja la ejecución del proceso de renombrado."""
      if self.is_processing:
         self.log("⚠ Ya hay un proceso en ejecución", "WARNING")
         return
      
      if not self.system_info:
         messagebox.showerror(
               "Error",
               "No hay información del sistema disponible.\n"
               "Por favor, reinicie la aplicación."
         )
         return
      
      workgroup = self.workgroup_entry.get().strip()
      
      if not workgroup:
         messagebox.showerror(
               "Campo Requerido",
               "Debe ingresar un nombre de grupo de trabajo (ciudad)."
         )
         self.workgroup_entry.focus()
         return
      
      # Validar workgroup
      is_valid, error = validate_workgroup_name(workgroup)
      if not is_valid:
         messagebox.showerror(
               "Error de Validación",
               f"El nombre de grupo de trabajo es inválido:\n\n{error}"
         )
         self.workgroup_entry.focus()
         return
      
      # Construir nuevo nombre
      suffix = self.suffix_var.get()
      new_name = build_new_hostname(self.system_info['serial'], suffix)
      
      # Validar nuevo nombre
      is_valid, error = validate_hostname(new_name)
      if not is_valid:
         messagebox.showerror(
               "Error de Validación",
               f"El nombre de equipo generado es inválido:\n\n{error}\n\n"
               f"Contacte al administrador del sistema."
         )
         return
      
      # Confirmar cambios con el usuario
      response = messagebox.askyesno(
         "⚠️ Confirmar Cambios Críticos",
         f"¿Confirma que desea aplicar los siguientes cambios?\n\n"
         f"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
         f"📌 NOMBRE DEL EQUIPO:\n"
         f"   Actual:  {self.system_info['current_name']}\n"
         f"   Nuevo:   {new_name}\n\n"
         f"📌 GRUPO DE TRABAJO:\n"
         f"   Actual:  {self.system_info['current_workgroup']}\n"
         f"   Nuevo:   {workgroup}\n\n"
         f"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
         f"⚠️ IMPORTANTE:\n"
         f"• El equipo se reiniciará automáticamente\n"
         f"• Guarde todo su trabajo antes de continuar\n"
         f"• Este proceso no se puede deshacer fácilmente\n\n"
         f"¿Desea continuar?",
         icon='warning'
      )
      
      if not response:
         self.log("✗ Operación cancelada por el usuario", "WARNING")
         return
      
      # Iniciar proceso en hilo separado
      self.is_processing = True
      self.btn_execute.configure(
         state="disabled",
         text="⏳ Procesando cambios...",
         fg_color=COLOR_BG_MEDIUM
      )
      self.clear_log()
      
      thread = threading.Thread(
         target=self.execute_rename,
         args=(new_name, workgroup),
         daemon=True
      )
      thread.start()
   
   def execute_rename(self, new_name, workgroup):
      """Ejecuta el proceso completo de renombrado del equipo."""
      try:
         self.log("━" * 75, "INFO")
         self.log("🚀 INICIANDO PROCESO DE RENOMBRADO DEL EQUIPO", "PROCESS")
         self.log("━" * 75, "INFO")
         self.update_status("Procesando cambios...", COLOR_WARNING)
         
         time.sleep(1)
         
         # Paso 1: Crear backup de configuración
         self.log("[1/4] Creando backup de configuración actual...", "PROCESS")
         backup_data = {
               "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
               "old_name": self.system_info['current_name'],
               "old_workgroup": self.system_info['current_workgroup'],
               "new_name": new_name,
               "new_workgroup": workgroup,
               "serial": self.system_info['serial'],
               "manufacturer": self.system_info['manufacturer'],
               "suffix": self.suffix_var.get()
         }
         
         if save_backup(backup_data):
               self.log(f"      ✓ Backup guardado en: {BACKUP_FILE}", "SUCCESS")
               self.log(f"      ✓ SHA-256: {backup_data.get('sha256', 'N/A')[:32]}...", "SUCCESS")
         else:
               self.log("      ⚠ No se pudo crear backup (continuando)", "WARNING")
         
         time.sleep(0.5)
         
         # Paso 2: Validar que el nuevo nombre sea diferente
         self.log("[2/4] Validando cambios necesarios...", "PROCESS")
         
         if new_name.upper() == self.system_info['current_name'].upper():
               self.log("      ℹ El nombre del equipo ya es correcto", "INFO")
               name_change_needed = False
         else:
               self.log(f"      ✓ Cambio de nombre requerido", "SUCCESS")
               name_change_needed = True
         
         if workgroup.upper() == self.system_info['current_workgroup'].upper():
               self.log("      ℹ El grupo de trabajo ya es correcto", "INFO")
               workgroup_change_needed = False
         else:
               self.log(f"      ✓ Cambio de grupo requerido", "SUCCESS")
               workgroup_change_needed = True
         
         if not name_change_needed and not workgroup_change_needed:
               self.log("      ℹ No hay cambios que aplicar", "INFO")
               raise Exception("El equipo ya tiene la configuración solicitada")
         
         time.sleep(0.5)
         
         # Paso 3: Aplicar cambios con PowerShell
         self.log("[3/4] Aplicando cambios en el sistema...", "PROCESS")
         self.log(f"      → Nuevo nombre: {new_name}", "INFO")
         self.log(f"      → Nuevo grupo: {workgroup}", "INFO")
         
         # Script PowerShell optimizado para renombrar
         ps_script = f'''
$ErrorActionPreference = 'Stop'

try {{
   # Renombrar equipo y cambiar grupo de trabajo
   Rename-Computer -NewName "{new_name}" -WorkGroupName "{workgroup}" -Force -PassThru -ErrorAction Stop | Out-Null
   
   Write-Output "SUCCESS: Cambios aplicados correctamente"
   exit 0
}} catch {{
   Write-Output "ERROR: $($_.Exception.Message)"
   exit 1
}}
'''
         
         success, output, error = run_powershell(ps_script, timeout=60)
         
         if not success or "ERROR:" in output:
               error_msg = output.split("ERROR:")[-1].strip() if "ERROR:" in output else error
               raise Exception(f"Error al aplicar cambios: {error_msg}")
         
         self.log("      ✓ Cambios aplicados correctamente en el sistema", "SUCCESS")
         time.sleep(0.5)
         
         # Paso 4: Preparar reinicio
         self.log("[4/4] Preparando reinicio del sistema...", "PROCESS")
         self.log("      ⏳ El equipo se reiniciará en 10 segundos", "WARNING")
         self.update_status("Reiniciando en 10 segundos...", COLOR_WARNING)
         
         # Countdown visual
         for i in range(10, 0, -1):
               self.log(f"      → Reiniciando en {i} segundo{'s' if i > 1 else ''}...", "INFO")
               time.sleep(1)
         
         self.log("━" * 75, "INFO")
         self.log("✓ PROCESO COMPLETADO EXITOSAMENTE", "SUCCESS")
         self.log("━" * 75, "INFO")
         self.log("🔄 Reiniciando equipo ahora...", "PROCESS")
         
         # Mostrar mensaje de éxito
         self.after(100, lambda: messagebox.showinfo(
               "✓ Cambios Aplicados Exitosamente",
               f"Los cambios han sido aplicados correctamente:\n\n"
               f"✓ Nombre: {new_name}\n"
               f"✓ Grupo: {workgroup}\n\n"
               f"El equipo se está reiniciando...",
               icon='info'
         ))
         
         # Ejecutar reinicio
         time.sleep(1)
         run_powershell("Restart-Computer -Force", timeout=10)
         
         self.update_status("Reinicio iniciado", COLOR_SUCCESS)
         
      except Exception as e:
         self.log("━" * 75, "ERROR")
         self.log(f"✗ ERROR EN EL PROCESO: {str(e)}", "ERROR")
         self.log("━" * 75, "ERROR")
         self.update_status("Error en el proceso", COLOR_ERROR)
         
         self.after(100, lambda: messagebox.showerror(
               "Error en Renombrado",
               f"No se pudieron aplicar los cambios:\n\n{str(e)}\n\n"
               "Posibles causas:\n"
               "• Permisos insuficientes\n"
               "• El equipo está en un dominio\n"
               "• Restricciones de políticas de grupo\n"
               "• Caracteres inválidos en el nombre\n\n"
               "Contacte al administrador del sistema."
         ))
      
      finally:
         # Restaurar estado de la interfaz
         self.is_processing = False
         self.after(100, lambda: self.btn_execute.configure(
               state="normal",
               text="🚀 Ejecutar Cambios de Renombrado",
               fg_color=COLOR_PRIMARY
         ))


# ============================================================================
# PUNTO DE ENTRADA PRINCIPAL
# ============================================================================

def main():
   """Función principal con verificación de privilegios de administrador."""
   
   # Verificar privilegios de administrador
   if not is_admin():
      response = messagebox.askyesno(
         "🔐 Privilegios de Administrador Requeridos",
         "Esta aplicación requiere privilegios de administrador\n"
         "para poder renombrar el equipo y modificar el grupo de trabajo.\n\n"
         "¿Desea reiniciar la aplicación como administrador?",
         icon='warning'
      )
      
      if response:
         run_as_admin()
      else:
         messagebox.showwarning(
               "Advertencia",
               "La aplicación NO funcionará correctamente sin privilegios de administrador.\n\n"
               "Los cambios de nombre y grupo de trabajo requieren permisos elevados."
         )
         # Permitir abrir la app para ver la interfaz, pero con funcionalidad limitada
   
   # Iniciar aplicación
   try:
      app = RenameApp()
      app.mainloop()
   except Exception as e:
      messagebox.showerror(
         "Error Fatal",
         f"No se pudo iniciar la aplicación:\n\n{e}\n\n"
         f"Versión: {__version__}\n"
         f"Autor: {__author__}\n"
         f"Contacto: {__company__}"
      )
      sys.exit(1)


if __name__ == "__main__":
   main()