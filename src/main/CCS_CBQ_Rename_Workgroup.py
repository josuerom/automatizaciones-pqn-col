"""
CCS_CBQ_Rename_Workgroup.py - Versión Mejorada
Autor: Josué Romero
Empresa: Stefanini / PQN
Fecha: 31/Octubre/2025

Descripción:
Renombra el equipo siguiendo el formato [7ULTIMOSDIGITOSDELSN]-[CCS/CBQ]-COL
y cambia el grupo de trabajo al nombre de la ciudad.
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

# ============================================================================
# CONFIGURACIÓN GLOBAL
# ============================================================================
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

APP_TITLE = "Renombrador de Equipos CCS-CBQ"
APP_VERSION = "v2.0"
APP_SIZE = "700x940"

# Colores profesionales
COLOR_PRIMARY = "#2196f3"
COLOR_SUCCESS = "#4caf50"
COLOR_WARNING = "#ff9800"
COLOR_ERROR = "#f44336"
COLOR_BG_DARK = "#1a1a1a"
COLOR_BG_LIGHT = "#2d2d2d"
COLOR_TEXT = "#e0e0e0"

# Fuentes
FONT_TITLE = ("Segoe UI", 26, "bold")
FONT_SUBTITLE = ("Segoe UI", 11)
FONT_CONSOLE = ("Consolas", 10)
FONT_BUTTON = ("Segoe UI", 12, "bold")
FONT_LABEL = ("Segoe UI", 11, "bold")

# Constantes de validación
MAX_NETBIOS_LENGTH = 15
SERIAL_LENGTH = 7
VALID_SUFFIXES = ["CCS", "CBQ"]
COUNTRY_CODE = "COL"

# Archivo de backup
BACKUP_FILE = Path("C:/ProgramData/CCS_CBQ_Renamer/backup_config.json")


# ============================================================================
# FUNCIONES AUXILIARES
# ============================================================================

def is_admin():
   """Verifica si el script tiene privilegios de administrador."""
   try:
      return ctypes.windll.shell32.IsUserAnAdmin()
   except:
      return False


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
         shell=True
      )
      
      success = result.returncode == 0
      output = result.stdout.strip()
      error = result.stderr.strip()
      
      return success, output, error
   except subprocess.TimeoutExpired:
      return False, "", "Timeout ejecutando comando"
   except Exception as e:
      return False, "", str(e)


def get_system_info():
   """
   Obtiene información del sistema usando PowerShell CIM.
   
   Returns:
      dict: Información del sistema o None si falla
   """
   try:
      # Obtener serial usando CIM
      success, serial, error = run_powershell(
         "(Get-CimInstance -ClassName Win32_BIOS).SerialNumber.Trim()"
      )
      if not success:
         return None
      
      # Obtener fabricante
      success, manufacturer, _ = run_powershell(
         "(Get-CimInstance -ClassName Win32_ComputerSystem).Manufacturer.Trim()"
      )
      
      # Obtener nombre actual
      success, current_name, _ = run_powershell(
         "$env:COMPUTERNAME"
      )
      
      # Obtener grupo de trabajo actual
      success, workgroup, _ = run_powershell(
         "(Get-CimInstance -ClassName Win32_ComputerSystem).Workgroup"
      )
      
      return {
         "serial": serial.strip(),
         "manufacturer": manufacturer.strip() if manufacturer else "Desconocido",
         "current_name": current_name.strip().upper(),
         "current_workgroup": workgroup.strip() if workgroup else "WORKGROUP"
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
      return False, "El nombre solo puede contener letras, números y guiones"
   
   if name.startswith('-') or name.endswith('-'):
      return False, "El nombre no puede empezar o terminar con guión"
   
   if '--' in name:
      return False, "El nombre no puede contener guiones consecutivos"
   
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
   
   # Permitir espacios en workgroups pero no al inicio/final
   name_stripped = name.strip()
   if len(name_stripped) != len(name):
      return False, "El nombre no puede tener espacios al inicio o final"
   
   if not re.match(r'^[A-Za-z0-9\s\-\.]+$', name):
      return False, "El nombre contiene caracteres no permitidos"
   
   return True, ""


def build_new_hostname(serial, suffix):
   """
   Construye el nuevo nombre del equipo.
   
   Args:
      serial: Serial del BIOS
      suffix: Sufijo (CCS o CBQ)
      
   Returns:
      str: Nuevo nombre del equipo
   """
   # Tomar últimos 7 caracteres del serial
   last7 = serial[-SERIAL_LENGTH:].upper() if len(serial) >= SERIAL_LENGTH else serial.upper().zfill(SERIAL_LENGTH)
   
   # Construir nombre
   new_name = f"{last7}-{suffix}-{COUNTRY_CODE}"
   
   return new_name


def save_backup(data):
   """Guarda backup de configuración actual."""
   try:
      BACKUP_FILE.parent.mkdir(parents=True, exist_ok=True)
      with open(BACKUP_FILE, 'w', encoding='utf-8') as f:
         json.dump(data, f, indent=2)
      return True
   except:
      return False


# ============================================================================
# CLASE PRINCIPAL
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
      self.after(300, self.check_prerequisites)
   
   def build_ui(self):
      """Construye la interfaz de usuario."""
      
      # Marco principal
      main_frame = ctk.CTkFrame(self, fg_color=COLOR_BG_DARK)
      main_frame.pack(fill="both", expand=True, padx=15, pady=15)
      
      # === ENCABEZADO ===
      header_frame = ctk.CTkFrame(main_frame, fg_color=COLOR_PRIMARY, corner_radius=10)
      header_frame.pack(fill="x", pady=(0, 15))
      
      title_label = ctk.CTkLabel(
         header_frame,
         text="🖥️ " + APP_TITLE,
         font=FONT_TITLE,
         text_color="white"
      )
      title_label.pack(pady=15)
      
      subtitle_label = ctk.CTkLabel(
         header_frame,
         text=f"Autor: Josué Romero  |  Stefanini / PQN  |  {APP_VERSION}",
         font=FONT_SUBTITLE,
         text_color="#e3f2fd"
      )
      subtitle_label.pack(pady=(0, 15))
      
      # === INFORMACIÓN DEL SISTEMA ===
      info_frame = ctk.CTkFrame(main_frame, fg_color=COLOR_BG_LIGHT, corner_radius=8)
      info_frame.pack(fill="x", pady=(0, 10))
      
      info_title = ctk.CTkLabel(
         info_frame,
         text="📋 Información Actual del Sistema",
         font=FONT_LABEL,
         text_color=COLOR_TEXT
      )
      info_title.pack(pady=(10, 5), anchor="w", padx=15)
      
      self.info_serial = ctk.CTkLabel(
         info_frame,
         text="Serial: Obteniendo...",
         font=FONT_SUBTITLE,
         text_color="#b0bec5",
         anchor="w"
      )
      self.info_serial.pack(pady=3, padx=15, anchor="w")
      
      self.info_manufacturer = ctk.CTkLabel(
         info_frame,
         text="Fabricante: Obteniendo...",
         font=FONT_SUBTITLE,
         text_color="#b0bec5",
         anchor="w"
      )
      self.info_manufacturer.pack(pady=3, padx=15, anchor="w")
      
      self.info_current_name = ctk.CTkLabel(
         info_frame,
         text="Nombre actual: Obteniendo...",
         font=FONT_SUBTITLE,
         text_color="#b0bec5",
         anchor="w"
      )
      self.info_current_name.pack(pady=3, padx=15, anchor="w")
      
      self.info_workgroup = ctk.CTkLabel(
         info_frame,
         text="Grupo de trabajo: Obteniendo...",
         font=FONT_SUBTITLE,
         text_color="#b0bec5",
         anchor="w"
      )
      self.info_workgroup.pack(pady=(3, 10), padx=15, anchor="w")
      
      # === CONFIGURACIÓN ===
      config_frame = ctk.CTkFrame(main_frame, fg_color=COLOR_BG_LIGHT, corner_radius=8)
      config_frame.pack(fill="x", pady=(0, 10))
      
      config_title = ctk.CTkLabel(
         config_frame,
         text="⚙️ Configuración Nueva",
         font=FONT_LABEL,
         text_color=COLOR_TEXT
      )
      config_title.pack(pady=(10, 10), anchor="w", padx=15)
      
      # Selector de sufijo
      suffix_label = ctk.CTkLabel(
         config_frame,
         text="Tipo de equipo:",
         font=FONT_SUBTITLE,
         text_color=COLOR_TEXT
      )
      suffix_label.pack(pady=(0, 5), padx=15, anchor="w")
      
      self.suffix_var = ctk.StringVar(value="CCS")
      suffix_frame = ctk.CTkFrame(config_frame, fg_color="transparent")
      suffix_frame.pack(pady=(0, 10), padx=15, anchor="w")
      
      ctk.CTkRadioButton(
         suffix_frame,
         text="CCS",
         variable=self.suffix_var,
         value="CCS",
         font=FONT_SUBTITLE,
         command=self.update_preview
      ).pack(side="left", padx=(0, 20))
      
      ctk.CTkRadioButton(
         suffix_frame,
         text="CBQ",
         variable=self.suffix_var,
         value="CBQ",
         font=FONT_SUBTITLE,
         command=self.update_preview
      ).pack(side="left")
      
      # Entrada de ciudad/grupo de trabajo
      workgroup_label = ctk.CTkLabel(
         config_frame,
         text="Nombre de la ciudad (grupo de trabajo):",
         font=FONT_SUBTITLE,
         text_color=COLOR_TEXT
      )
      workgroup_label.pack(pady=(0, 5), padx=15, anchor="w")
      
      self.workgroup_entry = ctk.CTkEntry(
         config_frame,
         placeholder_text="Ej: Bogota D.C, Medellin, Cali",
         width=400,
         height=35,
         font=FONT_SUBTITLE
      )
      self.workgroup_entry.pack(pady=(0, 5), padx=15, anchor="w")
      self.workgroup_entry.bind("<KeyRelease>", lambda e: self.validate_workgroup_input())
      
      self.workgroup_validation = ctk.CTkLabel(
         config_frame,
         text="",
         font=("Segoe UI", 9),
         text_color=COLOR_WARNING
      )
      self.workgroup_validation.pack(pady=(0, 10), padx=15, anchor="w")
      
      # Preview del nuevo nombre
      preview_label = ctk.CTkLabel(
         config_frame,
         text="Vista previa del nuevo nombre:",
         font=FONT_SUBTITLE,
         text_color=COLOR_TEXT
      )
      preview_label.pack(pady=(0, 5), padx=15, anchor="w")
      
      self.preview_name = ctk.CTkLabel(
         config_frame,
         text="--------",
         font=("Consolas", 14, "bold"),
         text_color=COLOR_PRIMARY
      )
      self.preview_name.pack(pady=(0, 10), padx=15, anchor="w")
      
      # === ÁREA DE LOGS ===
      log_label = ctk.CTkLabel(
         main_frame,
         text="📝 Registro de Actividad",
         font=FONT_LABEL,
         text_color=COLOR_TEXT,
         anchor="w"
      )
      log_label.pack(pady=(5, 5), anchor="w")
      
      self.text_log = ctk.CTkTextbox(
         main_frame,
         width=650,
         height=150,
         font=FONT_CONSOLE,
         fg_color=COLOR_BG_LIGHT,
         text_color=COLOR_TEXT,
         border_width=2,
         border_color=COLOR_PRIMARY,
         corner_radius=8
      )
      self.text_log.pack(pady=(0, 10))
      
      # === BARRA DE ESTADO ===
      self.status_label = ctk.CTkLabel(
         main_frame,
         text="Estado: Inicializando...",
         font=FONT_SUBTITLE,
         text_color=COLOR_WARNING
      )
      self.status_label.pack(pady=(0, 10))
      
      # === BOTONES ===
      button_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
      button_frame.pack(fill="x")
      
      self.btn_execute = ctk.CTkButton(
         button_frame,
         text="▶ Ejecutar Cambios",
         command=self.on_execute,
         font=FONT_BUTTON,
         height=40,
         corner_radius=8,
         fg_color=COLOR_PRIMARY,
         hover_color="#1976d2"
      )
      self.btn_execute.pack(side="left", expand=True, fill="x", padx=(0, 5))
      
      self.btn_clear = ctk.CTkButton(
         button_frame,
         text="🗑 Limpiar Log",
         command=self.clear_log,
         font=FONT_BUTTON,
         height=40,
         corner_radius=8,
         fg_color=COLOR_BG_LIGHT,
         hover_color="#424242"
      )
      self.btn_clear.pack(side="left", expand=True, fill="x", padx=(5, 0))
      
      # Atajos de teclado
      self.bind("<Return>", lambda e: self.on_execute())
      self.bind("<Escape>", lambda e: self.quit())
   
   def log(self, message, level="INFO"):
      """Registra un mensaje en el log."""
      icons = {"INFO": "ℹ", "SUCCESS": "✓", "WARNING": "⚠", "ERROR": "✗"}
      icon = icons.get(level, "•")
      
      self.text_log.configure(state="normal")
      self.text_log.insert("end", f"{icon} {message}\n")
      self.text_log.see("end")
      self.text_log.configure(state="disabled")
   
   def clear_log(self):
      """Limpia el log."""
      self.text_log.configure(state="normal")
      self.text_log.delete("1.0", "end")
      self.text_log.configure(state="disabled")
      self.log("Log limpiado")
   
   def update_status(self, text, color=COLOR_TEXT):
      """Actualiza el estado."""
      self.status_label.configure(text=f"Estado: {text}", text_color=color)
   
   def check_prerequisites(self):
      """Verifica prerequisitos del sistema."""
      self.log("Verificando prerequisitos...")
      
      # Verificar permisos de administrador
      if not is_admin():
         self.log("✗ Se requieren privilegios de administrador", "ERROR")
         self.update_status("Error: Sin privilegios", COLOR_ERROR)
         messagebox.showerror(
               "Privilegios Insuficientes",
               "Esta aplicación requiere privilegios de administrador.\n\n"
               "Por favor, ejecute como administrador."
         )
         self.btn_execute.configure(state="disabled")
         return
      
      self.log("✓ Privilegios de administrador confirmados", "SUCCESS")
      
      # Obtener información del sistema
      self.log("Obteniendo información del sistema...")
      self.system_info = get_system_info()
      
      if not self.system_info:
         self.log("✗ No se pudo obtener información del sistema", "ERROR")
         self.update_status("Error: Sin información", COLOR_ERROR)
         return
      
      # Actualizar labels
      self.info_serial.configure(
         text=f"Serial: {self.system_info['serial']}",
         text_color=COLOR_SUCCESS
      )
      self.info_manufacturer.configure(
         text=f"Fabricante: {self.system_info['manufacturer']}",
         text_color=COLOR_TEXT
      )
      self.info_current_name.configure(
         text=f"Nombre actual: {self.system_info['current_name']}",
         text_color=COLOR_TEXT
      )
      self.info_workgroup.configure(
         text=f"Grupo de trabajo: {self.system_info['current_workgroup']}",
         text_color=COLOR_TEXT
      )
      
      self.log(f"✓ Serial: {self.system_info['serial']}", "SUCCESS")
      self.log(f"✓ Nombre actual: {self.system_info['current_name']}", "SUCCESS")
      
      # Actualizar preview
      self.update_preview()
      
      self.update_status("Listo para ejecutar", COLOR_SUCCESS)
   
   def update_preview(self):
      """Actualiza el preview del nuevo nombre."""
      if not self.system_info:
         return
      
      suffix = self.suffix_var.get()
      new_name = build_new_hostname(self.system_info['serial'], suffix)
      
      is_valid, error = validate_hostname(new_name)
      
      if is_valid:
         self.preview_name.configure(text=new_name, text_color=COLOR_SUCCESS)
      else:
         self.preview_name.configure(text=f"{new_name} (Inválido)", text_color=COLOR_ERROR)
   
   def validate_workgroup_input(self):
      """Valida la entrada del grupo de trabajo en tiempo real."""
      workgroup = self.workgroup_entry.get().strip()
      
      if not workgroup:
         self.workgroup_validation.configure(text="")
         return
      
      is_valid, error = validate_workgroup_name(workgroup)
      
      if is_valid:
         self.workgroup_validation.configure(text="✓ Válido", text_color=COLOR_SUCCESS)
      else:
         self.workgroup_validation.configure(text=f"✗ {error}", text_color=COLOR_ERROR)
   
   def on_execute(self):
      """Maneja la ejecución del cambio."""
      if self.is_processing:
         return
      
      workgroup = self.workgroup_entry.get().strip()
      
      if not workgroup:
         messagebox.showerror("Error", "Debe ingresar un nombre de grupo de trabajo.")
         return
      
      # Validar workgroup
      is_valid, error = validate_workgroup_name(workgroup)
      if not is_valid:
         messagebox.showerror("Error de Validación", f"Nombre de grupo inválido:\n{error}")
         return
      
      # Construir nuevo nombre
      suffix = self.suffix_var.get()
      new_name = build_new_hostname(self.system_info['serial'], suffix)
      
      # Validar nuevo nombre
      is_valid, error = validate_hostname(new_name)
      if not is_valid:
         messagebox.showerror("Error de Validación", f"Nombre de equipo inválido:\n{error}")
         return
      
      # Confirmar
      response = messagebox.askyesno(
         "Confirmar Cambios",
         f"¿Confirma los siguientes cambios?\n\n"
         f"Nombre actual: {self.system_info['current_name']}\n"
         f"Nombre nuevo: {new_name}\n\n"
         f"Grupo actual: {self.system_info['current_workgroup']}\n"
         f"Grupo nuevo: {workgroup}\n\n"
         f"El equipo se reiniciará automáticamente.\n\n"
         f"¿Continuar?"
      )
      
      if not response:
         self.log("✗ Operación cancelada", "WARNING")
         return
      
      # Ejecutar en hilo separado
      self.is_processing = True
      self.btn_execute.configure(state="disabled", text="⏳ Procesando...")
      self.clear_log()
      
      thread = threading.Thread(
         target=self.execute_rename,
         args=(new_name, workgroup),
         daemon=True
      )
      thread.start()
   
   def execute_rename(self, new_name, workgroup):
      """Ejecuta el proceso de renombrado."""
      try:
         self.log("━" * 65, "INFO")
         self.log("🚀 Iniciando proceso de renombrado", "INFO")
         self.log("━" * 65, "INFO")
         
         # Guardar backup
         self.log("Creando backup de configuración actual...")
         backup_data = {
               "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
               "old_name": self.system_info['current_name'],
               "old_workgroup": self.system_info['current_workgroup'],
               "new_name": new_name,
               "new_workgroup": workgroup
         }
         
         if save_backup(backup_data):
               self.log("✓ Backup creado correctamente", "SUCCESS")
         else:
               self.log("⚠ No se pudo crear backup", "WARNING")
         
         # Ejecutar renombrado
         self.log(f"Cambiando nombre a: {new_name}...")
         self.log(f"Cambiando grupo a: {workgroup}...")
         
         ps_script = f'''
try {{
   Rename-Computer -NewName "{new_name}" -WorkGroupName "{workgroup}" -Force -PassThru -ErrorAction Stop | Out-Null
   Write-Output "SUCCESS"
}} catch {{
   Write-Output "ERROR: $($_.Exception.Message)"
   exit 1
}}
'''
         
         success, output, error = run_powershell(ps_script, timeout=60)
         
         if success and "SUCCESS" in output:
               self.log("✓ Cambios aplicados correctamente", "SUCCESS")
               self.log("⏳ Esperando 10 segundos antes de reiniciar...")
               time.sleep(10)
               
               self.log("🔄 Reiniciando equipo...")
               run_powershell("Restart-Computer -Force")
               
               self.after(100, lambda: messagebox.showinfo(
                  "Éxito",
                  "Cambios aplicados correctamente.\n"
                  "El equipo se está reiniciando..."
               ))
         else:
               raise Exception(error if error else output)
      
      except Exception as e:
         self.log(f"✗ Error: {str(e)}", "ERROR")
         self.after(100, lambda: messagebox.showerror(
               "Error",
               f"No se pudieron aplicar los cambios:\n\n{str(e)}"
         ))
      finally:
         self.is_processing = False
         self.after(100, lambda: self.btn_execute.configure(
               state="normal",
               text="▶ Ejecutar Cambios"
         ))


# ============================================================================
# PUNTO DE ENTRADA
# ============================================================================

if __name__ == "__main__":
   app = RenameApp()
   app.mainloop()