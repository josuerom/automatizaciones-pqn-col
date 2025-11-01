"""
PQN_COL_Equipment_Renamer.py - Versión Mejorada
Autor: Josué Romero
Empresa: Stefanini / PQN
Fecha: 31/Octubre/2025

Descripción:
Renombra equipos Windows basándose en el serial del BIOS con formato: XXXXXXX-PQN-COL
Requiere privilegios de administrador.
"""

import sys
import os
import subprocess
import ctypes
import time
import re
import json
import customtkinter as ctk
from tkinter import messagebox
from datetime import datetime
from pathlib import Path
import threading

# ============================================================================
# CONFIGURACIÓN GLOBAL
# ============================================================================
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

APP_TITLE = "Renombrador de Equipos PQN-COL"
APP_VERSION = "v2.0"
APP_SIZE = "750x850"

# Colores
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

# Configuración
LOG_DIR = Path("C:/ProgramData/PQN_COL_Renamer")
LOG_FILE = LOG_DIR / "renamer.log"
BACKUP_FILE = LOG_DIR / "backup_config.json"
SUFFIX = "-PQN-COL"
SERIAL_LENGTH = 7
MAX_NETBIOS_LENGTH = 15


# ============================================================================
# FUNCIONES AUXILIARES
# ============================================================================

def setup_logging():
   """Crea el directorio de logs."""
   try:
      LOG_DIR.mkdir(parents=True, exist_ok=True)
      return True
   except:
      return False


def log_to_file(message, level="INFO"):
   """Registra en archivo de log."""
   try:
      timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
      log_entry = f"[{timestamp}] [{level}] {message}\n"
      with open(LOG_FILE, "a", encoding="utf-8") as f:
         f.write(log_entry)
   except:
      pass


def is_admin():
   """Verifica privilegios de administrador."""
   try:
      return ctypes.windll.shell32.IsUserAnAdmin()
   except:
      return False


def run_powershell(command, timeout=30):
   """
   Ejecuta comando PowerShell.
   
   Returns:
      tuple: (success: bool, output: str, error: str)
   """
   try:
      result = subprocess.run(
         ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", command],
         capture_output=True,
         text=True,
         timeout=timeout
      )
      success = result.returncode == 0
      return success, result.stdout.strip(), result.stderr.strip()
   except subprocess.TimeoutExpired:
      return False, "", "Timeout"
   except Exception as e:
      return False, "", str(e)


def get_bios_serial():
   """Obtiene serial del BIOS usando PowerShell CIM."""
   success, output, error = run_powershell(
      "(Get-CimInstance -ClassName Win32_BIOS).SerialNumber.Trim()"
   )
   if success and output:
      serial = re.sub(r'\s+', '', output).strip()
      return serial if serial else None
   return None


def get_current_hostname():
   """Obtiene el nombre actual del equipo."""
   success, output, _ = run_powershell("$env:COMPUTERNAME")
   return output.strip().upper() if success else None


def get_manufacturer():
   """Obtiene el fabricante del equipo."""
   success, output, _ = run_powershell(
      "(Get-CimInstance -ClassName Win32_ComputerSystem).Manufacturer.Trim()"
   )
   return output.strip() if success else "Desconocido"


def validate_serial(serial):
   """Valida el serial."""
   if not serial:
      return False
   return bool(re.match(r'^[A-Za-z0-9\-]+$', serial))


def validate_hostname(name):
   """Valida el nombre del equipo."""
   if not name:
      return False, "Nombre vacío"
   
   if len(name) > MAX_NETBIOS_LENGTH:
      return False, f"Excede {MAX_NETBIOS_LENGTH} caracteres"
   
   if not re.match(r'^[A-Za-z0-9][A-Za-z0-9\-]*[A-Za-z0-9]$', name) and len(name) > 1:
      return False, "Caracteres inválidos"
   
   return True, ""


def build_hostname(serial):
   """Construye el nuevo hostname."""
   if not validate_serial(serial):
      return None
   
   if len(serial) >= SERIAL_LENGTH:
      serial_part = serial[-SERIAL_LENGTH:]
   else:
      serial_part = serial.zfill(SERIAL_LENGTH)
   
   hostname = f"{serial_part}{SUFFIX}".upper()
   
   if len(hostname) > MAX_NETBIOS_LENGTH:
      hostname = hostname[:MAX_NETBIOS_LENGTH]
   
   is_valid, error = validate_hostname(hostname)
   return hostname if is_valid else None


def save_backup(data):
   """Guarda backup de configuración."""
   try:
      with open(BACKUP_FILE, 'w', encoding='utf-8') as f:
         json.dump(data, f, indent=2)
      return True
   except:
      return False


def rename_computer(new_name):
   """Renombra el equipo."""
   command = f'Rename-Computer -NewName "{new_name}" -Force -PassThru -ErrorAction Stop'
   success, output, error = run_powershell(command, timeout=60)
   return success


def restart_computer(delay=15):
   """Reinicia el equipo después del delay especificado."""
   time.sleep(delay)
   run_powershell("Restart-Computer -Force")


# ============================================================================
# CLASE PRINCIPAL
# ============================================================================

class RenamerApp(ctk.CTk):
   def __init__(self):
      super().__init__()
      
      # Configuración
      self.title(f"{APP_TITLE} {APP_VERSION}")
      self.geometry(APP_SIZE)
      self.resizable(False, False)
      
      # Variables
      self.system_info = None
      self.is_processing = False
      
      # Construir interfaz
      self.build_ui()
      
      # Verificar prerequisitos
      self.after(500, self.check_prerequisites)
   
   def build_ui(self):
      """Construye la interfaz."""
      
      # Marco principal
      main_frame = ctk.CTkFrame(self, fg_color=COLOR_BG_DARK)
      main_frame.pack(fill="both", expand=True, padx=15, pady=15)
      
      # === ENCABEZADO ===
      header_frame = ctk.CTkFrame(main_frame, fg_color=COLOR_PRIMARY, corner_radius=10)
      header_frame.pack(fill="x", pady=(0, 15))
      
      title_label = ctk.CTkLabel(
         header_frame,
         text="💻 " + APP_TITLE,
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
      
      # === INFORMACIÓN ACTUAL ===
      info_frame = ctk.CTkFrame(main_frame, fg_color=COLOR_BG_LIGHT, corner_radius=8)
      info_frame.pack(fill="x", pady=(0, 10))
      
      info_title = ctk.CTkLabel(
         info_frame,
         text="📋 Información Actual del Sistema",
         font=FONT_LABEL,
         text_color=COLOR_TEXT
      )
      info_title.pack(pady=(10, 10), anchor="w", padx=15)
      
      self.info_serial = ctk.CTkLabel(
         info_frame,
         text="Serial BIOS: Obteniendo...",
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
      self.info_current_name.pack(pady=(3, 10), padx=15, anchor="w")
      
      # === NUEVO NOMBRE ===
      preview_frame = ctk.CTkFrame(main_frame, fg_color=COLOR_BG_LIGHT, corner_radius=8)
      preview_frame.pack(fill="x", pady=(0, 10))
      
      preview_title = ctk.CTkLabel(
         preview_frame,
         text="🎯 Nuevo Nombre (Preview)",
         font=FONT_LABEL,
         text_color=COLOR_TEXT
      )
      preview_title.pack(pady=(10, 10), anchor="w", padx=15)
      
      self.preview_name = ctk.CTkLabel(
         preview_frame,
         text="--------",
         font=("Consolas", 16, "bold"),
         text_color=COLOR_PRIMARY
      )
      self.preview_name.pack(pady=(0, 5), padx=15, anchor="w")
      
      self.preview_validation = ctk.CTkLabel(
         preview_frame,
         text="",
         font=("Segoe UI", 10),
         text_color=COLOR_SUCCESS
      )
      self.preview_validation.pack(pady=(0, 10), padx=15, anchor="w")
      
      # === LOG ===
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
         width=700,
         height=250,
         font=FONT_CONSOLE,
         fg_color=COLOR_BG_LIGHT,
         text_color=COLOR_TEXT,
         border_width=2,
         border_color=COLOR_PRIMARY,
         corner_radius=8
      )
      self.text_log.pack(pady=(0, 10))
      
      # === ESTADO ===
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
         text="▶ Aplicar Cambios y Reiniciar",
         command=self.on_execute,
         font=FONT_BUTTON,
         height=40,
         corner_radius=8,
         fg_color=COLOR_PRIMARY,
         hover_color="#1976d2"
      )
      self.btn_execute.pack(side="left", expand=True, fill="x", padx=(0, 5))
      
      self.btn_export = ctk.CTkButton(
         button_frame,
         text="💾 Exportar Log",
         command=self.export_log,
         font=FONT_BUTTON,
         height=40,
         corner_radius=8,
         fg_color=COLOR_BG_LIGHT,
         hover_color="#424242"
      )
      self.btn_export.pack(side="left", expand=True, fill="x", padx=(5, 0))
   
   def log(self, message, level="INFO"):
      """Registra mensaje en el log."""
      icons = {"INFO": "ℹ", "SUCCESS": "✓", "WARNING": "⚠", "ERROR": "✗"}
      icon = icons.get(level, "•")
      
      timestamp = datetime.now().strftime("%H:%M:%S")
      
      self.text_log.configure(state="normal")
      self.text_log.insert("end", f"[{timestamp}] {icon} {message}\n")
      self.text_log.see("end")
      self.text_log.configure(state="disabled")
      
      # También registrar en archivo
      log_to_file(message, level)
   
   def update_status(self, text, color=COLOR_TEXT):
      """Actualiza el estado."""
      self.status_label.configure(text=f"Estado: {text}", text_color=color)
   
   def check_prerequisites(self):
      """Verifica prerequisitos."""
      setup_logging()
      self.log("Verificando prerequisitos del sistema...")
      
      # Verificar privilegios
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
      
      serial = get_bios_serial()
      if not serial:
         self.log("✗ No se pudo obtener el serial del BIOS", "ERROR")
         self.update_status("Error: Sin serial", COLOR_ERROR)
         self.btn_execute.configure(state="disabled")
         return
      
      current_name = get_current_hostname()
      manufacturer = get_manufacturer()
      new_name = build_hostname(serial)
      
      if not new_name:
         self.log("✗ No se pudo construir un nombre válido", "ERROR")
         self.update_status("Error: Nombre inválido", COLOR_ERROR)
         self.btn_execute.configure(state="disabled")
         return
      
      self.system_info = {
         "serial": serial,
         "manufacturer": manufacturer,
         "current_name": current_name,
         "new_name": new_name
      }
      
      # Actualizar interfaz
      self.info_serial.configure(
         text=f"Serial BIOS: {serial}",
         text_color=COLOR_SUCCESS
      )
      self.info_manufacturer.configure(
         text=f"Fabricante: {manufacturer}",
         text_color=COLOR_TEXT
      )
      self.info_current_name.configure(
         text=f"Nombre actual: {current_name}",
         text_color=COLOR_TEXT
      )
      
      self.preview_name.configure(text=new_name)
      
      # Verificar si ya tiene el nombre correcto
      if current_name == new_name:
         self.log("✓ El equipo ya tiene el nombre correcto", "SUCCESS")
         self.preview_validation.configure(
               text="✓ El equipo ya está correctamente nombrado (no se requieren cambios)",
               text_color=COLOR_SUCCESS
         )
         self.update_status("No se requieren cambios", COLOR_SUCCESS)
         self.btn_execute.configure(state="disabled")
      else:
         self.log(f"✓ Nombre actual: {current_name}", "INFO")
         self.log(f"✓ Nuevo nombre: {new_name}", "INFO")
         self.preview_validation.configure(
               text=f"✓ El nombre cambiará de '{current_name}' a '{new_name}'",
               text_color=COLOR_WARNING
         )
         self.update_status("Listo para renombrar", COLOR_SUCCESS)
      
      self.log("━" * 70)
      self.log("✓ Sistema listo", "SUCCESS")
   
   def export_log(self):
      """Exporta el log a un archivo."""
      try:
         export_path = Path.home() / "Desktop" / f"Renamer_Log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
         
         log_content = self.text_log.get("1.0", "end")
         
         with open(export_path, 'w', encoding='utf-8') as f:
               f.write(f"LOG DEL RENOMBRADOR PQN-COL\n")
               f.write(f"Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
               f.write("=" * 70 + "\n\n")
               f.write(log_content)
         
         self.log(f"✓ Log exportado a: {export_path}", "SUCCESS")
         messagebox.showinfo("Éxito", f"Log exportado correctamente:\n\n{export_path}")
      except Exception as e:
         self.log(f"✗ Error al exportar log: {str(e)}", "ERROR")
         messagebox.showerror("Error", f"No se pudo exportar el log:\n\n{str(e)}")
   
   def on_execute(self):
      """Ejecuta el proceso de renombrado."""
      if self.is_processing:
         return
      
      if not self.system_info:
         messagebox.showerror("Error", "No hay información del sistema disponible.")
         return
      
      new_name = self.system_info['new_name']
      current_name = self.system_info['current_name']
      
      # Confirmar
      response = messagebox.askyesno(
         "Confirmar Renombrado",
         f"¿Confirma el cambio de nombre?\n\n"
         f"Nombre actual: {current_name}\n"
         f"Nombre nuevo:  {new_name}\n\n"
         f"⚠ El equipo se reiniciará automáticamente en 15 segundos.\n"
         f"⚠ Guarde todo su trabajo antes de continuar.\n\n"
         f"¿Continuar?"
      )
      
      if not response:
         self.log("✗ Operación cancelada por el usuario", "WARNING")
         return
      
      # Ejecutar en hilo separado
      self.is_processing = True
      self.btn_execute.configure(state="disabled", text="⏳ Procesando...")
      
      thread = threading.Thread(
         target=self.execute_rename,
         args=(new_name,),
         daemon=True
      )
      thread.start()
   
   def execute_rename(self, new_name):
      """Ejecuta el renombrado."""
      try:
         self.log("━" * 70)
         self.log("🚀 Iniciando proceso de renombrado", "INFO")
         self.log("━" * 70)
         
         # Crear backup
         self.log("Creando backup de configuración...")
         backup_data = {
               "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
               "old_name": self.system_info['current_name'],
               "new_name": new_name,
               "serial": self.system_info['serial']
         }
         
         if save_backup(backup_data):
               self.log("✓ Backup creado correctamente", "SUCCESS")
         else:
               self.log("⚠ No se pudo crear backup", "WARNING")
         
         # Renombrar
         self.log(f"Aplicando nuevo nombre: {new_name}...")
         
         if rename_computer(new_name):
               self.log("✓ Nombre aplicado correctamente", "SUCCESS")
               self.log("⏳ El equipo se reiniciará en 15 segundos...", "WARNING")
               
               self.after(100, lambda: messagebox.showinfo(
                  "Éxito",
                  f"✓ Cambio aplicado correctamente\n\n"
                  f"Nuevo nombre: {new_name}\n\n"
                  f"El equipo se reiniciará en 15 segundos.\n"
                  f"Guarde todo su trabajo."
               ))
               
               # Reiniciar
               restart_computer(15)
         else:
               raise Exception("Error al aplicar el nuevo nombre")
      
      except Exception as e:
         self.log(f"✗ Error: {str(e)}", "ERROR")
         self.after(100, lambda: messagebox.showerror(
               "Error",
               f"No se pudo completar el renombrado:\n\n{str(e)}"
         ))
      finally:
         self.is_processing = False
         self.after(100, lambda: self.btn_execute.configure(
               state="normal",
               text="▶ Aplicar Cambios y Reiniciar"
         ))


# ============================================================================
# PUNTO DE ENTRADA
# ============================================================================

if __name__ == "__main__":
   app = RenamerApp()
   app.mainloop()