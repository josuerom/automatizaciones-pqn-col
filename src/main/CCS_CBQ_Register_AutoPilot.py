"""
CCS_CBQ_Register_AutoPilot.py - Versión Mejorada
Autor: Josué Romero
Empresa: Stefanini / PQN
Fecha: 31/Octubre/2025

Descripción:
Instala el complemento y sube el dispositivo a la aplicación empresarial WindowsAutoPilotInfo.
"""

import subprocess
import customtkinter as ctk
import datetime
import threading
import time
from tkinter import messagebox
from pathlib import Path
import sys

# ============================================================================
# CONFIGURACIÓN GLOBAL
# ============================================================================
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# Constantes de la aplicación
APP_TITLE = "Registro AutoPilot CCS-CBQ"
APP_VERSION = "v2.0"
APP_SIZE = "650x650"
TEMP_DIR = Path("C:/Temp")
SCRIPT_NAME = "autopilot_register.ps1"

# Colores personalizados (Paleta profesional)
COLOR_PRIMARY = "#1e88e5"      # Azul principal
COLOR_SUCCESS = "#43a047"      # Verde éxito
COLOR_WARNING = "#fb8c00"      # Naranja advertencia
COLOR_ERROR = "#e53935"        # Rojo error
COLOR_BG_DARK = "#1a1a1a"      # Fondo oscuro
COLOR_BG_LIGHT = "#2d2d2d"     # Fondo claro
COLOR_TEXT = "#e0e0e0"         # Texto claro

# Fuentes
FONT_TITLE = ("Segoe UI", 24, "bold")
FONT_SUBTITLE = ("Segoe UI", 11)
FONT_CONSOLE = ("Consolas", 11)
FONT_BUTTON = ("Segoe UI", 12, "bold")


# ============================================================================
# FUNCIONES AUXILIARES
# ============================================================================

def ensure_temp_dir():
   """Asegura que existe el directorio temporal."""
   try:
      TEMP_DIR.mkdir(parents=True, exist_ok=True)
      return True
   except Exception as e:
      return False


def get_serial_number(max_retries=3):
   """
   Obtiene el número de serie del equipo usando PowerShell CIM.
   
   Args:
      max_retries: Número máximo de reintentos
      
   Returns:
      tuple: (success: bool, serial: str, error: str)
   """
   for attempt in range(max_retries):
      try:
         result = subprocess.run(
               [
                  "powershell", 
                  "-NoProfile",
                  "-Command", 
                  "(Get-CimInstance -ClassName Win32_BIOS).SerialNumber.Trim()"
               ],
               capture_output=True,
               text=True,
               check=True,
               timeout=10
         )
         serial = result.stdout.strip()
         
         if serial and serial != "":
               return True, serial, ""
         else:
               return False, "N/A", "Serial vacío"
               
      except subprocess.TimeoutExpired:
         if attempt < max_retries - 1:
               time.sleep(1)
               continue
         return False, "N/A", "Timeout al obtener serial"
      except Exception as e:
         if attempt < max_retries - 1:
               time.sleep(1)
               continue
         return False, "N/A", str(e)
   
   return False, "N/A", "Máximo de reintentos alcanzado"


def check_internet_connection():
   """Verifica si hay conexión a internet."""
   try:
      result = subprocess.run(
         ["ping", "-n", "1", "8.8.8.8"],
         capture_output=True,
         timeout=5
      )
      return result.returncode == 0
   except:
      return False


def check_powershell_version():
   """Verifica que PowerShell sea versión 5.1 o superior."""
   try:
      result = subprocess.run(
         ["powershell", "-Command", "$PSVersionTable.PSVersion.Major"],
         capture_output=True,
         text=True,
         timeout=5
      )
      version = int(result.stdout.strip())
      return version >= 5
   except:
      return False


# ============================================================================
# CLASE PRINCIPAL DE LA APLICACIÓN
# ============================================================================

class AutoPilotApp(ctk.CTk):
   def __init__(self):
      super().__init__()
      
      # Configuración de la ventana
      self.title(f"{APP_TITLE} {APP_VERSION}")
      self.geometry(APP_SIZE)
      self.resizable(False, False)
      
      # Variables de estado
      self.is_processing = False
      self.serial_number = None
      
      # Construir interfaz
      self.build_ui()
      
      # Verificar prerequisitos al iniciar
      self.after(500, self.check_prerequisites)
   
   def build_ui(self):
      """Construye la interfaz de usuario."""
      
      # Marco principal con padding
      main_frame = ctk.CTkFrame(self, fg_color=COLOR_BG_DARK)
      main_frame.pack(fill="both", expand=True, padx=15, pady=15)
      
      # === ENCABEZADO ===
      header_frame = ctk.CTkFrame(main_frame, fg_color=COLOR_PRIMARY, corner_radius=10)
      header_frame.pack(fill="x", pady=(0, 15))
      
      title_label = ctk.CTkLabel(
         header_frame,
         text="🚀 " + APP_TITLE,
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
         text="📋 Información del Sistema",
         font=("Segoe UI", 12, "bold"),
         text_color=COLOR_TEXT
      )
      info_title.pack(pady=(10, 5), anchor="w", padx=15)
      
      self.serial_label = ctk.CTkLabel(
         info_frame,
         text="Serial: Obteniendo...",
         font=FONT_SUBTITLE,
         text_color="#b0bec5",
         anchor="w"
      )
      self.serial_label.pack(pady=5, padx=15, anchor="w")
      
      self.status_label = ctk.CTkLabel(
         info_frame,
         text="Estado: Inicializando...",
         font=FONT_SUBTITLE,
         text_color="#b0bec5",
         anchor="w"
      )
      self.status_label.pack(pady=(0, 10), padx=15, anchor="w")
      
      # === ÁREA DE LOGS ===
      log_label = ctk.CTkLabel(
         main_frame,
         text="📝 Registro de Actividad",
         font=("Segoe UI", 12, "bold"),
         text_color=COLOR_TEXT,
         anchor="w"
      )
      log_label.pack(pady=(5, 5), anchor="w")
      
      self.output_box = ctk.CTkTextbox(
         main_frame,
         width=600,
         height=240,
         font=FONT_CONSOLE,
         fg_color=COLOR_BG_LIGHT,
         text_color=COLOR_TEXT,
         border_width=2,
         border_color=COLOR_PRIMARY,
         corner_radius=8
      )
      self.output_box.pack(pady=(0, 10))
      self.log("✓ Sistema inicializado correctamente")
      self.log("⚡ Esperando acción del usuario...")
      
      # === BARRA DE PROGRESO ===
      self.progress_bar = ctk.CTkProgressBar(
         main_frame,
         width=600,
         height=8,
         corner_radius=4,
         progress_color=COLOR_PRIMARY
      )
      self.progress_bar.pack(pady=(0, 15))
      self.progress_bar.set(0)
      
      # === BOTONES ===
      button_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
      button_frame.pack(fill="x")
      
      self.run_button = ctk.CTkButton(
         button_frame,
         text="▶ Ejecutar Registro",
         command=self.on_execute_clicked,
         font=FONT_BUTTON,
         height=40,
         corner_radius=8,
         fg_color=COLOR_PRIMARY,
         hover_color="#1565c0",
         border_width=2,
         border_color="#1976d2"
      )
      self.run_button.pack(side="left", expand=True, fill="x", padx=(0, 5))
      
      self.clear_button = ctk.CTkButton(
         button_frame,
         text="🗑 Limpiar Log",
         command=self.clear_log,
         font=FONT_BUTTON,
         height=40,
         corner_radius=8,
         fg_color=COLOR_BG_LIGHT,
         hover_color="#424242",
         border_width=2,
         border_color="#616161"
      )
      self.clear_button.pack(side="left", expand=True, fill="x", padx=(5, 0))
      
      # Atajos de teclado
      self.bind("<Return>", lambda e: self.on_execute_clicked())
      self.bind("<Escape>", lambda e: self.quit())
   
   def log(self, msg, level="INFO"):
      """
      Registra un mensaje en el output box con timestamp y color.
      
      Args:
         msg: Mensaje a registrar
         level: Nivel del log (INFO, SUCCESS, WARNING, ERROR)
      """
      timestamp = datetime.datetime.now().strftime("%H:%M:%S")
      
      # Mapeo de niveles a prefijos
      level_icons = {
         "INFO": "ℹ",
         "SUCCESS": "✓",
         "WARNING": "⚠",
         "ERROR": "✗"
      }
      
      icon = level_icons.get(level, "•")
      formatted_msg = f"[{timestamp}] {icon} {msg}\n"
      
      self.output_box.configure(state="normal")
      self.output_box.insert("end", formatted_msg)
      self.output_box.see("end")
      self.output_box.configure(state="disabled")
   
   def clear_log(self):
      """Limpia el área de logs."""
      self.output_box.configure(state="normal")
      self.output_box.delete("1.0", "end")
      self.output_box.configure(state="disabled")
      self.log("Log limpiado", "INFO")
   
   def update_status(self, text, color=COLOR_TEXT):
      """Actualiza el label de estado."""
      self.status_label.configure(text=f"Estado: {text}", text_color=color)
   
   def check_prerequisites(self):
      """Verifica los prerequisitos del sistema."""
      self.log("Verificando prerequisitos del sistema...", "INFO")
      
      # Verificar directorio temporal
      if not ensure_temp_dir():
         self.log("✗ No se pudo crear directorio temporal", "ERROR")
         self.update_status("Error: Sin acceso a C:/Temp", COLOR_ERROR)
         return
      
      # Obtener serial
      self.log("Obteniendo número de serie del equipo...", "INFO")
      success, serial, error = get_serial_number()
      
      if success:
         self.serial_number = serial
         self.serial_label.configure(text=f"Serial: {serial}", text_color=COLOR_SUCCESS)
         self.log(f"✓ Serial detectado: {serial}", "SUCCESS")
      else:
         self.serial_label.configure(text=f"Serial: Error ({error})", text_color=COLOR_ERROR)
         self.log(f"✗ Error al obtener serial: {error}", "ERROR")
      
      # Verificar PowerShell
      self.log("Verificando versión de PowerShell...", "INFO")
      if check_powershell_version():
         self.log("✓ PowerShell 5.1+ detectado", "SUCCESS")
      else:
         self.log("⚠ PowerShell 5.1+ no detectado", "WARNING")
      
      # Verificar conexión
      self.log("Verificando conexión a internet...", "INFO")
      if check_internet_connection():
         self.log("✓ Conexión a internet activa", "SUCCESS")
         self.update_status("Listo para ejecutar", COLOR_SUCCESS)
      else:
         self.log("⚠ Sin conexión a internet", "WARNING")
         self.update_status("Advertencia: Sin conexión", COLOR_WARNING)
      
      self.log("━" * 60, "INFO")
      self.log("✓ Verificación de prerequisitos completada", "SUCCESS")
   
   def on_execute_clicked(self):
      """Maneja el clic en el botón ejecutar."""
      if self.is_processing:
         self.log("⚠ Ya hay un proceso en ejecución", "WARNING")
         return
      
      # Confirmar acción
      response = messagebox.askyesno(
         "Confirmar Acción",
         "¿Desea iniciar el registro en AutoPilot?\n\n"
         "Este proceso:\n"
         "• Instalará el script Get-WindowsAutoPilotInfo\n"
         "• Subirá la información del dispositivo\n"
         "• Requerirá confirmación en PowerShell\n\n"
         "¿Continuar?"
      )
      
      if not response:
         self.log("✗ Operación cancelada por el usuario", "WARNING")
         return
      
      # Iniciar proceso en hilo separado
      self.is_processing = True
      self.run_button.configure(state="disabled", text="⏳ Procesando...")
      self.update_status("Ejecutando...", COLOR_WARNING)
      
      thread = threading.Thread(target=self.execute_autopilot_registration, daemon=True)
      thread.start()
   
   def execute_autopilot_registration(self):
      """Ejecuta el proceso de registro en AutoPilot."""
      try:
         self.log("━" * 60, "INFO")
         self.log("🚀 Iniciando proceso de registro AutoPilot", "INFO")
         self.log("━" * 60, "INFO")
         
         # Animar barra de progreso
         self.animate_progress(0.2)
         
         # Crear script de PowerShell
         self.log("Creando script de PowerShell...", "INFO")
         ps_script = """
# Script de registro AutoPilot
Write-Host "Instalando Get-WindowsAutoPilotInfo..." -ForegroundColor Cyan

try {
   Install-Script -Name Get-WindowsAutoPilotInfo -Force -Scope CurrentUser -Confirm:$false
   Write-Host "✓ Script instalado correctamente" -ForegroundColor Green
   
   Write-Host "`nSubiendo información a AutoPilot..." -ForegroundColor Cyan
   Get-WindowsAutoPilotInfo -Online
   
   Write-Host "`n✓ Proceso completado" -ForegroundColor Green
} catch {
   Write-Host "`n✗ Error: $($_.Exception.Message)" -ForegroundColor Red
   exit 1
}
"""
         
         script_path = TEMP_DIR / SCRIPT_NAME
         
         try:
               with open(script_path, "w", encoding="utf-8") as f:
                  f.write(ps_script)
               self.log(f"✓ Script creado en: {script_path}", "SUCCESS")
         except Exception as e:
               self.log(f"✗ Error al crear script: {e}", "ERROR")
               raise
         
         self.animate_progress(0.5)
         
         # Ejecutar PowerShell
         self.log("Abriendo ventana de PowerShell...", "INFO")
         self.log("⚠ IMPORTANTE: Cuando se solicite, confirme con 'Y' o 'S'", "WARNING")
         
         try:
               process = subprocess.Popen(
                  [
                     "powershell.exe",
                     "-ExecutionPolicy", "Bypass",
                     "-NoExit",
                     "-File", str(script_path)
                  ],
                  creationflags=subprocess.CREATE_NEW_CONSOLE
               )
               
               self.log("✓ Ventana de PowerShell abierta correctamente", "SUCCESS")
               self.log("📌 Continúe el proceso en la ventana de PowerShell", "INFO")
               self.animate_progress(1.0)
               
               # Mostrar mensaje informativo
               self.after(100, lambda: messagebox.showinfo(
                  "Proceso Iniciado",
                  "✓ Ventana de PowerShell abierta\n\n"
                  "📋 INSTRUCCIONES:\n"
                  "1. En la ventana de PowerShell, confirme con 'Y' o 'S'\n"
                  "2. Espere a que se complete la subida\n"
                  "3. Verifique el mensaje de éxito\n\n"
                  "⚠ No cierre esta aplicación hasta completar el proceso"
               ))
               
               self.log("━" * 60, "INFO")
               self.log("✓ Proceso iniciado exitosamente", "SUCCESS")
               self.update_status("Proceso en PowerShell", COLOR_SUCCESS)
               
         except Exception as e:
               self.log(f"✗ Error al abrir PowerShell: {e}", "ERROR")
               raise
      
      except Exception as e:
         self.log("━" * 60, "INFO")
         self.log(f"✗ ERROR CRÍTICO: {str(e)}", "ERROR")
         self.update_status("Error crítico", COLOR_ERROR)
         self.after(100, lambda: messagebox.showerror(
               "Error",
               f"No se pudo completar el proceso:\n\n{str(e)}"
         ))
      
      finally:
         # Restaurar estado de la interfaz
         self.is_processing = False
         self.after(100, lambda: self.run_button.configure(
               state="normal",
               text="▶ Ejecutar Registro"
         ))
   
   def animate_progress(self, target_value):
      """Anima la barra de progreso hasta el valor objetivo."""
      current = self.progress_bar.get()
      step = 0.05
      
      while current < target_value:
         current = min(current + step, target_value)
         self.progress_bar.set(current)
         self.update_idletasks()
         time.sleep(0.05)


# ============================================================================
# PUNTO DE ENTRADA
# ============================================================================

if __name__ == "__main__":
   try:
      app = AutoPilotApp()
      app.mainloop()
   except Exception as e:
      print(f"Error fatal: {e}")
      messagebox.showerror("Error Fatal", f"No se pudo iniciar la aplicación:\n\n{e}")
      sys.exit(1)