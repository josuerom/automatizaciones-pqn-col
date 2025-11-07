"""
CCS_CBQ_Register_AutoPilot.py - Versión Profesional Mejorada
Autor: Josué Romero
Empresa: Stefanini / PQN
Fecha: 06/Noviembre/2025
Versión: 3.0 Professional Edition

Licencia: Propiedad de Stefanini / PQN - Todos los derechos reservados
Copyright © 2025 Josué Romero - Stefanini / PQN

Descripción:
Sistema automatizado para inscripción de dispositivos en Windows AutoPilot.
Ejecuta comandos PowerShell con elevación de privilegios automática.
"""

import subprocess
import customtkinter as ctk
import datetime
import threading
import time
import sys
import os
import ctypes
from tkinter import messagebox
from pathlib import Path
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

# Constantes de la aplicación
APP_TITLE = "AutoPilot Registration System"
APP_VERSION = f"v{__version__}"
APP_SIZE = "700x700"

# Paleta de colores moderna y profesional (Esquema Azul-Cyan-Oscuro)
COLOR_PRIMARY = "#0d47a1"         # Azul profundo principal
COLOR_SECONDARY = "#00acc1"       # Cyan brillante
COLOR_ACCENT = "#00e5ff"          # Cyan claro
COLOR_SUCCESS = "#00c853"         # Verde éxito
COLOR_WARNING = "#ffa726"         # Naranja advertencia
COLOR_ERROR = "#d32f2f"           # Rojo error
COLOR_BG_DARK = "#0a0e27"         # Fondo oscuro principal
COLOR_BG_MEDIUM = "#1a1f3a"       # Fondo medio
COLOR_BG_LIGHT = "#242b4d"        # Fondo claro
COLOR_TEXT_WHITE = "#ffffff"      # Texto blanco
COLOR_TEXT_GRAY = "#b0bec5"       # Texto gris claro

# Fuentes (Incrementadas +2px según requerimiento)
FONT_TITLE = ("Segoe UI", 26, "bold")          
FONT_SUBTITLE = ("Segoe UI", 14)               
FONT_CONSOLE = ("Consolas", 14)                
FONT_BUTTON = ("Segoe UI", 14, "bold")         
FONT_INFO = ("Segoe UI", 14)                   


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
         # Modo desarrollo (.py)
         ctypes.windll.shell32.ShellExecuteW(
               None, "runas", sys.executable, f'"{sys.argv[0]}"', None, 1
         )
      else:
         # Modo producción (.exe)
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

def calculate_file_hash(filepath):
   """Calcula el hash SHA-256 del archivo."""
   try:
      sha256_hash = hashlib.sha256()
      with open(filepath, "rb") as f:
         for byte_block in iter(lambda: f.read(4096), b""):
               sha256_hash.update(byte_block)
      return sha256_hash.hexdigest()
   except:
      return "N/A"


def get_serial_number():
   """Obtiene el número de serie del equipo."""
   try:
      result = subprocess.run(
         ["powershell", "-NoProfile", "-Command",
            "(Get-CimInstance -ClassName Win32_BIOS).SerialNumber.Trim()"],
         capture_output=True,
         text=True,
         timeout=10
      )
      serial = result.stdout.strip()
      return serial if serial else "N/A"
   except:
      return "N/A"


def check_internet():
   """Verifica conexión a internet."""
   try:
      result = subprocess.run(
         ["ping", "-n", "2", "spradling.group"],
         capture_output=True,
         timeout=5
      )
      return result.returncode == 0
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
      self.serial_number = get_serial_number()
      
      # Construir interfaz
      self.build_ui()
      
      # Verificación inicial
      self.after(500, self.initial_check)
   
   def build_ui(self):
      """Construye la interfaz de usuario moderna."""
      
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
         text="⚡ " + APP_TITLE,
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
         text="💻 Información del Sistema",
         font=FONT_INFO,
         text_color=COLOR_TEXT_WHITE
      )
      info_title.pack(pady=(15, 10), anchor="w", padx=20)
      
      self.serial_label = ctk.CTkLabel(
         info_frame,
         text=f"Serial Number: {self.serial_number}",
         font=FONT_INFO,
         text_color=COLOR_TEXT_WHITE,
         anchor="w"
      )
      self.serial_label.pack(pady=5, padx=20, anchor="w")
      
      self.status_label = ctk.CTkLabel(
         info_frame,
         text="Estado: Listo para ejecutar",
         font=FONT_INFO,
         text_color=COLOR_SUCCESS,
         anchor="w"
      )
      self.status_label.pack(pady=(0, 15), padx=20, anchor="w")
      
      # === ÁREA DE LOGS ===
      log_label = ctk.CTkLabel(
         main_frame,
         text="📝 Registro de Actividad",
         font=FONT_INFO,
         text_color=COLOR_TEXT_WHITE,
         anchor="w"
      )
      log_label.pack(pady=(5, 8), anchor="w")
      
      self.output_box = ctk.CTkTextbox(
         main_frame,
         width=640,
         height=220,
         font=FONT_CONSOLE,
         fg_color=COLOR_BG_LIGHT,
         text_color=COLOR_TEXT_WHITE,
         border_width=2,
         border_color=COLOR_SECONDARY,
         corner_radius=10
      )
      self.output_box.pack(pady=(0, 15))
      self.log("✓ Sistema inicializado correctamente", "SUCCESS")
      self.log(f"✓ Licencia: {__license__}", "INFO")
      
      # === BARRA DE PROGRESO ===
      self.progress_bar = ctk.CTkProgressBar(
         main_frame,
         width=640,
         height=10,
         corner_radius=5,
         progress_color=COLOR_SECONDARY,
         fg_color=COLOR_BG_MEDIUM
      )
      self.progress_bar.pack(pady=(0, 20))
      self.progress_bar.set(0)
      
      # === BOTÓN PRINCIPAL (ÚNICO) ===
      self.run_button = ctk.CTkButton(
         main_frame,
         text="🚀 Inscribir Equipo Ahora",
         command=self.on_execute_clicked,
         font=FONT_BUTTON,
         height=50,
         corner_radius=10,
         fg_color=COLOR_PRIMARY,
         hover_color=COLOR_SECONDARY,
         border_width=3,
         border_color=COLOR_ACCENT,
         text_color=COLOR_TEXT_WHITE
      )
      self.run_button.pack(fill="x")
      
      # Atajos de teclado
      self.bind("<Return>", lambda e: self.on_execute_clicked())
      self.bind("<Escape>", lambda e: self.quit())
   
   def log(self, msg, level="INFO"):
      """Registra mensajes en el log con formato y color."""
      timestamp = datetime.datetime.now().strftime("%H:%M:%S")
      
      level_icons = {
         "INFO": "ℹ",
         "SUCCESS": "✓",
         "WARNING": "⚠",
         "ERROR": "✗",
         "PROCESS": "⚙"
      }
      
      icon = level_icons.get(level, "•")
      formatted_msg = f"[{timestamp}] {icon} {msg}\n"
      
      self.output_box.configure(state="normal")
      self.output_box.insert("end", formatted_msg)
      self.output_box.see("end")
      self.output_box.configure(state="disabled")
   
   def update_status(self, text, color=COLOR_TEXT_WHITE):
      """Actualiza el label de estado."""
      self.status_label.configure(text=f"Estado: {text}", text_color=color)
   
   def initial_check(self):
      """Verificación inicial del sistema."""
      self.log("Verificando prerequisitos del sistema...", "PROCESS")
      
      # Verificar conexión
      if check_internet():
         self.log("✓ Conexión a internet activa", "SUCCESS")
      else:
         self.log("⚠ Sin conexión a internet detectada", "WARNING")
      
      self.log("━" * 70, "INFO")
      self.log("✓ Sistema listo para inscripción AutoPilot", "SUCCESS")
   
   def on_execute_clicked(self):
      """Maneja el clic en el botón de ejecución."""
      if self.is_processing:
         self.log("⚠ Ya hay un proceso en ejecución", "WARNING")
         return
      
      # Confirmar acción
      response = messagebox.askyesno(
         "Confirmar Inscripción AutoPilot",
         "¿Desea iniciar la inscripción en Windows AutoPilot?\n\n"
         "Este proceso ejecutará automáticamente:\n"
         "• PowerShell con privilegios elevados\n"
         "• Instalación de Get-WindowsAutoPilotInfo\n"
         "• Registro del dispositivo en línea\n\n"
         "El proceso es completamente automático.\n\n"
         "¿Continuar?",
         icon='question'
      )
      
      if not response:
         self.log("✗ Operación cancelada por el usuario", "WARNING")
         return
      
      # Iniciar proceso
      self.is_processing = True
      self.run_button.configure(
         state="disabled",
         text="⏳ Procesando inscripción...",
         fg_color=COLOR_BG_MEDIUM
      )
      self.update_status("Ejecutando inscripción...", COLOR_WARNING)
      
      thread = threading.Thread(target=self.execute_autopilot, daemon=True)
      thread.start()
   
   def execute_autopilot(self):
      """Ejecuta el proceso completo de inscripción AutoPilot."""
      try:
         self.log("━" * 70, "INFO")
         self.log("🚀 INICIANDO INSCRIPCIÓN AUTOPILOT", "PROCESS")
         self.log("━" * 70, "INFO")
         
         self.animate_progress(0.1)
         
         # Crear script PowerShell automatizado
         self.log("Generando script PowerShell automatizado...", "PROCESS")
         
         ps_script = """
# Script Automatizado de Inscripción AutoPilot
# Copyright © 2025 Josué Romero - Stefanini / PQN

$ErrorActionPreference = 'Stop'

Write-Host ""
Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  INSCRIPCIÓN WINDOWS AUTOPILOT - PROCESO AUTOMATIZADO" -ForegroundColor White
Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""

# Paso 1: Configurar ExecutionPolicy Bypass
Write-Host "[1/4] Configurando política de ejecución..." -ForegroundColor Yellow
try {
   Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
   Write-Host "      ✓ Política de ejecución configurada" -ForegroundColor Green
} catch {
   Write-Host "      ✗ Error: $($_.Exception.Message)" -ForegroundColor Red
   exit 1
}

Write-Host ""

# Paso 2: Instalar NuGet Provider (requerido)
Write-Host "[2/4] Verificando NuGet Provider..." -ForegroundColor Yellow
try {
   if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
      Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser | Out-Null
      Write-Host "      ✓ NuGet Provider instalado" -ForegroundColor Green
   } else {
      Write-Host "      ✓ NuGet Provider ya está instalado" -ForegroundColor Green
   }
} catch {
   Write-Host "      ⚠ Continuando sin NuGet..." -ForegroundColor Yellow
}

Write-Host ""

# Paso 3: Instalar Get-WindowsAutoPilotInfo (con confirmación automática)
Write-Host "[3/4] Instalando Get-WindowsAutoPilotInfo..." -ForegroundColor Yellow
try {
   # Confiar en PSGallery automáticamente
   Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted -ErrorAction SilentlyContinue
   
   # Instalar el script con confirmación automática
   Install-Script -Name Get-WindowsAutoPilotInfo -Force -Scope CurrentUser -Confirm:$false -ErrorAction Stop
   Write-Host "      ✓ Get-WindowsAutoPilotInfo instalado correctamente" -ForegroundColor Green
} catch {
   Write-Host "      ✗ Error en instalación: $($_.Exception.Message)" -ForegroundColor Red
   Write-Host ""
   Write-Host "Presione cualquier tecla para cerrar..." -ForegroundColor Gray
   $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
   exit 1
}

Write-Host ""

# Paso 4: Ejecutar inscripción en línea
Write-Host "[4/4] Ejecutando inscripción en AutoPilot..." -ForegroundColor Yellow
Write-Host "      (Este proceso puede tomar varios minutos)" -ForegroundColor Gray
Write-Host ""

try {
   Get-WindowsAutoPilotInfo -Online
   Write-Host ""
   Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Green
   Write-Host "  ✓ INSCRIPCIÓN COMPLETADA EXITOSAMENTE" -ForegroundColor White
   Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Green
   Write-Host ""
   Write-Host "El dispositivo ha sido inscrito en Windows AutoPilot." -ForegroundColor Green
   Write-Host ""
} catch {
   Write-Host ""
   Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Red
   Write-Host "  ✗ ERROR EN LA INSCRIPCIÓN" -ForegroundColor White
   Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Red
   Write-Host ""
   Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
   Write-Host ""
   Write-Host "Presione cualquier tecla para cerrar..." -ForegroundColor Gray
   $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
   exit 1
}

Write-Host "Presione cualquier tecla para cerrar esta ventana..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
"""

         # Guardar script temporal
         temp_dir = Path(os.environ.get('TEMP', 'C:/Temp'))
         temp_dir.mkdir(parents=True, exist_ok=True)
         script_path = temp_dir / "autopilot_inscribe.ps1"
         
         with open(script_path, "w", encoding="utf-8") as f:
               f.write(ps_script)
         
         self.log(f"✓ Script creado: {script_path}", "SUCCESS")
         self.animate_progress(0.3)
         
         # Calcular hash del script
         script_hash = calculate_file_hash(script_path)
         self.log(f"✓ SHA-256: {script_hash[:32]}...", "INFO")
         self.animate_progress(0.4)
         
         # Ejecutar PowerShell con privilegios
         self.log("Iniciando PowerShell con privilegios elevados...", "PROCESS")
         self.log("⚠ Una ventana de PowerShell se abrirá automáticamente", "WARNING")
         
         self.animate_progress(0.6)
         
         # Ejecutar el script
         process = subprocess.Popen(
               [
                  "powershell.exe",
                  "-ExecutionPolicy", "Bypass",
                  "-NoProfile",
                  "-File", str(script_path)
               ],
               creationflags=subprocess.CREATE_NEW_CONSOLE
         )
         
         self.log("✓ PowerShell iniciado correctamente", "SUCCESS")
         self.log("✓ Todos los comandos se ejecutan automáticamente", "SUCCESS")
         self.animate_progress(0.9)
         
         time.sleep(1)
         self.animate_progress(1.0)
         
         self.log("━" * 70, "INFO")
         self.log("✓ PROCESO INICIADO EXITOSAMENTE", "SUCCESS")
         self.log("━" * 70, "INFO")
         self.log("📌 Continúe en la ventana de PowerShell", "INFO")
         self.log("📌 El proceso es completamente automático", "INFO")
         
         self.after(100, lambda: self.update_status(
               "Proceso completado - Verifique PowerShell",
               COLOR_SUCCESS
         ))
         
         # Mostrar mensaje de éxito
         self.after(200, lambda: messagebox.showinfo(
               "✓ Inscripción Iniciada",
               "El proceso de inscripción AutoPilot ha sido iniciado.\n\n"
               "✓ La ventana de PowerShell ejecutará todos los pasos automáticamente\n"
               "✓ NO se requiere intervención manual\n"
               "✓ Espere a que se complete el proceso\n\n"
               "Una vez finalizado, puede cerrar ambas ventanas.",
               icon='info'
         ))
         
      except Exception as e:
         self.log("━" * 70, "ERROR")
         self.log(f"✗ ERROR CRÍTICO: {str(e)}", "ERROR")
         self.log("━" * 70, "ERROR")
         self.update_status("Error crítico", COLOR_ERROR)
         
         self.after(100, lambda: messagebox.showerror(
               "Error en Inscripción",
               f"No se pudo completar la inscripción:\n\n{str(e)}\n\n"
               "Verifique que:\n"
               "• Tiene permisos de administrador\n"
               "• Hay conexión a internet\n"
               "• PowerShell está disponible"
         ))
      
      finally:
         # Restaurar interfaz
         self.is_processing = False
         self.after(100, lambda: self.run_button.configure(
               state="normal",
               text="🚀 Inscribir Equipo Ahora",
               fg_color=COLOR_PRIMARY
         ))
   
   def animate_progress(self, target_value):
      """Anima la barra de progreso."""
      current = self.progress_bar.get()
      step = 0.02
      
      while current < target_value:
         current = min(current + step, target_value)
         self.progress_bar.set(current)
         self.update_idletasks()
         time.sleep(0.03)


# ============================================================================
# PUNTO DE ENTRADA PRINCIPAL
# ============================================================================

def main():
   """Función principal con verificación de privilegios."""
   
   # Verificar privilegios de administrador
   if not is_admin():
      response = messagebox.askyesno(
         "Privilegios de Administrador Requeridos",
         "Esta aplicación requiere privilegios de administrador para funcionar correctamente.\n\n"
         "¿Desea reiniciar la aplicación como administrador?",
         icon='warning'
      )
      
      if response:
         run_as_admin()
      else:
         messagebox.showwarning(
               "Advertencia",
               "La aplicación puede no funcionar correctamente sin privilegios de administrador."
         )
   
   # Iniciar aplicación
   try:
      app = AutoPilotApp()
      app.mainloop()
   except Exception as e:
      messagebox.showerror(
         "Error Fatal",
         f"No se pudo iniciar la aplicación:\n\n{e}\n\n"
         f"Versión: {__version__}\n"
         f"Contacto: {__author__}"
      )
      sys.exit(1)


if __name__ == "__main__":
   main()