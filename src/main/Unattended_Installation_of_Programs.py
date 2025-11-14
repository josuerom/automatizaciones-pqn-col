"""
Unattended_Installation_of_Programs.py - Versión Profesional Mejorada
Autor: Josué Romero
Empresa: Stefanini / PQN
Fecha: 06/Noviembre/2025
Versión: 3.0 Professional Edition

Licencia: Propiedad de Stefanini / PQN - Todos los derechos reservados
Copyright © 2025 Josué Romero - Stefanini / PQN

Descripción:
Sistema automatizado para instalación desatendida de múltiples programas
sin intervención del usuario. Compatible con Windows 11.
"""

import os
import subprocess
import datetime
import time
import customtkinter as ctk
from tkinter import messagebox
from pathlib import Path
import threading
import hashlib
import sys
import ctypes

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

APP_TITLE = "Instalador Desatendido Multi-Programa"
APP_VERSION = f"v{__version__}"
APP_SIZE = "900x950"

# Paleta de colores profesional (Cian-Turquesa-Oscuro)
COLOR_PRIMARY = "#00838f"         # Cian profundo
COLOR_SECONDARY = "#00acc1"       # Cian medio
COLOR_ACCENT = "#26c6da"          # Cian brillante
COLOR_SUCCESS = "#00e676"         # Verde éxito
COLOR_WARNING = "#ffc107"         # Amarillo advertencia
COLOR_ERROR = "#d32f2f"           # Rojo error
COLOR_BG_DARK = "#0a0a0a"         # Negro profundo
COLOR_BG_MEDIUM = "#1a1a1a"       # Negro medio
COLOR_BG_LIGHT = "#2d2d2d"        # Gris oscuro
COLOR_TEXT_WHITE = "#ffffff"      # Texto blanco
COLOR_TEXT_GRAY = "#b0bec5"       # Texto gris claro

# Fuentes (Incrementadas +2px según requerimiento)
FONT_TITLE = ("Segoe UI", 28, "bold")          # Era 26
FONT_SUBTITLE = ("Segoe UI", 13)               # Era 11
FONT_CONSOLE = ("Consolas", 11)                # Era 9
FONT_BUTTON = ("Segoe UI", 14, "bold")         # Era 12
FONT_LABEL = ("Segoe UI", 13, "bold")          # Era 11
FONT_INFO = ("Segoe UI", 13)                   # Era 11

# Ruta de instaladores (fija)
INSTALLERS_PATH = Path("D:/Utilidades/Programas")
LOG_FILENAME = "install_log.txt"

# Definición de instaladores con banderas correctas
INSTALLERS = [
   {
      "id": "teamviewer_host",
      "name": "TeamViewer Host 15.71",
      "file": "1.1_TeamViewer_Host.exe",
      "args": "/S",  # Silencioso
      "timeout": 300,
      "category": "Soporte Remoto",
      "enabled": True,
      "description": "Cliente de soporte remoto (solo host)"
   },
   {
      "id": "teamviewer_full",
      "name": "TeamViewer Full Client 15.71",
      "file": "1.2_TeamViewer_Full_Client.exe",
      "args": "/S",  # Silencioso
      "timeout": 300,
      "category": "Soporte Remoto",
      "enabled": False,
      "description": "Cliente completo de soporte remoto"
   },
   {
      "id": "forticlient",
      "name": "FortiClient VPN 7.4.3",
      "file": "2_FortiClient.exe",
      "args": "/quiet /norestart",  # Silencioso sin reinicio
      "timeout": 600,
      "category": "Conectividad",
      "enabled": True,
      "description": "Cliente VPN corporativo"
   },
   {
      "id": "citrix",
      "name": "Citrix Workspace App 25.8",
      "file": "3_Citrix.exe",
      "args": "/silent /noreboot /AutoUpdateCheck=disabled",  # Silencioso
      "timeout": 600,
      "category": "Conectividad",
      "enabled": True,
      "description": "Acceso a aplicaciones virtualizadas"
   },
   {
      "id": "java8",
      "name": "Java 8 Update 341",
      "file": "4_Java8_341.exe",
      "args": "/s INSTALL_SILENT=1 AUTO_UPDATE=0 WEB_JAVA=1",  # Silencioso
      "timeout": 600,
      "category": "Runtime & Frameworks",
      "enabled": True,
      "description": "Java Runtime Environment 8"
   },
   {
      "id": "dotnet35",
      "name": ".NET Framework 3.5",
      "file": "5_NET_3.5.exe",
      "args": "/q /norestart",  # Silencioso sin reinicio
      "timeout": 900,
      "category": "Runtime & Frameworks",
      "enabled": True,
      "description": "Framework para aplicaciones .NET"
   },
   {
      "id": "adobe_reader",
      "name": "Adobe Acrobat Reader DC 2025",
      "file": "6_Reader.exe",
      "args": "/sAll /rs /msi EULA_ACCEPT=YES",  # Silencioso
      "timeout": 600,
      "category": "Esenciales",
      "enabled": True,
      "description": "Lector de documentos PDF"
   },
   {
      "id": "support_dell",
      "name": "SupportAssist Dell",
      "file": "7_SupportAssist_Dell.exe",
      "args": "/S /v/qn",  # Silencioso
      "timeout": 600,
      "category": "Soporte Hardware",
      "enabled": False,
      "description": "Soporte automático para equipos Dell"
   },
   {
      "id": "support_lenovo",
      "name": "SupportAssist Lenovo",
      "file": "7_SupportAssist_Lenovo.exe",
      "args": "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART",  # Silencioso
      "timeout": 600,
      "category": "Soporte Hardware",
      "enabled": False,
      "description": "Soporte automático para equipos Lenovo"
   },
   {
      "id": "teams",
      "name": "Microsoft Teams (Nuevo)",
      "file": "8_Teams.exe",
      "args": "/S",  # Silencioso
      "timeout": 900,
      "category": "Comunicaciones",
      "enabled": True,
      "description": "Plataforma de colaboración empresarial"
   },
   {
      "id": "chrome",
      "name": "Google Chrome Enterprise",
      "file": "9_ChromeEnterprise.msi",
      "args": "/qn /norestart",  # MSI silencioso
      "timeout": 600,
      "category": "Navegadores",
      "enabled": True,
      "description": "Navegador web corporativo"
   },
   {
      "id": "office365",
      "name": "Microsoft Office 365",
      "file": "10_Office365.exe",
      "config": "10_config.xml",
      "args": "/configure",  # Requiere XML
      "timeout": 1800,
      "category": "Productividad",
      "enabled": True,
      "description": "Suite ofimática completa"
   }
]


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
         f"No se pudo elevar privilegios:\n{e}"
      )
      sys.exit(1)


# ============================================================================
# FUNCIONES AUXILIARES
# ============================================================================

def calculate_hash(data):
   """Calcula hash SHA-256 de datos."""
   try:
      sha256_hash = hashlib.sha256()
      sha256_hash.update(str(data).encode('utf-8'))
      return sha256_hash.hexdigest()
   except:
      return "N/A"


def find_installer(filename):
   """Busca un instalador en la ruta fija."""
   installer_path = INSTALLERS_PATH / filename
   return str(installer_path) if installer_path.exists() else None


def run_installer(installer_path, arguments, timeout=600):
   """
   Ejecuta un instalador de forma desatendida.
   
   Args:
      installer_path: Ruta completa del instalador
      arguments: Argumentos de línea de comandos
      timeout: Timeout en segundos
   
   Returns:
      tuple: (success: bool, error_msg: str)
   """
   try:
      # Construir comando
      if installer_path.lower().endswith('.msi'):
         # Para instaladores MSI
         cmd = f'msiexec.exe /i "{installer_path}" {arguments}'
      else:
         # Para instaladores EXE
         cmd = f'"{installer_path}" {arguments}'
      
      # Ejecutar
      result = subprocess.run(
         cmd,
         shell=True,
         capture_output=True,
         text=True,
         timeout=timeout,
         creationflags=subprocess.CREATE_NO_WINDOW
      )
      
      # Códigos de retorno comunes de éxito
      success_codes = [0, 3010]  # 0=éxito, 3010=éxito pero requiere reinicio
      
      if result.returncode in success_codes:
         return True, ""
      else:
         return False, f"Código de salida: {result.returncode}"
         
   except subprocess.TimeoutExpired:
      return False, f"Timeout - La instalación excedió {timeout} segundos"
   except Exception as e:
      return False, str(e)


def write_log(path, content):
   """Escribe en el archivo de log."""
   try:
      with open(path, "a", encoding="utf-8") as f:
         timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
         f.write(f"[{timestamp}] {content}\n")
   except:
      pass


# ============================================================================
# CLASE PRINCIPAL DE LA APLICACIÓN
# ============================================================================

class InstallerApp(ctk.CTk):
   def __init__(self):
      super().__init__()
      
      # Configuración de ventana
      self.title(f"{APP_TITLE} {APP_VERSION}")
      self.geometry(APP_SIZE)
      self.resizable(False, False)
      
      # Variables de estado
      self.is_processing = False
      self.should_cancel = False
      self.installer_vars = {}
      self.log_path = None
      
      # Estadísticas
      self.stats = {
         "total": 0,
         "installed": 0,
         "skipped": 0,
         "failed": 0
      }
      
      # Construir interfaz
      self.build_ui()
      
      # Inicializar
      self.after(500, self.initialize)
   
   def build_ui(self):
      """Construye la interfaz de usuario moderna y profesional."""
      
      # Marco principal
      main_frame = ctk.CTkFrame(self, fg_color=COLOR_BG_DARK)
      main_frame.pack(fill="both", expand=True, padx=10, pady=10)  # Reducido padding general
      
      # === ENCABEZADO MODERNO ===
      header_frame = ctk.CTkFrame(
         main_frame,
         fg_color=COLOR_PRIMARY,
         corner_radius=12,
         border_width=2,
         border_color=COLOR_ACCENT
      )
      header_frame.pack(fill="x", pady=(0, 10))  # Reducido el padding inferior
      
      title_label = ctk.CTkLabel(
         header_frame,
         text="📦 " + APP_TITLE,
         font=FONT_TITLE,
         text_color=COLOR_TEXT_WHITE
      )
      title_label.pack(pady=(10, 3))  # Reducido el espacio entre el título y el marco
      
      subtitle_label = ctk.CTkLabel(
         header_frame,
         text=f"{__company__} | {__author__} | {APP_VERSION}",
         font=FONT_SUBTITLE,
         text_color=COLOR_ACCENT
      )
      subtitle_label.pack(pady=(0, 5))  # Reducido padding
      
      copyright_label = ctk.CTkLabel(
         header_frame,
         text=__copyright__,
         font=("Segoe UI", 11),
         text_color=COLOR_TEXT_GRAY
      )
      copyright_label.pack(pady=(0, 10))  # Reducido padding
      
      # === SELECCIÓN DE PROGRAMAS ===
      programs_frame = ctk.CTkFrame(
         main_frame,
         fg_color=COLOR_BG_MEDIUM,
         corner_radius=10
      )
      programs_frame.pack(fill="both", expand=True, pady=(0, 10))  # Reducido el padding inferior
      
      programs_title = ctk.CTkLabel(
         programs_frame,
         text="✅ Seleccione los Programas a Instalar",
         font=FONT_LABEL,
         text_color=COLOR_TEXT_WHITE
      )
      programs_title.pack(pady=(10, 5), anchor="w", padx=15)  # Reducido padding en el título
      
      # Scrollable frame para los programas
      programs_scroll = ctk.CTkScrollableFrame(
         programs_frame,
         width=820,
         height=300,
         fg_color=COLOR_BG_LIGHT,
         corner_radius=8
      )
      programs_scroll.pack(padx=15, pady=(0, 10))  # Reducido padding
      
      # Crear checkboxes agrupados por categoría
      categories = {}
      for installer in INSTALLERS:
         cat = installer["category"]
         if cat not in categories:
               categories[cat] = []
         categories[cat].append(installer)
      
      # Renderizar por categorías
      for category, items in sorted(categories.items()):
         # Título de categoría
         cat_label = ctk.CTkLabel(
               programs_scroll,
               text=f"▼ {category}",
               font=("Segoe UI", 13, "bold"),
               text_color=COLOR_ACCENT,
               anchor="w"
         )
         cat_label.pack(anchor="w", pady=(10, 5), padx=10)  # Reducido padding
         
         # Items de la categoría
         for installer in items:
               item_frame = ctk.CTkFrame(
                  programs_scroll,
                  fg_color=COLOR_BG_MEDIUM,
                  corner_radius=6,
                  border_width=1,
                  border_color=COLOR_SECONDARY
               )
               item_frame.pack(fill="x", pady=2, padx=10)  # Reducido padding entre los checkboxes
               
               # Variable del checkbox
               var = ctk.BooleanVar(value=installer["enabled"])
               self.installer_vars[installer["id"]] = var
               
               # Checkbox
               checkbox = ctk.CTkCheckBox(
                  item_frame,
                  text=installer["name"],
                  variable=var,
                  font=FONT_INFO,
                  text_color=COLOR_TEXT_WHITE,
                  fg_color=COLOR_PRIMARY,
                  hover_color=COLOR_SECONDARY
               )
               checkbox.pack(side="left", padx=10, pady=8)  # Reducido padding en checkbox
               
               # Descripción
               desc_label = ctk.CTkLabel(
                  item_frame,
                  text=f"• {installer['description']}",
                  font=("Segoe UI", 10),
                  text_color=COLOR_TEXT_GRAY
               )
               desc_label.pack(side="left", padx=(0, 10))  # Reducido el padding de la descripción
      
      # Botones de selección rápida
      quick_frame = ctk.CTkFrame(programs_frame, fg_color="transparent")
      quick_frame.pack(pady=(0, 10), padx=15)  # Reducido padding inferior
      
      ctk.CTkButton(
         quick_frame,
         text="✓ Seleccionar Todos",
         command=self.select_all,
         width=180,
         height=35,
         font=("Segoe UI", 12),
         fg_color=COLOR_PRIMARY,
         hover_color=COLOR_SECONDARY
      ).pack(side="left", padx=5)
      
      ctk.CTkButton(
         quick_frame,
         text="✗ Deseleccionar Todos",
         command=self.deselect_all,
         width=180,
         height=35,
         font=("Segoe UI", 12),
         fg_color=COLOR_BG_LIGHT,
         hover_color=COLOR_BG_MEDIUM
      ).pack(side="left", padx=5)
      
      ctk.CTkButton(
         quick_frame,
         text="⚡ Solo Esenciales",
         command=self.select_essentials,
         width=180,
         height=35,
         font=("Segoe UI", 12),
         fg_color=COLOR_SUCCESS,
         hover_color="#00c853"
      ).pack(side="left", padx=5)
      
      # === PROGRESO ===
      progress_frame = ctk.CTkFrame(
         main_frame,
         fg_color=COLOR_BG_MEDIUM,
         corner_radius=10
      )
      progress_frame.pack(fill="x", pady=(0, 10))  # Reducido padding inferior
      
      self.progress_label = ctk.CTkLabel(
         progress_frame,
         text="Listo para iniciar instalación",
         font=FONT_INFO,
         text_color=COLOR_TEXT_WHITE
      )
      self.progress_label.pack(pady=(10, 5))  # Reducido el espacio
      
      self.progress_bar = ctk.CTkProgressBar(
         progress_frame,
         width=820,
         height=12,
         corner_radius=6,
         progress_color=COLOR_ACCENT,
         fg_color=COLOR_BG_LIGHT
      )
      self.progress_bar.pack(pady=(0, 10), padx=15)  # Reducido padding
      self.progress_bar.set(0)
      
      # === ÁREA DE LOGS ===
      log_label = ctk.CTkLabel(
         main_frame,
         text="📝 Registro de Instalación",
         font=FONT_LABEL,
         text_color=COLOR_TEXT_WHITE,
         anchor="w"
      )
      log_label.pack(pady=(5, 5), anchor="w")  # Reducido padding
      
      self.textbox = ctk.CTkTextbox(
         main_frame,
         width=840,
         height=180,
         font=FONT_CONSOLE,
         fg_color=COLOR_BG_LIGHT,
         text_color=COLOR_TEXT_WHITE,
         border_width=2,
         border_color=COLOR_SECONDARY,
         corner_radius=10
      )
      self.textbox.pack(pady=(0, 10))  # Reducido padding inferior
      self.log("✓ Sistema inicializado correctamente", "SUCCESS")
      self.log(f"✓ Licencia: {__license__}", "INFO")
      
      # === BOTONES PRINCIPALES ===
      button_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
      button_frame.pack(fill="x", pady=(10, 5))  # Reducido el padding
      
      self.button_start = ctk.CTkButton(
         button_frame,
         text="🚀 Iniciar Instalación Desatendida",
         command=self.start_installation,
         font=FONT_BUTTON,
         height=50,
         corner_radius=10,
         fg_color=COLOR_PRIMARY,
         hover_color=COLOR_SECONDARY,
         border_width=3,
         border_color=COLOR_ACCENT,
         text_color=COLOR_TEXT_WHITE
      )
      self.button_start.pack(side="left", expand=True, fill="x", padx=(0, 5))  # Reducido padding en los botones
      
      self.button_cancel = ctk.CTkButton(
         button_frame,
         text="⏹ Cancelar Proceso",
         command=self.cancel_installation,
         font=FONT_BUTTON,
         height=50,
         corner_radius=10,
         fg_color=COLOR_ERROR,
         hover_color="#c62828",
         border_width=2,
         border_color=COLOR_ERROR,
         text_color=COLOR_TEXT_WHITE,
         state="disabled"
      )
      self.button_cancel.pack(side="left", fill="x", padx=(5, 0))  # Reducido padding en los botones
      
      # Atajos de teclado
      self.bind("<Escape>", lambda e: self.quit())

   
   def log(self, message, level="INFO"):
      """Registra mensajes en el log con formato."""
      timestamp = datetime.datetime.now().strftime("%H:%M:%S")
      icons = {
         "INFO": "ℹ",
         "SUCCESS": "✓",
         "WARNING": "⚠",
         "ERROR": "✗",
         "SKIP": "⏭",
         "PROCESS": "⚙"
      }
      icon = icons.get(level, "•")
      
      formatted_msg = f"[{timestamp}] {icon} {message}\n"
      
      self.textbox.configure(state="normal")
      self.textbox.insert("end", formatted_msg)
      self.textbox.see("end")
      self.textbox.configure(state="disabled")
      
      # Escribir en archivo de log
      if self.log_path:
         write_log(self.log_path, f"[{level}] {message}")
   
   def select_all(self):
      """Selecciona todos los programas."""
      for var in self.installer_vars.values():
         var.set(True)
      self.log("✓ Todos los programas seleccionados", "INFO")
   
   def deselect_all(self):
      """Deselecciona todos los programas."""
      for var in self.installer_vars.values():
         var.set(False)
      self.log("✗ Todos los programas deseleccionados", "INFO")
   
   def select_essentials(self):
      """Selecciona solo programas esenciales."""
      essentials = ["adobe_reader", "forticlient", "citrix", "java8", "teams", "chrome", "office365"]
      for inst_id, var in self.installer_vars.items():
         var.set(inst_id in essentials)
      self.log("⚡ Programas esenciales seleccionados", "INFO")
   
   def initialize(self):
      """Inicializa la aplicación y verifica prerequisitos."""
      self.log("Verificando prerequisitos del sistema...", "PROCESS")
      
      # Verificar que existe la carpeta de instaladores
      if not INSTALLERS_PATH.exists():
         self.log(f"✗ No se encontró la carpeta: {INSTALLERS_PATH}", "ERROR")
         messagebox.showerror(
               "Error de Configuración",
               f"No se encontró la carpeta de instaladores:\n\n"
               f"{INSTALLERS_PATH}\n\n"
               f"Asegúrese de que la carpeta existe y contiene los instaladores."
         )
         self.button_start.configure(state="disabled")
         return
      
      self.log(f"✓ Carpeta de instaladores encontrada: {INSTALLERS_PATH}", "SUCCESS")
      self.log_path = INSTALLERS_PATH / LOG_FILENAME
      
      # Verificar cuántos instaladores están disponibles
      available_count = 0
      for installer in INSTALLERS:
         if find_installer(installer["file"]):
               available_count += 1
      
      self.log(f"✓ Instaladores disponibles: {available_count}/{len(INSTALLERS)}", "SUCCESS")
      self.log("━" * 85, "INFO")
      self.log("✓ Sistema listo para instalar programas", "SUCCESS")
   
   def cancel_installation(self):
      """Cancela el proceso de instalación."""
      if self.is_processing:
         response = messagebox.askyesno(
               "Cancelar Instalación",
               "¿Está seguro de que desea cancelar el proceso?\n\n"
               "La instalación actual se completará antes de cancelar."
         )
         if response:
               self.should_cancel = True
               self.log("⚠ Cancelación solicitada por el usuario", "WARNING")
               self.button_cancel.configure(state="disabled")
   
   def start_installation(self):
      """Inicia el proceso de instalación desatendida."""
      # Verificar selección
      selected = [
         inst_id for inst_id, var in self.installer_vars.items()
         if var.get()
      ]
      
      if not selected:
         messagebox.showwarning(
               "Sin Selección",
               "Debe seleccionar al menos un programa para instalar."
         )
         return
      
      # Mostrar resumen y confirmar
      selected_names = [
         inst["name"] for inst in INSTALLERS
         if inst["id"] in selected
      ]
      
      confirmation_msg = f"Se instalarán {len(selected)} programas de forma desatendida:\n\n"
      for name in selected_names[:5]:
         confirmation_msg += f"• {name}\n"
      if len(selected_names) > 5:
         confirmation_msg += f"• ... y {len(selected_names) - 5} más\n"
      
      confirmation_msg += f"\n⚠️ IMPORTANTE:\n"
      confirmation_msg += f"• Este proceso puede tardar varios minutos\n"
      confirmation_msg += f"• No apague ni reinicie el equipo durante la instalación\n"
      confirmation_msg += f"• Las instalaciones se ejecutarán en modo silencioso\n\n"
      confirmation_msg += f"¿Desea continuar?"
      
      response = messagebox.askyesno(
         "Confirmar Instalación",
         confirmation_msg,
         icon='question'
      )
      
      if not response:
         self.log("✗ Instalación cancelada por el usuario", "WARNING")
         return
      
      # Iniciar proceso
      self.is_processing = True
      self.should_cancel = False
      self.button_start.configure(state="disabled")
      self.button_cancel.configure(state="normal")
      
      # Limpiar log
      self.textbox.configure(state="normal")
      self.textbox.delete("1.0", "end")
      self.textbox.configure(state="disabled")
      
      # Resetear estadísticas
      self.stats = {"total": 0, "installed": 0, "skipped": 0, "failed": 0}
      
      # Iniciar hilo de instalación
      thread = threading.Thread(target=self.install_programs, daemon=True)
      thread.start()
   
   def install_programs(self):
      """Ejecuta el proceso de instalación de los programas seleccionados."""
      try:
         self.log("━" * 85, "INFO")
         self.log("🚀 INICIANDO INSTALACIÓN DESATENDIDA DE PROGRAMAS", "PROCESS")
         self.log("━" * 85, "INFO")
         
         # Obtener lista de instaladores seleccionados
         selected_installers = [
               inst for inst in INSTALLERS
               if self.installer_vars[inst["id"]].get()
         ]
         
         total = len(selected_installers)
         self.stats["total"] = total
         completed = 0
         
         self.log(f"Total de programas a instalar: {total}", "INFO")
         self.log("", "INFO")
         
         # Instalar cada programa
         for installer in selected_installers:
               if self.should_cancel:
                  self.log("", "INFO")
                  self.log("✗ Proceso cancelado por el usuario", "WARNING")
                  break
               
               self.progress_label.configure(
                  text=f"Instalando ({completed + 1}/{total}): {installer['name']}"
               )
               
               self.log(f"──────────────────────────────────────────────────────────────", "INFO")
               self.log(f"[{completed + 1}/{total}] {installer['name']}", "PROCESS")
               self.log(f"Categoría: {installer['category']}", "INFO")
               
               # Buscar instalador
               installer_path = find_installer(installer["file"])
               
               if not installer_path:
                  self.log(f"      ✗ Archivo no encontrado: {installer['file']}", "ERROR")
                  self.stats["failed"] += 1
               else:
                  # Preparar argumentos
                  args = installer["args"]
                  
                  # Para Office 365, agregar ruta del config.xml
                  if installer["id"] == "office365" and "config" in installer:
                     config_path = find_installer(installer["config"])
                     if config_path:
                           args = f'{args} "{config_path}"'
                     else:
                        self.log("      ✗ No se encontró archivo de configuración XML", "ERROR")
                        self.stats["failed"] += 1
                        completed += 1
                        continue

               # Ejecutar instalador
               self.log("      ⚙ Ejecutando instalador en modo silencioso...", "PROCESS")
               success, error_msg = run_installer(installer_path, args, installer["timeout"])

               if success:
                  self.log("      ✓ Instalación completada con éxito", "SUCCESS")
                  self.stats["installed"] += 1
               else:
                  self.log(f"      ✗ Falló la instalación → {error_msg}", "ERROR")
                  self.stats["failed"] += 1

               completed += 1
               self.progress_bar.set(completed / total)
               time.sleep(0.4)

         # Finalización
         self.finish_installation()

      except Exception as e:
         self.log(f"✗ Error inesperado: {e}", "ERROR")
         self.finish_installation(force_error=True)


# ============================================================================
# FINALIZACIÓN DEL PROCESO
# ============================================================================
def finish_installation(self, force_error=False):
   """Finaliza el proceso, limpia estados y muestra un resumen."""
   self.is_processing = False
   self.button_start.configure(state="normal")
   self.button_cancel.configure(state="disabled")
   self.progress_label.configure(text="Proceso finalizado")

   self.log("")
   self.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━", "INFO")
   self.log("🏁 PROCESO COMPLETADO", "SUCCESS")
   self.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━", "INFO")

   # Mostrar estadísticas
   t = self.stats
   self.log(f"📌 Total seleccionados : {t['total']}", "INFO")
   self.log(f"📦 Instalados correctamente : {t['installed']}", "SUCCESS")
   self.log(f"⏭ Saltados / No encontrados : {t['skipped']}", "WARNING")
   self.log(f"❌ Fallidos : {t['failed']}", "ERROR")

   if force_error:
      self.log("⚠ El proceso terminó con errores inesperados.", "ERROR")

   # Mensaje final
   messagebox.showinfo(
      "Instalación Finalizada",
      f"Proceso completado.\n\n"
      f"Programas instalados: {t['installed']}\n"
      f"Fallidos: {t['failed']}\n"
      f"Saltados: {t['skipped']}\n\n"
      f"Puede revisar el log completo para más detalles."
   )


# ============================================================================
# PUNTO DE ENTRADA (MAIN)
# ============================================================================
if __name__ == "__main__":
   # Verificar si tiene permisos administrativos
   if not is_admin():
      messagebox.showwarning(
         "Permisos Requeridos",
         "Este instalador requiere privilegios administrativos.\n"
         "Se reiniciará automáticamente con permisos elevados."
      )
      run_as_admin()

   app = InstallerApp()
   app.mainloop()
