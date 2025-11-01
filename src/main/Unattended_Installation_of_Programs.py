"""
Unattended_Installation_of_Programs.py - Versión Mejorada
Autor: Josué Romero
Empresa: Stefanini / PQN
Fecha: 31/Octubre/2025

Descripción:
Instala múltiples programas de forma desatendida sin intervención del usuario.
"""

import os
import subprocess
import datetime
import time
import customtkinter as ctk
from tkinter import messagebox
from pathlib import Path
import threading

# ============================================================================
# CONFIGURACIÓN GLOBAL
# ============================================================================
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

APP_TITLE = "Instalador Desatendido PQN-COL"
APP_VERSION = "v2.0"
APP_SIZE = "850x890"

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
FONT_CONSOLE = ("Consolas", 9)
FONT_BUTTON = ("Segoe UI", 12, "bold")
FONT_LABEL = ("Segoe UI", 11, "bold")

# Rutas de búsqueda (prioridad: D: -> C:)
SEARCH_PATHS = [
   Path("D:/Utilidades/Programas"),
   Path("C:/Utilidades/Programas")
]

LOG_FILENAME = "install_log.txt"

# Definición de instaladores
INSTALLERS = [
   {
      "id": "reader",
      "name": "Adobe Acrobat Reader PDF",
      "file": "1_Reader.exe",
      "args": "/sAll /rs /rps /msi EULA_ACCEPT=YES",
      "check_cmd": 'reg query "HKLM\\Software\\Adobe\\Acrobat Reader"',
      "category": "Esenciales",
      "enabled": True
   },
   {
      "id": "forti",
      "name": "FortiClient VPN",
      "file": "2_FortiClient.exe",
      "args": "/quiet /norestart",
      "check_cmd": 'reg query "HKLM\\Software\\Fortinet\\FortiClient"',
      "category": "Conectividad",
      "enabled": True
   },
   {
      "id": "citrix",
      "name": "Citrix Workspace App",
      "file": "3_Citrix.exe",
      "args": "/silent /noreboot",
      "check_cmd": 'reg query "HKLM\\Software\\Citrix"',
      "category": "Conectividad",
      "enabled": True
   },
   {
      "id": "java",
      "name": "Java 8 Update 341",
      "file": "4_Java341.exe",
      "args": "/s",
      "check_cmd": 'reg query "HKLM\\Software\\JavaSoft\\Java Runtime Environment\\1.8"',
      "category": "Runtime",
      "enabled": True
   },
   {
      "id": "net35",
      "name": ".NET Framework 3.5",
      "file": "5_NET35.exe",
      "args": "/quiet /norestart",
      "check_cmd": 'reg query "HKLM\\Software\\Microsoft\\NET Framework Setup\\NDP\\v3.5"',
      "category": "Runtime",
      "enabled": True
   },
   {
      "id": "teamviewer",
      "name": "TeamViewer Host 2025",
      "file": "6_TeamViewerHost.exe",
      "args": "/S",
      "check_cmd": 'reg query "HKLM\\Software\\TeamViewer"',
      "category": "Soporte",
      "enabled": True
   },
   {
      "id": "support_dell",
      "name": "SupportAssist Dell",
      "file": "7_SupportAssistDell.exe",
      "args": "/quiet",
      "check_cmd": 'reg query "HKLM\\Software\\Dell\\SupportAssist"',
      "category": "Soporte",
      "enabled": False
   },
   {
      "id": "support_lenovo",
      "name": "SupportAssist Lenovo",
      "file": "7_SupportAssistLenovo.exe",
      "args": "/quiet",
      "check_cmd": 'reg query "HKLM\\Software\\Lenovo"',
      "category": "Soporte",
      "enabled": False
   },
   {
      "id": "teams",
      "name": "Microsoft Teams",
      "file": "8_Teams.exe",
      "args": "",
      "check_cmd": 'reg query "HKLM\\Software\\Microsoft\\Teams"',
      "category": "Comunicaciones",
      "enabled": True
   }
]

OFFICE_INSTALLER = {
   "id": "office",
   "name": "Microsoft Office 365",
   "file": "9_Office365.exe",
   "config": "9_config.xml",
   "check_cmd": 'reg query "HKLM\\Software\\Microsoft\\Office"',
   "category": "Productividad",
   "enabled": True
}


# ============================================================================
# FUNCIONES AUXILIARES
# ============================================================================

def find_base_folder():
   """Encuentra la carpeta base de instaladores."""
   for path in SEARCH_PATHS:
      if path.exists():
         return path
   return None


def find_installer(filename):
   """Busca un instalador específico."""
   base = find_base_folder()
   if base:
      installer_path = base / filename
      if installer_path.exists():
         return str(installer_path)
   return None


def check_if_installed(check_cmd):
   """Verifica si un programa ya está instalado."""
   try:
      result = subprocess.run(
         check_cmd,
         shell=True,
         capture_output=True,
         timeout=10
      )
      return result.returncode == 0
   except:
      return False


def run_installer(installer_path, arguments, timeout=600):
   """
   Ejecuta un instalador.
   
   Returns:
      tuple: (success: bool, error_msg: str)
   """
   try:
      cmd = [installer_path] + (arguments.split() if arguments else [])
      result = subprocess.run(
         cmd,
         capture_output=True,
         text=True,
         timeout=timeout
      )
      
      if result.returncode == 0:
         return True, ""
      else:
         return False, f"Código de salida: {result.returncode}"
   except subprocess.TimeoutExpired:
      return False, "Timeout (instalación tardó demasiado)"
   except Exception as e:
      return False, str(e)


def write_log(path, content):
   """Escribe en el archivo de log."""
   try:
      with open(path, "a", encoding="utf-8") as f:
         f.write(content + "\n")
   except:
      pass


# ============================================================================
# CLASE PRINCIPAL
# ============================================================================

class InstallerApp(ctk.CTk):
   def __init__(self):
      super().__init__()
      
      # Configuración
      self.title(f"{APP_TITLE} {APP_VERSION}")
      self.geometry(APP_SIZE)
      self.resizable(False, False)
      
      # Variables
      self.is_processing = False
      self.should_cancel = False
      self.installer_vars = {}
      self.office_var = None
      self.base_folder = None
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
      self.after(300, self.initialize)
   
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
         text="📦 " + APP_TITLE,
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
      
      # === SELECCIÓN DE PROGRAMAS ===
      programs_frame = ctk.CTkFrame(main_frame, fg_color=COLOR_BG_LIGHT, corner_radius=8)
      programs_frame.pack(fill="both", expand=True, pady=(0, 10))
      
      programs_title = ctk.CTkLabel(
         programs_frame,
         text="✅ Seleccione los Programas a Instalar",
         font=FONT_LABEL,
         text_color=COLOR_TEXT
      )
      programs_title.pack(pady=(10, 10), anchor="w", padx=15)
      
      # Scrollable frame
      programs_scroll = ctk.CTkScrollableFrame(
         programs_frame,
         width=780,
         height=280,
         fg_color="transparent"
      )
      programs_scroll.pack(padx=15, pady=(0, 10))
      
      # Crear checkboxes agrupados por categoría
      categories = {}
      for installer in INSTALLERS:
         cat = installer["category"]
         if cat not in categories:
               categories[cat] = []
         categories[cat].append(installer)
      
      for category, items in categories.items():
         # Título de categoría
         cat_label = ctk.CTkLabel(
               programs_scroll,
               text=f"  {category}",
               font=("Segoe UI", 11, "bold"),
               text_color=COLOR_PRIMARY,
               anchor="w"
         )
         cat_label.pack(anchor="w", pady=(10, 5))
         
         # Items de la categoría
         for installer in items:
               item_frame = ctk.CTkFrame(programs_scroll, fg_color="#242424", corner_radius=6)
               item_frame.pack(fill="x", pady=2, padx=5)
               
               var = ctk.BooleanVar(value=installer["enabled"])
               self.installer_vars[installer["id"]] = var
               
               checkbox = ctk.CTkCheckBox(
                  item_frame,
                  text=installer["name"],
                  variable=var,
                  font=("Segoe UI", 10)
               )
               checkbox.pack(side="left", padx=10, pady=8)
      
      # Office (separado)
      cat_label = ctk.CTkLabel(
         programs_scroll,
         text=f"  {OFFICE_INSTALLER['category']}",
         font=("Segoe UI", 11, "bold"),
         text_color=COLOR_PRIMARY,
         anchor="w"
      )
      cat_label.pack(anchor="w", pady=(10, 5))
      
      office_frame = ctk.CTkFrame(programs_scroll, fg_color="#242424", corner_radius=6)
      office_frame.pack(fill="x", pady=2, padx=5)
      
      self.office_var = ctk.BooleanVar(value=OFFICE_INSTALLER["enabled"])
      
      office_checkbox = ctk.CTkCheckBox(
         office_frame,
         text=OFFICE_INSTALLER["name"],
         variable=self.office_var,
         font=("Segoe UI", 10)
      )
      office_checkbox.pack(side="left", padx=10, pady=8)
      
      # Botones de selección rápida
      quick_frame = ctk.CTkFrame(programs_frame, fg_color="transparent")
      quick_frame.pack(pady=(0, 10), padx=15)
      
      ctk.CTkButton(
         quick_frame,
         text="Seleccionar Todos",
         command=self.select_all,
         width=150,
         height=30,
         font=("Segoe UI", 10)
      ).pack(side="left", padx=5)
      
      ctk.CTkButton(
         quick_frame,
         text="Deseleccionar Todos",
         command=self.deselect_all,
         width=150,
         height=30,
         font=("Segoe UI", 10)
      ).pack(side="left", padx=5)
      
      ctk.CTkButton(
         quick_frame,
         text="Solo Esenciales",
         command=self.select_essentials,
         width=150,
         height=30,
         font=("Segoe UI", 10)
      ).pack(side="left", padx=5)
      
      # === PROGRESO ===
      progress_frame = ctk.CTkFrame(main_frame, fg_color=COLOR_BG_LIGHT, corner_radius=8)
      progress_frame.pack(fill="x", pady=(0, 10))
      
      self.progress_label = ctk.CTkLabel(
         progress_frame,
         text="Listo para iniciar",
         font=FONT_SUBTITLE,
         text_color=COLOR_TEXT
      )
      self.progress_label.pack(pady=(10, 5))
      
      self.progress_bar = ctk.CTkProgressBar(
         progress_frame,
         width=780,
         height=10,
         corner_radius=5,
         progress_color=COLOR_PRIMARY
      )
      self.progress_bar.pack(pady=(0, 10), padx=15)
      self.progress_bar.set(0)
      
      # === LOG ===
      log_label = ctk.CTkLabel(
         main_frame,
         text="📝 Registro de Instalación",
         font=FONT_LABEL,
         text_color=COLOR_TEXT,
         anchor="w"
      )
      log_label.pack(pady=(5, 5), anchor="w")
      
      self.textbox = ctk.CTkTextbox(
         main_frame,
         width=800,
         height=150,
         font=FONT_CONSOLE,
         fg_color=COLOR_BG_LIGHT,
         text_color=COLOR_TEXT,
         border_width=2,
         border_color=COLOR_PRIMARY,
         corner_radius=8
      )
      self.textbox.pack(pady=(0, 10))
      
      # === BOTONES ===
      button_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
      button_frame.pack(fill="x")
      
      self.button_start = ctk.CTkButton(
         button_frame,
         text="▶ Iniciar Instalación",
         command=self.start_installation,
         font=FONT_BUTTON,
         height=40,
         corner_radius=8,
         fg_color=COLOR_PRIMARY,
         hover_color="#1976d2"
      )
      self.button_start.pack(side="left", expand=True, fill="x", padx=(0, 5))
      
      self.button_cancel = ctk.CTkButton(
         button_frame,
         text="⏹ Cancelar",
         command=self.cancel_installation,
         font=FONT_BUTTON,
         height=40,
         corner_radius=8,
         fg_color=COLOR_ERROR,
         hover_color="#c62828",
         state="disabled"
      )
      self.button_cancel.pack(side="left", expand=True, fill="x", padx=(5, 0))
   
   def log(self, message, level="INFO"):
      """Registra mensaje en el log."""
      icons = {
         "INFO": "ℹ",
         "SUCCESS": "✓",
         "WARNING": "⚠",
         "ERROR": "✗",
         "SKIP": "⏭"
      }
      icon = icons.get(level, "•")
      timestamp = datetime.datetime.now().strftime("%H:%M:%S")
      
      self.textbox.configure(state="normal")
      self.textbox.insert("end", f"[{timestamp}] {icon} {message}\n")
      self.textbox.see("end")
      self.textbox.configure(state="disabled")
      
      # Escribir en archivo
      if self.log_path:
         write_log(self.log_path, f"[{timestamp}] [{level}] {message}")
   
   def select_all(self):
      """Selecciona todos los programas."""
      for var in self.installer_vars.values():
         var.set(True)
      self.office_var.set(True)
      self.log("Todos los programas seleccionados")
   
   def deselect_all(self):
      """Deselecciona todos los programas."""
      for var in self.installer_vars.values():
         var.set(False)
      self.office_var.set(False)
      self.log("Todos los programas deseleccionados")
   
   def select_essentials(self):
      """Selecciona solo programas esenciales."""
      essentials = ["reader", "forti", "citrix", "java", "teams"]
      for inst_id, var in self.installer_vars.items():
         var.set(inst_id in essentials)
      self.office_var.set(True)
      self.log("Solo programas esenciales seleccionados")
   
   def initialize(self):
      """Inicializa la aplicación."""
      self.log("Inicializando sistema de instalación...")
      
      # Buscar carpeta base
      self.base_folder = find_base_folder()
      
      if not self.base_folder:
         self.log("✗ No se encontró la carpeta de instaladores", "ERROR")
         self.log("  Rutas buscadas:", "ERROR")
         for path in SEARCH_PATHS:
               self.log(f"    - {path}", "ERROR")
         messagebox.showerror(
               "Error",
               "No se encontró la carpeta de instaladores.\n\n"
               "Asegúrese de que existe:\n"
               "• D:\\Utilidades\\Programas\n"
               "• C:\\Utilidades\\Programas"
         )
         self.button_start.configure(state="disabled")
         return
      
      self.log(f"✓ Carpeta de instaladores encontrada: {self.base_folder}", "SUCCESS")
      self.log_path = self.base_folder / LOG_FILENAME
      self.log("✓ Sistema listo para instalar", "SUCCESS")
   
   def cancel_installation(self):
      """Cancela la instalación."""
      if self.is_processing:
         response = messagebox.askyesno(
               "Cancelar",
               "¿Está seguro de que desea cancelar?\n\n"
               "La instalación actual se completará."
         )
         if response:
               self.should_cancel = True
               self.log("Cancelación solicitada...", "WARNING")
               self.button_cancel.configure(state="disabled")
   
   def start_installation(self):
      """Inicia el proceso de instalación."""
      # Verificar selección
      selected = [
         inst_id for inst_id, var in self.installer_vars.items()
         if var.get()
      ]
      
      if self.office_var.get():
         selected.append("office")
      
      if not selected:
         messagebox.showwarning(
               "Sin Selección",
               "Debe seleccionar al menos un programa para instalar."
         )
         return
      
      # Confirmar
      response = messagebox.askyesno(
         "Confirmar Instalación",
         f"Se instalarán {len(selected)} programas.\n\n"
         f"⚠ Este proceso puede tardar varios minutos.\n"
         f"⚠ No apague ni reinicie el equipo durante la instalación.\n\n"
         f"¿Continuar?"
      )
      
      if not response:
         self.log("✗ Instalación cancelada por el usuario", "WARNING")
         return
      
      # Iniciar
      self.is_processing = True
      self.should_cancel = False
      self.button_start.configure(state="disabled")
      self.button_cancel.configure(state="normal")
      
      self.textbox.configure(state="normal")
      self.textbox.delete("1.0", "end")
      self.textbox.configure(state="disabled")
      
      # Resetear stats
      self.stats = {"total": 0, "installed": 0, "skipped": 0, "failed": 0}
      
      thread = threading.Thread(target=self.install_programs, daemon=True)
      thread.start()
   
   def install_programs(self):
      """Instala los programas seleccionados."""
      try:
         self.log("━" * 80)
         self.log("🚀 Iniciando instalación de programas", "INFO")
         self.log("━" * 80)
         
         # Obtener lista de instaladores seleccionados
         selected_installers = [
               inst for inst in INSTALLERS
               if self.installer_vars[inst["id"]].get()
         ]
         
         include_office = self.office_var.get()
         total = len(selected_installers) + (1 if include_office else 0)
         self.stats["total"] = total
         completed = 0
         
         # Instalar programas regulares
         for installer in selected_installers:
               if self.should_cancel:
                  self.log("✗ Proceso cancelado", "WARNING")
                  break
               
               self.progress_label.configure(text=f"Instalando: {installer['name']}")
               self.log(f"─── {installer['name']} ───", "INFO")
               
               # Verificar si ya está instalado
               if check_if_installed(installer["check_cmd"]):
                  self.log(f"⏭ Ya instalado, omitiendo...", "SKIP")
                  self.stats["skipped"] += 1
               else:
                  # Buscar instalador
                  installer_path = find_installer(installer["file"])
                  
                  if not installer_path:
                     self.log(f"✗ Archivo no encontrado: {installer['file']}", "ERROR")
                     self.stats["failed"] += 1
                  else:
                     self.log(f"  Ejecutando instalador...")
                     success, error = run_installer(installer_path, installer["args"])
                     
                     if success:
                           self.log(f"✓ Instalado correctamente", "SUCCESS")
                           self.stats["installed"] += 1
                     else:
                           self.log(f"✗ Error: {error}", "ERROR")
                           self.stats["failed"] += 1
               
               completed += 1
               self.progress_bar.set(completed / total)
               time.sleep(0.5)
         
         # Instalar Office
         if include_office and not self.should_cancel:
               self.progress_label.configure(text=f"Instalando: {OFFICE_INSTALLER['name']}")
               self.log(f"─── {OFFICE_INSTALLER['name']} ───", "INFO")
               
               if check_if_installed(OFFICE_INSTALLER["check_cmd"]):
                  self.log(f"⏭ Ya instalado, omitiendo...", "SKIP")
                  self.stats["skipped"] += 1
               else:
                  installer_path = find_installer(OFFICE_INSTALLER["file"])
                  config_path = find_installer(OFFICE_INSTALLER["config"])
                  
                  if not installer_path or not config_path:
                     self.log(f"✗ Archivos no encontrados", "ERROR")
                     self.stats["failed"] += 1
                  else:
                     self.log(f"  Ejecutando instalador (puede tardar varios minutos)...")
                     success, error = run_installer(
                           installer_path,
                           f"/configure {config_path}",
                           timeout=1800  # 30 minutos
                     )
                     
                     if success:
                           self.log(f"✓ Instalado correctamente", "SUCCESS")
                           self.stats["installed"] += 1
                     else:
                           self.log(f"✗ Error: {error}", "ERROR")
                           self.stats["failed"] += 1
               
               completed += 1
               self.progress_bar.set(completed / total)
         
         # Resumen
         self.log("━" * 80)
         self.log("📊 Resumen de Instalación", "INFO")
         self.log(f"  Total de programas: {self.stats['total']}")
         self.log(f"  ✓ Instalados:       {self.stats['installed']}", "SUCCESS")
         self.log(f"  ⏭ Omitidos:         {self.stats['skipped']}", "SKIP")
         self.log(f"  ✗ Fallidos:         {self.stats['failed']}", "ERROR" if self.stats['failed'] > 0 else "INFO")
         self.log("━" * 80)
         
         if not self.should_cancel:
               self.log("✓ Proceso completado", "SUCCESS")
               self.after(100, self.show_completion_dialog)
         
      except Exception as e:
         self.log(f"✗ Error crítico: {str(e)}", "ERROR")
      finally:
         self.is_processing = False
         self.after(100, lambda: self.button_start.configure(state="normal"))
         self.after(100, lambda: self.button_cancel.configure(state="disabled"))
         self.progress_label.configure(text="Proceso finalizado")
   
   def show_completion_dialog(self):
      """Muestra diálogo de completación."""
      message = (
         f"✓ Instalación completada\n\n"
         f"Programas instalados: {self.stats['installed']}\n"
         f"Programas omitidos: {self.stats['skipped']}\n"
         f"Programas fallidos: {self.stats['failed']}\n\n"
      )
      
      if self.stats['failed'] > 0:
         message += "⚠ Algunos programas no se instalaron correctamente.\n"
         message += "Revise el log para más detalles."
         messagebox.showwarning("Instalación Completada con Advertencias", message)
      else:
         message += "Todos los programas se instalaron correctamente."
         messagebox.showinfo("Instalación Completada", message)


# ============================================================================
# PUNTO DE ENTRADA
# ============================================================================

if __name__ == "__main__":
   app = InstallerApp()
   app.mainloop()