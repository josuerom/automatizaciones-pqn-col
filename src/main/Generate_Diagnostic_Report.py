"""
Generate_Diagnostic_Report.py - Versión Profesional Mejorada
Autor: Josué Romero
Empresa: Stefanini / PQN
Fecha: 06/Noviembre/2025
Versión: 3.0 Professional Edition

Licencia: Propiedad de Stefanini / PQN - Todos los derechos reservados
Copyright © 2025 Josué Romero - Stefanini / PQN

Descripción:
Sistema automatizado para generación de informes técnicos en PDF
sobre mantenimiento físico/lógico realizado en equipos.
"""

import customtkinter as ctk
from tkinter import messagebox
from datetime import datetime
import socket
import platform
import subprocess
import psutil
import os
import sys
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib.utils import ImageReader
import getpass
from pathlib import Path
import re
import hashlib
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
ctk.set_default_color_theme("green")

APP_TITLE = "Generador de Informes IT"
APP_VERSION = f"v{__version__}"
APP_SIZE = "800x990"

# Paleta de colores profesional (Esquema Naranja-Ámbar-Oscuro)
COLOR_PRIMARY = "#e65100"         # Naranja profundo
COLOR_SECONDARY = "#ff6f00"       # Naranja brillante
COLOR_ACCENT = "#ff9800"          # Ámbar
COLOR_SUCCESS = "#00c853"         # Verde éxito
COLOR_WARNING = "#ffc107"         # Amarillo advertencia
COLOR_ERROR = "#d32f2f"           # Rojo error
COLOR_BG_DARK = "#1a1410"         # Fondo oscuro principal
COLOR_BG_MEDIUM = "#2d2416"       # Fondo medio
COLOR_BG_LIGHT = "#3d321e"        # Fondo claro
COLOR_TEXT_WHITE = "#ffffff"      # Texto blanco
COLOR_TEXT_GRAY = "#b0bec5"       # Texto gris claro

# Fuentes (Incrementadas +2px según requerimiento)
FONT_TITLE = ("Segoe UI", 28, "bold")          # Era 26
FONT_SUBTITLE = ("Segoe UI", 13)               # Era 11
FONT_CONSOLE = ("Consolas", 12)                # Era 10
FONT_BUTTON = ("Segoe UI", 14, "bold")         # Era 12
FONT_LABEL = ("Segoe UI", 13, "bold")          # Era 11
FONT_INFO = ("Segoe UI", 13)                   # Era 11

# Rutas de logos (relativas al script)
LOGO_DIR = Path(__file__).parent.parent.parent / "static" / "logos"
LOGO_PROQUINAL = LOGO_DIR / "proquinal.png"
LOGO_MAYTE = LOGO_DIR / "mayte.png"
LOGO_STEFANINI = LOGO_DIR / "stefanini.png"


# ============================================================================
# FUNCIONES DE ELEVACIÓN DE PRIVILEGIOS
# ============================================================================

def is_admin():
   """Verifica si el script se está ejecutando como administrador."""
   try:
      return ctypes.windll.shell32.IsUserAnAdmin()
   except:
      return False


# ============================================================================
# FUNCIONES AUXILIARES
# ============================================================================

def calculate_file_hash(filepath):
   """Calcula el hash SHA-256 de un archivo."""
   try:
      sha256_hash = hashlib.sha256()
      with open(filepath, "rb") as f:
         for byte_block in iter(lambda: f.read(4096), b""):
               sha256_hash.update(byte_block)
      return sha256_hash.hexdigest()
   except:
      return "N/A"


def get_processor_info():
   """Obtiene información del procesador usando PowerShell CIM."""
   try:
      result = subprocess.run(
         ['powershell', '-NoProfile', '-Command',
            '(Get-CimInstance -ClassName Win32_Processor).Name'],
         capture_output=True,
         text=True,
         timeout=10,
         creationflags=subprocess.CREATE_NO_WINDOW
      )
      return result.stdout.strip() if result.returncode == 0 else "No disponible"
   except:
      return "No disponible"


def get_manufacturer():
   """Obtiene el fabricante del equipo usando PowerShell CIM."""
   try:
      result = subprocess.run(
         ['powershell', '-NoProfile', '-Command',
            '(Get-CimInstance -ClassName Win32_ComputerSystem).Manufacturer'],
         capture_output=True,
         text=True,
         timeout=10,
         creationflags=subprocess.CREATE_NO_WINDOW
      )
      return result.stdout.strip() if result.returncode == 0 else "No disponible"
   except:
      return "No disponible"


def get_model():
   """Obtiene el modelo del equipo."""
   try:
      result = subprocess.run(
         ['powershell', '-NoProfile', '-Command',
            '(Get-CimInstance -ClassName Win32_ComputerSystem).Model'],
         capture_output=True,
         text=True,
         timeout=10,
         creationflags=subprocess.CREATE_NO_WINDOW
      )
      return result.stdout.strip() if result.returncode == 0 else "No disponible"
   except:
      return "No disponible"


def get_bios_serial():
   """Obtiene el serial del BIOS."""
   try:
      result = subprocess.run(
         ['powershell', '-NoProfile', '-Command',
            '(Get-CimInstance -ClassName Win32_BIOS).SerialNumber'],
         capture_output=True,
         text=True,
         timeout=10,
         creationflags=subprocess.CREATE_NO_WINDOW
      )
      return result.stdout.strip() if result.returncode == 0 else "No disponible"
   except:
      return "No disponible"


def get_ram_info():
   """Obtiene información detallada de RAM."""
   try:
      total_gb = round(psutil.virtual_memory().total / (1024 ** 3), 2)
      available_gb = round(psutil.virtual_memory().available / (1024 ** 3), 2)
      used_gb = round((psutil.virtual_memory().total - psutil.virtual_memory().available) / (1024 ** 3), 2)
      percent = psutil.virtual_memory().percent
      
      return {
         "total": total_gb,
         "available": available_gb,
         "used": used_gb,
         "percent": percent
      }
   except:
      return {
         "total": 0,
         "available": 0,
         "used": 0,
         "percent": 0
      }


def get_disk_info():
   """Obtiene información de todos los discos."""
   disks = []
   try:
      for partition in psutil.disk_partitions():
         if 'cdrom' in partition.opts or partition.fstype == '':
               continue
         try:
               usage = psutil.disk_usage(partition.mountpoint)
               disks.append({
                  "drive": partition.device,
                  "total": round(usage.total / (1024 ** 3), 2),
                  "used": round(usage.used / (1024 ** 3), 2),
                  "free": round(usage.free / (1024 ** 3), 2),
                  "percent": usage.percent
               })
         except:
               continue
   except:
      pass
   return disks


def validate_ticket_number(ticket):
   """Valida formato de número de ticket."""
   if not ticket:
      return False, "El número de caso no puede estar vacío"
   
   if not re.match(r'^[0-9\-]+$', ticket):
      return False, "Solo puede contener números"
   
   if len(ticket) < 5:
      return False, "Debe tener al menos 5 digitos"
   
   return True, ""


def validate_fixed_asset(asset):
   """Valida formato de placa de activo fijo."""
   if not asset:
      return False, "La placa no puede estar vacía"
   
   if not re.match(r'^[0-9\-]+$', asset):
      return False, "Solo puede contener números"
   
   if len(asset) < 5:
      return False, "Debe tener al menos 5 digitos"
   
   return True, ""


def validate_technician_name(name):
   """Valida el nombre del técnico."""
   if not name:
      return False, "El nombre no puede estar vacío"
   
   if len(name) < 10:
      return False, "Debe tener al menos 10 caracteres"
   
   if not re.match(r'^[A-Za-zÁÉÍÓÚáéíóúÑñ\s]+$', name):
      return False, "Solo puede contener letras y espacios"
   
   return True, ""


# ============================================================================
# CLASE PRINCIPAL DE LA APLICACIÓN
# ============================================================================

class DiagnosticApp(ctk.CTk):
   def __init__(self):
      super().__init__()
      
      # Configuración de ventana
      self.title(f"{APP_TITLE} {APP_VERSION}")
      self.geometry(APP_SIZE)
      self.resizable(False, False)
      
      # Variables de estado
      self.system_info = None
      self.is_generating = False
      
      # Construir interfaz
      self.build_ui()
      
      # Obtener información del sistema
      self.after(500, self.load_system_info)
   
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
         text="📄 " + APP_TITLE,
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
         text="💻 Información del Sistema Detectada",
         font=FONT_LABEL,
         text_color=COLOR_TEXT_WHITE
      )
      info_title.pack(pady=(15, 10), anchor="w", padx=20)
      
      self.info_text = ctk.CTkTextbox(
         info_frame,
         width=740,
         height=90,
         font=("Consolas", 11),
         fg_color=COLOR_BG_LIGHT,
         text_color=COLOR_SUCCESS,
         border_width=2,
         border_color=COLOR_SECONDARY
      )
      self.info_text.pack(pady=(0, 15), padx=20)
      self.info_text.insert("end", "Cargando información del sistema...\n")
      self.info_text.configure(state="disabled")
      
      # === FORMULARIO ===
      form_frame = ctk.CTkFrame(
         main_frame,
         fg_color=COLOR_BG_MEDIUM,
         corner_radius=10
      )
      form_frame.pack(fill="x", pady=(0, 15))
      
      form_title = ctk.CTkLabel(
         form_frame,
         text="✏️ Datos del Informe",
         font=FONT_LABEL,
         text_color=COLOR_TEXT_WHITE
      )
      form_title.pack(pady=(15, 10), anchor="w", padx=20)
      
      # Técnico
      ctk.CTkLabel(
         form_frame,
         text="Nombre del técnico:",
         font=FONT_INFO,
         text_color=COLOR_TEXT_WHITE
      ).pack(pady=(0, 8), padx=20, anchor="w")
      
      self.tecnico_entry = ctk.CTkEntry(
         form_frame,
         placeholder_text="Ej: Juan Pérez García",
         width=500,
         height=40,
         font=FONT_INFO,
         fg_color=COLOR_BG_LIGHT,
         text_color=COLOR_TEXT_WHITE,
         placeholder_text_color=COLOR_TEXT_GRAY,
         border_color=COLOR_SECONDARY,
         border_width=2
      )
      self.tecnico_entry.pack(pady=(0, 15), padx=20, anchor="w")
      self.tecnico_entry.bind("<KeyRelease>", lambda e: self.validate_form())
      
      # Placa
      ctk.CTkLabel(
         form_frame,
         text="Número de placa (Activo Fijo):",
         font=FONT_INFO,
         text_color=COLOR_TEXT_WHITE
      ).pack(pady=(0, 8), padx=20, anchor="w")
      
      self.fixed_asset_entry = ctk.CTkEntry(
         form_frame,
         placeholder_text="Ej: 35094, AF-12345",
         width=500,
         height=40,
         font=FONT_INFO,
         fg_color=COLOR_BG_LIGHT,
         text_color=COLOR_TEXT_WHITE,
         placeholder_text_color=COLOR_TEXT_GRAY,
         border_color=COLOR_SECONDARY,
         border_width=2
      )
      self.fixed_asset_entry.pack(pady=(0, 15), padx=20, anchor="w")
      self.fixed_asset_entry.bind("<KeyRelease>", lambda e: self.validate_form())
      
      # Caso
      ctk.CTkLabel(
         form_frame,
         text="Número de caso en Mayté:",
         font=FONT_INFO,
         text_color=COLOR_TEXT_WHITE
      ).pack(pady=(0, 8), padx=20, anchor="w")
      
      self.ticket_entry = ctk.CTkEntry(
         form_frame,
         placeholder_text="Ej: 41233, CASE-2024-001",
         width=500,
         height=40,
         font=FONT_INFO,
         fg_color=COLOR_BG_LIGHT,
         text_color=COLOR_TEXT_WHITE,
         placeholder_text_color=COLOR_TEXT_GRAY,
         border_color=COLOR_SECONDARY,
         border_width=2
      )
      self.ticket_entry.pack(pady=(0, 8), padx=20, anchor="w")
      self.ticket_entry.bind("<KeyRelease>", lambda e: self.validate_form())
      self.ticket_entry.bind("<Return>", lambda e: self.generar_reporte())
      
      self.validation_label = ctk.CTkLabel(
         form_frame,
         text="",
         font=("Segoe UI", 11),
         text_color=COLOR_WARNING
      )
      self.validation_label.pack(pady=(0, 15), padx=20, anchor="w")
      
      # === ÁREA DE LOGS ===
      log_label = ctk.CTkLabel(
         main_frame,
         text="📝 Estado de Generación",
         font=FONT_LABEL,
         text_color=COLOR_TEXT_WHITE,
         anchor="w"
      )
      log_label.pack(pady=(5, 8), anchor="w")
      
      self.output_box = ctk.CTkTextbox(
         main_frame,
         width=740,
         height=120,
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
      
      # === BOTONES ===
      button_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
      button_frame.pack(fill="x")
      
      self.generate_button = ctk.CTkButton(
         button_frame,
         text="🚀 Generar Informe PDF",
         command=self.generar_reporte,
         font=FONT_BUTTON,
         height=50,
         corner_radius=10,
         fg_color=COLOR_PRIMARY,
         hover_color=COLOR_SECONDARY,
         border_width=3,
         border_color=COLOR_ACCENT,
         text_color=COLOR_TEXT_WHITE,
         state="disabled"
      )
      self.generate_button.pack(side="left", expand=True, fill="x", padx=(0, 8))
      
      self.clear_button = ctk.CTkButton(
         button_frame,
         text="🗑️ Limpiar Campos",
         command=self.clear_fields,
         font=FONT_BUTTON,
         height=50,
         corner_radius=10,
         fg_color=COLOR_BG_LIGHT,
         hover_color=COLOR_BG_MEDIUM,
         border_width=2,
         border_color=COLOR_SECONDARY,
         text_color=COLOR_TEXT_WHITE
      )
      self.clear_button.pack(side="left", fill="x", padx=(8, 0))
      
      # Atajos de teclado
      self.bind("<Escape>", lambda e: self.quit())
   
   def log(self, msg, level="INFO"):
      """Registra mensajes en el log con formato."""
      timestamp = datetime.now().strftime("%H:%M:%S")
      icons = {
         "INFO": "ℹ",
         "SUCCESS": "✓",
         "WARNING": "⚠",
         "ERROR": "✗",
         "PROCESS": "⚙"
      }
      icon = icons.get(level, "•")
      
      formatted_msg = f"[{timestamp}] {icon} {msg}\n"
      
      self.output_box.configure(state="normal")
      self.output_box.insert("end", formatted_msg)
      self.output_box.see("end")
      self.output_box.configure(state="disabled")
   
   def load_system_info(self):
      """Carga información del sistema."""
      self.log("Detectando configuración del hardware...", "PROCESS")
      
      try:
         hostname = socket.gethostname()
         os_info = f"{platform.system()} {platform.release()} {platform.version()}"
         processor = get_processor_info()
         manufacturer = get_manufacturer()
         model = get_model()
         serial = get_bios_serial()
         ram = get_ram_info()
         disks = get_disk_info()
         
         self.system_info = {
               "hostname": hostname,
               "os": os_info,
               "processor": processor,
               "manufacturer": manufacturer,
               "model": model,
               "serial": serial,
               "ram": ram,
               "disks": disks,
               "user": getpass.getuser()
         }
         
         # Actualizar cuadro de información
         self.info_text.configure(state="normal")
         self.info_text.delete("1.0", "end")
         self.info_text.insert("end", f"PC: {hostname} | Usuario: {getpass.getuser()}\n")
         self.info_text.insert("end", f"SO: {os_info}\n")
         self.info_text.insert("end", f"Fabricante: {manufacturer} | Modelo: {model}\n")
         self.info_text.insert("end", f"CPU: {processor[:60]}...\n")
         self.info_text.insert("end", f"RAM: {ram['total']}GB | Serial: {serial}")
         self.info_text.configure(state="disabled")
         
         self.log("✓ Información del sistema cargada correctamente", "SUCCESS")
         self.log(f"✓ Hardware: {manufacturer} {model}", "SUCCESS")
         self.log(f"✓ RAM: {ram['total']}GB | Discos: {len(disks)}", "SUCCESS")
         self.log("━" * 75, "INFO")
         self.log("✓ Listo para generar informes", "SUCCESS")
         
      except Exception as e:
         self.log(f"✗ Error al cargar información: {str(e)}", "ERROR")
         messagebox.showerror(
               "Error del Sistema",
               f"No se pudo obtener la información del sistema:\n\n{e}"
         )
   
   def validate_form(self):
      """Valida el formulario en tiempo real."""
      tecnico = self.tecnico_entry.get().strip()
      fixed_asset = self.fixed_asset_entry.get().strip()
      ticket = self.ticket_entry.get().strip()
      
      errors = []
      
      if tecnico:
         is_valid, msg = validate_technician_name(tecnico)
         if not is_valid:
               errors.append(f"Técnico: {msg}")
      else:
         errors.append("Falta nombre del técnico")
      
      if fixed_asset:
         is_valid, msg = validate_fixed_asset(fixed_asset)
         if not is_valid:
               errors.append(f"Placa: {msg}")
      else:
         errors.append("Falta número de placa")
      
      if ticket:
         is_valid, msg = validate_ticket_number(ticket)
         if not is_valid:
               errors.append(f"Caso: {msg}")
      else:
         errors.append("Falta número de caso")
      
      if errors:
         self.validation_label.configure(
               text="✗ " + " | ".join(errors),
               text_color=COLOR_ERROR
         )
         self.generate_button.configure(state="disabled")
      else:
         self.validation_label.configure(
               text="✓ Todos los campos son válidos",
               text_color=COLOR_SUCCESS
         )
         self.generate_button.configure(state="normal")
   
   def clear_fields(self):
      """Limpia todos los campos del formulario."""
      self.tecnico_entry.delete(0, "end")
      self.fixed_asset_entry.delete(0, "end")
      self.ticket_entry.delete(0, "end")
      self.validation_label.configure(text="")
      self.generate_button.configure(state="disabled")
      self.log("Campos limpiados correctamente", "INFO")
      self.tecnico_entry.focus()
   
   def generar_reporte(self):
      """Genera el informe PDF completo."""
      if self.is_generating:
         self.log("⚠ Ya hay un proceso de generación en curso", "WARNING")
         return
      
      tecnico = self.tecnico_entry.get().strip()
      ticket = self.ticket_entry.get().strip()
      fixed_asset = self.fixed_asset_entry.get().strip()
      
      if not all([tecnico, ticket, fixed_asset]):
         messagebox.showwarning(
               "Campos Requeridos",
               "Por favor completa todos los campos antes de generar el informe."
         )
         return
      
      self.is_generating = True
      self.generate_button.configure(
         state="disabled",
         text="⏳ Generando PDF...",
         fg_color=COLOR_BG_MEDIUM
      )
      
      try:
         self.log("━" * 75, "INFO")
         self.log("🚀 INICIANDO GENERACIÓN DE INFORME PDF", "PROCESS")
         self.log("━" * 75, "INFO")
         
         fecha_actual = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
         filename = f"Informe_Diagnostico_Caso_{ticket}.pdf"
         
         # Sanitizar nombre de archivo
         filename = re.sub(r'[<>:"/\\|?*]', '_', filename)
         
         # Rutas de salida
         documentos = Path.home() / "Documents"
         if not documentos.exists():
               documentos = Path.home() / "Documentos"
         
         ruta_docs = documentos / filename
         ruta_datos = Path("D:/Datos") / filename
         
         # Crear directorios
         documentos.mkdir(parents=True, exist_ok=True)
         self.log(f"✓ Directorio de salida: {documentos}", "SUCCESS")
         
         if Path("D:/").exists():
               Path("D:/Datos").mkdir(parents=True, exist_ok=True)
               self.log(f"✓ Directorio D:/Datos creado", "SUCCESS")
         
         self.log(f"Generando archivo: {filename}", "PROCESS")
         
         # Generar PDF
         self.crear_pdf(str(ruta_docs), tecnico, ticket, fixed_asset, fecha_actual)
         self.log(f"✓ PDF guardado en: {ruta_docs}", "SUCCESS")
         
         # Calcular hash
         file_hash = calculate_file_hash(str(ruta_docs))
         self.log(f"✓ SHA-256: {file_hash[:32]}...", "SUCCESS")
         
         # Copia en D:/
         if Path("D:/").exists():
               self.crear_pdf(str(ruta_datos), tecnico, ticket, fixed_asset, fecha_actual)
               self.log(f"✓ Copia de respaldo en: {ruta_datos}", "SUCCESS")
         
         # Abrir PDF automáticamente
         try:
               if Path("D:/").exists() and ruta_datos.exists():
                  os.startfile(str(ruta_datos))
               else:
                  os.startfile(str(ruta_docs))
               self.log("✓ PDF abierto automáticamente", "SUCCESS")
         except Exception as e:
               self.log(f"⚠ No se pudo abrir el PDF: {e}", "WARNING")
         
         self.log("━" * 75, "INFO")
         self.log("✓ INFORME GENERADO EXITOSAMENTE", "SUCCESS")
         self.log("━" * 75, "INFO")
         
         messagebox.showinfo(
               "✓ Informe Generado",
               f"El informe ha sido generado exitosamente:\n\n"
               f"📄 Archivo: {filename}\n"
               f"📁 Ubicación: {ruta_docs}\n"
               f"🔐 SHA-256: {file_hash[:16]}...\n\n"
               f"El archivo se ha abierto automáticamente.",
               icon='info'
         )
         
         # Cerrar aplicación después de 2 segundos
         self.after(2000, self.quit)
         
      except Exception as e:
         self.log("━" * 75, "ERROR")
         self.log(f"✗ ERROR CRÍTICO: {str(e)}", "ERROR")
         self.log("━" * 75, "ERROR")
         messagebox.showerror(
               "Error en Generación",
               f"No se pudo generar el informe:\n\n{str(e)}\n\n"
               "Posibles causas:\n"
               "• Falta librería reportlab\n"
               "• Sin permisos de escritura\n"
               "• Disco lleno\n"
               "• Logos no encontrados"
         )
      finally:
         self.is_generating = False
         self.generate_button.configure(
               state="normal",
               text="🚀 Generar Informe PDF",
               fg_color=COLOR_PRIMARY
         )

   def crear_pdf(self, path, tecnico, ticket, fixed_asset, fecha):
      """
      Crea el archivo PDF del informe con logos corporativos.

      Args:
         path: Ruta donde guardar el PDF
         tecnico: Nombre del técnico
         ticket: Número de caso
         fixed_asset: Placa de activo fijo
         fecha: Fecha y hora del informe
      """
      c = canvas.Canvas(path, pagesize=letter)
      width, height = letter

      # Nuevos colores (normalizados a 0..1)
      # NEGRO MATE: #0b0b0b
      BLACK_MATE = (0.043137, 0.043137, 0.043137)   # CAMBIO: antes blanco (1,1,1)
      # AZUL PROFESIONAL: #0b5cff
      AZUL = (0.043137, 0.360784, 1.0)              # CAMBIO: antes naranja (1,0.6,0)
      # GRIS OSCURO: #4a4a4a
      GRIS_OSCURO = (0.290196, 0.290196, 0.290196)  # CAMBIO: antes gris claro (0.9,0.9,0.9)

      # ===================================================================
      # ENCABEZADO CON LOGOS
      # ===================================================================
      try:
         # Logo Proquinal (Esquina superior izquierda)
         if LOGO_PROQUINAL.exists():
               c.drawImage(
                  str(LOGO_PROQUINAL),
                  50, height - 80,
                  width=100, height=40,
                  preserveAspectRatio=True,
                  mask='auto'
               )

         # Logo Mayté (Centro superior)
         if LOGO_MAYTE.exists():
               c.drawImage(
                  str(LOGO_MAYTE),
                  width/2 - 50, height - 80,
                  width=100, height=40,
                  preserveAspectRatio=True,
                  mask='auto'
               )

         # Logo Stefanini (Esquina superior derecha)
         if LOGO_STEFANINI.exists():
               c.drawImage(
                  str(LOGO_STEFANINI),
                  width - 150, height - 80,
                  width=100, height=40,
                  preserveAspectRatio=True,
                  mask='auto'
               )
      except Exception as e:
         self.log(f"⚠ Error cargando logos: {e}", "WARNING")

      # Línea divisoria (AZUL)
      c.setStrokeColorRGB(*AZUL)          # CAMBIO
      c.setLineWidth(2)
      c.line(50, height - 90, width - 50, height - 90)

      # ===================================================================
      # TÍTULO DEL INFORME
      # ===================================================================
      c.setFont("Helvetica-Bold", 18)
      c.setFillColorRGB(*AZUL)            # CAMBIO: título en AZUL
      c.drawCentredString(width / 2, height - 120, "INFORME DE DIAGNÓSTICO TÉCNICO")

      # Fecha y datos del caso (NEGRO MATE)
      c.setFont("Helvetica", 11)
      c.setFillColorRGB(*BLACK_MATE)      # CAMBIO
      c.drawString(50, height - 150, f"Fecha de generación: {fecha}")
      c.drawString(50, height - 165, f"Responsable: {tecnico}")
      c.drawString(50, height - 180, f"Caso Mayté: {ticket}")
      c.drawString(50, height - 195, f"Activo Fijo: {fixed_asset}")

      # Línea divisoria (AZUL más suave)
      c.setStrokeColorRGB(0.0, 0.3, 0.8)   # CAMBIO: tono de azul para la línea secundaria
      c.setLineWidth(1)
      c.line(50, height - 210, width - 50, height - 210)

      # ===================================================================
      # INFORMACIÓN DEL EQUIPO
      # ===================================================================
      y = height - 240
      c.setFont("Helvetica-Bold", 13)
      c.setFillColorRGB(0.043137, 0.360784, 1.0)   # CAMBIO: encabezado sección en AZUL
      c.drawString(50, y, "🔧 Información del Sistema")

      c.setFont("Helvetica", 10)
      c.setFillColorRGB(*BLACK_MATE)       # CAMBIO
      y -= 20
      sysinfo = self.system_info
      lines = [
         f"Equipo / Hostname: {sysinfo['hostname']}",
         f"Sistema Operativo: {sysinfo['os']}",
         f"Fabricante: {sysinfo['manufacturer']}",
         f"Modelo: {sysinfo['model']}",
         f"Serial BIOS: {sysinfo['serial']}",
         f"Procesador: {sysinfo['processor']}",
         f"RAM Total: {sysinfo['ram']['total']} GB",
      ]

      for line in lines:
         c.drawString(60, y, line)
         y -= 15

      # Espacio para discos
      y -= 10
      c.setFont("Helvetica-Bold", 13)
      c.setFillColorRGB(*AZUL)            # CAMBIO
      c.drawString(50, y, "🔧 Unidades de Almacenamiento:")
      y -= 15

      c.setFont("Helvetica", 10)
      c.setFillColorRGB(*BLACK_MATE)      # CAMBIO
      for disk in sysinfo["disks"]:
         c.drawString(60, y, f"{disk['drive']} - {disk['total']} GB (Usado {disk['used']} GB, Libre {disk['free']} GB)")
         y -= 13

      # ===================================================================
      # SECCIÓN DE OBSERVACIONES TÉCNICAS
      # ===================================================================
      y -= 25
      c.setFont("Helvetica-Bold", 13)
      c.setFillColorRGB(*AZUL)            # CAMBIO
      c.drawString(50, y, "🔧 Observaciones Técnicas / Mantenimiento Realizado:")
      y -= 18
      c.setFont("Helvetica-Oblique", 10)
      c.setFillColorRGB(*BLACK_MATE)      # CAMBIO
      c.drawString(60, y, "1. Limpieza física y lógica del equipo.")
      y -= 13
      c.drawString(60, y, "2. Cambio de pasta térmica al procesador.")
      y -= 13
      c.drawString(60, y, "3. Revisión de integridad de hardware (RAM, SSD, BIOS).")
      y -= 13
      c.drawString(60, y, "4. Optimización de arranque y procesos residuales.")
      y -= 13
      c.drawString(60, y, "5. Verificación de conectividad y servicios de red.")
      y -= 13
      c.drawString(60, y, "6. Actualizaciones de programas instalados y Windows.")
      y -= 13
      c.drawString(60, y, "7. Aplicación de políticas de SI - SpradlingGroup.")
      y -= 13
      c.drawString(60, y, "8. Presentar el equipo 1-2 veces al año para mantenimiento preventivo / correctivo.")

      # ===================================================================
      # FIRMA DIGITAL Y PIE DE PÁGINA
      # ===================================================================
      y = 100
      c.setFont("Helvetica-Bold", 11)
      c.setFillColorRGB(*AZUL)            # CAMBIO
      c.drawString(50, y + 40, "_______________________________")
      c.drawString(50, y + 25, tecnico)
      c.setFont("Helvetica", 9)
      c.setFillColorRGB(*BLACK_MATE)      # CAMBIO
      c.drawString(50, y + 12, "Técnico de Soporte Stefanini")

      # Línea divisoria final (AZUL)
      c.setStrokeColorRGB(*AZUL)          # CAMBIO
      c.setLineWidth(0.5)
      c.line(50, 80, width - 50, 80)

      # Hash de seguridad y pie de página (GRIS OSCURO)
      file_hash = calculate_file_hash(path)
      c.setFont("Helvetica", 7)
      c.setFillColorRGB(*GRIS_OSCURO)     # CAMBIO
      c.drawString(50, 65, f"SHA-256: {file_hash}")
      c.drawString(50, 50, "Documento Firmado")
      c.drawRightString(width - 50, 50, f"© 2025 @JOSUEROM | STEF - PQN | VERSIÓN {__version__}")

      # Guardar PDF
      c.save()


if __name__ == "__main__":
   # Elevar privilegios si no es administrador
   if not is_admin():
      ctypes.windll.shell32.ShellExecuteW(
            None,
            "runas",
            sys.executable,
            ' '.join([f'"{arg}"' for arg in sys.argv]),
            None,
            1
      )
      sys.exit(0)
   
   # Iniciar aplicación
   app = DiagnosticApp()
   app.mainloop()