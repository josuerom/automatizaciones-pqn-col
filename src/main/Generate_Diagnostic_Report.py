"""
Generate_Diagnostic_Report.py - Versión Mejorada
Autor: Josué Romero
Empresa: Stefanini / PQN
Fecha: 31/Octubre/2025

Descripción:
Creador de informe técnico-práctico en PDF sobre mantenimiento físico/lógico realizado.
"""

import customtkinter as ctk
from tkinter import messagebox
from datetime import datetime
import socket
import platform
import subprocess
import psutil
import os
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
import getpass
from pathlib import Path
import re

# ============================================================================
# CONFIGURACIÓN GLOBAL
# ============================================================================
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("green")

APP_TITLE = "Generador de Informes IT"
APP_VERSION = "v2.0"
APP_SIZE = "750x810"

# Colores profesionales
COLOR_PRIMARY = "#66bb6a"
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


# ============================================================================
# FUNCIONES DE DETECCIÓN DE HARDWARE
# ============================================================================

def get_processor_info():
   """Obtiene información del procesador usando PowerShell CIM."""
   try:
      result = subprocess.run(
         ['powershell', '-Command',
            '(Get-CimInstance -ClassName Win32_Processor).Name'],
         capture_output=True,
         text=True,
         timeout=10
      )
      return result.stdout.strip() if result.returncode == 0 else "No disponible"
   except:
      return "No disponible"


def get_manufacturer():
   """Obtiene el fabricante del equipo usando PowerShell CIM."""
   try:
      result = subprocess.run(
         ['powershell', '-Command',
            '(Get-CimInstance -ClassName Win32_ComputerSystem).Manufacturer'],
         capture_output=True,
         text=True,
         timeout=10
      )
      return result.stdout.strip() if result.returncode == 0 else "No disponible"
   except:
      return "No disponible"


def get_model():
   """Obtiene el modelo del equipo."""
   try:
      result = subprocess.run(
         ['powershell', '-Command',
            '(Get-CimInstance -ClassName Win32_ComputerSystem).Model'],
         capture_output=True,
         text=True,
         timeout=10
      )
      return result.stdout.strip() if result.returncode == 0 else "No disponible"
   except:
      return "No disponible"


def get_bios_serial():
   """Obtiene el serial del BIOS."""
   try:
      result = subprocess.run(
         ['powershell', '-Command',
            '(Get-CimInstance -ClassName Win32_BIOS).SerialNumber'],
         capture_output=True,
         text=True,
         timeout=10
      )
      return result.stdout.strip() if result.returncode == 0 else "No disponible"
   except:
      return "No disponible"


def get_ram_info():
   """Obtiene información detallada de RAM."""
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


def get_disk_info():
   """Obtiene información de todos los discos."""
   disks = []
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
   return disks


def validate_ticket_number(ticket):
   """Valida formato de número de ticket."""
   if not ticket:
      return False, "El número de caso no puede estar vacío"
   
   # Aceptar números, letras y guiones
   if not re.match(r'^[A-Za-z0-9\-]+$', ticket):
      return False, "El número de caso solo puede contener letras, números y guiones"
   
   if len(ticket) < 3:
      return False, "El número de caso debe tener al menos 3 caracteres"
   
   return True, ""


def validate_fixed_asset(asset):
   """Valida formato de placa de activo fijo."""
   if not asset:
      return False, "La placa no puede estar vacía"
   
   if not re.match(r'^[A-Za-z0-9\-]+$', asset):
      return False, "La placa solo puede contener letras, números y guiones"
   
   return True, ""


# ============================================================================
# CLASE PRINCIPAL
# ============================================================================

class DiagnosticApp(ctk.CTk):
   def __init__(self):
      super().__init__()
      
      # Configuración de ventana
      self.title(f"{APP_TITLE} {APP_VERSION}")
      self.geometry(APP_SIZE)
      self.resizable(False, False)
      
      # Variables
      self.system_info = None
      self.is_generating = False
      
      # Construir interfaz
      self.build_ui()
      
      # Obtener información del sistema
      self.after(500, self.load_system_info)
   
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
         text="📄 " + APP_TITLE,
         font=FONT_TITLE,
         text_color="white"
      )
      title_label.pack(pady=15)
      
      subtitle_label = ctk.CTkLabel(
         header_frame,
         text=f"Autor: Josué Romero  |  Stefanini / PQN  |  {APP_VERSION}",
         font=FONT_SUBTITLE,
         text_color="#e8f5e9"
      )
      subtitle_label.pack(pady=(0, 15))
      
      # === INFORMACIÓN DEL SISTEMA ===
      info_frame = ctk.CTkFrame(main_frame, fg_color=COLOR_BG_LIGHT, corner_radius=8)
      info_frame.pack(fill="x", pady=(0, 10))
      
      info_title = ctk.CTkLabel(
         info_frame,
         text="💻 Información del Sistema Detectada",
         font=FONT_LABEL,
         text_color=COLOR_TEXT
      )
      info_title.pack(pady=(10, 5), anchor="w", padx=15)
      
      self.info_text = ctk.CTkTextbox(
         info_frame,
         width=700,
         height=80,
         font=("Consolas", 9),
         fg_color="#1e1e1e",
         text_color=COLOR_SUCCESS,
         border_width=1,
         border_color="#424242"
      )
      self.info_text.pack(pady=(0, 10), padx=15)
      self.info_text.insert("end", "Cargando información del sistema...\n")
      
      # === FORMULARIO ===
      form_frame = ctk.CTkFrame(main_frame, fg_color=COLOR_BG_LIGHT, corner_radius=8)
      form_frame.pack(fill="x", pady=(0, 10))
      
      form_title = ctk.CTkLabel(
         form_frame,
         text="✏️ Datos del Informe",
         font=FONT_LABEL,
         text_color=COLOR_TEXT
      )
      form_title.pack(pady=(10, 10), anchor="w", padx=15)
      
      # Técnico
      ctk.CTkLabel(
         form_frame,
         text="Nombre del técnico:",
         font=FONT_SUBTITLE,
         text_color=COLOR_TEXT
      ).pack(pady=(0, 5), padx=15, anchor="w")
      
      self.tecnico_entry = ctk.CTkEntry(
         form_frame,
         placeholder_text="Ej: Juan Pérez",
         width=400,
         height=35,
         font=FONT_SUBTITLE
      )
      self.tecnico_entry.pack(pady=(0, 10), padx=15, anchor="w")
      self.tecnico_entry.bind("<KeyRelease>", lambda e: self.validate_form())
      
      # Placa
      ctk.CTkLabel(
         form_frame,
         text="Número de placa (AF):",
         font=FONT_SUBTITLE,
         text_color=COLOR_TEXT
      ).pack(pady=(0, 5), padx=15, anchor="w")
      
      self.fixed_asset_entry = ctk.CTkEntry(
         form_frame,
         placeholder_text="Ej: 35094",
         width=400,
         height=35,
         font=FONT_SUBTITLE
      )
      self.fixed_asset_entry.pack(pady=(0, 10), padx=15, anchor="w")
      self.fixed_asset_entry.bind("<KeyRelease>", lambda e: self.validate_form())
      
      # Caso
      ctk.CTkLabel(
         form_frame,
         text="Número de caso en Mayté:",
         font=FONT_SUBTITLE,
         text_color=COLOR_TEXT
      ).pack(pady=(0, 5), padx=15, anchor="w")
      
      self.ticket_entry = ctk.CTkEntry(
         form_frame,
         placeholder_text="Ej: 41233",
         width=400,
         height=35,
         font=FONT_SUBTITLE
      )
      self.ticket_entry.pack(pady=(0, 5), padx=15, anchor="w")
      self.ticket_entry.bind("<KeyRelease>", lambda e: self.validate_form())
      self.ticket_entry.bind("<Return>", lambda e: self.generar_reporte())
      
      self.validation_label = ctk.CTkLabel(
         form_frame,
         text="",
         font=("Segoe UI", 9),
         text_color=COLOR_WARNING
      )
      self.validation_label.pack(pady=(0, 10), padx=15, anchor="w")
      
      # === LOG ===
      log_label = ctk.CTkLabel(
         main_frame,
         text="📝 Estado de Generación",
         font=FONT_LABEL,
         text_color=COLOR_TEXT,
         anchor="w"
      )
      log_label.pack(pady=(5, 5), anchor="w")
      
      self.output_box = ctk.CTkTextbox(
         main_frame,
         width=700,
         height=100,
         font=FONT_CONSOLE,
         fg_color=COLOR_BG_LIGHT,
         text_color=COLOR_TEXT,
         border_width=2,
         border_color=COLOR_PRIMARY,
         corner_radius=8
      )
      self.output_box.pack(pady=(0, 10))
      self.log("✓ Sistema listo para generar informes")
      
      # === BOTONES ===
      button_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
      button_frame.pack(fill="x")
      
      self.generate_button = ctk.CTkButton(
         button_frame,
         text="▶ Generar Informe PDF",
         command=self.generar_reporte,
         font=FONT_BUTTON,
         height=40,
         corner_radius=8,
         fg_color=COLOR_PRIMARY,
         hover_color="#43a047"
      )
      self.generate_button.pack(side="left", expand=True, fill="x", padx=(0, 5))
      
      self.clear_button = ctk.CTkButton(
         button_frame,
         text="🗑 Limpiar Campos",
         command=self.clear_fields,
         font=FONT_BUTTON,
         height=40,
         corner_radius=8,
         fg_color=COLOR_BG_LIGHT,
         hover_color="#424242"
      )
      self.clear_button.pack(side="left", expand=True, fill="x", padx=(5, 0))
   
   def log(self, msg):
      """Registra mensaje en el log."""
      timestamp = datetime.now().strftime("%H:%M:%S")
      self.output_box.configure(state="normal")
      self.output_box.insert("end", f"[{timestamp}] {msg}\n")
      self.output_box.see("end")
      self.output_box.configure(state="disabled")
   
   def load_system_info(self):
      """Carga información del sistema."""
      self.log("Detectando configuración del hardware...")
      
      try:
         hostname = socket.gethostname()
         os_info = f"{platform.system()} {platform.release()}"
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
         self.info_text.insert("end", f"PC: {hostname} | OS: {os_info}\n")
         self.info_text.insert("end", f"Fabricante: {manufacturer} | Modelo: {model}\n")
         self.info_text.insert("end", f"Procesador: {processor}\n")
         self.info_text.insert("end", f"RAM: {ram['total']}GB | Serial: {serial}")
         self.info_text.configure(state="disabled")
         
         self.log("✓ Información del sistema cargada correctamente")
         
      except Exception as e:
         self.log(f"⚠ Error al cargar información: {str(e)}")
   
   def validate_form(self):
      """Valida el formulario en tiempo real."""
      tecnico = self.tecnico_entry.get().strip()
      fixed_asset = self.fixed_asset_entry.get().strip()
      ticket = self.ticket_entry.get().strip()
      
      errors = []
      
      if not tecnico:
         errors.append("Falta nombre del técnico")
      
      if fixed_asset:
         is_valid, msg = validate_fixed_asset(fixed_asset)
         if not is_valid:
               errors.append(msg)
      else:
         errors.append("Falta número de placa")
      
      if ticket:
         is_valid, msg = validate_ticket_number(ticket)
         if not is_valid:
               errors.append(msg)
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
      """Limpia todos los campos."""
      self.tecnico_entry.delete(0, "end")
      self.fixed_asset_entry.delete(0, "end")
      self.ticket_entry.delete(0, "end")
      self.validation_label.configure(text="")
      self.log("Campos limpiados")
   
   def generar_reporte(self):
      """Genera el informe PDF."""
      if self.is_generating:
         return
      
      tecnico = self.tecnico_entry.get().strip()
      ticket = self.ticket_entry.get().strip()
      fixed_asset = self.fixed_asset_entry.get().strip()
      
      if not all([tecnico, ticket, fixed_asset]):
         messagebox.showwarning("Campos Requeridos", "Por favor completa todos los campos.")
         return
      
      self.is_generating = True
      self.generate_button.configure(state="disabled", text="⏳ Generando...")
      
      try:
         self.log("━" * 70)
         self.log("🚀 Iniciando generación de informe...")
         
         fecha_actual = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
         filename = f"Informe Diagnóstico Caso {ticket}.pdf"
         
         # Rutas de salida
         documentos = Path.home() / "Documents"
         if not documentos.exists():
               documentos = Path.home() / "Documentos"
         
         ruta_docs = documentos / filename
         ruta_datos = Path("D:/Datos") / filename
         
         # Crear directorios
         documentos.mkdir(parents=True, exist_ok=True)
         if Path("D:/").exists():
               Path("D:/Datos").mkdir(parents=True, exist_ok=True)
         
         self.log(f"Generando PDF: {filename}")
         
         # Generar PDF
         self.crear_pdf(str(ruta_docs), tecnico, ticket, fixed_asset, fecha_actual)
         self.log(f"✓ Guardado en: {ruta_docs}")
         
         if Path("D:/").exists():
               self.crear_pdf(str(ruta_datos), tecnico, ticket, fixed_asset, fecha_actual)
               self.log(f"✓ Guardado en: {ruta_datos}")
         
         # Abrir PDF
         try:
               if Path("D:/").exists() and ruta_datos.exists():
                  os.startfile(str(ruta_datos))
               else:
                  os.startfile(str(ruta_docs))
               self.log("✓ PDF abierto automáticamente")
         except:
               self.log("⚠ No se pudo abrir el PDF automáticamente")
         
         self.log("━" * 70)
         self.log("✓ Informe generado exitosamente")
         
         messagebox.showinfo(
               "Éxito",
               f"Informe generado correctamente:\n\n{filename}\n\n"
               f"Ubicación: {ruta_docs}"
         )
         
         # Cerrar aplicación
         self.after(2000, self.quit)
         
      except Exception as e:
         self.log(f"✗ Error al generar informe: {str(e)}")
         messagebox.showerror("Error", f"No se pudo generar el informe:\n\n{str(e)}")
      finally:
         self.is_generating = False
         self.generate_button.configure(state="normal", text="▶ Generar Informe PDF")
   
   def crear_pdf(self, path, tecnico, ticket, fixed_asset, fecha):
      """Crea el archivo PDF del informe."""
      c = canvas.Canvas(path, pagesize=letter)
      width, height = letter
      
      # Encabezado
      c.setFont("Helvetica-Bold", 12)
      c.drawString(50, height - 50, "Proquinal S.A.S")
      c.drawRightString(width - 50, height - 50, "Stefanini Group CO")
      
      # Título
      c.setFont("Helvetica-Bold", 18)
      c.drawCentredString(width / 2, height - 90, "INFORME DE DIAGNÓSTICO")
      
      y = height - 130
      
      def draw_field(label, value, x, y):
         c.setFont("Helvetica-Bold", 10)
         c.drawString(x, y, label)
         offset = c.stringWidth(label, "Helvetica-Bold", 10)
         c.setFont("Helvetica", 10)
         c.drawString(x + offset + 5, y, value)
      
      # Información general
      draw_field("Técnico:", tecnico, 50, y)
      draw_field("Caso:", ticket, width / 2, y)
      y -= 20
      draw_field("Placa:", fixed_asset, 50, y)
      draw_field("Fecha:", fecha, width / 2, y)
      
      y -= 30
      c.line(50, y, width - 50, y)
      y -= 25
      
      # Diagnóstico
      c.setFont("Helvetica-Bold", 12)
      c.drawString(50, y, "Diagnóstico del Sistema")
      y -= 20
      
      # Información del sistema
      info = self.system_info
      ram = info['ram']
      disks = info['disks']
      
      items = [
         f"Usuario: {info['user']}",
         f"Host: {info['hostname']}",
         f"SO: {info['os']}",
         f"CPU: {info['processor'][:60]}",
         f"Fabricante: {info['manufacturer']}",
         f"Modelo: {info['model'][:40]}",
         f"Serial: {info['serial']}",
         f"RAM Total: {ram['total']} GB",
         f"RAM Usada: {ram['used']} GB ({ram['percent']}%)",
      ]
      
      for disk in disks:
         items.append(f"Disco {disk['drive']}: {disk['total']}GB (Libre: {disk['free']}GB)")
      
      c.setFont("Helvetica", 10)
      for item in items:
         if y < 100:
               c.showPage()
               y = height - 80
               c.setFont("Helvetica", 10)
         c.drawString(50, y, f"• {item}")
         y -= 15
      
      y -= 20
      
      # Procedimiento
      c.setFont("Helvetica-Bold", 12)
      c.drawString(50, y, "Procedimiento Realizado")
      y -= 20
      
      pasos = [
         "1. Se retiraron los tornillos de la carcasa inferior.",
         "2. Se accedió al hardware interno del equipo.",
         "3. Se desconectó la batería para evitar descargas.",
         "4. Se limpió el polvo con aire comprimido.",
         "5. Se retiró y limpió completamente la pasta térmica.",
         "6. Se aplicó nueva pasta térmica de alta calidad.",
         "7. Se reinstalaron todos los componentes.",
         "8. Se realizaron pruebas de encendido y estabilidad.",
         "9. Se validó el estado general del sistema.",
         "10. Se verificó el correcto funcionamiento."
      ]
      
      c.setFont("Helvetica", 10)
      for paso in pasos:
         if y < 100:
               c.showPage()
               y = height - 80
               c.setFont("Helvetica", 10)
         c.drawString(50, y, paso)
         y -= 18
      
      # Recomendaciones
      y -= 10
      c.setFont("Helvetica-Bold", 12)
      c.drawString(50, y, "Recomendaciones")
      y -= 20
      
      c.setFont("Helvetica", 10)
      recomendaciones = [
         "• Realizar mantenimiento preventivo 1-2 veces al año.",
         "• Evitar cargador conectado permanentemente durante el día.",
         "• Apagar el equipo preferiblemente una vez al día.",
         "• Mantener limpias las rejillas de ventilación."
      ]
      
      for rec in recomendaciones:
         if y < 80:
               c.showPage()
               y = height - 80
               c.setFont("Helvetica", 10)
         c.drawString(50, y, rec)
         y -= 18
      
      # Pie de página
      c.setFont("Helvetica-Oblique", 8)
      c.drawCentredString(width / 2, 20, f"Informe generado por {APP_TITLE} {APP_VERSION} | ©2025 @josuerom")
      
      c.save()


# ============================================================================
# PUNTO DE ENTRADA
# ============================================================================

if __name__ == "__main__":
   app = DiagnosticApp()
   app.mainloop()