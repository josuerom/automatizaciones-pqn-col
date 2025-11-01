"""
PQN_Access_Credentials.py - Versión Mejorada
Autor: Josué Romero
Empresa: Stefanini / PQN
Fecha: 31/Octubre/2025

Descripción:
Generador automático de credenciales corporativas de acceso.
"""

import os
import platform
import customtkinter as ctk
from tkinter import messagebox
from reportlab.lib.pagesizes import letter
from reportlab.platypus import Table, TableStyle, Paragraph, Spacer, SimpleDocTemplate
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.lib.units import inch
from pathlib import Path
import re
import random
import string
from datetime import datetime

# ============================================================================
# CONFIGURACIÓN GLOBAL
# ============================================================================
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("green")

APP_TITLE = "Generador de Credenciales PQN"
APP_VERSION = "v2.0"
APP_SIZE = "650x920"

# Colores
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
FONT_BUTTON = ("Segoe UI", 12, "bold")
FONT_LABEL = ("Segoe UI", 11, "bold")

# Constantes
DOMAIN = "@spradling.group"


# ============================================================================
# FUNCIONES AUXILIARES
# ============================================================================

def validate_username(username):
   """
   Valida formato de nombre de usuario.
   
   Returns:
      tuple: (is_valid: bool, error_msg: str)
   """
   if not username:
      return False, "El usuario no puede estar vacío"
   
   # Solo letras minúsculas, números y guiones
   if not re.match(r'^[a-z0-9\-]+$', username):
      return False, "Solo letras minúsculas, números y guiones"
   
   if len(username) < 3:
      return False, "Mínimo 3 caracteres"
   
   if username.startswith('-') or username.endswith('-'):
      return False, "No puede empezar ni terminar con guión"
   
   return True, ""


def validate_email_prefix(email_prefix):
   """
   Valida prefijo de correo.
   
   Returns:
      tuple: (is_valid: bool, error_msg: str)
   """
   if not email_prefix:
      return False, "El correo no puede estar vacío"
   
   # Solo letras minúsculas, números, puntos y guiones
   if not re.match(r'^[a-z0-9\.\-]+$', email_prefix):
      return False, "Solo letras minúsculas, números, puntos y guiones"
   
   if len(email_prefix) < 3:
      return False, "Mínimo 3 caracteres"
   
   if email_prefix.startswith('.') or email_prefix.endswith('.'):
      return False, "No puede empezar ni terminar con punto"
   
   if '..' in email_prefix:
      return False, "No puede tener puntos consecutivos"
   
   return True, ""


def validate_password(password):
   """
   Valida que la contraseña cumpla con requisitos mínimos.
   
   Returns:
      tuple: (is_valid: bool, error_msg: str)
   """
   if not password:
      return False, "La contraseña no puede estar vacía"
   
   if len(password) < 8:
      return False, "Mínimo 8 caracteres"
   
   if not any(c.isupper() for c in password):
      return False, "Debe contener al menos una mayúscula"
   
   if not any(c.islower() for c in password):
      return False, "Debe contener al menos una minúscula"
   
   if not any(c.isdigit() for c in password):
      return False, "Debe contener al menos un número"
   
   if not any(c in '!@#$%^&*()_+-=[]{}|;:,.<>?' for c in password):
      return False, "Debe contener al menos un carácter especial"
   
   return True, ""


def generate_secure_password(length=12):
   """
   Genera una contraseña segura aleatoria.
   
   Args:
      length: Longitud de la contraseña
      
   Returns:
      str: Contraseña generada
   """
   # Asegurar que tenga de cada tipo
   password = [
      random.choice(string.ascii_uppercase),
      random.choice(string.ascii_lowercase),
      random.choice(string.digits),
      random.choice('!@#$%^&*()_+-=')
   ]
   
   # Completar con caracteres aleatorios
   all_chars = string.ascii_letters + string.digits + '!@#$%^&*()_+-='
   password += [random.choice(all_chars) for _ in range(length - 4)]
   
   # Mezclar
   random.shuffle(password)
   
   return ''.join(password)


def open_pdf(path):
   """Abre un PDF con el visor predeterminado del sistema."""
   sistema = platform.system()
   try:
      if sistema == "Windows":
         os.startfile(path)
      elif sistema == "Darwin":  # macOS
         os.system(f"open '{path}'")
      else:  # Linux
         os.system(f"xdg-open '{path}'")
      return True
   except:
      return False


# ============================================================================
# CLASE PRINCIPAL
# ============================================================================

class CredencialesApp(ctk.CTk):
   def __init__(self):
      super().__init__()
      
      # Configuración
      self.title(f"{APP_TITLE} {APP_VERSION}")
      self.geometry(APP_SIZE)
      self.resizable(False, False)
      
      # Variables
      self.password_visible = False
      self.is_generating = False
      
      # Construir interfaz
      self.build_ui()
   
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
         text="🔐 " + APP_TITLE,
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
      
      # === FORMULARIO ===
      form_frame = ctk.CTkFrame(main_frame, fg_color=COLOR_BG_LIGHT, corner_radius=8)
      form_frame.pack(fill="x", pady=(0, 10))
      
      form_title = ctk.CTkLabel(
         form_frame,
         text="📝 Información de las Credenciales",
         font=FONT_LABEL,
         text_color=COLOR_TEXT
      )
      form_title.pack(pady=(15, 15), anchor="w", padx=15)
      
      # Correo corporativo
      ctk.CTkLabel(
         form_frame,
         text="Correo corporativo (sin @spradling.group):",
         font=FONT_SUBTITLE,
         text_color=COLOR_TEXT
      ).pack(pady=(0, 5), padx=15, anchor="w")
      
      correo_frame = ctk.CTkFrame(form_frame, fg_color="transparent")
      correo_frame.pack(pady=(0, 5), padx=15, anchor="w")
      
      self.correo_entry = ctk.CTkEntry(
         correo_frame,
         placeholder_text="Ej: juan.perez",
         width=300,
         height=35,
         font=FONT_SUBTITLE
      )
      self.correo_entry.pack(side="left")
      self.correo_entry.bind("<KeyRelease>", lambda e: self.validate_form())
      
      ctk.CTkLabel(
         correo_frame,
         text=DOMAIN,
         font=("Segoe UI", 11, "bold"),
         text_color=COLOR_PRIMARY
      ).pack(side="left", padx=(5, 0))
      
      self.correo_validation = ctk.CTkLabel(
         form_frame,
         text="",
         font=("Segoe UI", 9),
         text_color=COLOR_WARNING
      )
      self.correo_validation.pack(pady=(0, 10), padx=15, anchor="w")
      
      # Contraseña
      ctk.CTkLabel(
         form_frame,
         text="Contraseña de acceso:",
         font=FONT_SUBTITLE,
         text_color=COLOR_TEXT
      ).pack(pady=(0, 5), padx=15, anchor="w")
      
      pass_frame = ctk.CTkFrame(form_frame, fg_color="transparent")
      pass_frame.pack(pady=(0, 5), padx=15, anchor="w")
      
      self.pass_entry = ctk.CTkEntry(
         pass_frame,
         placeholder_text="Mínimo 8 caracteres",
         width=300,
         height=35,
         font=FONT_SUBTITLE,
         show="●"
      )
      self.pass_entry.pack(side="left")
      self.pass_entry.bind("<KeyRelease>", lambda e: self.validate_form())
      
      self.toggle_pass_btn = ctk.CTkButton(
         pass_frame,
         text="👁",
         command=self.toggle_password_visibility,
         width=35,
         height=35,
         font=("Segoe UI", 14)
      )
      self.toggle_pass_btn.pack(side="left", padx=(5, 0))
      
      self.generate_pass_btn = ctk.CTkButton(
         pass_frame,
         text="🎲",
         command=self.generate_password,
         width=35,
         height=35,
         font=("Segoe UI", 14),
         fg_color=COLOR_PRIMARY,
         hover_color="#43a047"
      )
      self.generate_pass_btn.pack(side="left", padx=(5, 0))
      
      self.pass_validation = ctk.CTkLabel(
         form_frame,
         text="",
         font=("Segoe UI", 9),
         text_color=COLOR_WARNING
      )
      self.pass_validation.pack(pady=(0, 10), padx=15, anchor="w")
      
      # Usuario de Windows/VPN
      ctk.CTkLabel(
         form_frame,
         text="Usuario de Windows:",
         font=FONT_SUBTITLE,
         text_color=COLOR_TEXT
      ).pack(pady=(0, 5), padx=15, anchor="w")
      
      self.user_entry = ctk.CTkEntry(
         form_frame,
         placeholder_text="Ej: perez-juan",
         width=300,
         height=35,
         font=FONT_SUBTITLE
      )
      self.user_entry.pack(pady=(0, 5), padx=15, anchor="w")
      self.user_entry.bind("<KeyRelease>", lambda e: self.validate_form())
      self.user_entry.bind("<Return>", lambda e: self.generate_pdf())
      
      self.user_validation = ctk.CTkLabel(
         form_frame,
         text="",
         font=("Segoe UI", 9),
         text_color=COLOR_WARNING
      )
      self.user_validation.pack(pady=(0, 15), padx=15, anchor="w")
      
      # === VISTA PREVIA ===
      preview_frame = ctk.CTkFrame(main_frame, fg_color=COLOR_BG_LIGHT, corner_radius=8)
      preview_frame.pack(fill="x", pady=(0, 10))
      
      preview_title = ctk.CTkLabel(
         preview_frame,
         text="👁 Vista Previa",
         font=FONT_LABEL,
         text_color=COLOR_TEXT
      )
      preview_title.pack(pady=(10, 10), anchor="w", padx=15)
      
      self.preview_text = ctk.CTkTextbox(
         preview_frame,
         width=590,
         height=200,
         font=("Consolas", 10),
         fg_color="#1e1e1e",
         text_color=COLOR_SUCCESS,
         border_width=1,
         border_color="#424242"
      )
      self.preview_text.pack(pady=(0, 15), padx=15)
      self.preview_text.insert("end", "Complete los campos para ver la vista previa...\n")
      
      # === VALIDACIÓN GENERAL ===
      self.validation_general = ctk.CTkLabel(
         main_frame,
         text="",
         font=("Segoe UI", 10, "bold"),
         text_color=COLOR_WARNING
      )
      self.validation_general.pack(pady=(0, 10))
      
      # === BOTONES ===
      button_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
      button_frame.pack(fill="x")
      
      self.generate_button = ctk.CTkButton(
         button_frame,
         text="▶ Generar PDF",
         command=self.generate_pdf,
         font=FONT_BUTTON,
         height=40,
         corner_radius=8,
         fg_color=COLOR_PRIMARY,
         hover_color="#43a047",
         state="disabled"
      )
      self.generate_button.pack(side="left", expand=True, fill="x", padx=(0, 5))
      
      self.clear_button = ctk.CTkButton(
         button_frame,
         text="🗑 Limpiar",
         command=self.clear_fields,
         font=FONT_BUTTON,
         height=40,
         corner_radius=8,
         fg_color=COLOR_BG_LIGHT,
         hover_color="#424242"
      )
      self.clear_button.pack(side="left", expand=True, fill="x", padx=(5, 0))
   
   def toggle_password_visibility(self):
      """Alterna visibilidad de la contraseña."""
      if self.password_visible:
         self.pass_entry.configure(show="●")
         self.password_visible = False
      else:
         self.pass_entry.configure(show="")
         self.password_visible = True
   
   def generate_password(self):
      """Genera una contraseña segura."""
      password = generate_secure_password(12)
      self.pass_entry.delete(0, "end")
      self.pass_entry.insert(0, password)
      self.validate_form()
      messagebox.showinfo(
         "Contraseña Generada",
         f"Se ha generado una contraseña segura.\n\n"
         f"⚠ Anótela antes de continuar:\n{password}"
      )
   
   def clear_fields(self):
      """Limpia todos los campos."""
      self.correo_entry.delete(0, "end")
      self.pass_entry.delete(0, "end")
      self.user_entry.delete(0, "end")
      self.correo_validation.configure(text="")
      self.pass_validation.configure(text="")
      self.user_validation.configure(text="")
      self.validation_general.configure(text="")
      self.preview_text.configure(state="normal")
      self.preview_text.delete("1.0", "end")
      self.preview_text.insert("end", "Complete los campos para ver la vista previa...\n")
      self.preview_text.configure(state="disabled")
   
   def validate_form(self):
      """Valida el formulario completo."""
      correo = self.correo_entry.get().strip()
      password = self.pass_entry.get().strip()
      usuario = self.user_entry.get().strip()
      
      all_valid = True
      
      # Validar correo
      if correo:
         is_valid, msg = validate_email_prefix(correo)
         if is_valid:
               self.correo_validation.configure(text="✓ Válido", text_color=COLOR_SUCCESS)
         else:
               self.correo_validation.configure(text=f"✗ {msg}", text_color=COLOR_ERROR)
               all_valid = False
      else:
         self.correo_validation.configure(text="")
         all_valid = False
      
      # Validar contraseña
      if password:
         is_valid, msg = validate_password(password)
         if is_valid:
               self.pass_validation.configure(text="✓ Contraseña segura", text_color=COLOR_SUCCESS)
         else:
               self.pass_validation.configure(text=f"✗ {msg}", text_color=COLOR_ERROR)
               all_valid = False
      else:
         self.pass_validation.configure(text="")
         all_valid = False
      
      # Validar usuario
      if usuario:
         is_valid, msg = validate_username(usuario)
         if is_valid:
               self.user_validation.configure(text="✓ Válido", text_color=COLOR_SUCCESS)
         else:
               self.user_validation.configure(text=f"✗ {msg}", text_color=COLOR_ERROR)
               all_valid = False
      else:
         self.user_validation.configure(text="")
         all_valid = False
      
      # Actualizar estado del botón
      if all_valid:
         self.generate_button.configure(state="normal")
         self.validation_general.configure(
               text="✓ Listo para generar PDF",
               text_color=COLOR_SUCCESS
         )
         self.update_preview(correo, password, usuario)
      else:
         self.generate_button.configure(state="disabled")
         self.validation_general.configure(text="")
   
   def update_preview(self, correo, password, usuario):
      """Actualiza la vista previa."""
      preview = f"""
╔════════════════════════════════════════════════════════╗
║     CREDENCIALES DE ACCESO CORPORATIVO PQN             ║
╚════════════════════════════════════════════════════════╝

📧 CORREO CORPORATIVO
Correo:     {correo}{DOMAIN}
Contraseña: {password}

📧 ACCESO A SISTEMAS
Usuario:    {usuario}
Contraseña: (La misma de arriba)

🌐 PLATAFORMAS DISPONIBLES
• S.O Windows
• FortiClient VPN
• Citrix
• Daruma
• Terranova
• Intranet PQN

⚠ IMPORTANTE
Guarde estas credenciales en un lugar seguro.
No comparta su contraseña con terceros.
      """
      
      self.preview_text.configure(state="normal")
      self.preview_text.delete("1.0", "end")
      self.preview_text.insert("end", preview.strip())
      self.preview_text.configure(state="disabled")
   
   def generate_pdf(self):
      """Genera el PDF de credenciales."""
      if self.is_generating:
         return
      
      correo = self.correo_entry.get().strip()
      password = self.pass_entry.get().strip()
      usuario = self.user_entry.get().strip()
      
      if not all([correo, password, usuario]):
         messagebox.showwarning("Campos Requeridos", "Complete todos los campos.")
         return
      
      self.is_generating = True
      self.generate_button.configure(state="disabled", text="⏳ Generando...")
      
      try:
         filename = "Tus Credenciales de Acceso PQN.pdf"
         
         # Rutas
         documentos = Path.home() / "Documents"
         if not documentos.exists():
               documentos = Path.home() / "Documentos"
         
         ruta_docs = documentos / filename
         ruta_datos = Path("D:/Datos") / filename
         
         # Crear directorios
         documentos.mkdir(parents=True, exist_ok=True)
         if Path("D:/").exists():
               Path("D:/Datos").mkdir(parents=True, exist_ok=True)
         
         # Generar PDF
         self.crear_pdf(str(ruta_docs), correo, password, usuario)
         
         if Path("D:/").exists():
               self.crear_pdf(str(ruta_datos), correo, password, usuario)
         
         # Abrir PDF
         if Path("D:/").exists() and ruta_datos.exists():
               open_pdf(str(ruta_datos))
         else:
               open_pdf(str(ruta_docs))
         
         messagebox.showinfo(
               "Éxito",
               f"✓ PDF generado correctamente\n\n"
               f"Ubicación: {ruta_docs}\n\n"
               f"El documento se abrió automáticamente."
         )
         
         # Cerrar aplicación
         self.after(1000, self.quit)
         
      except Exception as e:
         messagebox.showerror("Error", f"No se pudo generar el PDF:\n\n{str(e)}")
      finally:
         self.is_generating = False
         self.generate_button.configure(state="normal", text="▶ Generar PDF")
   
   def crear_pdf(self, path, correo, password, usuario):
      """Crea el PDF de credenciales."""
      doc = SimpleDocTemplate(
         str(path),
         pagesize=letter,
         rightMargin=50,
         leftMargin=50,
         topMargin=60,
         bottomMargin=40
      )
      
      story = []
      styles = getSampleStyleSheet()
      
      # Estilos personalizados
      title_style = ParagraphStyle(
         "title",
         fontName="Helvetica-Bold",
         fontSize=16,
         leading=20,
         alignment=TA_CENTER,
         spaceAfter=20,
         textColor=colors.HexColor("#4caf50")
      )
      
      normal_style = ParagraphStyle(
         "normal",
         fontName="Helvetica",
         fontSize=11,
         leading=15,
         alignment=TA_LEFT,
         spaceAfter=8
      )
      
      bold_style = ParagraphStyle(
         "bold",
         fontName="Helvetica-Bold",
         fontSize=11,
         leading=15,
         alignment=TA_LEFT,
         spaceAfter=8
      )
      
      # Título
      story.append(Paragraph("🔐 Credenciales de Acceso PQN", title_style))
      story.append(Spacer(1, 12))
      
      # Introducción
      story.append(Paragraph(
         "Estimado usuario, a continuación se presentan las credenciales de acceso "
         "que se le han asignado para su ingreso, cambio o transición de cargo en la empresa.",
         normal_style
      ))
      story.append(Spacer(1, 18))
      
      # Correo
      story.append(Paragraph("📧 Correo Corporativo", bold_style))
      story.append(Spacer(1, 6))
      story.append(Paragraph(f"Correo:     {correo}{DOMAIN}", normal_style))
      story.append(Paragraph(f"Contraseña: {password}", normal_style))
      story.append(Spacer(1, 12))
      story.append(Paragraph(
         "ℹ️ Todas las claves son complejas por política de seguridad interna.",
         normal_style
      ))
      story.append(Spacer(1, 18))
      
      # Tabla de plataformas
      story.append(Paragraph("🌐 Plataformas de Acceso", bold_style))
      story.append(Spacer(1, 10))
      
      def p(text):
         return Paragraph(text, normal_style)
      
      data = [
         [p("Plataforma"), p("URL / Ubicación"), p("Usuario"), p("Contraseña")],
         [p("Windows"), p("Sistema Operativo"), p(usuario), p("La misma")],
         [p("FortiClient VPN"), p("Programa instalado"), p(usuario), p("La misma")],
         [p("Citrix"), p("portal.proquinal.com/Citrix"), p(usuario), p("La misma")],
         [p("Daruma"), p("proquinal.darumasoftware.com"), p(usuario), p("La misma")],
         [p("Terranova"), p("secure.terranovasite.com"), p(f"{correo}{DOMAIN}"), p("La misma")],
         [p("Intranet"), p("pqnintranet.proquinal.com"), p("Nº de carnet"), p("Tu cédula")]
      ]
      
      table = Table(data, colWidths=[100, 180, 100, 80], repeatRows=1)
      table.setStyle(TableStyle([
         ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#4caf50")),
         ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
         ("ALIGN", (0, 0), (-1, -1), "LEFT"),
         ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
         ("FONTSIZE", (0, 0), (-1, -1), 9),
         ("BOX", (0, 0), (-1, -1), 1, colors.black),
         ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.grey),
         ("VALIGN", (0, 0), (-1, -1), "TOP"),
         ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f5f5f5")])
      ]))
      
      story.append(table)
      story.append(Spacer(1, 20))
      
      # Información importante
      story.append(Paragraph("⚠️ Importante", bold_style))
      story.append(Paragraph(
         "• Para solicitudes relacionadas con su equipo o accesos, cree un caso en Mayté "
         "(https://mayte.spradling.group) usando su correo y contraseña.",
         normal_style
      ))
      story.append(Spacer(1, 20))
      
      # Firma
      story.append(Paragraph("Atentamente,", normal_style))
      story.append(Spacer(1, 12))
      story.append(Paragraph("Equipo de Mesa de Servicio IT<br/>Proquinal S.A.", normal_style))
      story.append(Spacer(1, 20))
      
      # Pie
      footer_style = ParagraphStyle(
         "footer",
         fontName="Helvetica-Oblique",
         fontSize=8,
         alignment=TA_CENTER,
         textColor=colors.grey
      )
      story.append(Paragraph(
         f"Documento generado el {datetime.now().strftime('%d/%m/%Y %H:%M')} | "
         f"{APP_TITLE} {APP_VERSION}",
         footer_style
      ))
      
      doc.build(story)


# ============================================================================
# PUNTO DE ENTRADA
# ============================================================================

if __name__ == "__main__":
   app = CredencialesApp()
   app.mainloop()