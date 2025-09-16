"""
   PQN_Access_Credentials.py
   Autor: Josué Romero
   Empresa: Stefanini / PQN
   Fecha: 28/Agosto/2025

   Descripción:
   Generador automático de credenciales corporativos de acceso
"""

import os
import platform
import customtkinter as ctk
from tkinter import messagebox
from reportlab.lib.pagesizes import legal
from reportlab.platypus import Table, TableStyle, Paragraph, Spacer, SimpleDocTemplate
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from pathlib import Path


# Configuración de apariencia
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("green")


class CredencialesApp(ctk.CTk):
   def __init__(self):
      super().__init__()
      self.title("Auto Generador de Credenciales PQN")
      self.geometry("420x250")
      self.resizable(False, False)
      self.build_ui()

   def build_ui(self):
      self.title_label = ctk.CTkLabel(self, text="Auto Generador de Credenciales PQN",
                                       font=ctk.CTkFont(size=22, weight="bold"))
      self.title_label.pack(pady=10)

      self.correo_entry = ctk.CTkEntry(self, placeholder_text="josue.romero")
      self.correo_entry.pack(pady=7)

      self.pass_entry = ctk.CTkEntry(self, placeholder_text="contraseña", show="*")
      self.pass_entry.pack(pady=7)

      self.user_entry = ctk.CTkEntry(self, placeholder_text="romero-josue")
      self.user_entry.pack(pady=7)

      self.user_entry.bind("<Return>", lambda event: self.generate_pdf())
      self.generate_button = ctk.CTkButton(self, text="Ejecutar", command=self.generate_pdf)
      self.generate_button.pack(pady=15)

   def generate_pdf(self):
      correo = self.correo_entry.get().strip()
      password = self.pass_entry.get().strip()
      usuario = self.user_entry.get().strip()

      if not correo or not password or not usuario:
         messagebox.showwarning("Campos requeridos", "Por favor completa todos los campos.")
         return

      filename = "Tus Credenciales de Acceso PQN.pdf"
      documentos = Path.home() / "Documents"
      if not documentos.exists():
         documentos = Path.home() / "Documentos"
      ruta_docs = documentos / filename
      ruta_datos = Path("D:/Datos") / filename

      documentos.mkdir(parents=True, exist_ok=True)
      if Path("D:/").exists():
         Path("D:/Datos").mkdir(parents=True, exist_ok=True)

      # Estilos personalizados
      styles = getSampleStyleSheet()
      normal = ParagraphStyle("normal", fontName="Helvetica", fontSize=11, leading=15, alignment=TA_LEFT, spaceAfter=8)
      bold = ParagraphStyle("bold", fontName="Helvetica-Bold", fontSize=11, leading=15, alignment=TA_LEFT, spaceAfter=8)
      title = ParagraphStyle("title", fontName="Helvetica-Bold", fontSize=14, leading=18, alignment=TA_CENTER, spaceAfter=20)

      def abrir_pdf(path):
         sistema = platform.system()
         try:
            if sistema == "Windows":
               os.startfile(path)
            elif sistema == "Darwin":  # macOS
               os.system(f"open '{path}'")
            else:  # Linux
               os.system(f"xdg-open '{path}'")
         except Exception as e:
            print(f"No se pudo abrir el PDF automáticamente: {e}")

      def crear_pdf(path):
         doc = SimpleDocTemplate(str(path), pagesize=legal,
                                 rightMargin=50, leftMargin=50,
                                 topMargin=60, bottomMargin=40)
         story = []

         # Encabezado
         story.append(Paragraph("Credenciales de Acceso PQN", title))

         # Texto inicial
         story.append(Paragraph(
            "Estimado usuario. A continuación, le comparto las credenciales de acceso que se te han asignado "
            "por tu ingreso, cambio o transición de cargo en la empresa.", normal
         ))

         # Correo corporativo
         story.append(Spacer(1, 12))
         story.append(Paragraph("Correo Corporativo", bold))
         story.append(Spacer(1, 6))
         story.append(Paragraph(f"Correo:     {correo}@spradling.group", normal))
         story.append(Paragraph(f"Contraseña: {password}", normal))
         story.append(Spacer(1, 12))
         story.append(Paragraph(
            "“Sabemos que es difícil, pero todas las claves son así de complejas porque es una política interna.”", normal
         ))

         # Plataformas
         story.append(Spacer(1, 18))
         story.append(Paragraph("Plataformas de Acceso", bold))
         story.append(Spacer(1, 10))

         # Tabla con texto ajustado
         def p(text): return Paragraph(text, normal)

         data = [
            [p("Plataforma"), p("Página web"), p("Importancia"), p("Usuario de acceso"), p("Contraseña de acceso")],
            [p("S.O Windows"), p("No es página web"), p("Para iniciar sesión en el portátil"), p(usuario), p("La misma de arriba")],
            [p("FortiClient VPN"), p("No es página web – es un programa que ya está instalado en su portátil"),
             p("Para trabajar desde casa y tener acceso a las carpetas compartidas"), p(usuario), p("La misma de arriba")],
            [p("Citrix"), p("http://portal.proquinal.com/Citrix/PROQUINAL-INTERNOWeb/ – Instalado en el portátil"),
             p("Para acceso remoto a las aplicaciones PQN de tu área"), p(usuario), p("La misma de arriba")],
            [p("Daruma"), p("https://proquinal.darumasoftware.com/app.php/staff/"),
             p("Para ver todas las bases de datos y procesos PQN"), p(usuario), p("La misma de arriba")],
            [p("Terranova"), p("https://secure.terranovasite.com/Service"),
             p("Para realizar cursos cortos corporativos"), p("Tu correo"), p("La misma de arriba")],
            [p("Intranet"), p("http://pqnintranet.proquinal.com/PEP-PORTAL-WEB/appmanager/intranet/proquinal_es"),
             p("Para compartir información y recursos corporativos"), p("Tu número de carnet"), p("Tu primera contraseña es la cédula")]
         ]

         # Ajuste de tamaños para que quepa en la página letter (8.5 x 11 pulgadas)
         table = Table(data, colWidths=[70, 120, 130, 70, 80], repeatRows=1)

         table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
            ("ALIGN", (0, 0), (-1, -1), "LEFT"),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("BOX", (0, 0), (-1, -1), 1, colors.black),
            ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.black),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
         ]))
         story.append(table)

         # Texto final
         story.append(Spacer(1, 20))
         story.append(Paragraph("Importante", bold))
         story.append(Paragraph(
            "* Para cualquier solicitud relacionada con tu computador o accesos, debes crear un caso en Mayté, "
            "que es la plataforma de gestión de incidentes y requerimientos del área de IT.", normal
         ))
         story.append(Spacer(1, 10))
         story.append(Paragraph(
            '* Página web: <a href="https://mayte.spradling.group">https://mayte.spradling.group</a> '
            '(usa tu correo y contraseña para acceder) y diligencia el formulario de acuerdo con tu necesidad.',
            normal
         ))
         story.append(Spacer(1, 20))
         story.append(Paragraph("Atentamente,", normal))
         story.append(Spacer(1, 12))
         story.append(Paragraph("Equipo de la Mesa de Servicio IT<br/>Proquinal S.A.", normal))

         doc.build(story)

      crear_pdf(ruta_docs)
      if Path("D:/").exists():
         crear_pdf(ruta_datos)

      # Abrir el PDF generado automáticamente
      if Path("D:/").exists():
         abrir_pdf(ruta_datos)
      else:
         abrir_pdf(ruta_docs)

      # Cerrar la aplicación
      self.quit()
      self.destroy()


if __name__ == "__main__":
   app = CredencialesApp()
   app.mainloop()
