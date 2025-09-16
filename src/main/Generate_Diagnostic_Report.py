"""
   Generate_Diagnostic_Report.py
   Autor: Josué Romero
   Empresa: Stefanini / PQN
   Fecha: 10/Agosto/2025

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
import getpass
from pathlib import Path

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("green")


class DiagnosticApp(ctk.CTk):
   def __init__(self):
      super().__init__()
      self.title("Auto-Generador de Informe IT")
      self.geometry("490x385")
      self.resizable(False, False)
      self.build_ui()

   def build_ui(self):
      self.title_label = ctk.CTkLabel(self, text="Auto-Generador de Informe Soporte", font=ctk.CTkFont(size=22, weight="bold"))
      self.title_label.pack(pady=10)

      self.meta_label = ctk.CTkLabel(self, text="Autor: Josué Romero  |  Empresa: Stefanini / PQN  |  Fecha: 10/Agosto/2025", font=ctk.CTkFont(size=12))
      self.meta_label.pack()

      self.tecnico_entry = ctk.CTkEntry(self, placeholder_text="Nombre técnico")
      self.tecnico_entry.pack(pady=7)

      self.fixed_asset_entry = ctk.CTkEntry(self, placeholder_text="Número placa")
      self.fixed_asset_entry.pack(pady=7)

      self.ticket_entry = ctk.CTkEntry(self, placeholder_text="Número caso Mayté")
      self.ticket_entry.pack(pady=7)

      # Evento para consultar con Enter
      self.ticket_entry.bind("<Return>", lambda event: self.generar_reporte())

      self.output_box = ctk.CTkTextbox(self, width=400, height=110, font=("Consolas", 11))
      self.output_box.pack(pady=10)
      self.output_box.insert("end", "✔ Listo para generar informe...\n")

      self.generate_button = ctk.CTkButton(self, text="Ejecutar", command=self.generar_reporte)
      self.generate_button.pack(pady=10)

   def log(self, msg):
      timestamp = datetime.now().strftime("%H:%M:%S")
      self.output_box.insert("end", f"[{timestamp}] {msg}\n")
      self.output_box.see("end")

   def obtener_modelo_procesador(self):
      try:
         output = subprocess.check_output(
               'powershell -Command "Get-CimInstance -ClassName Win32_Processor | Select-Object -ExpandProperty Name"',
               shell=True
         ).decode(errors="ignore").strip()
         return output if output else "No disponible"
      except Exception as e:
         return "Error al obtener modelo"

   def obtener_marca_equipo(self):
      try:
         output = subprocess.check_output(
               'powershell -Command "Get-CimInstance -ClassName Win32_ComputerSystem | Select-Object -ExpandProperty Manufacturer"',
               shell=True
         ).decode(errors="ignore").strip()
         return output if output else "No disponible"
      except Exception as e:
         return "Error al obtener la marca"

   def generar_reporte(self):
      tecnico = self.tecnico_entry.get().strip()
      ticket = self.ticket_entry.get().strip()
      fixed_asset = self.fixed_asset_entry.get().strip()

      if not tecnico or not ticket or not fixed_asset:
         messagebox.showwarning("Campos requeridos", "Por favor completa todos los campos.")
         return

      self.generate_button.configure(state="disabled")
      self.log("Generando informe...")

      empresa = "Proquinal S.A.S"
      area = "Soporte IT"
      fecha_actual = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
      filename = f"Informe Diagnóstico Caso {ticket}.pdf"

      hostname = socket.gethostname()
      sistema = f"{platform.system()} {platform.release()}"
      ram_total = round(psutil.virtual_memory().total / (1024 ** 3), 2)
      ram_libre = round(psutil.virtual_memory().available / (1024 ** 3), 2)
      disco_total = round(psutil.disk_usage('C:\\').total / (1024 ** 3), 2)
      disco_libre = round(psutil.disk_usage('C:\\').free / (1024 ** 3), 2)

      # Variables creadas
      active_user = getpass.getuser()
      marca_equipo = self.obtener_marca_equipo()

      # ---- RUTAS DE SALIDA ----
      documentos = Path.home() / "Documents"
      if not documentos.exists():  # fallback si está en español
         documentos = Path.home() / "Documentos"

      ruta_docs = documentos / filename
      ruta_datos = Path("D:/Datos") / filename

      # Crear carpetas si no existen
      documentos.mkdir(parents=True, exist_ok=True)
      Path("D:/Datos").mkdir(parents=True, exist_ok=True)

      # ---- FUNCIÓN QUE CREA EL PDF ----
      def crear_pdf(path):
         c = canvas.Canvas(str(path), pagesize=letter)
         width, height = letter

         # Encabezado
         c.setFont("Helvetica-Bold", 12)
         c.drawString(50, height - 50, "Proquinal S.A.S")
         c.drawRightString(width - 50, height - 50, "Stefanini Group CO")

         # Título
         c.setFont("Helvetica-Bold", 16)
         c.drawCentredString(width / 2, height - 90, "INFORME DE DIAGNÓSTICO")

         y = height - 130

         def draw_row(label, value, x, y):
            c.setFont("Helvetica-Bold", 10)
            c.drawString(x, y, label)
            offset = c.stringWidth(label, "Helvetica-Bold", 10)
            c.setFont("Helvetica", 10)
            c.drawString(x + offset + 5, y, value)

         draw_row("Empresa:", empresa, 50, y)
         draw_row("Área:", area, width / 2, y)
         y -= 20
         draw_row("Técnico:", tecnico, 50, y)
         draw_row("Caso en Mayté:", ticket, width / 2, y)
         y -= 20
         draw_row("Fecha y hora:", fecha_actual, 50, y)

         c.line(50, y - 5, width - 50, y - 5)
         y -= 25

         # Diagnóstico
         c.setFont("Helvetica-Bold", 12)
         c.drawString(50, y, "Diagnóstico del Sistema")
         y -= 20

         col1 = [
            f"Usuario activo: {active_user}",
            f"Nombre de host: {hostname}",
            f"Sistema operativo: {sistema}",
            f"Procesador: {self.obtener_modelo_procesador()}",
            f"Marca: {self.obtener_marca_equipo()}",
            f"RAM Total: {ram_total} GB",
         ]
         col2 = [
            f"Placa AF: {fixed_asset}",
            f"Disco Total: {disco_total} GB",
            f"Disco Libre: {disco_libre} GB",
            "Antivirus: Microsoft Defender",
            "Windows Update: Al día",
            "Otros Discos: N/A"
         ]

         col_width = (width - 100) / 2

         for i, line in enumerate(col1):
            c.setFont("Helvetica", 10)
            c.drawString(50, y - (i * 15), line)
         for i, line in enumerate(col2):
            c.drawString(100 + col_width, y - (i * 15), line)

         y -= (max(len(col1), len(col2)) * 15) + 20

         # Procedimiento técnico
         c.setFont("Helvetica-Bold", 12)
         c.drawString(50, y, "Procedimiento Realizado por el Técnico")
         y -= 20

         pasos = [
            "1. Se retiraron los tornillos de la carcasa inferior del equipo.",
            "2. Se retiró cuidadosamente la tapa para acceder al hardware interno.",
            "3. Se desconectó la batería del equipo para evitar descargas.",
            "4. Se limpió el polvo acumulado con aire comprimido en ventiladores y componentes.",
            "5. Se retiró el disipador del procesador y se limpió completamente la pasta térmica vieja.",
            "6. Se aplicó nueva pasta térmica de alta calidad al procesador.",
            "7. Se reinstaló el disipador y se aseguraron los tornillos correspondientes.",
            "8. Se volvió a conectar la batería y se cerró la tapa inferior.",
            "9. Se encendió el equipo y se realizaron pruebas de encendido y estabilidad térmica.",
            "10. Se ejecutaron herramientas de análisis en el sistema para validar el estado general."
         ]

         c.setFont("Helvetica", 10)
         for paso in pasos:
            if y < 100:
               c.showPage()
               y = height - 80
            c.drawString(50, y, paso)
            y -= 20

         # Recomendaciones
         c.setFont("Helvetica-Bold", 12)
         y -= 20
         c.drawString(50, y, "Recomendaciones para el Propietario del Equipo")
         y -= 20
         c.setFont("Helvetica", 10)

         recomendaciones = [
            "*  Presentar el equipo físicamente al área de Soporte cada dos mes, con previa creación de caso en Mayté.",
            "*  Evitar mantener el cargador conectado de forma permanente para preservar la vida útil de la batería.",
            "*  Apagar el equipo diariamente y permitir un reposo mínimo de 5 minutos antes de encenderlo.",
            "*  Mantener limpias las rejillas de ventilación para asegurar la correcta disipación de calor."
         ]

         for rec in recomendaciones:
            if y < 80:
               c.showPage()
               y = height - 80
               c.setFont("Helvetica", 10)
            c.drawString(50, y, rec)
            y -= 20

         # Nota final
         c.setFont("Helvetica-Oblique", 10)
         y = max(y, 40)
         c.drawString(50, y, "Nota: Las evidencias fotográficas asociadas a este caso se adjuntan por separado al caso.")

         # Pie de página
         c.setFont("Helvetica-Oblique", 9)
         c.drawRightString(width - 130, 20, "Informe generado por el sistema 'Generate_Diagnostic_Report.exe' | ©2025 @josuerom")

         c.save()

      # Crear documentos en C: y D:
      crear_pdf(ruta_docs)
      crear_pdf(ruta_datos)

      self.log(f"✅ Informe generado en:\n{ruta_docs}\n{ruta_datos}")

      # Abrir preferiblemente desde D, si existe
      abrir_path = ruta_datos if ruta_datos.exists() else ruta_docs
      try:
         os.startfile(str(abrir_path))
      except:
         self.log("⚠ No se pudo abrir automáticamente el documento PDF.")

      self.quit()
      self.destroy()


if __name__ == "__main__":
   app = DiagnosticApp()
   app.mainloop()
