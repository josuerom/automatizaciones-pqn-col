"""
   CCS_CBQ_Register_AutoPilot.py
   Autor: Josué Romero
   Empresa: Stefanini / PQN
   Fecha: 14/Agosto/2025

   Descripción:
   Instala el complemento y sube el dispositivo a la aplicación empresarial AutoPilot.
"""

import subprocess
import customtkinter as ctk
import datetime
from tkinter import messagebox

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("green")

APP_TITLE = "Registro de Equipos CCS-CBQ AutoPilot"
APP_SIZE = "500x460"


def get_serial():
   """Obtiene el número de serie del equipo usando PowerShell (sin wmi/wmic)."""
   try:
      result = subprocess.run(
         ["powershell", "-Command", "(Get-CimInstance Win32_BIOS).SerialNumber.Trim()"],
         capture_output=True,
         text=True,
         check=True
      )
      return result.stdout.strip()
   except Exception:
      return "N/A"


class AutoPilotApp(ctk.CTk):
   def __init__(self):
      super().__init__()
      self.title(APP_TITLE)
      self.geometry(APP_SIZE)
      self.resizable(False, False)
      self.build_ui()

   def build_ui(self):
      self.title_label = ctk.CTkLabel(self, text=APP_TITLE, font=ctk.CTkFont(size=22, weight="bold"))
      self.title_label.pack(pady=10)

      self.meta_label = ctk.CTkLabel(
         self,
         text="Autor: Josué Romero  |  Empresa: Stefanini / PQN  |  Fecha: 14/Agosto/2025",
         font=ctk.CTkFont(size=12)
      )
      self.meta_label.pack()

      self.output_box = ctk.CTkTextbox(self, width=470, height=300, font=("Consolas", 12))
      self.output_box.pack(pady=15)
      self.output_box.insert("end", "Listo para iniciar.\n")

      self.run_button = ctk.CTkButton(self, text="Ejecutar", command=self.run_autopilot)
      self.run_button.pack(pady=10)

      # Permitir que Enter ejecute el proceso
      self.bind("<Return>", lambda event: self.run_autopilot())

   def log(self, msg):
      timestamp = datetime.datetime.now().strftime("%H:%M:%S")
      self.output_box.insert("end", f"[{timestamp}] {msg}\n")
      self.output_box.see("end")

   def run_powershell(self, command):
      """Ejecuta un comando de PowerShell y muestra salida en tiempo real en el output_box."""
      process = subprocess.Popen(
         ["powershell", "-Command", command],
         stdout=subprocess.PIPE,
         stderr=subprocess.PIPE,
         text=True
      )
      # Leer línea por línea
      for line in process.stdout:
         self.log(line.strip())
      for line in process.stderr:
         self.log(f"[ERROR] {line.strip()}")

      process.wait()
      return process.returncode

   def run_autopilot(self):
      self.run_button.configure(state="disabled")

      serial = get_serial()
      self.log(f"Serial del equipo: {serial}")
      self.log("Abriendo PowerShell con ExecutionPolicy Bypass...")

      # Creamos un script temporal con los comandos necesarios
      ps_script = r"""
      Install-Script -Name Get-WindowsAutoPilotInfo -Force -Scope CurrentUser -Confirm:$false
      Get-WindowsAutoPilotInfo -Online
      """

      script_path = "C:\\Temp\\autopilot_register.ps1"

      # Guardar script en disco
      try:
         with open(script_path, "w", encoding="utf-8") as f:
            f.write(ps_script)
      except Exception as e:
         self.log(f"[X] Error creando script: {e}")
         messagebox.showerror("Error", "No se pudo crear el script en disco.")
         self.run_button.configure(state="normal")
         return

      # Ejecutar en ventana nueva con bypass y esperando confirmación automática
      try:
         subprocess.Popen(
            [
               "powershell",
               "-ExecutionPolicy", "Bypass",
               "-NoExit",
               "-File", script_path
            ]
         )
         self.log("✓ Ventana de PowerShell abierta. Sigue el proceso en esa ventana.")
         messagebox.showinfo(
            "Proceso iniciado",
            "Se abrió una ventana de PowerShell.\n\n"
            "⚠️ IMPORTANTE: Cuando te pida confirmación, ingresa 'Y' o 'S' según tu idioma.\n"
            "Luego se ejecutará automáticamente la subida a AutoPilot."
         )
      except Exception as e:
         self.log(f"[X] Error ejecutando PowerShell: {e}")
         messagebox.showerror("Error", "No se pudo abrir PowerShell.")
         self.run_button.configure(state="normal")


if __name__ == "__main__":
   app = AutoPilotApp()
   app.mainloop()
