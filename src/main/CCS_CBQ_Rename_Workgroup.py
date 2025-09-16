"""
   CCS_CBQ_Rename_Workgroup.py
   Autor: Josué Romero
   Empresa: Stefanini / PQN
   Fecha: 14/Agosto/2025

   Descripción:
   Renombra el equipo siguiendo el formato [U7DSERIAL]-[CCS/CBQ]-COL
   y cambia el grupo de trabajo al nombre de la ciudad.
"""

import ctypes
import sys
import subprocess
import threading
import customtkinter as ctk
from tkinter import messagebox

# Configurar tema
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

APP_TITLE = "Auto-Renombrador de Equipos CCS-CBQ"
APP_SIZE = "500x500"


def run_powershell(script):
   """Ejecuta un script de PowerShell inline"""
   command = ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script]
   process = subprocess.run(command, capture_output=True, text=True, shell=True)
   return process.stdout.strip(), process.stderr.strip(), process.returncode

class RenameApp(ctk.CTk):
   def __init__(self):
      super().__init__()
      self.title(APP_TITLE)
      self.geometry(APP_SIZE)
      self.resizable(False, False)
      self.build_ui()

   def build_ui(self):
      self.label_title = ctk.CTkLabel(self, text=APP_TITLE, font=ctk.CTkFont(size=24, weight="bold"))
      self.label_title.pack(pady=(20, 10))

      self.label_meta = ctk.CTkLabel(self, text="Autor: Josué Romero  |  Empresa: Stefanini / PQN  |  Fecha: 14/Agosto/2025", font=ctk.CTkFont(size=12))
      self.label_meta.pack(pady=(0, 20))

      self.entry_label = ctk.CTkLabel(self, text="Nombre de la cuidad:")
      self.entry_label.pack()
      self.workgroup_entry = ctk.CTkEntry(self, placeholder_text="Ej: Bogota D.C, Medellin")
      self.workgroup_entry.pack(pady=(5, 15))

      self.text_log = ctk.CTkTextbox(self, width=480, height=250)
      self.text_log.pack(pady=10)

      # Evento para consultar con Enter
      self.entry_label.bind("<Return>", lambda event: self.execute_process())

      self.btn_run = ctk.CTkButton(self, text="Ejecutar", command=self.on_execute)
      self.btn_run.pack(pady=10)

      self.label_status = ctk.CTkLabel(self, text="", font=ctk.CTkFont(size=14, weight="bold"))
      self.label_status.pack()

   def log(self, message):
      self.text_log.configure(state="normal")
      self.text_log.insert("end", message + "\n")
      self.text_log.see("end")
      self.text_log.configure(state="disabled")

   def on_execute(self):
      self.btn_run.configure(state="disabled")
      self.text_log.configure(state="normal")
      self.text_log.delete("1.0", "end")
      self.text_log.configure(state="disabled")
      self.label_status.configure(text="Ejecutando...", text_color="yellow")
      threading.Thread(target=self.execute_process).start()

   def execute_process(self):
      try:
         workgroup = self.workgroup_entry.get().strip()
         if not workgroup:
            messagebox.showerror("Error", "Debe ingresar un nombre de grupo de trabajo.")
            self.finish_fail()
            return

         self.log("[*] Obteniendo información del sistema...")
         serial, _, _ = run_powershell("(Get-WmiObject win32_bios).SerialNumber.Trim()")
         manufacturer, _, _ = run_powershell("(Get-WmiObject win32_computersystem).Manufacturer.Trim()")
         current_name, _, _ = run_powershell("$env:COMPUTERNAME")
         last7 = serial[-7:].upper()
         new_name = f"{last7}-CCS-COL"

         self.log(f"[+] Fabricante detectado: {manufacturer}")
         self.log(f"[+] Serial detectado: {serial}")
         self.log(f"[+] Nombre actual del equipo: {current_name}")
         self.log(f"[+] Nombre correcto esperado: {new_name}")
         self.log(f"[+] Grupo de trabajo solicitado: {workgroup}")

         ps_script = f'''
            try {{
               Rename-Computer -NewName "{new_name}" -WorkGroupName "{workgroup}" -Force -PassThru -ErrorAction Stop
               Write-Output "[✓] Cambios aplicados correctamente. Reiniciando equipo..."
               Restart-Computer
            }} catch {{
               Write-Output "[X] Error al aplicar los cambios: $($_.Exception.Message)"
            }}
         '''

         out, err, code = run_powershell(ps_script)
         self.log(out if out else "[X] Error sin salida")
         if "[✓]" in out:
               self.finish_success()
         else:
               self.finish_fail()

      except Exception as e:
         self.log(f"[X] Excepción: {e}")
         self.finish_fail()

   def finish_success(self):
      self.label_status.configure(text="Cambios aplicados correctamente.", text_color="green")
      self.btn_run.configure(state="normal")

   def finish_fail(self):
      self.label_status.configure(text="Falló la operación.", text_color="red")
      self.btn_run.configure(state="normal")

if __name__ == "__main__":
   app = RenameApp()
   app.mainloop()
