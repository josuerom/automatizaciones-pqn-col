"""
   Optimize_System_Performance.py
   Autor: Josué Romero
   Empresa: Stefanini / PQN
   Fecha: 30/Agosto/2025

   Descripción:
   Optimiza el rendimiento del equipo ejecutando varias herramientas claves
"""

import sys
import ctypes
import subprocess
import threading
import customtkinter as ctk
from tkinter import messagebox

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

APP_TITLE = "Optimizador de Sistema Operativo"
APP_SIZE = "510x535"


def run_command(command, shell=True):
   process = subprocess.run(command, capture_output=True, text=True, shell=shell)
   return process.stdout.strip(), process.stderr.strip(), process.returncode

class OptimizeApp(ctk.CTk):
   def __init__(self):
      super().__init__()
      self.title(APP_TITLE)
      self.geometry(APP_SIZE)
      self.resizable(False, False)
      self.build_ui()

   def build_ui(self):
      self.label_title = ctk.CTkLabel(self, text="Auto Optimizador de Windows", font=ctk.CTkFont(size=24, weight="bold"))
      self.label_title.pack(pady=(15,10))

      self.label_meta = ctk.CTkLabel(self, text="Autor: Josué Romero  |  Empresa: Stefanini / PQN  |  Fecha: 30/Agosto/2025", font=ctk.CTkFont(size=12))
      self.label_meta.pack(pady=(0,20))

      self.text_log = ctk.CTkTextbox(self, width=500, height=340)
      self.text_log.pack(padx=10, pady=10)

      self.btn_run = ctk.CTkButton(self, text="Ejecutar", command=self.on_start)
      self.btn_run.pack(pady=10)

      # Evento para consultar con Enter
      self.btn_run.bind("<Return>", lambda event: self.on_start())

      self.label_status = ctk.CTkLabel(self, text="", font=ctk.CTkFont(size=14, weight="bold"))
      self.label_status.pack()

   def log(self, msg):
      self.text_log.configure(state="normal")
      self.text_log.insert("end", msg + "\n")
      self.text_log.see("end")
      self.text_log.configure(state="disabled")

   def on_start(self):
      self.btn_run.configure(state="disabled")
      self.text_log.configure(state="normal")
      self.text_log.delete("1.0", "end")
      self.text_log.configure(state="disabled")
      self.label_status.configure(text="Ejecutando procesos...", text_color="yellow")
      threading.Thread(target=self.optimize_system).start()

   def optimize_system(self):
      try:
         self.log("[+] Ejecutando limpieza de disco para todos los usuarios (cleanmgr)...")
         run_command('cleanmgr /tuneup:1')
         self.log("[✓] Limpieza de disco completada.")

         self.log("[+] Optimizando discos C: y D: (defrag)...")
         self.run_and_log("defrag C: /O")
         self.run_and_log("defrag D: /O")

         self.log("[+] Eliminando archivos temporales...")
         self.run_and_log('powershell Remove-Item -Path "$env:TEMP\\*" -Recurse -Force -ErrorAction SilentlyContinue')
         self.run_and_log('powershell Remove-Item -Path "C:\\Windows\\Temp\\*" -Recurse -Force -ErrorAction SilentlyContinue')
         self.run_and_log('powershell Remove-Item "C:\\Windows\\Prefetch\\*" -Force -ErrorAction SilentlyContinue')
         self.run_and_log('powershell Clear-RecycleBin -Force')

         self.log("[+] Reparando archivos del S.O (sfc /scannow). Esto puede tardar varios minutos...")
         self.run_and_log("sfc /scannow")

         self.log("[+] Reparando imagen del sistema con DISM. Esto puede tardar varios minutos...")
         self.run_and_log("DISM /Online /Cleanup-Image /ScanHealth")
         self.run_and_log("DISM /Online /Cleanup-Image /RestoreHealth")

         self.log("[+] Actualizando todos los programas instalados excepto Java SE 8 341. Esto puede tardar varios minutos...")
         update_programs = r'''
         Set-ExecutionPolicy Bypass -Scope Process -Force
         $raw = winget upgrade --accept-source-agreements --accept-package-agreements
         $apps = $raw | Select-Object -Skip 2 | ForEach-Object {
            ($_ -split '\s{2,}')[1]
         }
         foreach ($app in $apps) {
            if ($app -and $app -ne "Id") {
               if ($app -ne "Oracle.JavaRuntimeEnvironment") {
                  Write-Host "Buscando nueva version de [$app]"
                  winget upgrade --id $app --accept-source-agreements --accept-package-agreements -h
               }
               else {
                  Write-Host "Omitiendo actualización de [$app]"
               }
            }
         }

         '''
         self.run_and_log_ps(update_programs)

         # Este tarea es opcional aunque es muy demorada
         self.log("[+] Agendar diagnóstico de memoria RAM...")
         self.run_and_log("mdsched.exe")

         self.log("[✓] Optimización completada. Se recomienda reiniciar el equipo ahora mismo para que los cambios se apliquen.")

         if messagebox.askyesno("Reiniciar", "¿Desea reiniciar el equipo ahora?"):
            self.log("[!] Tarea de reinicio programada...")
            run_command("shutdown /r /t 3")

         self.finish_success()

      except Exception as e:
         self.log(f"[X] Error: {e}")
         self.finish_fail()

   def run_and_log(self, command):
      self.log(f"Ejecutando: {command}")
      out, err, code = run_command(command)
      if out:
         self.log(f"Salida: {out}")
      if err:
         self.log(f"Error: {err}")
      self.log(f"Código de salida: {code}")
      if code != 0:
         self.log(f"[!] Comando '{command}' terminó con error.")

   def run_and_log_ps(self, command):
      ps_command = f'powershell -Command "{command}"'
      self.log(f"Ejecutando (PowerShell): {command}")
      try:
         process = subprocess.run(
               ps_command,
               shell=True,
               stdout=subprocess.PIPE,
               stderr=subprocess.PIPE,
               text=True
         )
         out = process.stdout.strip()
         err = process.stderr.strip()
         code = process.returncode
         if out:
               self.log(f"Salida: {out}")
         if err:
               self.log(f"Error: {err}")
         self.log(f"Código de salida: {code}")
         if code != 0:
               self.log(f"[!] Comando PowerShell '{command}' terminó con error.")
         return out, err, code

      except Exception as e:
         self.log(f"Excepción al ejecutar comando PowerShell: {e}")
         return "", str(e), -1

   def finish_success(self):
      self.label_status.configure(text="Trabajo hecho.", text_color="green")
      self.update()
      self.after(3000, self.close_app)

   def finish_fail(self):
      self.label_status.configure(text="Una parte falló.", text_color="red")
      self.update()
      self.after(3000, self.close_app)

   def close_app(self):
      self.destroy()
      sys.exit(0)


if __name__ == "__main__":
   app = OptimizeApp()
   app.mainloop()
