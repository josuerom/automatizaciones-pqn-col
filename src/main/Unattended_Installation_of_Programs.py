"""
Unattended_Installation_of_Programs.py
Autor: Josué Romero
Empresa: Stefanini / PQN
Fecha: 15/Agosto/2025

Descripción:
Instala múltiples programas de forma desatendida sin intervención del usuario.
"""

import os
import subprocess
import datetime
import customtkinter as ctk
from tkinter import messagebox

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

APP_TITLE = "Instalación Desatendida de Programas PQN-COL"
APP_SIZE = "730x640"
LOG_FILENAME = "install_log.txt"

INSTALLERS = [
   {"name": "Adobe Acrobat Reader PDF", "file": "1_Reader.exe", "args": "/sAll /rs /rps /msi EULA_ACCEPT=YES"},
   {"name": "FortiClient VPN", "file": "2_FortiClient.exe", "args": "/quiet /norestart"},
   {"name": "Citrix Workspace App", "file": "3_Citrix.exe", "args": "/silent /noreboot"},
   {"name": "Java 8 Update 341", "file": "4_Java341.exe", "args": "/s"},
   {"name": ".NET Framework 3.5", "file": "5_NET35.exe", "args": "/quiet /norestart"},
   {"name": "TeamViewer Host 2025", "file": "6_TeamViewerHost.exe", "args": "/S"},
   {"name": "SupportAssist Dell", "file": "7_SupportAssistDell.exe", "args": "/quiet"},
   {"name": "SupportAssist Lenovo", "file": "7_SupportAssistLenovo.exe", "args": "/quiet"},
   {"name": "Microsoft Teams", "file": "8_Teams.exe", "args": ""}
]

OFFICE_INSTALLER = "9_Office365.exe"
OFFICE_CONFIG = "9_config.xml"

SEARCH_PATHS = [
   r"D:\Utilidades\Programas",   # prioridad primero D:
   r"C:\Utilidades\Programas"    # luego C:
]


def find_base_folder():
   """Devuelve la carpeta base donde están los instaladores (D: si existe, sino C:)"""
   for path in SEARCH_PATHS:
      if os.path.exists(path):
         return path
   return None


def find_installer(filename):
   """Busca un instalador en la carpeta base"""
   base_folder = find_base_folder()
   if base_folder:
      candidate = os.path.join(base_folder, filename)
      if os.path.exists(candidate):
         return candidate
   return None


def run_silent_installer(name, installer, arguments, log_path):
   if installer:
      try:
         subprocess.run([installer] + arguments.split(), check=True)
         write_log(log_path, f"[✓] {name} instalado correctamente.")
         return True, f"[✓] {name} instalado correctamente."
      except subprocess.CalledProcessError as e:
         write_log(log_path, f"[X] Error al instalar {name}: {str(e)}")
         return False, f"[X] Error al instalar {name}: {str(e)}"
   else:
      write_log(log_path, f"[!] Instalador no encontrado para {name}")
      return False, f"[!] Instalador no encontrado para {name}"


def write_log(path, content):
   with open(path, "a", encoding="utf-8") as f:
      f.write(content + "\n")


class InstallerApp(ctk.CTk):
   def __init__(self):
      super().__init__()
      self.title(APP_TITLE)
      self.geometry(APP_SIZE)
      self.resizable(False, False)
      self.build_ui()

   def build_ui(self):
      self.label_title = ctk.CTkLabel(self, text=APP_TITLE, font=ctk.CTkFont(size=24, weight="bold"))
      self.label_title.pack(pady=(15, 5))

      self.label_meta = ctk.CTkLabel(self,
                                    text="Autor: Josué Romero  |  Empresa: Stefanini / PQN  |  Fecha: 15/Agosto/2025",
                                    font=ctk.CTkFont(size=12))
      self.label_meta.pack()

      self.textbox = ctk.CTkTextbox(self, width=680, height=420)
      self.textbox.pack(padx=10, pady=10)

      self.button_start = ctk.CTkButton(self, text="Ejecutar", command=self.start_installation)
      self.button_start.pack(pady=(0, 10))

      # Barra de progreso
      self.progress = ctk.CTkProgressBar(self, width=680)
      self.progress.pack(pady=(0, 10))
      self.progress.set(0)

      # Evento Enter
      self.bind("<Return>", lambda event: self.start_installation())

   def log(self, message):
      self.textbox.configure(state="normal")
      self.textbox.insert("end", message + "\n")
      self.textbox.see("end")
      self.textbox.configure(state="disabled")

   def start_installation(self):
      self.button_start.configure(state="disabled")
      self.textbox.configure(state="normal")
      self.textbox.delete("1.0", "end")
      self.textbox.configure(state="disabled")
      self.progress.set(0)

      base_folder = find_base_folder()
      if not base_folder:
         messagebox.showerror("Error", "No se encontró la carpeta C: ni D:\\Utilidades\\Programas")
         self.button_start.configure(state="normal")
         return

      log_path = os.path.join(base_folder, LOG_FILENAME)
      write_log(log_path, f"\n===== REGISTRO DE INSTALACIÓN - {datetime.datetime.now()} =====")

      total = len(INSTALLERS) + 1  # +1 Office
      completed = 0

      for item in INSTALLERS:
         name = item["name"]
         file = find_installer(item["file"])
         args = item["args"]
         success, result = run_silent_installer(name, file, args, log_path)
         self.log(result)

         completed += 1
         self.progress.set(completed / total)
         self.update_idletasks()

      # Office
      self.log("[*] Verificando instalación de Office 365...")
      office_installer = find_installer(OFFICE_INSTALLER)
      config_xml = find_installer(OFFICE_CONFIG)

      if office_installer and config_xml:
         office_installed = self.check_office_installed()
         if office_installed:
               self.log("[!] Office ya está instalado. Se omite.")
               write_log(log_path, f"[i] Office ya estaba instalado - {datetime.datetime.now()}")
         else:
               self.log("[+] Instalando Office 365...")
               try:
                  subprocess.run([office_installer, "/configure", config_xml], check=True)
                  self.log("[✓] Office 365 instalado correctamente.")
                  write_log(log_path, f"[✓] Office 365 instalado correctamente - {datetime.datetime.now()}")
               except Exception as e:
                  self.log(f"[X] Error al instalar Office 365: {str(e)}")
                  write_log(log_path, f"[X] Error al instalar Office 365 - {datetime.datetime.now()} - {str(e)}")
      else:
         self.log("[!] Faltan 9_Office365.exe o 9_config.xml.")
         write_log(log_path, f"[!] Faltan archivos para instalar Office 365 - {datetime.datetime.now()}")

      completed += 1
      self.progress.set(completed / total)
      self.update_idletasks()

      self.log("\n[✓] Todas las instalaciones han finalizado.")
      self.log(f"[i] Puedes revisar el log en: {log_path}")
      self.button_start.configure(state="normal")

   def check_office_installed(self):
      try:
         output = subprocess.check_output(
               'reg query "HKLM\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall"',
               shell=True, text=True
         )
         return "Office" in output
      except:
         return False


if __name__ == "__main__":
   app = InstallerApp()
   app.mainloop()
