#!/usr/bin/env python3
"""
   PQN_COL_Domain_Manager.py
   Autor: Asistente Claude
   Basado en: PQN_COL_Equipment_Renamer.py por Josué Romero
   Empresa: Stefanini / PQN
   Fecha: 12/Septiembre/2025

   Descripción:
   Este programa:
   1. Crea punto de restauración
   2. Obtiene el serial del equipo y genera el nombre estándar
   3. Si está en dominio: solo renombra
   4. Si NO está en dominio: renombra Y une al dominio
   5. Se auto-ejecuta como administrador usando credenciales locales
   6. Reinicia automáticamente tras cambios exitosos
"""

import os
import sys
import time
import subprocess
import threading
import customtkinter as ctk
from pathlib import Path
from dotenv import load_dotenv

# --- Seteo de GUI ---
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

load_dotenv()

# --- Constantes ocultas ---
DOMAIN = os.getenv('DOMAIN')
DOMAIN_USER = os.getenv('DOMAIN_USER')
DOMAIN_PASS = os.getenv('DOMAIN_PASS')
LOCAL_ADMIN_USER = os.getenv('LOCAL_ADMIN_USER')
LOCAL_ADMIN_PASS = os.getenv('LOCAL_ADMIN_PASS')

APP_TITLE = "PQN-COL Domain Manager"
APP_SIZE = "600x500"

def is_admin():
    """Verifica si el script se ejecuta como administrador"""
    try:
        return os.getuid() == 0
    except AttributeError:
        import ctypes
        return ctypes.windll.shell32.IsUserAnAdmin() != 0

def run_as_admin():
    """Re-ejecuta el script como administrador usando credenciales locales"""
    try:
        script_path = sys.argv[0]
        if script_path.endswith('.py'):
            # Si es un script .py
            cmd = f'powershell.exe -Command "Start-Process python -ArgumentList \\"{script_path}\\" -Credential (New-Object System.Management.Automation.PSCredential(\\"{LOCAL_ADMIN_USER}\\", (ConvertTo-SecureString \\"{LOCAL_ADMIN_PASS}\\" -AsPlainText -Force))) -Verb RunAs -WindowStyle Hidden"'
        else:
            # Si es un .exe
            cmd = f'powershell.exe -Command "Start-Process \\"{script_path}\\" -Credential (New-Object System.Management.Automation.PSCredential(\\"{LOCAL_ADMIN_USER}\\", (ConvertTo-SecureString \\"{LOCAL_ADMIN_PASS}\\" -AsPlainText -Force))) -Verb RunAs"'
        
        subprocess.run(cmd, shell=True, check=True)
        sys.exit(0)
    except Exception as e:
        print(f"Error al ejecutar como administrador: {e}")
        return False

def run_powershell(script):
    """Ejecuta un script de PowerShell y retorna (stdout, stderr, returncode)"""
    try:
        cmd = ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script]
        result = subprocess.run(cmd, capture_output=True, text=True, shell=True, timeout=60)
        return result.stdout.strip(), result.stderr.strip(), result.returncode
    except subprocess.TimeoutExpired:
        return "", "Timeout ejecutando PowerShell", 1
    except Exception as e:
        return "", str(e), 1

def get_computer_serial():
    """Obtiene el número de serie del equipo"""
    scripts = [
        "(Get-CimInstance -ClassName Win32_BIOS).SerialNumber",
        "(Get-WmiObject -Class Win32_BIOS).SerialNumber",
        "(Get-CimInstance -ClassName Win32_ComputerSystemProduct).IdentifyingNumber",
        "(Get-WmiObject -Class Win32_ComputerSystemProduct).IdentifyingNumber"
    ]
    
    for script in scripts:
        try:
            stdout, stderr, code = run_powershell(script)
            if code == 0 and stdout.strip():
                return stdout.strip()
        except:
            continue
    
    return None

def build_hostname(serial):
    """Construye el nombre del host basado en el serial"""
    if not serial:
        return None
    
    # Extraer solo dígitos del serial
    digits = ''.join(c for c in serial if c.isdigit())
    
    if len(digits) >= 7:
        # Tomar los últimos 7 dígitos
        last_seven = digits[-7:]
    else:
        # Si no hay 7 dígitos, tomar caracteres alfanuméricos
        alnum = ''.join(c for c in serial if c.isalnum())
        if len(alnum) >= 7:
            last_seven = alnum[-7:].upper()
        else:
            # Rellenar con ceros si es necesario
            last_seven = alnum.ljust(7, '0').upper()
    
    return f"{last_seven}-PQN-COL"

def create_restore_point():
    """Crea un punto de restauración del sistema"""
    script = '''
    try {
        # Habilitar restauración del sistema si no está habilitada
        $status = Get-ComputerRestorePoint -ErrorAction SilentlyContinue
        if (-not $status) {
            Enable-ComputerRestore -Drive "C:\\" -Confirm:$false
            Start-Sleep -Seconds 3
        }
        
        # Crear punto de restauración
        Checkpoint-Computer -Description "Cambio de nombre de host" -RestorePointType "MODIFY_SETTINGS" -ErrorAction Stop
        Write-Output "SUCCESS: Punto de restauración creado"
    } catch {
        Write-Error "ERROR: No se pudo crear el punto de restauración: $($_.Exception.Message)"
        exit 1
    }
    '''
    return run_powershell(script)

def check_domain_status():
    """Verifica si el equipo está en el dominio"""
    script = '''
    try {
        $cs = Get-CimInstance -ClassName Win32_ComputerSystem
        $inDomain = $cs.PartOfDomain
        $domainName = $cs.Domain
        $computerName = $cs.Name
        
        Write-Output "DOMAIN_STATUS:$inDomain"
        Write-Output "DOMAIN_NAME:$domainName"
        Write-Output "COMPUTER_NAME:$computerName"
    } catch {
        Write-Error "ERROR: No se pudo verificar el estado del dominio"
        exit 1
    }
    '''
    
    stdout, stderr, code = run_powershell(script)
    if code != 0:
        return False, None, None
    
    in_domain = False
    domain_name = ""
    computer_name = ""
    
    for line in stdout.split('\n'):
        if line.startswith('DOMAIN_STATUS:'):
            in_domain = line.split(':', 1)[1].strip().lower() == 'true'
        elif line.startswith('DOMAIN_NAME:'):
            domain_name = line.split(':', 1)[1].strip().lower()
        elif line.startswith('COMPUTER_NAME:'):
            computer_name = line.split(':', 1)[1].strip()
    
    return in_domain, domain_name, computer_name

def rename_computer_only(new_name):
    """Renombra solo el equipo (cuando ya está en dominio)"""
    script = f'''
    try {{
        Rename-Computer -NewName "{new_name}" -Force -ErrorAction Stop
        Write-Output "SUCCESS: Equipo renombrado a {new_name}"
    }} catch {{
        Write-Error "ERROR: No se pudo renombrar el equipo: $($_.Exception.Message)"
        exit 1
    }}
    '''
    return run_powershell(script)

def join_domain_and_rename(new_name):
    """Une al dominio y renombra el equipo"""
    script = f'''
    try {{
        $secpass = ConvertTo-SecureString "{DOMAIN_PASS}" -AsPlainText -Force
        $cred = New-Object System.Management.Automation.PSCredential("{DOMAIN_USER}", $secpass)
        
        Add-Computer -DomainName "{DOMAIN}" -NewName "{new_name}" -Credential $cred -Force -ErrorAction Stop
        Write-Output "SUCCESS: Equipo unido al dominio {DOMAIN} con nombre {new_name}"
    }} catch {{
        Write-Error "ERROR: No se pudo unir al dominio: $($_.Exception.Message)"
        exit 1
    }}
    '''
    return run_powershell(script)

def restart_computer():
    """Reinicia el equipo inmediatamente"""
    try:
        subprocess.run("shutdown /r /t 0 /f", shell=True)
    except Exception as e:
        print(f"Error al reiniciar: {e}")

class DomainManagerApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry(APP_SIZE)
        self.resizable(False, False)
        
        # Variables de estado
        self.processing = False
        
        self.create_widgets()
        
        # Auto-ejecutar al iniciar
        self.after(1000, self.auto_execute)

    def create_widgets(self):
        # Título
        title_label = ctk.CTkLabel(
            self, 
            text="PQN-COL Domain Manager",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        title_label.pack(pady=(20, 5))
        
        # Subtítulo
        subtitle_label = ctk.CTkLabel(
            self,
            text="Gestión automática de nombres de host y unión al dominio",
            font=ctk.CTkFont(size=14)
        )
        subtitle_label.pack(pady=(0, 20))
        
        # Área de log
        self.log_text = ctk.CTkTextbox(self, width=550, height=300)
        self.log_text.pack(pady=10, padx=20)
        
        # Botón de ejecución manual
        self.execute_btn = ctk.CTkButton(
            self,
            text="Ejecutar Manualmente",
            command=self.manual_execute,
            width=200,
            height=35
        )
        self.execute_btn.pack(pady=10)
        
        # Estado
        self.status_label = ctk.CTkLabel(
            self,
            text="Listo para ejecutar...",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.status_label.pack(pady=(10, 20))

    def log_message(self, message):
        """Agrega un mensaje al log"""
        timestamp = time.strftime("%H:%M:%S")
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"[{timestamp}] {message}\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")
        self.update()

    def set_status(self, message, color="white"):
        """Actualiza el mensaje de estado"""
        self.status_label.configure(text=message, text_color=color)
        self.update()

    def auto_execute(self):
        """Ejecución automática al iniciar la aplicación"""
        if not self.processing:
            self.log_message("Iniciando ejecución automática...")
            self.execute_process()

    def manual_execute(self):
        """Ejecución manual mediante botón"""
        if not self.processing:
            self.log_message("Iniciando ejecución manual...")
            self.execute_process()

    def execute_process(self):
        """Ejecuta el proceso principal en un hilo separado"""
        if self.processing:
            return
        
        self.processing = True
        self.execute_btn.configure(state="disabled")
        self.set_status("Procesando...", "yellow")
        
        thread = threading.Thread(target=self.main_process)
        thread.daemon = True
        thread.start()

    def main_process(self):
        """Proceso principal del programa"""
        try:
            # Paso 1: Crear punto de restauración
            self.log_message("=== CREANDO PUNTO DE RESTAURACIÓN ===")
            self.set_status("Creando punto de restauración...", "yellow")
            
            stdout, stderr, code = create_restore_point()
            if code == 0:
                self.log_message("✓ Punto de restauración creado exitosamente")
            else:
                self.log_message(f"⚠ Advertencia al crear punto de restauración: {stderr}")
                self.log_message("Continuando con el proceso...")

            # Paso 2: Obtener serial y generar nombre
            self.log_message("\n=== OBTENIENDO INFORMACIÓN DEL EQUIPO ===")
            self.set_status("Obteniendo serial del equipo...", "yellow")
            
            serial = get_computer_serial()
            if not serial:
                self.log_message("✗ ERROR: No se pudo obtener el serial del equipo")
                self.finish_error()
                return
            
            self.log_message(f"✓ Serial obtenido: {serial}")
            
            new_hostname = build_hostname(serial)
            if not new_hostname:
                self.log_message("✗ ERROR: No se pudo generar el nombre del host")
                self.finish_error()
                return
            
            self.log_message(f"✓ Nombre de host generado: {new_hostname}")

            # Paso 3: Verificar estado del dominio
            self.log_message("\n=== VERIFICANDO ESTADO DEL DOMINIO ===")
            self.set_status("Verificando estado del dominio...", "yellow")
            
            in_domain, domain_name, current_name = check_domain_status()
            
            self.log_message(f"✓ Nombre actual: {current_name}")
            self.log_message(f"✓ En dominio: {'Sí' if in_domain else 'No'}")
            self.log_message(f"✓ Dominio actual: {domain_name if domain_name else 'N/A'}")

            # Verificar si ya tiene el nombre correcto
            if current_name.upper() == new_hostname.upper():
                self.log_message(f"✓ El equipo ya tiene el nombre correcto: {new_hostname}")
                if in_domain and domain_name.lower() == DOMAIN.lower():
                    self.log_message("✓ El equipo ya está correctamente configurado")
                    self.finish_success("Configuración ya correcta - No se requieren cambios")
                    return

            # Paso 4: Ejecutar acción según estado
            restart_required = False
            
            if in_domain and domain_name.lower() == DOMAIN.lower():
                # Solo renombrar
                self.log_message(f"\n=== RENOMBRANDO EQUIPO ===")
                self.set_status("Renombrando equipo...", "yellow")
                
                stdout, stderr, code = rename_computer_only(new_hostname)
                if code == 0:
                    self.log_message(f"✓ Equipo renombrado exitosamente a: {new_hostname}")
                    restart_required = True
                else:
                    self.log_message(f"✗ ERROR al renombrar: {stderr}")
                    self.finish_error()
                    return
                    
            else:
                # Unir al dominio y renombrar
                self.log_message(f"\n=== UNIENDO AL DOMINIO Y RENOMBRANDO ===")
                self.set_status("Uniendo al dominio...", "yellow")
                
                stdout, stderr, code = join_domain_and_rename(new_hostname)
                if code == 0:
                    self.log_message(f"✓ Equipo unido al dominio {DOMAIN} con nombre: {new_hostname}")
                    restart_required = True
                else:
                    self.log_message(f"✗ ERROR al unir al dominio: {stderr}")
                    self.finish_error()
                    return

            # Paso 5: Reiniciar si es necesario
            if restart_required:
                self.log_message("\n=== REINICIANDO EQUIPO ===")
                self.set_status("Reiniciando equipo en 10 segundos...", "green")
                
                for i in range(10, 0, -1):
                    self.log_message(f"Reiniciando en {i} segundos...")
                    time.sleep(1)
                
                self.log_message("✓ Reiniciando equipo...")
                restart_computer()
            
        except Exception as e:
            self.log_message(f"✗ ERROR INESPERADO: {str(e)}")
            self.finish_error()

    def finish_success(self, message="Proceso completado exitosamente"):
        """Finaliza el proceso con éxito"""
        self.processing = False
        self.set_status(message, "green")
        self.execute_btn.configure(state="normal")
        self.log_message(f"\n✓ {message}")

    def finish_error(self):
        """Finaliza el proceso con error"""
        self.processing = False
        self.set_status("Proceso finalizado con errores", "red")
        self.execute_btn.configure(state="normal")
        self.log_message("\n✗ Proceso finalizado con errores")

def main():
    # Verificar si se ejecuta como administrador
    if not is_admin():
        print("El programa necesita ejecutarse como administrador...")
        print("Intentando elevar privilegios...")
        run_as_admin()
        return

    # Crear y ejecutar la aplicación
    app = DomainManagerApp()
    app.mainloop()

if __name__ == "__main__":
    main()