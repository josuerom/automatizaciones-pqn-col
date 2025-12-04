"""
PQN_COL_Equipment_Renamer.py - Versión Profesional Mejorada
Autor: Josué Romero
Empresa: Stefanini / PQN
Fecha: 06/Noviembre/2025
Versión: 3.0 Professional Edition

Licencia: Propiedad de Stefanini / PQN - Todos los derechos reservados
Copyright © 2025 Josué Romero - Stefanini / PQN

Descripción:
Sistema automatizado para renombrar equipos Windows y unirlos a dominio.
Formato: [7DIGITOS]-[PQN/CCS/CBQ]-COL
"""

import sys
import os
import subprocess
import ctypes
import time
import re
import json
import customtkinter as ctk
from tkinter import messagebox
from datetime import datetime
from pathlib import Path
import threading
import hashlib

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
ctk.set_default_color_theme("blue")

APP_TITLE = "Renombrador de Equipos Multi-Site"
APP_VERSION = f"v{__version__}"
APP_SIZE = "800x980"

# Paleta de colores profesional (Rojo-Carmesí-Oscuro)
COLOR_PRIMARY = "#c62828"  # Rojo profundo
COLOR_SECONDARY = "#d32f2f"  # Rojo medio
COLOR_ACCENT = "#e53935"  # Rojo brillante
COLOR_SUCCESS = "#00e676"  # Verde éxito
COLOR_WARNING = "#ffc107"  # Amarillo advertencia
COLOR_ERROR = "#ff1744"  # Rojo error brillante
COLOR_BG_DARK = "#0a0a0a"  # Negro profundo
COLOR_BG_MEDIUM = "#1a1a1a"  # Negro medio
COLOR_BG_LIGHT = "#2d2d2d"  # Gris oscuro
COLOR_TEXT_WHITE = "#ffffff"  # Texto blanco
COLOR_TEXT_GRAY = "#b0bec5"  # Texto gris claro

# Fuentes (Incrementadas +2px según requerimiento)
FONT_TITLE = ("Segoe UI", 28, "bold")  # Era 26
FONT_SUBTITLE = ("Segoe UI", 13)  # Era 11
FONT_CONSOLE = ("Consolas", 12)  # Era 10
FONT_BUTTON = ("Segoe UI", 14, "bold")  # Era 12
FONT_LABEL = ("Segoe UI", 13, "bold")  # Era 11
FONT_INFO = ("Segoe UI", 13)  # Era 11

# Configuración de logs
LOG_DIR = Path("C:/ProgramData/PQN_Renamer")
LOG_FILE = LOG_DIR / "renamer.log"
BACKUP_FILE = LOG_DIR / "backup_config.json"

# Constantes de nomenclatura
SERIAL_LENGTH = 7
MAX_NETBIOS_LENGTH = 15
COUNTRY_CODE = "COL"
SITE_OPTIONS = ["PQN", "CCS", "CBQ"]

# Credenciales de dominio (para PQN)
DOMAIN_NAME = "proquinal.com"
DOMAIN_USER = "romero-josue"
DOMAIN_PASSWORD = "Jo320872.."


# ============================================================================
# FUNCIONES DE ELEVACIÓN DE PRIVILEGIOS
# ============================================================================


def is_admin():
    """Verifica si el script se está ejecutando como administrador."""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


def run_as_admin():
    """Reinicia el script con privilegios de administrador."""
    try:
        if sys.argv[0].endswith(".py"):
            ctypes.windll.shell32.ShellExecuteW(
                None, "runas", sys.executable, f'"{sys.argv[0]}"', None, 1
            )
        else:
            ctypes.windll.shell32.ShellExecuteW(
                None, "runas", sys.executable, " ".join(sys.argv), None, 1
            )
        sys.exit(0)
    except Exception as e:
        messagebox.showerror(
            "Error de Privilegios", f"No se pudo elevar privilegios:\n{e}"
        )
        sys.exit(1)


# ============================================================================
# FUNCIONES AUXILIARES
# ============================================================================


def setup_logging():
    """Crea el directorio de logs."""
    try:
        LOG_DIR.mkdir(parents=True, exist_ok=True)
        return True
    except:
        return False


def log_to_file(message, level="INFO"):
    """Registra en archivo de log."""
    try:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = f"[{timestamp}] [{level}] {message}\n"
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(log_entry)
    except:
        pass


def calculate_config_hash(data):
    """Calcula hash SHA-256 de la configuración."""
    try:
        sha256_hash = hashlib.sha256()
        sha256_hash.update(str(data).encode("utf-8"))
        return sha256_hash.hexdigest()
    except:
        return "N/A"


def run_powershell(command, timeout=30):
    """
    Ejecuta comando PowerShell con privilegios.

    Returns:
       tuple: (success: bool, output: str, error: str)
    """
    try:
        result = subprocess.run(
            [
                "powershell",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-Command",
                command,
            ],
            capture_output=True,
            text=True,
            timeout=timeout,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        success = result.returncode == 0
        return success, result.stdout.strip(), result.stderr.strip()
    except subprocess.TimeoutExpired:
        return False, "", "Timeout ejecutando comando PowerShell"
    except Exception as e:
        return False, "", str(e)


def get_bios_serial():
    """Obtiene serial del BIOS usando PowerShell CIM."""
    success, output, error = run_powershell(
        "(Get-CimInstance -ClassName Win32_BIOS).SerialNumber.Trim()"
    )
    if success and output:
        serial = re.sub(r"\s+", "", output).strip()
        return serial if serial else None
    return None


def get_current_hostname():
    """Obtiene el nombre actual del equipo."""
    success, output, _ = run_powershell("$env:COMPUTERNAME")
    return output.strip().upper() if success else None


def get_manufacturer():
    """Obtiene el fabricante del equipo."""
    success, output, _ = run_powershell(
        "(Get-CimInstance -ClassName Win32_ComputerSystem).Manufacturer.Trim()"
    )
    return output.strip() if success else "Desconocido"


def get_model():
    """Obtiene el modelo del equipo."""
    success, output, _ = run_powershell(
        "(Get-CimInstance -ClassName Win32_ComputerSystem).Model.Trim()"
    )
    return output.strip() if success else "Desconocido"


def is_in_domain():
    """Verifica si el equipo está unido a un dominio."""
    success, output, _ = run_powershell(
        "(Get-CimInstance -ClassName Win32_ComputerSystem).PartOfDomain"
    )
    return output.strip().lower() == "true" if success else False


def get_current_domain():
    """Obtiene el dominio actual."""
    success, output, _ = run_powershell(
        "(Get-CimInstance -ClassName Win32_ComputerSystem).Domain"
    )
    return output.strip() if success else "WORKGROUP"


def validate_serial(serial):
    """Valida el serial del BIOS."""
    if not serial:
        return False
    return bool(re.match(r"^[A-Za-z0-9\-]+$", serial))


def validate_hostname(name):
    """Valida el nombre del equipo según estándares NetBIOS."""
    if not name:
        return False, "Nombre vacío"

    if len(name) > MAX_NETBIOS_LENGTH:
        return False, f"Excede {MAX_NETBIOS_LENGTH} caracteres"

    if not re.match(r"^[A-Za-z0-9][A-Za-z0-9\-]*[A-Za-z0-9]$", name) and len(name) > 1:
        return False, "Caracteres inválidos o formato incorrecto"

    # Validar nombres reservados de Windows
    reserved = ["CON", "PRN", "AUX", "NUL", "COM1", "COM2", "LPT1", "LPT2"]
    if name.upper() in reserved:
        return False, f"'{name}' es un nombre reservado del sistema"

    return True, ""


def build_hostname(serial, site):
    """
    Construye el nuevo hostname según el formato establecido.

    Args:
       serial: Serial del BIOS
       site: Sitio (PQN, CCS, o CBQ)

    Returns:
       str: Nuevo hostname o None si es inválido
    """
    if not validate_serial(serial):
        return None

    # Limpiar serial (solo alfanuméricos)
    clean_serial = re.sub(r"[^A-Za-z0-9]", "", serial)

    # Tomar últimos 7 dígitos
    if len(clean_serial) >= SERIAL_LENGTH:
        serial_part = clean_serial[-SERIAL_LENGTH:]
    else:
        serial_part = clean_serial.zfill(SERIAL_LENGTH)

    # Construir nombre: [7DIGITOS]-[SITE]-COL
    hostname = f"{serial_part}-{site}-{COUNTRY_CODE}".upper()

    # Validar longitud NetBIOS
    if len(hostname) > MAX_NETBIOS_LENGTH:
        hostname = hostname[:MAX_NETBIOS_LENGTH]

    is_valid, error = validate_hostname(hostname)
    return hostname if is_valid else None


def save_backup(data):
    """Guarda backup de configuración con hash SHA-256."""
    try:
        data["sha256"] = calculate_config_hash(data)
        with open(BACKUP_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        return True
    except:
        return False


def rename_computer_modern(new_name):
    """
    Renombra el equipo usando método compatible con Windows 11 24H2.

    Args:
       new_name: Nuevo nombre del equipo

    Returns:
       bool: True si el renombrado fue exitoso
    """
    # Método 1: Usar Rename-Computer (PowerShell)
    ps_script = f"""
$ErrorActionPreference = 'Stop'
try {{
   Rename-Computer -NewName "{new_name}" -Force -PassThru -ErrorAction Stop | Out-Null
   Write-Output "SUCCESS"
   exit 0
}} catch {{
   Write-Output "ERROR: $($_.Exception.Message)"
   exit 1
}}
"""

    success, output, error = run_powershell(ps_script, timeout=60)

    if success and "SUCCESS" in output:
        return True

    # Método 2: Usar WMI como fallback (Windows 11 24H2 compatible)
    ps_script_wmi = f"""
$ErrorActionPreference = 'Stop'
try {{
   $computer = Get-WmiObject Win32_ComputerSystem
   $computer.Rename("{new_name}")
   Write-Output "SUCCESS"
   exit 0
}} catch {{
   Write-Output "ERROR: $($_.Exception.Message)"
   exit 1
}}
"""

    success, output, error = run_powershell(ps_script_wmi, timeout=60)
    return success and "SUCCESS" in output


def join_domain(domain, username, password, new_name):
    """
    Une el equipo a un dominio de Active Directory.

    Args:
       domain: Nombre del dominio
       username: Usuario del dominio
       password: Contraseña del usuario
       new_name: Nuevo nombre del equipo

    Returns:
       bool: True si la unión fue exitosa
    """
    # Crear credencial segura
    ps_script = f"""
$ErrorActionPreference = 'Stop'
try {{
   # Crear credencial
   $securePassword = ConvertTo-SecureString "{password}" -AsPlainText -Force
   $credential = New-Object System.Management.Automation.PSCredential("{username}@{domain}", $securePassword)
   
   # Unir al dominio con el nuevo nombre
   Add-Computer -DomainName "{domain}" -NewName "{new_name}" -Credential $credential -Force -ErrorAction Stop
   
   Write-Output "SUCCESS"
   exit 0
}} catch {{
   Write-Output "ERROR: $($_.Exception.Message)"
   exit 1
}}
"""

    success, output, error = run_powershell(ps_script, timeout=120)
    return success and "SUCCESS" in output


def restart_computer(delay=15):
    """Reinicia el equipo después del delay especificado."""
    time.sleep(delay)
    run_powershell("Restart-Computer -Force", timeout=13)


# ============================================================================
# CLASE PRINCIPAL DE LA APLICACIÓN
# ============================================================================


class RenamerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Configuración de ventana
        self.title(f"{APP_TITLE} {APP_VERSION}")
        self.geometry(APP_SIZE)
        self.resizable(False, False)

        # Variables de estado
        self.system_info = None
        self.is_processing = False

        # Construir interfaz
        self.build_ui()

        # Verificar prerequisitos
        self.after(500, self.check_prerequisites)

    def build_ui(self):
        """Construye la interfaz de usuario moderna y profesional."""

        # Marco principal
        main_frame = ctk.CTkFrame(self, fg_color=COLOR_BG_DARK)
        main_frame.pack(
            fill="both", expand=True, padx=10, pady=10
        )  # Reducimos el padding del marco principal.

        # === ENCABEZADO MODERNO ===
        header_frame = ctk.CTkFrame(
            main_frame,
            fg_color=COLOR_PRIMARY,
            corner_radius=12,
            border_width=2,
            border_color=COLOR_ACCENT,
        )
        header_frame.pack(
            fill="x", pady=(0, 10)
        )  # Reducimos el espacio inferior del encabezado.

        title_label = ctk.CTkLabel(
            header_frame,
            text="💻 " + APP_TITLE,
            font=FONT_TITLE,
            text_color=COLOR_TEXT_WHITE,
        )
        title_label.pack(pady=(10, 5))  # Reducimos el padding arriba y abajo.

        subtitle_label = ctk.CTkLabel(
            header_frame,
            text=f"{__company__} | {__author__} | {APP_VERSION}",
            font=FONT_SUBTITLE,
            text_color=COLOR_ACCENT,
        )
        subtitle_label.pack(pady=(0, 5))  # Reducimos el espacio inferior.

        copyright_label = ctk.CTkLabel(
            header_frame,
            text=__copyright__,
            font=("Segoe UI", 11),
            text_color=COLOR_TEXT_GRAY,
        )
        copyright_label.pack(pady=(0, 10))  # Reducimos el espacio inferior.

        # === INFORMACIÓN ACTUAL DEL SISTEMA ===
        info_frame = ctk.CTkFrame(
            main_frame, fg_color=COLOR_BG_MEDIUM, corner_radius=10
        )
        info_frame.pack(fill="x", pady=(0, 10))  # Reducimos el espaciado entre marcos.

        info_title = ctk.CTkLabel(
            info_frame,
            text="📋 Información Actual del Sistema",
            font=FONT_LABEL,
            text_color=COLOR_TEXT_WHITE,
        )
        info_title.pack(
            pady=(10, 5), anchor="w", padx=15
        )  # Reducimos el padding alrededor del título.

        # Reducción del espaciado entre las etiquetas de la información
        self.info_serial = ctk.CTkLabel(
            info_frame,
            text="Serial BIOS: Obteniendo información...",
            font=FONT_INFO,
            text_color=COLOR_TEXT_WHITE,
            anchor="w",
        )
        self.info_serial.pack(pady=3, padx=15, anchor="w")

        self.info_manufacturer = ctk.CTkLabel(
            info_frame,
            text="Fabricante: Obteniendo información...",
            font=FONT_INFO,
            text_color=COLOR_TEXT_WHITE,
            anchor="w",
        )
        self.info_manufacturer.pack(pady=3, padx=15, anchor="w")

        self.info_model = ctk.CTkLabel(
            info_frame,
            text="Modelo: Obteniendo información...",
            font=FONT_INFO,
            text_color=COLOR_TEXT_WHITE,
            anchor="w",
        )
        self.info_model.pack(pady=3, padx=15, anchor="w")

        self.info_current_name = ctk.CTkLabel(
            info_frame,
            text="Nombre actual: Obteniendo información...",
            font=FONT_INFO,
            text_color=COLOR_TEXT_WHITE,
            anchor="w",
        )
        self.info_current_name.pack(pady=3, padx=15, anchor="w")

        self.info_domain = ctk.CTkLabel(
            info_frame,
            text="Dominio/Grupo: Obteniendo información...",
            font=FONT_INFO,
            text_color=COLOR_TEXT_WHITE,
            anchor="w",
        )
        self.info_domain.pack(
            pady=(3, 10), padx=15, anchor="w"
        )  # Reducimos el espaciado aquí también.

        # === SELECCIÓN DE SITIO ===
        config_frame = ctk.CTkFrame(
            main_frame, fg_color=COLOR_BG_MEDIUM, corner_radius=10
        )
        config_frame.pack(
            fill="x", pady=(0, 10)
        )  # Reducimos el espaciado entre marcos.

        config_title = ctk.CTkLabel(
            config_frame,
            text="⚙️ Configuración de Nomenclatura",
            font=FONT_LABEL,
            text_color=COLOR_TEXT_WHITE,
        )
        config_title.pack(pady=(10, 5), anchor="w", padx=15)  # Reducimos el espaciado.

        site_label = ctk.CTkLabel(
            config_frame,
            text="Seleccione el sitio para la nomenclatura:",
            font=FONT_INFO,
            text_color=COLOR_TEXT_WHITE,
        )
        site_label.pack(pady=(0, 5), padx=15, anchor="w")  # Reducimos el espaciado.

        self.site_var = ctk.StringVar(value="PQN")
        site_buttons_frame = ctk.CTkFrame(config_frame, fg_color="transparent")
        site_buttons_frame.pack(
            pady=(0, 10), padx=15, anchor="w"
        )  # Reducimos el espaciado entre los botones.

        for site in SITE_OPTIONS:
            ctk.CTkRadioButton(
                site_buttons_frame,
                text=f"{site}-COL",
                variable=self.site_var,
                value=site,
                font=FONT_INFO,
                text_color=COLOR_TEXT_WHITE,
                fg_color=COLOR_PRIMARY,
                hover_color=COLOR_SECONDARY,
                command=self.update_preview,
            ).pack(
                side="left", padx=(0, 15)
            )  # Reducimos el padding entre los botones de selección.

        # === PREVIEW DEL NUEVO NOMBRE ===
        preview_frame = ctk.CTkFrame(
            main_frame, fg_color=COLOR_BG_MEDIUM, corner_radius=10
        )
        preview_frame.pack(
            fill="x", pady=(0, 10)
        )  # Reducimos el espaciado entre marcos.

        preview_title = ctk.CTkLabel(
            preview_frame,
            text="🎯 Vista Previa del Nuevo Nombre",
            font=FONT_LABEL,
            text_color=COLOR_TEXT_WHITE,
        )
        preview_title.pack(pady=(10, 5), anchor="w", padx=15)  # Reducimos el espaciado.

        self.preview_name = ctk.CTkLabel(
            preview_frame,
            text="-------",
            font=("Consolas", 18, "bold"),
            text_color=COLOR_ACCENT,
        )
        self.preview_name.pack(
            pady=(0, 5), padx=15, anchor="w"
        )  # Reducimos el espaciado entre el texto.

        self.preview_validation = ctk.CTkLabel(
            preview_frame, text="", font=("Segoe UI", 11), text_color=COLOR_SUCCESS
        )
        self.preview_validation.pack(
            pady=(0, 10), padx=15, anchor="w"
        )  # Reducimos el espaciado.

        # === ÁREA DE LOGS ===
        log_label = ctk.CTkLabel(
            main_frame,
            text="📝 Registro de Actividad",
            font=FONT_LABEL,
            text_color=COLOR_TEXT_WHITE,
            anchor="w",
        )
        log_label.pack(
            pady=(5, 5), anchor="w"
        )  # Reducimos el espaciado entre el título y el log.

        self.text_log = ctk.CTkTextbox(
            main_frame,
            width=740,
            height=200,
            font=FONT_CONSOLE,
            fg_color=COLOR_BG_LIGHT,
            text_color=COLOR_TEXT_WHITE,
            border_width=2,
            border_color=COLOR_SECONDARY,
            corner_radius=10,
        )
        self.text_log.pack(
            pady=(0, 10)
        )  # Reducimos el espaciado entre la caja de texto y la parte inferior.
        self.log("✓ Sistema inicializado correctamente", "SUCCESS")
        self.log(f"✓ Licencia: {__license__}", "INFO")

        # === BARRA DE ESTADO ===
        self.status_label = ctk.CTkLabel(
            main_frame,
            text="Estado: Inicializando sistema...",
            font=FONT_INFO,
            text_color=COLOR_WARNING,
        )
        self.status_label.pack(
            pady=(0, 10)
        )  # Reducimos el espaciado entre la barra de estado y la parte inferior.

        # === BOTONES ===
        button_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        button_frame.pack(
            fill="x", pady=(10, 5)
        )  # Reducimos el espaciado entre los botones.

        self.btn_execute = ctk.CTkButton(
            button_frame,
            text="🚀 Aplicar Cambios y Reiniciar",
            command=self.on_execute,
            font=FONT_BUTTON,
            height=50,
            corner_radius=10,
            fg_color=COLOR_PRIMARY,
            hover_color=COLOR_SECONDARY,
            border_width=3,
            border_color=COLOR_ACCENT,
            text_color=COLOR_TEXT_WHITE,
        )
        self.btn_execute.pack(
            side="left", expand=True, fill="x", padx=(0, 5)
        )  # Reducimos el padding lateral.

        self.btn_export = ctk.CTkButton(
            button_frame,
            text="💾 Exportar Log",
            command=self.export_log,
            font=FONT_BUTTON,
            height=50,
            corner_radius=10,
            fg_color=COLOR_BG_LIGHT,
            hover_color=COLOR_BG_MEDIUM,
            border_width=2,
            border_color=COLOR_SECONDARY,
            text_color=COLOR_TEXT_WHITE,
        )
        self.btn_export.pack(
            side="left", expand=True, fill="x", padx=(5, 0)
        )  # Reducimos el padding lateral y el espaciado entre botones.

        # Atajos de teclado
        self.bind("<Escape>", lambda e: self.quit())

    def log(self, message, level="INFO"):
        """Registra mensajes en el log con formato."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        icons = {
            "INFO": "ℹ",
            "SUCCESS": "✓",
            "WARNING": "⚠",
            "ERROR": "✗",
            "PROCESS": "⚙",
        }
        icon = icons.get(level, "•")

        formatted_msg = f"[{timestamp}] {icon} {message}\n"

        self.text_log.configure(state="normal")
        self.text_log.insert("end", formatted_msg)
        self.text_log.see("end")
        self.text_log.configure(state="disabled")

        # También registrar en archivo
        log_to_file(message, level)

    def update_status(self, text, color=COLOR_TEXT_WHITE):
        """Actualiza el label de estado."""
        self.status_label.configure(text=f"Estado: {text}", text_color=color)

    def check_prerequisites(self):
        """Verifica prerequisitos del sistema."""
        setup_logging()
        self.log("Verificando prerequisitos del sistema...", "PROCESS")

        # Verificar Windows 11
        try:
            success, output, _ = run_powershell(
                "(Get-WmiObject Win32_OperatingSystem).Caption"
            )
            if success:
                self.log(f"✓ Sistema operativo: {output}", "SUCCESS")
        except:
            pass

        # Obtener información del sistema
        self.log("Obteniendo información del hardware...", "PROCESS")

        serial = get_bios_serial()
        if not serial:
            self.log("✗ No se pudo obtener el serial del BIOS", "ERROR")
            self.update_status("Error: Sin serial del BIOS", COLOR_ERROR)
            messagebox.showerror(
                "Error del Sistema",
                "No se pudo obtener el serial del BIOS.\n\n"
                "Verifique que PowerShell esté disponible y funcional.",
            )
            self.btn_execute.configure(state="disabled")
            return

        current_name = get_current_hostname()
        manufacturer = get_manufacturer()
        model = get_model()
        in_domain = is_in_domain()
        current_domain = get_current_domain()

        self.system_info = {
            "serial": serial,
            "manufacturer": manufacturer,
            "model": model,
            "current_name": current_name,
            "in_domain": in_domain,
            "current_domain": current_domain,
        }

        # Actualizar interfaz con información obtenida
        self.info_serial.configure(
            text=f"Serial BIOS: {serial}", text_color=COLOR_SUCCESS
        )
        self.info_manufacturer.configure(
            text=f"Fabricante: {manufacturer}", text_color=COLOR_TEXT_WHITE
        )
        self.info_model.configure(text=f"Modelo: {model}", text_color=COLOR_TEXT_WHITE)
        self.info_current_name.configure(
            text=f"Nombre actual: {current_name}", text_color=COLOR_TEXT_WHITE
        )
        self.info_domain.configure(
            text=f"Dominio/Grupo: {current_domain} {'(Dominio)' if in_domain else '(Grupo de trabajo)'}",
            text_color=COLOR_TEXT_WHITE,
        )

        self.log(f"✓ Serial: {serial}", "SUCCESS")
        self.log(f"✓ Fabricante: {manufacturer}", "SUCCESS")
        self.log(f"✓ Modelo: {model}", "SUCCESS")
        self.log(f"✓ Nombre actual: {current_name}", "SUCCESS")
        self.log(f"✓ Dominio actual: {current_domain}", "SUCCESS")

        # Actualizar preview inicial
        self.update_preview()

        self.log("━" * 75, "INFO")
        self.log("✓ Sistema listo para renombrar", "SUCCESS")
        self.update_status("Listo para aplicar cambios", COLOR_SUCCESS)

    def update_preview(self):
        """Actualiza la vista previa del nuevo nombre."""
        if not self.system_info:
            return

        site = self.site_var.get()
        serial = self.system_info["serial"]
        new_name = build_hostname(serial, site)

        if not new_name:
            self.preview_name.configure(
                text="ERROR: Nombre inválido", text_color=COLOR_ERROR
            )
            self.preview_validation.configure(
                text="✗ No se pudo generar un nombre válido", text_color=COLOR_ERROR
            )
            return

        self.preview_name.configure(text=new_name, text_color=COLOR_SUCCESS)

        current_name = self.system_info["current_name"]

        if current_name == new_name:
            self.preview_validation.configure(
                text="✓ El equipo ya tiene el nombre correcto (no se requieren cambios)",
                text_color=COLOR_SUCCESS,
            )
        else:
            action_text = f"✓ Cambio: '{current_name}' → '{new_name}'"
            if site == "PQN":
                action_text += f" + Unión a dominio '{DOMAIN_NAME}'"
            self.preview_validation.configure(
                text=action_text, text_color=COLOR_WARNING
            )

    def export_log(self):
        """Exporta el log a un archivo en el escritorio."""
        try:
            desktop = Path.home() / "Desktop"
            export_path = (
                desktop / f"Renamer_Log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
            )

            log_content = self.text_log.get("1.0", "end")

            with open(export_path, "w", encoding="utf-8") as f:
                f.write(
                    f"═══════════════════════════════════════════════════════════════\n"
                )
                f.write(f"  LOG DEL RENOMBRADOR MULTI-SITE PQN/CCS/CBQ\n")
                f.write(
                    f"═══════════════════════════════════════════════════════════════\n"
                )
                f.write(f"Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"Versión: {APP_VERSION}\n")
                f.write(f"Autor: {__author__}\n")
                f.write(f"Empresa: {__company__}\n")
                f.write(
                    f"═══════════════════════════════════════════════════════════════\n\n"
                )
                f.write(log_content)
                f.write(
                    f"\n═══════════════════════════════════════════════════════════════\n"
                )
                f.write(f"Fin del log\n")
                f.write(
                    f"═══════════════════════════════════════════════════════════════\n"
                )

            self.log(f"✓ Log exportado a: {export_path}", "SUCCESS")
            messagebox.showinfo(
                "✓ Log Exportado",
                f"El log ha sido exportado correctamente:\n\n{export_path}",
            )
        except Exception as e:
            self.log(f"✗ Error al exportar log: {str(e)}", "ERROR")
            messagebox.showerror("Error", f"No se pudo exportar el log:\n\n{str(e)}")

    def on_execute(self):
        """Maneja la ejecución del proceso de renombrado."""
        if self.is_processing:
            self.log("⚠ Ya hay un proceso en ejecución", "WARNING")
            return

        if not self.system_info:
            messagebox.showerror(
                "Error",
                "No hay información del sistema disponible.\n"
                "Por favor, reinicie la aplicación.",
            )
            return

        site = self.site_var.get()
        serial = self.system_info["serial"]
        new_name = build_hostname(serial, site)
        current_name = self.system_info["current_name"]

        if not new_name:
            messagebox.showerror(
                "Error de Validación",
                "No se pudo generar un nombre válido para el equipo.\n\n"
                "Contacte al administrador del sistema.",
            )
            return

        # Verificar si ya tiene el nombre correcto
        if current_name == new_name and site != "PQN":
            messagebox.showinfo(
                "Sin Cambios",
                "El equipo ya tiene el nombre correcto.\n\n" "No se requieren cambios.",
            )
            return

        # Confirmar acción con el usuario
        confirmation_msg = f"¿Confirma que desea aplicar los siguientes cambios?\n\n"
        confirmation_msg += f"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
        confirmation_msg += f"📌 CAMBIOS A REALIZAR:\n\n"

        if current_name != new_name:
            confirmation_msg += f"• Nombre actual:  {current_name}\n"
            confirmation_msg += f"• Nombre nuevo:   {new_name}\n\n"
        else:
            confirmation_msg += f"• Nombre:         {new_name} (sin cambios)\n\n"

        if site == "PQN":
            confirmation_msg += f"• Acción adicional: Unir al dominio '{DOMAIN_NAME}'\n"
            confirmation_msg += f"• Usuario del dominio: {DOMAIN_USER}\n\n"

        confirmation_msg += f"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
        confirmation_msg += f"⚠️ IMPORTANTE:\n"
        confirmation_msg += (
            f"• El equipo se reiniciará automáticamente en 15 segundos\n"
        )
        confirmation_msg += f"• Guarde todo su trabajo antes de continuar\n"
        confirmation_msg += f"• Este proceso no se puede deshacer fácilmente\n\n"
        confirmation_msg += f"¿Desea continuar?"

        response = messagebox.askyesno(
            "⚠️ Confirmar Cambios Críticos", confirmation_msg, icon="warning"
        )

        if not response:
            self.log("✗ Operación cancelada por el usuario", "WARNING")
            return

        # Iniciar proceso en hilo separado
        self.is_processing = True
        self.btn_execute.configure(
            state="disabled", text="⏳ Procesando cambios...", fg_color=COLOR_BG_MEDIUM
        )
        self.update_status("Ejecutando cambios...", COLOR_WARNING)

        thread = threading.Thread(
            target=self.execute_rename, args=(new_name, site), daemon=True
        )
        thread.start()

    def execute_rename(self, new_name, site):
        """Ejecuta el proceso completo de renombrado y unión a dominio."""
        try:
            self.log("━" * 75, "INFO")
            self.log("🚀 INICIANDO PROCESO DE RENOMBRADO", "PROCESS")
            self.log("━" * 75, "INFO")

            time.sleep(1)

            # Paso 1: Crear backup de configuración
            self.log("[1/4] Creando backup de configuración actual...", "PROCESS")
            backup_data = {
                "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "old_name": self.system_info["current_name"],
                "new_name": new_name,
                "site": site,
                "serial": self.system_info["serial"],
                "manufacturer": self.system_info["manufacturer"],
                "model": self.system_info["model"],
                "old_domain": self.system_info["current_domain"],
            }

            if save_backup(backup_data):
                self.log(f"      ✓ Backup guardado en: {BACKUP_FILE}", "SUCCESS")
                self.log(
                    f"      ✓ SHA-256: {backup_data.get('sha256', 'N/A')[:32]}...",
                    "SUCCESS",
                )
            else:
                self.log("      ⚠ No se pudo crear backup (continuando)", "WARNING")

            time.sleep(0.5)

            # Paso 2: Determinar si es necesario renombrar
            self.log("[2/4] Validando cambios necesarios...", "PROCESS")

            current_name = self.system_info["current_name"]
            need_rename = current_name.upper() != new_name.upper()
            need_domain = site == "PQN" and not self.system_info["in_domain"]

            if need_rename:
                self.log(
                    f"      ✓ Cambio de nombre requerido: {current_name} → {new_name}",
                    "SUCCESS",
                )
            else:
                self.log(f"      ℹ El nombre ya es correcto: {new_name}", "INFO")

            if need_domain:
                self.log(f"      ✓ Unión a dominio requerida: {DOMAIN_NAME}", "SUCCESS")

            if not need_rename and not need_domain:
                self.log("      ℹ No hay cambios que aplicar", "INFO")
                raise Exception("El equipo ya tiene la configuración solicitada")

            time.sleep(0.5)

            # Paso 3: Aplicar cambios
            self.log("[3/4] Aplicando cambios en el sistema...", "PROCESS")

            if site == "PQN" and need_domain:
                # Para PQN: Renombrar Y unir al dominio en una sola operación
                self.log(
                    f"      → Uniendo al dominio '{DOMAIN_NAME}' con nombre '{new_name}'...",
                    "INFO",
                )

                success = join_domain(
                    DOMAIN_NAME, DOMAIN_USER, DOMAIN_PASSWORD, new_name
                )

                if not success:
                    raise Exception("Error al unir el equipo al dominio")

                self.log("      ✓ Equipo unido al dominio exitosamente", "SUCCESS")
                self.log(f"      ✓ Nombre aplicado: {new_name}", "SUCCESS")

            else:
                # Para CCS/CBQ: Solo renombrar (sin dominio)
                if need_rename:
                    self.log(f"      → Aplicando nuevo nombre: {new_name}...", "INFO")

                    success = rename_computer_modern(new_name)

                    if not success:
                        raise Exception("Error al aplicar el nuevo nombre")

                    self.log("      ✓ Nombre aplicado correctamente", "SUCCESS")

            time.sleep(0.5)

            # Paso 4: Preparar reinicio
            self.log("[4/4] Preparando reinicio del sistema...", "PROCESS")
            self.log("      ⏳ El equipo se reiniciará en 15 segundos", "WARNING")
            self.update_status("Reiniciando en 15 segundos...", COLOR_WARNING)

            # Countdown visual
            for i in range(15, 0, -1):
                self.log(
                    f"      → Reiniciando en {i} segundo{'s' if i > 1 else ''}...",
                    "INFO",
                )
                time.sleep(1)

            self.log("━" * 75, "INFO")
            self.log("✓ PROCESO COMPLETADO EXITOSAMENTE", "SUCCESS")
            self.log("━" * 75, "INFO")
            self.log("🔄 Reiniciando equipo ahora...", "PROCESS")

            # Mostrar mensaje de éxito
            self.after(
                100,
                lambda: messagebox.showinfo(
                    "✓ Cambios Aplicados Exitosamente",
                    f"Los cambios han sido aplicados correctamente:\n\n"
                    f"✓ Nombre: {new_name}\n"
                    f"✓ Sitio: {site}-COL\n"
                    + (f"✓ Dominio: {DOMAIN_NAME}\n" if site == "PQN" else "")
                    + f"\nEl equipo se está reiniciando...",
                    icon="info",
                ),
            )

            # Ejecutar reinicio
            time.sleep(1)
            restart_computer(0)

            self.update_status("Reinicio iniciado", COLOR_SUCCESS)

        except Exception as e:
            self.log("━" * 75, "ERROR")
            self.log(f"✗ ERROR EN EL PROCESO: {str(e)}", "ERROR")
            self.log("━" * 75, "ERROR")
            self.update_status("Error en el proceso", COLOR_ERROR)

            self.after(
                100,
                lambda: messagebox.showerror(
                    "Error en Renombrado",
                    f"No se pudieron aplicar los cambios:\n\n{str(e)}\n\n"
                    "Posibles causas:\n"
                    "• Permisos insuficientes\n"
                    "• El equipo ya está en un dominio\n"
                    "• Credenciales de dominio incorrectas\n"
                    "• Restricciones de políticas de grupo\n"
                    "• Problemas de conectividad con el dominio\n\n"
          