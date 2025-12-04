"""
PQN_Access_Credentials.py - Versión Profesional Mejorada
Autor: Josué Romero
Empresa: Stefanini / PQN
Fecha: 06/Noviembre/2025
Versión: 3.0 Professional Edition

Licencia: Propiedad de Stefanini / PQN - Todos los derechos reservados
Copyright © 2025 Josué Romero - Stefanini / PQN

Descripción:
Sistema automatizado para generación de credenciales corporativas de acceso
con tabla profesional y logos corporativos integrados en PDF.
"""

import os
import platform
import customtkinter as ctk
from tkinter import messagebox
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from pathlib import Path
import re
import random
import string
from datetime import datetime
import hashlib
import sys
import ctypes
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.utils import formatdate


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

APP_TITLE = "Generador de Credenciales PQN"
APP_VERSION = f"v{__version__}"
APP_SIZE = "700x950"

# Paleta de colores profesional (Morado-Púrpura-Oscuro)
COLOR_PRIMARY = "#6a1b9a"  # Morado profundo
COLOR_SECONDARY = "#8e24aa"  # Morado medio
COLOR_ACCENT = "#ab47bc"  # Morado claro
COLOR_SUCCESS = "#00e676"  # Verde éxito
COLOR_WARNING = "#ffc107"  # Amarillo advertencia
COLOR_ERROR = "#d32f2f"  # Rojo error
COLOR_BG_DARK = "#0a0a0a"  # Negro profundo
COLOR_BG_MEDIUM = "#1a1a1a"  # Negro medio
COLOR_BG_LIGHT = "#2d2d2d"  # Gris oscuro
COLOR_TEXT_WHITE = "#ffffff"  # Texto blanco
COLOR_TEXT_GRAY = "#b0bec5"  # Texto gris claro

# Fuentes (Incrementadas +2px según requerimiento)
FONT_TITLE = ("Segoe UI", 28, "bold")  # Era 26
FONT_SUBTITLE = ("Segoe UI", 13)  # Era 11
FONT_BUTTON = ("Segoe UI", 14, "bold")  # Era 12
FONT_LABEL = ("Segoe UI", 13, "bold")  # Era 11
FONT_INFO = ("Segoe UI", 13)  # Era 11
FONT_CONSOLE = ("Consolas", 12)  # Era 10

# Constantes
DOMAIN = "@spradling.group"


def resource_path(relative_path):
    """Obtiene la ruta de un recurso tanto si se ejecuta como .py o .exe"""
    try:
        base_path = sys._MEIPASS  # carpeta temporal del .exe
    except Exception:
        base_path = os.path.abspath(".")  # cuando ejecutas .py

    return os.path.join(base_path, relative_path)


# Rutas de logos
LOGO_PROQUINAL = Path(resource_path("static/logos/proquinal.png"))
LOGO_MAYTE = Path(resource_path("static/logos/mayte.png"))
LOGO_STEFANINI = Path(resource_path("static/logos/stefanini.png"))


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


def calculate_file_hash(filepath):
    """Calcula el hash SHA-256 de un archivo."""
    try:
        sha256_hash = hashlib.sha256()
        with open(filepath, "rb") as f:
            for byte_block in iter(lambda: f.read(4096), b""):
                sha256_hash.update(byte_block)
        return sha256_hash.hexdigest()
    except:
        return "N/A"


def validate_username(username):
    """Valida formato de nombre de usuario."""
    if not username:
        return False, "El usuario no puede estar vacío"

    if not re.match(r"^[a-z0-9\-]+$", username):
        return False, "Solo letras minúsculas, números y guiones"

    if len(username) < 3:
        return False, "Mínimo 3 caracteres"

    if username.startswith("-") or username.endswith("-"):
        return False, "No puede empezar ni terminar con guión"

    return True, ""


def validate_email_prefix(email_prefix):
    """Valida prefijo de correo."""
    if not email_prefix:
        return False, "El correo no puede estar vacío"

    if not re.match(r"^[a-z0-9\.\-]+$", email_prefix):
        return False, "Solo letras minúsculas, números, puntos y guiones"

    if len(email_prefix) < 3:
        return False, "Mínimo 3 caracteres"

    if email_prefix.startswith(".") or email_prefix.endswith("."):
        return False, "No puede empezar ni terminar con punto"

    if ".." in email_prefix:
        return False, "No puede tener puntos consecutivos"

    return True, ""


def validate_password(password):
    """Valida que la contraseña cumpla con requisitos mínimos."""
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

    if not any(c in "!@#$%^&*()_+-=[]{}|;:,.<>?" for c in password):
        return False, "Debe contener al menos un carácter especial"

    return True, ""


def generate_secure_password(length=12):
    """Genera una contraseña segura aleatoria."""
    password = [
        random.choice(string.ascii_uppercase),
        random.choice(string.ascii_lowercase),
        random.choice(string.digits),
        random.choice("!@#$%^&*()_+-="),
    ]

    all_chars = string.ascii_letters + string.digits + "!@#$%^&*()_+-="
    password += [random.choice(all_chars) for _ in range(length - 4)]

    random.shuffle(password)

    return "".join(password)


def open_pdf(path):
    """Abre un PDF con el visor predeterminado del sistema."""
    sistema = platform.system()
    try:
        if sistema == "Windows":
            os.startfile(path)
        elif sistema == "Darwin":
            os.system(f"open '{path}'")
        else:
            os.system(f"xdg-open '{path}'")
        return True
    except:
        return False


# ============================================================================
# CLASE PRINCIPAL DE LA APLICACIÓN
# ============================================================================


class CredencialesApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Configuración de ventana
        self.title(f"{APP_TITLE} {APP_VERSION}")
        self.geometry(APP_SIZE)
        self.resizable(True, True)

        # Variables de estado
        self.password_visible = False
        self.is_generating = False

        # Construir interfaz
        self.build_ui()

    def build_ui(self):
        """Construye la interfaz de usuario moderna y profesional con scroll."""

        # Marco principal con scroll
        main_scrollable = ctk.CTkScrollableFrame(
            self,
            fg_color=COLOR_BG_DARK,
            scrollbar_button_color=COLOR_PRIMARY,
            scrollbar_button_hover_color=COLOR_SECONDARY,
        )
        main_scrollable.pack(fill="both", expand=True, padx=10, pady=10)

        # === ENCABEZADO MODERNO ===
        header_frame = ctk.CTkFrame(
            main_scrollable,
            fg_color=COLOR_PRIMARY,
            corner_radius=12,
            border_width=2,
            border_color=COLOR_ACCENT,
        )
        header_frame.pack(fill="x", pady=(0, 10))

        title_label = ctk.CTkLabel(
            header_frame,
            text="🔐 " + APP_TITLE,
            font=FONT_TITLE,
            text_color=COLOR_TEXT_WHITE,
        )
        title_label.pack(pady=(10, 3))

        subtitle_label = ctk.CTkLabel(
            header_frame,
            text=f"{__company__} | {__author__} | {APP_VERSION}",
            font=FONT_SUBTITLE,
            text_color=COLOR_ACCENT,
        )
        subtitle_label.pack(pady=(0, 5))

        copyright_label = ctk.CTkLabel(
            header_frame,
            text=__copyright__,
            font=("Segoe UI", 11),
            text_color=COLOR_TEXT_GRAY,
        )
        copyright_label.pack(pady=(0, 10))

        # === FORMULARIO ===
        form_frame = ctk.CTkFrame(
            main_scrollable, fg_color=COLOR_BG_MEDIUM, corner_radius=10
        )
        form_frame.pack(fill="x", pady=(0, 10))

        form_title = ctk.CTkLabel(
            form_frame,
            text="📝 Información de las Credenciales",
            font=FONT_LABEL,
            text_color=COLOR_TEXT_WHITE,
        )
        form_title.pack(pady=(10, 10), anchor="w", padx=15)

        # Correo corporativo
        ctk.CTkLabel(
            form_frame,
            text="Correo corporativo (sin @spradling.group):",
            font=FONT_INFO,
            text_color=COLOR_TEXT_WHITE,
        ).pack(pady=(0, 5), padx=15, anchor="w")

        correo_frame = ctk.CTkFrame(form_frame, fg_color="transparent")
        correo_frame.pack(pady=(0, 5), padx=15, anchor="w", fill="x")

        self.correo_entry = ctk.CTkEntry(
            correo_frame,
            placeholder_text="Ej: juan.perez",
            height=40,
            font=FONT_INFO,
            fg_color=COLOR_BG_LIGHT,
            text_color=COLOR_TEXT_WHITE,
            placeholder_text_color=COLOR_TEXT_GRAY,
            border_color=COLOR_SECONDARY,
            border_width=2,
        )
        self.correo_entry.pack(side="left", fill="x", expand=True)
        self.correo_entry.bind("<KeyRelease>", lambda e: self.validate_form())

        ctk.CTkLabel(
            correo_frame,
            text=DOMAIN,
            font=("Segoe UI", 13, "bold"),
            text_color=COLOR_PRIMARY,
        ).pack(side="left", padx=(5, 0))

        self.correo_validation = ctk.CTkLabel(
            form_frame, text="", font=("Segoe UI", 11), text_color=COLOR_WARNING
        )
        self.correo_validation.pack(pady=(0, 10), padx=15, anchor="w")

        # Contraseña
        ctk.CTkLabel(
            form_frame,
            text="Contraseña de acceso:",
            font=FONT_INFO,
            text_color=COLOR_TEXT_WHITE,
        ).pack(pady=(0, 5), padx=15, anchor="w")

        pass_frame = ctk.CTkFrame(form_frame, fg_color="transparent")
        pass_frame.pack(pady=(0, 5), padx=15, anchor="w", fill="x")

        self.pass_entry = ctk.CTkEntry(
            pass_frame,
            placeholder_text="Mínimo 8 caracteres",
            height=40,
            font=FONT_INFO,
            fg_color=COLOR_BG_LIGHT,
            text_color=COLOR_TEXT_WHITE,
            placeholder_text_color=COLOR_TEXT_GRAY,
            border_color=COLOR_SECONDARY,
            border_width=2,
            show="●",
        )
        self.pass_entry.pack(side="left", fill="x", expand=True)
        self.pass_entry.bind("<KeyRelease>", lambda e: self.validate_form())

        self.toggle_pass_btn = ctk.CTkButton(
            pass_frame,
            text="👁",
            command=self.toggle_password_visibility,
            width=40,
            height=40,
            font=("Segoe UI", 16),
            fg_color=COLOR_BG_LIGHT,
            hover_color=COLOR_BG_MEDIUM,
        )
        self.toggle_pass_btn.pack(side="left", padx=(5, 0))

        self.generate_pass_btn = ctk.CTkButton(
            pass_frame,
            text="🎲",
            command=self.generate_password,
            width=40,
            height=40,
            font=("Segoe UI", 16),
            fg_color=COLOR_PRIMARY,
            hover_color=COLOR_SECONDARY,
        )
        self.generate_pass_btn.pack(side="left", padx=(5, 0))

        self.pass_validation = ctk.CTkLabel(
            form_frame, text="", font=("Segoe UI", 11), text_color=COLOR_WARNING
        )
        self.pass_validation.pack(pady=(0, 10), padx=15, anchor="w")

        # Usuario de Windows/VPN
        ctk.CTkLabel(
            form_frame,
            text="Usuario de Windows/VPN:",
            font=FONT_INFO,
            text_color=COLOR_TEXT_WHITE,
        ).pack(pady=(0, 5), padx=15, anchor="w")

        self.user_entry = ctk.CTkEntry(
            form_frame,
            placeholder_text="Ej: perez-pepito",
            height=40,
            font=FONT_INFO,
            fg_color=COLOR_BG_LIGHT,
            text_color=COLOR_TEXT_WHITE,
            placeholder_text_color=COLOR_TEXT_GRAY,
            border_color=COLOR_SECONDARY,
            border_width=2,
        )
        self.user_entry.pack(pady=(0, 5), padx=15, anchor="w", fill="x")
        self.user_entry.bind("<KeyRelease>", lambda e: self.validate_form())
        self.user_entry.bind("<Return>", lambda e: self.generate_pdf())

        self.user_validation = ctk.CTkLabel(
            form_frame, text="", font=("Segoe UI", 11), text_color=COLOR_WARNING
        )
        self.user_validation.pack(pady=(0, 10), padx=15, anchor="w")

        # === VISTA PREVIA ===
        preview_frame = ctk.CTkFrame(
            main_scrollable, fg_color=COLOR_BG_MEDIUM, corner_radius=10
        )
        preview_frame.pack(fill="x", pady=(0, 10))

        preview_title = ctk.CTkLabel(
            preview_frame,
            text="👁 Vista Previa",
            font=FONT_LABEL,
            text_color=COLOR_TEXT_WHITE,
        )
        preview_title.pack(pady=(10, 5), anchor="w", padx=15)

        self.preview_text = ctk.CTkTextbox(
            preview_frame,
            height=220,
            font=FONT_CONSOLE,
            fg_color=COLOR_BG_LIGHT,
            text_color=COLOR_SUCCESS,
            border_width=2,
            border_color=COLOR_SECONDARY,
        )
        self.preview_text.pack(pady=(0, 10), padx=15, fill="x")
        self.preview_text.insert(
            "end", "Complete los campos para ver la vista previa...\n"
        )
        self.preview_text.configure(state="disabled")

        # === VALIDACIÓN GENERAL ===
        self.validation_general = ctk.CTkLabel(
            main_scrollable,
            text="",
            font=("Segoe UI", 12, "bold"),
            text_color=COLOR_WARNING,
        )
        self.validation_general.pack(pady=(0, 10))

        # === BOTONES ===
        button_frame = ctk.CTkFrame(main_scrollable, fg_color="transparent")
        button_frame.pack(fill="x", pady=(10, 0))

        self.generate_button = ctk.CTkButton(
            button_frame,
            text="🚀 Generar PDF de Credenciales",
            command=self.generate_pdf,
            font=FONT_BUTTON,
            height=50,
            corner_radius=10,
            fg_color=COLOR_PRIMARY,
            hover_color=COLOR_SECONDARY,
            border_width=3,
            border_color=COLOR_ACCENT,
            text_color=COLOR_TEXT_WHITE,
            state="disabled",
        )
        self.generate_button.pack(side="left", expand=True, fill="x", padx=(0, 5))

        self.clear_button = ctk.CTkButton(
            button_frame,
            text="🗑️ Limpiar Campos",
            command=self.clear_fields,
            font=FONT_BUTTON,
            height=50,
            corner_radius=10,
            fg_color=COLOR_BG_LIGHT,
            hover_color=COLOR_BG_MEDIUM,
            border_width=2,
            border_color=COLOR_SECONDARY,
            text_color=COLOR_TEXT_WHITE,
        )
        self.clear_button.pack(side="left", fill="x", padx=(5, 0))

    def toggle_password_visibility(self):
        """Alterna visibilidad de la contraseña."""
        if self.password_visible:
            self.pass_entry.configure(show="●")
            self.password_visible = False
            self.toggle_pass_btn.configure(text="👁")
        else:
            self.pass_entry.configure(show="")
            self.password_visible = True
            self.toggle_pass_btn.configure(text="🙈")

    def generate_password(self):
        """Genera una contraseña segura."""
        password = generate_secure_password(12)
        self.pass_entry.delete(0, "end")
        self.pass_entry.insert(0, password)
        self.validate_form()
        messagebox.showinfo(
            "Contraseña Generada",
            f"Se ha generado una contraseña segura:\n\n"
            f"🔑 {password}\n\n"
            f"⚠️ IMPORTANTE: Anótela antes de continuar.",
        )

    def clear_fields(self):
        """Limpia todos los campos del formulario."""
        self.correo_entry.delete(0, "end")
        self.pass_entry.delete(0, "end")
        self.user_entry.delete(0, "end")
        self.correo_validation.configure(text="")
        self.pass_validation.configure(text="")
        self.user_validation.configure(text="")
        self.validation_general.configure(text="")
        self.preview_text.configure(state="normal")
        self.preview_text.delete("1.0", "end")
        self.preview_text.insert(
            "end", "Complete los campos para ver la vista previa...\n"
        )
        self.preview_text.configure(state="disabled")
        self.correo_entry.focus()

    def validate_form(self):
        """Valida el formulario completo en tiempo real."""
        correo = self.correo_entry.get().strip()
        password = self.pass_entry.get().strip()
        usuario = self.user_entry.get().strip()

        all_valid = True

        # Validar correo
        if correo:
            is_valid, msg = validate_email_prefix(correo)
            if is_valid:
                self.correo_validation.configure(
                    text="✓ Válido", text_color=COLOR_SUCCESS
                )
            else:
                self.correo_validation.configure(
                    text=f"✗ {msg}", text_color=COLOR_ERROR
                )
                all_valid = False
        else:
            self.correo_validation.configure(text="")
            all_valid = False

        # Validar contraseña
        if password:
            is_valid, msg = validate_password(password)
            if is_valid:
                self.pass_validation.configure(
                    text="✓ Contraseña segura", text_color=COLOR_SUCCESS
                )
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
                self.user_validation.configure(
                    text="✓ Válido", text_color=COLOR_SUCCESS
                )
            else:
                self.user_validation.configure(text=f"✗ {msg}", text_color=COLOR_ERROR)
                all_valid = False
        else:
            self.user_validation.configure(text="")
            all_valid = False

        # Actualizar estado del botón y vista previa
        if all_valid:
            self.generate_button.configure(state="normal")
            self.validation_general.configure(
                text="✓ Listo para generar PDF", text_color=COLOR_SUCCESS
            )
            self.update_preview(correo, password, usuario)
        else:
            self.generate_button.configure(state="disabled")
            self.validation_general.configure(text="")

    def update_preview(self, correo, password, usuario):
        """Actualiza la vista previa con las credenciales."""
        preview = f"""
╔══════════════════════════════════════════════════════════════╗
║     CREDENCIALES DE ACCESO CORPORATIVO PQN                   ║
╚══════════════════════════════════════════════════════════════╝

📧 CORREO CORPORATIVO
Usuario:    {correo}{DOMAIN}
Contraseña: {password}

🖥️ ACCESO A LAS PLATAFORMAS
Usuario:    {usuario}
Contraseña: {password}
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
            messagebox.showwarning(
                "Campos Requeridos",
                "Por favor completa todos los campos antes de generar el PDF.",
            )
            return

        self.is_generating = True
        self.generate_button.configure(
            state="disabled", text="⏳ Generando PDF...", fg_color=COLOR_BG_MEDIUM
        )

        try:
            filename = "Credenciales_Acceso_PQN.pdf"

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

            # Generar PDF
            self.crear_pdf(str(ruta_docs), correo, password, usuario)

            # Calcular hash
            file_hash = calculate_file_hash(str(ruta_docs))

            # Copia en D:/
            if Path("D:/").exists():
                self.crear_pdf(str(ruta_datos), correo, password, usuario)

            # Abrir PDF
            if Path("D:/").exists() and ruta_datos.exists():
                open_pdf(str(ruta_datos))
            else:
                open_pdf(str(ruta_docs))

            # === Enviar correo automáticamente ===
            try:
                correo_destino = f"{correo}{DOMAIN}"
                enviado = enviar_correo_con_pdf(correo_destino, str(ruta_docs))
                if enviado:
                    print(f"Correo enviado correctamente a {correo_destino}")
                else:
                    print("No se pudo enviar el correo con el adjunto.")
            except Exception as e:
                print(f"Error al intentar enviar el correo: {e}")

            # Cerrar aplicación después de 2 segundos
            self.after(2000, self.quit)

        except Exception as e:
            messagebox.showerror(
                "Error en Generación",
                f"No se pudo generar el PDF:\n\n{str(e)}\n\n"
                "Posibles causas:\n"
                "• Sin permisos de escritura\n"
                "• Disco lleno",
            )
        finally:
            self.is_generating = False
            self.generate_button.configure(
                state="normal",
                text="🚀 Generar PDF de Credenciales",
                fg_color=COLOR_PRIMARY,
            )

    # Diseño del documento
    def crear_pdf(self, path, correo, password, usuario):
        """Crea el PDF de credenciales con logos corporativos y tabla profesional."""
        c = canvas.Canvas(path, pagesize=letter)
        width, height = letter

        # Colores RGB normalizado
        MORADO = (0.416, 0.106, 0.604)
        VERDE = (0.0, 0.902, 0.463)
        NEGRO = (0.043, 0.043, 0.043)
        GRIS_OSCURO = (0.290, 0.290, 0.290)
        GRIS_CLARO = (0.95, 0.95, 0.95)
        BLANCO = (1.0, 1.0, 1.0)

        # === LOGOS EN ENCABEZADO ===
        try:
            if LOGO_PROQUINAL.exists():
                c.drawImage(
                    str(LOGO_PROQUINAL),
                    50,
                    height - 80,
                    width=100,
                    height=40,
                    preserveAspectRatio=True,
                    mask="auto",
                )
            if LOGO_MAYTE.exists():
                c.drawImage(
                    str(LOGO_MAYTE),
                    width - 150,
                    height - 80,
                    width=100,
                    height=40,
                    preserveAspectRatio=True,
                    mask="auto",
                )
        except:
            pass

        # Línea divisoria
        c.setStrokeColorRGB(*MORADO)
        c.setLineWidth(2)
        c.line(50, height - 90, width - 50, height - 90)

        # === TÍTULO ===
        c.setFont("Helvetica-Bold", 18)
        c.setFillColorRGB(*MORADO)
        c.drawCentredString(
            width / 2, height - 120, "CREDENCIALES DE ACCESO CORPORATIVO"
        )

        c.setFont("Helvetica", 11)
        c.setFillColorRGB(*NEGRO)
        c.drawCentredString(
            width / 2,
            height - 138,
            "Departamento de Tecnologías de la Información",
        )

        # Fecha y hora
        fecha_actual = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        c.setFont("Helvetica-Oblique", 9)
        c.setFillColorRGB(*GRIS_OSCURO)
        c.drawRightString(width - 50, height - 155, f"Generado el: {fecha_actual}")

        # === INFORMACIÓN BÁSICA ===
        info_y = height - 190
        c.setFont("Helvetica-Bold", 11)
        c.setFillColorRGB(*MORADO)
        c.drawString(50, info_y, "📧 Correo Corporativo:")
        c.setFont("Helvetica", 11)
        c.setFillColorRGB(*NEGRO)
        c.drawString(200, info_y, f"{correo}{DOMAIN}")

        c.setFont("Helvetica-Bold", 11)
        c.setFillColorRGB(*MORADO)
        c.drawString(50, info_y - 20, "👤 Usuario de red:")
        c.setFont("Helvetica", 11)
        c.setFillColorRGB(*NEGRO)
        c.drawString(200, info_y - 20, usuario)

        c.setFont("Helvetica-Bold", 11)
        c.setFillColorRGB(*MORADO)
        c.drawString(50, info_y - 40, "🔑 Contraseña de Acceso:")
        c.setFont("Helvetica", 11)
        c.setFillColorRGB(*NEGRO)
        c.drawString(200, info_y - 40, password)

        # === TABLA DE ACCESOS ===
        tabla_y_start = info_y - 80

        # Título de la tabla
        c.setFont("Helvetica-Bold", 12)
        c.setFillColorRGB(*MORADO)
        c.drawString(50, tabla_y_start, "🌐 PLATAFORMAS Y ACCESOS DISPONIBLES")

        # Datos de la tabla (ejemplo - ajusta según tus necesidades)
        datos_tabla = [
            {
                "nombre": "Equipo / Windows",
                "url": "https://somebooks.es/wp-content/uploads/2024/12/Unir-un-cliente-Windows-11-a-un-dominio-Windows-Server-2022-015.png",
                "para_que": "Para iniciar sesión en el equipo",
                "usuario": usuario,
                "password": password,
            },
            {
                "nombre": "VPN FortiClient",
                "url": "https://www.beyaz.net/files/elfinder/content_photo/icerik_dosyalari/ssl-vpn-kullanimi-iki.png",
                "para_que": "Conexión virtual a red corporativa desde WiFi libre",
                "usuario": usuario,
                "password": password,
            },
            {
                "nombre": "Citrix Workspace",
                "url": "https://pixel-si.uz/wp-content/uploads/2023/01/tab2_ciscoworkspace.webp",
                "para_que": "Programas SpradlingGroup en un solo sitio para trabajo remoto",
                "usuario": usuario,
                "password": password,
            },
            {
                "nombre": "Daruma",
                "url": "https://proquinal.darumasoftware.com/app.php/staff",
                "para_que": "Sistema de gestión documental y procesos",
                "usuario": usuario,
                "password": password,
            },
            {
                "nombre": "Terranova",
                "url": "",
                "para_que": "Aula virtual de cursos cortos corporativos",
                "usuario": correo,
                "password": password,
            },
            {
                "nombre": "Intranet PQN",
                "url": "http://pqnintranet.proquinal.com/PEP-PORTAL-WEB/appmanager/intranet/proquinal_es",
                "para_que": "Recursos internos y flujos Proquinal S.A.S",
                "usuario": "carnet / cédula",
                "password": password,
            },
        ]

        # Configuración de la tabla
        tabla_x = 50
        tabla_y = tabla_y_start - 25

        # Anchos de columnas (total = 512px para mantener margen)
        col_widths = [90, 130, 110, 90, 92]
        col_x_positions = [tabla_x]
        for w in col_widths[:-1]:
            col_x_positions.append(col_x_positions[-1] + w)

        # Altura de filas
        header_height = 25
        row_height = 35

        # Calcular altura total de la tabla
        total_height = header_height + (len(datos_tabla) * row_height)

        # Fondo de la tabla
        c.setFillColorRGB(*BLANCO)
        c.setStrokeColorRGB(*MORADO)
        c.setLineWidth(1.5)
        c.roundRect(
            tabla_x,
            tabla_y - total_height,
            sum(col_widths),
            total_height,
            8,
            stroke=1,
            fill=1,
        )

        # === ENCABEZADO DE LA TABLA ===
        c.setFillColorRGB(*MORADO)
        c.rect(
            tabla_x,
            tabla_y - header_height,
            sum(col_widths),
            header_height,
            stroke=0,
            fill=1,
        )

        # Esquinas redondeadas superiores (simuladas)
        c.setFillColorRGB(*MORADO)
        c.roundRect(
            tabla_x,
            tabla_y - header_height,
            sum(col_widths),
            header_height,
            8,
            stroke=0,
            fill=1,
        )
        c.setFillColorRGB(*BLANCO)
        c.rect(
            tabla_x, tabla_y - header_height - 5, sum(col_widths), 5, stroke=0, fill=1
        )

        # Texto del encabezado
        headers = ["Nombre", "URL", "Para qué es", "Usuario", "Contraseña"]
        c.setFont("Helvetica-Bold", 9)
        c.setFillColorRGB(*BLANCO)

        for i, header in enumerate(headers):
            text_x = col_x_positions[i] + col_widths[i] / 2
            c.drawCentredString(text_x, tabla_y - header_height + 8, header)

        # Líneas verticales del encabezado
        c.setStrokeColorRGB(*BLANCO)
        c.setLineWidth(1)
        for x_pos in col_x_positions[1:]:
            c.line(x_pos, tabla_y - header_height, x_pos, tabla_y)

        # === FILAS DE DATOS ===
        current_y = tabla_y - header_height

        for idx, fila in enumerate(datos_tabla):
            # Alternar color de fondo
            if idx % 2 == 0:
                c.setFillColorRGB(*BLANCO)
            else:
                c.setFillColorRGB(*GRIS_CLARO)

            c.rect(
                tabla_x,
                current_y - row_height,
                sum(col_widths),
                row_height,
                stroke=0,
                fill=1,
            )

            # Líneas horizontales
            c.setStrokeColorRGB(*GRIS_OSCURO)
            c.setLineWidth(0.5)
            c.line(tabla_x, current_y, tabla_x + sum(col_widths), current_y)

            # Líneas verticales
            for x_pos in col_x_positions[1:]:
                c.line(x_pos, current_y, x_pos, current_y - row_height)

            # Contenido de las celdas
            c.setFont("Helvetica-Bold", 8)
            c.setFillColorRGB(*NEGRO)

            # Nombre (centrado)
            text_x = col_x_positions[0] + col_widths[0] / 2
            c.drawCentredString(text_x, current_y - row_height / 2 + 2, fila["nombre"])

            # URL (clickeable y centrada)
            c.setFont("Helvetica", 7)
            c.setFillColorRGB(0, 0, 1)  # Azul para URL
            url_text = fila["url"].replace("https://", "").replace("http://", "")
            if len(url_text) > 22:
                url_text = url_text[:19] + "..."
            text_x = col_x_positions[1] + col_widths[1] / 2

            # Crear link clickeable
            link_rect = (
                col_x_positions[1] + 5,
                current_y - row_height + 5,
                col_x_positions[1] + col_widths[1] - 5,
                current_y - 5,
            )
            c.linkURL(fila["url"], link_rect, relative=0)
            c.drawCentredString(text_x, current_y - row_height / 2 + 2, url_text)

            # Para qué es (centrado)
            c.setFont("Helvetica", 8)
            c.setFillColorRGB(*NEGRO)
            text_x = col_x_positions[2] + col_widths[2] / 2
            c.drawCentredString(
                text_x, current_y - row_height / 2 + 2, fila["para_que"]
            )

            # Usuario (centrado)
            c.setFont("Helvetica", 7)
            text_x = col_x_positions[3] + col_widths[3] / 2
            usuario_text = fila["usuario"]
            if len(usuario_text) > 12:
                usuario_text = usuario_text[:9] + "..."
            c.drawCentredString(text_x, current_y - row_height / 2 + 2, usuario_text)

            # Contraseña (centrado y enmascarada)
            text_x = col_x_positions[4] + col_widths[4] / 2
            password_masked = "•" * min(len(fila["password"]), 10)
            c.drawCentredString(text_x, current_y - row_height / 2 + 2, password_masked)

            current_y -= row_height

        # Borde final de la tabla
        c.setStrokeColorRGB(*MORADO)
        c.setLineWidth(1.5)
        c.line(tabla_x, current_y, tabla_x + sum(col_widths), current_y)

        # === NOTAS DE SEGURIDAD ===
        nota_y = current_y - 30
        c.setFillColorRGB(*GRIS_OSCURO)
        c.setFont("Helvetica-Bold", 11)
        c.drawString(50, nota_y, "⚠️ IMPORTANTE:")

        c.setFont("Helvetica", 10)
        c.setFillColorRGB(*NEGRO)
        lineas = [
            "• Estas credenciales son personales e intransferibles.",
            "• No comparta su contraseña con ningún compañero (solo IT).",
            "• La contraseña debe cambiarse cada 28 días.",
            "• En caso de pérdida o amenaza digital, informe a Soporte IT inmediatamente.",
        ]

        y = nota_y - 20
        for linea in lineas:
            c.drawString(70, y, linea)
            y -= 14

        # === PIE DE PÁGINA ===
        c.setStrokeColorRGB(*MORADO)
        c.line(50, 80, width - 50, 80)

        c.setFont("Helvetica-Oblique", 9)
        c.setFillColorRGB(*GRIS_OSCURO)
        c.drawCentredString(
            width / 2,
            65,
            "Documento confidencial - Uso exclusivo del destinatario",
        )
        c.drawCentredString(
            width / 2, 52, "© 2025 PQN-COL | Todos los derechos reservados"
        )

        # === FIRMA DIGITAL (HASH) ===
        hash_text = calculate_file_hash(path)
        c.setFont("Courier", 8)
        c.setFillColorRGB(*GRIS_OSCURO)
        c.drawString(50, 35, f"SHA-256: {hash_text[:64]}")
        if len(hash_text) > 64:
            c.drawString(50, 25, hash_text[64:])

            # Logo Stefanini marca de agua (esquina inferior derecha - pequeño)
        try:
            if LOGO_STEFANINI.exists():
                c.saveState()
                c.setFillAlpha(0.2)  # Transparencia para marca de agua
                c.drawImage(
                    str(LOGO_STEFANINI),
                    width - 130,
                    20,
                    width=80,
                    height=30,
                    preserveAspectRatio=True,
                    mask="auto",
                )
                c.restoreState()
        except:
            pass

        # Finalizar PDF
        c.showPage()
        c.save()


# ============================================================================
# MÉTODO PRINCIPAL (ENTRY POINT)
# ============================================================================


def main():
    """Punto de entrada principal del programa."""
    try:
        # Verificar privilegios de administrador
        if platform.system() == "Windows" and not is_admin():
            respuesta = messagebox.askyesno(
                "Permisos Requeridos",
                "Oye pelao, este programa requiere privilegios de administrador para generar archivos.\n\n"
                "¿Desea reiniciarlo con permisos elevados?",
            )
            if respuesta:
                run_as_admin()
            else:
                messagebox.showinfo(
                    "Ejecución Cancelada",
                    "La aplicación no puede continuar sin permisos administrativos.",
                )
                sys.exit(0)

        # Iniciar aplicación
        app = CredencialesApp()
        app.mainloop()

    except KeyboardInterrupt:
        print("\nEjecución interrumpida por el usuario.")
        sys.exit(0)

    except Exception as e:
        messagebox.showerror("Error Fatal", f"Ocurrió un error inesperado:\n\n{str(e)}")
        sys.exit(1)


# ============================================================================
# FUNCIÓN PARA ENVÍO DE CORREO
# ============================================================================


def enviar_correo_con_pdf(destinatario, ruta_pdf):
    """
    Envía el PDF generado al correo corporativo indicado por el usuario.
    Usa SMTP seguro (TLS).
    """
    # Configuración del servidor de correo corporativo
    SMTP_SERVER = "smtp.office365.com"
    SMTP_PORT = 587
    SMTP_USER = "mayte@spradling.group"
    SMTP_PASS = "test"

    try:
        # Crear mensaje
        msg = MIMEMultipart()
        msg["From"] = SMTP_USER
        msg["To"] = destinatario
        msg["Date"] = formatdate(localtime=True)
        msg["Subject"] = "Tus Credenciales de Acceso Corporativo"

        cuerpo_html = f"""
      <html>
      <body style="font-size:13px; font-family:Segoe UI; color:#222;">
         <p>Buen día, estimado/a usuario,</p>
         <p>Adjunto encontrará un documento con sus credenciales de acceso corporativo generadas automáticamente por el sistema.</p>
         <p><b>Correo:</b> {destinatario}<br>
            <b>Fecha:</b> {datetime.now().strftime("%d/%m/%Y %H:%M")}</p>
         <p>Por favor guarde este documento en un lugar seguro ahora.</p>
         <hr>
         <p style="font-size:11x; color:#777;">Este correo fue enviado automáticamente. No debes responder este mensaje.<br>
         © 2025 Equipo Mesa de Ayuda Mayté - Todos los derechos reservados.</p>
      </body>
      </html>
      """
        msg.attach(MIMEText(cuerpo_html, "html"))

        # Adjuntar el PDF
        with open(ruta_pdf, "rb") as f:
            part = MIMEApplication(f.read(), Name=os.path.basename(ruta_pdf))
        part["Content-Disposition"] = (
            f'attachment; filename="{os.path.basename(ruta_pdf)}"'
        )
        msg.attach(part)

        # Envío del correo
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_USER, SMTP_PASS)
            server.send_message(msg)

        print(f"✅ Correo enviado exitosamente a [ {destinatario} ]")
        return True

    except Exception as e:
        print(f"❌ Error al enviar correo: {e}")
        return False


# ============================================================================
# EJECUCIÓN DEL SCRIPT
# ============================================================================

if __name__ == "__main__":
    main()
