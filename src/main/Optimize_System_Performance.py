"""
Optimize_System_Performance.py - Versión Mejorada Extendida
Autor: Josué Romero
Empresa: Stefanini / PQN
Fecha: 31/Octubre/2025
Modificado: Diciembre/2025

Descripción:
Optimiza el rendimiento del equipo ejecutando varias herramientas clave del sistema.
Incluye nuevas optimizaciones para Windows 11 24H2.
"""

import sys
import ctypes
import subprocess
import threading
import time
import customtkinter as ctk
from tkinter import messagebox
from datetime import datetime

# ============================================================================
# CONFIGURACIÓN GLOBAL
# ============================================================================
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

APP_TITLE = "Optimizador de Sistema"
APP_VERSION = "v3.5"
APP_SIZE = "800x900"

# Colores
COLOR_PRIMARY = "#42a5f5"
COLOR_SUCCESS = "#66bb6a"
COLOR_WARNING = "#ffa726"
COLOR_ERROR = "#ef5350"
COLOR_BG_DARK = "#1a1a1a"
COLOR_BG_LIGHT = "#2d2d2d"
COLOR_TEXT = "#e0e0e0"

# Fuentes
FONT_TITLE = ("Segoe UI", 26, "bold")
FONT_SUBTITLE = ("Segoe UI", 11)
FONT_CONSOLE = ("Consolas", 9)
FONT_BUTTON = ("Segoe UI", 12, "bold")
FONT_LABEL = ("Segoe UI", 11, "bold")


# ============================================================================
# DEFINICIÓN DE TAREAS DE OPTIMIZACIÓN
# ============================================================================

OPTIMIZATION_TASKS = [
    {
        "id": "cleanmgr",
        "name": "Limpieza de Disco",
        "description": "Elimina archivos basura del sistema",
        "command": "cleanmgr /verylowdisk /sagerun:1",
        "estimated_time": "2-6 min",
        "enabled": True,
        "critical": False,
        "category": "basic",
    },
    {
        "id": "defrag_c",
        "name": "Optimizar Disco C:",
        "description": "Desfragmenta y optimiza el disco principal",
        "command": "defrag C: /O /H",
        "estimated_time": "5-10 min",
        "enabled": True,
        "critical": False,
        "category": "basic",
    },
    {
        "id": "defrag_d",
        "name": "Optimizar Disco D:",
        "description": "Desfragmenta y optimiza el disco secundario",
        "command": "defrag D: /O /H",
        "estimated_time": "5-10 min",
        "enabled": False,
        "critical": False,
        "category": "basic",
    },
    {
        "id": "temp_files",
        "name": "Limpiar Archivos Temporales",
        "description": "Elimina temporales de Windows y usuario",
        "command": [
            'powershell Remove-Item -Path "$env:TEMP\\*" -Recurse -Force -ErrorAction SilentlyContinue',
            'powershell Remove-Item -Path "C:\\Windows\\Temp\\*" -Recurse -Force -ErrorAction SilentlyContinue',
            'powershell Remove-Item "C:\\Windows\\Prefetch\\*" -Force -ErrorAction SilentlyContinue',
            "powershell Clear-RecycleBin -Force -ErrorAction SilentlyContinue",
        ],
        "estimated_time": "1-2 min",
        "enabled": True,
        "critical": False,
        "category": "basic",
    },
    {
        "id": "sfc",
        "name": "Reparar Archivos del Sistema (SFC)",
        "description": "Verifica y repara archivos corruptos de Windows",
        "command": "sfc /scannow",
        "estimated_time": "10-20 min",
        "enabled": True,
        "critical": True,
        "category": "basic",
    },
    {
        "id": "dism_scan",
        "name": "Escanear Imagen del Sistema (DISM)",
        "description": "Escanea la integridad de la imagen de Windows",
        "command": "DISM /Online /Cleanup-Image /ScanHealth",
        "estimated_time": "5-10 min",
        "enabled": True,
        "critical": True,
        "category": "basic",
    },
    {
        "id": "dism_restore",
        "name": "Reparar Imagen del Sistema (DISM)",
        "description": "Repara la imagen de Windows si hay errores",
        "command": "DISM /Online /Cleanup-Image /RestoreHealth",
        "estimated_time": "10-30 min",
        "enabled": True,
        "critical": True,
        "category": "basic",
    },
    {
        "id": "winget_update",
        "name": "Actualizar Programas (Winget)",
        "description": "Actualiza los programas pendientes excepto Java 8-341",
        "command": "powershell",
        "estimated_time": "5-15 min",
        "enabled": True,
        "critical": False,
        "category": "basic",
    },
    # ============ NUEVAS TAREAS DE OPTIMIZACIÓN AVANZADA ============
    {
        "id": "disable_telemetry",
        "name": "Desactivar Telemetría de Windows",
        "description": "Deshabilita servicios de recopilación de datos",
        "command": [
            'reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\DataCollection" /v AllowTelemetry /t REG_DWORD /d 0 /f',
            'reg add "HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\DataCollection" /v AllowTelemetry /t REG_DWORD /d 0 /f',
            "sc config DiagTrack start= disabled",
            "sc stop DiagTrack",
            "sc config dmwappushservice start= disabled",
            "sc stop dmwappushservice",
        ],
        "estimated_time": "30 seg",
        "enabled": False,
        "critical": False,
        "category": "privacy",
    },
    {
        "id": "disable_cortana",
        "name": "Desactivar Cortana",
        "description": "Deshabilita el asistente Cortana",
        "command": [
            'reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\Windows Search" /v AllowCortana /t REG_DWORD /d 0 /f',
            'reg add "HKLM\\SOFTWARE\\Microsoft\\PolicyManager\\default\\Experience\\AllowCortana" /v value /t REG_DWORD /d 0 /f',
        ],
        "estimated_time": "10 seg",
        "enabled": False,
        "critical": False,
        "category": "privacy",
    },
    {
        "id": "disable_windows_ink",
        "name": "Desactivar Windows Ink",
        "description": "Deshabilita el área de trabajo de Windows Ink",
        "command": 'reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\WindowsInkWorkspace" /v AllowWindowsInkWorkspace /t REG_DWORD /d 0 /f',
        "estimated_time": "5 seg",
        "enabled": False,
        "critical": False,
        "category": "performance",
    },
    {
        "id": "disable_visual_effects",
        "name": "Optimizar Efectos Visuales",
        "description": "Configura efectos visuales para mejor rendimiento",
        "command": [
            'reg add "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\VisualEffects" /v VisualFXSetting /t REG_DWORD /d 2 /f',
            'reg add "HKCU\\Control Panel\\Desktop" /v UserPreferencesMask /t REG_BINARY /d 9012038010000000 /f',
            'reg add "HKCU\\Control Panel\\Desktop\\WindowMetrics" /v MinAnimate /t REG_SZ /d 0 /f',
            'reg add "HKCU\\Software\\Microsoft\\Windows\\DWM" /v EnableAeroPeek /t REG_DWORD /d 0 /f',
        ],
        "estimated_time": "15 seg",
        "enabled": False,
        "critical": False,
        "category": "performance",
    },
    {
        "id": "disable_startup_delay",
        "name": "Eliminar Retraso de Inicio",
        "description": "Reduce el tiempo de carga de programas al inicio",
        "command": 'reg add "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Serialize" /v StartupDelayInMSec /t REG_DWORD /d 0 /f',
        "estimated_time": "5 seg",
        "enabled": False,
        "critical": False,
        "category": "performance",
    },
    {
        "id": "disable_hibernation",
        "name": "Desactivar Hibernación",
        "description": "Libera espacio en disco (hiberfil.sys)",
        "command": "powercfg -h off",
        "estimated_time": "10 seg",
        "enabled": False,
        "critical": False,
        "category": "performance",
    },
    {
        "id": "optimize_power_plan",
        "name": "Configurar Plan de Energía Alto Rendimiento",
        "description": "Activa el plan de máximo rendimiento",
        "command": [
            "powercfg -duplicatescheme e9a42b02-d5df-448d-aa00-03f14749eb61",
            "powercfg -setactive 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c",
        ],
        "estimated_time": "10 seg",
        "enabled": False,
        "critical": False,
        "category": "performance",
    },
    {
        "id": "disable_windows_search",
        "name": "Desactivar Indexación de Windows Search",
        "description": "Reduce uso de disco y CPU",
        "command": [
            "sc config WSearch start= disabled",
            "sc stop WSearch",
        ],
        "estimated_time": "15 seg",
        "enabled": False,
        "critical": False,
        "category": "performance",
    },
    {
        "id": "disable_superfetch",
        "name": "Desactivar SysMain (Superfetch)",
        "description": "Útil para SSDs, reduce carga del sistema",
        "command": [
            "sc config SysMain start= disabled",
            "sc stop SysMain",
        ],
        "estimated_time": "10 seg",
        "enabled": False,
        "critical": False,
        "category": "performance",
    },
    {
        "id": "disable_windows_tips",
        "name": "Desactivar Consejos de Windows",
        "description": "Elimina notificaciones de sugerencias",
        "command": 'reg add "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\ContentDeliveryManager" /v SubscribedContent-338389Enabled /t REG_DWORD /d 0 /f',
        "estimated_time": "5 seg",
        "enabled": False,
        "critical": False,
        "category": "privacy",
    },
    {
        "id": "disable_activity_history",
        "name": "Desactivar Historial de Actividades",
        "description": "Desactiva el seguimiento de actividades",
        "command": [
            'reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\System" /v EnableActivityFeed /t REG_DWORD /d 0 /f',
            'reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\System" /v PublishUserActivities /t REG_DWORD /d 0 /f',
            'reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\System" /v UploadUserActivities /t REG_DWORD /d 0 /f',
        ],
        "estimated_time": "10 seg",
        "enabled": False,
        "critical": False,
        "category": "privacy",
    },
    {
        "id": "disable_transparency",
        "name": "Desactivar Transparencia",
        "description": "Mejora rendimiento gráfico",
        "command": 'reg add "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize" /v EnableTransparency /t REG_DWORD /d 0 /f',
        "estimated_time": "5 seg",
        "enabled": False,
        "critical": False,
        "category": "performance",
    },
    {
        "id": "clean_winsxs",
        "name": "Limpiar WinSxS",
        "description": "Reduce el tamaño de la carpeta de componentes",
        "command": "DISM /Online /Cleanup-Image /StartComponentCleanup /ResetBase",
        "estimated_time": "5-10 min",
        "enabled": False,
        "critical": False,
        "category": "maintenance",
    },
    {
        "id": "optimize_network",
        "name": "Optimizar Configuración de Red",
        "description": "Mejora la latencia y velocidad de red",
        "command": [
            "netsh int tcp set global autotuninglevel=normal",
            "netsh int tcp set global chimney=enabled",
            "netsh int tcp set global dca=enabled",
            "netsh int tcp set global netdma=enabled",
            "netsh int tcp set heuristics disabled",
        ],
        "estimated_time": "10 seg",
        "enabled": False,
        "critical": False,
        "category": "performance",
    },
    {
        "id": "clear_dns_cache",
        "name": "Limpiar Caché DNS",
        "description": "Refresca la resolución de nombres",
        "command": "ipconfig /flushdns",
        "estimated_time": "5 seg",
        "enabled": False,
        "critical": False,
        "category": "maintenance",
    },
    {
        "id": "disable_game_bar",
        "name": "Desactivar Xbox Game Bar",
        "description": "Libera recursos para mejor rendimiento",
        "command": [
            'reg add "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\GameDVR" /v AppCaptureEnabled /t REG_DWORD /d 0 /f',
            'reg add "HKCU\\System\\GameConfigStore" /v GameDVR_Enabled /t REG_DWORD /d 0 /f',
        ],
        "estimated_time": "10 seg",
        "enabled": False,
        "critical": False,
        "category": "performance",
    },
    {
        "id": "disable_windows_update_delivery",
        "name": "Desactivar Entrega de Actualizaciones P2P",
        "description": "Evita compartir ancho de banda",
        "command": 'reg add "HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\DeliveryOptimization\\Config" /v DODownloadMode /t REG_DWORD /d 0 /f',
        "estimated_time": "5 seg",
        "enabled": False,
        "critical": False,
        "category": "privacy",
    },
    {
        "id": "optimize_ssd",
        "name": "Optimizar SSD (TRIM)",
        "description": "Ejecuta comando TRIM en SSDs",
        "command": "defrag C: /L /O",
        "estimated_time": "2-5 min",
        "enabled": False,
        "critical": False,
        "category": "maintenance",
    },
]


# ============================================================================
# FUNCIONES AUXILIARES
# ============================================================================


def is_admin():
    """Verifica privilegios de administrador."""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


def run_command(command, shell=True, timeout=None):
    """
    Ejecuta un comando y retorna el resultado.

    Args:
       command: Comando a ejecutar
       shell: Usar shell
       timeout: Timeout en segundos

    Returns:
       tuple: (returncode, stdout, stderr)
    """
    try:
        result = subprocess.run(
            command, capture_output=True, text=True, shell=shell, timeout=timeout
        )
        return result.returncode, result.stdout.strip(), result.stderr.strip()
    except subprocess.TimeoutExpired:
        return -1, "", "Timeout"
    except Exception as e:
        return -1, "", str(e)


# ============================================================================
# CLASE PRINCIPAL
# ============================================================================


class OptimizeApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Configuración
        self.title(f"{APP_TITLE} {APP_VERSION}")
        self.geometry(APP_SIZE)
        self.resizable(True, True)

        # Variables
        self.is_processing = False
        self.should_cancel = False
        self.task_vars = {}
        self.current_task = None

        # Construir interfaz
        self.build_ui()

        # Verificar prerequisitos
        self.after(300, self.check_prerequisites)

    def build_ui(self):
        """Construye la interfaz con scroll."""

        # Marco principal con scroll
        main_scrollable = ctk.CTkScrollableFrame(
            self,
            fg_color=COLOR_BG_DARK,
            scrollbar_button_color=COLOR_PRIMARY,
            scrollbar_button_hover_color="#1976d2",
        )
        main_scrollable.pack(fill="both", expand=True, padx=15, pady=15)

        # === ENCABEZADO ===
        header_frame = ctk.CTkFrame(
            main_scrollable, fg_color=COLOR_PRIMARY, corner_radius=10
        )
        header_frame.pack(fill="x", pady=(0, 15))

        title_label = ctk.CTkLabel(
            header_frame, text="⚡ " + APP_TITLE, font=FONT_TITLE, text_color="white"
        )
        title_label.pack(pady=15)

        subtitle_label = ctk.CTkLabel(
            header_frame,
            text=f"Autor: Josué Romero  |  Stefanini / PQN  |  {APP_VERSION}",
            font=FONT_SUBTITLE,
            text_color="#e3f2fd",
        )
        subtitle_label.pack(pady=(0, 15))

        # === SELECCIÓN DE TAREAS ===
        tasks_frame = ctk.CTkFrame(
            main_scrollable, fg_color=COLOR_BG_LIGHT, corner_radius=8
        )
        tasks_frame.pack(fill="x", expand=True, pady=(0, 10))

        tasks_title = ctk.CTkLabel(
            tasks_frame,
            text="📋 Seleccione las Tareas a Ejecutar",
            font=FONT_LABEL,
            text_color=COLOR_TEXT,
        )
        tasks_title.pack(pady=(10, 10), anchor="w", padx=15)

        # Scrollable frame para tareas
        tasks_scroll = ctk.CTkScrollableFrame(
            tasks_frame, height=300, fg_color="transparent"
        )
        tasks_scroll.pack(padx=15, pady=(0, 10), fill="x")

        # Agrupar tareas por categoría
        categories = {
            "basic": "🔧 Tareas Básicas",
            "performance": "⚡ Optimización de Rendimiento",
            "privacy": "🔒 Privacidad y Telemetría",
            "maintenance": "🛠️ Mantenimiento Avanzado",
        }

        current_category = None
        for task in OPTIMIZATION_TASKS:
            # Mostrar encabezado de categoría si es nueva
            if task.get("category") != current_category:
                current_category = task.get("category")
                if current_category and current_category in categories:
                    cat_label = ctk.CTkLabel(
                        tasks_scroll,
                        text=categories[current_category],
                        font=("Segoe UI", 11, "bold"),
                        text_color=COLOR_PRIMARY,
                        anchor="w",
                    )
                    cat_label.pack(anchor="w", pady=(10, 5), padx=5)

            task_frame = ctk.CTkFrame(tasks_scroll, fg_color="#242424", corner_radius=6)
            task_frame.pack(fill="x", pady=3, padx=5)

            # Variable para el checkbox
            var = ctk.BooleanVar(value=task["enabled"])
            self.task_vars[task["id"]] = var

            # Checkbox y nombre
            checkbox = ctk.CTkCheckBox(task_frame, text="", variable=var, width=20)
            checkbox.pack(side="left", padx=10, pady=8)

            # Información de la tarea
            info_frame = ctk.CTkFrame(task_frame, fg_color="transparent")
            info_frame.pack(side="left", fill="x", expand=True, padx=5)

            name_label = ctk.CTkLabel(
                info_frame,
                text=task["name"],
                font=("Segoe UI", 10, "bold"),
                text_color=COLOR_SUCCESS if not task["critical"] else COLOR_WARNING,
                anchor="w",
            )
            name_label.pack(anchor="w")

            desc_label = ctk.CTkLabel(
                info_frame,
                text=f"{task['description']} • Tiempo estimado: {task['estimated_time']}",
                font=("Segoe UI", 9),
                text_color="#9e9e9e",
                anchor="w",
            )
            desc_label.pack(anchor="w")

        # Botones de selección rápida
        quick_select_frame = ctk.CTkFrame(tasks_frame, fg_color="transparent")
        quick_select_frame.pack(pady=(0, 10), padx=15)

        ctk.CTkButton(
            quick_select_frame,
            text="Seleccionar Todas",
            command=self.select_all_tasks,
            width=120,
            height=30,
            font=("Segoe UI", 10),
        ).pack(side="left", padx=3)

        ctk.CTkButton(
            quick_select_frame,
            text="Deseleccionar Todas",
            command=self.deselect_all_tasks,
            width=120,
            height=30,
            font=("Segoe UI", 10),
        ).pack(side="left", padx=3)

        ctk.CTkButton(
            quick_select_frame,
            text="Solo Básicas",
            command=self.select_quick_tasks,
            width=120,
            height=30,
            font=("Segoe UI", 10),
        ).pack(side="left", padx=3)

        ctk.CTkButton(
            quick_select_frame,
            text="Optimización Completa",
            command=self.select_performance_tasks,
            width=140,
            height=30,
            font=("Segoe UI", 10),
        ).pack(side="left", padx=3)

        # === PROGRESO ===
        progress_frame = ctk.CTkFrame(
            main_scrollable, fg_color=COLOR_BG_LIGHT, corner_radius=8
        )
        progress_frame.pack(fill="x", pady=(0, 10))

        self.progress_label = ctk.CTkLabel(
            progress_frame,
            text="Listo para iniciar",
            font=FONT_SUBTITLE,
            text_color=COLOR_TEXT,
        )
        self.progress_label.pack(pady=(10, 5))

        self.progress_bar = ctk.CTkProgressBar(
            progress_frame, height=10, corner_radius=5, progress_color=COLOR_PRIMARY
        )
        self.progress_bar.pack(pady=(0, 10), padx=15, fill="x")
        self.progress_bar.set(0)

        # === LOG ===
        log_label = ctk.CTkLabel(
            main_scrollable,
            text="📝 Registro de Actividad",
            font=FONT_LABEL,
            text_color=COLOR_TEXT,
            anchor="w",
        )
        log_label.pack(pady=(5, 5), anchor="w")

        self.text_log = ctk.CTkTextbox(
            main_scrollable,
            height=180,
            font=FONT_CONSOLE,
            fg_color=COLOR_BG_LIGHT,
            text_color=COLOR_TEXT,
            border_width=2,
            border_color=COLOR_PRIMARY,
            corner_radius=8,
        )
        self.text_log.pack(pady=(0, 10), fill="x")

        # === BOTONES ===
        button_frame = ctk.CTkFrame(main_scrollable, fg_color="transparent")
        button_frame.pack(fill="x")

        self.btn_run = ctk.CTkButton(
            button_frame,
            text="▶ Iniciar Optimización",
            command=self.on_start,
            font=FONT_BUTTON,
            height=40,
            corner_radius=8,
            fg_color=COLOR_PRIMARY,
            hover_color="#1976d2",
        )
        self.btn_run.pack(side="left", expand=True, fill="x", padx=(0, 5))

        self.btn_cancel = ctk.CTkButton(
            button_frame,
            text="⏹ Cancelar",
            command=self.cancel_operation,
            font=FONT_BUTTON,
            height=40,
            corner_radius=8,
            fg_color=COLOR_ERROR,
            hover_color="#c62828",
        )
        self.btn_cancel.pack(side="left", expand=True, fill="x", padx=(5, 0))

    def log(self, msg, level="INFO"):
        """Registra mensaje en el log."""
        icons = {
            "INFO": "ℹ",
            "SUCCESS": "✓",
            "WARNING": "⚠",
            "ERROR": "✗",
            "PROGRESS": "⏳",
        }
        icon = icons.get(level, "•")
        timestamp = datetime.now().strftime("%H:%M:%S")

        self.text_log.configure(state="normal")
        self.text_log.insert("end", f"[{timestamp}] {icon} {msg}\n")
        self.text_log.see("end")
        self.text_log.configure(state="disabled")

    def select_all_tasks(self):
        """Selecciona todas las tareas."""
        for var in self.task_vars.values():
            var.set(True)
        self.log("Todas las tareas seleccionadas")

    def deselect_all_tasks(self):
        """Deselecciona todas las tareas."""
        for var in self.task_vars.values():
            var.set(False)
        self.log("Todas las tareas deseleccionadas")

    def select_quick_tasks(self):
        """Selecciona solo tareas básicas/rápidas."""
        for task in OPTIMIZATION_TASKS:
            is_basic = task.get("category") == "basic"
            self.task_vars[task["id"]].set(is_basic)
        self.log("Tareas básicas seleccionadas")

    def select_performance_tasks(self):
        """Selecciona tareas de optimización completa."""
        performance_ids = [
            "cleanmgr",
            "temp_files",
            "defrag_c",
            "disable_visual_effects",
            "disable_startup_delay",
            "optimize_power_plan",
            "disable_superfetch",
            "disable_transparency",
            "optimize_network",
            "clear_dns_cache",
            "disable_game_bar",
        ]
        for task_id, var in self.task_vars.items():
            var.set(task_id in performance_ids)
        self.log("Optimización completa seleccionada")

    def check_prerequisites(self):
        """Verifica prerequisitos."""
        self.log("Verificando prerequisitos del sistema...")

        if not is_admin():
            self.log("✗ Se requieren privilegios de administrador", "ERROR")
            messagebox.showerror(
                "Privilegios Insuficientes",
                "Esta aplicación requiere privilegios de administrador.\n\n"
                "Por favor, ejecute como administrador.",
            )
            self.btn_run.configure(state="disabled")
            return

        self.log("✓ Privilegios de administrador confirmados", "SUCCESS")
        self.log("✓ Sistema listo para optimizar", "SUCCESS")

    def on_start(self):
        """Inicia el proceso de optimización."""
        # Verificar que al menos una tarea esté seleccionada
        selected = [tid for tid, var in self.task_vars.items() if var.get()]

        if not selected:
            messagebox.showwarning(
                "Sin Tareas", "Debe seleccionar al menos una tarea para ejecutar."
            )
            return

        # Confirmar
        response = messagebox.askyesno(
            "Confirmar Optimización",
            f"Se ejecutarán {len(selected)} tareas de optimización.\n\n"
            "⚠ Algunas tareas pueden tardar varios minutos.\n"
            "⚠ Se recomienda guardar todo su trabajo antes de continuar.\n\n"
            "¿Desea continuar?",
        )

        if not response:
            self.log("✗ Operación cancelada por el usuario", "WARNING")
            return

        # Iniciar en hilo separado
        self.is_processing = True
        self.should_cancel = False
        self.btn_run.configure(state="disabled")
        self.btn_cancel.configure(state="normal")

        self.text_log.configure(state="normal")
        self.text_log.delete("1.0", "end")
        self.text_log.configure(state="disabled")

        thread = threading.Thread(target=self.optimize_system, daemon=True)
        thread.start()

    def cancel_operation(self):
        """Cancela la operación en curso."""
        if self.is_processing:
            response = messagebox.askyesno(
                "Cancelar Operación",
                "¿Está seguro de que desea cancelar?\n\n"
                "La tarea actual se completará antes de detenerse.",
            )
            if response:
                self.should_cancel = True
                self.log("Cancelación solicitada...", "WARNING")
                self.btn_cancel.configure(state="disabled")

    def optimize_system(self):
        """Ejecuta las tareas de optimización."""
        try:
            self.log("━" * 75, "INFO")
            self.log("🚀 Iniciando proceso de optimización del sistema", "INFO")
            self.log("━" * 75, "INFO")

            # Obtener tareas seleccionadas
            selected_tasks = [
                task for task in OPTIMIZATION_TASKS if self.task_vars[task["id"]].get()
            ]

            total_tasks = len(selected_tasks)
            completed = 0

            for task in selected_tasks:
                if self.should_cancel:
                    self.log("✗ Proceso cancelado por el usuario", "WARNING")
                    break

                self.current_task = task
                self.update_progress_label(f"Ejecutando: {task['name']}")

                self.log(f"─── {task['name']} ───", "PROGRESS")
                self.log(f"Descripción: {task['description']}")
                self.log(f"Tiempo estimado: {task['estimated_time']}")

                # Ejecutar tarea
                if task["id"] == "winget_update":
                    success = self.update_programs()
                else:
                    success = self.execute_task(task)

                completed += 1
                progress = completed / total_tasks
                self.progress_bar.set(progress)

                if success:
                    self.log(f"✓ {task['name']} completado", "SUCCESS")
                else:
                    self.log(f"⚠ {task['name']} completado con advertencias", "WARNING")

                self.log("")  # Línea en blanco
                time.sleep(0.5)

            # Proceso completado
            if not self.should_cancel:
                self.log("━" * 75, "INFO")
                self.log("✓ Optimización completada exitosamente", "SUCCESS")
                self.log("━" * 75, "INFO")

                # Preguntar por reinicio
                self.after(100, self.ask_restart)

        except Exception as e:
            self.log(f"✗ Error crítico: {str(e)}", "ERROR")
        finally:
            self.is_processing = False
            self.current_task = None
            self.after(100, lambda: self.btn_run.configure(state="normal"))
            self.after(100, lambda: self.btn_cancel.configure(state="disabled"))
            self.update_progress_label("Proceso finalizado")

    def execute_task(self, task):
        """Ejecuta una tarea específica."""
        command = task["command"]

        if isinstance(command, list):
            # Múltiples comandos
            all_success = True
            for cmd in command:
                self.log(f"  → {cmd[:60]}...")
                code, out, err = run_command(cmd, timeout=300)
                if code != 0:
                    all_success = False
                    if err:
                        self.log(f"    Error: {err[:100]}", "ERROR")
            return all_success
        else:
            # Comando único
            self.log(f"  → {command[:60]}...")
            code, out, err = run_command(command, timeout=1800)  # 30 min timeout

            if code == 0:
                if out:
                    lines = out.split("\n")[:5]  # Primeras 5 líneas
                    for line in lines:
                        if line.strip():
                            self.log(f"    {line[:70]}")
                return True
            else:
                if err:
                    self.log(f"    Error: {err[:100]}", "ERROR")
                return False

    def update_programs(self):
        """Actualiza programas con winget."""
        self.log("  → Buscando actualizaciones disponibles...")

        ps_script = """
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
"""

        code, out, err = run_command(f'powershell -Command "{ps_script}"', timeout=1800)

        if out:
            lines = out.split("\n")
            for line in lines[:10]:  # Primeras 10 líneas
                if line.strip():
                    self.log(f"    {line[:70]}")

        return code == 0

    def update_progress_label(self, text):
        """Actualiza el label de progreso."""
        self.progress_label.configure(text=text)

    def ask_restart(self):
        """Pregunta si desea reiniciar."""
        response = messagebox.askyesno(
            "Optimización Completada",
            "✓ La optimización ha finalizado exitosamente.\n\n"
            "Se recomienda reiniciar el equipo para aplicar\n"
            "todos los cambios correctamente.\n\n"
            "¿Desea reiniciar ahora?",
        )

        if response:
            self.log("Reiniciando el sistema...", "INFO")
            run_command('shutdown /r /t 10 /c "Reinicio programado en 10 seg, guarde!"')
            messagebox.showinfo(
                "Reiniciando",
                "El sistema se reiniciará en 10 segundos.\n\n"
                "Guarde todo su trabajo ahora.",
            )
            self.after(2000, self.quit)


# ============================================================================
# PUNTO DE ENTRADA
# ============================================================================

if __name__ == "__main__":
    app = OptimizeApp()
    app.mainloop()
