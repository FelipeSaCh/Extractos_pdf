#ui_theme.py
import tkinter as tk
from tkinter import ttk

COLOR_BG = "#F4F6F8"
COLOR_SIDEBAR = "#FFFFFF"
COLOR_PRIMARY = "#2F6FED"
COLOR_PRIMARY_DARK = "#2457BE"
COLOR_TEXT = "#1F2933"
COLOR_MUTED = "#7B8794"
COLOR_SUCCESS = "#2E7D32"
COLOR_BORDER = "#E4E7EB"
COLOR_CANVAS_BG = "#525659"

FONT_TITLE = ("Segoe UI", 16, "bold")
FONT_SUBTITLE = ("Segoe UI", 9)
FONT_BUTTON = ("Segoe UI", 10)
FONT_INFO = ("Segoe UI", 9)
FONT_STATUS = ("Segoe UI", 9, "bold")


def aplicar_estilos(root: tk.Tk) -> ttk.Style:
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass

    style.configure("TFrame", background=COLOR_BG)
    style.configure("Sidebar.TFrame", background=COLOR_SIDEBAR)
    style.configure("Content.TFrame", background=COLOR_BG)

    style.configure(
        "Title.TLabel",
        background=COLOR_SIDEBAR,
        foreground=COLOR_TEXT,
        font=FONT_TITLE,
    )
    style.configure(
        "Subtitle.TLabel",
        background=COLOR_SIDEBAR,
        foreground=COLOR_MUTED,
        font=FONT_SUBTITLE,
    )
    style.configure(
        "Info.TLabel",
        background=COLOR_SIDEBAR,
        foreground=COLOR_MUTED,
        font=FONT_INFO,
    )
    style.configure(
        "Status.TLabel",
        background=COLOR_SIDEBAR,
        foreground=COLOR_SUCCESS,
        font=FONT_STATUS,
    )
    style.configure(
        "Footer.TLabel",
        background=COLOR_SIDEBAR,
        foreground=COLOR_MUTED,
        font=("Segoe UI", 8),
    )

    style.configure(
        "Primary.TButton",
        font=FONT_BUTTON,
        padding=10,
        background=COLOR_PRIMARY,
        foreground="white",
        borderwidth=0,
        focusthickness=0,
    )
    style.map(
        "Primary.TButton",
        background=[
            ("active", COLOR_PRIMARY_DARK),
            ("disabled", "#B9C6E4"),
        ],
        foreground=[("disabled", "#F0F0F0")],
    )

    style.configure(
        "Secondary.TButton",
        font=FONT_BUTTON,
        padding=10,
        background="#EDF1F7",
        foreground=COLOR_TEXT,
        borderwidth=0,
    )
    style.map(
        "Secondary.TButton",
        background=[("active", "#DDE4EF"), ("disabled", "#F3F4F6")],
        foreground=[("disabled", "#B0B7C3")],
    )

    style.configure("TPanedwindow", background=COLOR_BG)
    style.configure("TNotebook", background=COLOR_BG, borderwidth=0)
    style.configure("TNotebook.Tab", font=FONT_BUTTON, padding=(16, 8))
    style.map(
        "TNotebook.Tab",
        background=[("selected", "white")],
        foreground=[("selected", COLOR_PRIMARY)],
    )

    style.configure(
        "Treeview",
        font=("Segoe UI", 9),
        rowheight=26,
        background="white",
        fieldbackground="white",
        foreground=COLOR_TEXT,
    )
    style.configure(
        "Treeview.Heading",
        font=("Segoe UI", 9, "bold"),
        background="#EEF1F5",
        foreground=COLOR_TEXT,
        padding=6,
    )
    style.map(
        "Treeview",
        background=[("selected", COLOR_PRIMARY)],
        foreground=[("selected", "white")],
    )

    return style


def construir_warning_box(parent):
    warning_frame = tk.Frame(
        parent,
        bg="#FDECEA",
        highlightbackground="#D93025",
        highlightthickness=2,
        bd=0,
    )
    warning_frame.pack(fill=tk.X, pady=(0, 15))

    warning_label = tk.Label(
        warning_frame,
        text=(
            "⚠️ Recuerde que este programa únicamente puede usarse con"
            " archivos PDF que contengan texto (no imágenes escaneadas), y"
            " que el soporte actual es exclusivo para extractos con el"
            " formato de Bancolombia y Banco de Bogotá. Otros bancos pueden no ser compatibles hasta el momento"
        ),
        bg="#FDECEA",
        fg="#D93025",
        font=("Segoe UI", 8),
        wraplength=220,
        justify="left",
        padx=10,
        pady=8,
    )
    warning_label.pack(fill=tk.X)

    return warning_frame