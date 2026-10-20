import tkinter as tk
from tkinter import messagebox, ttk


def pedir_parametros_plano(parent):
    dialogo = tk.Toplevel(parent)
    dialogo.title("Parámetros del Archivo Plano")
    dialogo.geometry("350x220")
    dialogo.resizable(False, False)
    dialogo.transient(parent)
    dialogo.grab_set()

    frame = ttk.Frame(dialogo, padding=20)
    frame.pack(fill=tk.BOTH, expand=True)

    ttk.Label(frame, text="Número Consecutivo inicial:", font=("Segoe UI", 10)).pack(anchor="w", pady=(0, 5))
    var_consecutivo = tk.StringVar(value="1")
    entry_consecutivo = ttk.Entry(frame, textvariable=var_consecutivo, font=("Segoe UI", 10))
    entry_consecutivo.pack(fill=tk.X, pady=(0, 15))

    ttk.Label(frame, text="NIT de la Empresa (Tercero por defecto):", font=("Segoe UI", 10)).pack(anchor="w", pady=(0, 5))
    var_nit = tk.StringVar()
    ttk.Entry(frame, textvariable=var_nit, font=("Segoe UI", 10)).pack(fill=tk.X, pady=(0, 20))

    resultado = {}

    def confirmar(event=None):
        consecutivo = var_consecutivo.get().strip()
        resultado["consecutivo"] = consecutivo if consecutivo.isdigit() else "1"
        resultado["nit"] = var_nit.get().strip()
        dialogo.destroy()

    ttk.Button(frame, text="Confirmar y Generar", command=confirmar).pack(fill=tk.X)
    dialogo.bind("<Return>", confirmar)
    dialogo.bind("<Escape>", lambda e: dialogo.destroy())
    entry_consecutivo.focus_set()

    parent.wait_window(dialogo)
    return (resultado["consecutivo"], resultado["nit"]) if resultado else None


def generar_plano_con_dialogo(parent, excel_path):
    parametros = pedir_parametros_plano(parent)
    if parametros is None:
        return None

    from archivo_plano import GeneradorArchivoPlano

    consecutivo, nit = parametros
    try:
        parent.config(cursor="watch")
        parent.update_idletasks()
        ruta_plano = GeneradorArchivoPlano(excel_path).generar(
            numero_inicial=consecutivo, nit_empresa=nit
        )
    except Exception as e:
        messagebox.showerror("Error en Plano", f"Falló la generación del archivo plano:\n{e}", parent=parent)
        return None
    finally:
        parent.config(cursor="")

    messagebox.showinfo("Archivo Plano", f"El archivo plano se ha generado en:\n{ruta_plano}", parent=parent)
    return ruta_plano