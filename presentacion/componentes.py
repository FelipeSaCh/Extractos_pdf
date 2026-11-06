import tkinter as tk
from tkinter import ttk


def modal(ventana, parent):
    """Hace modal la ventana respecto a la raíz, liberando el grab al minimizar y recuperándolo al restaurar."""
    raiz = parent.winfo_toplevel()
    ventana.transient(raiz)

    def capturar():
        try:
            if ventana.winfo_exists() and ventana.winfo_viewable():
                ventana.grab_set()
        except tk.TclError:
            pass

    def al_minimizar(e):
        if e.widget is raiz and ventana.winfo_exists():
            ventana.grab_release()

    def al_restaurar(e):
        if e.widget is raiz and ventana.winfo_exists():
            ventana.deiconify()
            ventana.lift()
            ventana.focus_force()
            ventana.after(50, capturar)

    ids = [(ev, raiz.bind(ev, fn, "+")) for ev, fn in (("<Unmap>", al_minimizar), ("<Map>", al_restaurar))]

    def liberar(e):
        if e.widget is ventana:
            for ev, fid in ids:
                try:
                    raiz.unbind(ev, fid)
                except tk.TclError:
                    pass

    ventana.bind("<Destroy>", liberar, "+")
    ventana.update_idletasks()
    capturar()


def elegir_opcion(parent, titulo, texto, opciones, sugerida=None):
    """Diálogo de opciones excluyentes; devuelve la clave elegida o None si se cancela."""
    var = tk.StringVar(value=sugerida or next(iter(opciones)))
    res = {}

    ven = tk.Toplevel(parent)
    ven.title(titulo)
    ven.transient(parent)
    ven.resizable(False, False)
    frm = ttk.Frame(ven, padding=16)
    frm.pack(fill=tk.BOTH, expand=True)
    ttk.Label(frm, text=texto, font=("Segoe UI", 10)).pack(anchor="w", pady=(0, 8))
    for clave, descripcion in opciones.items():
        etiqueta = f"{descripcion}  (detectado)" if clave == sugerida else descripcion
        ttk.Radiobutton(frm, text=etiqueta, variable=var, value=clave).pack(anchor="w", pady=2)

    def aceptar(_=None):
        res["valor"] = var.get()
        ven.destroy()

    botones = ttk.Frame(frm)
    botones.pack(fill=tk.X, pady=(12, 0))
    ttk.Button(botones, text="Cancelar", command=ven.destroy).pack(side=tk.RIGHT)
    ttk.Button(botones, text="Aceptar", command=aceptar).pack(side=tk.RIGHT, padx=6)
    ven.bind("<Return>", aceptar)
    ven.bind("<Escape>", lambda e: ven.destroy())
    ven.grab_set()
    parent.wait_window(ven)
    return res.get("valor")