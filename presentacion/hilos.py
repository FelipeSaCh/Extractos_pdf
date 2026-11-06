import queue
import threading
import tkinter as tk
from tkinter import ttk

FUENTE = ("Segoe UI", 10)


class TrabajadorEnSegundoPlano:
    """Ejecuta tareas pesadas en un hilo secundario y entrega el resultado al hilo de Tkinter por sondeo de una cola.

    La tarea recibe la cola y encola ("progreso", valor, texto), ("exito", resultado, None) o ("error", excepcion, None).
    Ningún widget debe tocarse desde la tarea.
    """

    def __init__(self, root, intervalo_ms=50):
        self.root = root
        self.intervalo_ms = intervalo_ms

    def ejecutar(self, funcion_tarea, on_progreso=None, on_exito=None, on_error=None):
        cola = queue.Queue()

        def envoltorio():
            try:
                funcion_tarea(cola)
            except Exception as e:
                cola.put(("error", e, None))

        threading.Thread(target=envoltorio, daemon=True).start()
        self._sondear(cola, on_progreso, on_exito, on_error)

    def ejecutar_funcion(self, funcion, on_exito=None, on_error=None):
        """Atajo para una función sin progreso: su retorno va a on_exito."""
        self.ejecutar(lambda cola: cola.put(("exito", funcion(), None)), on_exito=on_exito, on_error=on_error)

    def _sondear(self, cola, on_progreso, on_exito, on_error):
        try:
            while True:
                tipo, valor, texto = cola.get_nowait()
                if tipo == "progreso":
                    if on_progreso:
                        on_progreso(valor, texto)
                elif tipo == "exito":
                    if on_exito:
                        on_exito(valor)
                    return
                elif tipo == "error":
                    if on_error:
                        on_error(valor)
                    return
        except queue.Empty:
            pass
        except tk.TclError:
            return
        try:
            self.root.after(self.intervalo_ms, lambda: self._sondear(cola, on_progreso, on_exito, on_error))
        except tk.TclError:
            pass


def con_carga(parent, texto, funcion, titulo="Archivo Plano"):
    """Ejecuta funcion() en segundo plano mostrando una ventana modal de espera; devuelve su resultado o relanza su error."""
    from presentacion.componentes import modal

    res = {}
    ven = tk.Toplevel(parent)
    ven.title(titulo)
    ven.resizable(False, False)
    ven.protocol("WM_DELETE_WINDOW", lambda: None)
    ttk.Label(ven, text=texto, font=FUENTE, padding=(20, 15, 20, 8)).pack()
    barra = ttk.Progressbar(ven, mode="indeterminate", length=300)
    barra.pack(padx=20, pady=(0, 18))
    barra.start(10)
    ven.update_idletasks()
    x = parent.winfo_rootx() + (parent.winfo_width() - ven.winfo_width()) // 2
    y = parent.winfo_rooty() + (parent.winfo_height() - ven.winfo_height()) // 2
    ven.geometry(f"+{max(x, 0)}+{max(y, 0)}")
    modal(ven, parent)

    def terminar(clave):
        def guardar(valor):
            res[clave] = valor
            barra.stop()
            ven.destroy()
        return guardar

    TrabajadorEnSegundoPlano(parent, 100).ejecutar_funcion(funcion, terminar("ok"), terminar("error"))
    parent.wait_window(ven)
    if "error" in res:
        raise res["error"]
    return res.get("ok")