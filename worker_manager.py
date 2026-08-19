#worker_manager.py
import queue
import threading


class TrabajadorEnSegundoPlano:
    """Ejecuta tareas pesadas en un hilo secundario y sincroniza el resultado

    con la interfaz Tkinter mediante sondeo (polling) de una cola.

    La tarea (`funcion_tarea`) recibe la cola como único argumento y es
    responsable de encolar su propio progreso (tupla ("progreso", valor, texto))
    y de finalizar encolando ("exito", resultado, None) o ("error", excepcion, None).
    Ningún widget de Tkinter debe tocarse desde dentro de `funcion_tarea`.
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

        self.root.after(
            self.intervalo_ms,
            lambda: self._sondear(cola, on_progreso, on_exito, on_error),
        )