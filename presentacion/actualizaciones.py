from tkinter import messagebox

from infraestructura import actualizador
from presentacion.hilos import TrabajadorEnSegundoPlano


def verificar(ventana, manual=False, al_cerrar=None):
    """Busca actualizaciones en segundo plano y pregunta al usuario. al_cerrar: función que cierra la app."""
    version = actualizador.__version__

    def con_resultado(info):
        if not info:
            if manual:
                messagebox.showinfo("Actualizaciones", f"Tienes la última versión (v{version}).", parent=ventana)
            return
        notas = f"\n\nNovedades:\n{info['notas']}" if info.get("notas") else ""
        msg = (f"Hay una nueva versión disponible: v{info['version']} (instalada: v{version}).{notas}\n\n"
               "¿Actualizar ahora? La aplicación se cerrará y se abrirá de nuevo.")
        if info.get("obligatoria"):
            messagebox.showinfo("Actualización obligatoria",
                                msg.replace("¿Actualizar ahora?", "Esta actualización es obligatoria."), parent=ventana)
        elif not messagebox.askyesno("Actualización disponible", msg, parent=ventana):
            return
        try:
            actualizador.instalar(info)
        except Exception as e:
            messagebox.showerror("Actualizaciones", f"No se pudo iniciar la actualización:\n{e}", parent=ventana)
            return
        (al_cerrar or ventana.destroy)()

    def con_error(error):
        if manual:
            messagebox.showwarning("Actualizaciones", f"No se pudo consultar el servidor:\n{error}", parent=ventana)

    TrabajadorEnSegundoPlano(ventana, 300).ejecutar_funcion(actualizador.buscar, con_resultado, con_error)


def programar(ventana, al_cerrar=None, demora_ms=3000):
    """Revisión automática al abrir; solo en la app instalada."""
    if actualizador.es_compilado():
        ventana.after(demora_ms, lambda: verificar(ventana, False, al_cerrar))