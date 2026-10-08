#!/usr/bin/env python3
"""Interfaz gráfica del convertidor universal (necesita convertidor_universal.py en la misma carpeta).

Comportamiento:
  * Al abrirse escanea la carpeta (la actual, o la de -i). Si detecta bases de datos "normales"
    (archivos), solo muestra una ventana pequeña de carga y convierte.
  * Si no detecta ninguna, ofrece elegir otra carpeta o conectarse a un servidor.
  * Los servidores (PostgreSQL, MySQL/MariaDB, SQL Server, Oracle u otros por URL) piden acceso
    en un formulario. La contraseña no se guarda en ningún lado.

Uso:
    python convertidor_gui.py
    python convertidor_gui.py -i C:\\datos -r
    python convertidor_gui.py --servidor          # abre directo el formulario de conexión
(En Windows, renómbralo a .pyw para que no aparezca la consola.)
"""

import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import convertidor_universal as cu

ROJO = "#b00020"


class App:
    def __init__(self, args):
        self.args = args
        self.salida_auto = args.salida == cu.SALIDA_POR_DEFECTO
        self.cola = queue.Queue()
        self.cancelar = threading.Event()
        self.marco = None
        self.pantalla = ""
        self.carpeta_escaneada = None

        self.root = tk.Tk()
        self.root.title("Convertidor de bases de datos")
        self.root.resizable(False, False)
        self.root.protocol("WM_DELETE_WINDOW", self.cerrar)
        self.root.after(100, self.sondear)

        if args.url:  # --url: conexión ya indicada, se convierte directo
            try:
                fuentes = cu.fuente_servidor(args)
                self.lanzar_conversion(fuentes, f"Servidor {fuentes[0][1]}")
            except cu.ErrorUsuario as e:
                self.pantalla_conexion()
                self.lbl_msg.configure(text=str(e), foreground=ROJO)
        elif args.servidor:
            self.pantalla_conexion()
        else:
            self.iniciar_escaneo()

    def ejecutar(self):
        self.root.mainloop()

    # ------------------------------------------------------------------ utilidades de ventana

    def _marco(self, ancho, alto):
        if self.marco is not None:
            self.marco.destroy()
        self.marco = ttk.Frame(self.root, padding=14)
        self.marco.pack(fill="both", expand=True)
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() - ancho) // 2
        y = (self.root.winfo_screenheight() - alto) // 3
        self.root.geometry(f"{ancho}x{alto}+{x}+{y}")
        return self.marco

    def cerrar(self):
        self.cancelar.set()
        self.root.destroy()

    @staticmethod
    def abrir(ruta):
        try:
            if sys.platform.startswith("win"):
                os.startfile(str(ruta))
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(ruta)])
            else:
                subprocess.Popen(["xdg-open", str(ruta)])
        except Exception as e:
            messagebox.showerror("No se pudo abrir", str(e))

    def _salida_en_carpeta(self, carpeta):
        if self.salida_auto:
            self.args.salida = str(Path(carpeta).resolve() / cu.SALIDA_POR_DEFECTO)

    # ------------------------------------------------------------------ cola de eventos (hilos -> ventana)

    def sondear(self):
        try:
            while True:
                tipo, datos = self.cola.get_nowait()
                manejador = getattr(self, f"_ev_{tipo}", None)
                if manejador:
                    manejador(**datos)
        except queue.Empty:
            pass
        self.root.after(100, self.sondear)

    def _en_hilo(self, funcion):
        threading.Thread(target=funcion, daemon=True).start()

    def _publicar(self, tipo, **datos):
        self.cola.put((tipo, datos))

    # ------------------------------------------------------------------ pantalla de carga (la pequeña)

    def pantalla_carga(self, texto):
        self.pantalla = "carga"
        m = self._marco(400, 135)
        self.lbl_estado = ttk.Label(m, text=texto, wraplength=370, justify="left")
        self.lbl_estado.pack(anchor="w", pady=(4, 10))
        self.barra = ttk.Progressbar(m, mode="indeterminate", length=370)
        self.barra.pack(fill="x")
        self.barra.start(12)
        self.btn_cancelar = ttk.Button(m, text="Cancelar", command=self.cancelar_proceso)
        self.btn_cancelar.pack(anchor="e", pady=(12, 0))

    def cancelar_proceso(self):
        self.cancelar.set()
        self.btn_cancelar.configure(state="disabled")
        self.lbl_estado.configure(text="Cancelando… (se guardará lo ya procesado)")

    def iniciar_escaneo(self):
        self.cancelar.clear()
        self._salida_en_carpeta(self.args.carpeta)
        self.pantalla_carga("Buscando bases de datos…")

        def tarea():
            try:
                self._publicar("escaneo", esc=cu.escanear(self.args))
            except cu.ErrorUsuario as e:
                self._publicar("error", texto=str(e))
            except Exception as e:
                self._publicar("error", texto=f"{type(e).__name__}: {e}")
        self._en_hilo(tarea)

    def lanzar_conversion(self, fuentes, origen, no_reconocidos=(), excluidos=()):
        self.cancelar.clear()
        self.pantalla_carga("Preparando…")

        def tarea():
            try:
                res = cu.convertir(
                    self.args, fuentes, origen, no_reconocidos, excluidos,
                    evento=self._publicar, cancelar=self.cancelar,
                )
                self._publicar("fin", res=res)
            except cu.ErrorUsuario as e:
                self._publicar("error", texto=str(e))
            except Exception as e:
                self._publicar("error", texto=f"{type(e).__name__}: {e}")
        self._en_hilo(tarea)

    # ------------------------------------------------------------------ eventos

    def _ev_escaneo(self, esc):
        if self.pantalla != "carga":
            return
        self.carpeta_escaneada = esc["carpeta"]
        if esc["fuentes"]:
            self.lanzar_conversion(esc["fuentes"], esc["origen"], esc["no_reconocidos"], esc["excluidos"])
        else:
            self.pantalla_inicio()

    def _ev_inicio(self, total):
        if self.pantalla != "carga":
            return
        self.barra.stop()
        self.barra.configure(mode="determinate", maximum=max(1, total), value=0)

    def _ev_archivo(self, i, total, nombre):
        if self.pantalla != "carga":
            return
        self.barra.configure(value=i - 1)
        corto = nombre if len(nombre) <= 45 else "…" + nombre[-44:]
        self.lbl_estado.configure(text=f"Convirtiendo {i} de {total}\n{corto}")

    def _ev_mensaje(self, texto):
        pass  # los detalles van al registro (.log); la ventana de carga se mantiene mínima

    def _ev_error(self, texto):
        messagebox.showerror("No se pudo continuar", texto)
        if self.args.url:
            self.pantalla_conexion()
        else:
            self.pantalla_inicio()

    def _ev_fin(self, res):
        self.args.url = None  # no conservar credenciales en memoria
        self.pantalla = "resultado"
        c = res["conteo"]
        m = self._marco(400, 170)
        titulo = "⚠️ Cancelado" if res["cancelado"] else "✅ Listo"
        resumen = f"{titulo}: {c['OK']} convertidos"
        if c["OMITIDO"]:
            resumen += f", {c['OMITIDO']} omitidos"
        if c["ERROR"]:
            resumen += f", {c['ERROR']} con error"
        ttk.Label(m, text=resumen, wraplength=370).pack(anchor="w", pady=(2, 2))
        ttk.Label(m, text=f"Tiempo: {res['segundos']:.1f} s. Los detalles están en el registro (.log) "
                          f"y en la pestaña '{cu.NOMBRE_INDICE}' del Excel.",
                  wraplength=370, justify="left").pack(anchor="w")
        fila = ttk.Frame(m)
        fila.pack(anchor="e", pady=(16, 0), fill="x")
        ttk.Button(fila, text="Cerrar", command=self.cerrar).pack(side="right")
        ttk.Button(fila, text="Ver registro", command=lambda: self.abrir(res["ruta_log"])).pack(side="right", padx=6)
        ttk.Button(fila, text="Abrir Excel", command=lambda: self.abrir(res["salida"])).pack(side="right")

    # ------------------------------------------------------------------ pantalla: nada detectado

    def pantalla_inicio(self):
        self.pantalla = "inicio"
        m = self._marco(430, 190)
        carpeta = self.carpeta_escaneada or Path(self.args.carpeta).resolve()
        ttk.Label(m, text=f"No se detectó ninguna base de datos en:\n{carpeta}",
                  wraplength=400, justify="left").pack(anchor="w", pady=(2, 14))
        ttk.Button(m, text="Elegir otra carpeta…", command=self.elegir_carpeta).pack(fill="x", pady=2)
        ttk.Button(m, text="Conectar a un servidor…", command=self.pantalla_conexion).pack(fill="x", pady=2)
        ttk.Button(m, text="Salir", command=self.cerrar).pack(fill="x", pady=2)

    def elegir_carpeta(self):
        carpeta = filedialog.askdirectory(title="Carpeta con bases de datos",
                                          initialdir=str(self.carpeta_escaneada or "."))
        if carpeta:
            self.args.carpeta = carpeta
            self.iniciar_escaneo()

    # ------------------------------------------------------------------ pantalla: acceso a servidor

    def pantalla_conexion(self):
        self.pantalla = "conexion"
        self.args.url = None
        m = self._marco(520, 600)
        m.columnconfigure(1, weight=1)

        self.v_motor = tk.StringVar(value="PostgreSQL")
        self.v_host, self.v_puerto = tk.StringVar(), tk.StringVar(value="5432")
        self.v_usuario, self.v_clave = tk.StringVar(), tk.StringVar()
        self.v_base, self.v_esquema = tk.StringVar(), tk.StringVar()
        self.v_url, self.v_limite = tk.StringVar(), tk.StringVar()
        self.v_win = tk.BooleanVar(value=False)
        self.v_salida = tk.StringVar(value=self.args.salida if not self.salida_auto
                                     else str(Path.cwd() / cu.SALIDA_POR_DEFECTO))

        def fila(r, texto, widget):
            ttk.Label(m, text=texto).grid(row=r, column=0, sticky="w", pady=3, padx=(0, 8))
            widget.grid(row=r, column=1, sticky="ew", pady=3)
            return widget

        cb = fila(0, "Motor", ttk.Combobox(m, textvariable=self.v_motor, values=list(cu.MOTORES), state="readonly"))
        cb.bind("<<ComboboxSelected>>", self._motor_cambio)
        self.ent_host = fila(1, "Servidor", ttk.Entry(m, textvariable=self.v_host))
        self.ent_puerto = fila(2, "Puerto", ttk.Entry(m, textvariable=self.v_puerto))
        self.chk_win = ttk.Checkbutton(m, text="Autenticación de Windows (SQL Server)",
                                       variable=self.v_win, command=self._auth_cambio)
        self.chk_win.grid(row=3, column=1, sticky="w")
        self.ent_usuario = fila(4, "Usuario", ttk.Entry(m, textvariable=self.v_usuario))
        self.ent_clave = fila(5, "Contraseña", ttk.Entry(m, textvariable=self.v_clave, show="•"))
        self.ent_base = fila(6, "Base de datos / servicio", ttk.Entry(m, textvariable=self.v_base))
        self.ent_esquema = fila(7, "Esquema (opcional)", ttk.Entry(m, textvariable=self.v_esquema))
        self.ent_url = fila(8, "URL (solo «Otro»)", ttk.Entry(m, textvariable=self.v_url))
        self.campos_servidor = [self.ent_host, self.ent_puerto, self.ent_usuario, self.ent_clave, self.ent_base]

        self.btn_probar = ttk.Button(m, text="Probar conexión y listar tablas", command=self.probar)
        self.btn_probar.grid(row=9, column=0, columnspan=2, sticky="ew", pady=(10, 2))
        self.lbl_msg = ttk.Label(m, text="", wraplength=480, justify="left")
        self.lbl_msg.grid(row=10, column=0, columnspan=2, sticky="w")

        marco_lista = ttk.Frame(m)
        marco_lista.grid(row=11, column=0, columnspan=2, sticky="nsew", pady=4)
        marco_lista.columnconfigure(0, weight=1)
        self.lista = tk.Listbox(marco_lista, selectmode="extended", height=7, exportselection=False)
        self.lista.grid(row=0, column=0, sticky="ew")
        scroll = ttk.Scrollbar(marco_lista, orient="vertical", command=self.lista.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        self.lista.configure(yscrollcommand=scroll.set)

        fila(12, "Máx. filas por tabla", ttk.Entry(m, textvariable=self.v_limite))
        marco_salida = ttk.Frame(m)
        marco_salida.columnconfigure(0, weight=1)
        ttk.Entry(marco_salida, textvariable=self.v_salida).grid(row=0, column=0, sticky="ew")
        ttk.Button(marco_salida, text="Examinar…", command=self.elegir_salida).grid(row=0, column=1, padx=(6, 0))
        fila(13, "Guardar en", marco_salida)

        botones = ttk.Frame(m)
        botones.grid(row=14, column=0, columnspan=2, sticky="e", pady=(14, 0))
        ttk.Button(botones, text="Salir", command=self.cerrar).pack(side="right")
        self.btn_convertir = ttk.Button(botones, text="Convertir", command=self.convertir_servidor, state="disabled")
        self.btn_convertir.pack(side="right", padx=6)

        self._motor_cambio()

    def _motor_cambio(self, *_):
        motor = self.v_motor.get()
        cfg = cu.MOTORES[motor]
        libre = cfg["driver"] is None
        self.v_puerto.set(str(cfg["puerto"]) if cfg["puerto"] else "")
        for w in self.campos_servidor:
            w.configure(state="disabled" if libre else "normal")
        self.ent_url.configure(state="normal" if libre else "disabled")
        es_sqlserver = motor == "SQL Server"
        self.chk_win.configure(state="normal" if es_sqlserver else "disabled")
        if not es_sqlserver:
            self.v_win.set(False)
        self._auth_cambio()

    def _auth_cambio(self):
        if self.v_motor.get() == "Otro (URL de SQLAlchemy)":
            return
        estado = "disabled" if self.v_win.get() else "normal"
        self.ent_usuario.configure(state=estado)
        self.ent_clave.configure(state=estado)

    def elegir_salida(self):
        ruta = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")],
                                            initialfile=Path(self.v_salida.get()).name)
        if ruta:
            self.v_salida.set(ruta)

    def _url_desde_formulario(self):
        return cu.construir_url(
            self.v_motor.get(), self.v_host.get(), self.v_puerto.get(), self.v_usuario.get(),
            self.v_clave.get(), self.v_base.get(), self.v_win.get(), self.v_url.get(),
        )

    def probar(self):
        try:
            url = self._url_desde_formulario()
        except cu.ErrorUsuario as e:
            self.lbl_msg.configure(text=str(e), foreground=ROJO)
            return
        esquema = self.v_esquema.get().strip() or None
        self.lbl_msg.configure(text="Conectando…", foreground="")
        self.btn_probar.configure(state="disabled")
        self.btn_convertir.configure(state="disabled")

        def tarea():
            try:
                tablas = cu.listar_tablas_servidor(url, esquema)
                self._publicar("conexion", ok=True, texto=f"Conexión correcta: {len(tablas)} tablas.", tablas=tablas)
            except cu.ErrorUsuario as e:
                self._publicar("conexion", ok=False, texto=str(e), tablas=[])
            except Exception as e:
                self._publicar("conexion", ok=False, texto=f"{type(e).__name__}: {e}", tablas=[])
        self._en_hilo(tarea)

    def _ev_conexion(self, ok, texto, tablas):
        if self.pantalla != "conexion":
            return
        self.btn_probar.configure(state="normal")
        self.lbl_msg.configure(text=texto, foreground="" if ok else ROJO)
        self.lista.delete(0, "end")
        for t in tablas:
            self.lista.insert("end", t)
        if tablas:
            self.lista.selection_set(0, "end")
        self.btn_convertir.configure(state="normal" if (ok and tablas) else "disabled")

    def convertir_servidor(self):
        try:
            url = self._url_desde_formulario()
            limite_txt = self.v_limite.get().strip()
            limite = int(limite_txt) if limite_txt else None
            if limite is not None and limite < 1:
                raise ValueError
        except cu.ErrorUsuario as e:
            self.lbl_msg.configure(text=str(e), foreground=ROJO)
            return
        except ValueError:
            self.lbl_msg.configure(text="El máximo de filas debe ser un número entero mayor que 0.", foreground=ROJO)
            return
        tablas = [self.lista.get(i) for i in self.lista.curselection()]
        if not tablas:
            self.lbl_msg.configure(text="Selecciona al menos una tabla.", foreground=ROJO)
            return

        self.args.url = url
        self.args.esquema = self.v_esquema.get().strip() or None
        self.args.tablas = tablas
        self.args.limite_filas = limite
        self.args.salida = self.v_salida.get().strip() or str(Path.cwd() / cu.SALIDA_POR_DEFECTO)
        self.args.log = None
        try:
            fuentes = cu.fuente_servidor(self.args)
        except cu.ErrorUsuario as e:
            self.lbl_msg.configure(text=str(e), foreground=ROJO)
            return
        self.v_clave.set("")  # la contraseña ya no se necesita en el formulario
        self.lanzar_conversion(fuentes, f"Servidor {fuentes[0][1]}")


def main():
    args = cu.parsear()
    args.sin_barra = True  # las barras de texto no aplican en la interfaz gráfica
    App(args).ejecutar()


if __name__ == "__main__":
    main()
