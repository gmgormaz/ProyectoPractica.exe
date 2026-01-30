from pathlib import Path
import pandas as pd
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox
from app.jobs import Prueba_Continuidad, Prueba_Aislamiento, Prueba_Caida_Tension, Prueba_Lazo, Prueba_Diferenciales
#Faltan las tablas :P
from app.jobs import Tabla_Continuidad, Tabla_Aislamiento, Tabla_Aislamiento_E_C, Tabla_T_EC, Tabla_Bucle_Falla, Tabla_C_T_T
from app.jobs import Tabla_Aislamiento_Trifasica, Tabla_Continuidad_Trifasica
from app.output import base_output_dir, reveal_in_explorer
import os
import sys
def crear_acceso_directo_si_no_existe():
    try:
        import win32com.client

        escritorio = os.path.join(os.environ["USERPROFILE"], "Desktop")
        acceso = os.path.join(escritorio, "ProcesadorCSV.lnk")
        if os.path.exists(acceso):
            return

        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(acceso)

        exe_path = sys.executable
        shortcut.Targetpath = exe_path
        shortcut.WorkingDirectory = os.path.dirname(exe_path)
        shortcut.IconLocation = exe_path
        shortcut.save()
    except Exception:
        pass


APP_TITLE = "Procesador CSV"

JOBS = [("continuidad", Prueba_Continuidad.run), ("aislamiento", Prueba_Aislamiento.run),
        ("Caida de Tensión", Prueba_Caida_Tension.run), ("Prueba de Lazo",Prueba_Lazo.run ),
        ("Diferenciales", Prueba_Diferenciales.run), ("Tabla Continuidad", Tabla_Continuidad.run),
        ("Tabla Aislamiento", Tabla_Aislamiento.run),("Tabla Aislamiento E-C", Tabla_Aislamiento_E_C.run)]


def pick_csv() -> Path | None:
    p = filedialog.askopenfilename(
        title="Selecciona un archivo CSV",
        filetypes=[("CSV", "*.csv")],
    )
    return Path(p) if p else None

def set_status(text: str):
    status_var.set(text)
    app.update_idletasks()

def do_continuidad():
    in_path = pick_csv()
    if not in_path:
        set_status("Cancelado.")
        return

    try:
        set_status("Procesando: Continuidad…")
        out_path = Prueba_Continuidad.run(in_path)
        set_status("Listo")
        messagebox.showinfo("Listo", f"Archivo generado:\n{out_path}")
        reveal_in_explorer(out_path)
    except Exception as e:
        set_status("Error")
        messagebox.showerror("Error", str(e))

def do_aislamiento():
    in_path = pick_csv()
    if not in_path:
        return

    try:
        out_path = Prueba_Aislamiento.run(in_path)
        messagebox.showinfo("Listo", f"Archivo generado:\n{out_path}")
        reveal_in_explorer(out_path)
    except Exception as e:
        messagebox.showerror("Error", str(e))


def do_caida_T():
    in_path = pick_csv()
    if not in_path:
        return

    try:
        df = pd.read_csv(in_path)
        n = len(df[df["Test Function"].isin(["Prueba de lazo sin disparos"])])

        win = ttk.Toplevel(app)
        win.title("I nominal")

        vars_in = []

        for i in range(n):
            ttk.Label(win, text=f"Circuito {i+1}").grid(row=i, column=0)

            v = ttk.StringVar(value="16")
            ttk.Entry(win, textvariable=v, width=6).grid(row=i, column=1)

            vars_in.append(v)

        def aceptar():
            try:
                corrientes = [int(v.get()) for v in vars_in]
            except:
                messagebox.showerror("Error", "Solo números enteros.")
                return

            win.destroy()
            out_path = Prueba_Caida_Tension.run(in_path, corrientes)
            messagebox.showinfo("Listo", f"Archivo generado:\n{out_path}")
            reveal_in_explorer(out_path)

        ttk.Button(win, text="Procesar", command=aceptar).grid(row=n+1, column=0, columnspan=2)

    except Exception as e:
        messagebox.showerror("Error", str(e))

def do_Lazo():
    in_path = pick_csv()
    if not in_path:
        return

    try:
        df = pd.read_csv(in_path)
        df = df[df["Test Function"].isin(["Prueba de lazo sin disparos"])]
        n = len(df)

        win = ttk.Toplevel(app)
        win.title("Datos Prueba de Lazo")

        filas = []

        for i in range(n):
            ttk.Label(win, text=f"Circuito {i+1}").grid(row=i, column=0)

            v_in = ttk.StringVar(value="16")
            e = ttk.Entry(win, textvariable=v_in, width=6)
            e.grid(row=i, column=1)

            v_c = ttk.StringVar(value="C")
            cb = ttk.Combobox(win, values=["B","C","D"], textvariable=v_c, width=4)
            cb.grid(row=i, column=2)

            filas.append((v_in, v_c))

        def aceptar():
            circuitos = []
            for vin, vc in filas:
                circuitos.append({
                    "In": int(vin.get()),
                    "curva": vc.get()
                })

            win.destroy()

            out_path = Prueba_Lazo.run(in_path, circuitos)
            messagebox.showinfo("Listo", f"Archivo generado:\n{out_path}")
            reveal_in_explorer(out_path)

        ttk.Button(win, text="Procesar", command=aceptar).grid(row=n+1, column=0, columnspan=3)

    except Exception as e:
        messagebox.showerror("Error", str(e))

def do_tabla_riso_entre_circuitos():
    win = ttk.Toplevel(app)
    win.title("Tabla Riso entre circuitos")

    v_dif = ttk.StringVar(value="1")
    vars_nc = []

    ttk.Label(win, text="Diferenciales").pack(anchor="w", padx=10, pady=(10,0))
    ttk.Entry(win, textvariable=v_dif, width=8).pack(anchor="w", padx=10)

    frame = ttk.Frame(win, padding=10)
    frame.pack()

    def set_difs():
        for w in frame.winfo_children():
            w.destroy()
        vars_nc.clear()

        try:
            n = int(v_dif.get())
            if n < 1:
                raise ValueError()
        except:
            messagebox.showerror("Error", "Diferenciales inválido")
            return

        for i in range(n):
            row = ttk.Frame(frame)
            row.pack(pady=2)
            ttk.Label(row, text=f"Dif {i+1}").pack(side="left")
            v = ttk.StringVar(value="1")
            ttk.Entry(row, textvariable=v, width=6).pack(side="left", padx=6)
            vars_nc.append(v)

    ttk.Button(win, text="Setear diferenciales", command=set_difs).pack(padx=10, pady=6)

    def crear():
        try:
            nc = [int(v.get()) for v in vars_nc]
            if any(x < 0 for x in nc):
                raise ValueError()
        except:
            messagebox.showerror("Error", "Circuitos por diferencial inválidos (enteros >= 0).")
            return

        win.destroy()
        try:
            out_path = Tabla_T_EC.run(nc)
            messagebox.showinfo("Listo", f"Tabla creada:\n{out_path}")
            reveal_in_explorer(out_path)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    ttk.Button(win, text="Crear tabla", bootstyle=SUCCESS, command=crear).pack(padx=10, pady=(0,10))

def do_diferenciales():
    in_path = pick_csv()
    if not in_path:
        return

    try:
        out_path = Prueba_Diferenciales.run(in_path)
        messagebox.showinfo("Listo", f"Archivo generado:\n{out_path}")
        reveal_in_explorer(out_path)
    except Exception as e:
        messagebox.showerror("Error", str(e))

def do_tabla_continuidad():
    win = ttk.Toplevel(app)
    win.title("Tabla Continuidad")

    v_al = ttk.StringVar(value="1")
    v_cir = ttk.StringVar(value="1")
    vars_cargas = []

    ttk.Label(win, text="Alimentadores").pack(anchor="w", padx=10, pady=(10,0))
    ttk.Entry(win, textvariable=v_al, width=8).pack(anchor="w", padx=10)

    ttk.Label(win, text="Circuitos").pack(anchor="w", padx=10, pady=(10,0))
    ttk.Entry(win, textvariable=v_cir, width=8).pack(anchor="w", padx=10)

    frame = ttk.Frame(win, padding=10)
    frame.pack()

    def set_cargas():
        for w in frame.winfo_children():
            w.destroy()
        vars_cargas.clear()

        try:
            n = int(v_cir.get())
        except:
            messagebox.showerror("Error", "Circuitos inválido")
            return

        for i in range(n):
            row = ttk.Frame(frame)
            row.pack()
            ttk.Label(row, text=f"C{i+1}").pack(side="left")
            v = ttk.StringVar(value="1")
            ttk.Entry(row, textvariable=v, width=6).pack(side="left", padx=6)
            vars_cargas.append(v)

    ttk.Button(win, text="Setear cargas", command=set_cargas).pack()

    def procesar():
        try:
            al = int(v_al.get())
            cargas = [int(v.get()) for v in vars_cargas]
        except:
            messagebox.showerror("Error", "Valores inválidos")
            return

        win.destroy()
        out = Tabla_Continuidad.run(al, cargas)
        messagebox.showinfo("Listo", f"Tabla creada:\n{out}")
        reveal_in_explorer(out)

    ttk.Button(win, text="Crear tabla", bootstyle=SUCCESS, command=procesar).pack(pady=8)

def do_tabla_aislamiento():
    win = ttk.Toplevel(app)
    win.title("Tabla Aislamiento")

    v_al = ttk.StringVar(value="1")
    v_cir = ttk.StringVar(value="1")

    ttk.Label(win, text="Alimentadores").pack(anchor="w", padx=10, pady=(10,0))
    ttk.Entry(win, textvariable=v_al, width=8).pack(anchor="w", padx=10)

    ttk.Label(win, text="Circuitos").pack(anchor="w", padx=10, pady=(10,0))
    ttk.Entry(win, textvariable=v_cir, width=8).pack(anchor="w", padx=10)

    def crear():
        try:
            al = int(v_al.get())
            cir = int(v_cir.get())
            if al < 0 or cir < 0:
                raise ValueError()
        except:
            messagebox.showerror("Error", "Valores inválidos (enteros >= 0).")
            return

        win.destroy()
        try:
            out_path = Tabla_Aislamiento.run(al, cir)
            messagebox.showinfo("Listo", f"Tabla creada:\n{out_path}")
            reveal_in_explorer(out_path)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    ttk.Button(win, text="Crear tabla", bootstyle=SUCCESS, command=crear).pack(padx=10, pady=10)

def do_tabla_bucle():
    win = ttk.Toplevel(app)
    win.title("Tabla Bucle - Configuración")

    v_total = ttk.StringVar(value="10")
    v_lg = ttk.BooleanVar(value=True)

    ttk.Label(win, text="Circuitos totales").pack(anchor="w", padx=10, pady=(10, 0))
    ttk.Entry(win, textvariable=v_total, width=10).pack(anchor="w", padx=10)

    ttk.Checkbutton(win, text="Incluir Línea General (R/S/T)", variable=v_lg).pack(anchor="w", padx=10, pady=(10, 0))

    def siguiente():
        try:
            total = int(v_total.get())
            if total < 1:
                raise ValueError()
        except:
            messagebox.showerror("Error", "Circuitos totales inválido (entero >= 1).")
            return

        lg = bool(v_lg.get())
        win.destroy()
        _seleccionar_monofasicos(total, lg)

    ttk.Button(win, text="Siguiente", bootstyle=SUCCESS, command=siguiente).pack(padx=10, pady=12)


def _seleccionar_monofasicos(total: int, incluir_linea_general: bool):
    win = ttk.Toplevel(app)
    win.title("Seleccionar monofásicos")

    ttk.Label(win, text="Marca los circuitos monofásicos:").pack(anchor="w", padx=10, pady=(10, 6))

    # scroll por si son muchos
    container = ttk.Frame(win)
    container.pack(fill=BOTH, expand=True, padx=10, pady=(0, 10))

    canvas = ttk.Canvas(container)
    scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
    scroll_frame = ttk.Frame(canvas)

    scroll_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill=BOTH, expand=True)
    scrollbar.pack(side="right", fill="y")

    # vars checkbox: True = monofásico
    vars_mono = []
    for i in range(1, total + 1):
        v = ttk.BooleanVar(value=False)
        row = ttk.Frame(scroll_frame)
        row.pack(fill=X, pady=2)

        ttk.Label(row, text=f"Circuito {i:02d}", width=14).pack(side="left")
        ttk.Checkbutton(row, text="Monofásico", variable=v).pack(side="left")
        vars_mono.append(v)

    def crear():
        monofasicos = [i for i, v in enumerate(vars_mono, start=1) if v.get()]

        try:
            out_path = Tabla_Bucle_Falla.run(total, monofasicos, incluir_linea_general)
            messagebox.showinfo("Listo", f"Tabla creada:\n{out_path}")
            reveal_in_explorer(out_path)
            win.destroy()
        except Exception as e:
            messagebox.showerror("Error", str(e))

    ttk.Button(win, text="Crear tabla", bootstyle=SUCCESS, command=crear).pack(padx=10, pady=(0, 10))

def do_tabla_caida():
    win = ttk.Toplevel(app)
    win.title("Tabla Caida T. - Configuración")

    v_total = ttk.StringVar(value="10")
    v_lg = ttk.BooleanVar(value=True)

    ttk.Label(win, text="Circuitos totales").pack(anchor="w", padx=10, pady=(10, 0))
    ttk.Entry(win, textvariable=v_total, width=10).pack(anchor="w", padx=10)

    ttk.Checkbutton(win, text="Incluir Línea General (R/S/T)", variable=v_lg).pack(anchor="w", padx=10, pady=(10, 0))

    def siguiente():
        try:
            total = int(v_total.get())
            if total < 1:
                raise ValueError()
        except:
            messagebox.showerror("Error", "Circuitos totales inválido (entero >= 1).")
            return

        lg = bool(v_lg.get())
        win.destroy()
        _seleccionar_monofasicos(total, lg)

    ttk.Button(win, text="Siguiente", bootstyle=SUCCESS, command=siguiente).pack(padx=10, pady=12)


def _seleccionar_monofasicos(total: int, incluir_linea_general: bool):
    win = ttk.Toplevel(app)
    win.title("Seleccionar monofásicos")

    ttk.Label(win, text="Marca los circuitos monofásicos:").pack(anchor="w", padx=10, pady=(10, 6))

    container = ttk.Frame(win)
    container.pack(fill=BOTH, expand=True, padx=10, pady=(0, 10))

    canvas = ttk.Canvas(container)
    scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
    scroll_frame = ttk.Frame(canvas)

    scroll_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill=BOTH, expand=True)
    scrollbar.pack(side="right", fill="y")

    vars_mono = []
    for i in range(1, total + 1):
        v = ttk.BooleanVar(value=False)
        row = ttk.Frame(scroll_frame)
        row.pack(fill=X, pady=2)

        ttk.Label(row, text=f"Circuito {i:02d}", width=14).pack(side="left")
        ttk.Checkbutton(row, text="Monofásico", variable=v).pack(side="left")
        vars_mono.append(v)

    def crear():
        monofasicos = [i for i, v in enumerate(vars_mono, start=1) if v.get()]

        try:
            out_path = Tabla_C_T_T.run(total, monofasicos, incluir_linea_general)
            messagebox.showinfo("Listo", f"Tabla creada:\n{out_path}")
            reveal_in_explorer(out_path)
            win.destroy()
        except Exception as e:
            messagebox.showerror("Error", str(e))

    ttk.Button(win, text="Crear tabla", bootstyle=SUCCESS, command=crear).pack(padx=10, pady=(0, 10))

def do_tablas_ec():
    win = ttk.Toplevel(app)
    win.title("Tablas Aislamiento")
    win.geometry("360x200")
    win.resizable(False, False)

    opcion = ttk.StringVar(value="simple")

    ttk.Label(
        win,
        text="¿Qué tabla quieres crear?",
        font=("Segoe UI", 11)
    ).pack(pady=(12, 8), anchor="w", padx=12)

    ttk.Radiobutton(win, text="Tabla Aislamiento", variable=opcion, value="simple").pack(anchor="w", padx=20)
    ttk.Radiobutton(win, text="Tabla Aislamiento E-C", variable=opcion, value="ec").pack(anchor="w", padx=20)
    ttk.Radiobutton(win, text="Tabla Trifásica Aislamiento E-C", variable=opcion, value="trifasica").pack(anchor="w", padx=20)
    ttk.Radiobutton(win, text="Tabla Trifásica Aislamiento", variable=opcion, value="simple tri").pack(anchor="w", padx=20)

    def continuar():
        sel = opcion.get()
        win.destroy()

        if sel == "simple":
            do_tabla_aislamiento()
        elif sel == "ec":
            do_tabla_A_E_C()
        elif sel == "simple tri":
            do_tt_riso()
        elif sel == "trifasica":
            do_tabla_riso_entre_circuitos()

    ttk.Button(win, text="Continuar", bootstyle=SUCCESS, command=continuar).pack(pady=14)

def do_tablas_con():
    win = ttk.Toplevel(app)
    win.title("Tablas Aislamiento")
    win.geometry("360x200")
    win.resizable(False, False)

    opcion = ttk.StringVar(value="simple")

    ttk.Label(
        win,
        text="¿Qué tabla quieres crear?",
        font=("Segoe UI", 11)
    ).pack(pady=(12, 8), anchor="w", padx=12)

    ttk.Radiobutton(win, text="Tabla de Continuidad", variable=opcion, value="simple").pack(anchor="w", padx=20)
    ttk.Radiobutton(win, text="Tabla Trifásica de Continuidad", variable=opcion, value="tri").pack(anchor="w", padx=20)


    def continuar():
        sel = opcion.get()
        win.destroy()

        if sel == "simple":
            do_tabla_continuidad()
        elif sel == "tri":
            do_tt_rl()

    ttk.Button(win, text="Continuar", bootstyle=SUCCESS, command=continuar).pack(pady=14)

def do_tt_rl():
    win = ttk.Toplevel(app)
    win.title("TT Rl (Rlo)")

    v_n = ttk.StringVar(value="10")
    v_lg = ttk.BooleanVar(value=True)

    ttk.Label(win, text="Cantidad de circuitos").pack(anchor="w", padx=10, pady=(10, 0))
    ttk.Entry(win, textvariable=v_n, width=10).pack(anchor="w", padx=10)

    ttk.Checkbutton(win, text="Incluir Línea General", variable=v_lg).pack(anchor="w", padx=10, pady=(8, 0))

    def crear():
        try:
            n = int(v_n.get())
            if n < 1:
                raise ValueError()
        except:
            messagebox.showerror("Error", "Cantidad de circuitos inválida (entero >= 1).")
            return

        try:
            out_path = Tabla_Continuidad_Trifasica.run(n, bool(v_lg.get()))
            messagebox.showinfo("Listo", f"Tabla creada:\n{out_path}")
            reveal_in_explorer(out_path)
            win.destroy()
        except Exception as e:
            messagebox.showerror("Error", str(e))

    ttk.Button(win, text="Crear tabla", bootstyle=SUCCESS, command=crear).pack(padx=10, pady=12)


def do_tt_riso():
    win = ttk.Toplevel(app)
    win.title("TT Riso")

    v_n = ttk.StringVar(value="10")
    v_lg = ttk.BooleanVar(value=True)

    ttk.Label(win, text="Cantidad de circuitos").pack(anchor="w", padx=10, pady=(10, 0))
    ttk.Entry(win, textvariable=v_n, width=10).pack(anchor="w", padx=10)

    ttk.Checkbutton(win, text="Incluir Línea General", variable=v_lg).pack(anchor="w", padx=10, pady=(8, 0))

    def crear():
        try:
            n = int(v_n.get())
            lg = bool(v_lg.get())
        except:
            messagebox.showerror("Error", "Ingresa un número entero válido.")
            return

        try:
            out_path = Tabla_Aislamiento_Trifasica.run(n, lg)
            messagebox.showinfo("Listo", f"Tabla creada:\n{out_path}")
            reveal_in_explorer(out_path)
            win.destroy()
        except Exception as e:
            messagebox.showerror("Error", str(e))

    ttk.Button(win, text="Crear tabla", bootstyle=SUCCESS, command=crear).pack(padx=10, pady=12)

def do_tabla_A_E_C():
    win = ttk.Toplevel(app)
    win.title("Tabla Aislamiento E-C")

    v_dif = ttk.StringVar(value="1")
    vars_nc = []

    ttk.Label(win, text="Diferenciales").pack(anchor="w", padx=10, pady=(10,0))
    ttk.Entry(win, textvariable=v_dif, width=8).pack(anchor="w", padx=10)

    frame = ttk.Frame(win, padding=10)
    frame.pack()

    def set_difs():
        for w in frame.winfo_children():
            w.destroy()
        vars_nc.clear()

        try:
            n = int(v_dif.get())
            if n < 1:
                raise ValueError()
        except:
            messagebox.showerror("Error", "Diferenciales inválido")
            return

        for i in range(n):
            row = ttk.Frame(frame)
            row.pack(pady=2)
            ttk.Label(row, text=f"Dif {i+1}").pack(side="left")
            v = ttk.StringVar(value="1")
            ttk.Entry(row, textvariable=v, width=6).pack(side="left", padx=6)
            vars_nc.append(v)

    ttk.Button(win, text="Setear diferenciales", command=set_difs).pack(padx=10, pady=6)

    
    def crear():
        try:
            nc = [int(v.get()) for v in vars_nc]
            if any(x < 1 for x in nc):
                raise ValueError()
        except:
            messagebox.showerror("Error", "Circuitos por diferencial inválidos (enteros >= 1).")
            return

        win.destroy()
        try:
            out_path = Tabla_Aislamiento_E_C.run(nc)
            messagebox.showinfo("Listo", f"Tabla creada:\n{out_path}")
            reveal_in_explorer(out_path)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    ttk.Button(win, text="Crear tabla", bootstyle=SUCCESS, command=crear).pack(padx=10, pady=(0,10))


def open_outputs():
    out_dir = base_output_dir()
    out_dir.mkdir(parents=True, exist_ok=True)
    reveal_in_explorer(out_dir)

# UI

if getattr(sys, "frozen", False):
    crear_acceso_directo_si_no_existe()

app = ttk.Window(themename="darkly")
app.title(APP_TITLE)
app.geometry("720x420")
app.minsize(720, 420)
app.resizable(True, True)

main = ttk.Frame(app, padding=14)
main.pack(fill=BOTH, expand=True)

main.columnconfigure(0, weight=1)
main.columnconfigure(1, weight=1)
main.rowconfigure(0, weight=1)
main.rowconfigure(1, weight=0)

card_pruebas = ttk.Labelframe(main, text="Pruebas", padding=12, bootstyle="light")
card_pruebas.grid(row=0, column=0, sticky="nsew", padx=(0, 8), pady=(0, 10))

ttk.Button(card_pruebas, text="Continuidad", bootstyle=PRIMARY, command=do_continuidad).pack(fill=X, pady=4)
ttk.Button(card_pruebas, text="Aislamiento", bootstyle=PRIMARY, command=do_aislamiento).pack(fill=X, pady=4)
ttk.Button(card_pruebas, text="Caída de Tensión", bootstyle=PRIMARY, command=do_caida_T).pack(fill=X, pady=4)
ttk.Button(card_pruebas, text="Prueba de Lazo", bootstyle=PRIMARY, command=do_Lazo).pack(fill=X, pady=4)
ttk.Button(card_pruebas, text="Diferenciales", bootstyle=PRIMARY, command=do_diferenciales).pack(fill=X, pady=4)

card_utils = ttk.Labelframe(main, text="Utilidades", padding=12, bootstyle="light")
card_utils.grid(row=0, column=1, sticky="nsew", padx=(8, 0), pady=(0, 10))


ttk.Button(card_utils, text="Tabla Imp. Bucle de Falla", bootstyle=PRIMARY, command=do_tabla_bucle).pack(fill=X, pady=4)
ttk.Button(card_utils, text="Tabla Caida Tensión Tri.", bootstyle=PRIMARY, command=do_tabla_caida).pack(fill=X, pady=4)
ttk.Button(card_utils, text="Tablas de Continuidad", bootstyle=PRIMARY, command=do_tablas_con).pack(fill=X, pady=4)
ttk.Button(card_utils, text="Tablas de Aislamiento", bootstyle=PRIMARY, command=do_tablas_ec).pack(fill=X, pady=4)


card_salida = ttk.Labelframe(main, text="Salida", padding=12, bootstyle="light")
card_salida.grid(row=1, column=0, columnspan=2, sticky="ew")

status_var = ttk.StringVar(value="Listo")
ttk.Label(card_salida, textvariable=status_var).pack(anchor=W)
ttk.Button(card_salida, text="Abrir carpeta OUTPUT", bootstyle=SECONDARY, command=open_outputs).pack(fill=X, pady=4)


app.mainloop()
