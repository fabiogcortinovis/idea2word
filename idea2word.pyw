# -*- coding: utf-8 -*-
"""
idea2word.pyw
Compone in proiezione ortogonale le 4 viste del nodo (3D, Z, Y, X) contenute
nel PDF di verifica di IDEA StatiCa Connection.

Avvio: doppio click (nessun terminale). Gli errori compaiono in finestre di
messaggio e vengono scritti in idea2word_errori.log accanto allo script.
Al primo avvio installa da solo le librerie mancanti.
"""
import os
import sys
import math
import traceback
import subprocess
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
LOG = os.path.join(HERE, "idea2word_errori.log")

# ----------------------------------------------------------------- parametri
PPM = 1.0          # px di uscita per mm reale (usato dalla calibrazione)
GAP = 30           # spazio tra le viste [px]
MARGIN = 20        # margine esterno [px]
THR = 40           # soglia "non bianco" (0-255) per il ritaglio
DIL = 4            # dilatazione per unire le parti di una stessa vista [px]
THR_SOFT = 8       # soglia "sensibile": cattura anche grigi molto chiari (calcestruzzo)
MIN_IMG = 300      # lato minimo per considerare un'immagine una "vista"

KEYS = ["3d", "z", "y", "x"]   # ordine di comparsa nel PDF
NAMES = {"3d": "Vista 3D", "z": "Pianta Z", "y": "Vista Y", "x": "Vista X"}
# posizione dei pannelli nella GUI = posizione delle viste nel risultato finale
GRID = {"x": (0, 0), "y": (0, 1), "3d": (1, 0), "z": (1, 1)}


# ----------------------------------------------------------- utilità errori
def log_error(text):
    try:
        with open(LOG, "a", encoding="utf-8") as f:
            f.write(text + "\n" + "-" * 60 + "\n")
    except Exception:
        pass


def show_error(title, text):
    log_error(f"{title}\n{text}")
    try:
        import tkinter as tk
        from tkinter import messagebox
        r = tk.Tk()
        r.withdraw()
        messagebox.showerror(title, text)
        r.destroy()
    except Exception:
        pass


# ---------------------------------------------------- installazione librerie
def _python_exe():
    exe = sys.executable
    if exe.lower().endswith("pythonw.exe"):
        cand = exe[:-len("pythonw.exe")] + "python.exe"
        if os.path.exists(cand):
            return cand
    return exe


def ensure_deps():
    needed = [("fitz", "pymupdf"), ("PIL", "pillow"),
              ("numpy", "numpy"), ("scipy", "scipy")]
    if os.name == "nt":
        needed.append(("win32com", "pywin32"))
    missing = []
    for mod, pkg in needed:
        try:
            __import__(mod)
        except ImportError:
            missing.append(pkg)
    if not missing:
        return
    import tkinter as tk
    r = tk.Tk()
    r.title("idea2word")
    tk.Label(r, padx=30, pady=20,
             text="Installazione librerie in corso:\n" + ", ".join(missing)
             + "\n\nAttendi, può richiedere un minuto...").pack()
    r.update()
    flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    err = ""
    for extra in ([], ["--user"]):
        p = subprocess.run([_python_exe(), "-m", "pip", "install", *extra,
                            *missing], capture_output=True, text=True,
                           creationflags=flags)
        if p.returncode == 0:
            err = ""
            break
        err = p.stdout[-800:] + p.stderr[-800:]
    r.destroy()
    if err:
        raise RuntimeError("Installazione librerie non riuscita "
                           "(controlla la connessione internet):\n" + err)


# ---------------------------------------------------------------- elaborazione
def load_views(pdf_path):
    """Restituisce le prime 4 immagini 'grandi' del PDF (PIL RGB), in ordine."""
    import fitz
    from PIL import Image
    doc = fitz.open(pdf_path)
    seen, out = set(), []
    for page in doc:
        for im in page.get_images(full=True):
            xref, w, h = im[0], im[2], im[3]
            if xref in seen or w < MIN_IMG or h < MIN_IMG:
                continue
            seen.add(xref)
            pix = fitz.Pixmap(doc, xref)
            if pix.n - pix.alpha >= 4:          # CMYK -> RGB
                pix = fitz.Pixmap(fitz.csRGB, pix)
            if pix.alpha:
                pix = fitz.Pixmap(pix, 0)       # scarta alpha
            img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
            out.append(img)
            if len(out) == 4:
                return out
    raise RuntimeError(
        f"Nel PDF ho trovato solo {len(out)} immagini di vista su 4.\n"
        "Il PDF non sembra un output IDEA StatiCa con le 4 viste "
        "incorporate come immagini.")


def crop_view(img):
    """Ritaglia sul contenuto.
    Simbolo degli assi: se e' separato dall'immagine da spazio bianco viene
    eliminato; se tocca/si sovrappone a parti dell'immagine (anche grigio molto
    chiaro, es. calcestruzzo) viene lasciato. Il ritaglio resta stretto
    sull'immagine in entrambi i casi."""
    import numpy as np
    from PIL import Image
    from scipy import ndimage as ndi
    arr = np.asarray(img.convert("RGB"), dtype=np.int16)
    diff = 255 - arr.min(axis=2)
    mask = diff > THR
    if not mask.any():
        return img
    # contenuto "sensibile": include grigi chiari, ma non le linee sottili
    sens = ndi.binary_opening(diff > THR_SOFT, structure=np.ones((3, 3))) | mask
    h, w = mask.shape
    lab, n = ndi.label(ndi.binary_dilation(mask, iterations=DIL),
                       structure=np.ones((3, 3)))
    drop = np.zeros_like(mask)
    for i, sl in enumerate(ndi.find_objects(lab), 1):
        ys, xs = sl
        bw, bh = xs.stop - xs.start, ys.stop - ys.start
        cx, cy = (xs.start + xs.stop) / 2, (ys.start + ys.stop) / 2
        cand = (bw < 0.2 * w and bh < 0.2 * h
                and cx < 0.35 * w and cy > 0.65 * h)
        if not cand:
            continue
        comp = lab == i
        zone = ndi.binary_dilation(comp, iterations=3)
        touches = (sens & ~comp & zone).any()   # c'e' altro contenuto attaccato?
        if not touches:
            drop |= comp
    content = sens & ~drop
    if not content.any():
        content = mask
    out = arr.copy()
    out[drop] = 255
    ys, xs = np.where(content)
    x0, x1 = max(xs.min() - 2, 0), min(xs.max() + 3, w)
    y0, y1 = max(ys.min() - 2, 0), min(ys.max() + 3, h)
    return Image.fromarray(out.astype(np.uint8)).crop((x0, y0, x1, y1))


def auto_scales(imgs):
    """Scala iniziale: Y di riferimento, X e Z allineate per estensione."""
    s = {"3d": 1.0, "y": 1.0}
    wy, hy = imgs["y"].size
    s["x"] = hy / imgs["x"].size[1]     # stessa quota Z
    s["z"] = wy / imgs["z"].size[0]     # stessa ascissa X
    return s


def compose(imgs, scales, offs):
    """Compone la tavola. offs[k] = (dx, dy) in px. Ritorna immagine PIL."""
    from PIL import Image
    sized = {}
    for k in KEYS:
        im = imgs[k]
        w = max(1, round(im.size[0] * scales[k]))
        h = max(1, round(im.size[1] * scales[k]))
        sized[k] = im.resize((w, h), Image.LANCZOS)
    wy, hy = sized["y"].size
    wx, hx = sized["x"].size
    wz, hz = sized["z"].size
    w3, h3 = sized["3d"].size
    pos = {"y": (0, 0)}
    pos["x"] = (-GAP - wx, (hy - hx) / 2)                 # allineata in Z
    pos["z"] = ((wy - wz) / 2, hy + GAP)                  # allineata in X
    pos["3d"] = (pos["z"][0] - GAP - w3, pos["z"][1] + (hz - h3) / 2)
    for k in KEYS:
        pos[k] = (pos[k][0] + offs[k][0], pos[k][1] + offs[k][1])
    x0 = min(pos[k][0] for k in KEYS)
    y0 = min(pos[k][1] for k in KEYS)
    x1 = max(pos[k][0] + sized[k].size[0] for k in KEYS)
    y1 = max(pos[k][1] + sized[k].size[1] for k in KEYS)
    W = int(math.ceil(x1 - x0)) + 2 * MARGIN
    H = int(math.ceil(y1 - y0)) + 2 * MARGIN
    canvas = Image.new("RGB", (W, H), "white")
    for k in KEYS:
        canvas.paste(sized[k], (round(pos[k][0] - x0 + MARGIN),
                                round(pos[k][1] - y0 + MARGIN)))
    return canvas


def insert_in_word(png_path):
    """Inserisce l'immagine nel documento Word aperto, alla posizione del cursore."""
    try:
        import win32com.client
    except ImportError:
        raise RuntimeError("Libreria pywin32 non disponibile: riavvia il programma.")
    try:
        word = win32com.client.GetActiveObject("Word.Application")
    except Exception:
        raise RuntimeError("Word non è aperto. Apri il documento, posiziona "
                           "il cursore dove vuoi l'immagine e riprova.")
    if word.Documents.Count == 0:
        raise RuntimeError("Non c'è nessun documento Word aperto.")
    sel = word.Selection
    sel.Collapse(0)
    shape = sel.InlineShapes.AddPicture(FileName=png_path, LinkToFile=False,
                                        SaveWithDocument=True)
    ps = word.ActiveDocument.PageSetup
    usable = ps.PageWidth - ps.LeftMargin - ps.RightMargin
    if shape.Width > usable:
        shape.LockAspectRatio = True
        shape.Width = usable


# ------------------------------------------------------------------------ GUI
def run_gui():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, simpledialog
    from PIL import ImageTk

    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    root = tk.Tk()
    root.title("idea2word")
    root.geometry("1420x800")

    def report(exc, tb, tbk):
        show_error("Errore", "".join(traceback.format_exception(exc, tb, tbk)))
    root.report_callback_exception = report

    st = {"pdf": None, "imgs": None, "photo": None, "job": None, "comp": None}
    scale_v = {k: tk.DoubleVar(value=1.0) for k in KEYS}
    dx_v = {k: tk.IntVar(value=0) for k in KEYS}
    dy_v = {k: tk.IntVar(value=0) for k in KEYS}

    # barra pulsanti
    bar = ttk.Frame(root, padding=6)
    bar.pack(fill="x")
    lbl = ttk.Label(bar, text="Nessun PDF aperto")
    main = ttk.Frame(root)
    main.pack(fill="both", expand=True)
    canvas = tk.Canvas(main, bg="#e8e8e8", highlightthickness=0)
    canvas.pack(side="left", fill="both", expand=True)
    panel = ttk.Frame(main, padding=6)
    panel.pack(side="right", fill="y")
    status = ttk.Label(root, anchor="w", padding=4,
                       text="Apri un PDF di IDEA StatiCa per iniziare.")
    status.pack(fill="x")

    def getf(var, default):
        try:
            return float(var.get())
        except (tk.TclError, ValueError):
            return default

    last = {k: 1.0 for k in KEYS}

    def params():
        sc = {}
        for k in KEYS:
            v = getf(scale_v[k], last[k])
            if v <= 0:
                v = last[k]
            last[k] = sc[k] = v
        of = {k: (int(getf(dx_v[k], 0)), int(getf(dy_v[k], 0))) for k in KEYS}
        return sc, of

    def render():
        st["job"] = None
        if not st["imgs"]:
            return
        sc, of = params()
        comp = compose(st["imgs"], sc, of)
        st["comp"] = comp
        cw, ch = max(canvas.winfo_width(), 50), max(canvas.winfo_height(), 50)
        z = min(cw / comp.size[0], ch / comp.size[1], 1.0)
        disp = comp.resize((max(1, int(comp.size[0] * z)),
                            max(1, int(comp.size[1] * z))))
        st["photo"] = ImageTk.PhotoImage(disp)
        canvas.delete("all")
        canvas.create_image(cw // 2, ch // 2, image=st["photo"])
        status.config(text=f"Tavola {comp.size[0]}×{comp.size[1]} px  -  "
                           f"anteprima al {z * 100:.0f}%")

    def refresh(*_):
        if st["job"]:
            root.after_cancel(st["job"])
        st["job"] = root.after(60, render)

    canvas.bind("<Configure>", refresh)

    def calibrate(key):
        if not st["imgs"]:
            return
        img = st["imgs"][key]
        win = tk.Toplevel(root)
        win.title(f"Calibra {NAMES[key]}: clicca 2 punti di distanza nota")
        z = min(1000 / img.size[0], 720 / img.size[1])
        z = max(z, 0.5)
        disp = img.resize((int(img.size[0] * z), int(img.size[1] * z)))
        ph = ImageTk.PhotoImage(disp)
        ttk.Label(win, padding=4, text="Clic sinistro: 2 punti (es. estremi "
                  "dell'ala o dell'altezza HEB360). Clic destro: azzera.").pack()
        cv = tk.Canvas(win, width=disp.size[0], height=disp.size[1])
        cv.pack()
        cv.create_image(0, 0, image=ph, anchor="nw")
        cv.image = ph
        pts = []

        def draw():
            cv.delete("m")
            for (x, y) in pts:
                cv.create_line(x - 8, y, x + 8, y, fill="red", tags="m")
                cv.create_line(x, y - 8, x, y + 8, fill="red", tags="m")
            if len(pts) == 2:
                cv.create_line(*pts[0], *pts[1], fill="red", dash=(4, 2), tags="m")

        def click(e):
            if len(pts) >= 2:
                return
            pts.append((e.x, e.y))
            draw()
            if len(pts) == 2:
                d = math.hypot(pts[0][0] - pts[1][0], pts[0][1] - pts[1][1]) / z
                D = simpledialog.askfloat(
                    "Distanza reale", "Distanza reale tra i due punti [mm]:",
                    parent=win, minvalue=0.001)
                if D is None:
                    pts.clear()
                    draw()
                    return
                if d < 2:
                    messagebox.showwarning("Punti troppo vicini",
                                           "I due punti coincidono.", parent=win)
                    pts.clear()
                    draw()
                    return
                scale_v[key].set(round(D * PPM / d, 4))
                win.destroy()
                refresh()

        def reset(_e):
            pts.clear()
            draw()
        cv.bind("<Button-1>", click)
        cv.bind("<Button-3>", reset)
        win.transient(root)
        win.grab_set()

    # pannelli: disposti come le viste nel risultato finale (griglia 2x2)
    for k in KEYS:
        r, c = GRID[k]
        f = ttk.LabelFrame(panel, text=NAMES[k], padding=6)
        f.grid(row=r, column=c, padx=3, pady=3, sticky="nsew")
        r1 = ttk.Frame(f)
        r1.pack(fill="x")
        ttk.Label(r1, text="Scala").pack(side="left")
        ttk.Spinbox(r1, textvariable=scale_v[k], from_=0.01, to=20,
                    increment=0.01, format="%.3f", width=7,
                    command=refresh).pack(side="left", padx=(4, 2))
        tk.Scale(f, variable=scale_v[k], from_=0.05, to=10, resolution=0.01,
                 orient="horizontal", showvalue=False, length=190,
                 command=lambda _v: refresh()).pack(fill="x")
        scale_v[k].trace_add("write", lambda *_: refresh())
        r2 = ttk.Frame(f)
        r2.pack(fill="x", pady=2)
        for txt, var in (("dx", dx_v[k]), ("dy", dy_v[k])):
            ttk.Label(r2, text=txt).pack(side="left")
            ttk.Spinbox(r2, textvariable=var, from_=-2000, to=2000, width=6,
                        command=refresh).pack(side="left", padx=(2, 8))
            var.trace_add("write", lambda *_: refresh())
        ttk.Button(f, text="Calibra con 2 punti",
                   command=lambda kk=k: calibrate(kk)).pack(fill="x", pady=2)

    ttk.Label(panel, wraplength=420, foreground="#555",
              text="I pannelli sono disposti come le viste nel risultato. "
                   "La scala si imposta con lo slider o scrivendo il valore. "
                   "Y è di riferimento: X e Z restano allineate al suo centro. "
                   "Per una scala reale calibra tutte e 4 le viste."
              ).grid(row=2, column=0, columnspan=2, pady=8)

    def open_pdf(path=None):
        path = path or filedialog.askopenfilename(
            title="Seleziona il PDF di IDEA StatiCa",
            filetypes=[("PDF", "*.pdf")])
        if not path:
            return
        raw = load_views(path)
        imgs = {k: crop_view(im) for k, im in zip(KEYS, raw)}
        st["pdf"], st["imgs"] = path, imgs
        for k, v in auto_scales(imgs).items():
            scale_v[k].set(round(v, 3))
        for k in KEYS:
            dx_v[k].set(0)
            dy_v[k].set(0)
        lbl.config(text=os.path.basename(path))
        render()

    def reset_auto():
        if st["imgs"]:
            for k, v in auto_scales(st["imgs"]).items():
                scale_v[k].set(round(v, 3))
            for k in KEYS:
                dx_v[k].set(0)
                dy_v[k].set(0)
            refresh()

    def export_png():
        if not st["imgs"]:
            return messagebox.showinfo("Nessun PDF", "Apri prima un PDF.")
        base = os.path.splitext(os.path.basename(st["pdf"]))[0]
        out = filedialog.asksaveasfilename(
            defaultextension=".png", filetypes=[("PNG", "*.png")],
            initialdir=os.path.dirname(st["pdf"]),
            initialfile=base + "_proiezioni.png")
        if not out:
            return
        sc, of = params()
        compose(st["imgs"], sc, of).save(out, dpi=(150, 150))
        status.config(text="Salvato: " + out)

    def to_word():
        if not st["imgs"]:
            return messagebox.showinfo("Nessun PDF", "Apri prima un PDF.")
        sc, of = params()
        fd, tmp = tempfile.mkstemp(suffix=".png")
        os.close(fd)
        try:
            compose(st["imgs"], sc, of).save(tmp, dpi=(150, 150))
            insert_in_word(tmp)
            status.config(text="Immagine inserita in Word alla posizione del cursore.")
        finally:
            try:
                os.remove(tmp)
            except OSError:
                pass

    ttk.Button(bar, text="Apri PDF", command=open_pdf).pack(side="left")
    lbl.pack(side="left", padx=10)
    ttk.Button(bar, text="Inserisci in Word (cursore)",
               command=to_word).pack(side="right", padx=3)
    ttk.Button(bar, text="Esporta PNG", command=export_png).pack(side="right", padx=3)
    ttk.Button(bar, text="Scale automatiche",
               command=reset_auto).pack(side="right", padx=3)

    if len(sys.argv) > 1 and sys.argv[1].lower().endswith(".pdf"):
        root.after(200, lambda: open_pdf(sys.argv[1]))
    root.mainloop()


def main():
    try:
        ensure_deps()
        run_gui()
    except Exception:
        show_error("idea2word - errore", traceback.format_exc())


if __name__ == "__main__":
    main()
