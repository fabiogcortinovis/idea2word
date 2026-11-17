# -*- coding: utf-8 -*-
"""
idea2word.pyw
Due schermate, per il PDF di verifica di IDEA StatiCa Connection:
 1) Viste ortogonali: compone le 4 viste del nodo (3D, Z, Y, X).
 2) Relazione: ritaglia e reimpagina le pagine (da "Membrature" in poi, senza le
    4 viste) in pagine di uguale formato max 170 x 240 mm da incollare in Word.

Avvio: doppio click (nessun terminale). Gli errori compaiono in finestre di
messaggio e vengono scritti in idea2word_errori.log accanto allo script.
Al primo avvio installa da solo le librerie mancanti.
"""
import os
import re
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
    # adatta alla larghezza utile: usa la scala in % (piu' robusta di Width)
    # e non fa fallire l'inserimento se Word rifiuta il ridimensionamento
    try:
        shape.LockAspectRatio = -1          # proporzioni sempre bloccate
    except Exception:
        log_error("Blocco proporzioni non riuscito:\n" + traceback.format_exc())
    try:
        ps = word.ActiveDocument.PageSetup
        usable = float(ps.PageWidth - ps.LeftMargin - ps.RightMargin)
        width = float(shape.Width)
        if usable > 50 and width > usable:
            pct = max(1.0, min(100.0, usable / width * 100.0))
            shape.ScaleWidth = pct
            shape.ScaleHeight = pct
    except Exception:
        log_error("Ridimensionamento in Word non riuscito:\n"
                  + traceback.format_exc())


# =============================================================================
# SCHERMATA 2 - RELAZIONE: impaginazione delle pagine del PDF per Word
# =============================================================================
OUT_W_MM = 170.0       # larghezza massima della pagina in Word [mm]
OUT_H_MM = 240.0       # altezza massima della pagina in Word [mm]
GAP_MAX_MM = 5.0       # spazio massimo tra due blocchi consecutivi [mm]
REN_DPI = 200          # risoluzione delle pagine prodotte
AN_DPI = 100           # risoluzione dell'analisi (solo per trovare gli spazi)
NOWHITE = 247          # grigio (0-255) sotto cui un pixel e' "contenuto"
MIN_BREAK_GAP = 3.0    # [pt] spazi piu' piccoli non sono mai punti di taglio
GAP_REF = 22.0         # [pt] spazio "da capitolo": taglio gratuito
PEN_GRID = 0.55        # costo di un taglio tra due righe di una tabella
PEN_SEAM_TAB = 0.30    # costo di un taglio dove la tabella continua tra pagine
PAGE_COST = 0.25       # costo fisso per ogni pagina (evita pagine inutili)
WASTE_W = 1.0          # peso dello spazio vuoto a fondo pagina
IMG_NOBREAK = 150      # le immagini raster >= 150 px non vengono mai tagliate


class PdfDoc:
    """Accesso al PDF tramite PyMuPDF (stessa interfaccia usata dall'impaginatore)."""

    def __init__(self, path):
        import fitz
        self.fitz = fitz
        self.doc = fitz.open(path)
        self.n = len(self.doc)

    def size(self, i):
        r = self.doc[i].rect
        return r.width, r.height

    def words(self, i):
        """[(x0, y0, x1, y1, testo, blocco, riga)]"""
        return [(w[0], w[1], w[2], w[3], w[4], w[5], w[6])
                for w in self.doc[i].get_text("words")]

    def images(self, i):
        """[(x0, y0, x1, y1, larghezza_px, altezza_px)] nell'ordine del PDF."""
        page = self.doc[i]
        out = []
        for im in page.get_images(full=True):
            xref, w, h = im[0], im[2], im[3]
            try:
                rects = page.get_image_rects(xref)
            except Exception:
                rects = []
            for r in rects:
                out.append((r.x0, r.y0, r.x1, r.y1, w, h))
        return out

    def render(self, i, clip, dpi, gray=False):
        from PIL import Image
        fitz = self.fitz
        z = dpi / 72.0
        pix = self.doc[i].get_pixmap(
            matrix=fitz.Matrix(z, z), clip=fitz.Rect(*clip), alpha=False,
            colorspace=fitz.csGRAY if gray else fitz.csRGB)
        return Image.frombytes("L" if gray else "RGB",
                               (pix.width, pix.height), pix.samples)


def _runs(flags):
    """Intervalli (inizio, fine inclusa) di valori True consecutivi."""
    out, s = [], None
    for i, v in enumerate(flags):
        if v and s is None:
            s = i
        elif not v and s is not None:
            out.append((s, i - 1))
            s = None
    if s is not None:
        out.append((s, len(flags) - 1))
    return out


def _row_info(row, wpx):
    """Se la riga di pixel e' un filetto orizzontale (bordo di tabella)
    ritorna (x_inizio, x_fine, grigio_medio), altrimenti None."""
    import numpy as np
    nzm = row < NOWHITE
    cnt = int(nzm.sum())
    if cnt < 0.5 * wpx:
        return None
    cols = np.flatnonzero(nzm)
    f, l = int(cols[0]), int(cols[-1])
    if cnt < 0.97 * (l - f + 1):
        return None
    vals = row[nzm]
    if vals.min() >= 150 and vals.mean() <= 232:
        return f, l, float(vals.mean())
    return None


def _grid_lines(g, a, b, wpx):
    """Filetti orizzontali (sottili) tra le righe a..b della pagina g."""
    info = [_row_info(g[r], wpx) for r in range(a, b + 1)]
    lines, r, n = [], 0, len(info)
    while r < n:
        if info[r] is None:
            r += 1
            continue
        s = r
        while r < n and info[r] is not None:
            r += 1
        if r - s <= 4:
            part = info[s:r]
            lines.append({"ra": a + s, "rb": a + r - 1,
                          "x0": min(p[0] for p in part),
                          "x1": max(p[1] for p in part) + 1,
                          "gray": sum(p[2] for p in part) / len(part)})
    return lines


def find_report_layout(doc):
    """Per ogni pagina: zona utile (senza intestazione e numero di pagina) e
    rettangoli da escludere (le prime 4 viste con le loro didascalie)."""
    excl, found = {}, 0
    for i in range(doc.n):
        for (x0, y0, x1, y1, w, h) in doc.images(i):
            if w < MIN_IMG or h < MIN_IMG:
                continue
            excl.setdefault(i, []).append((x0 - 1, y0 - 1, x1 + 1, y1 + 1))
            found += 1
            if found == 4:
                break
        if found == 4:
            break
    for i in list(excl):
        lines = {}
        for (x0, y0, x1, y1, t, b, l) in doc.words(i):
            lines.setdefault((b, l), []).append((x0, y0, x1, y1, t))
        for ws in lines.values():
            txt = " ".join(w[4] for w in ws)
            if re.search(r"Vista (assonometrica|in direzione [XYZ])", txt):
                excl[i].append((min(w[0] for w in ws) - 1, min(w[1] for w in ws) - 1,
                                max(w[2] for w in ws) + 1, max(w[3] for w in ws) + 1))
    bounds = []
    for i in range(doc.n):
        W, H = doc.size(i)
        ws = doc.words(i)
        top, bottom = 0.096 * H, 0.955 * H
        for w in ws:                                  # fine intestazione
            if w[4] == "Autore:" and w[1] < 0.15 * H:
                top = w[3] + 2
        low = [w[1] for w in ws if w[1] > 0.94 * H]   # numero di pagina
        if low:
            bottom = min(low) - 1
        if i == 0:                                    # si parte da "Membrature"
            for w in ws:
                if w[4] == "Membrature" and w[1] > top:
                    top = w[1] - 2
                    break
        bounds.append((top, bottom))
    return {"bounds": bounds, "excl": excl}


def analyze_pages(doc, lay):
    """Rende ogni pagina (solo zona utile) in grigi a bassa risoluzione."""
    import numpy as np
    S = 72.0 / AN_DPI
    pages = []
    for i in range(doc.n):
        W, H = doc.size(i)
        top, bot = lay["bounds"][i]
        g = np.array(doc.render(i, (0, top, W, bot), AN_DPI, gray=True),
                     dtype=np.uint8)
        for (x0, y0, x1, y1) in lay["excl"].get(i, []):
            a, b = max(0, int((y0 - top) / S)), min(g.shape[0], int(math.ceil((y1 - top) / S)))
            c, d = max(0, int(x0 / S)), min(g.shape[1], int(math.ceil(x1 / S)))
            if b > a and d > c:
                g[a:b, c:d] = 255
        pages.append({"g": g, "top": top, "S": S})
    return pages


def build_blocks(doc, lay, pages):
    """Spezza il documento in blocchi indivisibili (righe di testo, righe di
    tabella, immagini) ognuno con il costo di un'eventuale interruzione prima
    di esso. Ritorna (blocchi, X0, X1): X0..X1 e' il riquadro comune a tutte
    le pagine (il piu' largo)."""
    import numpy as np
    S = pages[0]["S"]
    cmin, cmax = None, None
    for pg in pages:
        cols = np.flatnonzero((pg["g"] < NOWHITE).any(axis=0))
        if cols.size:
            cmin = cols[0] if cmin is None else min(cmin, cols[0])
            cmax = cols[-1] if cmax is None else max(cmax, cols[-1])
    if cmin is None:
        raise RuntimeError("Non ho trovato contenuto nelle pagine del PDF.")
    X0 = max(0.0, cmin * S - 2)
    X1 = min(doc.size(0)[0], (cmax + 1) * S + 2)
    wpx = (X1 - X0) / S
    thr = int(math.ceil(MIN_BREAK_GAP / S))
    blocks, prev = [], None
    for i, pg in enumerate(pages):
        g, top = pg["g"], pg["top"]
        runs = _runs((g < NOWHITE).any(axis=1))
        merged = []
        for r in runs:                                   # unisce spazi minuscoli
            if merged and r[0] - merged[-1][1] - 1 < thr:
                merged[-1] = (merged[-1][0], r[1])
            else:
                merged.append(r)
        runs = merged
        if not runs:
            continue
        zones = [((y0 - top) / S, (y1 - top) / S)
                 for (x0, y0, x1, y1, w, h) in doc.images(i)
                 if min(w, h) >= IMG_NOBREAK]

        def inzone(r, zones=zones):
            return any(z0 + 1 < r < z1 - 1 for z0, z1 in zones)
        pb = []
        for ri, (a, b) in enumerate(runs):
            if ri > 0:
                prev_b = runs[ri - 1][1]
                gap = (a - prev_b - 1) * S
                if inzone((a + prev_b) / 2) and pb:      # spazio dentro un'immagine
                    pb[-1]["y1"] = top + (b + 1) * S
                    continue
                join = {"join": "gap", "gap": gap, "pen": max(0.0, 1 - gap / GAP_REF)}
            else:
                join = {"join": "gap", "gap": GAP_REF, "pen": 0.0}
            lines = [ln for ln in _grid_lines(g, a, b, wpx) if not inzone(ln["ra"])]
            starts_line = bool(lines) and lines[0]["ra"] - a <= 2
            cand = lines[2:] if starts_line else lines     # mai dopo la riga di intestazione
            cuts, last = [], a
            for ln in cand:
                if ln["ra"] - last >= 4 and b - ln["rb"] >= 3:
                    cuts.append(ln)
                    last = ln["ra"]
            starts = [a] + [c["ra"] for c in cuts]
            ends = [c["ra"] for c in cuts] + [b + 1]
            for pi, (s0, s1) in enumerate(zip(starts, ends)):
                blk = {"page": i, "y0": top + s0 * S, "y1": top + s1 * S}
                if pi == 0:
                    blk.update(join)
                    blk["line"] = None
                else:
                    c = cuts[pi - 1]
                    blk.update(join="grid", gap=0.0, pen=PEN_GRID,
                               line=(c["x0"] * S, c["x1"] * S, c["gray"], max(S, (c["rb"] - c["ra"] + 1) * S)))
                pb.append(blk)
        if not pb:
            continue
        if prev is not None:                              # giunzione tra pagine
            gA, bA = pages[prev[0]]["g"], prev[1]
            a0 = runs[0][0]
            infoB = _row_info(g[a0], wpx)
            lineA = any(_row_info(gA[r], wpx) for r in range(max(0, bA - 2), bA + 1))
            lineB = any(_row_info(g[r], wpx) for r in range(a0, min(a0 + 3, g.shape[0])))
            if lineA and lineB and infoB:
                f, l, mean = infoB                        # tabella che continua:
                ext, r = l - f + 1, a0                    # salta l'intestazione ripetuta
                while r <= runs[0][1] and (g[r] < NOWHITE).sum() >= 0.9 * ext:
                    r += 1
                pb[0]["y0"] = top + r * S
                if pb[0]["y1"] - pb[0]["y0"] < 2 * S:
                    pb.pop(0)
                if pb:
                    pb[0].update(join="seamtab", gap=0.0, pen=PEN_SEAM_TAB,
                                 line=(f * S, (l + 1) * S, mean, S))
        if pb:
            blocks.extend(pb)
        prev = (i, runs[-1][1])
    if not blocks:
        raise RuntimeError("Nessun contenuto da impaginare.")
    blocks[0].update(join=None, gap=0.0, pen=0.0)
    return blocks, X0, X1


def layout_report(blocks, X0, X1, out_w_mm, out_h_mm, gap_max_pt):
    """Sceglie scala e interruzioni di pagina (programmazione dinamica).
    Ritorna (k, altezza_pagina_in_pt_sorgente, pagine) con pagine = liste di
    indici di blocchi. k = pt di uscita per pt sorgente: e' il vincolo piu'
    limitante tra larghezza e altezza del blocco piu' alto."""
    Wout, Hout = out_w_mm / 25.4 * 72, out_h_mm / 25.4 * 72
    hmax = max(b["y1"] - b["y0"] for b in blocks)
    k = min(Wout / (X1 - X0), Hout / hmax)
    Hsrc = Hout / k
    n, INF = len(blocks), 1e18
    f, prv = [INF] * (n + 1), [0] * (n + 1)
    f[0] = 0.0
    for i in range(n):
        if f[i] >= INF:
            continue
        used = 0.0
        for j in range(i, n):
            b = blocks[j]
            if j > i:
                used += min(b["gap"], gap_max_pt)
            used += b["y1"] - b["y0"]
            if used > Hsrc + 1e-6:
                break
            cost = f[i] + PAGE_COST
            if j < n - 1:
                cost += WASTE_W * ((Hsrc - used) / Hsrc) ** 2 + blocks[j + 1]["pen"]
            if cost < f[j + 1]:
                f[j + 1], prv[j + 1] = cost, i
    pages, j = [], n
    while j > 0:
        pages.append(list(range(prv[j], j)))
        j = prv[j]
    pages.reverse()
    return k, Hsrc, pages


def render_report(doc, lay, blocks, pages, X0, X1, k, out_h_mm, gap_max_pt, dpi):
    """Disegna le pagine finali (PIL RGB, tutte della stessa dimensione)."""
    from PIL import Image, ImageDraw
    z = k * dpi / 72.0                       # px di uscita per pt sorgente
    Wpx = int(round((X1 - X0) * z))
    Hpx = int(round(out_h_mm / 25.4 * dpi))
    out = []
    for pg in pages:
        canvas = Image.new("RGB", (Wpx, Hpx), "white")
        draw = ImageDraw.Draw(canvas)
        y = 0.0
        groups = []                           # blocchi contigui -> un solo rendering
        for idx, bi in enumerate(pg):
            b = blocks[bi]
            if idx > 0:
                y += min(b["gap"], gap_max_pt) * z
            if (groups and b["join"] == "grid" and groups[-1]["page"] == b["page"]
                    and abs(groups[-1]["y1"] - b["y0"]) < 1e-6):
                groups[-1]["y1"] = b["y1"]
            else:
                groups.append({"page": b["page"], "y0": b["y0"], "y1": b["y1"], "ypx": y})
            y += (b["y1"] - b["y0"]) * z
        for gp in groups:
            eh = max(1, int(round((gp["y1"] - gp["y0"]) * z)))
            im = doc.render(gp["page"], (X0, gp["y0"], X1, gp["y1"]), 72.0 * z)
            base = Image.new("RGB", (Wpx, eh), "white")   # niente bordi neri da ritaglio
            base.paste(im.crop((0, 0, min(im.size[0], Wpx), min(im.size[1], eh))), (0, 0))
            im = base
            d2 = ImageDraw.Draw(im)
            for (x0, y0, x1, y1) in lay["excl"].get(gp["page"], []):
                d2.rectangle([(x0 - X0) * z, (y0 - gp["y0"]) * z,
                              (x1 - X0) * z, (y1 - gp["y0"]) * z], fill="white")
            canvas.paste(im, (0, int(round(gp["ypx"]))))
        # bordi tabella ricostruiti dove la pagina taglia una tabella
        first = blocks[pg[0]]
        if first["join"] == "seamtab" and first["line"]:
            lx0, lx1, gv, th = first["line"]
            draw.rectangle([(lx0 - X0) * z, 0, (lx1 - X0) * z, max(1, round(th * z)) - 1],
                           fill=(int(gv),) * 3)
        nxt = pg[-1] + 1
        if nxt < len(blocks) and blocks[nxt]["join"] == "grid" and blocks[nxt]["line"]:
            lx0, lx1, gv, th = blocks[nxt]["line"]
            draw.rectangle([(lx0 - X0) * z, y, (lx1 - X0) * z, y + max(1, round(th * z)) - 1],
                           fill=(int(gv),) * 3)
        out.append(canvas)
    return out


def insert_pages_in_word(paths, sizes_mm, page_break=True):
    """Inserisce le immagini una dopo l'altra nel documento Word aperto, a
    partire dal cursore, con dimensione reale in mm."""
    try:
        import win32com.client
    except ImportError:
        raise RuntimeError("Libreria pywin32 non disponibile: riavvia il programma.")
    try:
        word = win32com.client.GetActiveObject("Word.Application")
    except Exception:
        raise RuntimeError("Word non è aperto. Apri il documento, posiziona "
                           "il cursore dove vuoi le pagine e riprova.")
    if word.Documents.Count == 0:
        raise RuntimeError("Non c'è nessun documento Word aperto.")
    sel = word.Selection
    sel.Collapse(0)
    for n, (p, (w_mm, h_mm)) in enumerate(zip(paths, sizes_mm)):
        shape = sel.InlineShapes.AddPicture(FileName=p, LinkToFile=False,
                                            SaveWithDocument=True)
        try:
            shape.LockAspectRatio = -1
            shape.Width = w_mm / 25.4 * 72
            shape.Height = h_mm / 25.4 * 72
        except Exception:
            log_error("Ridimensionamento in Word non riuscito:\n" + traceback.format_exc())
        end = shape.Range.End
        sel.SetRange(end, end)
        if page_break and n < len(paths) - 1:
            sel.InsertBreak(7)                # wdPageBreak



# ------------------------------------------------------------------------ GUI
def build_report_tab(root, parent, shared):
    """Costruisce la schermata 2 (Relazione). Ritorna {'path': f, 'load': f}."""
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    from PIL import Image, ImageTk

    st = {"pdf": None, "doc": None, "lay": None, "blocks": None, "X": None,
          "imgs": [], "k": 1.0, "thumbs": {}, "photos": []}
    v_w = tk.StringVar(value=f"{OUT_W_MM:g}")
    v_h = tk.StringVar(value=f"{OUT_H_MM:g}")
    v_gap = tk.StringVar(value=f"{GAP_MAX_MM:g}")
    v_dpi = tk.StringVar(value=str(REN_DPI))
    v_brk = tk.BooleanVar(value=True)
    v_zoom = tk.IntVar(value=240)

    bar = ttk.Frame(parent, padding=6)
    bar.pack(fill="x")
    opt = ttk.Frame(parent, padding=(6, 0, 6, 6))
    opt.pack(fill="x")
    lbl = ttk.Label(bar, text="Nessun PDF aperto")
    view = ttk.Frame(parent)
    view.pack(fill="both", expand=True)
    cv = tk.Canvas(view, bg="#e8e8e8", highlightthickness=0)
    sb = ttk.Scrollbar(view, orient="vertical", command=cv.yview)
    cv.configure(yscrollcommand=sb.set)
    sb.pack(side="right", fill="y")
    cv.pack(side="left", fill="both", expand=True)
    status = ttk.Label(parent, anchor="w", padding=4,
                       text="Apri un PDF di IDEA StatiCa: le pagine partono da \"Membrature\".")
    status.pack(fill="x")

    def fnum(var, default, lo, hi):
        try:
            v = float(str(var.get()).replace(",", "."))
        except ValueError:
            return default
        return min(max(v, lo), hi)

    def busy(msg):
        status.config(text=msg)
        root.config(cursor="watch")
        root.update_idletasks()

    # ------------------------------------------------------------- anteprima
    def draw_preview(*_):
        cv.delete("all")
        st["photos"] = []
        imgs = st["imgs"]
        if not imgs:
            return
        tw = int(v_zoom.get())
        th = int(round(tw * imgs[0].size[1] / imgs[0].size[0]))
        cw = max(cv.winfo_width(), tw + 20)
        cols = max(1, (cw - 10) // (tw + 16))
        for i, im in enumerate(imgs):
            key = (i, tw)
            if key not in st["thumbs"]:
                st["thumbs"][key] = im.resize((tw, th), Image.LANCZOS)
            ph = ImageTk.PhotoImage(st["thumbs"][key])
            st["photos"].append(ph)
            x = 10 + (i % cols) * (tw + 16)
            y = 10 + (i // cols) * (th + 34)
            cv.create_rectangle(x - 1, y - 1, x + tw, y + th, outline="#888", fill="white")
            cid = cv.create_image(x, y, image=ph, anchor="nw")
            cv.create_text(x + tw // 2, y + th + 12, text=f"Pagina {i + 1} / {len(imgs)}")
            cv.tag_bind(cid, "<Double-Button-1>", lambda e, n=i: big(n))
        rows = (len(imgs) + cols - 1) // cols
        cv.configure(scrollregion=(0, 0, cw, 20 + rows * (th + 34)))

    def big(n):
        im = st["imgs"][n]
        win = tk.Toplevel(root)
        win.title(f"Pagina {n + 1} / {len(st['imgs'])}")
        z = min((root.winfo_screenheight() - 120) / im.size[1], 1.0)
        disp = im.resize((int(im.size[0] * z), int(im.size[1] * z)), Image.LANCZOS)
        ph = ImageTk.PhotoImage(disp)
        lb = ttk.Label(win, image=ph)
        lb.image = ph
        lb.pack()

    cv.bind("<Configure>", draw_preview)
    def wheel(e):
        cv.yview_scroll(-1 * (e.delta // 120), "units")
    cv.bind("<Enter>", lambda e: cv.bind_all("<MouseWheel>", wheel))
    cv.bind("<Leave>", lambda e: cv.unbind_all("<MouseWheel>"))

    # ------------------------------------------------------------ elaborazione
    def relayout(*_):
        if not st["blocks"]:
            return
        w = fnum(v_w, OUT_W_MM, 50, 400)
        h = fnum(v_h, OUT_H_MM, 50, 600)
        gap = fnum(v_gap, GAP_MAX_MM, 0, 50)
        dpi = int(fnum(v_dpi, REN_DPI, 72, 400))
        busy("Impaginazione e rendering in corso...")
        try:
            X0, X1 = st["X"]
            k, _H, pages = layout_report(st["blocks"], X0, X1, w, h, gap / 25.4 * 72)
            imgs = render_report(st["doc"], st["lay"], st["blocks"], pages,
                                 X0, X1, k, h, gap / 25.4 * 72, dpi)
        finally:
            root.config(cursor="")
        st["imgs"], st["k"], st["thumbs"], st["dpi"] = imgs, k, {}, dpi
        wmm = imgs[0].size[0] / dpi * 25.4
        hmm = imgs[0].size[1] / dpi * 25.4
        status.config(text=f"{len(imgs)} pagine da {wmm:.0f} × {hmm:.0f} mm  -  "
                           f"scala {k * 100:.0f}%  -  doppio click su una pagina per ingrandirla")
        draw_preview()

    def load(path=None):
        path = path or filedialog.askopenfilename(
            title="Seleziona il PDF di IDEA StatiCa", filetypes=[("PDF", "*.pdf")])
        if not path:
            return
        busy("Analisi del PDF in corso...")
        try:
            doc = PdfDoc(path)
            lay = find_report_layout(doc)
            blocks, X0, X1 = build_blocks(doc, lay, analyze_pages(doc, lay))
        except Exception:
            root.config(cursor="")
            raise
        st.update(pdf=path, doc=doc, lay=lay, blocks=blocks, X=(X0, X1))
        shared["pdf"] = path
        lbl.config(text=os.path.basename(path))
        relayout()

    def need_pages():
        if not st["imgs"]:
            messagebox.showinfo("Nessuna pagina", "Apri prima un PDF.")
            return False
        return True

    def save_tmp():
        dpi = st["dpi"]
        paths, sizes = [], []
        for im in st["imgs"]:
            fd, tmp = tempfile.mkstemp(suffix=".png")
            os.close(fd)
            im.save(tmp, dpi=(dpi, dpi))
            paths.append(tmp)
            sizes.append((im.size[0] / dpi * 25.4, im.size[1] / dpi * 25.4))
        return paths, sizes

    def to_word():
        if not need_pages():
            return
        paths, sizes = save_tmp()
        try:
            insert_pages_in_word(paths, sizes, v_brk.get())
            status.config(text=f"{len(paths)} pagine inserite in Word alla posizione del cursore.")
        finally:
            for p in paths:
                try:
                    os.remove(p)
                except OSError:
                    pass

    def export_png():
        if not need_pages():
            return
        d = filedialog.askdirectory(title="Cartella dove salvare le pagine PNG",
                                    initialdir=os.path.dirname(st["pdf"]))
        if not d:
            return
        base = os.path.splitext(os.path.basename(st["pdf"]))[0]
        dpi = st["dpi"]
        for i, im in enumerate(st["imgs"], 1):
            im.save(os.path.join(d, f"{base}_pag{i:02d}.png"), dpi=(dpi, dpi))
        status.config(text=f"Salvate {len(st['imgs'])} pagine in {d}")

    ttk.Button(bar, text="Apri PDF", command=load).pack(side="left")
    lbl.pack(side="left", padx=10)
    ttk.Button(bar, text="Inserisci in Word (cursore)", command=to_word).pack(side="right", padx=3)
    ttk.Button(bar, text="Salva pagine PNG", command=export_png).pack(side="right", padx=3)

    for txt, var, wd in (("Larghezza max [mm]", v_w, 5), ("Altezza max [mm]", v_h, 5),
                         ("Spazio tra blocchi [mm]", v_gap, 5), ("Risoluzione [dpi]", v_dpi, 5)):
        ttk.Label(opt, text=txt).pack(side="left", padx=(0, 3))
        e = ttk.Entry(opt, textvariable=var, width=wd)
        e.pack(side="left", padx=(0, 12))
        e.bind("<Return>", relayout)
    ttk.Checkbutton(opt, text="Interruzione di pagina tra le immagini",
                    variable=v_brk).pack(side="left", padx=(0, 12))
    ttk.Button(opt, text="Ricalcola", command=relayout).pack(side="left")
    ttk.Label(opt, text="Anteprima").pack(side="left", padx=(20, 3))
    tk.Scale(opt, variable=v_zoom, from_=120, to=520, orient="horizontal", showvalue=False,
             length=140, command=lambda _v: draw_preview()).pack(side="left")

    return {"path": lambda: st["pdf"], "load": load}


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
    root.title("idea2word - nodo IDEA StatiCa: viste e relazione per Word")
    root.geometry("1420x800")

    def report(exc, tb, tbk):
        show_error("Errore", "".join(traceback.format_exception(exc, tb, tbk)))
    root.report_callback_exception = report

    shared = {"pdf": None}      # PDF corrente condiviso tra le due schermate
    nb = ttk.Notebook(root)
    nb.pack(fill="both", expand=True)
    tab1 = ttk.Frame(nb)
    tab2 = ttk.Frame(nb)
    nb.add(tab1, text="  1 · Viste ortogonali  ")
    nb.add(tab2, text="  2 · Relazione (pagine)  ")

    st = {"pdf": None, "imgs": None, "photo": None, "job": None, "comp": None}
    scale_v = {k: tk.DoubleVar(value=1.0) for k in KEYS}
    dx_v = {k: tk.IntVar(value=0) for k in KEYS}
    dy_v = {k: tk.IntVar(value=0) for k in KEYS}

    # barra pulsanti
    bar = ttk.Frame(tab1, padding=6)
    bar.pack(fill="x")
    lbl = ttk.Label(bar, text="Nessun PDF aperto")
    main = ttk.Frame(tab1)
    main.pack(fill="both", expand=True)
    canvas = tk.Canvas(main, bg="#e8e8e8", highlightthickness=0)
    canvas.pack(side="left", fill="both", expand=True)
    panel = ttk.Frame(main, padding=6)
    panel.pack(side="right", fill="y")
    status = ttk.Label(tab1, anchor="w", padding=4,
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
        shared["pdf"] = path
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

    api1 = {"path": lambda: st["pdf"], "load": open_pdf}
    api2 = build_report_tab(root, tab2, shared)

    def on_tab(_e=None):
        # passando da una schermata all'altra carica da solo il PDF gia' aperto
        api = api1 if nb.index(nb.select()) == 0 else api2
        p = shared["pdf"]
        if p and api["path"]() != p and os.path.exists(p):
            api["load"](p)
    nb.bind("<<NotebookTabChanged>>", on_tab)

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
