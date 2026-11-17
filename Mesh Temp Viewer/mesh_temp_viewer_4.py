# -*- coding: utf-8 -*-
# =====================================================================
#  Mesh Temp Viewer 4   (Converter)   (Ansys Mechanical, IronPython 2.7)
#
#  - .mesh ve .temp dosyalarini okur (2D: X = radial, Y = eksenel)
#  - Modeli mesh + sicaklik renk skalasi (contour) ile gosterir
#  - Node secimi: tikla = sec/birak | Shift+surukle = kutu secim
#  - Saniye (TIME) secimi: combobox
#  - Secilen node'lari TEXT olarak verir:  TIME NODE X Y TEMP
#    (istenirse tum time'lar icin, istenirse mesh dosyasindaki 3. bolum degeri de ekli)
#
#  Mouse : tekerlek = zoom (imlec uzerinde) | orta tus / sol surukle = pan
#          tikla = node sec/birak | Shift+surukle (veya 'Box select mode') = kutu secim
#          Sag tus basili tut + surukle = kutu secim (her zaman)
#          Undo butonu / Ctrl+Z = son secimi geri al
#  v4: buyuk (GB) .temp icin tembel okuma (sadece blok konumlari taranir, sicakliklar gerektikce okunur),
#      onbellekli (cache) cizim, hizli node secimi, ilerleme penceresi (iptal edilebilir)
#          Birim: ekran (contour) ve export icin C veya K (varsayilan C); dosyadaki degerler C kabul edilir
# =====================================================================
import clr
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
import System
import re
import math
import time

from System.Windows.Forms import (Form, Button, Label, MessageBox, MessageBoxButtons, MessageBoxIcon,
                                  FormStartPosition, DialogResult, Panel, Cursors, Control, Keys,
                                  MouseButtons, Clipboard, ComboBox, ComboBoxStyle, CheckBox, ListBox,
                                  OpenFileDialog, SaveFileDialog, AnchorStyles, Screen, IWin32Window,
                                  CheckedListBox, TextBox, FormBorderStyle, Application, Timer, ProgressBar)
from System.Drawing import (Size, Point, Font, FontStyle, Color, PointF, Pen, SolidBrush, Bitmap, Graphics,
                            RectangleF)
from System.Drawing.Drawing2D import SmoothingMode, DashStyle

TITLE = u"Mesh Temp Viewer 4"
SW = 360            # sag panel genisligi
SUB_TRI = 6         # ucgen eleman alt bolme sayisi (renk yumusakligi)
SUB_QUAD = 4        # dortgen eleman alt bolme sayisi


# ---------------------------------------------------------------
# Dosya okuma
# ---------------------------------------------------------------
def tofloat(s):
    """Fortran tarzi sayilari da okur (ornek: 0.123-100)."""
    try:
        return float(s)
    except ValueError:
        m = re.match(r"^([-+]?\d*\.?\d+)([-+]\d+)$", s)
        if m:
            return float(m.group(1) + "E" + m.group(2))
        raise


def read_lines(path):
    return [unicode(l) for l in System.IO.File.ReadAllLines(path)]


class Cancelled(Exception):
    pass


ENC = System.Text.Encoding.GetEncoding(28591)       # latin-1: 1 karakter = 1 byte (dosya konumu = karakter konumu)
HDR_RE = re.compile(r"^(CASE|NODE|TIME|RPM)\s*=\s*(\S+)")
BLOCK_KEY = u"TEMPERATURES"


def parse_header_text(text):
    """Bir blogun ilk satirlari -> {'case','time','rpm','data_off'} (data_off = ilk veri satirinin indeksi)."""
    h = {"case": None, "time": None, "rpm": None, "data_off": None}
    for k, ln in enumerate(text.split("\n")[:12]):
        u = ln.strip().upper()
        if u.startswith("NODE TEMPERATURE"):
            h["data_off"] = k + 1
            break
        m = HDR_RE.match(u)
        if m:
            key, v = m.group(1), m.group(2)
            try:
                if key == "TIME":
                    h["time"] = tofloat(v)
                elif key == "CASE":
                    h["case"] = int(float(v))
                elif key == "RPM":
                    h["rpm"] = tofloat(v)
            except ValueError:
                pass
    return h


def parse_lines(lines, d0):
    """Veri satirlari -> (node id listesi, sicaklik listesi) (dosyadaki sirayla)."""
    ids, vals = [], []
    for ln in lines[d0:]:
        tk = ln.split()
        if len(tk) >= 2:
            try:
                a = int(tk[0])
                b = float(tk[1])
            except ValueError:
                try:
                    a = int(tk[0])
                    b = tofloat(tk[1])
                except ValueError:
                    continue
            ids.append(a)
            vals.append(b)
    return ids, vals


class TempFile(object):
    """Cok buyuk .temp dosyalari icin TEMBEL okuyucu.
    Acilista dosya bir kez taranir (sadece 'TEMPERATURES' satir baslarinin konumu bulunur); sicakliklar bir time
    secilince (veya export'ta) sadece o blok okunur. Node sirasi bloklarda ayni kabul edilir; export'ta secili
    node'larin satir sirasi ilk bloktan ogrenilir ve diger bloklarda dogrudan o satira gidilir (dogrulanir)."""

    def __init__(self, path, progress=None):
        self.path = path
        self.size = System.IO.FileInfo(path).Length
        self.blocks = []
        self.times = []
        self.block_of = {}
        self.pos_of = None
        self._cache = {}
        self._order = []
        self._scan(progress)

    def _open(self):
        return System.IO.FileStream(self.path, System.IO.FileMode.Open, System.IO.FileAccess.Read,
                                    System.IO.FileShare.ReadWrite, 1048576)

    def _scan(self, progress):
        CH = 16 * 1024 * 1024
        buf = System.Array.CreateInstance(System.Byte, CH)
        hits = set()
        tail = u""
        pos = 0
        fs = self._open()
        try:
            while True:
                n = fs.Read(buf, 0, CH)
                if n <= 0:
                    break
                text = tail + ENC.GetString(buf, 0, n)
                tbase = pos - len(tail)
                i = text.find(BLOCK_KEY)
                while i >= 0:
                    if i > 0:
                        prev = text[i - 1]
                    elif tbase == 0:
                        prev = u"\n"
                    else:
                        prev = None
                    if prev in (u"\n", u"\r"):
                        hits.add(tbase + i)
                    i = text.find(BLOCK_KEY, i + 1)
                pos += n
                tail = text[-16:]
                if progress is not None and not progress(float(pos) / max(self.size, 1),
                                                         u"Scanning temp file ... %d%%   (%d time blocks found)"
                                                         % (int(100.0 * pos / max(self.size, 1)), len(hits))):
                    raise Cancelled()
            offs = sorted(hits)
            for k, o in enumerate(offs):
                fs.Seek(o, System.IO.SeekOrigin.Begin)
                hb = System.Array.CreateInstance(System.Byte, 1200)
                got = fs.Read(hb, 0, 1200)
                h = parse_header_text(ENC.GetString(hb, 0, got))
                h["off"] = o
                h["end"] = offs[k + 1] if k + 1 < len(offs) else self.size
                self.blocks.append(h)
        finally:
            fs.Close()
        for k, b in enumerate(self.blocks):
            if b["time"] is not None:
                self.block_of[b["time"]] = k
        self.times = sorted(self.block_of.keys())

    def info(self, t):
        b = self.blocks[self.block_of[t]]
        return b["case"], b["rpm"]

    def read_block(self, k):
        b = self.blocks[k]
        n = int(b["end"] - b["off"])
        fs = self._open()
        try:
            fs.Seek(b["off"], System.IO.SeekOrigin.Begin)
            buf = System.Array.CreateInstance(System.Byte, n)
            got = 0
            while got < n:
                r = fs.Read(buf, got, n - got)
                if r <= 0:
                    break
                got += r
            return ENC.GetString(buf, 0, got)
        finally:
            fs.Close()

    def _d0(self, k):
        return self.blocks[k]["data_off"] or 6

    def block_full(self, t):
        """Bir time'in tum node sicakliklari {node: T} (son 3 time onbellekte)."""
        if t in self._cache:
            return self._cache[t]
        k = self.block_of[t]
        ids, vals = parse_lines(self.read_block(k).split("\n"), self._d0(k))
        if self.pos_of is None:
            self.pos_of = dict((n, i) for i, n in enumerate(ids))
        d = dict(zip(ids, vals))
        self._cache[t] = d
        self._order.append(t)
        if len(self._order) > 3:
            self._cache.pop(self._order.pop(0), None)
        return d

    def block_nodes(self, t, nids):
        """Sadece verilen node'larin sicakliklari (liste, nids sirasinda; yoksa None). Hizli yol: node'un satir
        sirasi bellidir, tum satirlar ayristirilmaz."""
        if t in self._cache:
            d = self._cache[t]
            return [d.get(n) for n in nids]
        if self.pos_of is None:
            self.block_full(t)
        k = self.block_of[t]
        d0 = self._d0(k)
        po = self.pos_of
        idxs = [po.get(n) for n in nids]
        mx = max([i for i in idxs if i is not None] or [0])
        lines = self.read_block(k).split("\n", d0 + mx + 2)
        out = []
        for n, i in zip(nids, idxs):
            if i is None:
                out.append(None)
                continue
            try:
                tk = lines[d0 + i].split()
                if int(tk[0]) != n:
                    raise ValueError("order")
                try:
                    out.append(float(tk[1]))
                except ValueError:
                    out.append(tofloat(tk[1]))
            except (ValueError, IndexError):
                d = self.block_full(t)                  # sira farkli: bu blok icin tam ayristirma
                return [d.get(m_) for m_ in nids]
        return out

    def global_range(self, progress):
        """Tum time'lardaki min / max sicaklik (butun dosya okunur)."""
        lo = hi = None
        nt = len(self.times)
        for i, t in enumerate(self.times):
            if not progress(float(i) / nt, u"Scanning all times for min / max ...  %d / %d" % (i + 1, nt)):
                raise Cancelled()
            k = self.block_of[t]
            ids, vals = parse_lines(self.read_block(k).split("\n"), self._d0(k))
            if vals:
                a, b = min(vals), max(vals)
                lo = a if lo is None else min(lo, a)
                hi = b if hi is None else max(hi, b)
        return (lo, hi) if lo is not None else (0.0, 1.0)


def parse_mesh(path):
    """-> (nodes {id: (X, Y)}, elems [(id, [node ids])], meshval {node: value})
    Kayitlar:  1 = node (3 satir)   2 = eleman (3 satir)   10 = node degeri (2 satir)"""
    L = read_lines(path)
    n = len(L)

    def toks(i):
        return L[i].split() if 0 <= i < n else []
    start = 4
    for i in range(0, n - 2):                     # ilk node kaydini bul (ust bilgi satirlarini atla)
        t0, t1, t2 = toks(i), toks(i + 1), toks(i + 2)
        if len(t0) >= 2 and t0[0] == "1" and len(t1) == 3 and t2 and t2[0].upper().startswith("1G"):
            try:
                [tofloat(x) for x in t1]
                start = i
                break
            except ValueError:
                pass
    nodes, elems, mv = {}, [], {}
    i = start
    while i < n:
        t = toks(i)
        if not t:
            i += 1
            continue
        k = t[0]
        try:
            if k == "1" and len(t) >= 2:
                c = toks(i + 1)
                nodes[int(t[1])] = (tofloat(c[0]), tofloat(c[1]))
                i += 3
            elif k == "2" and len(t) >= 2:
                cnt = int(toks(i + 1)[0])
                ids, j = [], i + 2
                while len(ids) < cnt and j < n:
                    ids.extend(int(x) for x in toks(j))
                    j += 1
                elems.append((int(t[1]), ids[:cnt]))
                i = j
            elif k == "10" and len(t) >= 2:
                mv[int(t[1])] = tofloat(toks(i + 1)[0])
                i += 2
            else:
                i += 1
        except (ValueError, IndexError):
            i += 1
    return nodes, elems, mv


# ---------------------------------------------------------------
# Model (elemanlar + renk icin alt poligonlar)
# ---------------------------------------------------------------
def build_elem(eid, ids, xy):
    cnt = len(ids)
    if cnt in (3, 6):
        nc, kind = 3, "tri"
    elif cnt in (4, 8):
        nc, kind = 4, "quad"
    else:
        return None
    if any(i not in xy for i in ids):
        return None
    corners = ids[:nc]
    mids = None
    if cnt > nc:                                  # orta node'lar: en yakin kenar ortasina gore sirala
        rest = list(ids[nc:])
        mids = []
        for e in range(nc):
            a, b = xy[corners[e]], xy[corners[(e + 1) % nc]]
            mx, my = (a[0] + b[0]) / 2.0, (a[1] + b[1]) / 2.0
            best = min(rest, key=lambda r: math.hypot(xy[r][0] - mx, xy[r][1] - my))
            rest.remove(best)
            mids.append(best)
    return {"id": eid, "kind": kind, "corners": corners, "mids": mids,
            "nodes": list(corners) + (list(mids) if mids else [])}


def tessellate(el, xy, nt=SUB_TRI, nq=SUB_QUAD):
    """Elemani kucuk poligonlara boler; her birinin merkezindeki sekil fonksiyonu agirliklari saklanir
    (renk = agirliklar x node sicakliklari)."""
    kind, md = el["kind"], el["mids"]
    P = [xy[i] for i in el["nodes"]]
    CS = [(-1, -1), (1, -1), (1, 1), (-1, 1)]
    MS = [(0, -1), (1, 0), (0, 1), (-1, 0)]

    def shp(a, b):
        if kind == "tri":
            L0, L1, L2 = 1.0 - a - b, a, b
            if md:
                return [L0 * (2 * L0 - 1), L1 * (2 * L1 - 1), L2 * (2 * L2 - 1),
                        4 * L0 * L1, 4 * L1 * L2, 4 * L2 * L0]
            return [L0, L1, L2]
        if md:
            w = [0.25 * (1 + a * cx) * (1 + b * cy) * (a * cx + b * cy - 1) for cx, cy in CS]
            for mx, my in MS:
                w.append(0.5 * (1 - a * a) * (1 + b * my) if mx == 0 else 0.5 * (1 + a * mx) * (1 - b * b))
            return w
        return [0.25 * (1 + a * cx) * (1 + b * cy) for cx, cy in CS]

    def pt(a, b):
        w = shp(a, b)
        X = sum(wi * p[0] for wi, p in zip(w, P))
        Y = sum(wi * p[1] for wi, p in zip(w, P))
        return (X, -Y)                            # cizim koordinati: yatay = X (radial), dikey = Y (eksenel, yukari)

    subs = []

    def add(verts):
        ca = sum(v[0] for v in verts) / float(len(verts))
        cb = sum(v[1] for v in verts) / float(len(verts))
        subs.append({"pts": [pt(a, b) for (a, b) in verts], "w": shp(ca, cb), "ci": 256})

    if kind == "tri":
        n = nt
        f = 1.0 / n
        for i in range(n):
            for j in range(n - i):
                add([(i * f, j * f), ((i + 1) * f, j * f), (i * f, (j + 1) * f)])
                if i + j < n - 1:
                    add([((i + 1) * f, j * f), ((i + 1) * f, (j + 1) * f), (i * f, (j + 1) * f)])
    else:
        n = nq
        f = 2.0 / n
        for i in range(n):
            for j in range(n):
                x0, x1, y0, y1 = -1 + i * f, -1 + (i + 1) * f, -1 + j * f, -1 + (j + 1) * f
                add([(x0, y0), (x1, y0), (x1, y1), (x0, y1)])
    el["subs"] = subs
    el["cpoly"] = [(xy[c][0], -xy[c][1]) for c in el["corners"]]
    el["ci"] = 256
    cn, ms = el["corners"], md
    ring = []
    for e in range(len(cn)):
        ring.append(xy[cn[e]])
        if ms:
            ring.append(xy[ms[e]])
    el["outline"] = [(p[0], -p[1]) for p in ring]
    xs = [p[0] for sp in subs for p in sp["pts"]]
    ys = [p[1] for sp in subs for p in sp["pts"]]
    el["bb"] = (min(xs), max(xs), min(ys), max(ys))


def build_model(nodes, raw_elems):
    els, skipped = [], 0
    k = max(1, min(SUB_TRI, int(math.sqrt(40000.0 / max(1, len(raw_elems))))))     # buyuk modelde alt bolme azalir
    for eid, ids in raw_elems:
        e = build_elem(eid, ids, nodes)
        if e is None:
            skipped += 1
            continue
        tessellate(e, nodes, k, min(SUB_QUAD, k))
        els.append(e)
    used = set()
    for e in els:
        used.update(e["nodes"])
    if not els:
        raise ValueError(u"Mesh dosyasindan hic eleman okunamadi.")
    xs = [e["bb"][0] for e in els] + [e["bb"][1] for e in els]
    ys = [e["bb"][2] for e in els] + [e["bb"][3] for e in els]
    return {"nodes": nodes, "elems": els, "used": used, "skipped": skipped,
            "box": (min(xs), max(xs), min(ys), max(ys))}


# ---------------------------------------------------------------
# Renk skalasi
# ---------------------------------------------------------------
def jet(t):
    t = max(0.0, min(1.0, t))
    r = min(1.0, max(0.0, 1.5 - abs(4 * t - 3)))
    g = min(1.0, max(0.0, 1.5 - abs(4 * t - 2)))
    b = min(1.0, max(0.0, 1.5 - abs(4 * t - 1)))
    return Color.FromArgb(int(r * 255), int(g * 255), int(b * 255))


def zoom_about(px, py, mx, my, W, H, r):
    ox, oy = mx - W / 2.0, my - H / 2.0
    return ox - (ox - px) * r, oy - (oy - py) * r


def clampf(v):
    return max(-1.0e6, min(1.0e6, v))


# ---------------------------------------------------------------
# Gorunum
# ---------------------------------------------------------------
class SceneView(Panel):
    """Model cizimi bir Bitmap'e (cache) yapilir; zoom / pan sirasinda bu bitmap olceklenip kaydirilarak hemen
    gosterilir, gercek (net) cizim hareket durunca (150 ms) yapilir. Secim isaretleri cache'in ustune cizilir,
    yani secim / kutu cizimi modeli yeniden cizmez."""
    def __init__(self, owner):
        Panel.__init__(self)
        self.owner = owner
        self.DoubleBuffered = True
        self.BackColor = Color.White
        self.model = None
        self.dn = {}                                  # node -> cizim koordinati
        self._grid = {}
        self._cell = 1.0
        self.avg_el = 1.0
        self.sel = set()
        self.s = 1.0
        self.px = self.py = self.cx = self.cy = 0.0
        self.mesh_on = True
        self.ids_on = False
        self.box_mode = False
        self.unit_txt = u"C"
        self.off = 0.0
        self._btn = None
        self.tmap = {}
        self.vmin, self.vmax = 0.0, 1.0
        self.has_t = False
        self._down = self._pan = self._rb = None
        self._drag = False
        self._fit_pending = False
        self._last = None
        self._cache = None
        self._cs = 1.0
        self._cofs = (0.0, 0.0)
        self._dirty = True
        self._stale = False
        self._timer = Timer()
        self._timer.Interval = 150
        self._timer.Tick += self._on_timer
        self.brushes = [SolidBrush(jet(i / 255.0)) for i in range(256)] + [SolidBrush(Color.LightGray)]
        self.pens = [Pen(b.Color, 1.0) for b in self.brushes]
        self.f_ui = Font("Segoe UI", 8.5)
        self.f_lab = Font("Segoe UI", 7.5)
        self.b_dark = SolidBrush(Color.FromArgb(40, 40, 40))
        self.b_green = SolidBrush(Color.FromArgb(0, 200, 110))
        self.p_glow = Pen(Color.FromArgb(0, 120, 60), 1.5)
        self.p_mesh = Pen(Color.FromArgb(70, 70, 70), 1.0)
        self.Paint += self._paint
        self.MouseDown += self._mdown
        self.MouseMove += self._mmove
        self.MouseUp += self._mup
        self.MouseWheel += self._wheel
        self.MouseEnter += lambda s, e: self.Focus()
        self.Resize += self._on_resize

    # ----- model / renkler -----
    def set_model(self, model):
        self.model = model
        self.dn = dict((nid, (xy[0], -xy[1])) for nid, xy in model["nodes"].items() if nid in model["used"])
        els = model["elems"]
        sample = els[:2000]
        self.avg_el = max(sum(e["bb"][1] - e["bb"][0] for e in sample) / float(len(sample)), 1e-12)
        self._cell = self.avg_el
        c = self._cell
        grid = {}
        for nid, (x, y) in self.dn.items():
            grid.setdefault((int(math.floor(x / c)), int(math.floor(y / c))), []).append(nid)
        self._grid = grid
        self.sel = set()
        self._fit_pending = True
        self._dirty = True
        self.Invalidate()

    def set_temps(self, tmap, vmin, vmax, off=0.0):
        self.tmap, self.vmin, self.vmax, self.off = tmap, vmin, vmax, off
        self.has_t = True
        self.recolor()
        self._dirty = True
        self.Invalidate()

    def mark_dirty(self):
        self._dirty = True
        self.Invalidate()

    def recolor(self):
        if self.model is None:
            return
        rng = self.vmax - self.vmin
        vmin = self.vmin
        get = self.tmap.get
        for el in self.model["elems"]:
            tv = [get(i) for i in el["nodes"]]
            subs = el["subs"]
            if None in tv:
                for sp in subs:
                    sp["ci"] = 256
                el["ci"] = 256
                continue
            tot = 0
            for sp in subs:
                T = 0.0
                for w, v in zip(sp["w"], tv):
                    T += w * v
                ci = 128 if rng <= 0 else int(255 * (T - vmin) / rng)
                ci = 0 if ci < 0 else (255 if ci > 255 else ci)
                sp["ci"] = ci
                tot += ci
            el["ci"] = tot // len(subs)

    # ----- gorunum degisimi -----
    def _view_changed(self):
        self._stale = True
        self._timer.Stop()
        self._timer.Start()
        self.Invalidate()

    def _on_timer(self, sender, e):
        self._timer.Stop()
        if self._stale:
            self._stale = False
            self._dirty = True
            self.Invalidate()

    def _on_resize(self, sender, e):
        W, H = self.Width, self.Height
        if (self._last is not None and not self._fit_pending and W > 50 and H > 50
                and self._last[0] > 50 and self._last[1] > 50):
            r = min(float(W) / self._last[0], float(H) / self._last[1])
            self.s *= r
            self.px *= r
            self.py *= r
        if W > 0 and H > 0:
            self._last = (W, H)
        self._view_changed()

    def _ofs(self):
        return (self.Width / 2.0 + self.px - self.cx * self.s,
                self.Height / 2.0 + self.py - self.cy * self.s)

    def fit_all(self):
        if self.model is None:
            return
        x0, x1, y0, y1 = self.model["box"]
        bw, bh = x1 - x0, y1 - y0
        big = max(bw, bh, 1e-12)
        bw, bh = max(bw, 0.25 * big), max(bh, 0.25 * big)
        self.cx, self.cy = (x0 + x1) / 2.0, (y0 + y1) / 2.0
        self.s = min((self.Width - 130.0) / (bw * 1.1), (self.Height - 60.0) / (bh * 1.1))
        self.px = self.py = 0.0
        self._view_changed()

    def zoom(self, f, mx, my):
        self.px, self.py = zoom_about(self.px, self.py, mx, my, self.Width, self.Height, f)
        self.s = max(1e-9, self.s * f)
        self._view_changed()

    # ----- cizim -----
    def _render_cache(self):
        W, H = self.Width, self.Height
        if W < 10 or H < 10:
            return
        bmp = Bitmap(W, H)
        g = Graphics.FromImage(bmp)
        try:
            g.Clear(Color.White)
            self._draw_model(g)
        finally:
            g.Dispose()
        if self._cache is not None:
            self._cache.Dispose()
        self._cache = bmp
        self._cs = self.s
        self._cofs = self._ofs()
        self._dirty = False

    def _paint(self, sender, e):
        g = e.Graphics
        if self._fit_pending and self.Width > 60 and self.Height > 60:
            self._fit_pending = False
            self.fit_all()
            self._stale = False
            self._timer.Stop()
        g.Clear(Color.White)
        if self.model is None:
            g.DrawString(u"Mesh / Temp dosyasi yuklenmedi.", self.f_ui, SolidBrush(Color.DimGray), 8.0, 8.0)
            return
        try:
            if self._cache is None or self._dirty:
                self._render_cache()
                self._stale = False
                self._timer.Stop()
            if self._cache is not None:
                if self._stale:
                    r = self.s / self._cs
                    ox, oy = self._ofs()
                    dx, dy = ox - self._cofs[0] * r, oy - self._cofs[1] * r
                    g.DrawImage(self._cache, RectangleF(dx, dy, self._cache.Width * r, self._cache.Height * r))
                else:
                    g.DrawImage(self._cache, 0, 0)
            self._draw_selection(g)
        except Exception as ex:
            g.DrawString(u"Draw error: %s" % str(ex), self.f_ui, SolidBrush(Color.Firebrick), 8.0, 24.0)
        g.SmoothingMode = SmoothingMode.AntiAlias
        g.DrawString(u"Selected nodes: %d    |  wheel: zoom  |  drag: pan  |  click: select node  |  right-drag / Shift+drag / Box mode: box select  |  Ctrl+Z: undo"
                     % len(self.sel), self.f_ui, SolidBrush(Color.DimGray), 8.0, 4.0)
        if self.has_t:
            self._draw_legend(g, self.f_ui)
        if self._rb is not None:
            x0, y0, x1, y1 = self._rb
            pen = Pen(Color.DarkOrange, 1.0)
            pen.DashStyle = DashStyle.Dash
            g.DrawRectangle(pen, min(x0, x1), min(y0, y1), abs(x1 - x0), abs(y1 - y0))

    def _draw_model(self, g):
        s = self.s
        ox, oy = self._ofs()
        W, H = float(self.Width), float(self.Height)
        g.SmoothingMode = SmoothingMode.HighSpeed
        br, pens = self.brushes, self.pens
        mesh_on = self.mesh_on
        occ = set()
        for el in self.model["elems"]:
            x0, x1, y0, y1 = el["bb"]
            sx0, sx1, sy0, sy1 = x0 * s + ox, x1 * s + ox, y0 * s + oy, y1 * s + oy
            if sx1 < -2 or sx0 > W + 2 or sy1 < -2 or sy0 > H + 2:
                continue
            px = sx1 - sx0
            if px < 1.5:                                   # cok kucuk eleman: tek piksel (ayni piksele bir kez)
                key = (int(sx0), int(sy0))
                if key in occ:
                    continue
                occ.add(key)
                g.FillRectangle(br[el["ci"]], sx0 - 0.5, sy0 - 0.5, 1.6, 1.6)
                continue
            if px < 6.0:                                   # kucuk eleman: tek renkli poligon
                pts = System.Array[PointF]([PointF(clampf(p[0] * s + ox), clampf(p[1] * s + oy)) for p in el["cpoly"]])
                g.FillPolygon(br[el["ci"]], pts)
            else:
                for sp in el["subs"]:
                    pts = System.Array[PointF]([PointF(clampf(p[0] * s + ox), clampf(p[1] * s + oy)) for p in sp["pts"]])
                    g.FillPolygon(br[sp["ci"]], pts)
                    g.DrawPolygon(pens[sp["ci"]], pts)
            if mesh_on and px >= 4.0:
                ring = System.Array[PointF]([PointF(clampf(p[0] * s + ox), clampf(p[1] * s + oy)) for p in el["outline"]])
                g.DrawPolygon(self.p_mesh, ring)
        g.SmoothingMode = SmoothingMode.AntiAlias
        px_el = self.avg_el * s
        if px_el >= 10.0:                                  # node noktalari (yeterince yakinlasinca)
            vis = []
            for nid, (x, y) in self.dn.items():
                sx, sy = x * s + ox, y * s + oy
                if -10 <= sx <= W + 10 and -10 <= sy <= H + 10:
                    vis.append((nid, sx, sy))
            if len(vis) <= 20000:
                for nid, sx, sy in vis:
                    g.FillRectangle(self.b_dark, sx - 1.5, sy - 1.5, 3.0, 3.0)
                if self.ids_on and px_el >= 30.0 and len(vis) <= 2000:
                    for nid, sx, sy in vis:
                        g.DrawString(str(nid), self.f_lab, self.b_dark, sx + 5.0, sy - 14.0)

    def _draw_selection(self, g):
        if not self.sel:
            return
        s = self.s
        ox, oy = self._ofs()
        W, H = float(self.Width), float(self.Height)
        vis = []
        dn = self.dn
        for nid in self.sel:
            p = dn.get(nid)
            if p is None:
                continue
            sx, sy = p[0] * s + ox, p[1] * s + oy
            if -10 <= sx <= W + 10 and -10 <= sy <= H + 10:
                vis.append((nid, sx, sy))
        g.SmoothingMode = SmoothingMode.AntiAlias
        if len(vis) > 3000:
            for nid, sx, sy in vis:
                g.FillRectangle(self.b_green, sx - 2.0, sy - 2.0, 4.0, 4.0)
            return
        for nid, sx, sy in vis:
            g.FillEllipse(self.b_green, sx - 5.5, sy - 5.5, 11.0, 11.0)
            g.DrawEllipse(self.p_glow, sx - 5.5, sy - 5.5, 11.0, 11.0)
        if len(vis) <= 300:
            for nid, sx, sy in vis:
                g.DrawString(str(nid), self.f_lab, self.b_dark, sx + 5.0, sy - 14.0)

    def _draw_legend(self, g, f):
        W, H = float(self.Width), float(self.Height)
        x, y0, y1, w = W - 95.0, 30.0, H - 30.0, 16.0
        n = 64
        h = (y1 - y0) / n
        for k in range(n):
            ci = int(255 * (1.0 - k / float(n - 1)))
            g.FillRectangle(self.brushes[ci], x, y0 + k * h, w, h + 1.0)
        g.DrawRectangle(Pen(Color.Gray, 1.0), x, y0, w, y1 - y0)
        g.DrawString(u"[%s]" % self.unit_txt, f, SolidBrush(Color.Black), x, y0 - 16.0)
        br = SolidBrush(Color.Black)
        for frac in (0.0, 0.25, 0.5, 0.75, 1.0):
            val = self.vmax - (self.vmax - self.vmin) * frac + self.off
            yy = y0 + (y1 - y0) * frac
            g.DrawLine(Pen(Color.Gray, 1.0), x + w, yy, x + w + 4.0, yy)
            g.DrawString(u"%.2f" % val, f, br, x + w + 5.0, yy - 8.0)

    # ----- mouse -----
    def _wheel(self, sender, e):
        self.zoom(1.2 if e.Delta > 0 else 1.0 / 1.2, e.X, e.Y)

    def _mdown(self, sender, e):
        if e.Button == MouseButtons.Left or e.Button == MouseButtons.Right:
            self._down = (e.X, e.Y, self.px, self.py)
            self._drag = False
            self._btn = e.Button
        elif e.Button == MouseButtons.Middle:
            self._pan = (e.X, e.Y, self.px, self.py)
            self.Cursor = Cursors.SizeAll

    def _mmove(self, sender, e):
        if self._pan is not None:
            x0, y0, p0x, p0y = self._pan
            self.px, self.py = p0x + (e.X - x0), p0y + (e.Y - y0)
            self._view_changed()
            return
        if self._down is None:
            return
        x0, y0, p0x, p0y = self._down
        if not self._drag and math.hypot(e.X - x0, e.Y - y0) > 4:
            self._drag = True
        if not self._drag:
            return
        if self._btn == MouseButtons.Right or self.box_mode or (Control.ModifierKeys & Keys.Shift) == Keys.Shift:
            self._rb = (x0, y0, e.X, e.Y)
            self.Invalidate()
        else:
            self.px, self.py = p0x + (e.X - x0), p0y + (e.Y - y0)
            self._view_changed()

    def _mup(self, sender, e):
        if e.Button == MouseButtons.Middle:
            self._pan = None
            self.Cursor = Cursors.Cross if self.box_mode else Cursors.Default
            return
        if self._down is None:
            return
        x0, y0 = self._down[0], self._down[1]
        if self._drag and self._rb is not None:
            xa, xb = sorted((x0, e.X))
            ya, yb = sorted((y0, e.Y))
            ox, oy = self._ofs()
            s = self.s
            hit = set()
            for nid, (x, y) in self.dn.items():
                sx, sy = x * s + ox, y * s + oy
                if xa <= sx <= xb and ya <= sy <= yb:
                    hit.add(nid)
            if hit - self.sel:
                self.owner.push_undo()
                self.sel |= hit
                self.owner.selection_changed()
        elif not self._drag and self._btn == MouseButtons.Left:
            self._pick(x0, y0)
        self._btn = None
        self._down = None
        self._drag = False
        self._rb = None
        self.Invalidate()

    def _pick(self, mx, my):
        ox, oy = self._ofs()
        s = self.s
        X, Y = (mx - ox) / s, (my - oy) / s
        cell = self._cell
        k = int(math.ceil((12.0 / s) / cell))
        if (2 * k + 1) ** 2 > 600:
            cands = list(self.dn.keys())
        else:
            ci, cj = int(math.floor(X / cell)), int(math.floor(Y / cell))
            cands = []
            for a in range(ci - k, ci + k + 1):
                for b in range(cj - k, cj + k + 1):
                    lst = self._grid.get((a, b))
                    if lst:
                        cands.extend(lst)
        best, bn = 12.0, None
        for nid in cands:
            x, y = self.dn[nid]
            d = math.hypot(x * s + ox - mx, y * s + oy - my)
            if d < best:
                best, bn = d, nid
        if bn is not None:
            self.owner.push_undo()
            if bn in self.sel:
                self.sel.discard(bn)
            else:
                self.sel.add(bn)
            self.owner.selection_changed()


# ---------------------------------------------------------------
# Pencere
# ---------------------------------------------------------------
class _HwndOwner(IWin32Window):
    def __init__(self, h):
        self._h = h

    def get_Handle(self):
        return self._h
    Handle = property(get_Handle)


def mechanical_owner():
    try:
        h = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle
        if h != System.IntPtr.Zero:
            return _HwndOwner(h)
    except Exception:
        pass
    return None


def pick_open(title, flt, current=u""):
    d = OpenFileDialog()
    d.Title = title
    d.Filter = flt
    try:
        if current and System.IO.File.Exists(current):
            d.InitialDirectory = System.IO.Path.GetDirectoryName(current)
    except Exception:
        pass
    if d.ShowDialog() == DialogResult.OK:
        return d.FileName
    return None


class ProgressForm(Form):
    """Uzun islemler icin ilerleme penceresi (Cancel ile iptal edilebilir). update() False donerse iptal edilmistir."""
    def __init__(self, title):
        self.Text = title
        self.ClientSize = Size(520, 112)
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.ControlBox = False
        self.TopMost = True
        self.cancelled = False
        self._t = 0.0
        self.lbl = Label()
        self.lbl.SetBounds(12, 12, 496, 36)
        self.Controls.Add(self.lbl)
        self.bar = ProgressBar()
        self.bar.Minimum = 0
        self.bar.Maximum = 100
        self.bar.SetBounds(12, 54, 496, 18)
        self.Controls.Add(self.bar)
        b = Button()
        b.Text = u"Cancel"
        b.SetBounds(420, 78, 88, 26)
        b.Click += self._cancel
        self.Controls.Add(b)

    def _cancel(self, sender, e):
        self.cancelled = True

    def update(self, frac, text=None):
        now = time.time()
        if now - self._t > 0.08 or frac >= 1.0:
            self._t = now
            if text is not None:
                self.lbl.Text = text
            self.bar.Value = max(0, min(100, int(frac * 100)))
            Application.DoEvents()
        return not self.cancelled


class StartForm(Form):
    """Ilk ekran: .mesh ve .temp dosyalarini browse ile sec."""
    def __init__(self, mesh_path=u"", temp_path=u""):
        self.Text = TITLE + u" - Input files"
        self.ClientSize = Size(580, 160)
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.mesh_path, self.temp_path = mesh_path, temp_path
        self.tbMesh = self._row(u".mesh file:", 20, mesh_path,
                                u"Mesh files (*.mesh)|*.mesh|All files (*.*)|*.*")
        self.tbTemp = self._row(u".temp file:", 64, temp_path,
                                u"Temp files (*.temp)|*.temp|All files (*.*)|*.*")
        ok = Button()
        ok.Text = u"Load"
        ok.SetBounds(390, 112, 85, 30)
        ok.Click += lambda s, e: self._ok()
        self.Controls.Add(ok)
        cn = Button()
        cn.Text = u"Cancel"
        cn.SetBounds(485, 112, 85, 30)
        cn.DialogResult = DialogResult.Cancel
        self.Controls.Add(cn)
        self.AcceptButton = ok
        self.CancelButton = cn

    def _row(self, text, y, value, flt):
        l = Label()
        l.Text = text
        l.SetBounds(10, y + 4, 75, 20)
        self.Controls.Add(l)
        tb = TextBox()
        tb.Text = value
        tb.SetBounds(90, y, 390, 24)
        self.Controls.Add(tb)
        b = Button()
        b.Text = u"Browse..."
        b.SetBounds(490, y - 1, 80, 26)

        def click(s, e):
            f = pick_open(text, flt, tb.Text.strip())
            if f:
                tb.Text = f
        b.Click += click
        self.Controls.Add(b)
        return tb

    def _ok(self):
        m, t = self.tbMesh.Text.strip().strip('"'), self.tbTemp.Text.strip().strip('"')
        for path, nm in ((m, u".mesh"), (t, u".temp")):
            if not path or not System.IO.File.Exists(path):
                MessageBox.Show(u"%s dosyasi secilmedi veya bulunamadi." % nm, TITLE,
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
                return
        self.mesh_path, self.temp_path = m, t
        self.DialogResult = DialogResult.OK


class ViewerForm(Form):
    def __init__(self):
        self.Text = TITLE
        wa = Screen.PrimaryScreen.WorkingArea
        self.ClientSize = Size(int(0.8 * min(1500, wa.Width - 30)), int(0.85 * min(900, wa.Height - 70)))
        self.MinimumSize = Size(840, 660)
        self.StartPosition = FormStartPosition.CenterScreen
        self.nodes, self.mv = {}, {}
        self.tf = None
        self.times = []
        self.grange = None
        self.cur_map = {}
        self.cur_time = None
        self._undo = []
        self.mesh_path = self.temp_path = u""
        self.KeyPreview = True
        self.KeyDown += self._on_key
        self.view = SceneView(self)
        self.side = Panel()
        self.Controls.Add(self.view)
        self.Controls.Add(self.side)
        self.Resize += lambda s, e: self._layout()
        self._layout()
        self._build_side()

    def _layout(self):
        W, H = self.ClientSize.Width, self.ClientSize.Height
        self.side.SetBounds(W - SW, 0, SW, H)
        self.view.SetBounds(0, 0, max(100, W - SW), H)

    # ----- sag panel -----
    def _lbl(self, text, x, y, w, h, anchor=None):
        l = Label()
        l.Text = text
        l.SetBounds(x, y, w, h)
        if anchor is not None:
            l.Anchor = anchor
        self.side.Controls.Add(l)
        return l

    def _btn(self, text, x, y, w, h, handler, anchor=None):
        b = Button()
        b.Text = text
        b.SetBounds(x, y, w, h)
        b.Click += handler
        if anchor is not None:
            b.Anchor = anchor
        self.side.Controls.Add(b)
        return b

    def _chk(self, text, x, y, w, checked, handler, anchor=None):
        c = CheckBox()
        c.Text = text
        c.SetBounds(x, y, w, 22)
        c.Checked = checked
        c.CheckedChanged += handler
        if anchor is not None:
            c.Anchor = anchor
        self.side.Controls.Add(c)
        return c

    def _build_side(self):
        H0 = self.side.Height
        TL = AnchorStyles.Top | AnchorStyles.Left
        BL = AnchorStyles.Bottom | AnchorStyles.Left
        self.lblMesh = self._lbl(u"Mesh: -", 8, 8, SW - 16, 18, TL)
        self.lblTemp = self._lbl(u"Temp: -", 8, 26, SW - 16, 18, TL)
        self._btn(u"Open files...", 8, 48, 120, 26, lambda s, e: self.ask_files(), TL)
        self._btn(u"Fit view", 136, 48, 90, 26, lambda s, e: self.view.fit_all(), TL)
        self._lbl(u"Time [s]:", 8, 86, 60, 20, TL)
        self.cbTime = ComboBox()
        self.cbTime.DropDownStyle = ComboBoxStyle.DropDownList
        self.cbTime.SetBounds(70, 82, 150, 24)
        self.cbTime.SelectedIndexChanged += lambda s, e: self.time_changed()
        self.side.Controls.Add(self.cbTime)
        self.lblTinfo = self._lbl(u"", 226, 86, SW - 234, 20, TL)
        self.chkMesh = self._chk(u"Mesh lines", 8, 114, 106, True, lambda s, e: self._opt(), TL)
        self.chkIds = self._chk(u"Node IDs", 118, 114, 94, False, lambda s, e: self._opt(), TL)
        self.chkBox = self._chk(u"Box select mode", 214, 114, 140, False, lambda s, e: self._opt(), TL)
        self.chkFix = self._chk(u"Fixed scale (all times)", 8, 136, 158, False, lambda s, e: self.fix_changed(), TL)
        self._lbl(u"Unit:", 176, 140, 34, 20, TL)
        self.cbUnit = ComboBox()
        self.cbUnit.DropDownStyle = ComboBoxStyle.DropDownList
        self.cbUnit.SetBounds(212, 136, 70, 24)
        self.cbUnit.Items.Add(u"C")
        self.cbUnit.Items.Add(u"K")
        self.cbUnit.SelectedIndex = 0
        self.cbUnit.SelectedIndexChanged += lambda s, e: self.time_changed()
        self.side.Controls.Add(self.cbUnit)
        self.lblScale = self._lbl(u"", 8, 160, SW - 16, 18, TL)
        self.lblHead = self._lbl(u"   Node          X            Y           Temp", 8, 184, SW - 16, 18, TL)
        self.lb = ListBox()
        self.lb.Font = Font("Consolas", 9.0)
        self.lb.SetBounds(8, 204, SW - 16, H0 - 204 - 300)
        self.lb.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        self.lb.HorizontalScrollbar = True
        self.side.Controls.Add(self.lb)
        self._btn(u"Select all", 8, H0 - 292, 90, 26, lambda s, e: self.select_all(), BL)
        self._btn(u"Clear", 102, H0 - 292, 70, 26, lambda s, e: self.clear_sel(), BL)
        self._btn(u"Undo (Ctrl+Z)", 176, H0 - 292, 120, 26, lambda s, e: self.undo(), BL)
        self._lbl(u"Export times:", 8, H0 - 256, 100, 20, BL)
        self._btn(u"All", 120, H0 - 260, 50, 24, lambda s, e: self.check_times(True), BL)
        self._btn(u"None", 174, H0 - 260, 50, 24, lambda s, e: self.check_times(False), BL)
        self._btn(u"Current", 228, H0 - 260, 70, 24, lambda s, e: self.check_current(), BL)
        self.clb = CheckedListBox()
        self.clb.CheckOnClick = True
        self.clb.SetBounds(8, H0 - 234, SW - 16, 108)
        self.clb.Anchor = BL
        self.side.Controls.Add(self.clb)
        self._lbl(u"Separator:", 8, H0 - 114, 62, 20, BL)
        self.cbSep = ComboBox()
        self.cbSep.DropDownStyle = ComboBoxStyle.DropDownList
        self.cbSep.SetBounds(74, H0 - 118, 90, 24)
        for t in (u"Tab", u"Space", u"Comma"):
            self.cbSep.Items.Add(t)
        self.cbSep.SelectedIndex = 0
        self.cbSep.Anchor = BL
        self.side.Controls.Add(self.cbSep)
        self._lbl(u"Export unit:", 172, H0 - 114, 70, 20, BL)
        self.cbExpUnit = ComboBox()
        self.cbExpUnit.DropDownStyle = ComboBoxStyle.DropDownList
        self.cbExpUnit.SetBounds(244, H0 - 118, 60, 24)
        self.cbExpUnit.Items.Add(u"C")
        self.cbExpUnit.Items.Add(u"K")
        self.cbExpUnit.SelectedIndex = 0
        self.cbExpUnit.Anchor = BL
        self.side.Controls.Add(self.cbExpUnit)
        self.chkMv = self._chk(u"Add mesh section-3 value column (MESHVAL)", 8, H0 - 90, SW - 16, False, lambda s, e: None, BL)
        self._btn(u"Copy to clipboard", 8, H0 - 62, 160, 28, lambda s, e: self.do_copy(), BL)
        self._btn(u"Save TXT...", 176, H0 - 62, 120, 28, lambda s, e: self.do_save(), BL)
        self.lblStatus = self._lbl(u"", 8, H0 - 30, SW - 16, 24, BL)
        self.lblStatus.Font = Font("Segoe UI", 8.0)

    def _opt(self):
        self.view.mesh_on = self.chkMesh.Checked
        self.view.ids_on = self.chkIds.Checked
        self.view.box_mode = self.chkBox.Checked
        self.view.Cursor = Cursors.Cross if self.chkBox.Checked else Cursors.Default
        self.view.mark_dirty()

    # ----- yukleme -----
    def ask_files(self):
        sf = StartForm(self.mesh_path, self.temp_path)
        if sf.ShowDialog(self) == DialogResult.OK:
            self.load_files(sf.mesh_path, sf.temp_path)

    def load_files(self, mesh_path, temp_path):
        pf = ProgressForm(u"Loading ...")
        pf.Show()
        try:
            pf.update(0.0, u"Reading mesh file ...")
            nodes, raw, mv = parse_mesh(mesh_path)
            pf.update(0.0, u"Building model (%d elements) ..." % len(raw))
            model = build_model(nodes, raw)
            tf = TempFile(temp_path, pf.update)
        except Cancelled:
            return False
        except Exception as ex:
            pf.Close()
            MessageBox.Show(u"Dosya okunamadi:\n%s" % str(ex), TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error)
            return False
        finally:
            pf.Close()
        if not tf.times:
            MessageBox.Show(u"Temp dosyasinda TIME bilgisi olan hicbir blok bulunamadi.", TITLE,
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return False
        self.mesh_path, self.temp_path = mesh_path, temp_path
        self._undo = []
        self.nodes, self.mv = nodes, mv
        self.tf = tf
        self.times = tf.times
        self.grange = None
        self.view.set_model(model)
        self.lblMesh.Text = u"Mesh: %s" % System.IO.Path.GetFileName(mesh_path)
        self.lblTemp.Text = u"Temp: %s" % System.IO.Path.GetFileName(temp_path)
        self.cbTime.Items.Clear()
        self.clb.Items.Clear()
        for k, t in enumerate(self.times):
            self.cbTime.Items.Add(u"%.6g" % t)
            self.clb.Items.Add(u"%.6g" % t, k == 0)
        msg = u"%d nodes, %d elements (%d skipped) | %d time(s) | %d section-3 values" % (
            len(nodes), len(model["elems"]), model["skipped"], len(self.times), len(mv))
        if len(tf.blocks) != len(self.times):
            msg += u" | WARNING: %d blocks, %d unique times (same TIME repeated: last one used)" % (len(tf.blocks), len(self.times))
        self.lblStatus.Text = msg
        self.cbTime.SelectedIndex = 0                 # time_changed() okur ve boyar
        first = tf.block_full(self.times[0])
        missing = len([n_ for n_ in model["used"] if n_ not in first])
        if missing:
            self.lblStatus.Text = msg + u" | %d mesh nodes without temp" % missing
        self.selection_changed()
        return True

    # ----- time / secim -----
    def _off(self):
        """Ekran birimi icin ofset (dosya degerleri C)."""
        return 273.15 if self.cbUnit.SelectedIndex == 1 else 0.0

    def fix_changed(self):
        if self.chkFix.Checked and self.grange is None and self.tf is not None:
            pf = ProgressForm(u"Scanning all times ...")
            pf.Show()
            ok = False
            try:
                self.grange = self.tf.global_range(pf.update)
                ok = True
            except Cancelled:
                pass
            except Exception as ex:
                MessageBox.Show(u"Okunamadi:\n%s" % str(ex), TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error)
            finally:
                pf.Close()
            if not ok:
                self.chkFix.Checked = False           # handler tekrar calisir
                return
        self.time_changed()

    def time_changed(self):
        i = self.cbTime.SelectedIndex
        if self.tf is None or i < 0 or i >= len(self.times):
            return
        t = self.times[i]
        self.Cursor = Cursors.WaitCursor
        try:
            tm = self.tf.block_full(t)
        except Exception as ex:
            self.Cursor = Cursors.Default
            MessageBox.Show(u"Time okunamadi:\n%s" % str(ex), TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error)
            return
        self.Cursor = Cursors.Default
        self.cur_time = t
        self.cur_map = tm
        off = self._off()
        un = u"K" if off else u"C"
        if self.chkFix.Checked and self.grange:
            lo, hi = self.grange
        else:
            used = self.view.model["used"]
            vals = [v for n_, v in tm.items() if n_ in used]
            lo, hi = (min(vals), max(vals)) if vals else (0.0, 1.0)
        case, rpm = self.tf.info(t)
        self.lblTinfo.Text = u"Case %s | RPM %s" % ("-" if case is None else case, "-" if rpm is None else "%.6g" % rpm)
        self.lblScale.Text = u"Scale [%s]: %.3f ... %.3f" % (un, lo + off, hi + off)
        self.lblHead.Text = u"   Node          X            Y         Temp [%s]" % un
        self.view.unit_txt = un
        self.view.set_temps(tm, lo, hi, off)
        self.selection_changed()

    def selection_changed(self):
        MAXLIST = 2000
        off = self._off()
        tm = self.cur_map
        sel = sorted(self.view.sel)
        items = []
        for nid in sel[:MAXLIST]:
            X, Y = self.nodes[nid]
            v = tm.get(nid)
            items.append(u"%6d %12.5f %12.5f %11s" % (nid, X, Y, u"%.3f" % (v + off) if v is not None else u"NaN"))
        if len(sel) > MAXLIST:
            items.append(u"... +%d more nodes (selected, not listed)" % (len(sel) - MAXLIST))
        self.lb.BeginUpdate()
        try:
            self.lb.Items.Clear()
            self.lb.Items.AddRange(System.Array[object](items))
        finally:
            self.lb.EndUpdate()
        self.view.Invalidate()

    def push_undo(self):
        self._undo.append(set(self.view.sel))
        if len(self._undo) > 50:
            del self._undo[0]

    def undo(self):
        if self._undo:
            self.view.sel = self._undo.pop()
            self.selection_changed()
            self.lblStatus.Text = u"Undo: %d node(s) selected" % len(self.view.sel)

    def _on_key(self, sender, e):
        if e.Control and e.KeyCode == Keys.Z:
            self.undo()
            e.Handled = True

    def select_all(self):
        self.push_undo()
        self.view.sel = set(self.view.dn.keys())
        self.selection_changed()

    def clear_sel(self):
        if self.view.sel:
            self.push_undo()
        self.view.sel = set()
        self.selection_changed()

    def check_times(self, state):
        for k in range(self.clb.Items.Count):
            self.clb.SetItemChecked(k, state)

    def check_current(self):
        i = self.cbTime.SelectedIndex
        for k in range(self.clb.Items.Count):
            self.clb.SetItemChecked(k, k == i)

    # ----- cikti -----
    def build_text(self):
        sel = sorted(self.view.sel)
        if not sel:
            MessageBox.Show(u"Once en az bir node sec.", TITLE, MessageBoxButtons.OK, MessageBoxIcon.Information)
            return None
        sep = [u"\t", u"  ", u","][max(0, self.cbSep.SelectedIndex)]
        times = [self.times[k] for k in range(self.clb.Items.Count) if self.clb.GetItemChecked(k)]
        if not times:
            MessageBox.Show(u"Export times listesinde en az bir time isaretle.", TITLE,
                            MessageBoxButtons.OK, MessageBoxIcon.Information)
            return None
        if len(sel) * len(times) > 20000000:
            if MessageBox.Show(u"%d node x %d time = %d deger. Bu cok buyuk, bellek yetmeyebilir. Devam edilsin mi?"
                               % (len(sel), len(times), len(sel) * len(times)), TITLE,
                               MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes:
                return None
        use_mv = self.chkMv.Checked
        off = 273.15 if self.cbExpUnit.SelectedIndex == 1 else 0.0
        cols = []
        pf = ProgressForm(u"Exporting ...")
        pf.Show()
        try:
            for ti, t in enumerate(times):
                if not pf.update(float(ti) / len(times), u"Reading time %s   (%d / %d)" % (u"%.6g" % t, ti + 1, len(times))):
                    return None
                cols.append(self.tf.block_nodes(t, sel))
        except Exception as ex:
            MessageBox.Show(u"Okunamadi:\n%s" % str(ex), TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error)
            return None
        finally:
            pf.Close()
        head = [u"NodeID", u"X_Loc", u"Y_Loc"] + [u"Time%d_T" % (k + 1) for k in range(len(times))]
        if use_mv:
            head.append(u"MeshVal")
        out = [sep.join(head)]
        for j, nid in enumerate(sel):
            X, Y = self.nodes[nid]
            row = [u"%d" % nid, u"%.10g" % X, u"%.10g" % Y]
            for c in cols:
                v = c[j]
                row.append(u"%.10g" % (v + off) if v is not None else u"NaN")
            if use_mv:
                mvv = self.mv.get(nid)
                row.append(u"%.9g" % mvv if mvv is not None else u"NaN")
            out.append(sep.join(row))
        return u"\r\n".join(out) + u"\r\n"

    def do_copy(self):
        txt = self.build_text()
        if txt is not None:
            Clipboard.SetText(txt)
            self.lblStatus.Text = u"Copied: %d node(s), unit %s" % (len(self.view.sel), self.cbExpUnit.Text)

    def do_save(self):
        txt = self.build_text()
        if txt is None:
            return
        d = SaveFileDialog()
        d.Title = u"Save TXT"
        d.Filter = u"Text files (*.txt)|*.txt|All files (*.*)|*.*"
        d.FileName = u"node_temperatures.txt"
        if d.ShowDialog() == DialogResult.OK:
            try:
                System.IO.File.WriteAllText(d.FileName, txt)
                self.lblStatus.Text = u"Saved: %s" % d.FileName
            except Exception as ex:
                MessageBox.Show(u"Kaydedilemedi:\n%s" % str(ex), TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error)


def main():
    owner = mechanical_owner()
    f = ViewerForm()
    mp, tp = u"", u""
    while True:
        sf = StartForm(mp, tp)
        r = sf.ShowDialog(owner) if owner is not None else sf.ShowDialog()
        if r != DialogResult.OK:
            return
        mp, tp = sf.mesh_path, sf.temp_path
        if f.load_files(mp, tp):
            break
    if owner is not None:
        f.ShowDialog(owner)
    else:
        f.ShowDialog()


main()
