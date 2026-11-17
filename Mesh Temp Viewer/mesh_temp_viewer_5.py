# -*- coding: utf-8 -*-
# =====================================================================
#  Mesh Temp Viewer 5   (Converter)   (Ansys Mechanical, IronPython 2.7)
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
#  v5: .disp (deflection) dosyasi opsiyonel; Show butonu ile contour; eleman gizleme (Hide elements);
#      zoom-out icin on-hesaplanmis genel gorunum (mip); sol ust yazilar duzeltildi
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
                                  CheckedListBox, TextBox, FormBorderStyle, Application, Timer, ProgressBar, ToolTip)
from System.Drawing import (Size, Point, Font, FontStyle, Color, PointF, Pen, SolidBrush, Bitmap, Graphics,
                            RectangleF, GraphicsUnit)
from System.Drawing.Drawing2D import SmoothingMode, DashStyle, InterpolationMode

TITLE = u"Mesh Temp Viewer 5"
SW = 360            # sag panel genisligi
SUB_TRI = 6         # ucgen eleman alt bolme sayisi (renk yumusakligi)
SUB_QUAD = 4        # dortgen eleman alt bolme sayisi


# ---------------------------------------------------------------
# Dosya okuma
# ---------------------------------------------------------------
def tofloat(s):
    """Fortran tarzi sayilari da okur (0.123-100, 0.0d0, 1.5D+02)."""
    s = s.replace("d", "E").replace("D", "E")
    try:
        return float(s)
    except ValueError:
        m = re.match(r"^([-+]?\d*\.?\d+)([-+]\d+)$", s)
        if m:
            return float(m.group(1) + "E" + m.group(2))
        raise


def fnum(tok):
    try:
        return float(tok)
    except ValueError:
        return tofloat(tok)


def read_lines(path):
    return [unicode(l) for l in System.IO.File.ReadAllLines(path)]


class Cancelled(Exception):
    pass


ENC = System.Text.Encoding.GetEncoding(28591)       # latin-1: 1 karakter = 1 byte (dosya konumu = karakter konumu)
HDR_RE = re.compile(r"^(CASE|NODE|TIME|RPM)\s*=\s*(\S+)")
KINDS = {"temp": {"key": u"TEMPERATURES", "ncomp": 1},
         "disp": {"key": u"DEFLECTIONS", "ncomp": 3}}        # disp: UX UY UZ (TX TY TZ okunmaz)


def parse_header_text(text):
    """Bir blogun ilk satirlari -> {'case','time','rpm','data_off'} (data_off = ilk veri satirinin indeksi)."""
    h = {"case": None, "time": None, "rpm": None, "data_off": None}
    for k, ln in enumerate(text.split("\n")[:14]):
        u = ln.strip().upper()
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
            continue
        if u.startswith("NODE"):                          # NODE TEMPERATURE (...) / NODE UX UY UZ TX TY TZ
            h["data_off"] = k + 1
            break
    return h


def parse_lines(lines, d0, ncomp=1):
    """Veri satirlari -> (node id listesi, [sutun listeleri]) (dosyadaki sirayla)."""
    ids = []
    cols = [[] for _ in range(ncomp)]
    end = ncomp + 1
    for ln in lines[d0:]:
        tk = ln.split()
        if len(tk) <= ncomp:
            continue
        try:
            a = int(tk[0])
            vs = [float(x) for x in tk[1:end]]
        except ValueError:
            try:
                a = int(tk[0])
                vs = [tofloat(x) for x in tk[1:end]]
            except ValueError:
                continue
        ids.append(a)
        for c in range(ncomp):
            cols[c].append(vs[c])
    return ids, cols


class FieldFile(object):
    """Cok buyuk .temp / .disp dosyalari icin TEMBEL okuyucu.
    Acilista dosya bir kez taranir (sadece blok baslarinin konumu bulunur); degerler bir time secilince
    (veya export'ta) sadece o blok okunur. Node sirasi bloklarda ayni kabul edilir; export'ta secili
    node'larin satir sirasi ilk bloktan ogrenilir ve diger bloklarda dogrudan o satira gidilir (dogrulanir)."""

    def __init__(self, path, kind, progress=None):
        self.path = path
        self.kind = kind
        self.key = KINDS[kind]["key"]
        self.ncomp = KINDS[kind]["ncomp"]
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
        key = self.key
        fs = self._open()
        try:
            while True:
                n = fs.Read(buf, 0, CH)
                if n <= 0:
                    break
                text = tail + ENC.GetString(buf, 0, n)
                tbase = pos - len(tail)
                i = text.find(key)
                while i >= 0:
                    if i > 0:
                        prev = text[i - 1]
                    elif tbase == 0:
                        prev = u"\n"
                    else:
                        prev = None
                    if prev in (u"\n", u"\r"):
                        hits.add(tbase + i)
                    i = text.find(key, i + 1)
                pos += n
                tail = text[-20:]
                if progress is not None and not progress(
                        float(pos) / max(self.size, 1),
                        u"Scanning %s file ... %d%%   (%d time blocks found)"
                        % (self.kind, int(100.0 * pos / max(self.size, 1)), len(hits))):
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

    def is_cached(self, t):
        return t in self._cache

    def block_cols(self, t):
        """Bir time'in (node idleri, sutunlar) (son 2 time onbellekte)."""
        if t in self._cache:
            return self._cache[t]
        k = self.block_of[t]
        ids, cols = parse_lines(self.read_block(k).split("\n"), self._d0(k), self.ncomp)
        if self.pos_of is None:
            self.pos_of = dict((n, i) for i, n in enumerate(ids))
        self._cache[t] = (ids, cols)
        self._order.append(t)
        if len(self._order) > 2:
            self._cache.pop(self._order.pop(0), None)
        return ids, cols

    def scalar(self, t, comp):
        """{node: deger}; comp = sutun indeksi veya -1 (vektor buyuklugu)."""
        ids, cols = self.block_cols(t)
        if comp >= 0:
            return dict(zip(ids, cols[comp]))
        if self.ncomp == 3:
            return dict(zip(ids, [math.sqrt(a * a + b * b + c * c) for a, b, c in zip(*cols)]))
        return dict(zip(ids, cols[0]))

    def block_nodes(self, t, nids, comps):
        """Sadece verilen node'larin degerleri: her comp icin bir liste (nids sirasinda; yoksa None).
        Hizli yol: node'un satir sirasi bellidir, tum satirlar ayristirilmaz."""
        if self.pos_of is None:
            self.block_cols(t)
        if t in self._cache:
            ids, cols = self._cache[t]
            idx = dict(zip(ids, range(len(ids))))
            return [[(cols[c][idx[n]] if n in idx else None) for n in nids] for c in comps]
        k = self.block_of[t]
        d0 = self._d0(k)
        po = self.pos_of
        idxs = [po.get(n) for n in nids]
        mx = max([i for i in idxs if i is not None] or [0])
        lines = self.read_block(k).split("\n", d0 + mx + 2)
        res = [[] for _ in comps]
        for n, i in zip(nids, idxs):
            if i is None:
                for r in res:
                    r.append(None)
                continue
            try:
                tk = lines[d0 + i].split()
                if int(tk[0]) != n:
                    raise ValueError("order")
                for j, c in enumerate(comps):
                    res[j].append(fnum(tk[1 + c]))
            except (ValueError, IndexError):
                ids, cols = self.block_cols(t)          # sira farkli: bu blok icin tam ayristirma
                idx = dict(zip(ids, range(len(ids))))
                return [[(cols[c][idx[m_]] if m_ in idx else None) for m_ in nids] for c in comps]
        return res

    def global_range(self, progress, comp):
        """Tum time'lardaki min / max (butun dosya okunur)."""
        lo = hi = None
        nt = len(self.times)
        for i, t in enumerate(self.times):
            if not progress(float(i) / nt, u"Scanning all times for min / max ...  %d / %d" % (i + 1, nt)):
                raise Cancelled()
            k = self.block_of[t]
            ids, cols = parse_lines(self.read_block(k).split("\n"), self._d0(k), self.ncomp)
            if comp >= 0:
                vals = cols[comp]
            elif self.ncomp == 3:
                vals = [math.sqrt(a * a + b * b + c * c) for a, b, c in zip(*cols)]
            else:
                vals = cols[0]
            if vals:
                a_, b_ = min(vals), max(vals)
                lo = a_ if lo is None else min(lo, a_)
                hi = b_ if hi is None else max(hi, b_)
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
    el["hid"] = False
    el["c"] = (sum(p[0] for p in el["cpoly"]) / len(el["cpoly"]), sum(p[1] for p in el["cpoly"]) / len(el["cpoly"]))
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
MIP_MAX = 4096                 # on-hesaplanmis genel gorunumun (mip) en buyuk kenari [piksel]
MIP_MIN_ELEMS = 15000          # bundan kucuk modellerde mip kullanilmaz


def point_in_poly(x, y, pts):
    inside = False
    n = len(pts)
    j = n - 1
    for i in range(n):
        xi, yi = pts[i]
        xj, yj = pts[j]
        if (yi > y) != (yj > y) and x < (xj - xi) * (y - yi) / (yj - yi) + xi:
            inside = not inside
        j = i
    return inside


class SceneView(Panel):
    """Cizim iki katmanli:
    - MIP: tum modelin sabit cozunurlukte (en fazla 4096 px) bir kez cizilmis hali. Uzaklastirilmis gorunumde
      (buyuk modellerde) zoom / pan sadece bu resmi kaydirir, yeniden cizim yok.
    - CACHE: yakinlastirilmis gorunumde sadece ekranda gorunen elemanlarin cizimi (bitmap). Zoom / pan sirasinda
      eski bitmap olceklenip gosterilir, net cizim hareket durunca (150 ms) yapilir.
    Secim isaretleri her zaman ustten cizilir (modeli yeniden cizmez)."""
    def __init__(self, owner):
        Panel.__init__(self)
        self.owner = owner
        self.DoubleBuffered = True
        self.BackColor = Color.White
        self.model = None
        self.dn = {}                                  # node -> cizim koordinati
        self._ngrid = None
        self._egrid = None
        self._cell = 1.0
        self.avg_el = 1.0
        self.sel = set()
        self.s = 1.0
        self.px = self.py = self.cx = self.cy = 0.0
        self.mesh_on = True
        self.ids_on = False
        self.box_mode = False
        self.hide_mode = False
        self.hide_stack = []
        self.hid_nodes = set()
        self.node_cnt = None
        self.n_hidden = 0
        self.title = u""
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
        self._mip = None
        self._mip_dirty = True
        self._mip_ok = False
        self._mip_geom = None
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
        self._ngrid = self._egrid = None
        self.sel = set()
        self.hide_stack = []
        self.hid_nodes = set()
        self.node_cnt = None
        self.n_hidden = 0
        self.has_t = False
        x0, x1, y0, y1 = model["box"]
        bw, bh = max(x1 - x0, 1e-12), max(y1 - y0, 1e-12)
        ms = MIP_MAX / max(bw, bh)
        self._mip_geom = (ms, int(bw * ms) + 3, int(bh * ms) + 3, x0, y0)
        self._mip_ok = len(els) > MIP_MIN_ELEMS
        self._fit_pending = True
        self._dirty = True
        self._mip_dirty = True
        self.Invalidate()

    def set_field(self, tmap, vmin, vmax, off, title):
        self.tmap, self.vmin, self.vmax, self.off, self.title = tmap, vmin, vmax, off, title
        self.has_t = True
        self.recolor()
        self.mark_dirty()

    def clear_field(self):
        self.has_t = False
        self.tmap = {}
        if self.model is not None:
            for el in self.model["elems"]:
                for sp in el["subs"]:
                    sp["ci"] = 256
                el["ci"] = 256
        self.mark_dirty()

    def mark_dirty(self):
        self._dirty = True
        self._mip_dirty = True
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

    # ----- eleman gizleme -----
    def _ensure_node_cnt(self):
        if self.node_cnt is None:
            cnt = {}
            for el in self.model["elems"]:
                if not el["hid"]:
                    for n in el["nodes"]:
                        cnt[n] = cnt.get(n, 0) + 1
            self.node_cnt = cnt

    def _hide_elems(self, els):
        new = [e for e in els if not e["hid"]]
        if not new:
            return
        self._ensure_node_cnt()
        for e in new:
            e["hid"] = True
            for n in e["nodes"]:
                c = self.node_cnt.get(n, 0) - 1
                self.node_cnt[n] = c
                if c <= 0:
                    self.hid_nodes.add(n)
        self.hide_stack.append(new)
        self.n_hidden += len(new)
        self.owner.note(u"Hidden %d element(s)  (total hidden: %d)" % (len(new), self.n_hidden))
        self.mark_dirty()

    def _unhide(self, els):
        self._ensure_node_cnt()
        for e in els:
            if e["hid"]:
                e["hid"] = False
                self.n_hidden -= 1
                for n in e["nodes"]:
                    self.node_cnt[n] = self.node_cnt.get(n, 0) + 1
                    self.hid_nodes.discard(n)

    def undo_hide(self):
        if self.hide_stack:
            g = self.hide_stack.pop()
            self._unhide(g)
            self.owner.note(u"Restored %d element(s)  (total hidden: %d)" % (len(g), self.n_hidden))
            self.mark_dirty()

    def unhide_all(self):
        if self.hide_stack:
            while self.hide_stack:
                self._unhide(self.hide_stack.pop())
            self.owner.note(u"All elements visible again")
            self.mark_dirty()

    # ----- gridler (ilk kullanimda kurulur) -----
    def _ensure_grids(self):
        if self._ngrid is not None:
            return
        self.Cursor = Cursors.WaitCursor
        try:
            c = self._cell
            ng = {}
            for nid, (x, y) in self.dn.items():
                ng.setdefault((int(math.floor(x / c)), int(math.floor(y / c))), []).append(nid)
            eg = {}
            for k, el in enumerate(self.model["elems"]):
                x0, x1, y0, y1 = el["bb"]
                for a in range(int(math.floor(x0 / c)), int(math.floor(x1 / c)) + 1):
                    for b in range(int(math.floor(y0 / c)), int(math.floor(y1 / c)) + 1):
                        eg.setdefault((a, b), []).append(k)
            self._ngrid, self._egrid = ng, eg
        finally:
            self._restore_cursor()

    def _restore_cursor(self):
        self.Cursor = Cursors.Hand if self.hide_mode else (Cursors.Cross if self.box_mode else Cursors.Default)

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
    def _use_mip(self):
        return self._mip_ok and self.s <= self._mip_geom[0] * 1.02

    def _render_mip(self):
        ms, mw, mh, x0, y0 = self._mip_geom
        self.Cursor = Cursors.WaitCursor
        try:
            bmp = Bitmap(mw, mh)
            g = Graphics.FromImage(bmp)
            try:
                g.Clear(Color.White)
                self._draw_model(g, ms, 1.0 - x0 * ms, 1.0 - y0 * ms, float(mw), float(mh), True)
            finally:
                g.Dispose()
        finally:
            self._restore_cursor()
        if self._mip is not None:
            self._mip.Dispose()
        self._mip = bmp
        self._mip_dirty = False

    def _draw_mip(self, g):
        if self._mip is None:
            return
        ms, mw, mh, x0, y0 = self._mip_geom
        s = self.s
        ox, oy = self._ofs()
        k = s / ms
        dx, dy = (x0 - 1.0 / ms) * s + ox, (y0 - 1.0 / ms) * s + oy
        W, H = float(self.Width), float(self.Height)
        a0, a1 = max(0.0, (0.0 - dx) / k), min(float(mw), (W - dx) / k)
        b0, b1 = max(0.0, (0.0 - dy) / k), min(float(mh), (H - dy) / k)
        if a1 <= a0 or b1 <= b0:
            return
        g.InterpolationMode = InterpolationMode.Bilinear
        g.DrawImage(self._mip, RectangleF(dx + a0 * k, dy + b0 * k, (a1 - a0) * k, (b1 - b0) * k),
                    RectangleF(a0, b0, a1 - a0, b1 - b0), GraphicsUnit.Pixel)

    def _render_cache(self):
        W, H = self.Width, self.Height
        if W < 10 or H < 10:
            return
        bmp = Bitmap(W, H)
        g = Graphics.FromImage(bmp)
        try:
            g.Clear(Color.White)
            ox, oy = self._ofs()
            self._draw_model(g, self.s, ox, oy, float(W), float(H), False)
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
            g.DrawString(u"Mesh / Temp / Disp dosyasi yuklenmedi.", self.f_ui, SolidBrush(Color.DimGray), 8.0, 8.0)
            return
        try:
            if self._use_mip():
                if self._mip is None or self._mip_dirty:
                    self._render_mip()
                self._draw_mip(g)
                self._stale = False
                self._timer.Stop()
            else:
                if self._cache is None or self._dirty:
                    self._render_cache()
                    self._stale = False
                    self._timer.Stop()
                if self._stale:
                    r = self.s / self._cs
                    if self._cache is not None and (0.2 < r < 5.0 or self._mip is None):
                        ox, oy = self._ofs()
                        dx, dy = ox - self._cofs[0] * r, oy - self._cofs[1] * r
                        g.DrawImage(self._cache, RectangleF(dx, dy, self._cache.Width * r, self._cache.Height * r))
                    else:
                        self._draw_mip(g)
                elif self._cache is not None:
                    g.DrawImage(self._cache, 0, 0)
            self._draw_selection(g)
        except Exception as ex:
            g.DrawString(u"Draw error: %s" % str(ex), self.f_ui, SolidBrush(Color.Firebrick), 8.0, 24.0)
        g.SmoothingMode = SmoothingMode.AntiAlias
        if self.hide_mode:
            txt = u"HIDE ELEMENTS: click an element to hide it | right-drag / Shift+drag = hide elements in the box | hidden: %d" % self.n_hidden
        else:
            txt = (u"Selected nodes: %d    |  wheel: zoom  |  drag: pan  |  click: select node  |  right-drag / Shift+drag / Box mode: box select  |  Ctrl+Z: undo"
                   % len(self.sel))
        g.DrawString(txt, self.f_ui, SolidBrush(Color.DarkOrange if self.hide_mode else Color.DimGray), 8.0, 4.0)
        if self.has_t:
            self._draw_legend(g, self.f_ui)
        if self._rb is not None:
            x0, y0, x1, y1 = self._rb
            pen = Pen(Color.DarkOrange, 1.0)
            pen.DashStyle = DashStyle.Dash
            g.DrawRectangle(pen, min(x0, x1), min(y0, y1), abs(x1 - x0), abs(y1 - y0))

    def _draw_model(self, g, s, ox, oy, W, H, mip):
        g.SmoothingMode = SmoothingMode.HighSpeed
        br, pens = self.brushes, self.pens
        mesh_on = self.mesh_on
        pm = self.p_mesh
        for el in self.model["elems"]:
            if el["hid"]:
                continue
            x0, x1, y0, y1 = el["bb"]
            sx0, sx1, sy0, sy1 = x0 * s + ox, x1 * s + ox, y0 * s + oy, y1 * s + oy
            if sx1 < -2 or sx0 > W + 2 or sy1 < -2 or sy0 > H + 2:
                continue
            px = max(sx1 - sx0, sy1 - sy0)                 # en uzun kenar (ince / uzun elemanlar kaybolmasin)
            if px < 1.0:
                g.FillRectangle(br[el["ci"]], sx0 - 0.5, sy0 - 0.5, 1.6, 1.6)
                continue
            if px < 6.0:
                pts = System.Array[PointF]([PointF(clampf(p[0] * s + ox), clampf(p[1] * s + oy)) for p in el["cpoly"]])
                g.FillPolygon(br[el["ci"]], pts)
            else:
                for sp in el["subs"]:
                    pts = System.Array[PointF]([PointF(clampf(p[0] * s + ox), clampf(p[1] * s + oy)) for p in sp["pts"]])
                    g.FillPolygon(br[sp["ci"]], pts)
                    g.DrawPolygon(pens[sp["ci"]], pts)
            if mesh_on and px >= 4.0:
                ring = System.Array[PointF]([PointF(clampf(p[0] * s + ox), clampf(p[1] * s + oy)) for p in el["outline"]])
                g.DrawPolygon(pm, ring)
        if mip:
            return
        g.SmoothingMode = SmoothingMode.AntiAlias
        px_el = self.avg_el * s
        if px_el >= 10.0:                                  # node noktalari (yeterince yakinlasinca)
            hid = self.hid_nodes
            vis = []
            for nid, (x, y) in self.dn.items():
                sx, sy = x * s + ox, y * s + oy
                if -10 <= sx <= W + 10 and -10 <= sy <= H + 10 and nid not in hid:
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
        hid = self.hid_nodes
        for nid in self.sel:
            p = dn.get(nid)
            if p is None or nid in hid:
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
        g.DrawString(self.title, f, SolidBrush(Color.Black), x - 20.0, y0 - 16.0)
        br = SolidBrush(Color.Black)
        for frac in (0.0, 0.25, 0.5, 0.75, 1.0):
            val = self.vmax - (self.vmax - self.vmin) * frac + self.off
            yy = y0 + (y1 - y0) * frac
            g.DrawLine(Pen(Color.Gray, 1.0), x + w, yy, x + w + 4.0, yy)
            g.DrawString(u"%.4g" % val, f, br, x + w + 5.0, yy - 8.0)

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
            self._restore_cursor()
            return
        if self._down is None:
            return
        x0, y0 = self._down[0], self._down[1]
        if self._drag and self._rb is not None:
            xa, xb = sorted((x0, e.X))
            ya, yb = sorted((y0, e.Y))
            ox, oy = self._ofs()
            s = self.s
            if self.hide_mode:
                hit = []
                for el in self.model["elems"]:
                    if el["hid"]:
                        continue
                    cx, cy = el["c"]
                    sx, sy = cx * s + ox, cy * s + oy
                    if xa <= sx <= xb and ya <= sy <= yb:
                        hit.append(el)
                self._hide_elems(hit)
            else:
                hid = self.hid_nodes
                hit = set()
                for nid, (x, y) in self.dn.items():
                    sx, sy = x * s + ox, y * s + oy
                    if xa <= sx <= xb and ya <= sy <= yb and nid not in hid:
                        hit.add(nid)
                if hit - self.sel:
                    self.owner.push_undo()
                    self.sel |= hit
                    self.owner.selection_changed()
        elif not self._drag and self._btn == MouseButtons.Left:
            if self.hide_mode:
                el = self._pick_elem(x0, y0)
                if el is not None:
                    self._hide_elems([el])
            else:
                self._pick(x0, y0)
        self._btn = None
        self._down = None
        self._drag = False
        self._rb = None
        self.Invalidate()

    def _pick_elem(self, mx, my):
        self._ensure_grids()
        ox, oy = self._ofs()
        s = self.s
        X, Y = (mx - ox) / s, (my - oy) / s
        c = self._cell
        els = self.model["elems"]
        best = None
        for k in self._egrid.get((int(math.floor(X / c)), int(math.floor(Y / c))), []):
            e = els[k]
            if e["hid"]:
                continue
            x0, x1, y0, y1 = e["bb"]
            if x0 <= X <= x1 and y0 <= Y <= y1 and point_in_poly(X, Y, e["cpoly"]):
                best = e
        return best

    def _pick(self, mx, my):
        self._ensure_grids()
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
                    lst = self._ngrid.get((a, b))
                    if lst:
                        cands.extend(lst)
        hid = self.hid_nodes
        best, bn = 12.0, None
        for nid in cands:
            if nid in hid:
                continue
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
    """Ilk ekran: .mesh (zorunlu) + .temp ve/veya .disp (opsiyonel, en az biri) dosyalarini sec."""
    def __init__(self, mesh_path=u"", temp_path=u"", disp_path=u""):
        self.Text = TITLE + u" - Input files"
        self.ClientSize = Size(600, 232)
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.mesh_path, self.temp_path, self.disp_path = mesh_path, temp_path, disp_path
        self.tbMesh = self._row(u".mesh file:", 20, mesh_path, u"Mesh files (*.mesh)|*.mesh|All files (*.*)|*.*")
        self.tbTemp = self._row(u".temp file:", 64, temp_path, u"Temp files (*.temp)|*.temp|All files (*.*)|*.*")
        self.tbDisp = self._row(u".disp file:", 108, disp_path, u"Disp files (*.disp)|*.disp|All files (*.*)|*.*")
        l = Label()
        l.Text = u".mesh is required. .temp and .disp are optional (at least one of them). Clear the box to skip a file."
        l.SetBounds(10, 150, 580, 20)
        self.Controls.Add(l)
        ok = Button()
        ok.Text = u"Load"
        ok.SetBounds(410, 186, 85, 30)
        ok.Click += lambda s, e: self._ok()
        self.Controls.Add(ok)
        cn = Button()
        cn.Text = u"Cancel"
        cn.SetBounds(505, 186, 85, 30)
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
        tb.SetBounds(90, y, 410, 24)
        self.Controls.Add(tb)
        b = Button()
        b.Text = u"Browse..."
        b.SetBounds(510, y - 1, 80, 26)

        def click(s, e):
            f = pick_open(text, flt, tb.Text.strip())
            if f:
                tb.Text = f
        b.Click += click
        self.Controls.Add(b)
        return tb

    def _ok(self):
        m = self.tbMesh.Text.strip().strip('"')
        t = self.tbTemp.Text.strip().strip('"')
        d = self.tbDisp.Text.strip().strip('"')
        if not m or not System.IO.File.Exists(m):
            MessageBox.Show(u".mesh dosyasi secilmedi veya bulunamadi.", TITLE, MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        if not t and not d:
            MessageBox.Show(u"En az bir .temp veya .disp dosyasi sec.", TITLE, MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        for path, nm in ((t, u".temp"), (d, u".disp")):
            if path and not System.IO.File.Exists(path):
                MessageBox.Show(u"%s dosyasi bulunamadi." % nm, TITLE, MessageBoxButtons.OK, MessageBoxIcon.Warning)
                return
        self.mesh_path, self.temp_path, self.disp_path = m, t, d
        self.DialogResult = DialogResult.OK


class ViewerForm(Form):
    def __init__(self):
        self.Text = TITLE
        wa = Screen.PrimaryScreen.WorkingArea
        self.ClientSize = Size(int(0.85 * min(1600, wa.Width - 30)), int(min(920, wa.Height - 60)))
        self.MinimumSize = Size(860, 740)
        self.StartPosition = FormStartPosition.CenterScreen
        self.nodes, self.mv = {}, {}
        self.tf = self.df = None
        self.times = []                       # temp ve disp time'larinin birlesimi
        self.fields = []
        self.shown_times = []
        self._shown = False
        self._updating = False
        self.cur_time = None
        self.cur_map = {}
        self.cur_fld = None
        self.grange = {}
        self._undo = []
        self.mesh_path = self.temp_path = self.disp_path = u""
        self._tt = ToolTip()
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
    def _lbl(self, text, x, y, w, h, anchor=None, ellipsis=False):
        l = Label()
        l.AutoSize = False
        l.AutoEllipsis = ellipsis
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

    def _combo(self, x, y, w, anchor):
        c = ComboBox()
        c.DropDownStyle = ComboBoxStyle.DropDownList
        c.SetBounds(x, y, w, 24)
        c.Anchor = anchor
        self.side.Controls.Add(c)
        return c

    def _build_side(self):
        H0 = self.side.Height
        TL = AnchorStyles.Top | AnchorStyles.Left
        BL = AnchorStyles.Bottom | AnchorStyles.Left
        self.lblMesh = self._lbl(u"Mesh: -", 8, 6, SW - 16, 16, TL, True)
        self.lblTemp = self._lbl(u"Temp: -", 8, 23, SW - 16, 16, TL, True)
        self.lblDisp = self._lbl(u"Disp: -", 8, 40, SW - 16, 16, TL, True)
        self._btn(u"Open files...", 8, 62, 120, 26, lambda s, e: self.ask_files(), TL)
        self._btn(u"Fit view", 136, 62, 90, 26, lambda s, e: self.view.fit_all(), TL)
        self._lbl(u"Field:", 8, 100, 40, 18, TL)
        self.cbField = self._combo(50, 96, 150, TL)
        self.cbField.SelectedIndexChanged += lambda s, e: self.field_changed()
        self._lbl(u"Unit:", 208, 100, 34, 18, TL)
        self.cbUnit = self._combo(244, 96, 60, TL)
        self.cbUnit.Items.Add(u"C")
        self.cbUnit.Items.Add(u"K")
        self.cbUnit.SelectedIndex = 0
        self.cbUnit.SelectedIndexChanged += lambda s, e: self.unit_changed()
        self._lbl(u"Time [s]:", 8, 130, 62, 18, TL)
        self.cbTime = self._combo(72, 126, 120, TL)
        self.cbTime.SelectedIndexChanged += lambda s, e: self.time_changed()
        self.lblTinfo = self._lbl(u"", 198, 130, SW - 206, 18, TL)
        self.chkMesh = self._chk(u"Mesh lines", 8, 156, 104, True, lambda s, e: self._opt_draw(), TL)
        self.chkIds = self._chk(u"Node IDs", 114, 156, 90, False, lambda s, e: self._opt_draw(), TL)
        self.chkBox = self._chk(u"Box select mode", 206, 156, 148, False, lambda s, e: self._opt_mode(), TL)
        self.chkHide = self._chk(u"Hide elements", 8, 180, 150, False, lambda s, e: self._opt_mode(), TL)
        self._btn(u"Undo hide", 162, 178, 86, 24, lambda s, e: self.view.undo_hide(), TL)
        self._btn(u"Unhide all", 252, 178, 96, 24, lambda s, e: self.view.unhide_all(), TL)
        self.chkFix = self._chk(u"Fixed scale (all times)", 8, 208, 172, False, lambda s, e: self.fix_changed(), TL)
        self.lblScale = self._lbl(u"", 184, 211, SW - 192, 18, TL)
        self.lblHead = self._lbl(u"   Node          X            Y           Value", 8, 234, SW - 16, 18, TL)
        self.lb = ListBox()
        self.lb.Font = Font("Consolas", 9.0)
        self.lb.SetBounds(8, 252, SW - 16, H0 - 252 - 330)
        self.lb.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        self.lb.HorizontalScrollbar = True
        self.side.Controls.Add(self.lb)
        self._btn(u"Select all", 8, H0 - 324, 90, 26, lambda s, e: self.select_all(), BL)
        self._btn(u"Clear", 102, H0 - 324, 70, 26, lambda s, e: self.clear_sel(), BL)
        self._btn(u"Undo (Ctrl+Z)", 176, H0 - 324, 120, 26, lambda s, e: self.undo(), BL)
        self._lbl(u"Export times:", 8, H0 - 290, 100, 18, BL)
        self._btn(u"All", 112, H0 - 294, 50, 24, lambda s, e: self.check_times(True), BL)
        self._btn(u"None", 166, H0 - 294, 50, 24, lambda s, e: self.check_times(False), BL)
        self._btn(u"Show", 220, H0 - 294, 70, 24, lambda s, e: self.show_clicked(), BL)
        self.clb = CheckedListBox()
        self.clb.CheckOnClick = True
        self.clb.SetBounds(8, H0 - 264, SW - 16, 92)
        self.clb.Anchor = BL
        self.side.Controls.Add(self.clb)
        self.chkExpT = self._chk(u"Temp", 8, H0 - 166, 70, True, lambda s, e: None, BL)
        self.chkExpD = self._chk(u"Disp (UX, UY)", 82, H0 - 166, 150, True, lambda s, e: None, BL)
        self._lbl(u"Separator:", 8, H0 - 136, 62, 18, BL)
        self.cbSep = self._combo(74, H0 - 140, 90, BL)
        for t in (u"Tab", u"Space", u"Comma"):
            self.cbSep.Items.Add(t)
        self.cbSep.SelectedIndex = 0
        self._lbl(u"Temp unit:", 172, H0 - 136, 66, 18, BL)
        self.cbExpUnit = self._combo(240, H0 - 140, 60, BL)
        self.cbExpUnit.Items.Add(u"C")
        self.cbExpUnit.Items.Add(u"K")
        self.cbExpUnit.SelectedIndex = 0
        self.chkMv = self._chk(u"Add mesh section-3 value column (MeshVal)", 8, H0 - 112, SW - 16, False, lambda s, e: None, BL)
        self._btn(u"Copy to clipboard", 8, H0 - 86, 160, 28, lambda s, e: self.do_copy(), BL)
        self._btn(u"Save TXT...", 176, H0 - 86, 120, 28, lambda s, e: self.do_save(), BL)
        self.lblStatus = self._lbl(u"", 8, H0 - 52, SW - 16, 48, BL)
        self.lblStatus.Font = Font("Segoe UI", 8.0)

    def note(self, text):
        self.lblStatus.Text = text

    def _opt_draw(self):
        self.view.mesh_on = self.chkMesh.Checked
        self.view.ids_on = self.chkIds.Checked
        self.view.mark_dirty()

    def _opt_mode(self):
        self.view.box_mode = self.chkBox.Checked
        self.view.hide_mode = self.chkHide.Checked
        self.view._restore_cursor()
        self.view.Invalidate()

    # ----- yukleme -----
    def ask_files(self):
        sf = StartForm(self.mesh_path, self.temp_path, self.disp_path)
        if sf.ShowDialog(self) == DialogResult.OK:
            self.load_files(sf.mesh_path, sf.temp_path, sf.disp_path)

    def _set_file_label(self, lbl, prefix, path):
        if path:
            lbl.Text = u"%s: %s" % (prefix, System.IO.Path.GetFileName(path))
            self._tt.SetToolTip(lbl, path)
        else:
            lbl.Text = u"%s: -" % prefix
            self._tt.SetToolTip(lbl, u"")

    def load_files(self, mesh_path, temp_path, disp_path):
        pf = ProgressForm(u"Loading ...")
        pf.Show()
        err = None
        tf = df = None
        try:
            pf.update(0.0, u"Reading mesh file ...")
            nodes, raw, mv = parse_mesh(mesh_path)
            pf.update(0.0, u"Building model (%d elements) ..." % len(raw))
            model = build_model(nodes, raw)
            if temp_path:
                tf = FieldFile(temp_path, "temp", pf.update)
            if disp_path:
                df = FieldFile(disp_path, "disp", pf.update)
        except Cancelled:
            err = u""
        except Exception as ex:
            err = u"Dosya okunamadi:\n%s" % (ex,)
        finally:
            pf.Close()
        if err is not None:
            if err:
                MessageBox.Show(err, TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error)
            return False
        for ff, nm in ((tf, u".temp"), (df, u".disp")):
            if ff is not None and not ff.times:
                MessageBox.Show(u"%s dosyasinda TIME bilgisi olan hicbir blok bulunamadi." % nm, TITLE,
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
                return False
        self.mesh_path, self.temp_path, self.disp_path = mesh_path, temp_path, disp_path
        self._undo = []
        self.nodes, self.mv = nodes, mv
        self.tf, self.df = tf, df
        tset = set()
        if tf is not None:
            tset.update(tf.times)
        if df is not None:
            tset.update(df.times)
        self.times = sorted(tset)
        self.grange = {}
        self._shown = False
        self.shown_times = []
        self.cur_map = {}
        self.cur_fld = None
        self.view.set_model(model)
        self._set_file_label(self.lblMesh, u"Mesh", mesh_path)
        self._set_file_label(self.lblTemp, u"Temp", temp_path)
        self._set_file_label(self.lblDisp, u"Disp", disp_path)
        self._updating = True
        self.fields = [{"name": u"(none)", "src": None, "comp": 0}]
        if tf is not None:
            self.fields.append({"name": u"Temp", "src": "temp", "comp": 0})
        if df is not None:
            self.fields.append({"name": u"Disp UX", "src": "disp", "comp": 0})
            self.fields.append({"name": u"Disp UY", "src": "disp", "comp": 1})
            self.fields.append({"name": u"Disp USUM", "src": "disp", "comp": -1})
        self.cbField.Items.Clear()
        for f_ in self.fields:
            self.cbField.Items.Add(f_["name"])
        self.cbField.SelectedIndex = 1 if len(self.fields) > 1 else 0
        self.cbTime.Items.Clear()
        self.clb.Items.Clear()
        for k, t in enumerate(self.times):
            self.clb.Items.Add(u"%.6g" % t, k == 0)
        self.chkExpT.Enabled = tf is not None
        self.chkExpT.Checked = tf is not None
        self.chkExpD.Enabled = df is not None
        self.chkExpD.Checked = df is not None
        self._updating = False
        self.cbUnit.Enabled = tf is not None
        self.lblTinfo.Text = u""
        self.lblScale.Text = u""
        self.lblHead.Text = u"   Node          X            Y           Value"
        msg = u"%d nodes, %d elements (%d skipped) | %d time(s)" % (len(nodes), len(model["elems"]), model["skipped"], len(self.times))
        for ff, nm in ((tf, u"temp"), (df, u"disp")):
            if ff is not None and len(ff.blocks) != len(ff.times):
                msg += u" | WARNING %s: %d blocks, %d unique times (repeated TIME: last used)" % (nm, len(ff.blocks), len(ff.times))
        msg += u" | Press Show to display a field."
        self.note(msg)
        self.selection_changed()
        return True

    # ----- alan (contour) -----
    def _off(self):
        """Temp icin ekran birimi ofseti (dosya degerleri C)."""
        return 273.15 if self.cbUnit.SelectedIndex == 1 else 0.0

    def _field(self):
        i = self.cbField.SelectedIndex
        if i <= 0 or i >= len(self.fields):
            return None
        return self.fields[i]

    def _src(self, fld):
        return self.tf if fld["src"] == "temp" else self.df

    def _checked_times(self):
        return [self.times[k] for k in range(self.clb.Items.Count) if self.clb.GetItemChecked(k)]

    def field_changed(self):
        if self._updating:
            return
        fld = self._field()
        self.cbUnit.Enabled = bool(fld is not None and fld["src"] == "temp") or (fld is None and self.tf is not None)
        if self._shown:
            if fld is None:
                self._clear_field()
            else:
                self._show(fld, self.shown_times or self._checked_times())

    def unit_changed(self):
        fld = self.cur_fld
        if self._shown and fld is not None and fld["src"] == "temp" and self.cur_time is not None:
            self._apply(fld, self.cur_time)

    def fix_changed(self):
        if self._shown and self.cur_fld is not None and self.cur_time is not None:
            self._apply(self.cur_fld, self.cur_time)

    def show_clicked(self):
        fld = self._field()
        if fld is None:
            MessageBox.Show(u"Once Field kutusundan bir alan sec (Temp / Disp UX / UY / USUM).", TITLE,
                            MessageBoxButtons.OK, MessageBoxIcon.Information)
            return
        times = self._checked_times()
        if not times:
            MessageBox.Show(u"Export times listesinde en az bir time isaretle.", TITLE,
                            MessageBoxButtons.OK, MessageBoxIcon.Information)
            return
        self._show(fld, times)

    def _show(self, fld, times):
        src = self._src(fld)
        avail = [t for t in times if t in src.block_of]
        if not avail:
            MessageBox.Show(u"Isaretli time'larin hicbiri %s dosyasinda yok." % (u"temp" if fld["src"] == "temp" else u"disp"),
                            TITLE, MessageBoxButtons.OK, MessageBoxIcon.Information)
            return
        self.shown_times = avail
        self._updating = True
        self.cbTime.Items.Clear()
        for t in avail:
            self.cbTime.Items.Add(u"%.6g" % t)
        self.cbTime.SelectedIndex = 0
        self._updating = False
        self._shown = True
        self._apply(fld, avail[0])

    def _clear_field(self):
        self._shown = False
        self.cur_map = {}
        self.cur_fld = None
        self.shown_times = []
        self._updating = True
        self.cbTime.Items.Clear()
        self._updating = False
        self.lblScale.Text = u""
        self.lblTinfo.Text = u""
        self.lblHead.Text = u"   Node          X            Y           Value"
        self.view.clear_field()
        self.selection_changed()

    def time_changed(self):
        if self._updating or not self._shown:
            return
        i = self.cbTime.SelectedIndex
        fld = self._field()
        if fld is None or i < 0 or i >= len(self.shown_times):
            return
        self._apply(fld, self.shown_times[i])

    def _load_field(self, fld, t):
        src = self._src(fld)
        if src.is_cached(t):
            return src.scalar(t, fld["comp"])
        pf = ProgressForm(u"Loading %s ..." % fld["name"])
        pf.Show()
        tm = None
        err = None
        try:
            pf.update(0.05, u"Reading %s at time %s ..." % (fld["name"], u"%.6g" % t))
            tm = src.scalar(t, fld["comp"])
            pf.update(1.0, u"Done")
        except Exception as ex:
            err = u"Okunamadi:\n%s" % (ex,)
        finally:
            pf.Close()
        if err is not None:
            MessageBox.Show(err, TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error)
            return None
        return tm

    def _ensure_grange(self, fld):
        key = (fld["src"], fld["comp"])
        if key in self.grange:
            return True
        pf = ProgressForm(u"Scanning all times ...")
        pf.Show()
        ok = False
        try:
            self.grange[key] = self._src(fld).global_range(pf.update, fld["comp"])
            ok = True
        except Cancelled:
            pass
        except Exception as ex:
            MessageBox.Show(u"Okunamadi:\n%s" % (ex,), TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error)
        finally:
            pf.Close()
        return ok

    def _apply(self, fld, t):
        tm = self._load_field(fld, t)
        if tm is None:
            return
        src = self._src(fld)
        self.cur_time, self.cur_map, self.cur_fld = t, tm, fld
        is_t = fld["src"] == "temp"
        off = self._off() if is_t else 0.0
        key = (fld["src"], fld["comp"])
        if self.chkFix.Checked and not self._ensure_grange(fld):
            self.chkFix.Checked = False                      # iptal: otomatik olcek (handler tekrar _apply cagirir)
            return
        if self.chkFix.Checked:
            lo, hi = self.grange[key]
        else:
            used = self.view.model["used"]
            vals = [v for n_, v in tm.items() if n_ in used]
            lo, hi = (min(vals), max(vals)) if vals else (0.0, 1.0)
        case, rpm = src.info(t)
        self.lblTinfo.Text = u"Case %s | RPM %s" % ("-" if case is None else case, "-" if rpm is None else "%.6g" % rpm)
        title = (u"Temp [%s]" % (u"K" if off else u"C")) if is_t else fld["name"]
        self.lblScale.Text = u"Scale: %.4g ... %.4g" % (lo + off, hi + off)
        self.lblHead.Text = u"   Node          X            Y      %s" % title
        self.view.set_field(tm, lo, hi, off, title)
        self.selection_changed()

    def selection_changed(self):
        MAXLIST = 2000
        fld = self.cur_fld if self._shown else None
        off = self._off() if (fld is not None and fld["src"] == "temp") else 0.0
        tm = self.cur_map
        sel = sorted(self.view.sel)
        items = []
        for nid in sel[:MAXLIST]:
            X, Y = self.nodes[nid]
            if fld is None:
                vt = u"-"
            else:
                v = tm.get(nid)
                vt = u"NaN" if v is None else (u"%.3f" % (v + off) if off or fld["src"] == "temp" else u"%.5g" % v)
            items.append(u"%6d %12.5f %12.5f %12s" % (nid, X, Y, vt))
        if len(sel) > MAXLIST:
            items.append(u"... +%d more nodes (selected, not listed)" % (len(sel) - MAXLIST))
        self.lb.BeginUpdate()
        try:
            self.lb.Items.Clear()
            self.lb.Items.AddRange(System.Array[object](items))
        finally:
            self.lb.EndUpdate()
        self.view.Invalidate()

    # ----- secim -----
    def push_undo(self):
        self._undo.append(set(self.view.sel))
        if len(self._undo) > 50:
            del self._undo[0]

    def undo(self):
        if self._undo:
            self.view.sel = self._undo.pop()
            self.selection_changed()
            self.note(u"Undo: %d node(s) selected" % len(self.view.sel))

    def _on_key(self, sender, e):
        if e.Control and e.KeyCode == Keys.Z:
            self.undo()
            e.Handled = True

    def select_all(self):
        self.push_undo()
        hid = self.view.hid_nodes
        self.view.sel = set(n for n in self.view.dn.keys() if n not in hid)
        self.selection_changed()

    def clear_sel(self):
        if self.view.sel:
            self.push_undo()
        self.view.sel = set()
        self.selection_changed()

    def check_times(self, state):
        for k in range(self.clb.Items.Count):
            self.clb.SetItemChecked(k, state)

    # ----- cikti -----
    def build_text(self):
        sel = sorted(self.view.sel)
        if not sel:
            MessageBox.Show(u"Once en az bir node sec.", TITLE, MessageBoxButtons.OK, MessageBoxIcon.Information)
            return None
        sep = [u"\t", u"  ", u","][max(0, self.cbSep.SelectedIndex)]
        times = self._checked_times()
        if not times:
            MessageBox.Show(u"Export times listesinde en az bir time isaretle.", TITLE,
                            MessageBoxButtons.OK, MessageBoxIcon.Information)
            return None
        use_t = self.tf is not None and self.chkExpT.Checked
        use_d = self.df is not None and self.chkExpD.Checked
        if not (use_t or use_d):
            MessageBox.Show(u"Export icin Temp ve / veya Disp kutusunu isaretle.", TITLE,
                            MessageBoxButtons.OK, MessageBoxIcon.Information)
            return None
        ncol = len(times) * ((1 if use_t else 0) + (2 if use_d else 0))
        if len(sel) * ncol > 20000000:
            if MessageBox.Show(u"%d node x %d sutun = %d deger. Bu cok buyuk, bellek yetmeyebilir. Devam edilsin mi?"
                               % (len(sel), ncol, len(sel) * ncol), TITLE,
                               MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes:
                return None
        use_mv = self.chkMv.Checked
        off = 273.15 if self.cbExpUnit.SelectedIndex == 1 else 0.0
        none = [None] * len(sel)
        ct, cd = [], []
        pf = ProgressForm(u"Exporting ...")
        pf.Show()
        err = None
        try:
            for ti, t in enumerate(times):
                if not pf.update(float(ti) / len(times), u"Reading time %s   (%d / %d)" % (u"%.6g" % t, ti + 1, len(times))):
                    raise Cancelled()
                if use_t:
                    ct.append(self.tf.block_nodes(t, sel, [0])[0] if t in self.tf.block_of else none)
                if use_d:
                    cd.append(self.df.block_nodes(t, sel, [0, 1]) if t in self.df.block_of else [none, none])
        except Cancelled:
            err = u""
        except Exception as ex:
            err = u"Okunamadi:\n%s" % (ex,)
        finally:
            pf.Close()
        if err is not None:
            if err:
                MessageBox.Show(err, TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error)
            return None
        head = [u"NodeID", u"X_Loc", u"Y_Loc"]
        if use_t:
            head += [u"Time%d_T" % (k + 1) for k in range(len(times))]
        if use_d:
            for k in range(len(times)):
                head += [u"Time%d_UX" % (k + 1), u"Time%d_UY" % (k + 1)]
        if use_mv:
            head.append(u"MeshVal")
        out = [sep.join(head)]
        for j, nid in enumerate(sel):
            X, Y = self.nodes[nid]
            row = [u"%d" % nid, u"%.10g" % X, u"%.10g" % Y]
            for c in ct:
                v = c[j]
                row.append(u"%.10g" % (v + off) if v is not None else u"NaN")
            for c in cd:
                for comp in (0, 1):
                    v = c[comp][j]
                    row.append(u"%.10g" % v if v is not None else u"NaN")
            if use_mv:
                mvv = self.mv.get(nid)
                row.append(u"%.9g" % mvv if mvv is not None else u"NaN")
            out.append(sep.join(row))
        return u"\r\n".join(out) + u"\r\n"

    def do_copy(self):
        txt = self.build_text()
        if txt is not None:
            Clipboard.SetText(txt)
            self.note(u"Copied: %d node(s)" % len(self.view.sel))

    def do_save(self):
        txt = self.build_text()
        if txt is None:
            return
        d = SaveFileDialog()
        d.Title = u"Save TXT"
        d.Filter = u"Text files (*.txt)|*.txt|All files (*.*)|*.*"
        d.FileName = u"node_values.txt"
        if d.ShowDialog() == DialogResult.OK:
            try:
                System.IO.File.WriteAllText(d.FileName, txt)
                self.note(u"Saved: %s" % d.FileName)
            except Exception as ex:
                MessageBox.Show(u"Kaydedilemedi:\n%s" % (ex,), TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error)


def main():
    owner = mechanical_owner()
    f = ViewerForm()
    mp = tp = dp = u""
    while True:
        sf = StartForm(mp, tp, dp)
        r = sf.ShowDialog(owner) if owner is not None else sf.ShowDialog()
        if r != DialogResult.OK:
            return
        mp, tp, dp = sf.mesh_path, sf.temp_path, sf.disp_path
        if f.load_files(mp, tp, dp):
            break
    if owner is not None:
        f.ShowDialog(owner)
    else:
        f.ShowDialog()


main()
