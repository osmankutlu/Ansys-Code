# -*- coding: utf-8 -*-
# =====================================================================
#  Mesh Temp Viewer 8   (Converter)   (Ansys Mechanical, IronPython 2.7)
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
#  v8: koordinat ekseni (sol alt: X radial, Y eksenel); Unselect modu; export penceresinde sadece var olan dosyaya ait secenekler aktif;
#      sag alttaki durum yazisi kaldirildi
#  v7: Export ayri pencere (agac -> liste, temp ve disp ayri dosya, UX/UY/USUM secimi, mm/m); Probe;
#      yatay renk cubugu (ekranin altinda); secili node listesi ListView (basliklar hizali); Fixed scale kaldirildi
#  v6: yeni sag panel duzeni (gruplar alt alta, secili node listesi en altta); tek yerden Field + Time secimi;
#      Undo (secim + gizleme) ve Unhide all yan yana; export basliklari T_<time> / D_<time>
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
                                  CheckedListBox, TextBox, FormBorderStyle, Application, Timer, ProgressBar, ToolTip, GroupBox,
                                  ListView, ListViewItem, View, TreeView, TreeNode, TreeViewAction, SelectionMode)
from System.Drawing import (Size, Point, Font, FontStyle, Color, PointF, Pen, SolidBrush, Bitmap, Graphics,
                            RectangleF, GraphicsUnit)
from System.Drawing.Drawing2D import SmoothingMode, DashStyle, InterpolationMode

TITLE = u"Mesh Temp Viewer 8"
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

    def node_values(self, t, nid):
        """Tek node'un tum sutunlari (tuple) - probe icin; blok onbellekteyse anlik."""
        ids, cols = self.block_cols(t)
        i = self.pos_of.get(nid) if self.pos_of is not None else None
        if i is None or i >= len(ids) or ids[i] != nid:
            try:
                i = ids.index(nid)
            except ValueError:
                return None
        return tuple(c[i] for c in cols)

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
        self.probe_on = False
        self.unsel_mode = False
        self.probes = []
        self.probe_text = {}
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
        self.f_leg_title = Font("Segoe UI", 15.0, FontStyle.Bold)
        self.f_leg = Font("Segoe UI", 12.0, FontStyle.Bold)
        self.f_probe = Font("Segoe UI", 9.5, FontStyle.Bold)
        self.f_axis = Font("Segoe UI", 10.0, FontStyle.Bold)
        self.b_leg_bg = SolidBrush(Color.FromArgb(225, 255, 255, 255))
        self.b_probe_bg = SolidBrush(Color.FromArgb(240, 255, 255, 215))
        self.p_probe = Pen(Color.FromArgb(200, 0, 160), 2.0)
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
        self.probes = []
        self.probe_text = {}
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
        self.owner.push_hide()
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
            self.owner.hides_cleared()
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
        self.Cursor = Cursors.Hand if self.hide_mode else (Cursors.Cross if (self.box_mode or self.probe_on or self.unsel_mode) else Cursors.Default)

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
        self.s = min((self.Width - 40.0) / (bw * 1.1), (self.Height - 150.0) / (bh * 1.1))
        self.px = 0.0
        self.py = -45.0                                # model biraz yukari: altta renk cubugu icin yer
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
            self._draw_probes(g)
        except Exception as ex:
            g.DrawString(u"Draw error: %s" % str(ex), self.f_ui, SolidBrush(Color.Firebrick), 8.0, 24.0)
        g.SmoothingMode = SmoothingMode.AntiAlias
        if self.unsel_mode:
            txt = u"UNSELECT: click a selected node, or drag a box (right-drag / Shift+drag) to remove nodes from the selection | selected: %d" % len(self.sel)
        elif self.hide_mode:
            txt = u"HIDE ELEMENTS: click an element to hide it | right-drag / Shift+drag = hide elements in the box | hidden: %d" % self.n_hidden
        else:
            txt = (u"Selected nodes: %d    |  wheel: zoom  |  drag: pan  |  click: select node  |  right-drag / Shift+drag / Box mode: box select  |  Ctrl+Z: undo"
                   % len(self.sel))
        g.DrawString(txt, self.f_ui, SolidBrush(Color.DarkOrange if (self.hide_mode or self.unsel_mode) else Color.DimGray), 8.0, 4.0)
        if self.has_t:
            self._draw_legend(g, self.f_ui)
        if self._rb is not None:
            x0, y0, x1, y1 = self._rb
            pen = Pen(Color.Firebrick if self.unsel_mode else Color.DarkOrange, 1.0)
            pen.DashStyle = DashStyle.Dash
            g.DrawRectangle(pen, min(x0, x1), min(y0, y1), abs(x1 - x0), abs(y1 - y0))
        self._draw_axes(g)

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

    def _draw_probes(self, g):
        if not self.probes:
            return
        s = self.s
        ox, oy = self._ofs()
        W, H = float(self.Width), float(self.Height)
        g.SmoothingMode = SmoothingMode.AntiAlias
        for nid in self.probes:
            p = self.dn.get(nid)
            if p is None or nid in self.hid_nodes:
                continue
            sx, sy = p[0] * s + ox, p[1] * s + oy
            g.DrawEllipse(self.p_probe, sx - 7.0, sy - 7.0, 14.0, 14.0)
            lines = self.probe_text.get(nid) or [u"Node %d" % nid]
            w, h = 0.0, 0.0
            hs = []
            for ln in lines:
                sz = g.MeasureString(ln, self.f_probe)
                w = max(w, sz.Width)
                h += sz.Height
                hs.append(sz.Height)
            bx = min(max(4.0, sx + 12.0), max(4.0, W - w - 14.0))
            by = min(max(4.0, sy - h - 10.0), max(4.0, H - h - 10.0))
            g.FillRectangle(self.b_probe_bg, bx, by, w + 8.0, h + 4.0)
            g.DrawRectangle(self.p_probe, bx, by, w + 8.0, h + 4.0)
            yy = by + 2.0
            for ln, hh in zip(lines, hs):
                g.DrawString(ln, self.f_probe, self.b_dark, bx + 4.0, yy)
                yy += hh

    def _draw_axes(self, g):
        """Sol altta koordinat ekseni: X (radial) saga, Y (eksenel) yukari."""
        H = float(self.Height)
        ox, oy, L = 36.0, H - 36.0, 66.0
        g.SmoothingMode = SmoothingMode.AntiAlias
        g.FillRectangle(self.b_leg_bg, 6.0, H - 140.0, 176.0, 134.0)
        pen = Pen(Color.Black, 2.5)
        g.DrawLine(pen, ox, oy, ox + L, oy)
        g.DrawLine(pen, ox, oy, ox, oy - L)
        blk = SolidBrush(Color.Black)
        g.FillPolygon(blk, System.Array[PointF]([PointF(ox + L + 8.0, oy), PointF(ox + L - 3.0, oy - 5.0), PointF(ox + L - 3.0, oy + 5.0)]))
        g.FillPolygon(blk, System.Array[PointF]([PointF(ox, oy - L - 8.0), PointF(ox - 5.0, oy - L + 3.0), PointF(ox + 5.0, oy - L + 3.0)]))
        g.DrawString(u"X (radial)", self.f_axis, blk, ox + L + 12.0, oy - 10.0)
        g.DrawString(u"Y (axial)", self.f_axis, blk, ox - 8.0, oy - L - 30.0)

    def _draw_legend(self, g, f):
        """Yatay renk cubugu ekranin altinda; ustte ne gosterildigi (Temp / Disp ..) buyuk yazilir."""
        W, H = float(self.Width), float(self.Height)
        avail = max(120.0, W - 190.0 - 40.0)
        bw = min(avail - 48.0, 780.0)
        x0 = 190.0 + 24.0 + (avail - 48.0 - bw) / 2.0
        bh = 20.0
        yb = H - 50.0
        g.FillRectangle(self.b_leg_bg, x0 - 24.0, yb - 42.0, bw + 48.0, 92.0)
        g.DrawString(self.title, self.f_leg_title, SolidBrush(Color.Black), x0 - 10.0, yb - 38.0)
        n = 128
        seg = bw / n
        for k in range(n):
            ci = int(255 * k / float(n - 1))
            g.FillRectangle(self.brushes[ci], x0 + k * seg, yb, seg + 1.0, bh)
        g.DrawRectangle(Pen(Color.Gray, 1.0), x0, yb, bw, bh)
        br = SolidBrush(Color.Black)
        tick = Pen(Color.Black, 1.0)
        for frac in (0.0, 0.25, 0.5, 0.75, 1.0):
            val = self.vmin + (self.vmax - self.vmin) * frac + self.off
            xx = x0 + bw * frac
            g.DrawLine(tick, xx, yb + bh, xx, yb + bh + 5.0)
            txt = u"%.5g" % val
            sz = g.MeasureString(txt, self.f_leg)
            g.DrawString(txt, self.f_leg, br, xx - sz.Width / 2.0, yb + bh + 5.0)

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
            elif self.unsel_mode:
                hit = set()
                for nid in self.sel:
                    p = self.dn.get(nid)
                    if p is None:
                        continue
                    sx, sy = p[0] * s + ox, p[1] * s + oy
                    if xa <= sx <= xb and ya <= sy <= yb:
                        hit.add(nid)
                if hit:
                    self.owner.push_undo()
                    self.sel -= hit
                    self.owner.selection_changed()
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
            elif self.probe_on:
                nid = self._find_node(x0, y0)
                if nid is not None:
                    self.owner.probe_toggle(nid)
            elif self.unsel_mode:
                nid = self._find_node(x0, y0, self.sel)
                if nid is not None:
                    self.owner.push_undo()
                    self.sel.discard(nid)
                    self.owner.selection_changed()
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

    def _find_node(self, mx, my, only=None):
        """Imlece en yakin (gorunen) node id'si (12 px icinde) veya None."""
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
            if nid in hid or (only is not None and nid not in only):
                continue
            x, y = self.dn[nid]
            d = math.hypot(x * s + ox - mx, y * s + oy - my)
            if d < best:
                best, bn = d, nid
        return bn

    def _pick(self, mx, my):
        bn = self._find_node(mx, my)
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


class ExportForm(Form):
    """Export penceresi: solda secilebilecekler (agac: Temp / Disp -> time'lar), ortada ekle / cikar butonlari,
    sagda export edilecekler. Temp ve Disp AYRI dosyalara yazilir."""
    def __init__(self, owner):
        self.owner = owner
        self.Text = u"Export TXT"
        self.ClientSize = Size(760, 536)
        self.StartPosition = FormStartPosition.CenterParent
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        self._busy = False
        self.chosen = set()                      # {(src, time)}
        self._right = []
        self.lblSel = Label()
        self.lblSel.AutoSize = False
        self.lblSel.SetBounds(12, 10, 736, 18)
        self.Controls.Add(self.lblSel)
        gl = GroupBox()
        gl.Text = u"Available  (check items, then press  >  )"
        gl.SetBounds(10, 34, 330, 336)
        self.Controls.Add(gl)
        self.tree = TreeView()
        self.tree.CheckBoxes = True
        self.tree.SetBounds(8, 20, 314, 306)
        self.tree.AfterCheck += self._after_check
        gl.Controls.Add(self.tree)
        for text, y, h in ((u">", 100, self._add_checked), (u">>", 136, self._add_all),
                           (u"<", 196, self._remove_sel), (u"<<", 232, self._remove_all)):
            b = Button()
            b.Text = text
            b.SetBounds(348, y, 64, 30)
            b.Click += (lambda hh: (lambda s_, e_: hh()))(h)
            self.Controls.Add(b)
        gr = GroupBox()
        gr.Text = u"To export  (select and press  <  to remove)"
        gr.SetBounds(420, 34, 330, 336)
        self.Controls.Add(gr)
        self.lbR = ListBox()
        self.lbR.SelectionMode = SelectionMode.MultiExtended
        self.lbR.SetBounds(8, 20, 314, 306)
        gr.Controls.Add(self.lbR)
        go = GroupBox()
        go.Text = u"Options"
        go.SetBounds(10, 376, 740, 100)
        self.Controls.Add(go)
        has_t = owner.tf is not None
        has_d = owner.df is not None
        dl1 = self._mk_lbl(go, u"Disp columns:", 10, 28, 90)
        self.chkUX = self._mk_chk(go, u"UX", 104, 24, 52, True)
        self.chkUY = self._mk_chk(go, u"UY", 160, 24, 52, True)
        self.chkUS = self._mk_chk(go, u"USUM", 216, 24, 70, False)
        dl2 = self._mk_lbl(go, u"Disp unit:", 320, 28, 64)
        self.cbDU = self._mk_combo(go, 388, 24, 70, [u"mm", u"m"])
        tl1 = self._mk_lbl(go, u"Temp unit:", 490, 28, 70)
        self.cbTU = self._mk_combo(go, 564, 24, 70, [u"C", u"K"])
        for c in (dl1, self.chkUX, self.chkUY, self.chkUS, dl2, self.cbDU):       # Disp secenekleri: sadece disp dosyasi varsa
            c.Enabled = has_d
        for c in (tl1, self.cbTU):                                                # Temp secenekleri: sadece temp dosyasi varsa
            c.Enabled = has_t
        self._mk_lbl(go, u"Separator:", 10, 66, 70)
        self.cbSep = self._mk_combo(go, 84, 62, 90, [u"Tab", u"Space", u"Comma"])
        self.chkMv = self._mk_chk(go, u"Add MeshVal column (mesh section 3)", 210, 62, 300, False)
        self.lblStatus = Label()
        self.lblStatus.AutoSize = False
        self.lblStatus.SetBounds(12, 486, 530, 42)
        self.Controls.Add(self.lblStatus)
        be = Button()
        be.Text = u"Export..."
        be.SetBounds(560, 488, 110, 32)
        be.Click += lambda s_, e_: self._do_export()
        self.Controls.Add(be)
        bc = Button()
        bc.Text = u"Close"
        bc.SetBounds(680, 488, 70, 32)
        bc.DialogResult = DialogResult.Cancel
        self.Controls.Add(bc)
        self.CancelButton = bc
        self._refresh_label()
        self._fill_tree()

    def _mk_lbl(self, parent, text, x, y, w):
        l = Label()
        l.AutoSize = False
        l.Text = text
        l.SetBounds(x, y, w, 18)
        parent.Controls.Add(l)
        return l

    def _mk_chk(self, parent, text, x, y, w, checked):
        c = CheckBox()
        c.Text = text
        c.SetBounds(x, y, w, 22)
        c.Checked = checked
        parent.Controls.Add(c)
        return c

    def _mk_combo(self, parent, x, y, w, items):
        c = ComboBox()
        c.DropDownStyle = ComboBoxStyle.DropDownList
        c.SetBounds(x, y, w, 24)
        for it in items:
            c.Items.Add(it)
        c.SelectedIndex = 0
        parent.Controls.Add(c)
        return c

    def _refresh_label(self):
        self.lblSel.Text = u"Selected nodes to export: %d" % len(self.owner.view.sel)

    def _sources(self):
        out = []
        if self.owner.tf is not None:
            out.append(("temp", u"Temp", self.owner.tf.times))
        if self.owner.df is not None:
            out.append(("disp", u"Disp", self.owner.df.times))
        return out

    def _fill_tree(self):
        self.tree.BeginUpdate()
        try:
            self.tree.Nodes.Clear()
            for src, label, times in self._sources():
                done = set(t for (s_, t) in self.chosen if s_ == src)
                root = TreeNode(label)
                root.Tag = src
                for t in times:
                    if t in done:
                        continue
                    ch = TreeNode(u"t = %.6g s" % t)
                    ch.Tag = t
                    root.Nodes.Add(ch)
                root.Text = u"%s   (%d available)" % (label, root.Nodes.Count)
                self.tree.Nodes.Add(root)
                root.Expand()
        finally:
            self.tree.EndUpdate()

    def _after_check(self, sender, e):
        if self._busy or e.Action == TreeViewAction.Unknown:
            return
        self._busy = True
        try:
            for ch in e.Node.Nodes:
                ch.Checked = e.Node.Checked
        finally:
            self._busy = False

    def _refresh_right(self):
        self._right = sorted(self.chosen, key=lambda c: (0 if c[0] == "temp" else 1, c[1]))
        self.lbR.BeginUpdate()
        try:
            self.lbR.Items.Clear()
            for src, t in self._right:
                self.lbR.Items.Add(u"%s     t = %.6g s" % (u"Temp" if src == "temp" else u"Disp", t))
        finally:
            self.lbR.EndUpdate()
        self._fill_tree()

    def _add_checked(self):
        for root in self.tree.Nodes:
            for ch in root.Nodes:
                if ch.Checked:
                    self.chosen.add((root.Tag, ch.Tag))
        self._refresh_right()

    def _add_all(self):
        for src, label, times in self._sources():
            for t in times:
                self.chosen.add((src, t))
        self._refresh_right()

    def _remove_sel(self):
        for i in sorted([i for i in self.lbR.SelectedIndices], reverse=True):
            if 0 <= i < len(self._right):
                self.chosen.discard(self._right[i])
        self._refresh_right()

    def _remove_all(self):
        self.chosen = set()
        self._refresh_right()

    def _do_export(self):
        comps = []
        if self.chkUX.Checked:
            comps.append(u"UX")
        if self.chkUY.Checked:
            comps.append(u"UY")
        if self.chkUS.Checked:
            comps.append(u"USUM")
        opts = {"temp_times": sorted(t for (s_, t) in self.chosen if s_ == "temp"),
                "disp_times": sorted(t for (s_, t) in self.chosen if s_ == "disp"),
                "comps": comps,
                "disp_unit": u"m" if self.cbDU.SelectedIndex == 1 else u"mm",
                "temp_unit": u"K" if self.cbTU.SelectedIndex == 1 else u"C",
                "sep": [u"\t", u"  ", u","][max(0, self.cbSep.SelectedIndex)],
                "mv": self.chkMv.Checked}
        self.lblStatus.Text = self.owner.run_export(opts)
        self._refresh_label()


class ViewerForm(Form):
    def __init__(self):
        self.Text = TITLE
        wa = Screen.PrimaryScreen.WorkingArea
        self.ClientSize = Size(int(0.85 * min(1600, wa.Width - 30)), int(min(900, wa.Height - 60)))
        self.MinimumSize = Size(860, 640)
        self.StartPosition = FormStartPosition.CenterScreen
        self.nodes, self.mv = {}, {}
        self.tf = self.df = None
        self.times = []
        self.fields = []
        self.disp_times = []                  # Time kutusundaki time'lar (secili alanin dosyasindakiler)
        self._shown = False
        self._updating = False
        self.cur_time = None
        self.cur_map = {}
        self.cur_fld = None
        self._undo = []
        self._last_dir = None
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

    # ----- kontrol uretici yardimcilar (parent = side panel veya bir GroupBox) -----
    def _lbl(self, parent, text, x, y, w, h, anchor=None, ellipsis=False):
        l = Label()
        l.AutoSize = False
        l.AutoEllipsis = ellipsis
        l.Text = text
        l.SetBounds(x, y, w, h)
        if anchor is not None:
            l.Anchor = anchor
        parent.Controls.Add(l)
        return l

    def _btn(self, parent, text, x, y, w, h, handler, anchor=None):
        b = Button()
        b.Text = text
        b.SetBounds(x, y, w, h)
        b.Click += handler
        if anchor is not None:
            b.Anchor = anchor
        parent.Controls.Add(b)
        return b

    def _chk(self, parent, text, x, y, w, checked, handler):
        c = CheckBox()
        c.Text = text
        c.SetBounds(x, y, w, 22)
        c.Checked = checked
        c.CheckedChanged += handler
        parent.Controls.Add(c)
        return c

    def _combo(self, parent, x, y, w):
        c = ComboBox()
        c.DropDownStyle = ComboBoxStyle.DropDownList
        c.SetBounds(x, y, w, 24)
        parent.Controls.Add(c)
        return c

    def _grp(self, title, y, h, anchor=None):
        g = GroupBox()
        g.Text = title
        g.SetBounds(6, y, SW - 12, h)
        if anchor is not None:
            g.Anchor = anchor
        self.side.Controls.Add(g)
        return g

    # ----- sag panel: yukaridan asagiya  Files / Display / Select-Hide-Probe / Export / Selected nodes / status -----
    def _build_side(self):
        H0 = self.side.Height
        sp = self.side
        TL = AnchorStyles.Top | AnchorStyles.Left
        TLR = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        # --- dosyalar ---
        self.lblMesh = self._lbl(sp, u"Mesh: -", 8, 4, SW - 16, 16, TLR, True)
        self.lblTemp = self._lbl(sp, u"Temp: -", 8, 21, SW - 16, 16, TLR, True)
        self.lblDisp = self._lbl(sp, u"Disp: -", 8, 38, SW - 16, 16, TLR, True)
        self._btn(sp, u"Open files...", 8, 58, 120, 26, lambda s, e: self.ask_files(), TL)
        self._btn(sp, u"Fit view", 136, 58, 90, 26, lambda s, e: self.view.fit_all(), TL)
        # --- ekranda goster ---
        g = self._grp(u"Display on screen", 90, 104, TLR)
        self._lbl(g, u"Field:", 8, 22, 40, 18)
        self.cbField = self._combo(g, 50, 18, 150)
        self.cbField.SelectedIndexChanged += lambda s, e: self.field_changed()
        self._lbl(g, u"Unit:", 208, 22, 34, 18)
        self.cbUnit = self._combo(g, 244, 18, 64)
        self.cbUnit.Items.Add(u"C")
        self.cbUnit.Items.Add(u"K")
        self.cbUnit.SelectedIndex = 0
        self.cbUnit.SelectedIndexChanged += lambda s, e: self.unit_changed()
        self._lbl(g, u"Time [s]:", 8, 52, 58, 18)
        self.cbTime = self._combo(g, 68, 48, 104)
        self.cbTime.Enabled = False
        self.cbTime.SelectedIndexChanged += lambda s, e: self.time_changed()
        self.btnPrev = self._btn(g, u"<", 176, 47, 26, 26, lambda s, e: self.step_time(-1))
        self.btnNext = self._btn(g, u">", 204, 47, 26, 26, lambda s, e: self.step_time(1))
        self.btnPrev.Enabled = False
        self.btnNext.Enabled = False
        self.lblTinfo = self._lbl(g, u"", 236, 52, 100, 18)
        self.chkMesh = self._chk(g, u"Mesh lines", 8, 76, 104, True, lambda s, e: self._opt_draw())
        self.chkIds = self._chk(g, u"Node IDs", 116, 76, 100, False, lambda s, e: self._opt_draw())
        # --- secim / gizleme / probe ---
        g2 = self._grp(u"Select / Hide / Probe", 198, 110, TLR)
        self.chkBox = self._chk(g2, u"Box select mode", 8, 16, 146, False, lambda s, e: self._opt_mode())
        self.chkHide = self._chk(g2, u"Hide elements mode", 158, 16, 176, False, lambda s, e: self._opt_mode())
        self._btn(g2, u"Select all", 8, 42, 80, 26, lambda s, e: self.select_all())
        self._btn(g2, u"Clear", 92, 42, 64, 26, lambda s, e: self.clear_sel())
        self._btn(g2, u"Undo", 160, 42, 84, 26, lambda s, e: self.undo())
        self._btn(g2, u"Unhide all", 248, 42, 88, 26, lambda s, e: self.view.unhide_all())
        self.btnProbe = self._btn(g2, u"Probe: OFF", 8, 76, 100, 26, lambda s, e: self.toggle_probe())
        self.btnUnsel = self._btn(g2, u"Unselect: OFF", 112, 76, 110, 26, lambda s, e: self.toggle_unsel())
        tip = self._lbl(g2, u"Right-drag / Shift+drag = box", 228, 78, 108, 24)
        tip.Font = Font("Segoe UI", 7.5)
        # --- export ---
        self._btn(sp, u"Export TXT...", 8, 316, 170, 30, lambda s, e: self.open_export(), TL)
        # --- secili node'lar (en altta, kalan alani doldurur) ---
        gy = 354
        self.grpSel = self._grp(u"Selected nodes", gy, H0 - gy - 8,
                                AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right)
        self.lv = ListView()
        self.lv.View = View.Details
        self.lv.FullRowSelect = True
        self.lv.GridLines = True
        self.lv.HideSelection = False
        self.lv.Font = Font("Consolas", 9.0)
        self.lv.Columns.Add(u"Node", 64)
        self.lv.Columns.Add(u"X", 90)
        self.lv.Columns.Add(u"Y", 90)
        self.lv.Columns.Add(u"Value", 74)
        self.lv.SetBounds(8, 18, SW - 12 - 16, H0 - gy - 8 - 26)
        self.lv.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        self.grpSel.Controls.Add(self.lv)

    def note(self, text):
        pass                                  # sag alttaki durum yazisi kaldirildi

    def _opt_draw(self):
        self.view.mesh_on = self.chkMesh.Checked
        self.view.ids_on = self.chkIds.Checked
        self.view.mark_dirty()

    def _opt_mode(self):
        if self.chkHide.Checked:
            if self.view.probe_on:
                self._set_probe(False)
            if self.view.unsel_mode:
                self._set_unsel(False)
        self.view.box_mode = self.chkBox.Checked
        self.view.hide_mode = self.chkHide.Checked
        self.view._restore_cursor()
        self.view.Invalidate()

    # ----- probe -----
    def toggle_probe(self):
        self._set_probe(not self.view.probe_on)

    def _set_probe(self, on):
        v = self.view
        v.probe_on = on
        self.btnProbe.Text = u"Probe: ON" if on else u"Probe: OFF"
        self.btnProbe.BackColor = Color.FromArgb(255, 205, 90) if on else Control.DefaultBackColor
        if on and self.chkHide.Checked:
            self.chkHide.Checked = False
        if on and v.unsel_mode:
            self._set_unsel(False)
        if not on:
            v.probes = []
            v.probe_text = {}
        v._restore_cursor()
        v.Invalidate()

    def toggle_unsel(self):
        self._set_unsel(not self.view.unsel_mode)

    def _set_unsel(self, on):
        v = self.view
        v.unsel_mode = on
        self.btnUnsel.Text = u"Unselect: ON" if on else u"Unselect: OFF"
        self.btnUnsel.BackColor = Color.FromArgb(255, 150, 150) if on else Control.DefaultBackColor
        if on:
            if self.chkHide.Checked:
                self.chkHide.Checked = False
            if v.probe_on:
                self._set_probe(False)
        v._restore_cursor()
        v.Invalidate()

    def probe_toggle(self, nid):
        v = self.view
        if nid in v.probes:
            v.probes.remove(nid)
        else:
            v.probes.append(nid)
        self.refresh_probes()

    def _probe_lines(self, nid):
        lines = [u"Node %d" % nid]
        fld = self.cur_fld if self._shown else None
        if fld is None:
            return lines
        try:
            vals = self._src(fld).node_values(self.cur_time, nid)
        except Exception:
            vals = None
        if vals is None:
            lines.append(u"(no data)")
        elif fld["src"] == "temp":
            off = self._off()
            lines.append(u"T = %.3f %s" % (vals[0] + off, u"K" if off else u"C"))
        else:
            ux, uy, uz = vals[0], vals[1], vals[2]
            lines.append(u"UX %.5g     UY %.5g     USUM %.5g   [mm]" % (ux, uy, math.sqrt(ux * ux + uy * uy + uz * uz)))
        return lines

    def refresh_probes(self):
        v = self.view
        v.probe_text = dict((nid, self._probe_lines(nid)) for nid in v.probes)
        v.Invalidate()

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
        self._shown = False
        self.disp_times = []
        self.cur_map = {}
        self.cur_fld = None
        self.cur_time = None
        self.view.set_model(model)
        self._set_probe(False)
        self._set_unsel(False)
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
        self.cbField.SelectedIndex = 0
        self.cbTime.Items.Clear()
        self.cbTime.Enabled = False
        self.btnPrev.Enabled = False
        self.btnNext.Enabled = False
        self._updating = False
        self.cbUnit.Enabled = tf is not None
        self.lblTinfo.Text = u""
        self.lv.Columns[3].Text = u"Value"
        msg = u"%d nodes, %d elements (%d skipped) | %d time(s)" % (len(nodes), len(model["elems"]), model["skipped"], len(self.times))
        for ff, nm in ((tf, u"temp"), (df, u"disp")):
            if ff is not None and len(ff.blocks) != len(ff.times):
                msg += u" | WARNING %s: %d blocks, %d unique times (repeated TIME: last used)" % (nm, len(ff.blocks), len(ff.times))
        msg += u" | Pick a Field to display it."
        self.note(msg)
        self.selection_changed()
        return True

    # ----- ekranda alan gosterme (Field + Time) -----
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

    def field_changed(self):
        if self._updating:
            return
        fld = self._field()
        if fld is None:
            self.cbUnit.Enabled = self.tf is not None
            self._clear_field()
            return
        self.cbUnit.Enabled = fld["src"] == "temp"
        src = self._src(fld)
        keep = self.cur_time
        self.disp_times = list(src.times)
        idx = self.disp_times.index(keep) if keep in self.disp_times else 0
        self._updating = True
        self.cbTime.Items.Clear()
        for t in self.disp_times:
            self.cbTime.Items.Add(u"%.6g" % t)
        self.cbTime.SelectedIndex = idx
        self._updating = False
        self.cbTime.Enabled = True
        self.btnPrev.Enabled = True
        self.btnNext.Enabled = True
        self._apply(fld, self.disp_times[idx])

    def time_changed(self):
        if self._updating or (self.tf is None and self.df is None):
            return
        fld = self._field()
        i = self.cbTime.SelectedIndex
        if fld is None or i < 0 or i >= len(self.disp_times):
            return
        self._apply(fld, self.disp_times[i])

    def step_time(self, d):
        i = self.cbTime.SelectedIndex + d
        if 0 <= i < len(self.disp_times):
            self.cbTime.SelectedIndex = i             # time_changed() gosterir

    def unit_changed(self):
        fld = self.cur_fld
        if self._shown and fld is not None and fld["src"] == "temp" and self.cur_time is not None:
            self._apply(fld, self.cur_time)

    def _clear_field(self):
        self._shown = False
        self.cur_map = {}
        self.cur_fld = None
        self.disp_times = []
        self._updating = True
        self.cbTime.Items.Clear()
        self._updating = False
        self.cbTime.Enabled = False
        self.btnPrev.Enabled = False
        self.btnNext.Enabled = False
        self.lblTinfo.Text = u""
        self.lv.Columns[3].Text = u"Value"
        self.view.clear_field()
        self.refresh_probes()
        self.selection_changed()

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

    def _apply(self, fld, t):
        tm = self._load_field(fld, t)
        if tm is None:
            return
        src = self._src(fld)
        self.cur_time, self.cur_map, self.cur_fld = t, tm, fld
        self._shown = True
        is_t = fld["src"] == "temp"
        off = self._off() if is_t else 0.0
        used = self.view.model["used"]
        vals = [v for n_, v in tm.items() if n_ in used]
        lo, hi = (min(vals), max(vals)) if vals else (0.0, 1.0)
        case, rpm = src.info(t)
        self.lblTinfo.Text = u"Case %s | RPM %s" % ("-" if case is None else case, "-" if rpm is None else "%.6g" % rpm)
        name = (u"Temp [%s]" % (u"K" if off else u"C")) if is_t else (u"%s [mm]" % fld["name"])
        self.lv.Columns[3].Text = name
        self.view.set_field(tm, lo, hi, off, u"%s      Time = %.6g s" % (name, t))
        self.refresh_probes()
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
                vt = u"NaN" if v is None else (u"%.3f" % (v + off) if fld["src"] == "temp" else u"%.5g" % v)
            items.append(ListViewItem(System.Array[str]([u"%d" % nid, u"%.5f" % X, u"%.5f" % Y, vt])))
        if len(sel) > MAXLIST:
            items.append(ListViewItem(System.Array[str]([u"...", u"+%d more" % (len(sel) - MAXLIST), u"", u""])))
        self.lv.BeginUpdate()
        try:
            self.lv.Items.Clear()
            self.lv.Items.AddRange(System.Array[ListViewItem](items))
        finally:
            self.lv.EndUpdate()
        self.grpSel.Text = u"Selected nodes (%d)" % len(sel)
        self.view.Invalidate()

    # ----- secim / geri alma (secim + gizleme tek Undo'da) -----
    def push_undo(self):
        self._undo.append(("sel", set(self.view.sel)))
        if len(self._undo) > 100:
            del self._undo[0]

    def push_hide(self):
        self._undo.append(("hide", None))
        if len(self._undo) > 100:
            del self._undo[0]

    def hides_cleared(self):
        self._undo = [e for e in self._undo if e[0] != "hide"]

    def undo(self):
        while self._undo:
            kind, data = self._undo.pop()
            if kind == "sel":
                self.view.sel = data
                self.selection_changed()
                self.note(u"Undo: %d node(s) selected" % len(self.view.sel))
                return
            if self.view.hide_stack:
                self.view.undo_hide()
                return

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

    # ----- export -----
    def open_export(self):
        if self.tf is None and self.df is None:
            return
        ExportForm(self).ShowDialog(self)

    def _ask_save(self, title, default):
        d = SaveFileDialog()
        d.Title = title
        d.Filter = u"Text files (*.txt)|*.txt|All files (*.*)|*.*"
        d.FileName = default
        if self._last_dir:
            d.InitialDirectory = self._last_dir
        if d.ShowDialog() == DialogResult.OK:
            self._last_dir = System.IO.Path.GetDirectoryName(d.FileName)
            return d.FileName
        return None

    def _read_times(self, ff, times, sel, comps, title):
        """-> (veri, hata). veri[k] = [her comp icin node listesi] (times sirasinda). Iptalde hata = u""."""
        pf = ProgressForm(u"Exporting %s ..." % title)
        pf.Show()
        out = []
        err = None
        try:
            for ti, t in enumerate(times):
                if not pf.update(float(ti) / len(times), u"%s: reading time %s   (%d / %d)" % (title, u"%.6g" % t, ti + 1, len(times))):
                    raise Cancelled()
                out.append(ff.block_nodes(t, sel, comps))
        except Cancelled:
            err = u""
        except Exception as ex:
            err = u"Okunamadi:\n%s" % (ex,)
        finally:
            pf.Close()
        return (out, err)

    def _big_ok(self, nvalues):
        if nvalues > 20000000:
            return MessageBox.Show(u"%d deger yazilacak. Bu cok buyuk, bellek yetmeyebilir. Devam edilsin mi?" % nvalues, TITLE,
                                   MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes
        return True

    def run_export(self, o):
        sel = sorted(self.view.sel)
        if not sel:
            MessageBox.Show(u"Once ana pencerede en az bir node sec.", TITLE, MessageBoxButtons.OK, MessageBoxIcon.Information)
            return u"No nodes selected."
        want_t = bool(o["temp_times"]) and self.tf is not None
        want_d = bool(o["disp_times"]) and self.df is not None
        if not (want_t or want_d):
            MessageBox.Show(u"Sag listeye en az bir Temp veya Disp time'i ekle.", TITLE, MessageBoxButtons.OK, MessageBoxIcon.Information)
            return u"Nothing to export."
        if want_d and not o["comps"]:
            MessageBox.Show(u"Disp icin en az bir sutun sec (UX / UY / USUM).", TITLE, MessageBoxButtons.OK, MessageBoxIcon.Information)
            return u"No Disp column selected."
        results = []
        if want_t:
            r = self._export_temp(sel, o)
            if r:
                results.append(r)
        if want_d:
            r = self._export_disp(sel, o)
            if r:
                results.append(r)
        return u"   |   ".join(results) if results else u"Nothing saved."

    def _export_temp(self, sel, o):
        times = o["temp_times"]
        if not self._big_ok(len(sel) * len(times)):
            return u""
        path = self._ask_save(u"Save Temp TXT", u"node_temp.txt")
        if not path:
            return u""
        data, err = self._read_times(self.tf, times, sel, [0], u"Temp")
        if err is not None:
            if err:
                MessageBox.Show(err, TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error)
            return u"Temp export cancelled."
        off = 273.15 if o["temp_unit"] == u"K" else 0.0
        sep = o["sep"]
        head = [u"NodeID", u"X_Loc", u"Y_Loc"] + [u"T_%.6g" % t for t in times]
        if o["mv"]:
            head.append(u"MeshVal")
        out = [sep.join(head)]
        for j, nid in enumerate(sel):
            X, Y = self.nodes[nid]
            row = [u"%d" % nid, u"%.10g" % X, u"%.10g" % Y]
            for d in data:
                v = d[0][j]
                row.append(u"%.10g" % (v + off) if v is not None else u"NaN")
            if o["mv"]:
                m = self.mv.get(nid)
                row.append(u"%.9g" % m if m is not None else u"NaN")
            out.append(sep.join(row))
        return self._write(path, out, u"Temp", len(sel), len(times))

    def _export_disp(self, sel, o):
        times = o["disp_times"]
        comps = o["comps"]
        if not self._big_ok(len(sel) * len(times) * len(comps)):
            return u""
        path = self._ask_save(u"Save Disp TXT", u"node_disp.txt")
        if not path:
            return u""
        need = [0, 1, 2] if u"USUM" in comps else sorted(set(c for c, nm in ((0, u"UX"), (1, u"UY")) if nm in comps))
        data, err = self._read_times(self.df, times, sel, need, u"Disp")
        if err is not None:
            if err:
                MessageBox.Show(err, TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error)
            return u"Disp export cancelled."
        f = 0.001 if o["disp_unit"] == u"m" else 1.0            # dosya mm kabul edilir
        sep = o["sep"]
        names = {u"UX": u"DX", u"UY": u"DY", u"USUM": u"DSUM"}
        head = [u"NodeID", u"X_Loc", u"Y_Loc"]
        for t in times:
            if len(comps) == 1:
                head.append(u"D_%.6g" % t)
            else:
                head += [u"%s_%.6g" % (names[c], t) for c in comps]
        if o["mv"]:
            head.append(u"MeshVal")
        out = [sep.join(head)]
        for j, nid in enumerate(sel):
            X, Y = self.nodes[nid]
            row = [u"%d" % nid, u"%.10g" % X, u"%.10g" % Y]
            for d in data:
                for c in comps:
                    if c == u"UX":
                        v = d[need.index(0)][j]
                    elif c == u"UY":
                        v = d[need.index(1)][j]
                    else:
                        a_, b_, c_ = d[0][j], d[1][j], d[2][j]
                        v = math.sqrt(a_ * a_ + b_ * b_ + c_ * c_) if (a_ is not None and b_ is not None and c_ is not None) else None
                    row.append(u"%.10g" % (v * f) if v is not None else u"NaN")
            if o["mv"]:
                m = self.mv.get(nid)
                row.append(u"%.9g" % m if m is not None else u"NaN")
            out.append(sep.join(row))
        return self._write(path, out, u"Disp", len(sel), len(times))

    def _write(self, path, lines, label, nn, nt):
        try:
            System.IO.File.WriteAllText(path, u"\r\n".join(lines) + u"\r\n")
        except Exception as ex:
            MessageBox.Show(u"Kaydedilemedi:\n%s" % (ex,), TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error)
            return u"%s: could not save." % label
        self.note(u"Saved %s: %s" % (label, path))
        return u"%s saved (%d nodes x %d times): %s" % (label, nn, nt, System.IO.Path.GetFileName(path))


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
