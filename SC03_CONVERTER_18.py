# -*- coding: utf-8 -*-
# =====================================================================
#  SC03 CONVERTER   (v18)   (Ansys Mechanical, IronPython 2.7)
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
#  v18: cizim hatti ESKI haline (v15: mip genel gorunum + gorunum onbellegi, arka planda cizim, yeni resim hazir olana kadar eski resim
#       olceklenip gosterilir) geri alindi; tile piramidi kaldirildi. Arayuz / secim / Fit / bos yere tiklama ayni.
#  v17: Fit = ekranda kalan (gizlenmemis) elemanlari sigdirir; zoom sirasinda gizlenmis / olmayan yerlerin eski tile'lari artik gosterilmez
#       (gizleme sonrasi eski tile'lar gecersiz, cok kaba tile buyutulmez); bos yere tiklama gecici (sari) secimi iptal eder, 'Cancel picks' kalkti
#  v16: TILE PIRAMIDI cizim: model bir kez 512 px'lik parcalar (tile) halinde, yakinlik seviyelerine gore cizilir ve onbellekte tutulur;
#       zoom / pan sirasinda hazir tile'lar gosterilir, sadece eksik tile'lar arka planda cizilir (bastan cizim yok);
#       listeye eklenen / Hide others'ta kullanilan node'larin ekran isareti kalkar ('Show listed' ile acilabilir); 'Cancel picks' (eski Discard picks)
#  v15: secim iki asamali (node'a tikla = SARI gecici isaret, 'Add to list' = listeye ekle + YESIL isaret); renk cubugunda +/- ve Auto kalkti,
#       her yeni time / gizleme sonrasi renk skalasi otomatik; 'Rendering' yazisi kalkti; Hide others / Show all hizlandirildi
#       (renkler arka planda cizim sirasinda hesaplanir, toplu gizleme); sag panel genisligi surukleyerek ayarlanir
#  v14: ONERI 3 yerlesimi (koyu tema, sol arac seridi, ustte Field/Unit/Time, sagda secili node tablosu); Show all; secili node ID yazisi yok;
#       Unit: temp C/K, disp mm/m (son secim hatirlanir); sadece SAG TUS ile kutu secim; bilgi yazilari kalkti; eksen takimi arka karesiz;
#       ekranda gorunen node'larda MAX / MIN etiketi (Max/Min dugmesiyle acilir-kapanir); renk cubugunda sabit Min / Max + renk skalasi +/- ve Auto
#  v13: Hide others (secili node'lara dokunmayan elemanlari gizler); renk olcegi gorunen (gizlenmemis) node'lara gore
#       her gizleme degisiminde yenilenir; renk cubugu etiketleri kendi penceresine sigdirildi; baslik SC03 CONVERTER
#  v12: Mesh lines / Nodes dugmeleri varsayilan KAPALI (Node IDs kaldirildi: Nodes = node noktalari); donmus gorunumde de hizli mip;
#       renk cubugu bir kez cizilip onbellekte tutulur; hareket sirasinda buyuk secim isaretleri atlanir
#  v11: model gorunumu ESKI haline dondu (X yatay, Y dikey yukari); sadece eksen takimi X yukari (kirmizi) / Y saga (yesil);
#       Hide elements / Mesh lines / Node IDs dugme oldu, Box select kutusu kalkti, Rotate (Z ekseni, 90 derece) eklendi
#  v10: pencereler MODELESS (Ansys kullanilabilir kalir); agir cizimler arka plan thread'inde + animasyonlu zoom;
#       eksen takimi: X yukari (kirmizi), Y saga (yesil); gorunum X yukari / Y saga
#  v9: Export: tek time listesi (Shift / Ctrl ile coklu secim), Temp / Disp kutulari (varsayilan bos), kayit adi bir kez
#      sorulur (_temp / _disp eklenir), bitince 'Export completed' mesaji
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
import threading

from System.Windows.Forms import (Form, Button, Label, MessageBox, MessageBoxButtons, MessageBoxIcon,
                                  FormStartPosition, DialogResult, Panel, Cursors, Control, Keys,
                                  MouseButtons, Clipboard, ComboBox, ComboBoxStyle, CheckBox, ListBox,
                                  OpenFileDialog, SaveFileDialog, AnchorStyles, Screen, IWin32Window,
                                  CheckedListBox, TextBox, FormBorderStyle, Application, Timer, ProgressBar, ToolTip, GroupBox,
                                  ListView, ListViewItem, View, TreeView, TreeNode, TreeViewAction, SelectionMode,
                                  FlatStyle, BorderStyle)
from System.Drawing import (Size, Point, Font, FontStyle, Color, PointF, Pen, SolidBrush, Bitmap, Graphics,
                            RectangleF, GraphicsUnit, StringFormat, StringAlignment)
from System.Drawing.Drawing2D import SmoothingMode, DashStyle, InterpolationMode

TITLE = u"SC03 CONVERTER"
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
        return (X, -Y)                            # cizim koordinati: yatay = model X, dikey = model Y (yukari)

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
    el["cw"] = shp(1.0 / 3.0, 1.0 / 3.0) if kind == "tri" else shp(0.0, 0.0)      # eleman merkezindeki sekil fonksiyonlari
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
BULK_HIDE = 3000               # bu kadar elemandan fazlasi gizlenirken / gosterilirken toplu (hizli) yol kullanilir


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
    """Cizim:
    - MIP: tum modelin sabit cozunurlukte (en fazla 4096 px) cizilmis hali. Uzaklastirilmis gorunumde (buyuk modellerde)
      zoom / pan sadece bu resmi olcekler / kaydirir.
    - CACHE: yakinlastirilmis gorunumde ekranda gorunen elemanlarin cizimi.
    Bu resimlerin ASIL cizimi ARKA PLAN thread'inde yapilir: zoom / pan sirasinda arayuz hic beklemez, eski resim
    olceklenip kaydirilarak gosterilir, yeni net resim hazir olunca yerine gecer. Zoom animasyonludur.
    Secim / probe isaretleri her zaman ustten cizilir."""
    def __init__(self, owner):
        Panel.__init__(self)
        self.owner = owner
        self.DoubleBuffered = True
        self.bg = Color.FromArgb(22, 24, 28)
        self.BackColor = self.bg
        self.model = None
        self.dn = {}                                  # node -> cizim koordinati
        self._ngrid = None
        self._egrid = None
        self._grid_thread = None
        self._cell = 1.0
        self.avg_el = 1.0
        self.sel = set()
        self.s = 1.0
        self.px = self.py = self.cx = self.cy = 0.0
        self.rot = 0.0                                # gorunum donusu (radyan, Z ekseni etrafinda, saat yonu)
        self._rot_target = None
        self.mesh_on = False                          # varsayilan kapali (daha hizli)
        self.nodes_on = False                         # node noktalarini goster / sakla (varsayilan kapali)
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
        # --- arka plan cizim durumu ---
        self._content_gen = 0                         # renk / gizleme / mesh cizgisi degisince artar
        self._view_gen = 0                            # zoom / pan / boyut degisince artar
        self._moving = False                          # gorunum su an degisiyor (hareket durana kadar cizim baslamaz)
        self._cache = None
        self._cache_cgen = -1
        self._cache_sig = None
        self._cs = 1.0
        self._cxf = None
        self._mip = None
        self._mip_cgen = -1
        self._mip_ok = False
        self._mip_ms = 1.0                            # mip olcegi [piksel / model birimi]
        self._mip_rot = 0.0                           # mip'in cizildigi donus
        self._mip_xf = None
        self._mip_w = self._mip_h = 0
        self._leg_bmp = None                          # renk cubugu onbellegi
        self._leg_key = None
        self.fac = 1.0                                # goruntulenen deger = ham deger * fac + off
        self.dext = None                              # (min, min node, max, max node): gorunen node'larin veri uclari
        self.minmax_on = True                         # ekrandaki node'larda MAX / MIN etiketi
        self._mm = None                               # ekrandaki node'larin (min node, min, max node, max) sonucu
        self._mm_sig = None
        self._mm_job = None
        self.pend = set()                             # tiklanmis ama listeye EKLENMEMIS (gecici, sari) node'lar
        self.listed_on = False                        # listedeki node'lari ekranda yesil goster (varsayilan kapali)
        self.vis_nodes = None                         # gizleme varken gorunen node'lar (onbellek)
        self._job = None
        self._result = None
        self._anim = None
        self._timer = Timer()
        self._timer.Interval = 150
        self._timer.Tick += self._on_timer
        self._poll = Timer()
        self._poll.Interval = 50
        self._poll.Tick += self._on_poll
        self._poll.Start()
        self._anim_timer = Timer()
        self._anim_timer.Interval = 15
        self._anim_timer.Tick += self._on_anim
        # --- cizim nesneleri (bir kez uretilir, her paint'te yeniden yaratilmaz) ---
        self.brushes = [SolidBrush(jet(i / 255.0)) for i in range(256)] + [SolidBrush(Color.LightGray)]
        self.f_ui = Font("Segoe UI", 8.5)
        self.f_lab = Font("Segoe UI", 7.5)
        self.b_dark = SolidBrush(Color.FromArgb(40, 40, 40))
        self.b_green = SolidBrush(Color.FromArgb(0, 200, 110))
        self.p_glow = Pen(Color.FromArgb(0, 120, 60), 1.5)
        self.f_leg_title = Font("Segoe UI", 15.0, FontStyle.Bold)
        self.f_leg = Font("Segoe UI", 12.0, FontStyle.Bold)
        self.f_probe = Font("Segoe UI", 9.5, FontStyle.Bold)
        self.f_axis = Font("Segoe UI", 15.0, FontStyle.Bold)
        self.b_leg_bg = SolidBrush(Color.FromArgb(225, 255, 255, 255))
        self.b_probe_bg = SolidBrush(Color.FromArgb(240, 255, 255, 215))
        self.p_probe = Pen(Color.FromArgb(200, 0, 160), 2.0)
        self.b_black = SolidBrush(Color.Black)
        self.b_dim = SolidBrush(Color.FromArgb(170, 176, 186))
        self.b_orange = SolidBrush(Color.DarkOrange)
        self.b_red = SolidBrush(Color.Firebrick)
        self.p_gray = Pen(Color.Gray, 1.0)
        self.p_black = Pen(Color.Black, 1.0)
        self.p_axx = Pen(Color.FromArgb(255, 70, 70), 3.0)
        self.p_axy = Pen(Color.FromArgb(70, 220, 100), 3.0)
        self.b_axx = SolidBrush(Color.FromArgb(255, 70, 70))
        self.b_axy = SolidBrush(Color.FromArgb(70, 220, 100))
        self.b_white = SolidBrush(Color.White)
        self.b_cmin = SolidBrush(Color.FromArgb(30, 70, 210))
        self.b_cmax = SolidBrush(Color.FromArgb(200, 30, 30))
        self.b_mmax = SolidBrush(Color.FromArgb(215, 45, 45))
        self.b_mmin = SolidBrush(Color.FromArgb(40, 95, 220))
        self.p_mmax = Pen(Color.FromArgb(255, 120, 120), 3.0)
        self.p_mmin = Pen(Color.FromArgb(120, 170, 255), 3.0)
        self.b_btn = SolidBrush(Color.FromArgb(58, 64, 76))
        self.p_btn = Pen(Color.FromArgb(130, 138, 150), 1.0)
        self.f_btn = Font("Segoe UI", 13.0, FontStyle.Bold)
        self.b_pend = SolidBrush(Color.FromArgb(255, 205, 40))
        self.p_pend = Pen(Color.FromArgb(150, 100, 0), 1.5)
        self.sf_c = StringFormat()
        self.sf_c.Alignment = StringAlignment.Center
        self.sf_c.LineAlignment = StringAlignment.Center
        self.p_rb_o = Pen(Color.DarkOrange, 1.0)
        self.p_rb_o.DashStyle = DashStyle.Dash
        self.p_rb_r = Pen(Color.Firebrick, 1.0)
        self.p_rb_r.DashStyle = DashStyle.Dash
        self.Paint += self._paint
        self.MouseDown += self._mdown
        self.MouseMove += self._mmove
        self.MouseUp += self._mup
        self.MouseWheel += self._wheel
        self.MouseEnter += lambda s, e: self.Focus()
        self.Resize += self._on_resize

    # ----- model / renkler -----
    def set_model(self, model):
        self._content_gen += 1                        # calisan cizimleri iptal eder
        self._view_gen += 1
        self.model = model
        self.dn = dict((nid, (xy[0], -xy[1])) for nid, xy in model["nodes"].items() if nid in model["used"])
        els = model["elems"]
        sample = els[:2000]
        self.avg_el = max(sum(e["bb"][1] - e["bb"][0] for e in sample) / float(len(sample)), 1e-12)
        self._cell = self.avg_el
        self._ngrid = self._egrid = None
        th = threading.Thread(target=self._build_grids, args=(model, self.dn, self._cell))   # tiklama gridleri arka planda
        th.daemon = True
        self._grid_thread = th
        th.start()
        self.sel = set()
        self.pend = set()
        self.vis_nodes = None
        self.hide_stack = []
        self.hid_nodes = set()
        self.node_cnt = None
        self.n_hidden = 0
        self.probes = []
        self.probe_text = {}
        self.has_t = False
        self._mm = None
        self._mm_sig = None
        x0, x1, y0, y1 = model["box"]
        bw, bh = max(x1 - x0, 1e-12), max(y1 - y0, 1e-12)
        self._mip_ms = MIP_MAX / max(bw, bh)
        self._mip_ok = len(els) > MIP_MIN_ELEMS
        for b in (self._cache, self._mip):
            if b is not None:
                b.Dispose()
        self._cache = self._mip = None
        self._cache_cgen = self._mip_cgen = -1
        self._cache_sig = None
        self._result = None
        self._anim = None
        self._fit_pending = True
        self.Invalidate()

    def set_field(self, tmap, vmin, vmax, off, title, fac=1.0, ext=None):
        self.tmap, self.vmin, self.vmax, self.off, self.title = tmap, vmin, vmax, off, title
        self.fac = fac
        self.dext = ext
        self.has_t = True
        self.mark_dirty()

    def clear_field(self):
        self.has_t = False
        self.tmap = {}
        self.dext = None
        self._mm = None
        self._mm_sig = None
        self.mark_dirty()

    def mark_dirty(self):
        self._content_gen += 1
        self.Invalidate()

    # ----- eleman gizleme -----
    def _ensure_node_cnt(self):
        if self.node_cnt is None:
            cnt = {}
            for el in self.model["elems"]:
                if not el["hid"]:
                    for n in el["nodes"]:
                        cnt[n] = cnt.get(n, 0) + 1
            self.node_cnt = cnt

    def _recompute_hid_nodes(self):
        """Toplu gizleme / gosterme: gorunen node kumesi bir turda yeniden kurulur (eleman eleman guncellemeden cok hizli)."""
        vis = set()
        for el in self.model["elems"]:
            if not el["hid"]:
                vis.update(el["nodes"])
        self.vis_nodes = vis
        self.hid_nodes = self.model["used"] - vis
        self.node_cnt = None

    def visible_nodes(self):
        """Gizlenmemis elemanlara ait node'lar."""
        if self.n_hidden == 0:
            return self.model["used"]
        if self.vis_nodes is None:
            vis = set()
            for el in self.model["elems"]:
                if not el["hid"]:
                    vis.update(el["nodes"])
            self.vis_nodes = vis
        return self.vis_nodes

    def _hide_elems(self, els):
        new = [e for e in els if not e["hid"]]
        if not new:
            return
        if len(new) <= BULK_HIDE:
            self._ensure_node_cnt()                      # sayaclar gizlemeden ONCE (gorunen elemanlara gore) kurulur
            cnt = self.node_cnt
            for e in new:
                e["hid"] = True
                for n in e["nodes"]:
                    c = cnt.get(n, 0) - 1
                    cnt[n] = c
                    if c <= 0:
                        self.hid_nodes.add(n)
            self.vis_nodes = None
        else:
            for e in new:
                e["hid"] = True
            self._recompute_hid_nodes()
        self.hide_stack.append(new)
        self.owner.push_hide()
        self.n_hidden += len(new)
        self.mark_dirty()
        self.owner.hidden_changed("hide")

    def _unhide(self, els):
        els = [e for e in els if e["hid"]]
        if not els:
            return
        if len(els) <= BULK_HIDE:
            self._ensure_node_cnt()
            cnt = self.node_cnt
            for e in els:
                e["hid"] = False
                self.n_hidden -= 1
                for n in e["nodes"]:
                    cnt[n] = cnt.get(n, 0) + 1
                    self.hid_nodes.discard(n)
            self.vis_nodes = None
        else:
            for e in els:
                e["hid"] = False
            self.n_hidden -= len(els)
            self._recompute_hid_nodes()

    def hide_others(self, keep):
        """Secili node'lara dokunmayan (node'larindan hicbiri secili olmayan) tum elemanlari gizler. -> gizlenen sayisi"""
        els = [el for el in self.model["elems"] if not el["hid"] and keep.isdisjoint(el["nodes"])]
        self._hide_elems(els)
        return len(els)

    def undo_hide(self):
        if self.hide_stack:
            g = self.hide_stack.pop()
            self._unhide(g)
            self.mark_dirty()
            self.owner.hidden_changed("unhide")

    def unhide_all(self):
        """Hepsini goster: tek turda bayraklar sifirlanir (eleman eleman guncelleme yok)."""
        if self.hide_stack or self.n_hidden:
            for el in self.model["elems"]:
                el["hid"] = False
            self.hide_stack = []
            self.n_hidden = 0
            self.hid_nodes = set()
            self.vis_nodes = None
            self.node_cnt = None
            self.owner.hides_cleared()
            self.mark_dirty()
            self.owner.hidden_changed("unhide")

    # ----- gridler (ilk kullanimda kurulur) -----
    def _build_grids(self, model, dn, c):
        """Node ve eleman icin tiklama gridleri (arka plan thread'inde kurulur)."""
        try:
            ng = {}
            for nid, (x, y) in dn.items():
                ng.setdefault((int(math.floor(x / c)), int(math.floor(y / c))), []).append(nid)
            eg = {}
            for k, el in enumerate(model["elems"]):
                x0, x1, y0, y1 = el["bb"]
                for a in range(int(math.floor(x0 / c)), int(math.floor(x1 / c)) + 1):
                    for b in range(int(math.floor(y0 / c)), int(math.floor(y1 / c)) + 1):
                        eg.setdefault((a, b), []).append(k)
            if self.model is model:
                self._ngrid, self._egrid = ng, eg
        except Exception:
            pass

    def _ensure_grids(self):
        if self._ngrid is not None:
            return
        self.Cursor = Cursors.WaitCursor
        try:
            th = self._grid_thread
            if th is not None and th.is_alive():
                th.join()
            if self._ngrid is None:
                self._build_grids(self.model, self.dn, self._cell)
        finally:
            self._restore_cursor()

    def _restore_cursor(self):
        self.Cursor = Cursors.Hand if self.hide_mode else (Cursors.Cross if (self.box_mode or self.probe_on or self.unsel_mode) else Cursors.Default)

    # ----- gorunum degisimi -----
    def _view_changed(self):
        self._view_gen += 1
        self._moving = True
        self._timer.Stop()
        self._timer.Start()
        self.Invalidate()

    def _on_timer(self, sender, e):                   # hareket durdu (150 ms) -> net cizim baslayabilir
        self._timer.Stop()
        self._moving = False
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

    def _mat(self):
        c, sn = math.cos(self.rot), math.sin(self.rot)
        s = self.s
        return (s * c, -s * sn, s * sn, s * c)

    def _xf(self):
        """Taban (model cizim) koordinatindan ekrana: ekran = M * taban + o  -> (m00, m01, m10, m11, ox, oy)."""
        m00, m01, m10, m11 = self._mat()
        ox = self.Width / 2.0 + self.px - (m00 * self.cx + m01 * self.cy)
        oy = self.Height / 2.0 + self.py - (m10 * self.cx + m11 * self.cy)
        return (m00, m01, m10, m11, ox, oy)

    def _to_base(self, x, y, xf):
        m00, m01, m10, m11, ox, oy = xf
        det = m00 * m11 - m01 * m10
        dx, dy = x - ox, y - oy
        return ((m11 * dx - m01 * dy) / det, (-m10 * dx + m00 * dy) / det)

    def _visible_box(self):
        """Gizlenmemis elemanlarin siniri (Fit bunu sigdirir); gizleme yoksa tum model."""
        if self.n_hidden == 0:
            return self.model["box"]
        x0 = x1 = y0 = y1 = None
        for el in self.model["elems"]:
            if el["hid"]:
                continue
            a0, a1, b0, b1 = el["bb"]
            if x0 is None:
                x0, x1, y0, y1 = a0, a1, b0, b1
            else:
                if a0 < x0:
                    x0 = a0
                if a1 > x1:
                    x1 = a1
                if b0 < y0:
                    y0 = b0
                if b1 > y1:
                    y1 = b1
        return (x0, x1, y0, y1) if x0 is not None else self.model["box"]

    def fit_all(self):
        if self.model is None:
            return
        x0, x1, y0, y1 = self._visible_box()
        self.cx, self.cy = (x0 + x1) / 2.0, (y0 + y1) / 2.0
        c, sn = math.cos(self.rot), math.sin(self.rot)
        pts = [(x0, y0), (x1, y0), (x1, y1), (x0, y1)]
        xs = [c * (x - self.cx) - sn * (y - self.cy) for x, y in pts]
        ys = [sn * (x - self.cx) + c * (y - self.cy) for x, y in pts]
        bw, bh = max(xs) - min(xs), max(ys) - min(ys)
        big = max(bw, bh, 1e-12)
        bw, bh = max(bw, 0.25 * big), max(bh, 0.25 * big)
        self.s = min((self.Width - 40.0) / (bw * 1.1), (self.Height - 170.0) / (bh * 1.1))
        self.px = 0.0
        self.py = -55.0                                # model biraz yukari: altta renk cubugu icin yer
        self._anim = None
        self._view_changed()

    def rotate(self, ddeg):
        """Modeli Z ekseni etrafinda (ekran duzleminde) doner; ekran merkezindeki nokta merkezde kalir."""
        if self.model is None:
            return
        if self._rot_target is None:
            self.cx, self.cy = self._to_base(self.Width / 2.0, self.Height / 2.0, self._xf())
            self.px = self.py = 0.0
            self._rot_target = self.rot
        self._rot_target += math.radians(ddeg)
        self._anim_timer.Start()

    def zoom(self, f, mx, my):
        """Animasyonlu zoom: hedef olcek birikir, imlecin altindaki nokta sabit kalir."""
        bx, by = self._to_base(mx, my, self._xf())
        base = self._anim["target"] if self._anim is not None else self.s
        self._anim = {"mx": float(mx), "my": float(my), "Xa": bx, "Ya": by, "target": max(1e-9, base * f)}
        self._anim_timer.Start()

    def _on_anim(self, sender, e):
        a = self._anim
        rt = self._rot_target
        if a is None and rt is None:
            self._anim_timer.Stop()
            return
        if rt is not None:
            d = rt - self.rot
            if abs(d) < 0.01:
                self.rot = rt % (2.0 * math.pi)
                self._rot_target = None
            else:
                self.rot += d * 0.35
        if a is not None:
            cur, tgt = self.s, a["target"]
            lr = math.log(tgt / cur)
            if abs(lr) < 0.004:
                new, done = tgt, True
            else:
                new, done = cur * math.exp(lr * 0.4), False
            self.s = new
            m00, m01, m10, m11 = self._mat()
            dx, dy = a["Xa"] - self.cx, a["Ya"] - self.cy
            self.px = a["mx"] - self.Width / 2.0 - (m00 * dx + m01 * dy)
            self.py = a["my"] - self.Height / 2.0 - (m10 * dx + m11 * dy)
            if done:
                self._anim = None
        self._view_changed()
        if self._anim is None and self._rot_target is None:
            self._anim_timer.Stop()

    # ----- cizim: agir isler ARKA PLAN thread'inde -----
    # MIP: tum modelin sabit cozunurlukte (en fazla 4096 px) cizilmis hali (uzaklastirilmis gorunum).
    # CACHE: yakinlastirilmis gorunumde ekranda gorunen elemanlarin cizimi.
    # Yeni resim hazirlanirken eski resim olceklenip kaydirilarak gosterilir; hazir olunca yerine gecer.
    def _use_mip(self):
        return self._mip_ok and self.s <= self._mip_ms * 1.02

    def _mip_geom_for(self, rot):
        """Mip bitmap'inin boyutu ve donusumu; donus mip'e GOMULUR, boylece zoom / pan hizli (eksene paralel) cizilir."""
        x0, x1, y0, y1 = self.model["box"]
        cx, cy = (x0 + x1) / 2.0, (y0 + y1) / 2.0
        c, sn = math.cos(rot), math.sin(rot)
        ms = self._mip_ms
        pts = [(x0, y0), (x1, y0), (x1, y1), (x0, y1)]
        xs = [c * (x - cx) - sn * (y - cy) for x, y in pts]
        ys = [sn * (x - cx) + c * (y - cy) for x, y in pts]
        mw = int((max(xs) - min(xs)) * ms) + 3
        mh = int((max(ys) - min(ys)) * ms) + 3
        ox = 1.0 - ms * (c * cx - sn * cy) - ms * min(xs)
        oy = 1.0 - ms * (sn * cx + c * cy) - ms * min(ys)
        return mw, mh, (ms * c, -ms * sn, ms * sn, ms * c, ox, oy)

    def _sig(self):
        return (self.s, self.rot, self.px, self.py, self.cx, self.cy, self.Width, self.Height)

    def _need_render(self):
        if self.model is None or self.Width < 10 or self.Height < 10 or self._moving:
            return None
        if self._use_mip():
            if self._mip is None or self._mip_cgen != self._content_gen or abs(self._mip_rot - self.rot) > 1e-9:
                return "mip"
            return None
        if self._cache is None or self._cache_cgen != self._content_gen or self._cache_sig != self._sig():
            return "cache"
        return None

    def _start_job(self, kind):
        if kind == "mip":
            mw, mh, mxf = self._mip_geom_for(self.rot)
            job = {"kind": "mip", "s": self._mip_ms, "xf": mxf, "W": mw, "H": mh, "rot": self.rot}
        else:
            job = {"kind": "cache", "s": self.s, "xf": self._xf(), "W": self.Width, "H": self.Height, "sig": self._sig()}
        job["cgen"] = self._content_gen
        job["vgen"] = self._view_gen
        job["alive"] = True
        job["has_t"], job["tmap"] = self.has_t, self.tmap                  # alan verisi: renkler cizim thread'inde hesaplanir
        job["vmin"], job["vmax"] = self.vmin, self.vmax
        self._job = job
        th = threading.Thread(target=self._run_job, args=(job,))
        th.daemon = True
        th.start()

    def _run_job(self, job):
        bmp = None
        try:
            bmp = Bitmap(job["W"], job["H"])
            g = Graphics.FromImage(bmp)
            try:
                g.Clear(self.bg)
                ok = self._draw_model(g, job["xf"], float(job["W"]), float(job["H"]), job["kind"] == "mip", job)
            finally:
                g.Dispose()
            if ok:
                self._result = (job, bmp)
                bmp = None
        except Exception:
            pass
        finally:
            if bmp is not None:
                bmp.Dispose()
            job["alive"] = False

    def _on_poll(self, sender, e):
        try:
            r = self._result
            if r is not None:
                self._result = None
                job, bmp = r
                if job["cgen"] != self._content_gen or self.model is None:
                    bmp.Dispose()
                elif job["kind"] == "mip":
                    old = self._mip
                    self._mip, self._mip_cgen = bmp, job["cgen"]
                    self._mip_rot, self._mip_xf, self._mip_w, self._mip_h = job["rot"], job["xf"], job["W"], job["H"]
                    if old is not None:
                        old.Dispose()
                    self.Invalidate()
                else:
                    old = self._cache
                    self._cache, self._cache_cgen, self._cache_sig = bmp, job["cgen"], job["sig"]
                    self._cs, self._cxf = job["s"], job["xf"]
                    if old is not None:
                        old.Dispose()
                    self.Invalidate()
            if self._job is not None and not self._job["alive"]:
                self._job = None
                self.Invalidate()
            if self._job is None:
                k = self._need_render()
                if k is not None:
                    self._start_job(k)
                    self.Invalidate()
            if self.minmax_on and self.has_t and not self._moving and self.model is not None:       # ekrandaki node'larin max / min'i
                now = (self._content_gen, self._sig())
                j = self._mm_job
                if j is not None and not j["alive"]:
                    self._mm_job = None
                    if j["res"] is not None and j["sig"] == now:
                        self._mm, self._mm_sig = j["res"], j["sig"]
                        self.Invalidate()
                    j = None
                if j is None and self._mm_sig != now:
                    job = {"sig": now, "xf": self._xf(), "W": float(self.Width), "H": float(self.Height),
                           "alive": True, "res": None, "cgen": self._content_gen}
                    self._mm_job = job
                    th = threading.Thread(target=self._run_mm, args=(job, self.tmap, self.dn, self.hid_nodes))
                    th.daemon = True
                    th.start()
        except Exception:
            pass

    def _run_mm(self, job, tmap, dn, hid):
        """ARKA PLAN: gorunum penceresindeki (gizlenmemis) node'larin min / max degeri ve node'lari."""
        try:
            m00, m01, m10, m11, ox, oy = job["xf"]
            W, H = job["W"], job["H"]
            lo = hi = None
            nlo = nhi = None
            n = 0
            get = tmap.get
            for nid, (x, y) in dn.items():
                n += 1
                if (n & 8191) == 0 and self._content_gen != job["cgen"]:
                    return
                v = get(nid)
                if v is None:
                    continue
                sx = m00 * x + m01 * y + ox
                if sx < 0.0 or sx > W:
                    continue
                sy = m10 * x + m11 * y + oy
                if sy < 0.0 or sy > H or nid in hid:
                    continue
                if lo is None or v < lo:
                    lo, nlo = v, nid
                if hi is None or v > hi:
                    hi, nhi = v, nid
            if lo is not None:
                job["res"] = (nlo, lo, nhi, hi)
        except Exception:
            pass
        finally:
            job["alive"] = False

    def _draw_minmax(self, g):
        if not (self.minmax_on and self.has_t) or self._mm is None:
            return
        m00, m01, m10, m11, ox, oy = self._xf()
        W, H = float(self.Width), float(self.Height)
        nlo, vlo, nhi, vhi = self._mm
        for tag, nid, v, bb, pp, up in ((u"MAX", nhi, vhi, self.b_mmax, self.p_mmax, True), (u"MIN", nlo, vlo, self.b_mmin, self.p_mmin, False)):
            p = self.dn.get(nid)
            if p is None or nid in self.hid_nodes:
                continue
            sx, sy = m00 * p[0] + m01 * p[1] + ox, m10 * p[0] + m11 * p[1] + oy
            if sx < 0.0 or sx > W or sy < 0.0 or sy > H:
                continue
            g.DrawEllipse(pp, sx - 9.0, sy - 9.0, 18.0, 18.0)
            txt = u"%s  %.5g    [%d]" % (tag, v * self.fac + self.off, nid)
            sz = g.MeasureString(txt, self.f_probe)
            bx = min(max(4.0, sx + 14.0), max(4.0, W - sz.Width - 16.0))
            by = sy - sz.Height - 14.0 if up else sy + 12.0
            by = min(max(4.0, by), max(4.0, H - sz.Height - 8.0))
            g.FillRectangle(bb, bx, by, sz.Width + 8.0, sz.Height + 4.0)
            g.DrawString(txt, self.f_probe, self.b_white, bx + 4.0, by + 2.0)

    def _blit(self, g, img, iw, ih, ixf, region=True):
        """img (piksel koordinati ixf ile taban koordinatina bagli) resmini guncel gorunume basar.
        Eksene paralelse (donus ayni) hizli dikdortgen cizimi, degilse paralelkenar cizimi."""
        xf = self._xf()
        W, H = float(self.Width), float(self.Height)
        if region:
            us, vs = [], []
            for (sx, sy) in ((0.0, 0.0), (W, 0.0), (W, H), (0.0, H)):      # ekran kosellerinin resim pikseli karsiligi
                bx, by = self._to_base(sx, sy, xf)
                us.append(ixf[0] * bx + ixf[1] * by + ixf[4])
                vs.append(ixf[2] * bx + ixf[3] * by + ixf[5])
            a0, a1 = max(0.0, min(us)), min(float(iw), max(us))
            b0, b1 = max(0.0, min(vs)), min(float(ih), max(vs))
        else:
            a0, a1, b0, b1 = 0.0, float(iw), 0.0, float(ih)
        if a1 <= a0 or b1 <= b0:
            return

        def fwd(u, v):
            bx, by = self._to_base(u, v, ixf)
            return (xf[0] * bx + xf[1] * by + xf[4], xf[2] * bx + xf[3] * by + xf[5])
        p00, p10, p01 = fwd(a0, b0), fwd(a1, b0), fwd(a0, b1)
        src = RectangleF(a0, b0, a1 - a0, b1 - b0)
        if abs(p10[1] - p00[1]) < 1e-6 and abs(p01[0] - p00[0]) < 1e-6 and p10[0] > p00[0] and p01[1] > p00[1]:
            g.DrawImage(img, RectangleF(p00[0], p00[1], p10[0] - p00[0], p01[1] - p00[1]), src, GraphicsUnit.Pixel)
        else:
            g.DrawImage(img, System.Array[PointF]([PointF(p00[0], p00[1]), PointF(p10[0], p10[1]), PointF(p01[0], p01[1])]),
                        src, GraphicsUnit.Pixel)

    def _draw_mip(self, g):
        if self._mip is None or self._mip_xf is None:
            return
        g.InterpolationMode = InterpolationMode.Bilinear
        self._blit(g, self._mip, self._mip_w, self._mip_h, self._mip_xf, True)

    def _draw_cache(self, g):
        """Eski cache resmini guncel gorunume (olcek / kaydirma / donus) cevirerek basar."""
        if self._cache is None or self._cxf is None:
            return
        g.InterpolationMode = InterpolationMode.Bilinear
        self._blit(g, self._cache, self._cache.Width, self._cache.Height, self._cxf, False)

    def _draw_base(self, g):
        """Hazir olan en iyi resmi gosterir (hazir degilse eskisini cevirerek); yeni resim arka planda hazirlanir."""
        if self._use_mip():
            if self._mip is not None:
                self._draw_mip(g)
            elif self._cache is not None:
                self._draw_cache(g)
            return
        if self._cache is not None:
            if self._cache_sig == self._sig() and self._cache_cgen == self._content_gen:
                g.DrawImage(self._cache, 0, 0)
            else:
                r = self.s / self._cs
                if (0.2 < r < 5.0) or self._mip is None:
                    self._draw_cache(g)
                else:
                    self._draw_mip(g)
        elif self._mip is not None:
            self._draw_mip(g)

    def _paint(self, sender, e):
        g = e.Graphics
        if self._fit_pending and self.Width > 60 and self.Height > 60:
            self._fit_pending = False
            self.fit_all()
        g.Clear(self.bg)
        if self.model is None:
            g.DrawString(u"Mesh / Temp / Disp dosyasi yuklenmedi.", self.f_ui, self.b_dim, 8.0, 8.0)
            return
        try:
            self._draw_base(g)
            self._draw_selection(g)
            self._draw_probes(g)
        except Exception as ex:
            g.DrawString(u"Draw error: %s" % str(ex), self.f_ui, self.b_red, 8.0, 24.0)
        g.SmoothingMode = SmoothingMode.AntiAlias
        self._draw_minmax(g)
        if self.has_t:
            self._draw_legend(g, self.f_ui)
        if self._rb is not None:
            x0, y0, x1, y1 = self._rb
            g.DrawRectangle(self.p_rb_r if self.unsel_mode else self.p_rb_o, min(x0, x1), min(y0, y1), abs(x1 - x0), abs(y1 - y0))
        self._draw_axes(g)

    def _draw_model(self, g, xf, W, H, mip, job):
        """ARKA PLAN thread'inde calisir (kendi firca / kalemlerini kullanir). Degisiklik olursa False doner (iptal).
        Eleman renkleri burada, sadece cizilen elemanlar icin hesaplanir (alan / time degisimi arayuzu bekletmez)."""
        g.SmoothingMode = SmoothingMode.HighSpeed
        br = [SolidBrush(jet(i / 255.0)) for i in range(256)] + [SolidBrush(Color.LightGray)]
        pm = Pen(Color.FromArgb(70, 70, 70), 1.0)
        mesh_on = self.mesh_on
        has_t = job["has_t"]
        get = job["tmap"].get
        vmin = job["vmin"]
        rng = job["vmax"] - vmin
        m00, m01, m10, m11, ox, oy = xf
        s = math.sqrt(abs(m00 * m11 - m01 * m10))
        gc, gv = job["cgen"], job["vgen"]
        is_cache = not mip
        bx_, by_ = [], []                                              # gorunur alanin taban koordinatindaki sinirlari
        for (sx, sy) in ((-2.0, -2.0), (W + 2.0, -2.0), (W + 2.0, H + 2.0), (-2.0, H + 2.0)):
            b = self._to_base(sx, sy, xf)
            bx_.append(b[0])
            by_.append(b[1])
        vx0, vx1, vy0, vy1 = min(bx_), max(bx_), min(by_), max(by_)

        def ci_of(ws, tv):
            T = 0.0
            for w, v in zip(ws, tv):
                T += w * v
            if rng <= 0:
                return 128
            c = int(255 * (T - vmin) / rng)
            return 0 if c < 0 else (255 if c > 255 else c)
        n = 0
        for el in self.model["elems"]:
            n += 1
            if (n & 511) == 0:
                if self._content_gen != gc or (is_cache and self._view_gen != gv):
                    return False
            if el["hid"]:
                continue
            x0, x1, y0, y1 = el["bb"]
            if x1 < vx0 or x0 > vx1 or y1 < vy0 or y0 > vy1:
                continue
            px = max(x1 - x0, y1 - y0) * s                         # en uzun kenar (ince / uzun elemanlar kaybolmasin)
            ok = False
            if has_t:
                tv = [get(i) for i in el["nodes"]]
                ok = None not in tv
            if px < 1.0:
                ci = ci_of(el["cw"], tv) if ok else 256
                g.FillRectangle(br[ci], m00 * x0 + m01 * y0 + ox - 0.5, m10 * x0 + m11 * y0 + oy - 0.5, 1.6, 1.6)
                continue
            if px < 6.0:
                ci = ci_of(el["cw"], tv) if ok else 256
                pts = System.Array[PointF]([PointF(clampf(m00 * p[0] + m01 * p[1] + ox), clampf(m10 * p[0] + m11 * p[1] + oy)) for p in el["cpoly"]])
                g.FillPolygon(br[ci], pts)
            else:
                for sp in el["subs"]:
                    ci = ci_of(sp["w"], tv) if ok else 256
                    pts = System.Array[PointF]([PointF(clampf(m00 * p[0] + m01 * p[1] + ox), clampf(m10 * p[0] + m11 * p[1] + oy)) for p in sp["pts"]])
                    g.FillPolygon(br[ci], pts)
            if mesh_on and px >= 4.0:
                ring = System.Array[PointF]([PointF(clampf(m00 * p[0] + m01 * p[1] + ox), clampf(m10 * p[0] + m11 * p[1] + oy)) for p in el["outline"]])
                g.DrawPolygon(pm, ring)
        if mip or not self.nodes_on:
            return True
        g.SmoothingMode = SmoothingMode.AntiAlias
        if self.avg_el * s >= 4.0:                                 # node noktalari (Nodes: ON ve yeterince yakinlasinca)
            dark = SolidBrush(Color.FromArgb(40, 40, 40))
            hid = self.hid_nodes
            vis = []
            for nid, (x, y) in self.dn.items():
                sx, sy = m00 * x + m01 * y + ox, m10 * x + m11 * y + oy
                if -10 <= sx <= W + 10 and -10 <= sy <= H + 10 and nid not in hid:
                    vis.append((sx, sy))
            if len(vis) <= 20000:
                for sx, sy in vis:
                    g.FillRectangle(dark, sx - 1.5, sy - 1.5, 3.0, 3.0)
        return True

    def _draw_selection(self, g):
        """Sari = tiklanmis ama listeye eklenmemis (gecici); yesil = listeye eklenmis."""
        show_sel = self.listed_on and bool(self.sel)
        if not self.pend and not show_sel:
            return
        if self._moving and ((len(self.sel) if show_sel else 0) + len(self.pend)) > 1500:     # hareket sirasinda buyuk secim isaretleri atlanir (akici zoom)
            return
        m00, m01, m10, m11, ox, oy = self._xf()
        W, H = float(self.Width), float(self.Height)
        dn = self.dn
        hid = self.hid_nodes
        g.SmoothingMode = SmoothingMode.AntiAlias
        for grp, brush, pen in ((self.pend, self.b_pend, self.p_pend), (self.sel if self.listed_on else (), self.b_green, self.p_glow)):
            if not grp:
                continue
            vis = []
            for nid in grp:
                p = dn.get(nid)
                if p is None or nid in hid:
                    continue
                sx, sy = m00 * p[0] + m01 * p[1] + ox, m10 * p[0] + m11 * p[1] + oy
                if -10 <= sx <= W + 10 and -10 <= sy <= H + 10:
                    vis.append((sx, sy))
            if len(vis) > 3000:
                for sx, sy in vis:
                    g.FillRectangle(brush, sx - 2.0, sy - 2.0, 4.0, 4.0)
            else:
                for sx, sy in vis:
                    g.FillEllipse(brush, sx - 5.5, sy - 5.5, 11.0, 11.0)
                    g.DrawEllipse(pen, sx - 5.5, sy - 5.5, 11.0, 11.0)

    def _draw_probes(self, g):
        if not self.probes:
            return
        m00, m01, m10, m11, ox, oy = self._xf()
        W, H = float(self.Width), float(self.Height)
        g.SmoothingMode = SmoothingMode.AntiAlias
        for nid in self.probes:
            p = self.dn.get(nid)
            if p is None or nid in self.hid_nodes:
                continue
            sx, sy = m00 * p[0] + m01 * p[1] + ox, m10 * p[0] + m11 * p[1] + oy
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

    def _arrow(self, g, pen, brush, ox, oy, dx, dy, L, label):
        tx, ty = ox + dx * L, oy + dy * L
        g.DrawLine(pen, ox, oy, tx, ty)
        qx, qy = -dy, dx
        g.FillPolygon(brush, System.Array[PointF]([PointF(tx + dx * 10.0, ty + dy * 10.0),
                                                   PointF(tx - dx * 3.0 + qx * 6.0, ty - dy * 3.0 + qy * 6.0),
                                                   PointF(tx - dx * 3.0 - qx * 6.0, ty - dy * 3.0 - qy * 6.0)]))
        g.DrawString(label, self.f_axis, brush, tx + dx * 20.0 - 9.0, ty + dy * 20.0 - 13.0)

    def _draw_axes(self, g):
        """Sol altta eksen takimi: X yukari (kirmizi), Y saga (yesil); model donunce birlikte doner."""
        H = float(self.Height)
        ox, oy, L = 76.0, H - 76.0, 38.0
        c, sn = math.cos(self.rot), math.sin(self.rot)
        g.SmoothingMode = SmoothingMode.AntiAlias
        self._arrow(g, self.p_axx, self.b_axx, ox, oy, sn, -c, L, u"X")
        self._arrow(g, self.p_axy, self.b_axy, ox, oy, c, sn, L, u"Y")

    def _draw_legend(self, g, f):
        """Yatay renk cubugu ekranin altinda. Ustte: ne gosterildigi (buyuk) ve SABIT Min / Max (gorunen node'larin veri uclari);
        cubugun altinda: renk skalasi etiketleri (skala her time / gizleme degisiminde otomatik: mavi = min, kirmizi = max).
        Cubuk bir bitmap'e bir kez cizilir, her paint'te tek DrawImage ile basilir."""
        W, H = float(self.Width), float(self.Height)
        M = 64.0
        LH = 120
        avail = max(2.0 * M + 80.0, W - 150.0 - 16.0)
        pw = min(avail, 780.0 + 2.0 * M)
        bw = pw - 2.0 * M
        px0 = 150.0 + (avail - pw) / 2.0
        top = H - LH
        fac, off = self.fac, self.off
        key = (self.title, round(self.vmin, 9), round(self.vmax, 9), round(off, 6), round(fac, 9), self.dext, int(pw))
        if key != self._leg_key or self._leg_bmp is None:
            if self._leg_bmp is not None:
                self._leg_bmp.Dispose()
            bmp = Bitmap(int(pw), LH)
            lg = Graphics.FromImage(bmp)
            try:
                lg.Clear(Color.FromArgb(236, 255, 255, 255))
                lg.SmoothingMode = SmoothingMode.AntiAlias
                lg.DrawString(self.title, self.f_leg_title, self.b_black, 14.0, 4.0)
                if self.dext is not None:
                    smin = u"Min  %.5g" % (self.dext[0] * fac + off)
                    smax = u"Max  %.5g" % (self.dext[2] * fac + off)
                    lg.DrawString(smin, self.f_leg, self.b_cmin, M, 34.0)
                    szx = lg.MeasureString(smax, self.f_leg)
                    lg.DrawString(smax, self.f_leg, self.b_cmax, M + bw - szx.Width, 34.0)
                yb, bh = 66.0, 20.0
                n = 128
                seg = bw / n
                for k in range(n):
                    ci = int(255 * k / float(n - 1))
                    lg.FillRectangle(self.brushes[ci], M + k * seg, yb, seg + 1.0, bh)
                lg.DrawRectangle(self.p_gray, M, yb, bw, bh)
                for frac in (0.0, 0.25, 0.5, 0.75, 1.0):
                    val = (self.vmin + (self.vmax - self.vmin) * frac) * fac + off
                    xx = M + bw * frac
                    lg.DrawLine(self.p_black, xx, yb + bh, xx, yb + bh + 5.0)
                    txt = u"%.5g" % val
                    sz = lg.MeasureString(txt, self.f_leg)
                    lx = max(2.0, min(xx - sz.Width / 2.0, pw - sz.Width - 2.0))
                    lg.DrawString(txt, self.f_leg, self.b_black, lx, yb + bh + 5.0)
            finally:
                lg.Dispose()
            self._leg_bmp, self._leg_key = bmp, key
        g.DrawImage(self._leg_bmp, int(px0), int(top))

    # ----- mouse -----
    def _wheel(self, sender, e):
        self.zoom(1.25 if e.Delta > 0 else 0.8, e.X, e.Y)

    def _mdown(self, sender, e):
        self._anim = None
        if self._rot_target is not None:
            self.rot = self._rot_target % (2.0 * math.pi)
            self._rot_target = None
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
        if self._btn == MouseButtons.Right:                                  # kutu secim: sadece sag tus
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
            m00, m01, m10, m11, ox, oy = self._xf()
            if self.hide_mode:
                hit = []
                for el in self.model["elems"]:
                    if el["hid"]:
                        continue
                    cx, cy = el["c"]
                    sx, sy = m00 * cx + m01 * cy + ox, m10 * cx + m11 * cy + oy
                    if xa <= sx <= xb and ya <= sy <= yb:
                        hit.append(el)
                self._hide_elems(hit)
            elif self.unsel_mode:
                hit = set()
                for nid in (self.sel | self.pend):
                    p = self.dn.get(nid)
                    if p is None:
                        continue
                    sx, sy = m00 * p[0] + m01 * p[1] + ox, m10 * p[0] + m11 * p[1] + oy
                    if xa <= sx <= xb and ya <= sy <= yb:
                        hit.add(nid)
                if hit:
                    self.owner.push_undo()
                    self.sel -= hit
                    self.pend -= hit
                    self.owner.selection_changed()
            else:
                hid = self.hid_nodes
                hit = set()
                for nid, (x, y) in self.dn.items():
                    sx, sy = m00 * x + m01 * y + ox, m10 * x + m11 * y + oy
                    if xa <= sx <= xb and ya <= sy <= yb and nid not in hid:
                        hit.add(nid)
                new = hit - self.sel - self.pend                      # kutu secim: gecici (sari) isaret
                if new:
                    self.owner.push_undo()
                    self.pend |= new
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
                nid = self._find_node(x0, y0, self.sel | self.pend)
                if nid is not None:
                    self.owner.push_undo()
                    self.sel.discard(nid)
                    self.pend.discard(nid)
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
        X, Y = self._to_base(mx, my, self._xf())
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
        xf = self._xf()
        m00, m01, m10, m11, ox, oy = xf
        X, Y = self._to_base(mx, my, xf)
        cell = self._cell
        k = int(math.ceil((12.0 / self.s) / cell))
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
            d = math.hypot(m00 * x + m01 * y + ox - mx, m10 * x + m11 * y + oy - my)
            if d < best:
                best, bn = d, nid
        return bn

    def _pick(self, mx, my):
        """Tiklama: node GECICI (sari) isaretlenir / isareti kalkar; zaten listede olana dokunmaz.
        Node olmayan bos bir yere tiklamak gecici (sari) secimi iptal eder."""
        bn = self._find_node(mx, my)
        if bn is None:
            if self.pend:
                self.owner.push_undo()
                self.pend = set()
                self.owner.selection_changed()
            return
        if bn in self.sel:
            return
        self.owner.push_undo()
        if bn in self.pend:
            self.pend.discard(bn)
        else:
            self.pend.add(bn)
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
    def __init__(self, mesh_path=u"", temp_path=u"", disp_path=u"", callback=None):
        self.callback = callback
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
        if self.callback is None:
            self.DialogResult = DialogResult.OK
            return
        self.Enabled = False                          # yukleme sirasinda cift tiklama olmasin
        ok = False
        try:
            ok = self.callback(m, t, d)
        finally:
            self.Enabled = True
        if ok:
            self.Close()


class ExportForm(Form):
    """Export penceresi: solda TEK time listesi (Shift = aralik, Ctrl = tek tek), ortada ekle / cikar butonlari,
    sagda export edilecek time'lar. Altta ne export edilecegi isaretlenir (Temp ve / veya Disp, varsayilan ikisi de bos).
    Kayit adi bir kez sorulur; sonuna _temp / _disp eklenerek iki ayri dosya yazilir."""
    def __init__(self, owner):
        self.owner = owner
        self.Text = TITLE + u" - Export TXT"
        self.ClientSize = Size(760, 514)
        self.StartPosition = FormStartPosition.CenterParent
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.chosen = set()                      # secilen time'lar
        self._left = []
        self._right = []
        tset = set()
        if owner.tf is not None:
            tset.update(owner.tf.times)
        if owner.df is not None:
            tset.update(owner.df.times)
        self.all_times = sorted(tset)
        has_t = owner.tf is not None
        has_d = owner.df is not None
        self.lblSel = Label()
        self.lblSel.AutoSize = False
        self.lblSel.SetBounds(12, 10, 736, 18)
        self.Controls.Add(self.lblSel)
        gl = GroupBox()
        gl.Text = u"Times  (Shift = range, Ctrl = single, then press  >  )"
        gl.SetBounds(10, 34, 330, 300)
        self.Controls.Add(gl)
        self.lbL = ListBox()
        self.lbL.SelectionMode = SelectionMode.MultiExtended
        self.lbL.SetBounds(8, 20, 314, 270)
        gl.Controls.Add(self.lbL)
        for text, y, h in ((u">", 100, self._add_sel), (u">>", 136, self._add_all),
                           (u"<", 196, self._remove_sel), (u"<<", 232, self._remove_all)):
            b = Button()
            b.Text = text
            b.SetBounds(348, y, 64, 30)
            b.Click += (lambda hh: (lambda s_, e_: hh()))(h)
            self.Controls.Add(b)
        gr = GroupBox()
        gr.Text = u"To export  (Shift / Ctrl, then press  <  to remove)"
        gr.SetBounds(420, 34, 330, 300)
        self.Controls.Add(gr)
        self.lbR = ListBox()
        self.lbR.SelectionMode = SelectionMode.MultiExtended
        self.lbR.SetBounds(8, 20, 314, 270)
        gr.Controls.Add(self.lbR)
        go = GroupBox()
        go.Text = u"Export"
        go.SetBounds(10, 340, 740, 122)
        self.Controls.Add(go)
        # --- Temp ---
        self.chkT = self._mk_chk(go, u"Temp", 10, 22, 62, False)
        self.chkT.Enabled = has_t
        self.lblTU = self._mk_lbl(go, u"Temp unit:", 90, 26, 66)
        self.cbTU = self._mk_combo(go, 160, 22, 64, [u"C", u"K"])
        # --- Disp ---
        self.chkD = self._mk_chk(go, u"Disp", 10, 56, 62, False)
        self.chkD.Enabled = has_d
        self.lblDC = self._mk_lbl(go, u"Disp columns:", 90, 60, 86)
        self.chkUX = self._mk_chk(go, u"UX", 180, 56, 52, True)
        self.chkUY = self._mk_chk(go, u"UY", 236, 56, 52, True)
        self.chkUS = self._mk_chk(go, u"USUM", 292, 56, 70, False)
        self.lblDU = self._mk_lbl(go, u"Disp unit:", 380, 60, 64)
        self.cbDU = self._mk_combo(go, 448, 56, 70, [u"mm", u"m"])
        # --- ortak ---
        self._mk_lbl(go, u"Separator:", 10, 94, 70)
        self.cbSep = self._mk_combo(go, 84, 90, 90, [u"Tab", u"Space", u"Comma"])
        self.chkMv = self._mk_chk(go, u"Add MeshVal column (mesh section 3)", 210, 90, 300, False)
        self.chkT.CheckedChanged += lambda s_, e_: self._enable()
        self.chkD.CheckedChanged += lambda s_, e_: self._enable()
        be = Button()
        be.Text = u"Export..."
        be.SetBounds(560, 472, 110, 32)
        be.Click += lambda s_, e_: self._do_export()
        self.Controls.Add(be)
        bc = Button()
        bc.Text = u"Close"
        bc.SetBounds(680, 472, 70, 32)
        bc.DialogResult = DialogResult.Cancel
        self.Controls.Add(bc)
        self.CancelButton = bc
        self._enable()
        self._refresh_label()
        self._refresh_lists()

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

    def _enable(self):
        """Temp secenekleri Temp isaretliyse, Disp secenekleri (UX / UY / USUM, birim) Disp isaretliyse aktif."""
        t, d = self.chkT.Checked, self.chkD.Checked
        for c in (self.lblTU, self.cbTU):
            c.Enabled = t
        for c in (self.lblDC, self.chkUX, self.chkUY, self.chkUS, self.lblDU, self.cbDU):
            c.Enabled = d

    def _refresh_label(self):
        self.lblSel.Text = u"Selected nodes to export: %d" % len(self.owner.view.sel)

    def _refresh_lists(self):
        self._left = [t for t in self.all_times if t not in self.chosen]
        self._right = sorted(self.chosen)
        for lb, items in ((self.lbL, self._left), (self.lbR, self._right)):
            lb.BeginUpdate()
            try:
                lb.Items.Clear()
                lb.Items.AddRange(System.Array[object]([u"t = %.6g s" % t for t in items]))
            finally:
                lb.EndUpdate()

    def _add_sel(self):
        for i in [i for i in self.lbL.SelectedIndices]:
            if 0 <= i < len(self._left):
                self.chosen.add(self._left[i])
        self._refresh_lists()

    def _add_all(self):
        self.chosen.update(self.all_times)
        self._refresh_lists()

    def _remove_sel(self):
        for i in [i for i in self.lbR.SelectedIndices]:
            if 0 <= i < len(self._right):
                self.chosen.discard(self._right[i])
        self._refresh_lists()

    def _remove_all(self):
        self.chosen = set()
        self._refresh_lists()

    def _do_export(self):
        comps = []
        if self.chkUX.Checked:
            comps.append(u"UX")
        if self.chkUY.Checked:
            comps.append(u"UY")
        if self.chkUS.Checked:
            comps.append(u"USUM")
        opts = {"times": sorted(self.chosen),
                "do_temp": self.chkT.Checked,
                "do_disp": self.chkD.Checked,
                "comps": comps,
                "disp_unit": u"m" if self.cbDU.SelectedIndex == 1 else u"mm",
                "temp_unit": u"K" if self.cbTU.SelectedIndex == 1 else u"C",
                "sep": [u"\t", u"  ", u","][max(0, self.cbSep.SelectedIndex)],
                "mv": self.chkMv.Checked}
        self.owner.run_export(opts)
        self._refresh_label()


# ---------------------------------------------------------------
# Koyu tema renkleri / sol arac seridi dugmesi
# ---------------------------------------------------------------
C_WIN = (22, 24, 28)
C_PANEL = (33, 37, 43)
C_CARD = (44, 49, 57)
C_BORDER = (70, 77, 88)
C_TEXT = (232, 235, 240)
C_MUTED = (150, 158, 170)
C_COMBO = (30, 34, 40)
GRP = {"view": (41, 121, 214), "sel": (38, 158, 98), "hide": (232, 126, 30), "tool": (140, 80, 200),
       "exp": (0, 150, 160), "file": (110, 120, 135)}
TOPH, RAILW, SIDEW = 72, 116, 480
_BR, _PN = {}, {}


def rgb(t):
    return Color.FromArgb(t[0], t[1], t[2])


def mixc(c, bg, a):
    return tuple(int(bg[i] * (1.0 - a) + c[i] * a) for i in range(3))


def brush_of(t):
    b = _BR.get(t)
    if b is None:
        b = _BR[t] = SolidBrush(rgb(t))
    return b


def pen_of(t, w=1.0):
    k = (t, w)
    p_ = _PN.get(k)
    if p_ is None:
        p_ = _PN[k] = Pen(rgb(t), w)
    return p_


class RailButton(Panel):
    """Sol arac seridi dugmesi: ustte buyuk simge, altta kucuk etiket. toggle ise set_on() ile vurgulanir."""
    def __init__(self, glyph, text, grp, handler):
        Panel.__init__(self)
        self.DoubleBuffered = True
        self.glyph, self.text_, self.grp = glyph, text, grp
        self.on = False
        self._hover = False
        self._down = False
        self._handler = handler
        self.Cursor = Cursors.Hand
        self.f_g = Font("Segoe UI Symbol", 15.0)
        self.f_t = Font("Segoe UI", 7.5, FontStyle.Bold)
        self.sf = StringFormat()
        self.sf.Alignment = StringAlignment.Center
        self.sf.LineAlignment = StringAlignment.Center
        self.Paint += self._paint
        self.MouseEnter += lambda s, e: self._set(hover=True)
        self.MouseLeave += lambda s, e: self._set(hover=False, down=False)
        self.MouseDown += lambda s, e: self._set(down=True)
        self.MouseUp += self._up

    def _set(self, hover=None, down=None):
        if hover is not None:
            self._hover = hover
        if down is not None:
            self._down = down
        self.Invalidate()

    def _up(self, sender, e):
        was = self._down
        self._down = False
        self.Invalidate()
        if was and 0 <= e.X <= self.Width and 0 <= e.Y <= self.Height:
            self._handler()

    def set_on(self, on):
        self.on = on
        self.Invalidate()

    def _paint(self, sender, e):
        g = e.Graphics
        W, H = self.Width, self.Height
        c = GRP[self.grp]
        if self.on:
            fill, border, fg, tx = c, mixc(c, (255, 255, 255), 0.45), (255, 255, 255), (255, 255, 255)
        else:
            a = 0.52 if self._down else (0.40 if self._hover else 0.26)
            fill, border = mixc(c, C_CARD, a), mixc(c, C_CARD, 0.62)
            fg, tx = mixc(c, (255, 255, 255), 0.55), (222, 226, 234)
        g.Clear(rgb(C_PANEL))
        g.FillRectangle(brush_of(fill), 0, 0, W - 1, H - 1)
        g.DrawRectangle(pen_of(border, 1.0), 0, 0, W - 1, H - 1)
        g.DrawString(self.glyph, self.f_g, brush_of(fg), RectangleF(0.0, 1.0, float(W), float(H - 17)), self.sf)
        g.DrawString(self.text_, self.f_t, brush_of(tx), RectangleF(0.0, float(H - 17), float(W), 15.0), self.sf)


class ViewerForm(Form):
    def __init__(self):
        self.Text = TITLE
        self.BackColor = rgb(C_WIN)
        wa = Screen.PrimaryScreen.WorkingArea
        self.ClientSize = Size(int(0.88 * min(1700, wa.Width - 30)), int(min(920, wa.Height - 60)))
        self.MinimumSize = Size(1000, 640)
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
        self._unit_mem = {"temp": 0, "disp": 0}     # alan basina son secilen birim (Temp: C / K, Disp: mm / m)
        self._ext_cur = None                        # (alan verisi, veri uclari) onbellegi: Show all hizli
        self.side_w = SIDEW
        self._sdrag = False
        self._undo = []
        self._last_dir = None
        self._export = None
        self._start = None
        self.FormClosed += self._on_closed
        self.mesh_path = self.temp_path = self.disp_path = u""
        self._tt = ToolTip()
        self.KeyPreview = True
        self.KeyDown += self._on_key
        self.view = SceneView(self)
        self.top = Panel()
        self.rail = Panel()
        self.side = Panel()
        self.splitter = Panel()                      # sag paneli surukleyerek genisletip daraltmak icin
        self.splitter.BackColor = rgb(C_BORDER)
        self.splitter.Cursor = Cursors.VSplit
        self.splitter.MouseDown += lambda s, e: setattr(self, "_sdrag", True)
        self.splitter.MouseUp += lambda s, e: setattr(self, "_sdrag", False)
        self.splitter.MouseMove += self._split_move
        for c in (self.view, self.top, self.rail, self.side, self.splitter):
            self.Controls.Add(c)
        self.Resize += lambda s, e: self._layout()
        self._layout()
        self._build_top()
        self._build_rail()
        self._build_side()

    def _layout(self):
        W, H = self.ClientSize.Width, self.ClientSize.Height
        sw = max(320, min(self.side_w, W - RAILW - 320))
        self.side_w = sw
        self.top.SetBounds(0, 0, W, TOPH)
        self.rail.SetBounds(0, TOPH, RAILW, H - TOPH)
        self.side.SetBounds(W - sw, TOPH, sw, H - TOPH)
        self.splitter.SetBounds(W - sw - 6, TOPH, 6, H - TOPH)
        self.view.SetBounds(RAILW, TOPH, max(100, W - RAILW - sw - 6), H - TOPH)
        self._fit_cols()

    def _split_move(self, sender, e):
        if self._sdrag:
            W = self.ClientSize.Width
            self.side_w = W - (self.splitter.Left + e.X)
            self._layout()

    def _fit_cols(self):
        """Tablonun son (deger) sutunu kalan genisligi doldurur."""
        if not self.__dict__.get("_cols_ready"):
            return
        w = self.lv.Width - (70 + 110 + 110) - 26
        if w >= 70:
            self.lv.Columns[3].Width = w

    # ----- koyu tema yardimcilari -----
    def _dlabel(self, parent, text, x, y, w, h=24, muted=False, ellipsis=False, bold=False, size=10.0):
        l = Label()
        l.AutoSize = False
        l.AutoEllipsis = ellipsis
        l.Text = text
        l.SetBounds(x, y, w, h)
        l.BackColor = parent.BackColor
        l.ForeColor = rgb(C_MUTED if muted else C_TEXT)
        l.Font = Font("Segoe UI", size, FontStyle.Bold if bold else FontStyle.Regular)
        parent.Controls.Add(l)
        return l

    def _dcombo(self, parent, x, y, w):
        c = ComboBox()
        c.DropDownStyle = ComboBoxStyle.DropDownList
        c.FlatStyle = FlatStyle.Flat
        c.BackColor = rgb(C_COMBO)
        c.ForeColor = rgb(C_TEXT)
        c.Font = Font("Segoe UI", 10.5)
        c.SetBounds(x, y, w, 30)
        parent.Controls.Add(c)
        return c

    def _dbtn(self, parent, text, x, y, w, h, handler, grp=None, size=10.0, solid=False):
        b = Button()
        b.Text = text
        b.SetBounds(x, y, w, h)
        b.FlatStyle = FlatStyle.Flat
        b.UseVisualStyleBackColor = False
        c = GRP[grp] if grp else C_CARD
        b.BackColor = rgb(c if solid else mixc(c, C_CARD, 0.35) if grp else C_CARD)
        b.ForeColor = rgb((255, 255, 255) if solid else C_TEXT)
        b.FlatAppearance.BorderColor = rgb(mixc(c, C_CARD, 0.65) if grp else C_BORDER)
        b.Font = Font("Segoe UI", size, FontStyle.Bold if solid else FontStyle.Regular)
        b.Click += handler
        parent.Controls.Add(b)
        return b

    # ----- ust cubuk: Field / Unit / Time / dosyalar -----
    def _build_top(self):
        t = self.top
        t.BackColor = rgb(C_PANEL)
        TL = AnchorStyles.Top | AnchorStyles.Left
        self._dlabel(t, u"Field", 14, 22, 44, bold=True)
        self.cbField = self._dcombo(t, 62, 18, 190)
        self.cbField.SelectedIndexChanged += lambda s, e: self.field_changed()
        self._dlabel(t, u"Unit", 270, 22, 40, bold=True)
        self.cbUnit = self._dcombo(t, 312, 18, 90)
        self.cbUnit.Items.Add(u"-")
        self.cbUnit.SelectedIndex = 0
        self.cbUnit.Enabled = False
        self.cbUnit.SelectedIndexChanged += lambda s, e: self.unit_changed()
        self._dlabel(t, u"Time [s]", 420, 22, 70, bold=True)
        self.cbTime = self._dcombo(t, 494, 18, 110)
        self.cbTime.Enabled = False
        self.cbTime.SelectedIndexChanged += lambda s, e: self.time_changed()
        self.btnPrev = self._dbtn(t, u"\u25c0", 612, 17, 40, 34, lambda s, e: self.step_time(-1), 'view')
        self.btnNext = self._dbtn(t, u"\u25b6", 658, 17, 40, 34, lambda s, e: self.step_time(1), 'view')
        self.btnPrev.Enabled = False
        self.btnNext.Enabled = False
        self.lblTinfo = self._dlabel(t, u"", 712, 22, 130, muted=True, size=9.0)
        W = t.Width
        self.lblFiles = self._dlabel(t, u"", 850, 20, W - 850 - 168, 30, muted=True, ellipsis=True, size=9.0)
        self.lblFiles.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        ob = self._dbtn(t, u"Open files\u2026", W - 154, 17, 140, 36, lambda s, e: self.ask_files(), 'file')
        ob.Anchor = AnchorStyles.Top | AnchorStyles.Right

    # ----- sol arac seridi -----
    def _build_rail(self):
        r = self.rail
        r.BackColor = rgb(C_PANEL)
        r.AutoScroll = True
        self.rb = {}
        spec = [
            ("view", [("fit", u"\u21f1", u"Fit", self.fit_view, u"Fit model to window"),
                      ("rotl", u"\u27f2", u"Rot L", lambda: self.view.rotate(-90.0), u"Rotate view 90\u00b0 counter-clockwise"),
                      ("rotr", u"\u27f3", u"Rot R", lambda: self.view.rotate(90.0), u"Rotate view 90\u00b0 clockwise"),
                      ("mesh", u"\u25a6", u"Mesh", self.toggle_mesh, u"Show / hide mesh lines"),
                      ("nodes", u"\u2237", u"Nodes", self.toggle_nodes, u"Show / hide node points")]),
            ("sel", [("selall", u"\u2611", u"Select all", self.select_all, u"Pick all visible nodes (yellow); press Add to list to keep them"),
                     ("clear", u"\u2715", u"Clear", self.clear_sel, u"Clear the picked and listed nodes"),
                     ("unsel", u"\u2296", u"Unselect", self.toggle_unsel, u"Click a node or right-drag a box to remove nodes from the picks / list"),
                     ("undo", u"\u21b6", u"Undo", self.undo, u"Undo the last selection / hide step (Ctrl+Z)")]),
            ("hide", [("hide", u"\u25a8", u"Hide el.", self.toggle_hide, u"Click an element (or right-drag a box) to hide it"),
                      ("hideo", u"\u25e7", u"Hide others", self.hide_others_clicked, u"Hide the elements that do not touch the selected nodes"),
                      ("showall", u"\u25eb", u"Show all", self.show_all, u"Show all hidden elements again")]),
            ("tool", [("probe", u"\u2295", u"Probe", self.toggle_probe, u"Click a node to show its values on the screen"),
                      ("minmax", u"\u2195", u"Max/Min", self.toggle_minmax, u"Show / hide the MAX and MIN labels of the nodes on the screen")])]
        y = 10
        for grp, items in spec:
            strip = Panel()
            strip.BackColor = rgb(GRP[grp])
            strip.SetBounds(3, y, 4, len(items) * 56 - 4)
            r.Controls.Add(strip)
            for key, glyph, text, handler, tip in items:
                b = RailButton(glyph, text, grp, handler)
                b.SetBounds(12, y, RAILW - 20, 52)
                r.Controls.Add(b)
                self._tt.SetToolTip(b, tip)
                self.rb[key] = b
                y += 56
            y += 12
        self.rb["minmax"].set_on(True)

    # ----- sag panel: secili node tablosu + Export -----
    def _build_side(self):
        s_ = self.side
        s_.BackColor = rgb(C_PANEL)
        H0 = s_.Height
        TLR = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        self.lblSel = self._dlabel(s_, u"Selected nodes (0)", 14, 8, 360, 26, bold=True, size=11.0)
        self.lblSel.Anchor = TLR
        self.btnAdd = self._dbtn(s_, u"Add to list (0)", 12, 38, 190, 34, lambda s, e: self.add_pending(), 'sel', size=10.0, solid=True)
        self.btnListed = self._dbtn(s_, u"Show listed: OFF", 210, 38, 160, 34, lambda s, e: self.toggle_listed(), 'file', size=10.0)
        self.btnAdd.Enabled = False
        self._tt.SetToolTip(self.btnAdd, u"Add the picked (yellow) nodes to this list. Clicking an empty spot of the view cancels the picks.")
        self._tt.SetToolTip(self.btnListed, u"Show the nodes in this list on the screen as green marks.")
        self.lv = ListView()
        self.lv.View = View.Details
        self.lv.FullRowSelect = True
        self.lv.GridLines = False
        self.lv.HideSelection = False
        self.lv.BorderStyle = getattr(BorderStyle, "None")
        self.lv.BackColor = rgb(C_CARD)
        self.lv.ForeColor = rgb(C_TEXT)
        self.lv.Font = Font("Consolas", 10.0)
        self.lv.Columns.Add(u"Node", 70)
        self.lv.Columns.Add(u"X", 110)
        self.lv.Columns.Add(u"Y", 110)
        self.lv.Columns.Add(u"Value", 140)
        self.lv.SetBounds(12, 80, s_.Width - 24, H0 - 80 - 76)
        self.lv.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        self.lv.OwnerDraw = True
        self.lv.DrawColumnHeader += self._lv_head
        self.lv.DrawItem += self._lv_default
        self.lv.DrawSubItem += self._lv_default
        s_.Controls.Add(self.lv)
        eb = self._dbtn(s_, u"Export TXT\u2026", 12, H0 - 62, s_.Width - 24, 50, lambda s, e: self.open_export(), 'exp', size=13.0, solid=True)
        eb.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        self._cols_ready = True
        self._fit_cols()

    def _lv_head(self, sender, e):
        try:
            e.Graphics.FillRectangle(brush_of((58, 64, 74)), e.Bounds)
            e.Graphics.DrawString(e.Header.Text, Font("Consolas", 10.0, FontStyle.Bold), brush_of(C_TEXT), e.Bounds.X + 4, e.Bounds.Y + 4)
        except Exception:
            e.DrawDefault = True

    def _lv_default(self, sender, e):
        e.DrawDefault = True

    def note(self, text):
        pass

    # ----- sol serit islevleri -----
    def fit_view(self):
        self.view.fit_all()

    def toggle_mesh(self):
        v = self.view
        v.mesh_on = not v.mesh_on
        self.rb["mesh"].set_on(v.mesh_on)
        v.mark_dirty()

    def toggle_nodes(self):
        v = self.view
        v.nodes_on = not v.nodes_on
        self.rb["nodes"].set_on(v.nodes_on)
        v.mark_dirty()

    def toggle_minmax(self):
        v = self.view
        v.minmax_on = not v.minmax_on
        self.rb["minmax"].set_on(v.minmax_on)
        v.Invalidate()

    def show_all(self):
        self.view.unhide_all()

    # ----- modlar: Hide elements / Probe / Unselect (ayni anda sadece biri) -----
    def toggle_hide(self):
        self._set_hide(not self.view.hide_mode)

    def _set_hide(self, on):
        v = self.view
        v.hide_mode = on
        self.rb["hide"].set_on(on)
        if on:
            if v.probe_on:
                self._set_probe(False)
            if v.unsel_mode:
                self._set_unsel(False)
        v._restore_cursor()
        v.Invalidate()

    def toggle_probe(self):
        self._set_probe(not self.view.probe_on)

    def _set_probe(self, on):
        v = self.view
        v.probe_on = on
        self.rb["probe"].set_on(on)
        if on and v.hide_mode:
            self._set_hide(False)
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
        self.rb["unsel"].set_on(on)
        if on:
            if v.hide_mode:
                self._set_hide(False)
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
            f_ = self._dfac()
            ux, uy, uz = vals[0] * f_, vals[1] * f_, vals[2] * f_
            lines.append(u"UX %.5g     UY %.5g     USUM %.5g   [%s]" % (ux, uy, math.sqrt(ux * ux + uy * uy + uz * uz), u"m" if f_ != 1.0 else u"mm"))
        return lines

    def refresh_probes(self):
        v = self.view
        v.probe_text = dict((nid, self._probe_lines(nid)) for nid in v.probes)
        v.Invalidate()

    # ----- dosyalar -----
    def ask_files(self):
        if self._start is not None:
            self._start.Activate()
            return
        sf = StartForm(self.mesh_path, self.temp_path, self.disp_path, lambda m, t, d: self.load_files(m, t, d))
        sf.FormClosed += lambda s_, e_: setattr(self, "_start", None)
        self._start = sf
        sf.Show(self)

    def _on_closed(self, sender, e):
        """Pencere kapaninca buyuk veriler ve cizim kaynaklari birakilir."""
        v = self.view
        try:
            v._content_gen += 1
            v._view_gen += 1
            for t in (v._timer, v._poll, v._anim_timer):
                t.Stop()
            for b in (v._cache, v._mip, v._leg_bmp):
                if b is not None:
                    b.Dispose()
            v._cache = v._mip = v._leg_bmp = None
            v.model = None
            v.dn = {}
            v._ngrid = v._egrid = None
        except Exception:
            pass
        self.tf = self.df = None
        self.cur_map = {}

    def _update_files_label(self):
        parts = []
        for pre, path in ((u"Mesh", self.mesh_path), (u"Temp", self.temp_path), (u"Disp", self.disp_path)):
            parts.append(u"%s: %s" % (pre, System.IO.Path.GetFileName(path) if path else u"-"))
        self.lblFiles.Text = u"   |   ".join(parts)
        self._tt.SetToolTip(self.lblFiles, u"\n".join(p for p in (self.mesh_path, self.temp_path, self.disp_path) if p))

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
        self._ext_cur = None
        self.view.set_model(model)
        self._set_probe(False)
        self._set_unsel(False)
        self._set_hide(False)
        self._update_files_label()
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
        self._setup_unit_combo(None)
        self.lblTinfo.Text = u""
        self.lv.Columns[3].Text = u"Value"
        self.selection_changed()
        return True

    # ----- ekranda alan gosterme (Field + Unit + Time) -----
    def _field(self):
        i = self.cbField.SelectedIndex
        if i <= 0 or i >= len(self.fields):
            return None
        return self.fields[i]

    def _src(self, fld):
        return self.tf if fld["src"] == "temp" else self.df

    def _off(self):
        """Temp icin ekran birimi ofseti (dosya degerleri C)."""
        fld = self._field()
        return 273.15 if (fld is not None and fld["src"] == "temp" and self.cbUnit.SelectedIndex == 1) else 0.0

    def _dfac(self):
        """Disp icin ekran birimi carpani (dosya degerleri mm)."""
        fld = self._field()
        return 0.001 if (fld is not None and fld["src"] == "disp" and self.cbUnit.SelectedIndex == 1) else 1.0

    def _setup_unit_combo(self, kind):
        """Unit kutusu secili alana gore: Temp -> C / K, Disp -> mm / m; son secim alan basina hatirlanir."""
        self._updating = True
        self.cbUnit.Items.Clear()
        items = {"temp": [u"C", u"K"], "disp": [u"mm", u"m"]}.get(kind, [u"-"])
        for it in items:
            self.cbUnit.Items.Add(it)
        self.cbUnit.SelectedIndex = self._unit_mem.get(kind, 0) if kind else 0
        self.cbUnit.Enabled = kind is not None
        self._updating = False

    def field_changed(self):
        if self._updating:
            return
        fld = self._field()
        if fld is None:
            self._setup_unit_combo(None)
            self._clear_field()
            return
        self._setup_unit_combo(fld["src"])
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
        if self._updating:
            return
        fld = self._field()
        if fld is None:
            return
        self._unit_mem[fld["src"]] = self.cbUnit.SelectedIndex
        if self._shown and self.cur_fld is not None and self.cur_map:
            self._push_field()
            self.refresh_probes()
            self.selection_changed()

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

    def _extremes(self, tm):
        """Gorunen (gizlenmemis) node'larin veri min / max degeri ve node'lari: (min, min node, max, max node).
        Gizleme yokken sonuc alan verisi basina onbellekte tutulur (Show all anlik)."""
        v = self.view
        full = v.n_hidden == 0
        if full and self._ext_cur is not None and self._ext_cur[0] is tm:
            return self._ext_cur[1]
        lo = hi = None
        nlo = nhi = None
        if full:
            used = v.model["used"]
            for n_, val in tm.items():
                if n_ not in used:
                    continue
                if lo is None or val < lo:
                    lo, nlo = val, n_
                if hi is None or val > hi:
                    hi, nhi = val, n_
        else:
            get = tm.get
            for n_ in v.visible_nodes():
                val = get(n_)
                if val is None:
                    continue
                if lo is None or val < lo:
                    lo, nlo = val, n_
                if hi is None or val > hi:
                    hi, nhi = val, n_
        ext = (0.0, None, 1.0, None) if lo is None else (lo, nlo, hi, nhi)
        if full:
            self._ext_cur = (tm, ext)
        return ext

    def _push_field(self):
        """Renk skalasi HER ZAMAN otomatik: gorunen node'larin min (mavi) ... max (kirmizi) degeri."""
        fld, tm, t = self.cur_fld, self.cur_map, self.cur_time
        is_t = fld["src"] == "temp"
        off = self._off() if is_t else 0.0
        fac = 1.0 if is_t else self._dfac()
        ext = self._extremes(tm)
        unit = (u"K" if off else u"C") if is_t else (u"m" if fac != 1.0 else u"mm")
        name = (u"Temp [%s]" % unit) if is_t else (u"%s [%s]" % (fld["name"], unit))
        self.lv.Columns[3].Text = name
        self.view.set_field(tm, ext[0], ext[2], off, u"%s      Time = %.6g s" % (name, t), fac, ext)

    def hidden_changed(self, kind="hide"):
        """Eleman gizlendi / geri getirildi: renk skalasi ve Min / Max kalan node'lara gore yenilenir."""
        if not (self._shown and self.cur_fld is not None and self.cur_map and self.view.model is not None):
            return
        if kind == "hide" and self._ext_cur is not None and self._ext_cur[0] is self.cur_map:
            ext = self._ext_cur[1]
            hid = self.view.hid_nodes
            if ext[1] not in hid and ext[3] not in hid:     # min ve max node'u hala gorunuyor: uclar degismez
                return
        self._push_field()

    def _apply(self, fld, t):
        tm = self._load_field(fld, t)
        if tm is None:
            return
        src = self._src(fld)
        self.cur_time, self.cur_map, self.cur_fld = t, tm, fld
        self._shown = True
        case, rpm = src.info(t)
        self.lblTinfo.Text = u"Case %s | RPM %s" % ("-" if case is None else case, "-" if rpm is None else "%.6g" % rpm)
        self._push_field()
        self.refresh_probes()
        self.selection_changed()

    def hide_others_clicked(self):
        v = self.view
        keep = set(v.sel) | set(v.pend)
        if not keep:
            MessageBox.Show(u"Once ekranda en az bir node sec.", TITLE, MessageBoxButtons.OK, MessageBoxIcon.Information)
            return
        if v.pend:
            self.push_undo()
            v.pend = set()                           # secim kullanildi: ekrandaki sari isaretler kalkar
            self.selection_changed()
        n = v.hide_others(keep)
        if n == 0:
            MessageBox.Show(u"Gizlenecek baska eleman yok.", TITLE, MessageBoxButtons.OK, MessageBoxIcon.Information)

    def selection_changed(self):
        MAXLIST = 2000
        v = self.view
        fld = self.cur_fld if self._shown else None
        off = self._off() if (fld is not None and fld["src"] == "temp") else 0.0
        fac = self._dfac() if (fld is not None and fld["src"] == "disp") else 1.0
        tm = self.cur_map
        sel = sorted(v.sel)
        items = []
        for nid in sel[:MAXLIST]:
            X, Y = self.nodes[nid]
            if fld is None:
                vt = u"-"
            else:
                val = tm.get(nid)
                vt = u"NaN" if val is None else (u"%.3f" % (val + off) if fld["src"] == "temp" else u"%.5g" % (val * fac))
            items.append(ListViewItem(System.Array[str]([u"%d" % nid, u"%.5f" % X, u"%.5f" % Y, vt])))
        if len(sel) > MAXLIST:
            items.append(ListViewItem(System.Array[str]([u"...", u"+%d more" % (len(sel) - MAXLIST), u"", u""])))
        self.lv.BeginUpdate()
        try:
            self.lv.Items.Clear()
            self.lv.Items.AddRange(System.Array[ListViewItem](items))
        finally:
            self.lv.EndUpdate()
        self.lblSel.Text = u"Selected nodes (%d)" % len(sel)
        npend = len(v.pend)
        self.btnAdd.Text = u"Add to list (%d)" % npend
        self.btnAdd.Enabled = npend > 0
        v.Invalidate()

    # ----- secim (sari = gecici, yesil = listede) / geri alma -----
    def add_pending(self):
        v = self.view
        if not v.pend:
            return
        self.push_undo()
        v.sel |= v.pend
        v.pend = set()
        self.selection_changed()

    def toggle_listed(self):
        v = self.view
        v.listed_on = not v.listed_on
        self.btnListed.Text = u"Show listed: %s" % (u"ON" if v.listed_on else u"OFF")
        self.btnListed.BackColor = rgb(mixc(GRP["sel"], C_CARD, 0.8) if v.listed_on else mixc(GRP["file"], C_CARD, 0.35))
        v.Invalidate()

    def push_undo(self):
        self._undo.append(("sel", (set(self.view.sel), set(self.view.pend))))
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
                self.view.sel, self.view.pend = data
                self.selection_changed()
                return
            if self.view.hide_stack:
                self.view.undo_hide()
                return

    def _on_key(self, sender, e):
        if e.Control and e.KeyCode == Keys.Z:
            self.undo()
            e.Handled = True

    def select_all(self):
        """Gorunen tum node'lari gecici (sari) isaretler; listeye almak icin Add to list."""
        self.push_undo()
        v = self.view
        hid = v.hid_nodes
        v.pend = set(n for n in v.dn.keys() if n not in hid) - v.sel
        self.selection_changed()

    def clear_sel(self):
        v = self.view
        if v.sel or v.pend:
            self.push_undo()
        v.sel = set()
        v.pend = set()
        self.selection_changed()

    # ----- export -----
    def open_export(self):
        if self.tf is None and self.df is None:
            return
        if self._export is not None:
            self._export.Activate()
            return
        ex = ExportForm(self)
        ex.FormClosed += lambda s_, e_: setattr(self, "_export", None)
        self._export = ex
        ex.Show(self)                                 # modeless: Ansys ve ana pencere kullanilabilir kalir

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

    def _suffixed(self, base, tag):
        """node_values.txt + 'temp' -> node_values_temp.txt"""
        d = System.IO.Path.GetDirectoryName(base)
        n = System.IO.Path.GetFileNameWithoutExtension(base)
        ext = System.IO.Path.GetExtension(base) or u".txt"
        return System.IO.Path.Combine(d, u"%s_%s%s" % (n, tag, ext))

    def _read_times(self, ff, times, sel, comps, title):
        """-> (veri, hata). veri[k] = [her comp icin node listesi] (times sirasinda; dosyada olmayan time -> None'lar).
        Iptalde hata = u""."""
        pf = ProgressForm(u"Exporting %s ..." % title)
        pf.Show()
        out = []
        err = None
        none = [None] * len(sel)
        try:
            for ti, t in enumerate(times):
                if not pf.update(float(ti) / len(times), u"%s: reading time %s   (%d / %d)" % (title, u"%.6g" % t, ti + 1, len(times))):
                    raise Cancelled()
                if t in ff.block_of:
                    out.append(ff.block_nodes(t, sel, comps))
                else:
                    out.append([none for _ in comps])
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
            MessageBox.Show(u"Listede node yok. Once ekranda node sec (sari) ve \"Add to list\" ile listeye ekle.", TITLE, MessageBoxButtons.OK, MessageBoxIcon.Information)
            return False
        times = o["times"]
        if not times:
            MessageBox.Show(u"Sag listeye en az bir time ekle.", TITLE, MessageBoxButtons.OK, MessageBoxIcon.Information)
            return False
        want_t = bool(o["do_temp"]) and self.tf is not None
        want_d = bool(o["do_disp"]) and self.df is not None
        if not (want_t or want_d):
            MessageBox.Show(u"Export icin Temp ve / veya Disp kutusunu isaretle.", TITLE, MessageBoxButtons.OK, MessageBoxIcon.Information)
            return False
        if want_d and not o["comps"]:
            MessageBox.Show(u"Disp icin en az bir sutun sec (UX / UY / USUM).", TITLE, MessageBoxButtons.OK, MessageBoxIcon.Information)
            return False
        if not self._big_ok(len(sel) * len(times) * ((1 if want_t else 0) + (len(o["comps"]) if want_d else 0))):
            return False
        base = self._ask_save(u"Save TXT  (\"_temp\" / \"_disp\" is added to the name)", u"node_values.txt")
        if not base:
            return False
        paths = {}
        if want_t:
            paths["temp"] = self._suffixed(base, u"temp")
        if want_d:
            paths["disp"] = self._suffixed(base, u"disp")
        exist = [pth for pth in paths.values() if System.IO.File.Exists(pth)]
        if exist:
            if MessageBox.Show(u"These files already exist:\n\n%s\n\nOverwrite?" % u"\n".join(exist), TITLE,
                               MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes:
                return False
        done = []
        if want_t:
            r = self._export_temp(sel, o, paths["temp"])
            if r:
                done.append(r)
        if want_d:
            r = self._export_disp(sel, o, paths["disp"])
            if r:
                done.append(r)
        if done:
            MessageBox.Show(u"Export completed.\n\n" + u"\n\n".join(done), TITLE, MessageBoxButtons.OK, MessageBoxIcon.Information)
            return True
        return False

    def _export_temp(self, sel, o, path):
        times = o["times"]
        data, err = self._read_times(self.tf, times, sel, [0], u"Temp")
        if err is not None:
            if err:
                MessageBox.Show(err, TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error)
            return None
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

    def _export_disp(self, sel, o, path):
        times = o["times"]
        comps = o["comps"]
        need = [0, 1, 2] if u"USUM" in comps else sorted(set(c for c, nm in ((0, u"UX"), (1, u"UY")) if nm in comps))
        data, err = self._read_times(self.df, times, sel, need, u"Disp")
        if err is not None:
            if err:
                MessageBox.Show(err, TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error)
            return None
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
            return None
        return u"%s  (%d nodes x %d times):\n%s" % (label, nn, nt, path)


def main():
    owner = mechanical_owner()
    f = ViewerForm()

    def start(m, t, d):
        ok = f.load_files(m, t, d)
        if ok and not f.Visible:
            if owner is not None:
                f.Show(owner)                         # modeless: Ansys kullanilabilir kalir
            else:
                f.Show()
        return ok

    sf = StartForm(u"", u"", u"", start)
    if owner is not None:
        sf.Show(owner)
    else:
        sf.Show()


main()
