# -*- coding: utf-8 -*-
# =====================================================================
#  SC03 CONVERTER 25   (Ansys Mechanical, IronPython 2.7)
#  25: hizli motor (C#) 2D'ye de uygulandi (eleman cizimi, gizleme, secim, tiklama); 3D / 2D secim bayraklari motorda: kutu secimi, Select all,
#      Add to list, Clear milyonlarca node'da bile anlik (Python kumeleri yok).
#  24: 3D icin HIZLI MOTOR (C#, calisirken derlenir): node okuma, izdusum + z-tamponlu cizim, kutu / tiklama secimi ve gizleme derlenmis kodda;
#      milyonlarca node'da her karede TUM node'lar cizilir (seyreltme yok), dondurme / zoom akici. Derlenemezse onceki Python yolu kullanilir.
#  23: 3D secimi (tiklama, kutu, select all) cizimden BAGIMSIZ: her seferinde tum node'lar guncel gorunumle projekte edilir (onizleme / eski cizim
#      sonucu kullanilmaz). Cizim hizlandi (ayni pikseldeki ayni renkli node'lar tek cizilir); ekranin ustunde 'ekrandaki / toplam node' sayaci.
#  22: 3D'de mesh'teki TUM node'lar (gizlenmeyenler) her zaman ekranda cizilir ve secilebilir: z-tampon / hucre secimi yok; Select all, kutu secim ve
#      tiklama bagimsiz olarak butun node'lar uzerinde calisir (zoom / aci degisince secilebilen node sayisi degismez).
#  21: koordinat satirlarinda bosluksuz yazilmis sayilar (Fortran E16.9: negatif sayi onceki alana yapisir) artik okunur (3D'de node'lar eksik kaliyordu);
#      okunamayan node kaydi olursa uyari verilir.
#  20: 3D'de Unselect yok, yerine 'Hide' (secili node'lari gizler); 3D acilinca bilgi penceresi cikmaz.
#  19: 2D + 3D tek kod. Z koordinati olan mesh = 3D (node bulutu: elemanlar okunmaz, z-tampon, hizli onizleme); 2D eskisi gibi.
#      Soldaki serit tek ekrana sigar (3D'de); Iso / +-X / +-Y / +-Z gorunumleri; Rotate: aci kutusu (varsayilan 5) + eksen secimi
#      (3D: X Y Z, 2D: X Y); Select all 3D'de ekranda cizilenleri secer; Export / Start pencerelerinin Close / Cancel dugmeleri kapatir.
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
#  3D TEST 1: Z koordinati olan mesh = 3D. 3D modelde dis yuzey (yuzler) orthografik izdusumle cizilir; dondurme / kaydirma / zoom,
#       node secimi, hide, probe, max/min, contour, export (Z_Loc, UZ) calisir. 2D modeller eskisi gibi acilir.
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
import heapq
import itertools

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
from System.Drawing import Rectangle
from System.Drawing.Imaging import ImageLockMode, PixelFormat
from System.Runtime.InteropServices import Marshal

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


_NUM_RE = re.compile(r"[-+]?(?:\d+\.\d*|\.\d+|\d+)(?:[EeDd][-+]?\d+)?")


def xyz_of(line):
    """Koordinat satirindan (X, Y, Z). Once bosluga gore, olmazsa (sayilar bitisik yazilmissa: 1.0E+02-2.5E+01) regex ile ayirir."""
    try:
        v = [tofloat(x) for x in line.split()]
    except ValueError:
        v = []
    if len(v) < 2:
        v = [tofloat(x) for x in _NUM_RE.findall(line)]
    if len(v) < 2:
        raise ValueError(line)
    return (v[0], v[1], v[2] if len(v) > 2 else 0.0)


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
            vals = cols[comp]
        elif self.ncomp == 3:
            vals = [math.sqrt(a * a + b * b + c * c) for a, b, c in zip(*cols)]
        else:
            vals = cols[0]
        m = ScalarMap(zip(ids, vals))
        m.ids, m.vals = ids, vals
        return m

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
                nodes[int(t[1])] = xyz_of(L[i + 1])
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
        self.R2 = (1.0, 0.0, 0.0, 0.0, 1.0, 0.0, 0.0, 0.0, 1.0)      # 2D: modelin X / Y eksenleri etrafindaki donusu (3x3)
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
        self._mip_rot = None                          # mip'in cizildigi gorunum matrisi
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
        self.R2 = (1.0, 0.0, 0.0, 0.0, 1.0, 0.0, 0.0, 0.0, 1.0)
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

    def _L(self):
        """Birim olcekli gorunum matrisi (taban koordinati -> ekran yonu): X / Y ekseni donusu (R2) ve ekran duzlemi donusu (rot)."""
        R = self.R2
        a00, a01, a10, a11 = R[0], -R[1], -R[3], R[4]             # taban Y'si asagi bakar: F * A * F
        c, sn = math.cos(self.rot), math.sin(self.rot)
        return (a00 * c + a01 * sn, -a00 * sn + a01 * c, a10 * c + a11 * sn, -a10 * sn + a11 * c)

    def _mat(self):
        L = self._L()
        s = self.s
        return (s * L[0], s * L[1], s * L[2], s * L[3])

    def rotate_axis(self, axis, deg):
        """Modeli GLOBAL X (0) / Y (1) ekseni etrafinda deg derece doner (sag el kurali); ekran merkezindeki nokta merkezde kalir."""
        if self.model is None:
            return
        self.cx, self.cy = self._to_base(self.Width / 2.0, self.Height / 2.0, self._xf())
        self.px = self.py = 0.0
        self._anim = None
        self.R2 = _ortho(_mmul(self.R2, _rot_axis(axis, math.radians(deg))))
        self._view_changed()

    def _xf(self):
        """Taban (model cizim) koordinatindan ekrana: ekran = M * taban + o  -> (m00, m01, m10, m11, ox, oy)."""
        m00, m01, m10, m11 = self._mat()
        ox = self.Width / 2.0 + self.px - (m00 * self.cx + m01 * self.cy)
        oy = self.Height / 2.0 + self.py - (m10 * self.cx + m11 * self.cy)
        return (m00, m01, m10, m11, ox, oy)

    def _to_base(self, x, y, xf):
        m00, m01, m10, m11, ox, oy = xf
        det = m00 * m11 - m01 * m10
        if abs(det) < 1e-12:
            det = 1e-12
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
        L = self._L()
        pts = [(x0, y0), (x1, y0), (x1, y1), (x0, y1)]
        xs = [L[0] * (x - self.cx) + L[1] * (y - self.cy) for x, y in pts]
        ys = [L[2] * (x - self.cx) + L[3] * (y - self.cy) for x, y in pts]
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

    def _mip_geom_for(self, L):
        """Mip bitmap'inin boyutu ve donusumu; gorunum matrisi mip'e GOMULUR, boylece zoom / pan hizli (eksene paralel) cizilir."""
        x0, x1, y0, y1 = self.model["box"]
        cx, cy = (x0 + x1) / 2.0, (y0 + y1) / 2.0
        ms = self._mip_ms
        pts = [(x0, y0), (x1, y0), (x1, y1), (x0, y1)]
        xs = [L[0] * (x - cx) + L[1] * (y - cy) for x, y in pts]
        ys = [L[2] * (x - cx) + L[3] * (y - cy) for x, y in pts]
        mw = int((max(xs) - min(xs)) * ms) + 3
        mh = int((max(ys) - min(ys)) * ms) + 3
        ox = 1.0 - ms * (L[0] * cx + L[1] * cy) - ms * min(xs)
        oy = 1.0 - ms * (L[2] * cx + L[3] * cy) - ms * min(ys)
        return mw, mh, (ms * L[0], ms * L[1], ms * L[2], ms * L[3], ox, oy)

    def _sig(self):
        return (self.s, self._L(), self.px, self.py, self.cx, self.cy, self.Width, self.Height)

    def _need_render(self):
        if self.model is None or self.Width < 10 or self.Height < 10 or self._moving:
            return None
        if self._use_mip():
            if self._mip is None or self._mip_cgen != self._content_gen or self._mip_rot != self._L():
                return "mip"
            return None
        if self._cache is None or self._cache_cgen != self._content_gen or self._cache_sig != self._sig():
            return "cache"
        return None

    def _start_job(self, kind):
        if kind == "mip":
            mw, mh, mxf = self._mip_geom_for(self._L())
            job = {"kind": "mip", "s": self._mip_ms, "xf": mxf, "W": mw, "H": mh, "rot": self._L()}
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
        """Sol altta eksen takimi: X (kirmizi), Y (yesil); model donunce / egilince birlikte doner."""
        H = float(self.Height)
        ox, oy, L = 76.0, H - 76.0, 38.0
        M = self._L()                                    # taban (X, -Y) -> ekran yonu
        g.SmoothingMode = SmoothingMode.AntiAlias
        self._arrow(g, self.p_axx, self.b_axx, ox, oy, M[0], M[2], L, u"X")
        self._arrow(g, self.p_axy, self.b_axy, ox, oy, -M[1], -M[3], L, u"Y")

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
# 3D = NODE BULUTU: Z koordinati olan mesh 3D sayilir; elemanlar okunmaz, sadece node'lar cizilir
# ---------------------------------------------------------------
import array as _arrmod
_arr = _arrmod.array
PREVIEW_POINTS = 30000         # donerken cizilen en fazla nokta sayisi (hareket bitince TUM node'lar cizilir)
BIG_SEL = 3000                 # bundan fazla secili node: secim renkleri cizimin icine islenir

def _fl(tok):
    try:
        return float(tok)
    except ValueError:
        return tofloat(tok)

class NodeTable(object):
    """Node koordinatlari (kompakt diziler). nodes[nid] -> (X, Y, Z). Numaralar 1..N ardisiksa dict kullanilmaz."""
    def __init__(self, ids, X, Y, Z):
        self.ids, self.X, self.Y, self.Z = ids, X, Y, Z
        n = len(ids)
        self.first = ids[0] if n else 1
        c = True
        first = self.first
        for i in range(n):
            if ids[i] != first + i:
                c = False
                break
        self.contig = c
        self._map = None

    def index(self, nid):
        if self.contig:
            i = nid - self.first
            return i if 0 <= i < len(self.ids) else -1
        if self._map is None:
            self._map = dict((v, i) for i, v in enumerate(self.ids))
        return self._map.get(nid, -1)

    def id_of(self, i):
        return i + self.first if self.contig else self.ids[i]

    def __getitem__(self, nid):
        i = self.index(nid)
        if i < 0:
            raise KeyError(nid)
        return (self.X[i], self.Y[i], self.Z[i])

    def __contains__(self, nid):
        return self.index(nid) >= 0

    def __len__(self):
        return len(self.ids)

    def keys(self):
        return self.ids

class AllNodes(object):
    """Modeldeki tum node'lar (model['used']): uyelik testi ve uzunluk."""
    def __init__(self, nt):
        self.nt = nt

    def __contains__(self, nid):
        return self.nt.index(nid) >= 0

    def __len__(self):
        return len(self.nt)

    def __iter__(self):
        return iter(self.nt.ids)

class ElemCount(object):
    def __init__(self, n):
        self.n = n

    def __len__(self):
        return self.n

class ArrayMap(object):
    """MeshVal icin (mesh 3. bolum degerleri): .get(nid)."""
    def __init__(self, ids, vals):
        self.ids, self.vals, self._d = ids, vals, None

    def get(self, nid, default=None):
        if self._d is None:
            self._d = dict(zip(self.ids, self.vals))
        return self._d.get(nid, default)

    def __len__(self):
        return len(self.ids)

def _mmul(A, B):
    return tuple(A[3 * i] * B[j] + A[3 * i + 1] * B[3 + j] + A[3 * i + 2] * B[6 + j] for i in range(3) for j in range(3))

def _rot_axis(axis, a):
    """Eksen (0 = X, 1 = Y, 2 = Z) etrafinda a radyan donus matrisi (sag el kurali)."""
    c, s_ = math.cos(a), math.sin(a)
    if axis == 0:
        return (1.0, 0.0, 0.0, 0.0, c, -s_, 0.0, s_, c)
    if axis == 1:
        return (c, 0.0, s_, 0.0, 1.0, 0.0, -s_, 0.0, c)
    return (c, -s_, 0.0, s_, c, 0.0, 0.0, 0.0, 1.0)

def _ortho(R):
    """Tekrarlanan carpimlardaki kucuk sapmayi giderir (satirlari birim ve dik yapar)."""
    def nrm(v):
        L = math.sqrt(v[0] * v[0] + v[1] * v[1] + v[2] * v[2]) or 1.0
        return (v[0] / L, v[1] / L, v[2] / L)

    def crs(a, b):
        return (a[1] * b[2] - a[2] * b[1], a[2] * b[0] - a[0] * b[2], a[0] * b[1] - a[1] * b[0])
    z = nrm(R[6:9])
    x = nrm(crs(R[3:6], z))
    y = crs(z, x)
    return x + y + z

def _R_azel(az, el):
    ca, sa = math.cos(az), math.sin(az)
    ce, se = math.cos(el), math.sin(el)
    return (ca, -sa, 0.0, se * sa, se * ca, ce, -ce * sa, -ce * ca, se)

def _look(zv, up):
    z = tuple(float(v) for v in zv)
    x = (up[1] * z[2] - up[2] * z[1], up[2] * z[0] - up[0] * z[2], up[0] * z[1] - up[1] * z[0])
    y = (z[1] * x[2] - z[2] * x[1], z[2] * x[0] - z[0] * x[2], z[0] * x[1] - z[1] * x[0])
    return _ortho(x + y + z)

def _coeffs(R, s, W, H, pan, ctr):
    """Model (X, Y, Z) -> ekran: sx = a0 X + a1 Y + a2 Z + a3, sy = b0 ..., derinlik = c0 ... (ortografik izdusum)."""
    cx, cy, cz = ctr
    ox, oy = W / 2.0 + pan[0], H / 2.0 + pan[1]
    return ((s * R[0], s * R[1], s * R[2], ox - s * (R[0] * cx + R[1] * cy + R[2] * cz)),
            (-s * R[3], -s * R[4], -s * R[5], oy + s * (R[3] * cx + R[4] * cy + R[5] * cz)),
            (R[6], R[7], R[8], -(R[6] * cx + R[7] * cy + R[8] * cz)))


def _same_R(A, B):
    return max(abs(A[i] - B[i]) for i in range(9)) < 1e-9

def jet_rgb(t):
    t = max(0.0, min(1.0, t))
    r = min(1.0, max(0.0, 1.5 - abs(4 * t - 3)))
    g = min(1.0, max(0.0, 1.5 - abs(4 * t - 2)))
    b = min(1.0, max(0.0, 1.5 - abs(4 * t - 1)))
    return (int(r * 255), int(g * 255), int(b * 255))

class NodeKeys(object):
    """view.dn yerine: Select all icin node numaralari (.keys())."""
    def __init__(self, nt):
        self.nt = nt

    def keys(self):
        return self.nt.ids

    def __len__(self):
        return len(self.nt)

class HidNodesCloud(object):
    """view.hid_nodes: gizlenen node'lar."""
    def __init__(self, view):
        self.v = view

    def __contains__(self, nid):
        v = self.v
        if v.n_hidden == 0 or v.nhid is None:
            return False
        i = v.model["nodes"].index(nid)
        return i < 0 or bool(v.nhid[i])

class PointCloudView(SceneView):
    """3D gorunum = NODE BULUTU: her node bir nokta (elemanlar / yuzeyler kullanilmaz). Yazilim z-tamponu: ayni piksel hucresine dusen
    noktalardan en yakin olan cizilir (ic node'lar on yuzeyin arkasinda gizlenir). Renk = contour, hafif derinlik golgesi. Gizleme node'lari
    gizler. Sol tus surukle = dondur, orta tus (veya Ctrl+sol) = kaydir, tekerlek = zoom, sag tus surukle = kutu secim, tiklama = node sec.
    Dondururken seyreltilmis (az noktali) hizli onizleme, durunca tum noktalar."""
    def __init__(self, owner):
        SceneView.__init__(self, owner)
        self.R = _R_azel(math.radians(-35.0), math.radians(28.0))       # gorunum (dondurme) matrisi: satirlar = ekran sag, yukari, kameraya
        self.s3 = 1.0                                 # piksel / model birimi
        self.pan3 = [0.0, -40.0]
        self.ctr3 = (0.0, 0.0, 0.0)
        self.rad3 = 1.0
        self._img3 = None
        self._img3_sig = None
        self._img3_mode = None
        self._img3_cgen = -1
        self._img3_hv = -1
        self._img3_selv = None
        self._img3_big = False
        self._job3 = None
        self._res3 = None
        self._n_full = None                           # son TAM cizimde ekrandaki (gizlenmemis) node sayisi
        self.sel_version = 0                          # secim her degistiginde artar (ViewerForm.selection_changed)
        self._hide_ver = 0
        self._g3 = None
        self.nhid = None                              # node gizli bayragi (bytearray)
        self.p_azz = Pen(Color.FromArgb(90, 150, 255), 3.0)
        self.b_azz = SolidBrush(Color.FromArgb(90, 150, 255))
        self.hid_nodes = HidNodesCloud(self)

    def _selsig(self):
        return (self.sel_version, self.listed_on)

    def _bigsel(self):
        return (len(self.pend) + (len(self.sel) if self.listed_on else 0)) > BIG_SEL

    def _rot9(self):
        return self.R

    def _vsig3(self):
        return (self.R, self.s3, self.pan3[0], self.pan3[1], self.ctr3, self.Width, self.Height)

    def _aff(self):
        """Son cizimin (resim) guncel gorunume (kaydirma / zoom) donusumu: (olcek, ofset x, ofset y)."""
        sig = self._img3_sig
        if sig is None:
            return (1.0, 0.0, 0.0)
        R0, s_img, px_img, py_img, ctr, W0, H0 = sig
        if _same_R(R0, self.R) and ctr == self.ctr3 and W0 == self.Width and H0 == self.Height:
            r = self.s3 / s_img
            cxs, cys = self.Width / 2.0, self.Height / 2.0
            return (r, cxs + self.pan3[0] - r * (cxs + px_img), cys + self.pan3[1] - r * (cys + py_img))
        return (1.0, 0.0, 0.0)

    def project(self, X, Y, Z):
        """Model noktasi -> ekran (x, y, derinlik) (guncel gorunum)."""
        r = self._rot9()
        cx, cy, cz = self.ctr3
        x, y, z = X - cx, Y - cy, Z - cz
        return (self.Width / 2.0 + self.pan3[0] + self.s3 * (r[0] * x + r[1] * y + r[2] * z),
                self.Height / 2.0 + self.pan3[1] - self.s3 * (r[3] * x + r[4] * y + r[5] * z),
                r[6] * x + r[7] * y + r[8] * z)

    def _commit_hide(self):
        self._hide_ver += 1
        self._vis_cache = None
        self.mark_dirty()

    def visible_nodes(self):
        nt = self.model["nodes"]
        return [nt.id_of(i) for i in self._visible_node_indices()]

    def _visible_box3(self):
        m = self.model
        if self.n_hidden == 0:
            return m["box3"]
        nt = m["nodes"]
        idx = list(self._visible_node_indices())
        if not idx:
            return m["box3"]
        xs = [nt.X[i] for i in idx]
        ys = [nt.Y[i] for i in idx]
        zs = [nt.Z[i] for i in idx]
        return (min(xs), max(xs), min(ys), max(ys), min(zs), max(zs))

    def fit_all(self):
        """Ekranda kalan (gizlenmemis) kismi sigdirir."""
        if self.model is None:
            return
        x0, x1, y0, y1, z0, z1 = self._visible_box3()
        self.ctr3 = ((x0 + x1) / 2.0, (y0 + y1) / 2.0, (z0 + z1) / 2.0)
        self.rad3 = max(0.5 * math.sqrt((x1 - x0) ** 2 + (y1 - y0) ** 2 + (z1 - z0) ** 2), 1e-12)
        self.s3 = 0.5 * min(self.Width - 40.0, self.Height - 170.0) / self.rad3
        self.pan3 = [0.0, -50.0]
        self._view_changed()

    def rotate(self, ddeg):
        self.rotate_axis(2, ddeg)

    def rotate_axis(self, axis, deg):
        """Modeli GLOBAL X / Y / Z ekseni (0 / 1 / 2) etrafinda deg derece doner (sag el kurali: + = o eksenin ucundan bakinca saat yonu tersi)."""
        if self.model is None:
            return
        self.R = _ortho(_mmul(self.R, _rot_axis(axis, math.radians(deg))))
        self._view_changed()

    def set_preset(self, name):
        """iso, +x, -x, +y, -y, +z, -z: kamera o eksenin o tarafindan bakar."""
        if name == "iso":
            self.R = _R_azel(math.radians(-35.0), math.radians(28.0))
        else:
            sg = 1.0 if name[0] == "+" else -1.0
            ax = name[1]
            zv = {"x": (sg, 0.0, 0.0), "y": (0.0, sg, 0.0), "z": (0.0, 0.0, sg)}[ax]
            self.R = _look(zv, (0.0, 1.0, 0.0) if ax == "z" else (0.0, 0.0, 1.0))
        self.fit_all()

    def zoom(self, f, mx, my):
        dx, dy = mx - self.Width / 2.0, my - self.Height / 2.0
        self.pan3 = [dx - (dx - self.pan3[0]) * f, dy - (dy - self.pan3[1]) * f]
        self.s3 = max(1e-9, self.s3 * f)
        self._view_changed()

    def _on_resize(self, sender, e):
        self._view_changed()

    def _wheel(self, sender, e):
        self.zoom(1.2 if e.Delta > 0 else 1.0 / 1.2, e.X, e.Y)

    def _need_render3(self):
        if self.model is None or self.Width < 10 or self.Height < 10:
            return None
        if self._img3 is None:
            return "preview"                           # ilk resim: once hizli onizleme, sonra tum node'lar
        same = (self._img3_sig == self._vsig3() and self._img3_cgen == self._content_gen and self._img3_hv == self._hide_ver)
        seldirty = (self._img3_big or self._bigsel()) and self._img3_selv != self._selsig()
        if same and not seldirty and (self._img3_mode == "full" or self._moving):
            return None
        return "preview" if (self._moving or seldirty) else "full"

    def _run_job3(self, job):
        bmp = None
        try:
            bmp = Bitmap(job["W"], job["H"])
            g = Graphics.FromImage(bmp)
            try:
                g.Clear(self.bg)
                res = self._draw_faces(g, job)
            finally:
                g.Dispose()
            if res is not None:
                self._res3 = (job, bmp, res)
                bmp = None
        except Exception:
            pass
        finally:
            if bmp is not None:
                bmp.Dispose()
            job["alive"] = False


    def _on_poll(self, sender, e):
        try:
            r = self._res3
            if r is not None:
                self._res3 = None
                job, bmp, res = r
                ok = self.model is not None and job["cgen"] == self._content_gen and res["hv"] == self._hide_ver
                if ok and job["mode"] == "full" and job["sig"] != self._vsig3():
                    ok = False
                if ok:
                    old = self._img3
                    self._img3, self._img3_sig, self._img3_mode, self._img3_cgen = bmp, job["sig"], job["mode"], job["cgen"]
                    self._img3_hv = res["hv"]
                    self._img3_selv, self._img3_big = job["selv"], job["bigsel"]
                    if job["mode"] == "full":
                        self._mm = res["mm"] if (self.minmax_on and self.has_t) else None
                        self._n_full = res["n"]
                    if old is not None:
                        old.Dispose()
                    self.Invalidate()
                else:
                    bmp.Dispose()
            if self._job3 is not None and not self._job3["alive"]:
                self._job3 = None
            if self._job3 is None and self.model is not None:
                mode = self._need_render3()
                if mode is not None:
                    self._start_job3(mode)
        except Exception:
            pass

    def _draw_img3(self, g):
        if self._img3 is None:
            return
        r, ox, oy = self._aff()
        if abs(r - 1.0) < 1e-9 and abs(ox) < 1e-6 and abs(oy) < 1e-6:
            g.DrawImage(self._img3, 0, 0)
        else:
            g.DrawImage(self._img3, RectangleF(ox, oy, self._img3.Width * r, self._img3.Height * r))

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
            self._draw_img3(g)
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
        if self._n_full is not None:
            tag = u""
            if getattr(self, "cloud", None) is not None:
                tag = u"   [fast engine]"
            elif _NATIVE["err"]:
                tag = u"   [fast engine OFF: %s]" % _NATIVE["err"][:90]
            g.DrawString(u"Nodes on screen: %d  |  total: %d%s" % (self._n_full, self.model["nn"] - self.n_hidden, tag), self.f_ui, self.b_dim, 8.0, 6.0)

    def _scr(self, nid):
        """Node'un GUNCEL gorunumdeki ekran konumu (gizliyse None). Cizimden bagimsiz, dogrudan izdusum."""
        nt = self.model["nodes"]
        i = nt.index(nid)
        if i < 0 or (self.nhid is not None and self.nhid[i]):
            return None
        return self.project(nt.X[i], nt.Y[i], nt.Z[i])

    def _draw_selection(self, g):
        show_sel = self.listed_on and bool(self.sel)
        if not self.pend and not show_sel:
            return
        if self._bigsel() or self._img3_big:                   # cok sayida secili node: renkler cizimin icine islenir
            return
        W, H = float(self.Width), float(self.Height)
        g.SmoothingMode = SmoothingMode.AntiAlias
        for grp, brush, pen in ((self.pend, self.b_pend, self.p_pend), (self.sel if self.listed_on else (), self.b_green, self.p_glow)):
            for nid in grp:
                p = self._scr(nid)
                if p is not None and -10 <= p[0] <= W + 10 and -10 <= p[1] <= H + 10:
                    g.FillEllipse(brush, p[0] - 5.5, p[1] - 5.5, 11.0, 11.0)
                    g.DrawEllipse(pen, p[0] - 5.5, p[1] - 5.5, 11.0, 11.0)

    def _draw_probes(self, g):
        if not self.probes:
            return
        W, H = float(self.Width), float(self.Height)
        g.SmoothingMode = SmoothingMode.AntiAlias
        for nid in self.probes:
            p = self._scr(nid)
            if p is None:
                continue
            sx, sy = p[0], p[1]
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

    def _draw_minmax(self, g):
        if not (self.minmax_on and self.has_t) or self._mm is None:
            return
        W, H = float(self.Width), float(self.Height)
        nlo, vlo, nhi, vhi = self._mm
        for tag, nid, v, bb, pp, up in ((u"MAX", nhi, vhi, self.b_mmax, self.p_mmax, True), (u"MIN", nlo, vlo, self.b_mmin, self.p_mmin, False)):
            p = self._scr(nid)
            if p is None or p[0] < 0.0 or p[0] > W or p[1] < 0.0 or p[1] > H:
                continue
            sx, sy = p[0], p[1]
            g.DrawEllipse(pp, sx - 9.0, sy - 9.0, 18.0, 18.0)
            txt = u"%s  %.5g    [%d]" % (tag, v * self.fac + self.off, nid)
            sz = g.MeasureString(txt, self.f_probe)
            bx = min(max(4.0, sx + 14.0), max(4.0, W - sz.Width - 16.0))
            by = sy - sz.Height - 14.0 if up else sy + 12.0
            by = min(max(4.0, by), max(4.0, H - sz.Height - 8.0))
            g.FillRectangle(bb, bx, by, sz.Width + 8.0, sz.Height + 4.0)
            g.DrawString(txt, self.f_probe, self.b_white, bx + 4.0, by + 2.0)

    def _draw_axes(self, g):
        """Sol altta 3 eksenli takim (X kirmizi, Y yesil, Z mavi); gorunume gore doner."""
        H = float(self.Height)
        ox, oy, L = 76.0, H - 76.0, 38.0
        r = self._rot9()
        g.SmoothingMode = SmoothingMode.AntiAlias
        self._arrow(g, self.p_axx, self.b_axx, ox, oy, r[0], -r[3], L, u"X")
        self._arrow(g, self.p_axy, self.b_axy, ox, oy, r[1], -r[4], L, u"Y")
        self._arrow(g, self.p_azz, self.b_azz, ox, oy, r[2], -r[5], L, u"Z")

    def _mdown(self, sender, e):
        self._g3 = {"x": e.X, "y": e.Y, "b": e.Button, "R": self.R, "px": self.pan3[0], "py": self.pan3[1],
                    "drag": False, "ctrl": (Control.ModifierKeys & Keys.Control) == Keys.Control}

    def _mmove(self, sender, e):
        g3 = self._g3
        if g3 is None:
            return
        dx, dy = e.X - g3["x"], e.Y - g3["y"]
        if not g3["drag"] and math.hypot(dx, dy) > 4:
            g3["drag"] = True
        if not g3["drag"]:
            return
        b = g3["b"]
        if b == MouseButtons.Right:
            self._rb = (g3["x"], g3["y"], e.X, e.Y)
            self.Invalidate()
        elif b == MouseButtons.Middle or g3["ctrl"]:
            self.pan3 = [g3["px"] + dx, g3["py"] + dy]
            self._view_changed()
        else:
            self.R = _ortho(_mmul(_rot_axis(0, dy * 0.008), _mmul(_rot_axis(1, dx * 0.008), g3["R"])))      # ekranin yatay / dikey ekseni etrafinda
            self._view_changed()

    def _mup(self, sender, e):
        g3 = self._g3
        if g3 is None:
            return
        self._g3 = None
        x0, y0 = g3["x"], g3["y"]
        if g3["drag"] and self._rb is not None and g3["b"] == MouseButtons.Right:
            xa, xb = sorted((x0, e.X))
            ya, yb = sorted((y0, e.Y))
            self._box_action(xa, xb, ya, yb)
        elif not g3["drag"] and g3["b"] == MouseButtons.Left:
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
        self._rb = None
        self.Invalidate()

    def _find_node(self, mx, my, only=None):
        """Imlece en yakin node (12 px icinde; arkada kalanlar dahil hepsi); esitlikte one olan. Guncel gorunumle dogrudan hesaplanir."""
        nt = self.model["nodes"]
        (a0, a1, a2, a3), (b0, b1, b2, b3), (c0, c1, c2, c3) = _coeffs(self.R, self.s3, self.Width, self.Height, self.pan3, self.ctr3)
        nh = self.nhid
        best, bn, bz = 12.0, None, None
        for i, (x, y, z) in enumerate(zip(nt.X, nt.Y, nt.Z)):
            dx = a0 * x + a1 * y + a2 * z + a3 - mx
            if dx > 12.0 or dx < -12.0:
                continue
            dy = b0 * x + b1 * y + b2 * z + b3 - my
            if dy > 12.0 or dy < -12.0 or nh[i]:
                continue
            d = math.hypot(dx, dy)
            if d > best + 1e-9:
                continue
            nid = nt.id_of(i)
            if only is not None and nid not in only:
                continue
            dep = c0 * x + c1 * y + c2 * z
            if d < best - 1e-9 or bn is None or dep > bz:
                best, bn, bz = d, nid, dep
        return bn

    def _nodes_in_box(self, xa, xb, ya, yb):
        """Ekran koordinatlarinda [xa,xb] x [ya,yb] kutusundaki TUM (gizlenmemis) node indeksleri. Her seferinde guncel gorunumle tum node'lar
        projekte edilir: onizleme / eski cizim sonucu kullanilmaz, derinlikten bagimsiz."""
        nt = self.model["nodes"]
        (a0, a1, a2, a3), (b0, b1, b2, b3), c = _coeffs(self.R, self.s3, self.Width, self.Height, self.pan3, self.ctr3)
        nh = self.nhid
        out = []
        for i, (x, y, z) in enumerate(zip(nt.X, nt.Y, nt.Z)):
            u = a0 * x + a1 * y + a2 * z + a3
            if u < xa or u > xb:
                continue
            w = b0 * x + b1 * y + b2 * z + b3
            if w < ya or w > yb:
                continue
            if nh[i]:
                continue
            out.append(i)
        return out

    def _box_action(self, xa, xb, ya, yb):
        """Sag tus kutusu: Hide modunda kutudaki node'lari gizler; Unselect modunda secimden cikarir; yoksa gecici (sari) secer. Kutudaki TUM node'lar."""
        nt = self.model["nodes"]
        inside = self._nodes_in_box(xa, xb, ya, yb)
        if self.hide_mode:
            self._hide_elems(inside)
            return
        inside = set(nt.id_of(i) for i in inside)
        if self.unsel_mode:
            hit = inside & (self.sel | self.pend)
            if hit:
                self.owner.push_undo()
                self.sel -= hit
                self.pend -= hit
                self.owner.selection_changed()
        else:
            new = inside - self.sel - self.pend
            if new:
                self.owner.push_undo()
                self.pend |= new
                self.owner.selection_changed()

    def set_model(self, model):
        self._content_gen += 1
        self._view_gen += 1
        self.model = model
        self.dn = NodeKeys(model["nodes"])
        self.sel = set()
        self.pend = set()
        self.vis_nodes = None
        self.hide_stack = []
        self.nhid = bytearray(model["nn"])
        self.n_hidden = 0
        self._hide_ver += 1
        self.probes = []
        self.probe_text = {}
        self.has_t = False
        self._mm = None
        self._mm_sig = None
        self._n_full = None
        if self._img3 is not None:
            self._img3.Dispose()
        self._img3 = None
        self._img3_sig = None
        self._img3_selv = None
        self._fit_pending = True
        self.Invalidate()

    def _hide_elems(self, idxs):
        nh = self.nhid
        new = [i for i in idxs if not nh[i]]
        if not new:
            return
        self.hide_stack.append(bytearray(nh))
        if len(self.hide_stack) > 25:
            del self.hide_stack[0]
        for i in new:
            nh[i] = 1
        self.n_hidden += len(new)
        self.owner.push_hide()
        self._commit_hide()
        self.owner.hidden_changed("hide")

    def undo_hide(self):
        if not self.hide_stack:
            return
        self.nhid = self.hide_stack.pop()
        self.n_hidden = sum(self.nhid)
        self._commit_hide()
        self.owner.hidden_changed("undo")

    def unhide_all(self):
        if self.n_hidden == 0:
            return
        self.nhid = bytearray(self.model["nn"])
        self.n_hidden = 0
        self.hide_stack = []
        self.owner.hides_cleared()
        self._commit_hide()
        self.owner.hidden_changed("unhide")

    def hide_others(self, keep):
        nt = self.model["nodes"]
        kidx = set(nt.index(n) for n in keep)
        kidx.discard(-1)
        if not kidx:
            return 0
        old = self.n_hidden
        self.hide_stack.append(bytearray(self.nhid))
        if len(self.hide_stack) > 25:
            del self.hide_stack[0]
        nh = bytearray([1]) * self.model["nn"]
        for i in kidx:
            nh[i] = 0
        self.nhid = nh
        self.n_hidden = self.model["nn"] - len(kidx)
        self.owner.push_hide()
        self._commit_hide()
        self.owner.hidden_changed("hide")
        return max(0, self.n_hidden - old)

    def _visible_node_indices(self):
        nn = self.model["nn"]
        if self.n_hidden == 0:
            return range(nn)
        nh = self.nhid
        return [i for i in range(nn) if not nh[i]]

    def _pick_elem(self, mx, my):
        nid = self._find_node(mx, my)
        return None if nid is None else self.model["nodes"].index(nid)

    def _start_job3(self, mode):
        nt = self.model["nodes"]
        job = {"mode": mode, "sig": self._vsig3(), "rot": self._rot9(), "s": self.s3, "pan": (self.pan3[0], self.pan3[1]),
               "ctr": self.ctr3, "W": self.Width, "H": self.Height, "cgen": self._content_gen, "hv": self._hide_ver,
               "vgen": self._view_gen, "has_t": self.has_t, "tmap": self.tmap, "vmin": self.vmin, "vmax": self.vmax,
               "first": nt.first, "contig": nt.contig, "nhid": self.nhid, "rad": self.rad3,
               "pend": self.pend, "sel": self.sel, "showsel": self.listed_on, "bigsel": self._bigsel(), "selv": self._selsig(),
               "track": bool(self.minmax_on), "alive": True}
        self._job3 = job
        th = threading.Thread(target=self._run_job3, args=(job,))
        th.daemon = True
        th.start()

    def _draw_faces(self, g, job):
        """(adi eski: _run_job3 cagirir) TUM gizlenmemis node'lari izdusumle ve ciz. Ayni pikselde, ayni renk / derinlik katmaninda olan node'lar gorsel olarak
        ayni oldugu icin tek cizilir (secim bundan ETKILENMEZ: secim tum node'lar uzerinden ayrica yapilir). Iptalde None."""
        m = self.model
        nt = m["nodes"]
        NX, NY, NZ, ids = nt.X, nt.Y, nt.Z, nt.ids
        first, contig = job["first"], job["contig"]
        nn = m["nn"]
        nh = job["nhid"]
        (a0, a1, a2, a3), (b0, b1, b2, b3), (c0, c1, c2, c3) = _coeffs(job["rot"], job["s"], job["W"], job["H"], job["pan"], job["ctr"])
        W, H = job["W"], job["H"]
        has_t = job["has_t"]
        get = job["tmap"].get
        vmin = job["vmin"]
        rng = job["vmax"] - vmin
        gc, gv = job["cgen"], job["vgen"]
        preview = job["mode"] == "preview"
        pend, sel, showsel, big = job["pend"], job["sel"], job["showsel"], job["bigsel"]
        track = has_t and job["track"] and not preview
        rad = job["rad"] if job["rad"] > 0 else 1.0
        zlo, zk = -rad, 7.0 / (2.0 * rad)
        step = 1
        ps = 3.0 if nn <= 250000 else 2.0
        if preview and nn > PREVIEW_POINTS:
            step = int(math.ceil(nn / float(PREVIEW_POINTS)))
            ps = min(8.0, 2.0 + 0.8 * math.sqrt(step))
        hp = ps * 0.5
        groups = {}
        seen = {}
        lo = hi = None
        nlo = nhi = None
        ic = 0
        count = 0
        for i in range(0, nn, step):
            ic += 1
            if (ic & 131071) == 0 and (self._content_gen != gc or (not preview and self._view_gen != gv)):
                return None
            if nh[i]:
                continue
            x, y, z = NX[i], NY[i], NZ[i]
            sx = a0 * x + a1 * y + a2 * z + a3
            if sx < 0.0 or sx >= W:
                continue
            sy = b0 * x + b1 * y + b2 * z + b3
            if sy < 0.0 or sy >= H:
                continue
            count += 1
            nid = i + first if contig else ids[i]
            ci = 256
            if big and nid in pend:
                ci = 257
            elif big and showsel and nid in sel:
                ci = 258
            elif has_t:
                v = get(nid)
                if v is not None:
                    ci = 128 if rng <= 0 else int(255 * (v - vmin) / rng)
                    ci = 0 if ci < 0 else (255 if ci > 255 else ci)
                    if track:
                        if lo is None or v < lo:
                            lo, nlo = v, nid
                        if hi is None or v > hi:
                            hi, nhi = v, nid
            li = int((c0 * x + c1 * y + c2 * z + c3 - zlo) * zk)
            li = 0 if li < 0 else (7 if li > 7 else li)
            k = (li, ci)
            st = seen.get(k)
            pk = int(sy) * 8192 + int(sx)
            if st is None:
                seen[k] = st = set()
                groups[k] = []
            elif pk in st:
                continue
            st.add(pk)
            groups[k].append(RectangleF(sx - hp, sy - hp, ps, ps))
        g.SmoothingMode = SmoothingMode.HighSpeed
        for k in sorted(groups):                              # uzaktan yakina (derinlik katmanlari)
            li, ci = k
            if ci == 257:
                base, kk = (255, 215, 0), 1.0
            elif ci == 258:
                base, kk = (60, 200, 90), 1.0
            else:
                base = (190, 192, 198) if ci == 256 else jet_rgb(ci / 255.0)
                kk = 0.65 + 0.35 * li / 7.0
            b = SolidBrush(Color.FromArgb(int(base[0] * kk), int(base[1] * kk), int(base[2] * kk)))
            g.FillRectangles(b, System.Array[RectangleF](groups[k]))
            b.Dispose()
        mm = (nlo, lo, nhi, hi) if lo is not None else None
        return {"mm": mm, "n": count, "hv": job["hv"]}


def read_nodes3d(path, progress=None):
    """Mesh dosyasinin SADECE node bolumunu akisla okur (eleman bolumune gelince durur). -> (NodeTable, 3D_mu)"""
    size = max(int(System.IO.FileInfo(path).Length), 1)
    ids, X, Y, Z = _arr("i"), _arr("d"), _arr("d"), _arr("d")
    f = open(path, "r")
    try:
        head = []
        for _ in range(12):
            ln = f.readline()
            if not ln:
                break
            head.append(ln)
        start = 4
        for i in range(len(head) - 2):
            t0, t1, t2 = head[i].split(), head[i + 1].split(), head[i + 2].split()
            if len(t0) >= 2 and t0[0] == "1" and t2 and t2[0].upper().startswith("1G"):
                try:
                    xyz_of(head[i + 1])
                    start = i
                    break
                except ValueError:
                    pass
        it = iter(itertools.chain(head[start:], f))
        consumed = 0
        nrec = 0
        nbad = 0
        sample = u""
        while True:
            try:
                ln = next(it)
            except StopIteration:
                break
            t = ln.split()
            if not t:
                continue
            if t[0] == "2" and len(t) >= 2:
                break                                          # eleman bolumu basladi: node'lar bitti
            if t[0] == "1" and len(t) >= 2:
                try:
                    cl = next(it)
                    gl = next(it)
                    nid = int(t[1])
                    x, y, z = xyz_of(cl)
                except (ValueError, IndexError):
                    nbad += 1
                    if not sample:
                        sample = (ln.rstrip() + u"\n" + cl.rstrip())[:200]
                    continue
                except StopIteration:
                    break
                ids.append(nid)
                X.append(x)
                Y.append(y)
                Z.append(z)
                consumed += len(ln) + len(cl) + len(gl) + 3
                nrec += 1
                if progress is not None and (nrec & 65535) == 0:
                    if not progress(min(consumed / float(size), 0.99), u"Reading nodes ...  %d" % nrec):
                        raise Cancelled()
    finally:
        f.close()
    if not len(ids):
        raise ValueError(u"Mesh dosyasinda node bulunamadi.")
    ext = max(max(X) - min(X), max(Y) - min(Y), 1e-12)
    is3 = (max(Z) - min(Z)) > 1e-9 * ext
    return NodeTable(ids, X, Y, Z), is3, nbad, sample

def cloud_model(nt):
    return {"is3d": True, "nodes": nt, "nn": len(nt), "ne": 0, "used": AllNodes(nt), "elems": ElemCount(0), "skipped": 0,
            "counts": {}, "bad": {},
            "box3": (min(nt.X), max(nt.X), min(nt.Y), max(nt.Y), min(nt.Z), max(nt.Z))}


# ---------------------------------------------------------------
# HIZLI MOTOR (C#): 3D node bulutu ve 2D mesh icin okuma, cizim, secim, gizleme - derlenmis kod (milyonlarca node / elemanda Python'dan yuzlerce kat
# hizli). Calisma aninda Mechanical'in kendi .NET derleyicisiyle (CSharpCodeProvider) derlenir; derlenemezse Python yoluna dusulur.
# ---------------------------------------------------------------
SC3_NATIVE_SRC = r'''using System;
using System.IO;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;

namespace SC3
{
    public static class Fio
    {
        static readonly Regex numRe = new Regex(@"[-+]?(?:\d+\.\d*|\.\d+|\d+)(?:[EeDd][-+]?\d+)?", RegexOptions.Compiled);
        static readonly Regex fortRe = new Regex(@"^([-+]?\d*\.?\d+)([-+]\d+)$", RegexOptions.Compiled);
        public static readonly char[] seps = new char[] { ' ', '\t' };

        public static bool ParseNum(string s, out double v)
        {
            s = s.Replace('d', 'E').Replace('D', 'E');
            if (double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out v)) return true;
            Match m = fortRe.Match(s);
            if (m.Success)
                return double.TryParse(m.Groups[1].Value + "E" + m.Groups[2].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out v);
            v = 0.0;
            return false;
        }

        // koordinat satiri: once bosluga gore, olmazsa (bitisik yazilmis sayilar) regex ile
        public static bool ParseXyz(string line, out double x, out double y, out double z)
        {
            x = y = z = 0.0;
            string[] t = line.Split(seps, StringSplitOptions.RemoveEmptyEntries);
            double[] v = new double[3];
            int n = 0;
            bool ok = true;
            for (int i = 0; i < t.Length; i++)
            {
                double d;
                if (!ParseNum(t[i], out d)) { ok = false; break; }
                if (n < 3) v[n] = d;
                n++;
            }
            if (!ok || n < 2)
            {
                n = 0;
                MatchCollection mc = numRe.Matches(line);
                for (int i = 0; i < mc.Count; i++)
                {
                    double d;
                    if (!ParseNum(mc[i].Value, out d)) { n = 0; break; }
                    if (n < 3) v[n] = d;
                    n++;
                }
            }
            if (n < 2) return false;
            x = v[0]; y = v[1]; z = n > 2 ? v[2] : 0.0;
            return true;
        }
    }

    // Node secimi (bit bayraklari): 1 = gecici (sari), 2 = listede (yesil). Hem 3D bulut hem 2D mesh kullanir; hepsi derlenmis kodda
    // (milyonlarca node'da Python kumeleri yerine).
    public abstract class SelBase
    {
        public int N;
        public int[] Ids;
        public byte[] Sel;
        public abstract bool NodeVisible(int i);

        public int SelCount(int mask)
        {
            int n = 0;
            for (int i = 0; i < N; i++) if ((Sel[i] & mask) != 0) n++;
            return n;
        }
        public bool SelHas(int mask, int i) { return i >= 0 && i < N && (Sel[i] & mask) != 0; }
        public void SelSetOne(int mask, int i, bool on)
        {
            if (i < 0 || i >= N) return;
            if (on) Sel[i] |= (byte)mask; else Sel[i] &= (byte)~mask;
        }
        public void SelClear(int mask) { byte m = (byte)~mask; for (int i = 0; i < N; i++) Sel[i] &= m; }
        // idx'teki node'lara bit ekler (excl bitlerinden birine sahip olanlar haric); eklenen sayisi
        public int SelAddIdx(int mask, int[] idx, int excl)
        {
            int n = 0;
            for (int k = 0; k < idx.Length; k++)
            {
                int i = idx[k];
                if (i < 0 || i >= N || (Sel[i] & excl) != 0 || (Sel[i] & mask) != 0) continue;
                Sel[i] |= (byte)mask; n++;
            }
            return n;
        }
        public int SelRemoveIdx(int mask, int[] idx)
        {
            int n = 0;
            for (int k = 0; k < idx.Length; k++)
            {
                int i = idx[k];
                if (i < 0 || i >= N || (Sel[i] & mask) == 0) continue;
                Sel[i] &= (byte)~mask; n++;
            }
            return n;
        }
        public int SelOrMask(int dst, int src)
        {
            int n = 0;
            for (int i = 0; i < N; i++) if ((Sel[i] & src) != 0 && (Sel[i] & dst) == 0) { Sel[i] |= (byte)dst; n++; }
            return n;
        }
        public int SelAndNotMask(int dst, int src)
        {
            int n = 0;
            for (int i = 0; i < N; i++) if ((Sel[i] & src) != 0 && (Sel[i] & dst) != 0) { Sel[i] &= (byte)~dst; n++; }
            return n;
        }
        public int SelMove(int from, int to)
        {
            int n = 0;
            for (int i = 0; i < N; i++) if ((Sel[i] & from) != 0) { Sel[i] = (byte)((Sel[i] & ~from) | to); n++; }
            return n;
        }
        // gorunen tum node'lara bit ekler (excl bitlerine sahip olanlar haric)
        public int SelAllVisible(int mask, int excl)
        {
            int n = 0;
            for (int i = 0; i < N; i++) if (NodeVisible(i) && (Sel[i] & excl) == 0) { Sel[i] |= (byte)mask; n++; }
            return n;
        }
        public byte[] SelSnapshot() { return (byte[])Sel.Clone(); }
        public void SelRestore(byte[] s) { Buffer.BlockCopy(s, 0, Sel, 0, N); }
        public int[] SelIndices(int mask)
        {
            List<int> r = new List<int>();
            for (int i = 0; i < N; i++) if ((Sel[i] & mask) != 0) r.Add(i);
            return r.ToArray();
        }
        public int[] SelIds(int mask)
        {
            List<int> r = new List<int>();
            for (int i = 0; i < N; i++) if ((Sel[i] & mask) != 0) r.Add(Ids[i]);
            return r.ToArray();
        }
        // en kucuk k node numarasi (artan)
        public int[] SelSmallestIds(int mask, int k)
        {
            int[] all = SelIds(mask);
            if (all.Length <= k) { Array.Sort(all); return all; }
            Array.Sort(all);
            int[] r = new int[k];
            Array.Copy(all, r, k);
            return r;
        }
    }
}


namespace SC3
{
    // SC03 3D node bulutu motoru: node okuma, izdusum + z-tamponlu cizim (piksel dizisine), kutu / tiklama secimi, gizleme.
    // Saf yonetilen kod (System.Drawing yok); Python tarafi piksel dizisini Bitmap'e kopyalar.
    public sealed class Cloud : SelBase
    {
        public double[] X, Y, Z;
        public int First;
        public bool Contig;
        public byte[] Hidden;
        public float[] Vals;
        public int BadCount;
        public string BadSample = "";
        public bool Is3D;
        public double MinX, MaxX, MinY, MaxY, MinZ, MaxZ;
        public int LastCount, LastMinIdx, LastMaxIdx;
        public float[] ZBuf;
        Dictionary<int, int> map;
        readonly object sync = new object();

        public override bool NodeVisible(int i) { return Hidden[i] == 0; }

        // Mesh dosyasinin SADECE node bolumunu okur (eleman kaydina "2" gelince durur).
        public static Cloud Load(string path)
        {
            List<int> ids = new List<int>(1 << 20);
            List<double> xs = new List<double>(1 << 20), ys = new List<double>(1 << 20), zs = new List<double>(1 << 20);
            int bad = 0;
            string sample = "";
            using (StreamReader sr = new StreamReader(path, System.Text.Encoding.ASCII, false, 1 << 20))
            {
                string ln;
                while ((ln = sr.ReadLine()) != null)
                {
                    string[] t = ln.Split(Fio.seps, StringSplitOptions.RemoveEmptyEntries);
                    if (t.Length == 0) continue;
                    if (t[0] == "2" && t.Length >= 2) break;
                    if (t[0] == "1" && t.Length >= 2)
                    {
                        string cl = sr.ReadLine();
                        string gl = sr.ReadLine();
                        if (cl == null || gl == null) break;
                        int id;
                        double x, y, z;
                        if (int.TryParse(t[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out id) && Fio.ParseXyz(cl, out x, out y, out z))
                        {
                            ids.Add(id); xs.Add(x); ys.Add(y); zs.Add(z);
                        }
                        else
                        {
                            bad++;
                            if (sample.Length == 0) { sample = ln.TrimEnd() + "\n" + cl.TrimEnd(); if (sample.Length > 200) sample = sample.Substring(0, 200); }
                        }
                    }
                }
            }
            Cloud c = new Cloud();
            c.N = ids.Count;
            c.Ids = ids.ToArray(); c.X = xs.ToArray(); c.Y = ys.ToArray(); c.Z = zs.ToArray();
            c.BadCount = bad; c.BadSample = sample;
            c.Init();
            return c;
        }

        // Python tarafindan dizilerle de kurulabilir (test / yedek yol)
        public static Cloud FromArrays(int[] ids, double[] x, double[] y, double[] z)
        {
            Cloud c = new Cloud();
            c.N = ids.Length; c.Ids = ids; c.X = x; c.Y = y; c.Z = z;
            c.Init();
            return c;
        }

        void Init()
        {
            Hidden = new byte[N];
            Sel = new byte[N];
            Vals = new float[N];
            for (int i = 0; i < N; i++) Vals[i] = float.NaN;
            First = N > 0 ? Ids[0] : 1;
            Contig = true;
            for (int i = 0; i < N; i++) if (Ids[i] != First + i) { Contig = false; break; }
            if (!Contig)
            {
                map = new Dictionary<int, int>(N);
                for (int i = 0; i < N; i++) map[Ids[i]] = i;
            }
            if (N > 0)
            {
                MinX = MaxX = X[0]; MinY = MaxY = Y[0]; MinZ = MaxZ = Z[0];
                for (int i = 1; i < N; i++)
                {
                    if (X[i] < MinX) MinX = X[i]; if (X[i] > MaxX) MaxX = X[i];
                    if (Y[i] < MinY) MinY = Y[i]; if (Y[i] > MaxY) MaxY = Y[i];
                    if (Z[i] < MinZ) MinZ = Z[i]; if (Z[i] > MaxZ) MaxZ = Z[i];
                }
                double ext = Math.Max(Math.Max(MaxX - MinX, MaxY - MinY), 1e-12);
                Is3D = (MaxZ - MinZ) > 1e-9 * ext;
            }
        }

        public int IndexOf(int id)
        {
            if (Contig) { int i = id - First; return (i >= 0 && i < N) ? i : -1; }
            int j;
            return map.TryGetValue(id, out j) ? j : -1;
        }

        public int[] IndexMany(int[] ids)
        {
            int[] r = new int[ids.Length];
            for (int k = 0; k < ids.Length; k++) r[k] = IndexOf(ids[k]);
            return r;
        }

        public int[] IdsOf(int[] idx)
        {
            int[] r = new int[idx.Length];
            for (int k = 0; k < idx.Length; k++) r[k] = Ids[idx[k]];
            return r;
        }

        // ---------------- degerler / secim / gizleme ----------------
        public void SetValues(int[] ids, double[] vals)
        {
            lock (sync)
            {
                for (int i = 0; i < N; i++) Vals[i] = float.NaN;
                for (int k = 0; k < ids.Length; k++) { int i = IndexOf(ids[k]); if (i >= 0) Vals[i] = (float)vals[k]; }
            }
        }

        public void ClearValues()
        {
            lock (sync) { for (int i = 0; i < N; i++) Vals[i] = float.NaN; }
        }

        public int HideIdx(int[] idx)
        {
            int n = 0;
            lock (sync) { for (int k = 0; k < idx.Length; k++) { int i = idx[k]; if (i >= 0 && i < N && Hidden[i] == 0) { Hidden[i] = 1; n++; } } }
            return n;
        }

        public int HideAllExcept(int[] keep)
        {
            lock (sync)
            {
                for (int i = 0; i < N; i++) Hidden[i] = 1;
                for (int k = 0; k < keep.Length; k++) { int i = keep[k]; if (i >= 0 && i < N) Hidden[i] = 0; }
            }
            return CountHidden();
        }

        public void ClearHidden() { lock (sync) { Array.Clear(Hidden, 0, N); } }

        public int CountHidden()
        {
            int n = 0;
            for (int i = 0; i < N; i++) if (Hidden[i] != 0) n++;
            return n;
        }

        public byte[] CopyHidden() { lock (sync) { return (byte[])Hidden.Clone(); } }

        public void SetHidden(byte[] h) { lock (sync) { Buffer.BlockCopy(h, 0, Hidden, 0, N); } }

        public int[] HiddenIds()
        {
            List<int> r = new List<int>();
            for (int i = 0; i < N; i++) if (Hidden[i] != 0) r.Add(Ids[i]);
            return r.ToArray();
        }

        // gizlenmemis node'larin sinir kutusu: minx, maxx, miny, maxy, minz, maxz (gizlenmemis yoksa null)
        public double[] VisibleBounds()
        {
            double[] b = null;
            for (int i = 0; i < N; i++)
            {
                if (Hidden[i] != 0) continue;
                if (b == null) b = new double[] { X[i], X[i], Y[i], Y[i], Z[i], Z[i] };
                else
                {
                    if (X[i] < b[0]) b[0] = X[i]; if (X[i] > b[1]) b[1] = X[i];
                    if (Y[i] < b[2]) b[2] = Y[i]; if (Y[i] > b[3]) b[3] = Y[i];
                    if (Z[i] < b[4]) b[4] = Z[i]; if (Z[i] > b[5]) b[5] = Z[i];
                }
            }
            return b;
        }

        // ---------------- izdusum katsayilari ----------------
        // sx = a0 X + a1 Y + a2 Z + a3 ; sy = b0 X + ... ; derinlik = c0 X + ...  (R: 9 sayi, satirlar = ekran sag, yukari, kameraya)
        static double[] Coeffs(double[] R, double s, double W, double H, double panX, double panY, double cx, double cy, double cz)
        {
            double ox = W / 2.0 + panX, oy = H / 2.0 + panY;
            return new double[] {
                s * R[0], s * R[1], s * R[2], ox - s * (R[0] * cx + R[1] * cy + R[2] * cz),
                -s * R[3], -s * R[4], -s * R[5], oy + s * (R[3] * cx + R[4] * cy + R[5] * cz),
                R[6], R[7], R[8], -(R[6] * cx + R[7] * cy + R[8] * cz) };
        }

        // ---------------- cizim ----------------
        // px: W*H ARGB; palette: 256 renk (contour). TUM gizlenmemis node'lar; piksel basina en yakin (z-tampon).
        public int Render(int[] px, int W, int H, double[] R, double s, double panX, double panY, double cx, double cy, double cz,
                          int ps, int bg, int[] palette, bool hasT, double vmin, double vmax, bool bigSel, bool showSel, double rad, bool track)
        {
            lock (sync)
            {
                int total = W * H;
                if (ZBuf == null || ZBuf.Length != total) ZBuf = new float[total];
                float[] zb = ZBuf;
                for (int i = 0; i < total; i++) { px[i] = bg; zb[i] = float.NegativeInfinity; }
                double[] k = Coeffs(R, s, W, H, panX, panY, cx, cy, cz);
                double a0 = k[0], a1 = k[1], a2 = k[2], a3 = k[3], b0 = k[4], b1 = k[5], b2 = k[6], b3 = k[7], c0 = k[8], c1 = k[9], c2 = k[10], c3 = k[11];
                double rng = vmax - vmin;
                double r2 = rad > 0 ? 2.0 * rad : 2.0;
                int count = 0;
                int minI = -1, maxI = -1;
                float lo = float.MaxValue, hi = float.MinValue;
                int off0 = -((ps - 1) / 2);
                int yellow = unchecked((int)0xFFFFD700), green = unchecked((int)0xFF3CC85A), gray = unchecked((int)0xFFBEC0C6);
                for (int i = 0; i < N; i++)
                {
                    if (Hidden[i] != 0) continue;
                    double x = X[i], y = Y[i], z = Z[i];
                    double sx = a0 * x + a1 * y + a2 * z + a3;
                    if (sx < 0.0 || sx >= W) continue;
                    double sy = b0 * x + b1 * y + b2 * z + b3;
                    if (sy < 0.0 || sy >= H) continue;
                    count++;
                    double d = c0 * x + c1 * y + c2 * z + c3;
                    int col;
                    bool shade = true;
                    byte sf = Sel[i];
                    if (bigSel && (sf & 1) != 0) { col = yellow; shade = false; }
                    else if (bigSel && showSel && (sf & 2) != 0) { col = green; shade = false; }
                    else
                    {
                        col = gray;
                        if (hasT)
                        {
                            float v = Vals[i];
                            if (!float.IsNaN(v))
                            {
                                int ci = rng <= 0 ? 128 : (int)(255.0 * (v - vmin) / rng);
                                if (ci < 0) ci = 0; else if (ci > 255) ci = 255;
                                col = palette[ci];
                                if (track)
                                {
                                    if (v < lo) { lo = v; minI = i; }
                                    if (v > hi) { hi = v; maxI = i; }
                                }
                            }
                        }
                    }
                    if (shade)
                    {
                        double t = (d + rad) / r2;
                        if (t < 0.0) t = 0.0; else if (t > 1.0) t = 1.0;
                        double f = 0.65 + 0.35 * t;
                        int r = (int)(((col >> 16) & 255) * f), g = (int)(((col >> 8) & 255) * f), b = (int)((col & 255) * f);
                        col = unchecked((int)0xFF000000) | (r << 16) | (g << 8) | b;
                    }
                    int ix = (int)sx + off0, iy = (int)sy + off0;
                    float fd = (float)d;
                    for (int yy = iy; yy < iy + ps; yy++)
                    {
                        if (yy < 0 || yy >= H) continue;
                        int row = yy * W;
                        for (int xx = ix; xx < ix + ps; xx++)
                        {
                            if (xx < 0 || xx >= W) continue;
                            int p = row + xx;
                            if (fd > zb[p]) { zb[p] = fd; px[p] = col; }
                        }
                    }
                }
                LastCount = count; LastMinIdx = minI; LastMaxIdx = maxI;
                return count;
            }
        }

        // ---------------- secim (cizimden bagimsiz, guncel gorunumle) ----------------
        public int[] BoxSelect(double[] R, double s, double W, double H, double panX, double panY, double cx, double cy, double cz,
                               double xa, double xb, double ya, double yb)
        {
            double[] k = Coeffs(R, s, W, H, panX, panY, cx, cy, cz);
            double a0 = k[0], a1 = k[1], a2 = k[2], a3 = k[3], b0 = k[4], b1 = k[5], b2 = k[6], b3 = k[7];
            List<int> r = new List<int>();
            for (int i = 0; i < N; i++)
            {
                double x = X[i], y = Y[i], z = Z[i];
                double u = a0 * x + a1 * y + a2 * z + a3;
                if (u < xa || u > xb) continue;
                double w = b0 * x + b1 * y + b2 * z + b3;
                if (w < ya || w > yb) continue;
                if (Hidden[i] != 0) continue;
                r.Add(i);
            }
            return r.ToArray();
        }

        // imlece en yakin node (rad piksel icinde); esitlikte one olan. Yoksa -1.
        public int Pick(double[] R, double s, double W, double H, double panX, double panY, double cx, double cy, double cz, double mx, double my, double rad)
        {
            double[] k = Coeffs(R, s, W, H, panX, panY, cx, cy, cz);
            double a0 = k[0], a1 = k[1], a2 = k[2], a3 = k[3], b0 = k[4], b1 = k[5], b2 = k[6], b3 = k[7], c0 = k[8], c1 = k[9], c2 = k[10];
            double best = rad, bz = 0.0;
            int bi = -1;
            for (int i = 0; i < N; i++)
            {
                double x = X[i], y = Y[i], z = Z[i];
                double dx = a0 * x + a1 * y + a2 * z + a3 - mx;
                if (dx > rad || dx < -rad) continue;
                double dy = b0 * x + b1 * y + b2 * z + b3 - my;
                if (dy > rad || dy < -rad || Hidden[i] != 0) continue;
                double d = Math.Sqrt(dx * dx + dy * dy);
                if (d > best + 1e-9) continue;
                double dep = c0 * x + c1 * y + c2 * z;
                if (bi < 0 || d < best - 1e-9 || dep > bz) { best = d; bi = i; bz = dep; }
            }
            return bi;
        }

        // gizlenmemis node'lar arasinda degeri olanlarin min / max node indeksi: {minIdx, maxIdx} (yoksa -1)
        public int[] Extremes()
        {
            int mi = -1, ma = -1;
            float lo = float.MaxValue, hi = float.MinValue;
            for (int i = 0; i < N; i++)
            {
                if (Hidden[i] != 0) continue;
                float v = Vals[i];
                if (float.IsNaN(v)) continue;
                if (v < lo) { lo = v; mi = i; }
                if (v > hi) { hi = v; ma = i; }
            }
            return new int[] { mi, ma };
        }
    }
}


namespace SC3
{
    // 2D mesh motoru: node + eleman okuma, eleman cizimi (konturlu, yazilim rasterizer), gizleme, secim, tiklama.
    // Taban koordinati = (X, -Y) (ekranda Y yukari). Ekran = M * taban + o.
    public sealed class Mesh2D : SelBase
    {
        public double[] X, Y;
        public int First;
        public bool Contig;
        Dictionary<int, int> map;
        public int NE;
        public int[] Conn;      // eleman basina 8: koseler (nc) + orta node'lar (nc, varsa), kalani -1
        public byte[] Nc;       // 3 / 4
        public byte[] HasMid;
        public int[] EId;
        public byte[] EHid;
        public int[] NCnt;      // node basina gorunen eleman sayisi
        public byte[] Used;
        public float[] Vals;
        public double[] Bb;     // eleman basina 4: x0, x1, y0, y1 (taban)
        public double[] Cen;    // eleman basina 2: kose cokgeninin merkezi (taban)
        public double AvgEl, BoxX0, BoxX1, BoxY0, BoxY1;
        public int Skipped, BadNodes;
        public int[] MvIds;
        public double[] MvVals;
        public int HiddenElems;
        double[] SX, SY;
        readonly object sync = new object();

        public override bool NodeVisible(int i) { return Used[i] != 0 && NCnt[i] > 0; }

        public int IndexOf(int id)
        {
            if (Contig) { int i = id - First; return (i >= 0 && i < N) ? i : -1; }
            int j;
            return map.TryGetValue(id, out j) ? j : -1;
        }
        public int[] IndexMany(int[] ids)
        {
            int[] r = new int[ids.Length];
            for (int k = 0; k < ids.Length; k++) r[k] = IndexOf(ids[k]);
            return r;
        }
        public int[] IdsOf(int[] idx)
        {
            int[] r = new int[idx.Length];
            for (int k = 0; k < idx.Length; k++) r[k] = Ids[idx[k]];
            return r;
        }

        // ---------------- okuma ----------------
        public static Mesh2D Load(string path)
        {
            List<int> ids = new List<int>(1 << 18);
            List<double> xs = new List<double>(1 << 18), ys = new List<double>(1 << 18);
            List<int> rawIds = new List<int>(1 << 20), rawOff = new List<int>(1 << 18), rawEid = new List<int>(1 << 18);
            List<int> mvi = new List<int>();
            List<double> mvv = new List<double>();
            int bad = 0;
            using (StreamReader sr = new StreamReader(path, System.Text.Encoding.ASCII, false, 1 << 20))
            {
                string ln;
                while ((ln = sr.ReadLine()) != null)
                {
                    string[] t = ln.Split(Fio.seps, StringSplitOptions.RemoveEmptyEntries);
                    if (t.Length < 2) continue;
                    int k0;
                    if (!int.TryParse(t[0], NumberStyles.Integer, CultureInfo.InvariantCulture, out k0)) continue;
                    if (k0 == 1)
                    {
                        string cl = sr.ReadLine();
                        string gl = sr.ReadLine();
                        if (cl == null || gl == null) break;
                        int id; double x, y, z;
                        if (int.TryParse(t[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out id) && Fio.ParseXyz(cl, out x, out y, out z))
                        { ids.Add(id); xs.Add(x); ys.Add(y); }
                        else bad++;
                    }
                    else if (k0 == 2)
                    {
                        string hl = sr.ReadLine();
                        if (hl == null) break;
                        string[] ht = hl.Split(Fio.seps, StringSplitOptions.RemoveEmptyEntries);
                        int cnt, eid;
                        if (ht.Length == 0 || !int.TryParse(ht[0], NumberStyles.Integer, CultureInfo.InvariantCulture, out cnt) || !int.TryParse(t[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out eid)) continue;
                        int start = rawIds.Count;
                        bool ok = true;
                        int got = 0;
                        while (got < cnt)
                        {
                            string nl = sr.ReadLine();
                            if (nl == null) { ok = false; break; }
                            string[] nt = nl.Split(Fio.seps, StringSplitOptions.RemoveEmptyEntries);
                            for (int q = 0; q < nt.Length && got < cnt; q++)
                            {
                                int v;
                                if (!int.TryParse(nt[q], NumberStyles.Integer, CultureInfo.InvariantCulture, out v)) { ok = false; break; }
                                rawIds.Add(v); got++;
                            }
                            if (!ok) break;
                        }
                        if (!ok) { rawIds.RemoveRange(start, rawIds.Count - start); continue; }
                        rawOff.Add(start); rawEid.Add(eid);
                    }
                    else if (k0 == 10)
                    {
                        string vl = sr.ReadLine();
                        if (vl == null) break;
                        string[] vt = vl.Split(Fio.seps, StringSplitOptions.RemoveEmptyEntries);
                        int id; double v;
                        if (vt.Length > 0 && int.TryParse(t[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out id) && Fio.ParseNum(vt[0], out v)) { mvi.Add(id); mvv.Add(v); }
                    }
                }
            }
            Mesh2D m = new Mesh2D();
            m.N = ids.Count; m.Ids = ids.ToArray(); m.X = xs.ToArray(); m.Y = ys.ToArray(); m.BadNodes = bad;
            m.MvIds = mvi.ToArray(); m.MvVals = mvv.ToArray();
            m.InitNodes();
            m.BuildElements(rawIds, rawOff, rawEid);
            return m;
        }

        void InitNodes()
        {
            First = N > 0 ? Ids[0] : 1;
            Contig = true;
            for (int i = 0; i < N; i++) if (Ids[i] != First + i) { Contig = false; break; }
            if (!Contig) { map = new Dictionary<int, int>(N); for (int i = 0; i < N; i++) map[Ids[i]] = i; }
            Sel = new byte[N]; Used = new byte[N]; NCnt = new int[N]; Vals = new float[N];
            for (int i = 0; i < N; i++) Vals[i] = float.NaN;
        }

        void BuildElements(List<int> raw, List<int> off, List<int> eids)
        {
            int ne = off.Count;
            List<int> conn = new List<int>(ne * 8), eid = new List<int>(ne);
            List<byte> nc = new List<byte>(ne), hm = new List<byte>(ne);
            int skipped = 0;
            int[] rest = new int[4];
            for (int e = 0; e < ne; e++)
            {
                int s = off[e];
                int cnt = (e + 1 < ne ? off[e + 1] : raw.Count) - s;
                int c, mids;
                if (cnt == 3 || cnt == 6) c = 3; else if (cnt == 4 || cnt == 8) c = 4; else { skipped++; continue; }
                mids = cnt - c;
                int[] ix = new int[cnt];
                bool ok = true;
                for (int k = 0; k < cnt; k++) { ix[k] = IndexOf(raw[s + k]); if (ix[k] < 0) { ok = false; break; } }
                if (!ok) { skipped++; continue; }
                int[] row = new int[8];
                for (int k = 0; k < 8; k++) row[k] = -1;
                for (int k = 0; k < c; k++) row[k] = ix[k];
                if (mids > 0)
                {
                    bool[] used = new bool[mids];
                    for (int k = 0; k < c; k++)
                    {
                        int a = ix[k], b = ix[(k + 1) % c];
                        double mx = (X[a] + X[b]) / 2.0, my = (Y[a] + Y[b]) / 2.0;
                        int best = -1; double bd = double.MaxValue;
                        for (int q = 0; q < mids; q++)
                        {
                            if (used[q]) continue;
                            int n = ix[c + q];
                            double d = Math.Sqrt((X[n] - mx) * (X[n] - mx) + (Y[n] - my) * (Y[n] - my));
                            if (d < bd) { bd = d; best = q; }
                        }
                        used[best] = true;
                        row[c + k] = ix[c + best];
                    }
                }
                for (int k = 0; k < 8; k++) conn.Add(row[k]);
                eid.Add(eids[e]); nc.Add((byte)c); hm.Add((byte)(mids > 0 ? 1 : 0));
            }
            NE = eid.Count; Skipped = skipped;
            Conn = conn.ToArray(); EId = eid.ToArray(); Nc = nc.ToArray(); HasMid = hm.ToArray();
            EHid = new byte[NE]; Bb = new double[NE * 4]; Cen = new double[NE * 2];
            double x0 = double.MaxValue, x1 = double.MinValue, y0 = double.MaxValue, y1 = double.MinValue;
            double sumw = 0; int nw = Math.Min(NE, 2000);
            for (int e = 0; e < NE; e++)
            {
                double bx0 = double.MaxValue, bx1 = double.MinValue, by0 = double.MaxValue, by1 = double.MinValue, cx = 0, cy = 0;
                for (int k = 0; k < 8; k++)
                {
                    int n = Conn[e * 8 + k];
                    if (n < 0) continue;
                    double bx = X[n], by = -Y[n];
                    if (bx < bx0) bx0 = bx; if (bx > bx1) bx1 = bx; if (by < by0) by0 = by; if (by > by1) by1 = by;
                    if (k < Nc[e]) { cx += bx; cy += by; }
                    Used[n] = 1; NCnt[n]++;
                }
                Bb[e * 4] = bx0; Bb[e * 4 + 1] = bx1; Bb[e * 4 + 2] = by0; Bb[e * 4 + 3] = by1;
                Cen[e * 2] = cx / Nc[e]; Cen[e * 2 + 1] = cy / Nc[e];
                if (bx0 < x0) x0 = bx0; if (bx1 > x1) x1 = bx1; if (by0 < y0) y0 = by0; if (by1 > y1) y1 = by1;
                if (e < nw) sumw += bx1 - bx0;
            }
            AvgEl = Math.Max(nw > 0 ? sumw / nw : 1.0, 1e-12);
            BoxX0 = x0; BoxX1 = x1; BoxY0 = y0; BoxY1 = y1;
        }

        // ---------------- degerler ----------------
        public void SetValues(int[] ids, double[] vals)
        {
            lock (sync)
            {
                for (int i = 0; i < N; i++) Vals[i] = float.NaN;
                for (int k = 0; k < ids.Length; k++) { int i = IndexOf(ids[k]); if (i >= 0) Vals[i] = (float)vals[k]; }
            }
        }
        public void ClearValues() { lock (sync) { for (int i = 0; i < N; i++) Vals[i] = float.NaN; } }
        public int[] Extremes()
        {
            int mi = -1, ma = -1; float lo = float.MaxValue, hi = float.MinValue;
            for (int i = 0; i < N; i++)
            {
                if (!NodeVisible(i)) continue;
                float v = Vals[i];
                if (float.IsNaN(v)) continue;
                if (v < lo) { lo = v; mi = i; }
                if (v > hi) { hi = v; ma = i; }
            }
            return new int[] { mi, ma };
        }

        // ---------------- gizleme ----------------
        void Recount()
        {
            Array.Clear(NCnt, 0, N);
            int h = 0;
            for (int e = 0; e < NE; e++)
            {
                if (EHid[e] != 0) { h++; continue; }
                for (int k = 0; k < 8; k++) { int n = Conn[e * 8 + k]; if (n >= 0) NCnt[n]++; }
            }
            HiddenElems = h;
        }
        public int HideElems(int[] idx)
        {
            int n = 0;
            lock (sync)
            {
                for (int q = 0; q < idx.Length; q++)
                {
                    int e = idx[q];
                    if (e < 0 || e >= NE || EHid[e] != 0) continue;
                    EHid[e] = 1; n++;
                    for (int k = 0; k < 8; k++) { int nd = Conn[e * 8 + k]; if (nd >= 0) NCnt[nd]--; }
                }
                HiddenElems += n;
            }
            return n;
        }
        // secili node'lara dokunmayan gorunen elemanlari gizler; gizlenen sayisi
        public int HideOthers(int[] keepNodes)
        {
            byte[] keep = new byte[N];
            for (int q = 0; q < keepNodes.Length; q++) { int i = keepNodes[q]; if (i >= 0 && i < N) keep[i] = 1; }
            List<int> hide = new List<int>();
            for (int e = 0; e < NE; e++)
            {
                if (EHid[e] != 0) continue;
                bool touch = false;
                for (int k = 0; k < 8; k++) { int n = Conn[e * 8 + k]; if (n >= 0 && keep[n] != 0) { touch = true; break; } }
                if (!touch) hide.Add(e);
            }
            return HideElems(hide.ToArray());
        }
        public void UnhideAll() { lock (sync) { Array.Clear(EHid, 0, NE); Recount(); } }
        public byte[] CopyEHid() { lock (sync) { return (byte[])EHid.Clone(); } }
        public void SetEHid(byte[] h) { lock (sync) { Buffer.BlockCopy(h, 0, EHid, 0, NE); Recount(); } }
        // gizlenmemis elemanlarin sinir kutusu (taban): x0, x1, y0, y1 (yoksa null)
        public double[] VisibleBox()
        {
            double[] b = null;
            for (int e = 0; e < NE; e++)
            {
                if (EHid[e] != 0) continue;
                double a0 = Bb[e * 4], a1 = Bb[e * 4 + 1], c0 = Bb[e * 4 + 2], c1 = Bb[e * 4 + 3];
                if (b == null) b = new double[] { a0, a1, c0, c1 };
                else { if (a0 < b[0]) b[0] = a0; if (a1 > b[1]) b[1] = a1; if (c0 < b[2]) b[2] = c0; if (c1 > b[3]) b[3] = c1; }
            }
            return b;
        }
        public int[] VisibleNodeIds()
        {
            List<int> r = new List<int>();
            for (int i = 0; i < N; i++) if (NodeVisible(i)) r.Add(Ids[i]);
            return r.ToArray();
        }
        public bool IsNodeVisible(int i) { return i >= 0 && i < N && NodeVisible(i); }

        // ---------------- secim / tiklama (ekran = M * taban + o) ----------------
        public int[] BoxNodes(double m00, double m01, double m10, double m11, double ox, double oy, double xa, double xb, double ya, double yb)
        {
            List<int> r = new List<int>();
            for (int i = 0; i < N; i++)
            {
                if (!NodeVisible(i)) continue;
                double bx = X[i], by = -Y[i];
                double sx = m00 * bx + m01 * by + ox;
                if (sx < xa || sx > xb) continue;
                double sy = m10 * bx + m11 * by + oy;
                if (sy < ya || sy > yb) continue;
                r.Add(i);
            }
            return r.ToArray();
        }
        // maske ait (secili) node'lar arasinda en yakin; mask = 0 -> hepsi
        public int PickNode(double m00, double m01, double m10, double m11, double ox, double oy, double mx, double my, double rad, int mask)
        {
            double best = rad; int bi = -1;
            for (int i = 0; i < N; i++)
            {
                if (!NodeVisible(i)) continue;
                if (mask != 0 && (Sel[i] & mask) == 0) continue;
                double bx = X[i], by = -Y[i];
                double dx = m00 * bx + m01 * by + ox - mx;
                if (dx > rad || dx < -rad) continue;
                double dy = m10 * bx + m11 * by + oy - my;
                if (dy > rad || dy < -rad) continue;
                double d = Math.Sqrt(dx * dx + dy * dy);
                if (d <= best) { best = d; bi = i; }
            }
            return bi;
        }
        public int[] ElemsInBox(double m00, double m01, double m10, double m11, double ox, double oy, double xa, double xb, double ya, double yb)
        {
            List<int> r = new List<int>();
            for (int e = 0; e < NE; e++)
            {
                if (EHid[e] != 0) continue;
                double cx = Cen[e * 2], cy = Cen[e * 2 + 1];
                double sx = m00 * cx + m01 * cy + ox, sy = m10 * cx + m11 * cy + oy;
                if (sx >= xa && sx <= xb && sy >= ya && sy <= yb) r.Add(e);
            }
            return r.ToArray();
        }
        public int PickElem(double m00, double m01, double m10, double m11, double ox, double oy, double mx, double my)
        {
            double det = m00 * m11 - m01 * m10;
            if (Math.Abs(det) < 1e-12) det = 1e-12;
            double dx = mx - ox, dy = my - oy;
            double X0 = (m11 * dx - m01 * dy) / det, Y0 = (-m10 * dx + m00 * dy) / det;
            int best = -1;
            for (int e = 0; e < NE; e++)
            {
                if (EHid[e] != 0) continue;
                if (X0 < Bb[e * 4] || X0 > Bb[e * 4 + 1] || Y0 < Bb[e * 4 + 2] || Y0 > Bb[e * 4 + 3]) continue;
                int c = Nc[e];
                bool inside = false;
                int j = c - 1;
                for (int i = 0; i < c; i++)
                {
                    double xi = X[Conn[e * 8 + i]], yi = -Y[Conn[e * 8 + i]], xj = X[Conn[e * 8 + j]], yj = -Y[Conn[e * 8 + j]];
                    if ((yi > Y0) != (yj > Y0) && X0 < (xj - xi) * (Y0 - yi) / (yj - yi) + xi) inside = !inside;
                    j = i;
                }
                if (inside) best = e;
            }
            return best;
        }
        // ekrandaki (pencere icindeki) gorunen node'lar arasinda degeri olanlarin min / max node indeksi
        public int[] MinMaxOnScreen(double m00, double m01, double m10, double m11, double ox, double oy, double W, double H)
        {
            int mi = -1, ma = -1; float lo = float.MaxValue, hi = float.MinValue;
            for (int i = 0; i < N; i++)
            {
                if (!NodeVisible(i)) continue;
                float v = Vals[i];
                if (float.IsNaN(v)) continue;
                double bx = X[i], by = -Y[i];
                double sx = m00 * bx + m01 * by + ox;
                if (sx < 0 || sx > W) continue;
                double sy = m10 * bx + m11 * by + oy;
                if (sy < 0 || sy > H) continue;
                if (v < lo) { lo = v; mi = i; }
                if (v > hi) { hi = v; ma = i; }
            }
            return new int[] { mi, ma };
        }

        // ---------------- sekil fonksiyonlari ----------------
        static readonly double[] CSx = { -1, 1, 1, -1 }, CSy = { -1, -1, 1, 1 };
        static readonly double[] MSx = { 0, 1, 0, -1 }, MSy = { -1, 0, 1, 0 };
        // w: kose (nc) + orta (nc) agirliklari
        static void Shape(int nc, bool mid, double a, double b, double[] w)
        {
            if (nc == 3)
            {
                double L0 = 1.0 - a - b, L1 = a, L2 = b;
                if (mid)
                {
                    w[0] = L0 * (2 * L0 - 1); w[1] = L1 * (2 * L1 - 1); w[2] = L2 * (2 * L2 - 1);
                    w[3] = 4 * L0 * L1; w[4] = 4 * L1 * L2; w[5] = 4 * L2 * L0;
                }
                else { w[0] = L0; w[1] = L1; w[2] = L2; }
                return;
            }
            if (mid)
            {
                for (int k = 0; k < 4; k++) w[k] = 0.25 * (1 + a * CSx[k]) * (1 + b * CSy[k]) * (a * CSx[k] + b * CSy[k] - 1);
                for (int k = 0; k < 4; k++)
                    w[4 + k] = MSx[k] == 0 ? 0.5 * (1 - a * a) * (1 + b * MSy[k]) : 0.5 * (1 + a * MSx[k]) * (1 - b * b);
            }
            else for (int k = 0; k < 4; k++) w[k] = 0.25 * (1 + a * CSx[k]) * (1 + b * CSy[k]);
        }

        // ---------------- cizim ----------------
        int[] cpx; int cW, cH, cBg; double cVmin, cRng; int[] cPal; bool cGray;
        static readonly int LightGray = unchecked((int)0xFFD3D3D3);

        int ColorOf(double v)
        {
            if (cRng <= 0) return cPal[128];
            int ci = (int)(255.0 * (v - cVmin) / cRng);
            return cPal[ci < 0 ? 0 : (ci > 255 ? 255 : ci)];
        }

        // ucgen: tepe degerleri v0..v2 (gray ise duz gri); baricentrik dolgu (kenarda ufak tolerans)
        void FillTri(double x0, double y0, double v0, double x1, double y1, double v1, double x2, double y2, double v2, bool gray)
        {
            double minx = Math.Min(x0, Math.Min(x1, x2)), maxx = Math.Max(x0, Math.Max(x1, x2));
            double miny = Math.Min(y0, Math.Min(y1, y2)), maxy = Math.Max(y0, Math.Max(y1, y2));
            int ix0 = (int)Math.Floor(minx - 0.5); int ix1 = (int)Math.Ceiling(maxx + 0.5);
            int iy0 = (int)Math.Floor(miny - 0.5); int iy1 = (int)Math.Ceiling(maxy + 0.5);
            if (ix0 < 0) ix0 = 0; if (iy0 < 0) iy0 = 0; if (ix1 > cW - 1) ix1 = cW - 1; if (iy1 > cH - 1) iy1 = cH - 1;
            if (ix1 < ix0 || iy1 < iy0) return;
            double den = (y1 - y2) * (x0 - x2) + (x2 - x1) * (y0 - y2);
            if (Math.Abs(den) < 1e-12)
            {
                int col = gray ? LightGray : ColorOf((v0 + v1 + v2) / 3.0);
                int cx = (int)((x0 + x1 + x2) / 3.0), cy = (int)((y0 + y1 + y2) / 3.0);
                if (cx >= 0 && cx < cW && cy >= 0 && cy < cH) cpx[cy * cW + cx] = col;
                return;
            }
            double id = 1.0 / den;
            const double tol = -0.02;
            int flat = gray ? LightGray : 0;
            for (int y = iy0; y <= iy1; y++)
            {
                double py = y + 0.5;
                int row = y * cW;
                for (int x = ix0; x <= ix1; x++)
                {
                    double pxx = x + 0.5;
                    double w0 = ((y1 - y2) * (pxx - x2) + (x2 - x1) * (py - y2)) * id;
                    if (w0 < tol) continue;
                    double w1 = ((y2 - y0) * (pxx - x2) + (x0 - x2) * (py - y2)) * id;
                    if (w1 < tol) continue;
                    double w2 = 1.0 - w0 - w1;
                    if (w2 < tol) continue;
                    cpx[row + x] = gray ? flat : ColorOf(w0 * v0 + w1 * v1 + w2 * v2);
                }
            }
        }

        void Line(double xa, double ya, double xb, double yb, int col)
        {
            int x0 = (int)xa, y0 = (int)ya, x1 = (int)xb, y1 = (int)yb;
            if (Math.Abs(x0) > 20000 || Math.Abs(y0) > 20000 || Math.Abs(x1) > 20000 || Math.Abs(y1) > 20000) return;
            int dx = Math.Abs(x1 - x0), dy = -Math.Abs(y1 - y0), sx = x0 < x1 ? 1 : -1, sy = y0 < y1 ? 1 : -1, err = dx + dy;
            while (true)
            {
                if (x0 >= 0 && x0 < cW && y0 >= 0 && y0 < cH) cpx[y0 * cW + x0] = col;
                if (x0 == x1 && y0 == y1) break;
                int e2 = 2 * err;
                if (e2 >= dy) { err += dy; x0 += sx; }
                if (e2 <= dx) { err += dx; y0 += sy; }
            }
        }

        void Square(double x, double y, int r, int col)
        {
            int ix = (int)x, iy = (int)y;
            for (int yy = iy - r; yy <= iy + r; yy++)
            {
                if (yy < 0 || yy >= cH) continue;
                for (int xx = ix - r; xx <= ix + r; xx++) if (xx >= 0 && xx < cW) cpx[yy * cW + xx] = col;
            }
        }

        public int Render(int[] px, int W, int H, double m00, double m01, double m10, double m11, double ox, double oy, int bg, int[] palette,
                          bool hasT, double vmin, double vmax, bool meshOn, bool nodesOn, bool bigSel, bool showSel)
        {
            lock (sync)
            {
                cpx = px; cW = W; cH = H; cBg = bg; cPal = palette; cVmin = vmin; cRng = vmax - vmin;
                int total = W * H;
                for (int i = 0; i < total; i++) px[i] = bg;
                if (SX == null || SX.Length != N) { SX = new double[N]; SY = new double[N]; }
                for (int i = 0; i < N; i++)
                {
                    double bx = X[i], by = -Y[i];
                    SX[i] = m00 * bx + m01 * by + ox; SY[i] = m10 * bx + m11 * by + oy;
                }
                double sc = Math.Sqrt(Math.Abs(m00 * m11 - m01 * m10));
                int drawn = 0;
                double[] w = new double[8];
                double[] ex = new double[8], ey = new double[8], ev = new double[8];
                int meshCol = unchecked((int)0xFF464646);
                for (int e = 0; e < NE; e++)
                {
                    if (EHid[e] != 0) continue;
                    int nc = Nc[e];
                    bool mid = HasMid[e] != 0;
                    int nn = mid ? 2 * nc : nc;
                    double minx = double.MaxValue, maxx = double.MinValue, miny = double.MaxValue, maxy = double.MinValue;
                    bool ok = hasT;
                    for (int k = 0; k < nn; k++)
                    {
                        int n = Conn[e * 8 + k];
                        double sx = SX[n], sy = SY[n];
                        ex[k] = sx; ey[k] = sy;
                        if (sx < minx) minx = sx; if (sx > maxx) maxx = sx; if (sy < miny) miny = sy; if (sy > maxy) maxy = sy;
                        if (ok) { float v = Vals[n]; if (float.IsNaN(v)) ok = false; else ev[k] = v; }
                    }
                    if (maxx < -2 || minx > W + 2 || maxy < -2 || miny > H + 2) continue;
                    drawn++;
                    double size = Math.Max(maxx - minx, maxy - miny);
                    bool gray = !ok;
                    if (size < 1.0)
                    {
                        double cv = 0;
                        if (ok) { Shape(nc, mid, nc == 3 ? 1.0 / 3.0 : 0.0, nc == 3 ? 1.0 / 3.0 : 0.0, w); for (int k = 0; k < nn; k++) cv += w[k] * ev[k]; }
                        int cx = (int)((minx + maxx) * 0.5), cy = (int)((miny + maxy) * 0.5);
                        if (cx >= 0 && cx < W && cy >= 0 && cy < H) px[cy * W + cx] = gray ? LightGray : ColorOf(cv);
                        continue;
                    }
                    if (size < 6.0)
                    {
                        double cv = 0;
                        if (ok) { Shape(nc, mid, nc == 3 ? 1.0 / 3.0 : 0.0, nc == 3 ? 1.0 / 3.0 : 0.0, w); for (int k = 0; k < nn; k++) cv += w[k] * ev[k]; }
                        FillTri(ex[0], ey[0], cv, ex[1], ey[1], cv, ex[2], ey[2], cv, gray);
                        if (nc == 4) FillTri(ex[0], ey[0], cv, ex[2], ey[2], cv, ex[3], ey[3], cv, gray);
                    }
                    else if (nc == 3 && !mid)
                    {
                        FillTri(ex[0], ey[0], ev[0], ex[1], ey[1], ev[1], ex[2], ey[2], ev[2], gray);
                    }
                    else
                    {
                        int n = mid ? (size < 20.0 ? 2 : 4) : (nc == 3 ? 1 : (size < 20.0 ? 2 : 4));
                        // parametrik izgara: ucgen icin (i,j) i+j<=n, dortgen icin (i,j) 0..n
                        int m = n + 1;
                        double[] gx = new double[m * m], gy = new double[m * m], gv = new double[m * m];
                        bool[] has = new bool[m * m];
                        for (int i = 0; i < m; i++)
                            for (int j = 0; j < m; j++)
                            {
                                if (nc == 3 && i + j > n) continue;
                                double a, b;
                                if (nc == 3) { a = (double)i / n; b = (double)j / n; }
                                else { a = -1.0 + 2.0 * i / n; b = -1.0 + 2.0 * j / n; }
                                Shape(nc, mid, a, b, w);
                                double sx = 0, sy = 0, sv = 0;
                                for (int k = 0; k < nn; k++) { sx += w[k] * ex[k]; sy += w[k] * ey[k]; if (ok) sv += w[k] * ev[k]; }
                                gx[i * m + j] = sx; gy[i * m + j] = sy; gv[i * m + j] = sv; has[i * m + j] = true;
                            }
                        for (int i = 0; i < n; i++)
                            for (int j = 0; j < n; j++)
                            {
                                if (nc == 3)
                                {
                                    if (i + j > n - 1) continue;
                                    int p00 = i * m + j, p10 = (i + 1) * m + j, p01 = i * m + j + 1;
                                    FillTri(gx[p00], gy[p00], gv[p00], gx[p10], gy[p10], gv[p10], gx[p01], gy[p01], gv[p01], gray);
                                    if (i + j < n - 1)
                                    {
                                        int p11 = (i + 1) * m + j + 1;
                                        FillTri(gx[p10], gy[p10], gv[p10], gx[p11], gy[p11], gv[p11], gx[p01], gy[p01], gv[p01], gray);
                                    }
                                }
                                else
                                {
                                    int p00 = i * m + j, p10 = (i + 1) * m + j, p01 = i * m + j + 1, p11 = (i + 1) * m + j + 1;
                                    FillTri(gx[p00], gy[p00], gv[p00], gx[p10], gy[p10], gv[p10], gx[p11], gy[p11], gv[p11], gray);
                                    FillTri(gx[p00], gy[p00], gv[p00], gx[p11], gy[p11], gv[p11], gx[p01], gy[p01], gv[p01], gray);
                                }
                            }
                    }
                    if (meshOn && size >= 4.0)
                    {
                        for (int k = 0; k < nc; k++)
                        {
                            int kn = (k + 1) % nc;
                            if (mid)
                            {
                                Line(ex[k], ey[k], ex[nc + k], ey[nc + k], meshCol);
                                Line(ex[nc + k], ey[nc + k], ex[kn], ey[kn], meshCol);
                            }
                            else Line(ex[k], ey[k], ex[kn], ey[kn], meshCol);
                        }
                    }
                }
                if (nodesOn && AvgEl * sc >= 4.0)
                {
                    int dark = unchecked((int)0xFF282828), cnt = 0;
                    for (int i = 0; i < N; i++) if (NodeVisible(i) && SX[i] >= -10 && SX[i] <= W + 10 && SY[i] >= -10 && SY[i] <= H + 10) { cnt++; if (cnt > 20000) break; }
                    if (cnt <= 20000)
                        for (int i = 0; i < N; i++) if (NodeVisible(i) && SX[i] >= -10 && SX[i] <= W + 10 && SY[i] >= -10 && SY[i] <= H + 10) Square(SX[i], SY[i], 1, dark);
                }
                if (bigSel)
                {
                    int yellow = unchecked((int)0xFFFFD700), green = unchecked((int)0xFF3CC85A);
                    for (int i = 0; i < N; i++)
                    {
                        byte sf = Sel[i];
                        if (sf == 0 || !NodeVisible(i)) continue;
                        if ((sf & 1) != 0) Square(SX[i], SY[i], 2, yellow);
                        else if (showSel && (sf & 2) != 0) Square(SX[i], SY[i], 2, green);
                    }
                }
                return drawn;
            }
        }
    }
}
'''

_NATIVE = {"type": None, "mesh2d": None, "tried": False, "err": u"", "lock": threading.Lock()}
SEL_PEND, SEL_LIST = 1, 2


def native_cloud_type():
    """C# motorunu (SC3.Cloud / SC3.Mesh2D) bir kez derler; olmazsa None (uygulama Python yolunu kullanir)."""
    st = _NATIVE
    with st["lock"]:
        if st["tried"]:
            return st["type"]
        st["tried"] = True
        try:
            clr.AddReference("System")
            from Microsoft.CSharp import CSharpCodeProvider
            from System.CodeDom.Compiler import CompilerParameters
            cp = CompilerParameters()
            cp.GenerateInMemory = True
            cp.GenerateExecutable = False
            cp.CompilerOptions = "/optimize+"
            cp.ReferencedAssemblies.Add("System.dll")
            cp.ReferencedAssemblies.Add("System.Core.dll")
            res = CSharpCodeProvider().CompileAssemblyFromSource(cp, SC3_NATIVE_SRC)
            if res.Errors.HasErrors:
                st["err"] = u"; ".join(unicode(e) for e in list(res.Errors)[:3])
                return None
            clr.AddReference(res.CompiledAssembly)
            import SC3
            st["type"] = SC3.Cloud
            st["mesh2d"] = SC3.Mesh2D
        except Exception as ex:
            st["err"] = unicode(ex)
            st["type"] = None
            st["mesh2d"] = None
        return st["type"]


class NativeNodeTable(object):
    """C# motorundaki node'lar (NodeTable ile ayni arayuz): nodes[nid] -> (X, Y, Z)."""
    def __init__(self, cloud):
        self.c = cloud
        self.X, self.Y, self.Z, self.ids = cloud.X, cloud.Y, cloud.Z, cloud.Ids
        self.first = cloud.First
        self.contig = bool(cloud.Contig)

    def index(self, nid):
        return self.c.IndexOf(nid)

    def id_of(self, i):
        return i + self.first if self.contig else self.ids[i]

    def __getitem__(self, nid):
        i = self.c.IndexOf(nid)
        if i < 0:
            raise KeyError(nid)
        return (self.X[i], self.Y[i], self.Z[i])

    def __contains__(self, nid):
        return self.c.IndexOf(nid) >= 0

    def __len__(self):
        return self.c.N

    def keys(self):
        return self.ids


def cloud_model_native(cloud):
    nt = NativeNodeTable(cloud)
    return {"is3d": True, "nodes": nt, "nn": cloud.N, "ne": 0, "used": AllNodes(nt), "elems": ElemCount(0), "skipped": 0,
            "counts": {}, "bad": {}, "cloud": cloud,
            "box3": (cloud.MinX, cloud.MaxX, cloud.MinY, cloud.MaxY, cloud.MinZ, cloud.MaxZ)}


class Native2DNodes(object):
    """2D mesh'in node'lari (C# motorunda): nodes[nid] -> (X, Y, 0)."""
    def __init__(self, mesh):
        self.m = mesh
        self.X, self.Y, self.ids = mesh.X, mesh.Y, mesh.Ids

    def index(self, nid):
        return self.m.IndexOf(nid)

    def id_of(self, i):
        return self.ids[i]

    def __getitem__(self, nid):
        i = self.m.IndexOf(nid)
        if i < 0:
            raise KeyError(nid)
        return (self.X[i], self.Y[i], 0.0)

    def __contains__(self, nid):
        return self.m.IndexOf(nid) >= 0

    def __len__(self):
        return self.m.N

    def keys(self):
        return self.ids


class UsedNodes2D(object):
    """model['used']: en az bir elemana ait node'lar."""
    def __init__(self, mesh):
        self.m = mesh

    def __contains__(self, nid):
        i = self.m.IndexOf(nid)
        return i >= 0 and self.m.Used[i] != 0

    def __len__(self):
        return sum(1 for v in self.m.Used if v)


class NodeBase2D(object):
    """view.dn (2D): node -> taban koordinati (X, -Y)."""
    def __init__(self, mesh):
        self.m = mesh

    def get(self, nid, default=None):
        m = self.m
        i = m.IndexOf(nid)
        if i < 0 or m.Used[i] == 0:
            return default
        return (m.X[i], -m.Y[i])

    def __getitem__(self, nid):
        p = self.get(nid)
        if p is None:
            raise KeyError(nid)
        return p

    def __contains__(self, nid):
        return self.get(nid) is not None

    def keys(self):
        return self.m.VisibleNodeIds()

    def items(self):
        m = self.m
        return [(m.Ids[i], (m.X[i], -m.Y[i])) for i in range(m.N) if m.Used[i]]


class HidNodes2D(object):
    """view.hid_nodes (2D): hicbir gorunen elemana ait olmayan node'lar."""
    def __init__(self, mesh):
        self.m = mesh

    def __contains__(self, nid):
        i = self.m.IndexOf(nid)
        return i < 0 or not self.m.IsNodeVisible(i)


def mesh_model_native(mesh):
    return {"mesh2": mesh, "nodes": Native2DNodes(mesh), "elems": ElemCount(mesh.NE), "used": UsedNodes2D(mesh), "skipped": mesh.Skipped,
            "box": (mesh.BoxX0, mesh.BoxX1, mesh.BoxY0, mesh.BoxY1)}


class ScalarMap(dict):
    """{node: deger}; .ids / .vals: ayni veri liste olarak (C# motoruna hizli aktarim icin)."""
    pass


class NSet(object):
    """Motordaki secim bayraklarini (bit) Python kumesi gibi gosterir; milyonlarca node'da bile hizli (islemler derlenmis kodda)."""
    def __init__(self, view, mask):
        self.v, self.mask = view, mask

    def __len__(self):
        return self.v.engine.SelCount(self.mask)

    def __nonzero__(self):
        return self.v.engine.SelCount(self.mask) > 0
    __bool__ = __nonzero__

    def __contains__(self, nid):
        e = self.v.engine
        i = e.IndexOf(nid)
        return i >= 0 and e.SelHas(self.mask, i)

    def __iter__(self):
        return iter(list(self.v.engine.SelIds(self.mask)))

    def add(self, nid):
        e = self.v.engine
        e.SelSetOne(self.mask, e.IndexOf(nid), True)

    def discard(self, nid):
        e = self.v.engine
        e.SelSetOne(self.mask, e.IndexOf(nid), False)

    def clear(self):
        self.v.engine.SelClear(self.mask)

    def smallest(self, k):
        return list(self.v.engine.SelSmallestIds(self.mask, k))

    def _idx(self, other):
        e = self.v.engine
        return e.IndexMany(System.Array[int](list(other)))

    def __ior__(self, other):
        e = self.v.engine
        if isinstance(other, NSet):
            e.SelOrMask(self.mask, other.mask)
        else:
            e.SelAddIdx(self.mask, self._idx(other), 0)
        return self

    def __isub__(self, other):
        e = self.v.engine
        if isinstance(other, NSet):
            e.SelAndNotMask(self.mask, other.mask)
        else:
            e.SelRemoveIdx(self.mask, self._idx(other))
        return self

    def __or__(self, other):
        return set(self) | set(other)

    def __and__(self, other):
        return set(self) & set(other)

    def __sub__(self, other):
        return set(self) - set(other)

    def copy(self):
        return set(self)


class NativeSelMixin(object):
    """Hizli gorunumlerde secim (sari = gecici, yesil = listede) motordaki bit bayraklariyla tutulur; view.sel / view.pend kume gibi kullanilir."""
    nat_sel = False
    engine = None

    def _nat_init(self, engine):
        self.engine = engine
        self._nset_pend = NSet(self, SEL_PEND)
        self._nset_sel = NSet(self, SEL_LIST)
        self.nat_sel = True
        engine.SelClear(3)
        self._vals_for = None

    def _nat_assign(self, nset, value):
        if value is nset:
            return
        e = self.engine
        e.SelClear(nset.mask)
        if isinstance(value, NSet):
            e.SelOrMask(nset.mask, value.mask)
        elif value:
            e.SelAddIdx(nset.mask, e.IndexMany(System.Array[int](list(value))), 0)

    def _get_pend(self):
        return self._nset_pend if self.nat_sel else self.__dict__.get("_pend_plain", set())

    def _get_sel(self):
        return self._nset_sel if self.nat_sel else self.__dict__.get("_sel_plain", set())

    pend = property(_get_pend)
    sel = property(_get_sel)

    def __setattr__(self, name, value):
        """view.pend = ... / view.sel = ... atamalari (kume gibi): hizli modda motordaki bayraklara yazilir."""
        if name == "pend" or name == "sel":
            if self.nat_sel:
                self._nat_assign(self._nset_pend if name == "pend" else self._nset_sel, value)
            else:
                self.__dict__["_pend_plain" if name == "pend" else "_sel_plain"] = value
            return
        super(NativeSelMixin, self).__setattr__(name, value)

    def sel_snapshot(self):
        return self.engine.SelSnapshot()

    def sel_restore(self, snap):
        self.engine.SelRestore(snap)

    def sel_select_all(self):
        self.engine.SelClear(SEL_PEND)
        self.engine.SelAllVisible(SEL_PEND, SEL_LIST)

    def sel_indices(self):
        return self.engine.SelIndices(SEL_PEND | SEL_LIST)

    def ensure_values(self, tmap):
        """Alan degerlerini (renk) motora aktarir (ayni veri icin bir kez)."""
        if self.engine is None or tmap is self._vals_for:
            return
        ids = getattr(tmap, "ids", None)
        vals = getattr(tmap, "vals", None)
        if ids is None:
            ids, vals = list(tmap.keys()), list(tmap.values())
        self.engine.SetValues(System.Array[int](ids), System.Array[float](vals))
        self._vals_for = tmap

    def _palette(self):
        if self.__dict__.get("_pal") is None:
            vals = []
            for i in range(256):
                r, g, b = jet_rgb(i / 255.0)
                a = (255 << 24) | (r << 16) | (g << 8) | b
                vals.append(a - (1 << 32) if a >= (1 << 31) else a)
            self._pal = System.Array[int](vals)
        return self._pal

    def _copy_bitmap(self, px, W, H):
        bmp = Bitmap(W, H, PixelFormat.Format32bppArgb)
        bd = bmp.LockBits(Rectangle(0, 0, W, H), ImageLockMode.WriteOnly, PixelFormat.Format32bppArgb)
        Marshal.Copy(px, 0, bd.Scan0, W * H)
        bmp.UnlockBits(bd)
        return bmp


class FastCloudView(NativeSelMixin, PointCloudView):
    """PointCloudView'in C# motorlu hali: cizim, kutu / tiklama secimi, secim bayraklari, gizleme derlenmis kodda; her karede TUM node'lar cizilir
    (seyreltme yok). Motor yoksa (model 'cloud' icermiyorsa) ust sinifin Python yolu kullanilir."""
    def __init__(self, owner):
        PointCloudView.__init__(self, owner)
        self.cloud = None
        self._pal = None
        self._px = None
        self._vals_for = None

    def set_model(self, model):
        PointCloudView.set_model(self, model)
        self.cloud = model.get("cloud")
        self._vals_for = None
        if self.cloud is not None:
            self._nat_init(self.cloud)
            self.nhid = self.cloud.Hidden                   # byte[]: Python'dan nhid[i] ile okunabilir
            self.n_hidden = 0
        else:
            self.nat_sel = False
            self.engine = None

    def set_field(self, tmap, vmin, vmax, off, title, fac=1.0, ext=None):
        self.ensure_values(tmap)
        PointCloudView.set_field(self, tmap, vmin, vmax, off, title, fac, ext)

    def clear_field(self):
        if self.cloud is not None:
            self.cloud.ClearValues()
            self._vals_for = None
        PointCloudView.clear_field(self)

    # ----- cizim -----
    def _need_render3(self):
        if self.cloud is None:
            return PointCloudView._need_render3(self)
        if self.model is None or self.Width < 10 or self.Height < 10:
            return None
        same = (self._img3 is not None and self._img3_sig == self._vsig3() and self._img3_cgen == self._content_gen
                and self._img3_hv == self._hide_ver)
        seldirty = self._img3_selv != self._selsig() and (self._img3_big or self._bigsel())          # sayim sadece secim degistiyse
        if same and not seldirty and (self._img3_mode == "full" or self._moving):
            return None
        return "preview" if self._moving else "full"           # 'preview' = hareket sirasindaki kare (eski olsa da gosterilir); yine TUM node'lar cizilir

    def _run_job3(self, job):
        if self.cloud is None:
            return PointCloudView._run_job3(self, job)
        bmp = None
        try:
            c = self.cloud
            W, H = job["W"], job["H"]
            if self._px is None or len(self._px) != W * H:
                self._px = System.Array.CreateInstance(System.Int32, W * H)
            nn = self.model["nn"]
            sx, sy = job["pan"]
            cx, cy, cz = job["ctr"]
            n = c.Render(self._px, W, H, System.Array[float](list(job["rot"])), job["s"], sx, sy, cx, cy, cz,
                         3 if nn <= 250000 else 2, self.bg.ToArgb(), self._palette(), bool(job["has_t"]), job["vmin"], job["vmax"],
                         bool(job["bigsel"]), bool(job["showsel"]), job["rad"], bool(job["track"] and job["has_t"]))
            if self._content_gen != job["cgen"] or (job["mode"] == "full" and self._view_gen != job["vgen"]):
                return                                              # arada veri / (duran gorunumde) gorunum degisti: bu kare atilir
            bmp = self._copy_bitmap(self._px, W, H)
            mm = None
            if job["track"] and job["has_t"] and c.LastMinIdx >= 0 and c.LastMaxIdx >= 0:
                nt = self.model["nodes"]
                a, b = nt.id_of(c.LastMinIdx), nt.id_of(c.LastMaxIdx)
                get = job["tmap"].get
                mm = (a, get(a), b, get(b))
            self._res3 = (job, bmp, {"mm": mm, "n": n, "hv": job["hv"]})
            bmp = None
        except Exception:
            pass
        finally:
            if bmp is not None:
                bmp.Dispose()
            job["alive"] = False

    # ----- secim (cizimden bagimsiz, tum node'lar; secim bayraklari motorda) -----
    def _vp(self):
        sx, sy = self.pan3
        cx, cy, cz = self.ctr3
        return (System.Array[float](list(self.R)), self.s3, float(self.Width), float(self.Height), sx, sy, cx, cy, cz)

    def _nodes_in_box(self, xa, xb, ya, yb):
        if self.cloud is None:
            return PointCloudView._nodes_in_box(self, xa, xb, ya, yb)
        R, s, W, H, px, py, cx, cy, cz = self._vp()
        return self.cloud.BoxSelect(R, s, W, H, px, py, cx, cy, cz, xa, xb, ya, yb)             # int[] node indeksleri

    def _box_action(self, xa, xb, ya, yb):
        if self.cloud is None:
            return PointCloudView._box_action(self, xa, xb, ya, yb)
        inside = self._nodes_in_box(xa, xb, ya, yb)
        if self.hide_mode:
            self._hide_elems(inside)
            return
        snap = self.cloud.SelSnapshot()
        n = self.cloud.SelRemoveIdx(3, inside) if self.unsel_mode else self.cloud.SelAddIdx(SEL_PEND, inside, 3)
        if n:
            self.owner.push_undo_data(snap)
            self.owner.selection_changed()

    def _find_node(self, mx, my, only=None):
        if self.cloud is None or only is not None:
            return PointCloudView._find_node(self, mx, my, only)
        R, s, W, H, px, py, cx, cy, cz = self._vp()
        i = self.cloud.Pick(R, s, W, H, px, py, cx, cy, cz, mx, my, 12.0)
        return None if i < 0 else self.model["nodes"].id_of(i)

    # ----- gizleme (motorda) -----
    def _push_snapshot(self):
        self.hide_stack.append(self.cloud.CopyHidden())
        if len(self.hide_stack) > 25:
            del self.hide_stack[0]

    def _hide_elems(self, idxs):
        if self.cloud is None:
            return PointCloudView._hide_elems(self, idxs)
        arr = idxs if hasattr(idxs, "Length") else System.Array[int](list(idxs))
        self._push_snapshot()
        n = self.cloud.HideIdx(arr)
        if n == 0:
            self.hide_stack.pop()
            return
        self.n_hidden += n
        self.owner.push_hide()
        self._commit_hide()
        self.owner.hidden_changed("hide")

    def undo_hide(self):
        if self.cloud is None:
            return PointCloudView.undo_hide(self)
        if not self.hide_stack:
            return
        self.cloud.SetHidden(self.hide_stack.pop())
        self.n_hidden = self.cloud.CountHidden()
        self._commit_hide()
        self.owner.hidden_changed("undo")

    def unhide_all(self):
        if self.cloud is None:
            return PointCloudView.unhide_all(self)
        if self.n_hidden == 0:
            return
        self.cloud.ClearHidden()
        self.n_hidden = 0
        self.hide_stack = []
        self.owner.hides_cleared()
        self._commit_hide()
        self.owner.hidden_changed("unhide")

    def hide_others_idx(self, kidx):
        """Verilen node indekslerine (secili) dokunmayanlari gizler; gizlenen sayisi."""
        old = self.n_hidden
        self._push_snapshot()
        self.n_hidden = self.cloud.HideAllExcept(kidx)
        self.owner.push_hide()
        self._commit_hide()
        self.owner.hidden_changed("hide")
        return max(0, self.n_hidden - old)

    def hide_others(self, keep):
        if self.cloud is None:
            return PointCloudView.hide_others(self, keep)
        kidx = self.cloud.IndexMany(System.Array[int](list(keep)))
        if not any(i >= 0 for i in kidx):
            return 0
        return self.hide_others_idx(kidx)

    def _visible_box3(self):
        if self.cloud is None or self.n_hidden == 0:
            return PointCloudView._visible_box3(self)
        b = self.cloud.VisibleBounds()
        return self.model["box3"] if b is None else tuple(b)


class FastSceneView(NativeSelMixin, SceneView):
    """2D gorunumun C# motorlu hali: elemanlar (konturlu), mesh cizgileri, node noktalari motorda cizilir; gizleme, secim, tiklama motorda.
    Her karede yeniden cizilir (mip / onbellek yok). Motor yoksa (model 'mesh2' icermiyorsa) eski Python yolu kullanilir."""
    def __init__(self, owner):
        SceneView.__init__(self, owner)
        self.mesh = None
        self.sel_version = 0
        self._img2 = None
        self._img2_sig = None
        self._img2_mode = None
        self._img2_cgen = -1
        self._img2_xf = None
        self._img2_selv = None
        self._img2_big = False
        self._px = None
        self._pal = None
        self._vals_for = None

    def _selsig(self):
        return (self.sel_version, self.listed_on)

    def _bigsel2(self):
        return (len(self.pend) + (len(self.sel) if self.listed_on else 0)) > BIG_SEL

    def set_model(self, model):
        m = model.get("mesh2")
        self.mesh = m
        if m is None:
            self.nat_sel = False
            self.engine = None
            return SceneView.set_model(self, model)
        self._content_gen += 1
        self._view_gen += 1
        self.model = model
        self.dn = NodeBase2D(m)
        self.avg_el = m.AvgEl
        self._cell = self.avg_el
        self._ngrid = self._egrid = None
        self._grid_thread = None
        self._nat_init(m)
        self.vis_nodes = None
        self.hide_stack = []
        self.hid_nodes = HidNodes2D(m)
        self.node_cnt = None
        self.n_hidden = 0
        self.probes = []
        self.probe_text = {}
        self.has_t = False
        self._mm = None
        self._mm_sig = None
        self._mip_ok = False
        for b in (self._cache, self._mip, self._img2):
            if b is not None:
                b.Dispose()
        self._cache = self._mip = self._img2 = None
        self._img2_sig = None
        self._img2_selv = None
        self._cache_cgen = self._mip_cgen = -1
        self._cache_sig = None
        self._result = None
        self._anim = None
        self.R2 = (1.0, 0.0, 0.0, 0.0, 1.0, 0.0, 0.0, 0.0, 1.0)
        self._fit_pending = True
        self.Invalidate()

    def set_field(self, tmap, vmin, vmax, off, title, fac=1.0, ext=None):
        if self.mesh is not None:
            self.ensure_values(tmap)
        SceneView.set_field(self, tmap, vmin, vmax, off, title, fac, ext)

    def clear_field(self):
        if self.mesh is not None:
            self.mesh.ClearValues()
            self._vals_for = None
        SceneView.clear_field(self)

    # ----- gizleme (motorda) -----
    def _push_snapshot(self):
        self.hide_stack.append(self.mesh.CopyEHid())
        if len(self.hide_stack) > 25:
            del self.hide_stack[0]

    def _hide_elems(self, els):
        if self.mesh is None:
            return SceneView._hide_elems(self, els)
        arr = els if hasattr(els, "Length") else System.Array[int](list(els))
        self._push_snapshot()
        n = self.mesh.HideElems(arr)
        if n == 0:
            self.hide_stack.pop()
            return
        self.n_hidden = self.mesh.HiddenElems
        self.owner.push_hide()
        self.mark_dirty()
        self.owner.hidden_changed("hide")

    def undo_hide(self):
        if self.mesh is None:
            return SceneView.undo_hide(self)
        if not self.hide_stack:
            return
        self.mesh.SetEHid(self.hide_stack.pop())
        self.n_hidden = self.mesh.HiddenElems
        self.mark_dirty()
        self.owner.hidden_changed("unhide")

    def unhide_all(self):
        if self.mesh is None:
            return SceneView.unhide_all(self)
        if self.hide_stack or self.n_hidden:
            self.mesh.UnhideAll()
            self.hide_stack = []
            self.n_hidden = 0
            self.owner.hides_cleared()
            self.mark_dirty()
            self.owner.hidden_changed("unhide")

    def hide_others_idx(self, kidx):
        """Verilen node indekslerine (secili) dokunmayan elemanlari gizler; gizlenen eleman sayisi."""
        self._push_snapshot()
        n = self.mesh.HideOthers(kidx)
        if n == 0:
            self.hide_stack.pop()
            return 0
        self.n_hidden = self.mesh.HiddenElems
        self.owner.push_hide()
        self.mark_dirty()
        self.owner.hidden_changed("hide")
        return n

    def hide_others(self, keep):
        if self.mesh is None:
            return SceneView.hide_others(self, keep)
        kidx = self.mesh.IndexMany(System.Array[int](list(keep)))
        return self.hide_others_idx(kidx)

    def visible_nodes(self):
        if self.mesh is None:
            return SceneView.visible_nodes(self)
        return list(self.mesh.VisibleNodeIds())

    def _visible_box(self):
        if self.mesh is None or self.n_hidden == 0:
            return SceneView._visible_box(self)
        b = self.mesh.VisibleBox()
        return self.model["box"] if b is None else tuple(b)

    # ----- cizim (arka plan, motorda) -----
    def _need_render(self):
        if self.mesh is None:
            return SceneView._need_render(self)
        if self.model is None or self.Width < 10 or self.Height < 10:
            return None
        same = self._img2 is not None and self._img2_sig == self._sig() and self._img2_cgen == self._content_gen
        seldirty = self._img2_selv != self._selsig() and (self._img2_big or self._bigsel2())
        if same and not seldirty and (self._img2_mode == "full" or self._moving):
            return None
        return "preview" if self._moving else "full"           # 'preview' = hareket sirasindaki kare (eski olsa da gosterilir)

    def _start_job(self, kind):
        if self.mesh is None:
            return SceneView._start_job(self, kind)
        job = {"mode": kind, "xf": self._xf(), "W": self.Width, "H": self.Height, "sig": self._sig(), "cgen": self._content_gen,
               "vgen": self._view_gen, "alive": True, "has_t": self.has_t, "tmap": self.tmap, "vmin": self.vmin, "vmax": self.vmax,
               "mesh_on": self.mesh_on, "nodes_on": self.nodes_on, "bigsel": self._bigsel2(), "showsel": self.listed_on,
               "selv": self._selsig(), "track": bool(self.minmax_on)}
        self._job = job
        th = threading.Thread(target=self._run_job, args=(job,))
        th.daemon = True
        th.start()

    def _run_job(self, job):
        if self.mesh is None:
            return SceneView._run_job(self, job)
        bmp = None
        try:
            m = self.mesh
            W, H = job["W"], job["H"]
            if self._px is None or len(self._px) != W * H:
                self._px = System.Array.CreateInstance(System.Int32, W * H)
            m00, m01, m10, m11, ox, oy = job["xf"]
            n = m.Render(self._px, W, H, m00, m01, m10, m11, ox, oy, self.bg.ToArgb(), self._palette(), bool(job["has_t"]), job["vmin"],
                         job["vmax"], bool(job["mesh_on"]), bool(job["nodes_on"]), bool(job["bigsel"]), bool(job["showsel"]))
            if self._content_gen != job["cgen"]:
                return
            bmp = self._copy_bitmap(self._px, W, H)
            mm = None
            if job["track"] and job["has_t"]:
                ei = m.MinMaxOnScreen(m00, m01, m10, m11, ox, oy, float(W), float(H))
                if ei[0] >= 0 and ei[1] >= 0:
                    a, b = m.Ids[ei[0]], m.Ids[ei[1]]
                    get = job["tmap"].get
                    mm = (a, get(a), b, get(b))
            self._result = (job, bmp, {"mm": mm, "n": n})
            bmp = None
        except Exception:
            pass
        finally:
            if bmp is not None:
                bmp.Dispose()
            job["alive"] = False

    def _on_poll(self, sender, e):
        if self.mesh is None:
            return SceneView._on_poll(self, sender, e)
        try:
            r = self._result
            if r is not None:
                self._result = None
                job, bmp, res = r
                if job["cgen"] != self._content_gen or self.model is None:
                    bmp.Dispose()
                else:
                    old = self._img2
                    self._img2, self._img2_sig, self._img2_mode, self._img2_cgen = bmp, job["sig"], job["mode"], job["cgen"]
                    self._img2_xf, self._img2_selv, self._img2_big = job["xf"], job["selv"], job["bigsel"]
                    if job["mode"] == "full":
                        self._mm = res["mm"] if (self.minmax_on and self.has_t) else None
                    if old is not None:
                        old.Dispose()
                    self.Invalidate()
            if self._job is not None and not self._job["alive"]:
                self._job = None
            if self._job is None and self.model is not None:
                k = self._need_render()
                if k is not None:
                    self._start_job(k)
        except Exception:
            pass

    def _draw_base(self, g):
        if self.mesh is None:
            return SceneView._draw_base(self, g)
        if self._img2 is None:
            return
        if self._img2_sig == self._sig() and self._img2_cgen == self._content_gen:
            g.DrawImage(self._img2, 0, 0)
        else:
            g.InterpolationMode = InterpolationMode.Bilinear
            self._blit(g, self._img2, self._img2.Width, self._img2.Height, self._img2_xf, False)

    def _paint(self, sender, e):
        SceneView._paint(self, sender, e)
        if self.mesh is not None and self.model is not None:
            e.Graphics.DrawString(u"[fast engine]", self.f_ui, self.b_dim, 8.0, 6.0)

    def _draw_selection(self, g):
        if self.mesh is None:
            return SceneView._draw_selection(self, g)
        show_sel = self.listed_on and len(self.sel) > 0
        if len(self.pend) == 0 and not show_sel:
            return
        if self._bigsel2() or self._img2_big:                  # cok sayida secili node: isaretler cizimin icine islenir
            return
        m00, m01, m10, m11, ox, oy = self._xf()
        W, H = float(self.Width), float(self.Height)
        dn = self.dn
        hid = self.hid_nodes
        g.SmoothingMode = SmoothingMode.AntiAlias
        for grp, brush, pen in ((self.pend, self.b_pend, self.p_pend), (self.sel if self.listed_on else (), self.b_green, self.p_glow)):
            for nid in grp:
                p = dn.get(nid)
                if p is None or nid in hid:
                    continue
                sx, sy = m00 * p[0] + m01 * p[1] + ox, m10 * p[0] + m11 * p[1] + oy
                if -10 <= sx <= W + 10 and -10 <= sy <= H + 10:
                    g.FillEllipse(brush, sx - 5.5, sy - 5.5, 11.0, 11.0)
                    g.DrawEllipse(pen, sx - 5.5, sy - 5.5, 11.0, 11.0)

    # ----- tiklama / kutu (motorda) -----
    def _find_node(self, mx, my, only=None):
        if self.mesh is None:
            return SceneView._find_node(self, mx, my, only)
        m00, m01, m10, m11, ox, oy = self._xf()
        i = self.mesh.PickNode(m00, m01, m10, m11, ox, oy, mx, my, 12.0, 0 if only is None else 3)       # only: secili node'lar arasinda
        return None if i < 0 else self.mesh.Ids[i]

    def _pick_elem(self, mx, my):
        if self.mesh is None:
            return SceneView._pick_elem(self, mx, my)
        m00, m01, m10, m11, ox, oy = self._xf()
        e = self.mesh.PickElem(m00, m01, m10, m11, ox, oy, mx, my)
        return None if e < 0 else e

    def _mup(self, sender, e):
        if self.mesh is None:
            return SceneView._mup(self, sender, e)
        if e.Button == MouseButtons.Middle:
            self._pan = None
            self._restore_cursor()
            return
        if self._down is None:
            return
        x0, y0 = self._down[0], self._down[1]
        m = self.mesh
        if self._drag and self._rb is not None:
            xa, xb = sorted((x0, e.X))
            ya, yb = sorted((y0, e.Y))
            m00, m01, m10, m11, ox, oy = self._xf()
            if self.hide_mode:
                self._hide_elems(m.ElemsInBox(m00, m01, m10, m11, ox, oy, xa, xb, ya, yb))
            else:
                idx = m.BoxNodes(m00, m01, m10, m11, ox, oy, xa, xb, ya, yb)
                snap = m.SelSnapshot()
                n = m.SelRemoveIdx(3, idx) if self.unsel_mode else m.SelAddIdx(SEL_PEND, idx, 3)
                if n:
                    self.owner.push_undo_data(snap)
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
                nid = self._find_node(x0, y0, True)
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
        cn.Click += lambda s_, e_: self.Close()          # modeless pencerede DialogResult kapatmaz: acikca Close
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
        self.chkUZ = self._mk_chk(go, u"UZ", 292, 56, 52, bool(owner.is3d))
        self.chkUZ.Visible = bool(owner.is3d)                                  # UZ sadece 3D modelde
        self.chkUS = self._mk_chk(go, u"USUM", 348, 56, 70, False)
        self.lblDU = self._mk_lbl(go, u"Disp unit:", 440, 60, 64)
        self.cbDU = self._mk_combo(go, 508, 56, 70, [u"mm", u"m"])
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
        bc.Click += lambda s_, e_: self.Close()
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
        for c in (self.lblDC, self.chkUX, self.chkUY, self.chkUZ, self.chkUS, self.lblDU, self.cbDU):
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
        if self.chkUZ.Checked and self.chkUZ.Visible:
            comps.append(u"UZ")
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
        self.f_gs = Font("Segoe UI Symbol", 11.5)
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
        small = H < 46
        lh = 14 if small else 17
        g.DrawString(self.glyph, self.f_gs if small else self.f_g, brush_of(fg), RectangleF(0.0, 1.0, float(W), float(H - lh)), self.sf)
        g.DrawString(self.text_, self.f_t, brush_of(tx), RectangleF(0.0, float(H - lh), float(W), 15.0), self.sf)


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
        self.is3d = False
        self.view2d = FastSceneView(self)
        self.view3d = FastCloudView(self)
        th_ = threading.Thread(target=native_cloud_type)            # C# motoru arka planda derlenir
        th_.daemon = True
        th_.start()
        self.view3d.Visible = False
        self.view = self.view2d
        self.top = Panel()
        self.rail = Panel()
        self.side = Panel()
        self.splitter = Panel()                      # sag paneli surukleyerek genisletip daraltmak icin
        self.splitter.BackColor = rgb(C_BORDER)
        self.splitter.Cursor = Cursors.VSplit
        self.splitter.MouseDown += lambda s, e: setattr(self, "_sdrag", True)
        self.splitter.MouseUp += lambda s, e: setattr(self, "_sdrag", False)
        self.splitter.MouseMove += self._split_move
        for c in (self.view2d, self.view3d, self.top, self.rail, self.side, self.splitter):
            self.Controls.Add(c)
        self.Resize += lambda s, e: self._layout()
        self._layout()
        self._build_top()
        self._build_rail()
        self._build_side()
        self._layout_rail()

    def _layout(self):
        W, H = self.ClientSize.Width, self.ClientSize.Height
        sw = max(320, min(self.side_w, W - RAILW - 320))
        self.side_w = sw
        self.top.SetBounds(0, 0, W, TOPH)
        self.rail.SetBounds(0, TOPH, RAILW, H - TOPH)
        self.side.SetBounds(W - sw, TOPH, sw, H - TOPH)
        self.splitter.SetBounds(W - sw - 6, TOPH, 6, H - TOPH)
        for v_ in (self.view2d, self.view3d):
            v_.SetBounds(RAILW, TOPH, max(100, W - RAILW - sw - 6), H - TOPH)
        self._fit_cols()
        if self.__dict__.get("_rail_groups"):
            self._layout_rail()                       # serit yuksekligi pencereye uyar

    def _split_move(self, sender, e):
        if self._sdrag:
            W = self.ClientSize.Width
            self.side_w = W - (self.splitter.Left + e.X)
            self._layout()

    def _fit_cols(self):
        """Tablonun son (deger) sutunu kalan genisligi doldurur."""
        if not self.__dict__.get("_cols_ready"):
            return
        w = self.lv.Width - (70 + 110 + 110 + self.lv.Columns[3].Width) - 26
        if w >= 70:
            self.lv.Columns[4].Width = w

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
        # anahtar: (simge, etiket, islev, ipucu, 2D'de gorunur, 3D'de gorunur)
        B = {
            "fit": (u"\u21f1", u"Fit", self.fit_view, u"Fit model to window", 1, 1),
            "mesh": (u"\u25a6", u"Mesh", self.toggle_mesh, u"Show / hide mesh lines", 1, 0),
            "nodes": (u"\u2237", u"Nodes", self.toggle_nodes, u"Show / hide node points", 1, 0),
            "iso": (u"\u25c8", u"Iso", lambda: self._preset("iso"), u"Isometric view", 0, 1),
            "px": (u"+X", u"", lambda: self._preset("+x"), u"View from +X (looking toward -X)", 0, 1),
            "mx": (u"\u2212X", u"", lambda: self._preset("-x"), u"View from -X (looking toward +X)", 0, 1),
            "py": (u"+Y", u"", lambda: self._preset("+y"), u"View from +Y (looking toward -Y)", 0, 1),
            "my": (u"\u2212Y", u"", lambda: self._preset("-y"), u"View from -Y (looking toward +Y)", 0, 1),
            "pz": (u"+Z", u"", lambda: self._preset("+z"), u"View from +Z (looking down, toward -Z)", 0, 1),
            "mz": (u"\u2212Z", u"", lambda: self._preset("-z"), u"View from -Z (looking up, toward +Z)", 0, 1),
            "rotl": (u"\u27f2", u"Rot L", lambda: self.rotate_clicked(-1), u"Rotate by the angle above about the chosen axis (right-hand rule, counter-clockwise)", 1, 1),
            "rotr": (u"\u27f3", u"Rot R", lambda: self.rotate_clicked(1), u"Rotate by the angle above about the chosen axis, the other way (clockwise)", 1, 1),
            "selall": (u"\u2611", u"Select all", self.select_all, u"Pick all (not hidden) nodes (yellow); press Add to list to keep them", 1, 1),
            "clear": (u"\u2715", u"Clear", self.clear_sel, u"Clear the picked and listed nodes", 1, 1),
            "unsel": (u"\u2296", u"Unselect", self.toggle_unsel, u"Click a node or right-drag a box to remove nodes from the picks / list", 1, 0),
            "undo": (u"\u21b6", u"Undo", self.undo, u"Undo the last selection / hide step (Ctrl+Z)", 1, 1),
            "hide": (u"\u25a8", u"Hide el.", self.toggle_hide, u"Click an element (or right-drag a box) to hide it", 1, 0),
            "hidesel": (u"\u25a8", u"Hide", self.hide_selected_clicked, u"Hide the selected (picked / listed) nodes", 0, 1),
            "hideo": (u"\u25e7", u"Hide others", self.hide_others_clicked, u"Hide what does not touch the selected nodes", 1, 1),
            "showall": (u"\u25eb", u"Show all", self.show_all, u"Show all hidden parts again", 1, 1),
            "probe": (u"\u2295", u"Probe", self.toggle_probe, u"Click a node to show its values on the screen", 1, 1),
            "minmax": (u"\u2195", u"Max/Min", self.toggle_minmax, u"Show / hide the MAX and MIN labels of the nodes on the screen", 1, 1)}
        # grup: satirlar (her satir 1 veya 2 dugme; "ROT" = aci / eksen kutulari)
        self._rail_groups = [
            ("view", [("fit",), ("mesh",), ("nodes",)]),
            ("view", [("iso",), ("px", "mx"), ("py", "my"), ("pz", "mz")]),
            ("view", [("ROT",), ("rotl",), ("rotr",)]),
            ("sel", [("selall",), ("clear",), ("unsel",), ("undo",)]),
            ("hide", [("hide",), ("hidesel",), ("hideo",), ("showall",)]),
            ("tool", [("probe",), ("minmax",)])]
        self._rail_vis = {}
        for key, (glyph, text, handler, tip, v2, v3) in B.items():
            grp = [g_ for g_, rows in self._rail_groups if any(key in row for row in rows)][0]
            b = RailButton(glyph, text, grp, handler)
            r.Controls.Add(b)
            self._tt.SetToolTip(b, tip)
            self.rb[key] = b
            self._rail_vis[key] = (v2, v3)
        self._rail_strips = []
        for grp, rows in self._rail_groups:
            st = Panel()
            st.BackColor = rgb(GRP[grp])
            r.Controls.Add(st)
            self._rail_strips.append(st)
        self.tbAng = TextBox()                                    # donme acisi (derece)
        self.tbAng.Text = u"5"
        self.tbAng.BackColor = rgb(C_COMBO)
        self.tbAng.ForeColor = rgb(C_TEXT)
        self.tbAng.BorderStyle = BorderStyle.FixedSingle
        self.tbAng.Font = Font("Segoe UI", 11.0)
        r.Controls.Add(self.tbAng)
        self._tt.SetToolTip(self.tbAng, u"Rotation angle in degrees (default 5). Type any value.")
        self.lblDeg = self._dlabel(r, u"deg", 0, 0, 36, 20, muted=True, size=9.0)
        self.cbAxis = self._dcombo(r, 0, 0, RAILW - 20)           # 3D: donme ekseni
        for t_ in (u"Axis X", u"Axis Y"):
            self.cbAxis.Items.Add(t_)
        self.cbAxis.SelectedIndex = 0
        self._axis_mode = False
        self._tt.SetToolTip(self.cbAxis, u"The model rotates about this (global) axis")
        self.rb["minmax"].set_on(True)
        self._layout_rail()

    def _layout_rail(self):
        """Gorunur dugmeleri (2D / 3D'ye gore) alt alta dizer; yukseklik pencereye gore ayarlanir, serit kaydirma cubugu cikarmaz."""
        is3 = self.is3d
        for k, b in self.rb.items():
            b.Visible = False
        vis_rows = []
        for gi, (grp, rows) in enumerate(self._rail_groups):
            vr = []
            for row in rows:
                if row[0] == "ROT":
                    vr.append(row)
                else:
                    keys = [k for k in row if self._rail_vis[k][1 if is3 else 0]]
                    if keys:
                        vr.append(tuple(keys))
            vis_rows.append(vr)
        nbtn = sum(1 for vr in vis_rows for row in vr if row[0] != "ROT")
        nrot = sum(1 for vr in vis_rows for row in vr if row[0] == "ROT")
        ngrp = sum(1 for vr in vis_rows if vr)
        avail = max(self.rail.Height, 200)
        fixed = 10 + nrot * 62 + ngrp * 8 + 6
        bh = int((avail - fixed) / float(max(nbtn, 1))) - 4
        bh = max(30, min(52, bh))
        need = fixed + nbtn * (bh + 4) > avail
        self.rail.AutoScroll = need                       # sigiyorsa kaydirma cubugu hic cikmaz
        W1, W2 = RAILW - 20, (RAILW - 20 - 4) // 2
        y = 10
        for (grp, rows), vr, st in zip(self._rail_groups, vis_rows, self._rail_strips):
            y0 = y
            for row in vr:
                if row[0] == "ROT":
                    self.tbAng.SetBounds(12, y, 56, 26)
                    self.lblDeg.SetBounds(74, y + 3, 36, 20)
                    self.cbAxis.SetBounds(12, y + 30, W1, 30)
                    self.tbAng.Visible = self.lblDeg.Visible = self.cbAxis.Visible = True
                    y += 62
                    continue
                for i, k in enumerate(row):
                    b = self.rb[k]
                    if len(row) == 1:
                        b.SetBounds(12, y, W1, bh)
                    else:
                        b.SetBounds(12 + i * (W2 + 4), y, W2, bh)
                    b.Visible = True
                y += bh + 4
            st.Visible = y > y0
            if y > y0:
                st.SetBounds(3, y0, 4, y - y0 - 4)
                y += 8

    def rotate_clicked(self, sign):
        """Rot L (sign -1) / Rot R (+1): acisi kutudan (varsayilan 5), 3D'de secili eksen etrafinda doner."""
        try:
            a = float(self.tbAng.Text.strip().replace(u",", u"."))
        except ValueError:
            MessageBox.Show(u"Rotate angle must be a number (degrees).", TITLE, MessageBoxButtons.OK, MessageBoxIcon.Information)
            return
        self.view.rotate_axis(self.cbAxis.SelectedIndex, -sign * a)          # Rot L = sag el kurali (+), Rot R = (-)

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
        self.lv.Columns.Add(u"Z", 0)
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
            if self.is3d:
                lines.append(u"UX %.5g   UY %.5g   UZ %.5g   USUM %.5g  [%s]" % (ux, uy, uz, math.sqrt(ux * ux + uy * uy + uz * uz), u"m" if f_ != 1.0 else u"mm"))
            else:
                lines.append(u"UX %.5g     UY %.5g     USUM %.5g   [%s]" % (ux, uy, math.sqrt(ux * ux + uy * uy + uz * uz), u"m" if f_ != 1.0 else u"mm"))
        return lines

    def refresh_probes(self):
        v = self.view
        v.probe_text = dict((nid, self._probe_lines(nid)) for nid in v.probes)
        v.Invalidate()

    # ----- dosyalar -----
    def _hd0(self):
        """Export / tablo koordinat basligi: 2D -> X_Loc Y_Loc, 3D -> X_Loc Y_Loc Z_Loc."""
        return [u"NodeID", u"X_Loc", u"Y_Loc"] + ([u"Z_Loc"] if self.is3d else [])

    def _preset(self, name):
        if self.is3d:
            self.view3d.set_preset(name)

    def _set_dim_mode(self, is3):
        """2D / 3D gorunum degisimi: aktif gorunum, 3D on ayar dugmeleri, tablodaki Z sutunu."""
        self.is3d = is3
        self.view = self.view3d if is3 else self.view2d
        self.view2d.Visible = not is3
        self.view3d.Visible = is3
        for v_ in (self.view2d, self.view3d):
            try:
                if v_ is self.view:
                    v_._poll.Start()
                else:
                    v_._poll.Stop()
            except Exception:
                pass
        if is3 != self._axis_mode:                       # eksen listesi: 2D = X Y, 3D = X Y Z
            self._axis_mode = is3
            self.cbAxis.Items.Clear()
            for t_ in ((u"Axis X", u"Axis Y", u"Axis Z") if is3 else (u"Axis X", u"Axis Y")):
                self.cbAxis.Items.Add(t_)
            self.cbAxis.SelectedIndex = 2 if is3 else 0
        self._set_hide(False)
        if is3:
            self._set_unsel(False)
        self._layout_rail()
        self.lv.Columns[3].Width = 110 if is3 else 0
        self._layout()
        self._fit_cols()

    def ask_files(self):
        if self._start is not None:
            self._start.Activate()
            return
        sf = StartForm(self.mesh_path, self.temp_path, self.disp_path, lambda m, t, d: self.load_files(m, t, d))
        sf.FormClosed += lambda s_, e_: setattr(self, "_start", None)
        self._start = sf
        sf.Show(self)

    def _on_closed(self, sender, e):
        """Pencere kapaninca buyuk veriler ve cizim kaynaklari birakilir; acik yardimci pencereler de kapatilir."""
        for fm in (self._start, self._export):
            try:
                if fm is not None:
                    fm.Close()
            except Exception:
                pass
        for v in (self.view2d, self.view3d):
            try:
                v._content_gen += 1
                v._view_gen += 1
                for t in (v._timer, v._poll, v._anim_timer):
                    t.Stop()
                for nm in ("_cache", "_mip", "_leg_bmp", "_img3", "_img2"):
                    b = getattr(v, nm, None)
                    if b is not None:
                        b.Dispose()
                    setattr(v, nm, None)
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
        info3 = None
        nbad, bad_sample = 0, u""
        try:
            pf.update(0.0, u"Reading mesh file ...")
            t_nodes = time.time()
            Cl = native_cloud_type()
            cloud = None
            if Cl is not None:
                try:
                    pf.update(0.1, u"Reading nodes ...")
                    cloud = Cl.Load(mesh_path)
                except Exception:
                    cloud = None
            if cloud is not None:
                is3, nbad, bad_sample = bool(cloud.Is3D), cloud.BadCount, cloud.BadSample
                if is3:
                    model, mv, info3 = cloud_model_native(cloud), ArrayMap(_arr("i"), _arr("d")), {"t": time.time() - t_nodes}
                    nodes = model["nodes"]
            else:
                nt3, is3, nbad, bad_sample = read_nodes3d(mesh_path, pf.update)
                if is3:
                    model, mv, info3 = cloud_model(nt3), ArrayMap(_arr("i"), _arr("d")), {"t": time.time() - t_nodes}
                    nodes = nt3
            if not is3:
                m2 = None
                M2 = _NATIVE.get("mesh2d")
                if M2 is not None:
                    try:
                        pf.update(0.1, u"Reading mesh (fast engine) ...")
                        m2 = M2.Load(mesh_path)
                        if m2.NE == 0:
                            m2 = None
                    except Exception:
                        m2 = None
                if m2 is not None:
                    model, mv = mesh_model_native(m2), ArrayMap(m2.MvIds, m2.MvVals)
                    nodes = model["nodes"]
                else:
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
        self._set_dim_mode(is3)
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
            if is3:
                self.fields.append({"name": u"Disp UZ", "src": "disp", "comp": 2})
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
        self.lv.Columns[4].Text = u"Value"
        self.selection_changed()
        if is3 and nbad:
            MessageBox.Show(u"%d node kaydi okunamadi (modelde eksik node olabilir). Ornek kayit:\n\n%s" % (nbad, bad_sample), TITLE,
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
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
        self.lv.Columns[4].Text = u"Value"
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
        if getattr(v, "engine", None) is not None:                    # C# motoru: gizlenmemis node'lar arasinda min / max
            v.ensure_values(tm)
            ei = v.engine.Extremes()
            if ei[0] >= 0 and ei[1] >= 0:
                nt = v.model["nodes"]
                nlo, nhi = nt.id_of(ei[0]), nt.id_of(ei[1])
                lo, hi = tm.get(nlo), tm.get(nhi)
        elif full:
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
            for n_ in ([] if getattr(v, "engine", None) is not None else v.visible_nodes()):
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
        self.lv.Columns[4].Text = name
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

    def hide_selected_clicked(self):
        """3D: secili (sari + listedeki) node'lari gizler."""
        v = self.view
        if getattr(v, "nat_sel", False):
            idx = v.sel_indices()
            if len(idx) == 0:
                MessageBox.Show(u"Once ekranda en az bir node sec.", TITLE, MessageBoxButtons.OK, MessageBoxIcon.Information)
                return
            if len(v.pend):
                self.push_undo()
                v.pend = set()
                self.selection_changed()
            v._hide_elems(idx)
            return
        keep = set(v.sel) | set(v.pend)
        if not keep:
            MessageBox.Show(u"Once ekranda en az bir node sec.", TITLE, MessageBoxButtons.OK, MessageBoxIcon.Information)
            return
        nt = v.model["nodes"]
        idx = [i for i in (nt.index(n) for n in keep) if i >= 0]
        if v.pend:
            self.push_undo()
            v.pend = set()
            self.selection_changed()
        v._hide_elems(idx)

    def hide_others_clicked(self):
        v = self.view
        if getattr(v, "nat_sel", False):
            idx = v.sel_indices()
            if len(idx) == 0:
                MessageBox.Show(u"Once ekranda en az bir node sec.", TITLE, MessageBoxButtons.OK, MessageBoxIcon.Information)
                return
            if len(v.pend):
                self.push_undo()
                v.pend = set()                       # secim kullanildi: ekrandaki sari isaretler kalkar
                self.selection_changed()
            n = v.hide_others_idx(idx)
            if n == 0:
                MessageBox.Show(u"Gizlenecek baska eleman yok.", TITLE, MessageBoxButtons.OK, MessageBoxIcon.Information)
            return
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
        self.view.sel_version = getattr(self.view, "sel_version", 0) + 1       # 3D: cok sayida secim cizime islenir
        MAXLIST = 2000
        v = self.view
        fld = self.cur_fld if self._shown else None
        off = self._off() if (fld is not None and fld["src"] == "temp") else 0.0
        fac = self._dfac() if (fld is not None and fld["src"] == "disp") else 1.0
        tm = self.cur_map
        nsel = len(v.sel)
        sel = v.sel.smallest(MAXLIST) if hasattr(v.sel, "smallest") else heapq.nsmallest(MAXLIST, v.sel)         # buyuk kumeyi tamamen siralamaz
        items = []
        for nid in sel:
            p3 = self.nodes[nid]
            X, Y = p3[0], p3[1]
            if fld is None:
                vt = u"-"
            else:
                val = tm.get(nid)
                vt = u"NaN" if val is None else (u"%.3f" % (val + off) if fld["src"] == "temp" else u"%.5g" % (val * fac))
            items.append(ListViewItem(System.Array[str]([u"%d" % nid, u"%.5f" % X, u"%.5f" % Y, (u"%.5f" % p3[2]) if self.is3d else u"", vt])))
        if nsel > MAXLIST:
            items.append(ListViewItem(System.Array[str]([u"...", u"+%d more" % (nsel - MAXLIST), u"", u"", u""])))
        self.lv.BeginUpdate()
        try:
            self.lv.Items.Clear()
            self.lv.Items.AddRange(System.Array[ListViewItem](items))
        finally:
            self.lv.EndUpdate()
        self.lblSel.Text = u"Selected nodes (%d)" % nsel
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
        if getattr(self.view, "nat_sel", False):                   # hizli gorunum: secim bayraklarinin kopyasi (milisaniye)
            self._undo.append(("selN", self.view.sel_snapshot()))
        else:
            self._undo.append(("sel", (set(self.view.sel), set(self.view.pend))))
        if len(self._undo) > 100:
            del self._undo[0]

    def push_undo_data(self, snap):
        """Hizli gorunum: degisiklikten ONCE alinmis secim kopyasini geri alma yigitina koyar."""
        self._undo.append(("selN", snap))
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
            if kind == "selN":
                if getattr(self.view, "nat_sel", False):
                    self.view.sel_restore(data)
                    self.selection_changed()
                    return
                continue
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
        if getattr(v, "nat_sel", False):               # hizli gorunum: motorda (milyonlarca node'da anlik)
            v.sel_select_all()
        elif self.is3d:                                # 3D: gizlenmemis TUM node'lar (aralik / kume islemleriyle hizli)
            nt = v.model["nodes"]
            allids = set(range(nt.first, nt.first + len(nt))) if nt.contig else set(nt.ids)
            if v.n_hidden:
                if getattr(v, "cloud", None) is not None:
                    allids -= set(v.cloud.HiddenIds())
                else:
                    nh = v.nhid
                    allids -= set(nt.id_of(i) for i in range(len(nt)) if nh[i])
            v.pend = allids - v.sel
        else:
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
        head = self._hd0() + [u"T_%.6g" % t for t in times]
        if o["mv"]:
            head.append(u"MeshVal")
        out = [sep.join(head)]
        for j, nid in enumerate(sel):
            p3 = self.nodes[nid]
            X, Y = p3[0], p3[1]
            row = [u"%d" % nid, u"%.10g" % X, u"%.10g" % Y] + ([u"%.10g" % p3[2]] if self.is3d else [])
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
        need = [0, 1, 2] if u"USUM" in comps else sorted(set(c for c, nm in ((0, u"UX"), (1, u"UY"), (2, u"UZ")) if nm in comps))
        data, err = self._read_times(self.df, times, sel, need, u"Disp")
        if err is not None:
            if err:
                MessageBox.Show(err, TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error)
            return None
        f = 0.001 if o["disp_unit"] == u"m" else 1.0            # dosya mm kabul edilir
        sep = o["sep"]
        names = {u"UX": u"DX", u"UY": u"DY", u"UZ": u"DZ", u"USUM": u"DSUM"}
        head = self._hd0()
        for t in times:
            if len(comps) == 1:
                head.append(u"D_%.6g" % t)
            else:
                head += [u"%s_%.6g" % (names[c], t) for c in comps]
        if o["mv"]:
            head.append(u"MeshVal")
        out = [sep.join(head)]
        for j, nid in enumerate(sel):
            p3 = self.nodes[nid]
            X, Y = p3[0], p3[1]
            row = [u"%d" % nid, u"%.10g" % X, u"%.10g" % Y] + ([u"%.10g" % p3[2]] if self.is3d else [])
            for d in data:
                for c in comps:
                    if c == u"UX":
                        v = d[need.index(0)][j]
                    elif c == u"UY":
                        v = d[need.index(1)][j]
                    elif c == u"UZ":
                        v = d[need.index(2)][j]
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
