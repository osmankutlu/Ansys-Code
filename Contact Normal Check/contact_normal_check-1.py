# -*- coding: utf-8 -*-
# =====================================================================
#  Contact Normal Viewer / Flip   (Ansys Mechanical 2D, IronPython 2.7)
#
#  Sayfa 1 : kontak listesi (agac)
#  Sayfa 2 : sol  -> secili kontaklar (check = flip, Side = contact/target/both)
#            sag  -> Model view  (mesh'ten cizilen, tamamen interaktif; oklar =
#                    eleman dugum sirasindan hesaplanan normal)
#                    Mechanical view (Mechanical'in kendi Element Normals goruntusu)
#  Goruntu: -Z'den bakis, X yukari, Y saga
#  Mouse : tekerlek = rotate | Ctrl+tekerlek = pan (dikey) | Ctrl+Shift+tekerlek = pan (yatay)
#          Alt+tekerlek = zoom | sol surukle = pan | tikla = eleman sec/birak
#          Shift+surukle = kutu ile eleman sec
# =====================================================================
import clr
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
import System
import re

# ==== CORE START ====
import math


def compute_segments(elems, xy, edge_nodes, side):
    """elems: [(eid, corner_ids, all_ids)]; xy: {nid: (x, y)}; returns list of segment dicts.
    Normal = right-hand side of the edge direction in element node order = (ty, -tx)."""
    segs = []
    for eid, corner, alln in elems:
        n = len(corner)
        if n < 2:
            continue
        for k in range(n):
            a, b = corner[k], corner[(k + 1) % n]
            if a == b or a not in edge_nodes or b not in edge_nodes:
                continue
            pa, pb = xy.get(a), xy.get(b)
            if pa is None or pb is None:
                continue
            tx, ty = pb[0] - pa[0], pb[1] - pa[1]
            L = math.hypot(tx, ty)
            if L == 0:
                continue
            mid = alln[n + k] if len(alln) == 2 * n else None
            segs.append({"side": side, "eid": eid, "a": a, "b": b, "m": mid,
                         "ax": pa[0], "ay": pa[1], "bx": pb[0], "by": pb[1],
                         "mx": (pa[0] + pb[0]) / 2.0, "my": (pa[1] + pb[1]) / 2.0,
                         "nx": ty / L, "ny": -tx / L, "len": L})
    return segs


def build_flip_whole(name, side):
    L = ["! ContactNormalCheck - normal flip (ESURF,,REVERSE)",
         "! Contact: %s" % name,
         "! FLIP whole side=%s" % side]
    if side in ("both", "contact"):
        L += ["esel,s,type,,cid", "esurf,,reverse"]
    if side in ("both", "target"):
        L += ["esel,s,type,,tid", "esurf,,reverse"]
    L.append("allsel,all")
    return "\n".join(L) + "\n"


def build_flip_elems(name, segs):
    """Flip only the given segments (selected by their nodes)."""
    L = ["! ContactNormalCheck - normal flip of selected elements (ESURF,,REVERSE)",
         "! Contact: %s" % name,
         "! If the mesh changes, re-run the tool and update this command."]
    for kind, typ in (("contact", "cid"), ("target", "tid")):
        group = [s for s in segs if s["side"] == kind]
        if not group:
            continue
        L.append("! FLIP elems side=%s n=%d" % (kind, len(group)))
        L.append("esel,none")
        for s in group:
            L.append("nsel,s,node,,%d" % s["a"])
            L.append("nsel,a,node,,%d" % s["b"])
            if s["m"]:
                L.append("nsel,a,node,,%d" % s["m"])
            L.append("esln,a,1")
        L.append("esel,r,type,,%s" % typ)
        L.append("esurf,,reverse")
        L.append("allsel,all")
    return "\n".join(L) + "\n"
# ==== CORE END ====

from System.Windows.Forms import (Form, Button, Label, MessageBox, MessageBoxButtons,
                                  MessageBoxIcon, FormStartPosition, FormBorderStyle,
                                  DialogResult, GroupBox, FlatStyle, TextBox, TreeView,
                                  TreeNode, Application, ComboBox, ComboBoxStyle,
                                  ListView, ListViewItem, View, ScrollBars, Panel,
                                  Cursor, Cursors, Control, Keys, MouseButtons,
                                  CheckedListBox, CheckState)
from System.Drawing import (Size, Point, Font, FontStyle, Color, Bitmap, PointF, Pen,
                            SolidBrush, Rectangle, RectangleF)
from System.Drawing.Drawing2D import Matrix, SmoothingMode, DashStyle

SNIPPET_NAME = u"ContactNormalFlip"
IMG_W, IMG_H = 2000, 1250
PALETTE = [(150, 165, 182), (155, 180, 148), (238, 228, 214), (214, 214, 230),
           (232, 200, 200), (200, 225, 225)]


# ---------------------------------------------------------------
# Contact helpers
# ---------------------------------------------------------------
def obj_id(c):
    try:
        return int(c.ObjectId)
    except Exception:
        return None


def label_of(c):
    oid = obj_id(c)
    return c.Name + (u"  [ID %d]" % oid if oid is not None else u"")


def ukey_of(c, idx):
    oid = obj_id(c)
    return oid if oid is not None else 1000000 + idx


def build_contact_tree(connections):
    tree = []
    for conn_group in connections.Children:
        group_contacts = []
        for contact in conn_group.Children:
            if "ContactRegion" in contact.GetType().ToString():
                group_contacts.append(contact)
        if group_contacts:
            tree.append((conn_group.Name, group_contacts))
    return tree


def find_flip_snippet(contact):
    for child in contact.Children:
        try:
            if "CommandSnippet" in child.GetType().ToString() and child.Name == SNIPPET_NAME:
                return child
        except Exception:
            pass
    return None


# ---------------------------------------------------------------
# Model scene (from mesh data)
# ---------------------------------------------------------------
def load_scene(log):
    """Reads bodies / 2D elements / node coordinates from the Mechanical mesh."""
    mesh = ExtAPI.DataModel.MeshDataByName("Global")
    bodies = []
    try:
        for asm in ExtAPI.DataModel.GeoData.Assemblies:
            for part in asm.Parts:
                for body in part.Bodies:
                    eids = []
                    for face in body.Faces:
                        eids.extend(list(mesh.MeshRegionById(face.Id).ElementIds))
                    if eids:
                        bodies.append((u"%s\\%s [%d]" % (part.Name, body.Name, body.Id), eids))
    except Exception as ex:
        log(u"body/face lookup failed (%s) - drawing all elements as one body" % str(ex))
        bodies = []
    if not bodies:
        bodies = [(u"Model", [e.Id for e in mesh.Elements])]
    xy = {}
    scene = {"bodies": [], "xy": xy, "mesh": mesh}
    for bi, (name, eids) in enumerate(bodies):
        elems = []
        for eid in eids:
            el = mesh.ElementById(eid)
            alln = list(el.NodeIds)
            try:
                corner = list(el.CornerNodeIds)
            except Exception:
                corner = alln
            for nid in alln:
                if nid not in xy:
                    nd = mesh.NodeById(nid)
                    xy[nid] = (nd.X, nd.Y)
            poly = System.Array[PointF]([PointF(xy[n][1], -xy[n][0]) for n in corner if n in xy])
            elems.append((eid, corner, alln, poly))
        scene["bodies"].append({"name": name, "elems": elems, "visible": True,
                                "color": Color.FromArgb(*PALETTE[bi % len(PALETTE)])})
    xs = [p[0] for p in xy.values()]
    ys = [p[1] for p in xy.values()]
    scene["box"] = (min(ys), max(ys), -max(xs), -min(xs))      # draw coords (X=y, Y=-x)
    return scene


def contact_segments(scene, c, log):
    mesh = scene["mesh"]
    out = []
    for side, loc in (("contact", c.SourceLocation), ("target", c.TargetLocation)):
        nodes = set()
        for gid in list(loc.Ids):
            try:
                nodes.update(list(mesh.MeshRegionById(gid).NodeIds))
            except Exception as ex:
                log(u"%s: edge %s mesh lookup failed (%s)" % (label_of(c), gid, str(ex)))
        for b in scene["bodies"]:
            out.extend(compute_segments([(e[0], e[1], e[2]) for e in b["elems"]],
                                        scene["xy"], nodes, side))
    return out


# ---------------------------------------------------------------
# Mechanical graphics capture (reference view)
# ---------------------------------------------------------------
def normal_props(c):
    names = [n for n in dir(c) if "normal" in n.lower() and not n.startswith("_")]
    pri = [n for n in names if "element" in n.lower()]
    rest = [n for n in names if n not in pri and
            not any(k in n.lower() for k in ("stiff", "updat", "factor", "toler"))]
    return pri + rest


def set_normals(c, want, log):
    for n in normal_props(c):
        try:
            cur = getattr(c, n)
        except Exception:
            continue
        before = str(cur)
        ok = False
        try:
            if isinstance(cur, bool):
                setattr(c, n, want)
                ok = True
            else:
                T = cur.GetType()
                if T.IsEnum:
                    cands = (u"Yes", u"On", u"Show", u"True") if want else (u"No", u"Off", u"Hide", u"False")
                    for cand in cands:
                        try:
                            setattr(c, n, System.Enum.Parse(T, cand))
                            ok = True
                            break
                        except Exception:
                            pass
        except Exception:
            pass
        if ok:
            try:
                after = str(getattr(c, n))
            except Exception:
                after = u"?"
            if after != before:
                return n, cur
    log(u"could not switch Element Normals on/off automatically")
    return None, None


def restore_normals(c, name, old):
    if name is not None:
        try:
            setattr(c, name, old)
        except Exception:
            pass


def _val(q):
    try:
        return float(q.Value)
    except Exception:
        return float(q)


def _like(old, newval):
    try:
        return type(old)(newval, old.Unit)
    except Exception:
        return newval


_MODEL_BOX = []


def zoom_camera_to(segs, scene, log):
    import math as _m
    cam = ExtAPI.Graphics.Camera
    cam.SetFit()
    xs = [s["ax"] for s in segs] + [s["bx"] for s in segs]
    ys = [s["ay"] for s in segs] + [s["by"] for s in segs]
    cb = (min(xs), max(xs), min(ys), max(ys))
    allx = [p[0] for p in scene["xy"].values()]
    ally = [p[1] for p in scene["xy"].values()]
    mb = (min(allx), max(allx), min(ally), max(ally))
    H0 = cam.SceneHeight
    F0 = cam.FocalPoint
    h0 = _val(H0)
    mw, mh = mb[1] - mb[0], mb[3] - mb[2]
    cw, ch = cb[1] - cb[0], cb[3] - cb[2]
    ext = max(mh, mw / 1.6)
    best, u = None, 1.0
    for cand in (1.0, 1e-3, 1e3, 1e-6, 1e6, 0.0254, 39.3701):
        r = abs(_m.log(h0 / (cand * ext)))
        if best is None or r < best:
            best, u = r, cand
    tgt = max(ch * 1.8, cw * 1.8 / 1.6, ext * 0.02)
    new_h = h0 * tgt / ext
    fx, fy, fz = _val(F0.X), _val(F0.Y), _val(F0.Z)
    cam.FocalPoint = type(F0)(fx + u * ((cb[0] + cb[1]) / 2.0 - (mb[0] + mb[1]) / 2.0),
                              fy + u * ((cb[2] + cb[3]) / 2.0 - (mb[2] + mb[3]) / 2.0), fz)
    cam.SceneHeight = _like(H0, new_h)


def redraw():
    Application.DoEvents()
    for nm in ("Refresh", "Redraw", "Update"):
        try:
            getattr(ExtAPI.Graphics, nm)()
            break
        except Exception:
            pass
    System.Threading.Thread.Sleep(400)
    Application.DoEvents()


def set_mesh_display(on):
    try:
        ExtAPI.Graphics.ViewOptions.ShowMesh = on
        return True
    except Exception:
        return False


def export_image(path):
    gr = ExtAPI.Graphics
    try:
        st = Ansys.Mechanical.Graphics.GraphicsImageExportSettings()
        st.Resolution = GraphicsResolutionType.EnhancedResolution
        st.Background = GraphicsBackgroundType.White
        st.Width = IMG_W
        st.Height = IMG_H
        st.CurrentGraphicsDisplay = False
        gr.ExportImage(path, GraphicsImageExportFormat.PNG, st)
    except Exception:
        gr.ExportImage(path, GraphicsImageExportFormat.PNG)


def load_bitmap(path):
    ms = System.IO.MemoryStream(System.IO.File.ReadAllBytes(path))
    b = Bitmap(ms)
    img = Bitmap(b)
    b.Dispose()
    ms.Dispose()
    return img


def capture_contact(cd, idx, outdir, log, rescale=False):
    """Mechanical screenshot with Element Normals on (view from -Z = 'Back', image later rotated 90 CW)."""
    c = cd["c"]
    L = lambda m: log(u"%s: %s" % (label_of(c), m))
    name, old = None, None
    try:
        Tree.Activate([c])
        try:
            ExtAPI.Graphics.Camera.SetSpecificViewOrientation(ViewOrientationType.Back)
        except Exception:
            pass
        done = False
        if cd["segs"]:
            try:
                zoom_camera_to(cd["segs"], cd["scene"], L)
                done = True
            except Exception as ex:
                L(u"camera zoom failed (%s) - full fit" % str(ex))
        if not done:
            try:
                ExtAPI.Graphics.Camera.SetFit()
            except Exception:
                pass
        # normals are switched on AFTER the camera is set so the arrows are scaled to this view
        if rescale:
            for n in dir(c):
                if "rescale" in n.lower():
                    try:
                        getattr(c, n)()
                        break
                    except Exception:
                        pass
            n0, o0 = set_normals(c, False, L)
            redraw()
        name, old = set_normals(c, True, L)
        redraw()
        path = System.IO.Path.Combine(outdir, u"contact_%d.png" % idx)
        if System.IO.File.Exists(path):
            System.IO.File.Delete(path)
        export_image(path)
        return load_bitmap(path)
    except Exception as ex:
        L(u"capture failed: %s" % str(ex))
        return None
    finally:
        restore_normals(c, name, old)


# ---------------------------------------------------------------
# Interactive view
# ---------------------------------------------------------------
def _dist_pt_seg(px, py, ax, ay, bx, by):
    dx, dy = bx - ax, by - ay
    d2 = dx * dx + dy * dy
    t = 0.0 if d2 == 0 else max(0.0, min(1.0, ((px - ax) * dx + (py - ay) * dy) / d2))
    cx, cy = ax + t * dx, ay + t * dy
    return math.hypot(px - cx, py - cy)


class SceneView(Panel):
    def __init__(self, owner):
        Panel.__init__(self)
        self.owner = owner
        self.DoubleBuffered = True
        self.BackColor = Color.White
        self.mode = u"model"
        self.scene = None
        self.segs = []
        self.sel = set()
        self.image = None
        self.mesh_on = False
        self.s = 1.0
        self.theta = 0.0
        self.px = 0.0
        self.py = 0.0
        self.cx = 0.0
        self.cy = 0.0
        self.arrow = 1.0
        self._down = None
        self._drag = False
        self._rb = None
        self.Paint += self._paint
        self.MouseDown += self._mdown
        self.MouseMove += self._mmove
        self.MouseUp += self._mup
        self.MouseWheel += self._wheel
        self.MouseEnter += lambda s, e: self.Focus()
        self.Resize += lambda s, e: self.Invalidate()

    # ----- transforms -----
    def _matrix(self):
        m = Matrix()
        m.Translate(self.Width / 2.0 + self.px, self.Height / 2.0 + self.py)
        m.Rotate(self.theta)
        m.Scale(self.s, self.s)
        m.Translate(-self.cx, -self.cy)
        return m

    def _to_screen(self, pts):
        arr = System.Array[PointF]([PointF(x, y) for (x, y) in pts])
        self._matrix().TransformPoints(arr)
        return [(p.X, p.Y) for p in arr]

    # ----- fit / zoom -----
    def _fit_box(self, box):
        x0, x1, y0, y1 = box
        bw, bh = x1 - x0, y1 - y0
        big = max(bw, bh, 1e-9)
        bw, bh = max(bw, 0.25 * big), max(bh, 0.25 * big)
        self.cx, self.cy = (x0 + x1) / 2.0, (y0 + y1) / 2.0
        self.s = min((self.Width - 40.0) / (bw * 1.5), (self.Height - 40.0) / (bh * 1.5))
        self.px = self.py = 0.0

    def fit(self):
        if self.mode == u"image":
            if self.image is not None:
                self.s = min((self.Width - 20.0) / self.image.Height,
                             (self.Height - 20.0) / self.image.Width)
                self.px = self.py = 0.0
        elif self.segs:
            xs, ys = [], []
            for g in self.segs:
                for (x, y) in ((g["ax"], g["ay"]), (g["bx"], g["by"])):
                    xs.append(y)
                    ys.append(-x)
            self._fit_box((min(xs), max(xs), min(ys), max(ys)))
            self.rescale()
        else:
            self.fit_all()
        self.Invalidate()

    def fit_all(self):
        if self.scene is not None:
            self._fit_box(self.scene["box"])
        self.Invalidate()

    def zoom(self, f):
        self.s = max(1e-9, self.s * f)
        self.Invalidate()

    def rescale(self):
        self.arrow = 34.0 / self.s
        self.Invalidate()

    # ----- painting -----
    def _paint(self, sender, e):
        g = e.Graphics
        g.Clear(Color.White)
        g.SmoothingMode = SmoothingMode.AntiAlias
        if self.mode == u"image":
            if self.image is not None:
                g.TranslateTransform(self.Width / 2.0 + self.px, self.Height / 2.0 + self.py)
                g.RotateTransform(self.theta + 90.0)
                g.ScaleTransform(self.s, self.s)
                g.DrawImage(self.image, -self.image.Width / 2.0, -self.image.Height / 2.0,
                            float(self.image.Width), float(self.image.Height))
                g.ResetTransform()
        elif self.scene is not None:
            g.Transform = self._matrix()
            inv = 1.0 / self.s
            for b in self.scene["bodies"]:
                if not b["visible"]:
                    continue
                br = SolidBrush(b["color"])
                pen = Pen(Color.FromArgb(120, 120, 120) if self.mesh_on else b["color"], inv)
                for el in b["elems"]:
                    if el[3].Length >= 3:
                        g.FillPolygon(br, el[3])
                        g.DrawPolygon(pen, el[3])
            self._draw_segments(g, inv)
            g.ResetTransform()
        # overlay
        f = Font("Segoe UI", 8.0)
        txt = u"%s view | view from -Z (X up, Y right) | rot %.0f deg | selected %d" % (
            u"Mechanical" if self.mode == u"image" else u"Model", self.theta % 360, len(self.sel))
        g.DrawString(txt, f, Brushes_Gray(), 6.0, 4.0)
        if self._rb is not None:
            x0, y0, x1, y1 = self._rb
            pen = Pen(Color.DarkOrange, 1.0)
            pen.DashStyle = DashStyle.Dash
            g.DrawRectangle(pen, min(x0, x1), min(y0, y1), abs(x1 - x0), abs(y1 - y0))

    def _draw_segments(self, g, inv):
        red, blue = Color.Crimson, Color.RoyalBlue
        L = self.arrow
        for i, sg in enumerate(self.segs):
            col = red if sg["side"] == "contact" else blue
            A = PointF(sg["ay"], -sg["ax"])
            B = PointF(sg["by"], -sg["bx"])
            M = PointF(sg["my"], -sg["mx"])
            dx, dy = sg["ny"], -sg["nx"]                       # normal in draw coords
            T = PointF(M.X + dx * L, M.Y + dy * L)
            if i in self.sel:
                hp = Pen(Color.Gold, 9.0 * inv)
                g.DrawLine(hp, A, B)
                g.DrawLine(Pen(Color.Gold, 6.0 * inv), M, T)
            g.DrawLine(Pen(col, 3.0 * inv), A, B)
            ap = Pen(col, 1.8 * inv)
            g.DrawLine(ap, M, T)
            for ang in (150.0, -150.0):
                r = math.radians(ang)
                bx = dx * math.cos(r) - dy * math.sin(r)
                by = dx * math.sin(r) + dy * math.cos(r)
                g.DrawLine(ap, T, PointF(T.X + bx * L * 0.3, T.Y + by * L * 0.3))

    # ----- interaction -----
    def _wheel(self, sender, e):
        step = 1 if e.Delta > 0 else -1
        mods = Control.ModifierKeys
        ctrl = (mods & Keys.Control) == Keys.Control
        shift = (mods & Keys.Shift) == Keys.Shift
        alt = (mods & Keys.Alt) == Keys.Alt
        if alt:
            self.zoom(1.2 if step > 0 else 1 / 1.2)
            return
        if ctrl and shift:
            self.px += 40 * step
        elif ctrl:
            self.py += 40 * step
        else:
            self.theta += 5.0 * step
        self.Invalidate()

    def _mdown(self, sender, e):
        if e.Button == MouseButtons.Left:
            self._down = (e.X, e.Y, self.px, self.py)
            self._drag = False

    def _mmove(self, sender, e):
        if self._down is None:
            return
        x0, y0, p0x, p0y = self._down
        if not self._drag and math.hypot(e.X - x0, e.Y - y0) > 4:
            self._drag = True
        if not self._drag:
            return
        if (Control.ModifierKeys & Keys.Shift) == Keys.Shift and self.mode == u"model":
            self._rb = (x0, y0, e.X, e.Y)
        else:
            self.px = p0x + (e.X - x0)
            self.py = p0y + (e.Y - y0)
        self.Invalidate()

    def _mup(self, sender, e):
        if self._down is None:
            return
        x0, y0 = self._down[0], self._down[1]
        if self.mode == u"model" and self.segs:
            if self._drag and self._rb is not None:
                xa, xb = sorted((x0, e.X))
                ya, yb = sorted((y0, e.Y))
                pts = self._to_screen([(g["my"], -g["mx"]) for g in self.segs])
                for i, (sx, sy) in enumerate(pts):
                    if xa <= sx <= xb and ya <= sy <= yb:
                        self.sel.add(i)
                self.owner.selection_changed()
            elif not self._drag:
                self._pick(x0, y0)
        self._down = None
        self._drag = False
        self._rb = None
        self.Invalidate()

    def _pick(self, mx, my):
        best, bi = 12.0, None
        for i, g in enumerate(self.segs):
            dx, dy = g["ny"], -g["nx"]
            pts = self._to_screen([(g["ay"], -g["ax"]), (g["by"], -g["bx"]), (g["my"], -g["mx"]),
                                   (g["my"] + dx * self.arrow, -g["mx"] + dy * self.arrow)])
            d = min(_dist_pt_seg(mx, my, pts[0][0], pts[0][1], pts[1][0], pts[1][1]),
                    _dist_pt_seg(mx, my, pts[2][0], pts[2][1], pts[3][0], pts[3][1]))
            if d < best:
                best, bi = d, i
        if bi is not None:
            if bi in self.sel:
                self.sel.discard(bi)
            else:
                self.sel.add(bi)
            self.owner.selection_changed()


def Brushes_Gray():
    return SolidBrush(Color.DimGray)


class BodiesForm(Form):
    def __init__(self, scene, view):
        self.Text = u"Show / Hide geometry"
        self.ClientSize = Size(380, 300)
        self.StartPosition = FormStartPosition.CenterParent
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        self._scene = scene
        self._view = view
        lb = CheckedListBox()
        lb.CheckOnClick = True
        lb.Location = Point(10, 10)
        lb.Size = Size(360, 240)
        for b in scene["bodies"]:
            lb.Items.Add(b["name"], b["visible"])
        lb.ItemCheck += self._on_check
        self.Controls.Add(lb)
        bt = Button()
        bt.Text = u"Close"
        bt.Location = Point(280, 260)
        bt.Size = Size(90, 28)
        bt.DialogResult = DialogResult.OK
        self.Controls.Add(bt)

    def _on_check(self, s, e):
        self._scene["bodies"][e.Index]["visible"] = (e.NewValue == CheckState.Checked)
        self._view.Invalidate()


# ---------------------------------------------------------------
# UI
# ---------------------------------------------------------------
class NormalForm(Form):
    def __init__(self, contact_tree):
        self.Text = u"Contact Normal Check"
        self.ClientSize = Size(1320, 740)
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        self._tree = contact_tree
        self._sel_keys = set()
        self._cds = []
        self._scene = None
        base = Form().Font
        self.base_font = base
        self.font_bold = Font(base, FontStyle.Bold)
        self.font_title = Font(base.FontFamily, 9, FontStyle.Bold)
        self._build_select_panel()

    def _btn(self, text, x, y, w, h, handler, bold=False):
        b = Button()
        b.Text = text
        b.Font = self.font_bold if bold else self.base_font
        b.FlatStyle = FlatStyle.Standard
        b.Location = Point(x, y)
        b.Size = Size(w, h)
        if handler is not None:
            b.Click += handler
        return b

    # ---------------- panel 1 ----------------
    def _build_select_panel(self):
        self.Controls.Clear()
        self.Text = u"Contact Normal Check  -  Select Contacts"
        grp = GroupBox()
        grp.Text = u"Contacts  (check the ones whose normals you want to see)"
        grp.Font = self.font_title
        grp.Location = Point(14, 10)
        grp.Size = Size(1290, 670)
        self.Controls.Add(grp)
        self.tv = TreeView()
        self.tv.Location = Point(10, 24)
        self.tv.Size = Size(1270, 636)
        self.tv.CheckBoxes = True
        self.tv.ShowLines = True
        self.tv.AfterCheck += self._on_node_check
        grp.Controls.Add(self.tv)
        self._ordered, self._groups = [], []
        for (gname, contacts) in self._tree:
            gn = TreeNode(gname + u"  [" + str(len(contacts)) + u" contacts]")
            gn.Tag = u"GROUP"
            gn.NodeFont = self.font_bold
            self._groups.append(gn)
            for c in contacts:
                ch = TreeNode(label_of(c))
                ch.Tag = u"CONTACT"
                ch.NodeFont = self.base_font
                ch.Checked = (ukey_of(c, len(self._ordered)) in self._sel_keys)
                self._ordered.append((ch, c))
                gn.Nodes.Add(ch)
            self.tv.Nodes.Add(gn)
            gn.Expand()
        self._updating = False
        self.Controls.Add(self._btn(u"Select All", 14, 696, 100, 30, lambda s, e: self._set_all(True)))
        self.Controls.Add(self._btn(u"Clear", 124, 696, 80, 30, lambda s, e: self._set_all(False)))
        bc = self._btn(u"Cancel", 1070, 696, 100, 30, None)
        bc.DialogResult = DialogResult.Cancel
        self.Controls.Add(bc)
        self.CancelButton = bc
        self.Controls.Add(self._btn(u"Show Normals >", 1180, 696, 124, 30, self._on_show, True))

    def _on_node_check(self, sender, e):
        if self._updating:
            return
        self._updating = True
        node = e.Node
        if str(node.Tag) == u"GROUP":
            for ch in node.Nodes:
                ch.Checked = node.Checked
        elif node.Parent is not None:
            node.Parent.Checked = all(ch.Checked for ch in node.Parent.Nodes)
        self._updating = False

    def _set_all(self, v):
        self._updating = True
        for g in self._groups:
            g.Checked = v
            for ch in g.Nodes:
                ch.Checked = v
        self._updating = False

    def _on_show(self, s, e):
        sel = []
        for i, (n, c) in enumerate(self._ordered):
            if n.Checked:
                sel.append((ukey_of(c, i), c))
        if not sel:
            MessageBox.Show(u"Please select at least one contact.", u"Warning",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        self._sel_keys = set(k for k, c in sel)
        lines = []
        Cursor.Current = Cursors.WaitCursor
        old = self._cds
        self._cds = []
        try:
            if self._scene is None:
                try:
                    self._scene = load_scene(lines.append)
                except Exception as ex:
                    lines.append(u"Model view unavailable (mesh read failed): %s" % str(ex))
            outdir = System.IO.Path.Combine(System.IO.Path.GetTempPath(), u"contact_normal_check")
            System.IO.Directory.CreateDirectory(outdir)
            try:
                prev = list(Tree.ActiveObjects)
            except Exception:
                prev = []
            self.Hide()
            for i, (k, c) in enumerate(sel):
                cd = {"c": c, "key": k, "segs": [], "sel": set(), "side": u"both",
                      "shot": None, "scene": self._scene}
                if self._scene is not None:
                    try:
                        cd["segs"] = contact_segments(self._scene, c, lines.append)
                    except Exception as ex:
                        lines.append(u"%s: segment extraction failed (%s)" % (label_of(c), str(ex)))
                    if not cd["segs"]:
                        lines.append(u"%s: no contact/target segments found on the mesh" % label_of(c))
                cd["shot"] = capture_contact(cd, i, outdir, lines.append)
                cd["idx"] = i
                self._cds.append(cd)
            try:
                if prev:
                    Tree.Activate(prev)
            except Exception:
                pass
        finally:
            Cursor.Current = Cursors.Default
            self.Show()
        self._build_result_panel(lines)

    # ---------------- panel 2 ----------------
    def _build_result_panel(self, lines):
        self.Controls.Clear()
        self.Text = u"Contact Normal Check  -  Normals"

        g1 = GroupBox()
        g1.Text = u"Contacts  (check = FLIP)"
        g1.Font = self.font_title
        g1.Location = Point(14, 10)
        g1.Size = Size(380, 520)
        self.Controls.Add(g1)
        lv = ListView()
        lv.View = View.Details
        lv.CheckBoxes = True
        lv.FullRowSelect = True
        lv.HideSelection = False
        lv.GridLines = True
        lv.MultiSelect = False
        lv.Font = self.base_font
        lv.Location = Point(10, 24)
        lv.Size = Size(360, 440)
        lv.Columns.Add(u"Contact", 215)
        lv.Columns.Add(u"Side", 70)
        lv.Columns.Add(u"Sel", 45)
        for cd in self._cds:
            it = ListViewItem(label_of(cd["c"]))
            it.SubItems.Add(cd["side"])
            it.SubItems.Add(u"0")
            lv.Items.Add(it)
        lv.SelectedIndexChanged += self._on_pick
        g1.Controls.Add(lv)
        self.lv = lv
        lbl = Label()
        lbl.Text = u"Sides to flip (this contact):"
        lbl.Font = self.base_font
        lbl.Location = Point(10, 476)
        lbl.Size = Size(170, 20)
        g1.Controls.Add(lbl)
        self.cb_side = ComboBox()
        self.cb_side.DropDownStyle = ComboBoxStyle.DropDownList
        self.cb_side.Font = self.base_font
        for t in (u"both", u"contact", u"target"):
            self.cb_side.Items.Add(t)
        self.cb_side.SelectedIndex = 0
        self.cb_side.Location = Point(184, 473)
        self.cb_side.Size = Size(100, 24)
        self.cb_side.SelectedIndexChanged += self._on_side
        g1.Controls.Add(self.cb_side)

        g2 = GroupBox()
        g2.Text = u"Normals  (red = contact, blue = target)"
        g2.Font = self.font_title
        g2.Location = Point(404, 10)
        g2.Size = Size(900, 590)
        self.Controls.Add(g2)
        self.view = SceneView(self)
        self.view.scene = self._scene
        self.view.Location = Point(10, 24)
        self.view.Size = Size(880, 520)
        g2.Controls.Add(self.view)
        v = self.view
        row = [(u"Zoom +", lambda s, e: v.zoom(1.3)), (u"Zoom -", lambda s, e: v.zoom(1 / 1.3)),
               (u"Fit", lambda s, e: v.fit()), (u"Fit all", lambda s, e: v.fit_all()),
               (u"Mesh", self._on_mesh), (u"Rescale", self._on_rescale),
               (u"Bodies...", self._on_bodies), (u"Clear sel", self._on_clear_sel),
               (u"Reset rot", self._on_reset_rot), (u"Mechanical view", self._on_toggle_mode)]
        x = 10
        for t, h in row:
            w = 112 if t == u"Mechanical view" else 82
            g2.Controls.Add(self._btn(t, x, 552, w, 28, h))
            x += w + 6
            if t == u"Mechanical view":
                self.btn_mode = g2.Controls[g2.Controls.Count - 1]

        self.log = TextBox()
        self.log.Multiline = True
        self.log.ReadOnly = True
        self.log.ScrollBars = ScrollBars.Vertical
        self.log.Font = self.base_font
        self.log.Location = Point(14, 610)
        self.log.Size = Size(1290, 76)
        self.log.Text = u"\r\n".join(lines) + (u"\r\n" if lines else u"")
        self.Controls.Add(self.log)

        self.Controls.Add(self._btn(u"< Back", 14, 696, 90, 30, lambda s, e: self._build_select_panel()))
        bc = self._btn(u"Close", 1070, 696, 90, 30, None)
        bc.DialogResult = DialogResult.Cancel
        self.Controls.Add(bc)
        self.CancelButton = bc
        self.Controls.Add(self._btn(u"Flip Selected", 1180, 696, 124, 30, self._on_flip, True))

        self._cur = None
        if lv.Items.Count > 0:
            lv.Items[0].Selected = True

    def _cur_cd(self):
        if self.lv.SelectedIndices.Count == 0:
            return None
        return self._cds[self.lv.SelectedIndices[0]]

    def _on_pick(self, s, e):
        cd = self._cur_cd()
        if cd is None:
            return
        self._cur = cd
        self.view.segs = cd["segs"]
        self.view.sel = cd["sel"]
        self.view.image = cd["shot"]
        self.cb_side.SelectedItem = cd["side"]
        if self.view.mode == u"image" and cd["shot"] is None:
            self.view.mode = u"model"
            self.btn_mode.Text = u"Mechanical view"
        self.view.fit()

    def _on_side(self, s, e):
        cd = self._cur
        if cd is None:
            return
        cd["side"] = self.cb_side.SelectedItem
        self.lv.Items[cd["idx"]].SubItems[1].Text = cd["side"]

    def selection_changed(self):
        cd = self._cur
        if cd is not None:
            self.lv.Items[cd["idx"]].SubItems[2].Text = str(len(cd["sel"]))

    def _on_mesh(self, s, e):
        v = self.view
        if v.mode == u"image":
            self._mech_state = not getattr(self, "_mech_state", False)
            set_mesh_display(self._mech_state)
            self._recapture(False)
        else:
            v.mesh_on = not v.mesh_on
            v.Invalidate()

    def _on_rescale(self, s, e):
        if self.view.mode == u"image":
            self._recapture(True)
        else:
            self.view.rescale()

    def _recapture(self, rescale):
        cd = self._cur
        if cd is None:
            return
        outdir = System.IO.Path.Combine(System.IO.Path.GetTempPath(), u"contact_normal_check")
        lines = []
        Cursor.Current = Cursors.WaitCursor
        self.Hide()
        try:
            img = capture_contact(cd, cd["idx"], outdir, lines.append, rescale)
        finally:
            Cursor.Current = Cursors.Default
            self.Show()
        if img is not None:
            cd["shot"] = img
            self.view.image = img
        for m in lines:
            self.log.AppendText(m + u"\r\n")
        self.view.Invalidate()

    def _on_bodies(self, s, e):
        if self._scene is None:
            return
        BodiesForm(self._scene, self.view).ShowDialog()

    def _on_clear_sel(self, s, e):
        cd = self._cur
        if cd is not None:
            cd["sel"].clear()
            self.selection_changed()
            self.view.Invalidate()

    def _on_reset_rot(self, s, e):
        self.view.theta = 0.0
        self.view.Invalidate()

    def _on_toggle_mode(self, s, e):
        v = self.view
        if v.mode == u"model":
            if v.image is None:
                self.log.AppendText(u"No Mechanical image for this contact.\r\n")
                return
            v.mode = u"image"
            self.btn_mode.Text = u"Model view"
        else:
            v.mode = u"model"
            self.btn_mode.Text = u"Mechanical view"
        v.fit()

    def _on_flip(self, s, e):
        n = 0
        for cd in self._cds:
            c = cd["c"]
            it = self.lv.Items[cd["idx"]]
            old = find_flip_snippet(c)
            chosen = [cd["segs"][i] for i in sorted(cd["sel"]) if i < len(cd["segs"])]
            if chosen:
                text = build_flip_elems(c.Name, chosen)
                what = u"%d selected element(s)" % len(chosen)
            elif it.Checked:
                text = build_flip_whole(c.Name, cd["side"])
                what = u"whole side: %s" % cd["side"]
            else:
                if old is not None:
                    old.Delete()
                    self.log.AppendText(u"Flip command removed: %s\r\n" % label_of(c))
                continue
            if old is not None:
                old.Delete()
            sn = c.AddCommandSnippet()
            sn.Name = SNIPPET_NAME
            sn.Input = text
            self.log.AppendText(u"Flip command written (%s): %s\r\n" % (what, label_of(c)))
            n += 1
        if n:
            self.log.AppendText(u"Done. Flip is applied at the next solve; it is not reflected "
                                u"in either view. Uses cid / tid of each contact.\r\n")


# ---------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------
model = ExtAPI.DataModel.Project.Model
contact_tree = build_contact_tree(model.Connections)
if sum(len(v) for (_, v) in contact_tree) == 0:
    MessageBox.Show(u"No contacts found in the model.", u"Info",
                    MessageBoxButtons.OK, MessageBoxIcon.Information)
else:
    NormalForm(contact_tree).ShowDialog()
