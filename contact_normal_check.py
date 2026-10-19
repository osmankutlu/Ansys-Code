# -*- coding: utf-8 -*-
# =====================================================================
#  Contact Normal Viewer / Flip   (Ansys Mechanical 2D, IronPython 2.7)
#
#  1) Tum kontaklar listelenir, istediklerini isaretlersin
#  2) Her kontak icin "Element Normals = Yes" acilir, Mechanical grafik
#     penceresinden goruntu alinip popup'ta gosterilir (dat gerekmez)
#  3) Ters gordugun kontaklari isaretle -> kontagin altina
#     ESURF,,REVERSE iceren Command eklenir (cid / tid parametreleri ile)
# =====================================================================
import clr
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
import System
import re
from System.Windows.Forms import (Form, Button, Label, MessageBox, MessageBoxButtons,
                                  MessageBoxIcon, FormStartPosition, FormBorderStyle,
                                  DialogResult, GroupBox, FlatStyle, TextBox, TreeView,
                                  TreeNode, Application, ComboBox, ComboBoxStyle,
                                  ListView, ListViewItem, View, ScrollBars, CheckBox,
                                  PictureBox, PictureBoxSizeMode, Panel, Cursor, Cursors)
from System.Drawing import Size, Point, Font, FontStyle, Color, Bitmap

SNIPPET_NAME = u"ContactNormalFlip"
IMG_W, IMG_H = 1800, 1100


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


def flip_side_of(contact):
    sn = find_flip_snippet(contact)
    if sn is None:
        return None
    try:
        m = re.findall(r"!\s*FLIP side=(\w+)", sn.Input)
        return m[0] if m else u"both"
    except Exception:
        return u"both"


def build_flip_snippet(name, side):
    L = [u"! ContactNormalCheck - normal flip (ESURF,,REVERSE)",
         u"! Contact: %s" % name,
         u"! FLIP side=%s" % side]
    if side in (u"both", u"contact"):
        L += [u"esel,s,type,,cid", u"esurf,,reverse"]
    if side in (u"both", u"target"):
        L += [u"esel,s,type,,tid", u"esurf,,reverse"]
    L.append(u"allsel,all")
    return u"\n".join(L) + u"\n"


# ---------------------------------------------------------------
# Graphics: Element Normals on/off + camera zoom + image export
# ---------------------------------------------------------------
def normal_props(c):
    names = [n for n in dir(c) if "normal" in n.lower() and not n.startswith("_")]
    pri = [n for n in names if "element" in n.lower()]
    rest = [n for n in names if n not in pri and
            not any(k in n.lower() for k in ("stiff", "updat", "factor", "toler"))]
    return pri + rest, names


def set_normals(c, want, log):
    """Turns the 'Element Normals' display on/off. Returns (prop_name, old_value) or (None, None)."""
    order, allnames = normal_props(c)
    vals = []
    for n in allnames:
        try:
            vals.append(u"%s=%s" % (n, getattr(c, n)))
        except Exception:
            vals.append(n)
    log(u"normal-related properties: " + u", ".join(vals))
    for n in order:
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
            log(u"Element Normals: '%s' %s -> %s" % (n, before, after))
            if after != before:
                return n, cur
    return None, None


def restore_normals(c, name, old):
    if name is None:
        return
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


def mesh_boxes(c):
    """(contact_box, model_box) as (xmin,xmax,ymin,ymax) from the mesh nodes."""
    mesh = ExtAPI.DataModel.MeshDataByName("Global")
    ids = list(c.SourceLocation.Ids) + list(c.TargetLocation.Ids)
    xs, ys = [], []
    for gid in ids:
        reg = mesh.MeshRegionById(gid)
        for nid in reg.NodeIds:
            n = mesh.NodeById(nid)
            xs.append(n.X)
            ys.append(n.Y)
    cb = (min(xs), max(xs), min(ys), max(ys))
    if not _MODEL_BOX:
        mx, my = [], []
        for n in mesh.Nodes:
            mx.append(n.X)
            my.append(n.Y)
        _MODEL_BOX.append((min(mx), max(mx), min(my), max(my)))
    return cb, _MODEL_BOX[0]


def zoom_to_contact(c, log):
    import math
    cam = ExtAPI.Graphics.Camera
    cam.SetFit()
    cb, mb = mesh_boxes(c)
    H0 = cam.SceneHeight
    F0 = cam.FocalPoint
    h0 = _val(H0)
    mw, mh = mb[1] - mb[0], mb[3] - mb[2]
    cw, ch = cb[1] - cb[0], cb[3] - cb[2]
    ext = max(mh, mw / 1.6)
    if ext <= 0:
        raise Exception("model box is empty")
    best, u = None, 1.0
    for cand in (1.0, 1e-3, 1e3, 1e-6, 1e6, 0.0254, 39.3701):
        r = abs(math.log(h0 / (cand * ext)))
        if best is None or r < best:
            best, u = r, cand
    tgt = max(ch * 1.8, cw * 1.8 / 1.6, ext * 0.02)
    new_h = h0 * tgt / ext
    ccx, ccy = (cb[0] + cb[1]) / 2.0, (cb[2] + cb[3]) / 2.0
    mcx, mcy = (mb[0] + mb[1]) / 2.0, (mb[2] + mb[3]) / 2.0
    fx, fy, fz = _val(F0.X), _val(F0.Y), _val(F0.Z)
    cam.FocalPoint = type(F0)(fx + u * (ccx - mcx), fy + u * (ccy - mcy), fz)
    cam.SceneHeight = _like(H0, new_h)
    log(u"zoom applied (unit factor %g, scene height %g -> %g)" % (u, h0, new_h))


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
    data = System.IO.File.ReadAllBytes(path)
    ms = System.IO.MemoryStream(data)
    b = Bitmap(ms)
    img = Bitmap(b)
    b.Dispose()
    ms.Dispose()
    return img


def capture_contact(c, idx, auto_normals, zoom_geo, outdir, log):
    """Returns (bitmap or None, message)."""
    name, old = (None, None)
    L = lambda m: log(u"%s: %s" % (label_of(c), m))
    try:
        Tree.Activate([c])
        if auto_normals:
            name, old = set_normals(c, True, L)
            if name is None:
                L(u"could not switch Element Normals on - set Display > Element Normals = Yes "
                  u"manually and run again with the option unchecked.")
        try:
            ExtAPI.Graphics.Camera.SetSpecificViewOrientation(ViewOrientationType.Front)
        except Exception:
            pass
        zoomed = False
        if zoom_geo:
            try:
                zoom_to_contact(c, L)
                zoomed = True
            except Exception as ex:
                L(u"zoom failed (%s) - using full fit" % str(ex))
        if not zoomed:
            try:
                ExtAPI.Graphics.Camera.SetFit()
            except Exception:
                pass
        redraw()
        path = System.IO.Path.Combine(outdir, u"contact_%d.png" % idx)
        if System.IO.File.Exists(path):
            System.IO.File.Delete(path)
        export_image(path)
        return load_bitmap(path), u"ok"
    except Exception as ex:
        return None, u"export failed: %s" % str(ex)
    finally:
        restore_normals(c, name, old)


# ---------------------------------------------------------------
# UI
# ---------------------------------------------------------------
class NormalForm(Form):
    def __init__(self, contact_tree):
        self.Text = u"Contact Normal Check"
        self.ClientSize = Size(1180, 680)
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        self._tree = contact_tree
        self._sel_keys = set()
        self._shots = []
        self._orig = None
        self._zoom = 1.0
        base = Form().Font
        self.base_font = base
        self.font_bold = Font(base, FontStyle.Bold)
        self.font_title = Font(base.FontFamily, 9, FontStyle.Bold)
        self._opt_auto = True
        self._opt_zoom = True
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

    # ---------------- panel 1 : contact list ----------------
    def _build_select_panel(self):
        self.Controls.Clear()
        self.Text = u"Contact Normal Check  -  Select Contacts"

        grp = GroupBox()
        grp.Text = u"Contacts  (check the ones whose normals you want to see)"
        grp.Font = self.font_title
        grp.Location = Point(14, 10)
        grp.Size = Size(1150, 520)
        self.Controls.Add(grp)

        self.tv = TreeView()
        self.tv.Location = Point(10, 24)
        self.tv.Size = Size(1130, 486)
        self.tv.CheckBoxes = True
        self.tv.ShowLines = True
        self.tv.AfterCheck += self._on_node_check
        grp.Controls.Add(self.tv)

        self._ordered = []
        self._groups = []
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

        self.chk_auto = CheckBox()
        self.chk_auto.Text = u"Turn on 'Element Normals' for the capture (switched back afterwards)"
        self.chk_auto.Font = self.base_font
        self.chk_auto.Checked = self._opt_auto
        self.chk_auto.Location = Point(20, 540)
        self.chk_auto.Size = Size(520, 22)
        self.Controls.Add(self.chk_auto)

        self.chk_zoom = CheckBox()
        self.chk_zoom.Text = u"Zoom camera onto each contact (experimental)"
        self.chk_zoom.Font = self.base_font
        self.chk_zoom.Checked = self._opt_zoom
        self.chk_zoom.Location = Point(20, 564)
        self.chk_zoom.Size = Size(520, 22)
        self.Controls.Add(self.chk_zoom)

        self.Controls.Add(self._btn(u"Select All", 14, 630, 100, 30, self._on_select_all))
        self.Controls.Add(self._btn(u"Clear", 124, 630, 80, 30, self._on_clear))
        bc = self._btn(u"Cancel", 930, 630, 100, 30, None)
        bc.DialogResult = DialogResult.Cancel
        self.Controls.Add(bc)
        self.CancelButton = bc
        self.Controls.Add(self._btn(u"Show Normals >", 1040, 630, 124, 30, self._on_show, True))

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

    def _on_select_all(self, s, e):
        self._set_all(True)

    def _on_clear(self, s, e):
        self._set_all(False)

    def _on_show(self, s, e):
        self._sel = []
        self._uk = {}
        for i, (n, c) in enumerate(self._ordered):
            if n.Checked:
                self._sel.append(c)
                self._uk[id(c)] = ukey_of(c, i)
        self._sel_keys = set(self._uk.values())
        self._opt_auto = self.chk_auto.Checked
        self._opt_zoom = self.chk_zoom.Checked
        if not self._sel:
            MessageBox.Show(u"Please select at least one contact.", u"Warning",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        outdir = System.IO.Path.Combine(System.IO.Path.GetTempPath(), u"contact_normal_check")
        System.IO.Directory.CreateDirectory(outdir)
        try:
            prev = list(Tree.ActiveObjects)
        except Exception:
            prev = []
        lines = []
        shots = []
        self.Hide()
        Cursor.Current = Cursors.WaitCursor
        try:
            for i, c in enumerate(self._sel):
                img, msg = capture_contact(c, i, self._opt_auto, self._opt_zoom,
                                           outdir, lines.append)
                if img is None:
                    lines.append(u"%s: %s" % (label_of(c), msg))
                shots.append(img)
            try:
                if prev:
                    Tree.Activate(prev)
            except Exception:
                pass
        finally:
            Cursor.Current = Cursors.Default
            self.Show()
        self._shots = shots
        self._build_result_panel(lines)

    # ---------------- panel 2 : images + flip ----------------
    def _build_result_panel(self, lines):
        self.Controls.Clear()
        self.Text = u"Contact Normal Check  -  Normals"

        g1 = GroupBox()
        g1.Text = u"Contacts  (check = FLIP)"
        g1.Font = self.font_title
        g1.Location = Point(14, 10)
        g1.Size = Size(360, 470)
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
        lv.Size = Size(340, 436)
        lv.Columns.Add(u"Contact", 235)
        lv.Columns.Add(u"Img", 40)
        lv.Columns.Add(u"Cmd", 55)
        for i, c in enumerate(self._sel):
            it = ListViewItem(label_of(c))
            it.SubItems.Add(u"ok" if self._shots[i] is not None else u"-")
            side = flip_side_of(c)
            it.SubItems.Add(side if side else u"")
            it.Checked = side is not None
            lv.Items.Add(it)
        lv.SelectedIndexChanged += self._on_pick
        g1.Controls.Add(lv)
        self.lv = lv

        g2 = GroupBox()
        g2.Text = u"Element normals  (red = contact, blue = target)"
        g2.Font = self.font_title
        g2.Location = Point(384, 10)
        g2.Size = Size(780, 470)
        self.Controls.Add(g2)

        self.panel = Panel()
        self.panel.AutoScroll = True
        self.panel.BackColor = Color.White
        self.panel.Location = Point(10, 24)
        self.panel.Size = Size(760, 400)
        self.panel.MouseEnter += lambda s, e: self.panel.Focus()
        self.panel.MouseWheel += self._on_wheel
        g2.Controls.Add(self.panel)

        self.pic = PictureBox()
        self.pic.SizeMode = PictureBoxSizeMode.StretchImage
        self.pic.Location = Point(0, 0)
        self.pic.Size = Size(10, 10)
        self.panel.Controls.Add(self.pic)

        g2.Controls.Add(self._btn(u"Zoom +", 10, 432, 80, 26, lambda s, e: self._set_zoom(self._zoom * 1.4)))
        g2.Controls.Add(self._btn(u"Zoom -", 96, 432, 80, 26, lambda s, e: self._set_zoom(self._zoom / 1.4)))
        g2.Controls.Add(self._btn(u"Fit", 182, 432, 80, 26, lambda s, e: self._fit()))
        g2.Controls.Add(self._btn(u"100 %", 268, 432, 80, 26, lambda s, e: self._set_zoom(1.0)))

        lbl = Label()
        lbl.Text = u"Sides to flip:"
        lbl.Font = self.base_font
        lbl.Location = Point(16, 494)
        lbl.Size = Size(90, 20)
        self.Controls.Add(lbl)
        self.cb_side = ComboBox()
        self.cb_side.DropDownStyle = ComboBoxStyle.DropDownList
        self.cb_side.Font = self.base_font
        self.cb_side.Items.Add(u"Both (contact + target)")
        self.cb_side.Items.Add(u"Contact side only")
        self.cb_side.Items.Add(u"Target side only")
        self.cb_side.SelectedIndex = 0
        self.cb_side.Location = Point(110, 491)
        self.cb_side.Size = Size(220, 24)
        self.Controls.Add(self.cb_side)

        self.log = TextBox()
        self.log.Multiline = True
        self.log.ReadOnly = True
        self.log.ScrollBars = ScrollBars.Vertical
        self.log.Font = self.base_font
        self.log.Location = Point(14, 524)
        self.log.Size = Size(1150, 90)
        self.log.Text = u"\r\n".join(lines) + (u"\r\n" if lines else u"")
        self.Controls.Add(self.log)

        self.Controls.Add(self._btn(u"< Back", 14, 630, 90, 30, self._on_back))
        bc = self._btn(u"Close", 940, 630, 90, 30, None)
        bc.DialogResult = DialogResult.Cancel
        self.Controls.Add(bc)
        self.CancelButton = bc
        self.Controls.Add(self._btn(u"Flip Selected", 1040, 630, 124, 30, self._on_flip, True))

        if lv.Items.Count > 0:
            lv.Items[0].Selected = True

    def _on_pick(self, s, e):
        if self.lv.SelectedIndices.Count == 0:
            return
        img = self._shots[self.lv.SelectedIndices[0]]
        self._orig = img
        if img is None:
            self.pic.Image = None
            return
        self.pic.Image = img
        self._fit()

    def _fit(self):
        if self._orig is None:
            return
        z = min(float(self.panel.ClientSize.Width) / self._orig.Width,
                float(self.panel.ClientSize.Height) / self._orig.Height)
        self._set_zoom(z)

    def _set_zoom(self, z):
        if self._orig is None:
            return
        z = max(0.05, min(z, 6.0))
        self._zoom = z
        self.pic.Size = Size(int(self._orig.Width * z), int(self._orig.Height * z))

    def _on_wheel(self, s, e):
        self._set_zoom(self._zoom * (1.2 if e.Delta > 0 else 1.0 / 1.2))

    def _on_back(self, s, e):
        self._build_select_panel()

    def _on_flip(self, s, e):
        side = (u"both", u"contact", u"target")[self.cb_side.SelectedIndex]
        n_set = 0
        for i, c in enumerate(self._sel):
            it = self.lv.Items[i]
            old = find_flip_snippet(c)
            if it.Checked:
                if old is not None:
                    old.Delete()
                sn = c.AddCommandSnippet()
                sn.Name = SNIPPET_NAME
                sn.Input = build_flip_snippet(c.Name, side)
                it.SubItems[2].Text = side
                self.log.AppendText(u"Flip command written (%s): %s\r\n" % (side, label_of(c)))
                n_set += 1
            elif old is not None:
                old.Delete()
                it.SubItems[2].Text = u""
                self.log.AppendText(u"Flip command removed: %s\r\n" % label_of(c))
        if n_set:
            self.log.AppendText(u"Done. Flip is applied at the next solve (the Element Normals "
                                u"display does not show it). Uses cid / tid of each contact.\r\n")


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
