# -*- coding: utf-8 -*-
# =============================================================================
# Averaged / Unaveraged % Check - ANSYS Mechanical (IronPython, RST/DPF)
#
#   - Gerilmeler dogrudan RST'den DPF ile okunur (ElementalNodal = unaveraged).
#   - Coklu Named Selection + coklu step secimi.
#   - Her eleman-node noktasinda EQV / S1 / S3 hesaplanir, node bazinda:
#         Averaged   = o node'a bagli eleman degerlerinin ortalamasi
#         Unaveraged = o node'daki |max| eleman degeri
#         %          = abs((Unaveraged - Averaged) / Unaveraged) * 100
#   - Her (NS, Step) icin bir DPF Python Result olusturulup plot edilir.
# =============================================================================
import os
import math
import time
import clr
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
from System.Windows.Forms import (
    Form, Label, Button, CheckBox, GroupBox, ComboBox, ListBox, SelectionMode,
    TextBox, OpenFileDialog, DialogResult, FormStartPosition, FormBorderStyle,
    ComboBoxStyle, MessageBox, MessageBoxButtons, MessageBoxIcon, Application,
    FlatStyle
)
from System.Drawing import Size, Point, Color, Font, FontStyle
from System import Int32
from System.Collections.Generic import List as _NetList

import mech_dpf
import Ans.DataProcessing as dpf
mech_dpf.setExtAPI(ExtAPI)

CTRL_FONT = Font("Segoe UI", 9)
HDR_FONT  = Font("Segoe UI", 9, FontStyle.Bold)
BTN_BLUE  = Color.FromArgb(0, 120, 215)

STRESS_TYPES = [
    ("EQV", "Equivalent (von Mises)"),
    ("S1",  "Maximum Principal"),
    ("S3",  "Minimum Principal"),
]


# =============================================================================
# GENERAL HELPERS
# =============================================================================
def safe_str(o):
    if o is None: return ""
    try: return str(o)
    except:
        try: return "".join(c if ord(c) < 128 else "?" for c in unicode(o))
        except: return "?"

def ascii_tag(name, maxlen=25):
    s = "".join(c if (c.isalnum() and ord(c) < 128) else "_" for c in safe_str(name))
    return s[:maxlen]

def fmt_duration(sec):
    s = int(sec)
    return "{0} sec".format(s) if s < 60 else "{0} min {1} sec".format(s // 60, s % 60)

def _to_net_int_list(ids):
    lst = _NetList[Int32]()
    for i in ids: lst.Add(int(i))
    return lst

def _no_scroll(cmb):
    def _b(s, e):
        try: e.Handled = True
        except: pass
    cmb.MouseWheel += _b
    return cmb

def _btn(text, w, h, blue=False):
    b = Button(); b.Text = text; b.Size = Size(w, h); b.Font = CTRL_FONT
    if blue:
        b.Font = HDR_FONT; b.BackColor = BTN_BLUE
        b.ForeColor = Color.White; b.FlatStyle = FlatStyle.Flat
    return b

def _lbl(text, x, y, w, bold=True):
    l = Label(); l.Text = text; l.Location = Point(x, y)
    l.Size = Size(w, 20); l.Font = HDR_FONT if bold else CTRL_FONT
    return l


# =============================================================================
# DPF BACKEND  (Force/Disp/Temp aracindaki kalip)
# =============================================================================
def dpf_load(rst_path):
    ds = dpf.DataSources(rst_path)
    return ds, dpf.Model(ds)

def get_load_step_info(dpf_model):
    """[(load_step_no, last_set_id, time), ...]  — her load step'in son seti."""
    tf = dpf_model.TimeFreqSupport
    n_sets = int(tf.NumberSets)
    try:    tvals = list(tf.TimeFreqs.Data)
    except: tvals = [0.0] * n_sets

    try:
        step_ids = [int(x) for x in tf.StepIds.Data]
        if len(step_ids) == n_sets:
            seen = {}
            for i in range(n_sets):
                seen[step_ids[i]] = (i + 1, float(tvals[i]) if i < len(tvals) else 0.0)
            return sorted([(ls, sid, t) for ls, (sid, t) in seen.items()])
    except: pass

    try:
        nps = [int(x) for x in tf.NumberSetsPerStep.Data]
        out = []; cum = 0
        for k, c in enumerate(nps):
            cum += c
            out.append((k + 1, cum, float(tvals[cum - 1]) if cum - 1 < len(tvals) else 0.0))
        return out
    except: pass

    return [(i + 1, i + 1, float(tvals[i]) if i < len(tvals) else 0.0) for i in range(n_sets)]

def fc_count(fc):
    try: return fc.FieldCount
    except:
        try: return len(fc)
        except: return 1

def fc_get_field(fc, i):
    try: return fc[i]
    except: return fc.GetFieldByIndex(i)

def get_mesh(ds, dpf_model):
    try: return dpf_model.Mesh
    except: pass
    try:
        mop = dpf.operators.mesh.mesh_provider()
        mop.inputs.data_sources.Connect(ds)
        return mop.outputs.mesh.GetData()
    except: pass
    try:
        mop = dpf.Operator("mapdl::rst::MeshProvider")
        mop.Connect(4, ds)
        return mop.GetOutputAsMeshedRegion(0)
    except: pass
    return None

def element_node_ids(mesh, eid):
    try: el = mesh.ElementById(eid)
    except: return None
    if el is None: return None
    try: return list(el.NodeIds)
    except: pass
    try: return [n.Id for n in el.Nodes]
    except: pass
    try: return list(el.CornerNodeIds)
    except: pass
    return None

def ns_elem_ids(ns):
    """NS (body/face/element) -> mesh eleman ID listesi."""
    try:
        loc = ns.Location
        st = str(loc.SelectionType)
    except:
        return []
    out = set()
    if "MeshElement" in st:
        for i in loc.Ids: out.add(int(i))
        return sorted(out)
    try:
        md = ExtAPI.DataModel.MeshDataByName("Global")
        for gid in loc.Ids:
            try:
                for e in md.MeshRegionById(gid).ElementIds: out.add(int(e))
            except: pass
    except: pass
    return sorted(out)

def read_stress_elemnodal(ds, set_id, elem_ids):
    op = dpf.operators.result.stress()
    op.inputs.data_sources.Connect(ds)
    op.inputs.time_scoping.Connect([set_id])
    sc = dpf.Scoping(); sc.Location = "Elemental"
    try:    sc.Ids = _to_net_int_list(elem_ids)
    except: sc.Ids = elem_ids
    op.inputs.mesh_scoping.Connect(sc)
    try: op.inputs.requested_location.Connect("ElementalNodal")
    except: pass
    fc = op.outputs.fields_container.GetData()

    out = {}   # eid -> (values, comp)
    for fi in range(fc_count(fc)):
        f = fc_get_field(fc, fi)
        try:    comp = int(f.ComponentCount)
        except: comp = 6
        for eid in list(f.ScopingIds):
            if eid in out: continue
            try: out[eid] = (list(f.GetEntityDataById(eid)), comp)
            except: pass
    return out


# =============================================================================
# STRESS SCALARS  (DPF bilesen sirasi: XX YY ZZ XY YZ XZ)
# =============================================================================
def s_eqv(sx, sy, sz, sxy, syz, sxz):
    return math.sqrt(0.5 * ((sx - sy) ** 2 + (sy - sz) ** 2 + (sz - sx) ** 2)
                     + 3.0 * (sxy * sxy + syz * syz + sxz * sxz))

def s_principal(sx, sy, sz, sxy, syz, sxz):
    """(S1, S3) — trigonometrik ozdeger cozumu."""
    m = (sx + sy + sz) / 3.0
    a, b, c = sx - m, sy - m, sz - m
    J2 = 0.5 * (a * a + b * b + c * c) + sxy * sxy + syz * syz + sxz * sxz
    if J2 < 1e-30:
        return m, m
    J3 = (a * b * c + 2.0 * sxy * syz * sxz
          - a * syz * syz - b * sxz * sxz - c * sxy * sxy)
    r = math.sqrt(J2 / 3.0)
    cos3 = max(-1.0, min(1.0, J3 / (2.0 * r ** 3)))
    th = math.acos(cos3) / 3.0
    return m + 2.0 * r * math.cos(th), m + 2.0 * r * math.cos(th + 2.0 * math.pi / 3.0)

def compute_pct_multi(elem_ids, edata, mesh, keys, conn_cache):
    """Secili tum stress tipleri icin tek geciste hesaplar.
    Donen: {key: (pct, avg, una)}"""
    need_p = ("S1" in keys) or ("S3" in keys)
    node_vals = dict((k, {}) for k in keys)
    for eid in elem_ids:
        d = edata.get(eid)
        if d is None: continue
        vals, comp = d
        if comp < 6: continue
        nids = conn_cache.get(eid)
        if nids is None:
            nids = element_node_ids(mesh, eid) or []
            conn_cache[eid] = nids
        # Quadratic elemanlarda rst sadece corner node degerlerini tutar;
        # node listesinde corner'lar once geldigi icin ilk n tanesi eslesir.
        n = min(len(vals) // comp, len(nids))
        for j in range(n):
            b = j * comp
            v = vals[b:b + 6]
            nid = nids[j]
            if "EQV" in keys:
                node_vals["EQV"].setdefault(nid, []).append(
                    s_eqv(v[0], v[1], v[2], v[3], v[4], v[5]))
            if need_p:
                p1, p3 = s_principal(v[0], v[1], v[2], v[3], v[4], v[5])
                if "S1" in keys: node_vals["S1"].setdefault(nid, []).append(p1)
                if "S3" in keys: node_vals["S3"].setdefault(nid, []).append(p3)

    out = {}
    for k in keys:
        pct, avg, una = {}, {}, {}
        for nid, lst in node_vals[k].items():
            a = sum(lst) / len(lst)
            u = max(lst, key=abs)
            avg[nid] = a; una[nid] = u
            pct[nid] = abs((u - a) / u) * 100.0 if abs(u) > 1e-12 else 0.0
        out[k] = (pct, avg, una)
    return out


# =============================================================================
# PYTHON RESULT  (forward ile plot)
# =============================================================================
def build_dpf_code(node_vals, result_name):
    pairs = list(node_vals.items()); chunk = 200; lines = []
    for i in range(0, len(pairs), chunk):
        seg = ", ".join("({0}, {1:.4f})".format(k, v) for k, v in pairs[i:i + chunk])
        lines.append("    for _k, _v in [{0}]: node_vals[_k] = _v".format(seg))
    wf_key = "avgpct_" + ascii_tag(result_name, 45)
    code = [
        "# Averaged/Unaveraged % - " + result_name,
        "def post_started(sender, analysis):  # Do not change this line",
        "    define_dpf_workflow(analysis)",
        "def define_dpf_workflow(analysis):",
        "    import mech_dpf",
        "    import Ans.DataProcessing as dpf",
        "    mech_dpf.setExtAPI(ExtAPI)",
        "    node_vals = dict()",
        "\n".join(lines),
        "    node_ids = sorted(node_vals.keys())",
        "    scoping = dpf.Scoping()",
        "    scoping.Ids = node_ids",
        "    scoping.Location = 'Nodal'",
        "    field = dpf.Field()",
        "    field.Location = 'Nodal'",
        "    field.Scoping = scoping",
        "    field.Data = [node_vals[nid] for nid in node_ids]",
        "    fc = dpf.FieldsContainer()",
        "    fc.Labels = ['time']",
        "    ld = dict(); ld['time'] = 1",
        "    fc.Add(field, ld)",
        "    fwd = dpf.operators.utility.forward()",
        "    fwd.inputs.any.Connect(fc)",
        "    wf = dpf.Workflow()",
        "    wf.Add(fwd)",
        "    wf.SetOutputContour(fwd)",
        "    wf.Record('" + wf_key + "', True)",
        "    this.WorkflowId = wf.GetRecordedId()",
    ]
    return "\n".join(code)


# =============================================================================
# DIALOG
# =============================================================================
class StepPickerDialog(Form):
    """Soldan saga step secici (Force/Disp/Temp aracindaki ile ayni)."""
    def __init__(self, step_nos, preselected=None, step_labels=None):
        self._step_nos      = list(step_nos)
        self._step_labels   = step_labels or {}
        self.selected_steps = list(preselected) if preselected else []

        self.Text            = "Select Steps"
        self.StartPosition   = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox     = False; self.MinimizeBox = False
        self.ClientSize      = Size(560, 420)

        self.Controls.Add(_lbl("Available Steps", 15, 12, 210))
        self.lst_avail = ListBox()
        self.lst_avail.Location = Point(15, 35); self.lst_avail.Size = Size(210, 320)
        self.lst_avail.SelectionMode = SelectionMode.MultiExtended
        self.lst_avail.Font = CTRL_FONT
        self.lst_avail.DoubleClick += self._on_add
        self.Controls.Add(self.lst_avail)

        btn_add = Button(); btn_add.Text = "Add >"
        btn_add.Location = Point(237, 130); btn_add.Size = Size(86, 32)
        btn_add.ForeColor = Color.FromArgb(0, 100, 200); btn_add.Click += self._on_add
        self.Controls.Add(btn_add)

        btn_all = Button(); btn_all.Text = "All >>"
        btn_all.Location = Point(237, 168); btn_all.Size = Size(86, 32)
        btn_all.ForeColor = Color.FromArgb(0, 100, 200); btn_all.Click += self._on_add_all
        self.Controls.Add(btn_all)

        btn_del = Button(); btn_del.Text = "< Remove"
        btn_del.Location = Point(237, 214); btn_del.Size = Size(86, 32)
        btn_del.ForeColor = Color.FromArgb(200, 0, 0); btn_del.Click += self._on_del
        self.Controls.Add(btn_del)

        btn_clr = Button(); btn_clr.Text = "<< Clear"
        btn_clr.Location = Point(237, 252); btn_clr.Size = Size(86, 32)
        btn_clr.ForeColor = Color.FromArgb(200, 0, 0); btn_clr.Click += self._on_clear
        self.Controls.Add(btn_clr)

        self.Controls.Add(_lbl("Selected Steps", 335, 12, 210))
        self.lst_sel = ListBox()
        self.lst_sel.Location = Point(335, 35); self.lst_sel.Size = Size(210, 320)
        self.lst_sel.SelectionMode = SelectionMode.MultiExtended
        self.lst_sel.Font = CTRL_FONT
        self.lst_sel.DoubleClick += self._on_del
        self.Controls.Add(self.lst_sel)

        btn_ok = Button(); btn_ok.Text = "OK"
        btn_ok.Location = Point(365, 372); btn_ok.Size = Size(80, 30)
        btn_ok.Click += self._on_ok; self.Controls.Add(btn_ok)

        btn_cn = Button(); btn_cn.Text = "Cancel"
        btn_cn.Location = Point(455, 372); btn_cn.Size = Size(90, 30)
        btn_cn.DialogResult = DialogResult.Cancel
        self.Controls.Add(btn_cn); self.CancelButton = btn_cn

        self._populate()

    def _label(self, s):
        return self._step_labels.get(s, "Step {0}".format(s))

    def _populate(self):
        sel = set(self.selected_steps)
        self.lst_avail.Items.Clear(); self.lst_sel.Items.Clear()
        for s in self._step_nos:
            if s not in sel: self.lst_avail.Items.Add(self._label(s))
        for s in sorted(sel):
            self.lst_sel.Items.Add(self._label(s))

    def _parse(self, item):
        txt = safe_str(item)
        for s in self._step_nos:
            if self._label(s) == txt: return s
        return None

    def _on_add(self, s, e):
        for i in range(self.lst_avail.Items.Count):
            if self.lst_avail.GetSelected(i):
                st = self._parse(self.lst_avail.Items[i])
                if st is not None and st not in self.selected_steps:
                    self.selected_steps.append(st)
        self.selected_steps.sort(); self._populate()

    def _on_add_all(self, s, e):
        self.selected_steps = list(self._step_nos); self._populate()

    def _on_del(self, s, e):
        rem = set()
        for i in range(self.lst_sel.Items.Count):
            if self.lst_sel.GetSelected(i):
                st = self._parse(self.lst_sel.Items[i])
                if st is not None: rem.add(st)
        self.selected_steps = [x for x in self.selected_steps if x not in rem]
        self._populate()

    def _on_clear(self, s, e):
        self.selected_steps = []; self._populate()

    def _on_ok(self, s, e):
        if not self.selected_steps:
            MessageBox.Show("Please select at least one step.", "Warning"); return
        self.DialogResult = DialogResult.OK; self.Close()


class AvgDialog(Form):
    ROW_H = 28

    def __init__(self, analyses, ns_list):
        self._analyses = analyses
        self._ns_list  = ns_list
        self._ns_names = [safe_str(n.Name) for n in ns_list]
        self._steps    = []          # [(ls, set_id, t)]
        self._step_labels = {}
        self._sel_step_nos = []
        self._ns_rows  = []          # [(lbl, combo, btn_add, btn_rem)]
        self.ds = None; self.dpf_model = None; self.rst_path = ""

        self.Text = "Averaged / Unaveraged % Check (RST)"
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False; self.MinimizeBox = False
        self.BackColor = Color.White
        self.AutoScroll = True
        self._W = 480; self._X = 12

        # Analysis
        self.lbl_an = _lbl("Analysis:", 0, 0, self._W); self.Controls.Add(self.lbl_an)
        self.cmb_an = _no_scroll(ComboBox())
        self.cmb_an.Size = Size(self._W, 24)
        self.cmb_an.DropDownStyle = ComboBoxStyle.DropDownList; self.cmb_an.Font = CTRL_FONT
        idx = 0
        for i, a in enumerate(analyses):
            at = safe_str(a.AnalysisType)
            self.cmb_an.Items.Add("[{0}] {1}  ({2})".format(i + 1, safe_str(a.Name), at))
            if idx == 0 and "static" in at.lower(): idx = i
        self.Controls.Add(self.cmb_an)

        # RST
        self.lbl_rst = _lbl("RST File:", 0, 0, self._W); self.Controls.Add(self.lbl_rst)
        self.txt_rst = TextBox(); self.txt_rst.ReadOnly = True
        self.txt_rst.Size = Size(self._W - 88, 24)
        self.txt_rst.Font = CTRL_FONT; self.txt_rst.BackColor = Color.White
        self.Controls.Add(self.txt_rst)
        self.btn_rst = _btn("Browse...", 82, 24)
        self.btn_rst.Click += self._on_browse; self.Controls.Add(self.btn_rst)

        # Named Selections (dropdown satirlari)
        self.lbl_ns = _lbl("Named Selections:", 0, 0, self._W); self.Controls.Add(self.lbl_ns)
        self._add_ns_row()

        # Stress types (coklu secim)
        self.grp_st = GroupBox(); self.grp_st.Text = "Stress Type(s)"; self.grp_st.Font = HDR_FONT
        self.grp_st.Size = Size(self._W, 50)
        self._chks = []
        for i, (_k, txt) in enumerate(STRESS_TYPES):
            cb = CheckBox(); cb.Text = txt; cb.Font = CTRL_FONT
            cb.Location = Point(10 + i * 158, 20); cb.Size = Size(156, 22)
            cb.Checked = (i == 0)
            self.grp_st.Controls.Add(cb); self._chks.append(cb)
        self.Controls.Add(self.grp_st)

        # Steps (soldan saga picker)
        self.lbl_steps = _lbl("Steps:", 0, 0, 60); self.Controls.Add(self.lbl_steps)
        self.btn_steps = _btn("Select Steps...  (none selected)", 300, 26)
        self.btn_steps.Click += self._on_pick_steps
        self.Controls.Add(self.btn_steps)

        # Run / Cancel
        self.btn_ok = _btn("Run", 90, 30, True); self.btn_ok.Click += self._ok
        self.Controls.Add(self.btn_ok)
        self.btn_cn = _btn("Cancel", 90, 30); self.btn_cn.DialogResult = DialogResult.Cancel
        self.Controls.Add(self.btn_cn); self.CancelButton = self.btn_cn

        self.cmb_an.SelectedIndexChanged += self._on_analysis
        self.cmb_an.SelectedIndex = idx
        self._on_analysis()
        self._relayout()

    # ---------------- layout ----------------
    def _relayout(self):
        X = self._X; W = self._W; y = 12
        self.lbl_an.Location = Point(X, y); y += 22
        self.cmb_an.Location = Point(X, y); y += 32
        self.lbl_rst.Location = Point(X, y); y += 22
        self.txt_rst.Location = Point(X, y)
        self.btn_rst.Location = Point(X + W - 82, y - 1); y += 34

        self.lbl_ns.Location = Point(X, y); y += 22
        cmb_w = W - 2 * (26 + 4) - 4 - 34
        for i, (lbl, cmb, ba, br) in enumerate(self._ns_rows):
            lbl.Location = Point(X, y + 4); lbl.Text = "{0}.".format(i + 1)
            cmb.Location = Point(X + 34, y); cmb.Size = Size(cmb_w, 24)
            ba.Location  = Point(X + 34 + cmb_w + 4, y)
            br.Location  = Point(X + 34 + cmb_w + 4 + 30, y)
            br.Enabled   = len(self._ns_rows) > 1
            y += self.ROW_H
        y += 6

        self.grp_st.Location = Point(X, y); y += 60
        self.lbl_steps.Location = Point(X, y + 4)
        self.btn_steps.Location = Point(X + 64, y); y += 42

        self.btn_ok.Location = Point(X + W - 186, y)
        self.btn_cn.Location = Point(X + W - 90, y)
        self.ClientSize = Size(W + 2 * X, min(y + 44, 800))

    # ---------------- NS rows ----------------
    def _add_ns_row(self, s=None, e=None):
        lbl = _lbl("", 0, 0, 30, bold=False)
        cmb = _no_scroll(ComboBox())
        cmb.DropDownStyle = ComboBoxStyle.DropDownList; cmb.Font = CTRL_FONT
        cmb.Items.Add("-- Select --")
        for n in self._ns_names: cmb.Items.Add(n)
        cmb.SelectedIndex = 0
        ba = Button(); ba.Text = "+"; ba.Size = Size(26, 24)
        ba.BackColor = BTN_BLUE; ba.ForeColor = Color.White; ba.FlatStyle = FlatStyle.Flat
        ba.Click += self._add_ns_row
        br = Button(); br.Text = "-"; br.Size = Size(26, 24)
        br.BackColor = Color.FromArgb(200, 0, 0); br.ForeColor = Color.White
        br.FlatStyle = FlatStyle.Flat
        br.Click += self._rem_ns_row
        for c in (lbl, cmb, ba, br): self.Controls.Add(c)
        self._ns_rows.append((lbl, cmb, ba, br))
        if s is not None: self._relayout()

    def _rem_ns_row(self, s, e):
        if len(self._ns_rows) <= 1: return
        target = None
        for row in self._ns_rows:
            if row[3] is s: target = row; break
        if target is None: return
        for c in target: self.Controls.Remove(c)
        self._ns_rows.remove(target)
        self._relayout()

    # ---------------- steps ----------------
    def _update_step_btn(self):
        n = len(self._sel_step_nos); tot = len(self._steps)
        if n == 0:     self.btn_steps.Text = "Select Steps...  (none selected)"
        elif n == tot: self.btn_steps.Text = "Select Steps...  (all {0} selected)".format(tot)
        else:          self.btn_steps.Text = "Select Steps...  ({0} of {1} selected)".format(n, tot)

    def _on_pick_steps(self, s, e):
        if not self._steps:
            MessageBox.Show("No result steps available (check RST file).", "Warning"); return
        dlg = StepPickerDialog([ls for ls, _sid, _t in self._steps],
                               preselected=self._sel_step_nos,
                               step_labels=self._step_labels)
        if dlg.ShowDialog() == DialogResult.OK:
            self._sel_step_nos = sorted(dlg.selected_steps)
            self._update_step_btn()

    # ---------------- RST ----------------
    def _on_analysis(self, s=None, e=None):
        a = self._analyses[self.cmb_an.SelectedIndex]
        try:    rst = safe_str(a.ResultFileName)
        except: rst = ""
        self._load_rst(rst if rst and os.path.exists(rst) else "")

    def _on_browse(self, s, e):
        dlg = OpenFileDialog()
        dlg.Filter = "Ansys Result File (*.rst)|*.rst|All files (*.*)|*.*"
        if self.rst_path and os.path.exists(self.rst_path):
            dlg.InitialDirectory = os.path.dirname(self.rst_path)
        if dlg.ShowDialog() == DialogResult.OK:
            self._load_rst(dlg.FileName)

    def _load_rst(self, path):
        self._steps = []; self._step_labels = {}; self._sel_step_nos = []
        self.ds = None; self.dpf_model = None; self.rst_path = path
        self.txt_rst.Text = path if path else "(no result file - solve or browse)"
        if path:
            try:
                self.ds, self.dpf_model = dpf_load(path)
                self._steps = get_load_step_info(self.dpf_model)
            except Exception as ex:
                MessageBox.Show("RST could not be read:\n" + safe_str(ex), "Error")
                self.ds = None; self.dpf_model = None; self._steps = []
        self._step_labels = dict((ls, "Step {0}  (t={1:.4g}s)".format(ls, t))
                                 for ls, _sid, t in self._steps)
        if self._steps:
            self._sel_step_nos = [self._steps[-1][0]]   # varsayilan: son step
        self._update_step_btn()

    # ---------------- OK ----------------
    def _ok(self, s, e):
        errs = []
        if self.ds is None: errs.append("- No valid RST file.")
        sel_ns = []; seen = set()
        for _l, cmb, _a, _r in self._ns_rows:
            if cmb.SelectedIndex > 0:
                i = cmb.SelectedIndex - 1
                if i not in seen:
                    seen.add(i); sel_ns.append(self._ns_list[i])
        keys = [STRESS_TYPES[i][0] for i, cb in enumerate(self._chks) if cb.Checked]
        step_map = dict((ls, (ls, sid, t)) for ls, sid, t in self._steps)
        sel_st = [step_map[n] for n in self._sel_step_nos if n in step_map]
        if not sel_ns: errs.append("- Select at least one Named Selection.")
        if not keys:   errs.append("- Select at least one stress type.")
        if not sel_st: errs.append("- Select at least one step.")
        if errs:
            MessageBox.Show("\n".join(errs), "Missing Input",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning); return
        self.analysis  = self._analyses[self.cmb_an.SelectedIndex]
        self.sel_ns    = sel_ns
        self.sel_steps = sel_st
        self.keys      = keys
        self.DialogResult = DialogResult.OK; self.Close()


class ProgressDialog(Form):
    def __init__(self):
        self.Text = "Processing..."; self.TopMost = True
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedSingle
        self.MaximizeBox = False; self.ClientSize = Size(440, 72)
        self.l1 = _lbl("Starting...", 15, 14, 410)
        self.l2 = _lbl("", 15, 38, 410, bold=False)
        self.l2.ForeColor = Color.DimGray
        self.Controls.Add(self.l1); self.Controls.Add(self.l2)
    def update(self, a, b=""):
        self.l1.Text = a; self.l2.Text = b; Application.DoEvents()


# =============================================================================
# ENTRY POINT
# =============================================================================
def run():
    model    = ExtAPI.DataModel.Project.Model
    analyses = list(model.Analyses)
    ns_list  = list(model.NamedSelections.Children) if model.NamedSelections else []
    if not analyses or not ns_list:
        MessageBox.Show("No analysis or Named Selection found.", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error); return

    dlg = AvgDialog(analyses, ns_list)
    if dlg.ShowDialog() != DialogResult.OK: return

    t0 = time.time()
    keys = dlg.keys
    key_txt = dict(STRESS_TYPES)
    sol = dlg.analysis.Solution
    prog = ProgressDialog(); prog.Show()
    errors = []; results = []

    prog.update("Reading mesh from RST...")
    mesh = get_mesh(dlg.ds, dlg.dpf_model)
    if mesh is None:
        prog.Close()
        MessageBox.Show("Mesh could not be read from RST.", "Error"); return

    ns_elems = []
    for ns in dlg.sel_ns:
        ids = ns_elem_ids(ns)
        if ids: ns_elems.append((ns, ids))
        else:   errors.append("{0}: no elements found (NS must be body/element scoped)".format(safe_str(ns.Name)))
    if not ns_elems:
        prog.Close()
        MessageBox.Show("\n".join(errors), "Error"); return

    union = set()
    for _ns, ids in ns_elems: union.update(ids)
    union = sorted(union)
    conn_cache = {}

    for ls, sid, t in dlg.sel_steps:
        prog.update("Step {0}: reading stress from RST...".format(ls),
                    "{0} elements".format(len(union)))
        try:
            edata = read_stress_elemnodal(dlg.ds, sid, union)
        except Exception as ex:
            errors.append("Step {0}: DPF read failed - {1}".format(ls, safe_str(ex))); continue
        if not edata:
            errors.append("Step {0}: empty stress field".format(ls)); continue

        for ns, eids in ns_elems:
            ns_name = safe_str(ns.Name)
            prog.update("Step {0}: {1}".format(ls, ns_name), "Computing averaged / unaveraged...")
            multi = compute_pct_multi(eids, edata, mesh, keys, conn_cache)
            for key in keys:
                pct, avg, una = multi[key]
                if not pct:
                    errors.append("{0} | Step {1} | {2}: no matching nodes".format(ns_name, ls, key))
                    continue

                rname = "AvgErr%_{0}_{1}_Step{2}".format(key, ascii_tag(ns_name), ls)
                for ch in list(sol.Children):
                    if ch.Name == rname:
                        try: ch.Delete()
                        except: pass
                pr = sol.AddPythonResult()
                pr.Name = rname
                pr.Text = build_dpf_code(pct, rname)
                try: pr.Connect()
                except: pass

                nmax = max(pct, key=pct.get)
                line = ("{0} | Step {1} | {2}: Nodes={3}  Max%={4:.2f} @ Node {5}  "
                        "(Avg={6:.3f}, Unavg={7:.3f})").format(
                            ns_name, ls, key, len(pct), pct[nmax], nmax, avg[nmax], una[nmax])
                results.append((pr, key_txt[key], line))

    if results:
        prog.update("Evaluating Python Results...", "{0} result(s)".format(len(results)))
        try: sol.EvaluateAllResults()
        except: pass
        for pr, txt, line in results:
            try:
                c = pr.AddComment(); c.Text = txt + "  |  " + line
            except: pass
        try: results[-1][0].Activate()
        except: pass

    prog.Close()
    msg = ["Averaged / Unaveraged % Check  —  " + ", ".join(key_txt[k] for k in keys),
           "% = abs((Unaveraged - Averaged) / Unaveraged) * 100\n"]
    if results: msg += ["Completed:"] + [l for _p, _t, l in results]
    if errors:  msg += ["\nERRORS:"] + errors
    msg.append("\nDuration: " + fmt_duration(time.time() - t0))
    MessageBox.Show("\n".join(msg), "Averaged / Unaveraged %",
                    MessageBoxButtons.OK, MessageBoxIcon.Information)


run()
