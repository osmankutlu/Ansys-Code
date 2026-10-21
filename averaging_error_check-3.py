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
    Form, Label, Button, RadioButton, GroupBox, CheckedListBox, ComboBox,
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

SCALAR_FN = {
    "EQV": lambda v: s_eqv(v[0], v[1], v[2], v[3], v[4], v[5]),
    "S1":  lambda v: s_principal(v[0], v[1], v[2], v[3], v[4], v[5])[0],
    "S3":  lambda v: s_principal(v[0], v[1], v[2], v[3], v[4], v[5])[1],
}

def compute_pct(elem_ids, edata, mesh, key, conn_cache):
    fn = SCALAR_FN[key]
    node_vals = {}
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
            node_vals.setdefault(nids[j], []).append(fn(vals[b:b + 6]))

    pct, avg, una = {}, {}, {}
    for nid, lst in node_vals.items():
        a = sum(lst) / len(lst)
        u = max(lst, key=abs)
        avg[nid] = a; una[nid] = u
        pct[nid] = abs((u - a) / u) * 100.0 if abs(u) > 1e-12 else 0.0
    return pct, avg, una


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
class AvgDialog(Form):
    def __init__(self, analyses, ns_list):
        self._analyses = analyses
        self._ns_list  = ns_list
        self._steps    = []          # [(ls, set_id, t)]
        self.ds = None; self.dpf_model = None; self.rst_path = ""

        self.Text = "Averaged / Unaveraged % Check (RST)"
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False; self.MinimizeBox = False
        self.BackColor = Color.White
        W = 460; X = 12; y = 12

        # Analysis
        self.Controls.Add(_lbl("Analysis:", X, y, W)); y += 22
        self.cmb_an = _no_scroll(ComboBox())
        self.cmb_an.Location = Point(X, y); self.cmb_an.Size = Size(W, 24)
        self.cmb_an.DropDownStyle = ComboBoxStyle.DropDownList; self.cmb_an.Font = CTRL_FONT
        idx = 0
        for i, a in enumerate(analyses):
            at = safe_str(a.AnalysisType)
            self.cmb_an.Items.Add("[{0}] {1}  ({2})".format(i + 1, safe_str(a.Name), at))
            if idx == 0 and "static" in at.lower(): idx = i
        self.Controls.Add(self.cmb_an); y += 32

        # RST
        self.Controls.Add(_lbl("RST File:", X, y, W)); y += 22
        self.txt_rst = TextBox(); self.txt_rst.ReadOnly = True
        self.txt_rst.Location = Point(X, y); self.txt_rst.Size = Size(W - 88, 24)
        self.txt_rst.Font = CTRL_FONT; self.txt_rst.BackColor = Color.White
        self.Controls.Add(self.txt_rst)
        b = _btn("Browse...", 82, 24); b.Location = Point(X + W - 82, y - 1)
        b.Click += self._on_browse; self.Controls.Add(b)
        y += 34

        # Named Selections
        self.Controls.Add(_lbl("Named Selections:", X, y, 200))
        self._add_all_clear(X + W - 112, y - 2, lambda: self.lst_ns)
        y += 24
        self.lst_ns = CheckedListBox(); self.lst_ns.CheckOnClick = True
        self.lst_ns.Location = Point(X, y); self.lst_ns.Size = Size(W, 120)
        self.lst_ns.Font = CTRL_FONT
        for ns in ns_list: self.lst_ns.Items.Add(safe_str(ns.Name))
        self.Controls.Add(self.lst_ns); y += 128

        # Stress type
        grp = GroupBox(); grp.Text = "Stress Type"; grp.Font = HDR_FONT
        grp.Location = Point(X, y); grp.Size = Size(W, 50)
        self._rbs = []
        for i, (_k, txt) in enumerate(STRESS_TYPES):
            rb = RadioButton(); rb.Text = txt; rb.Font = CTRL_FONT
            rb.Location = Point(10 + i * 150, 20); rb.Size = Size(148, 22)
            rb.Checked = (i == 0)
            grp.Controls.Add(rb); self._rbs.append(rb)
        self.Controls.Add(grp); y += 60

        # Steps
        self.Controls.Add(_lbl("Steps:", X, y, 200))
        self._add_all_clear(X + W - 112, y - 2, lambda: self.lst_steps)
        y += 24
        self.lst_steps = CheckedListBox(); self.lst_steps.CheckOnClick = True
        self.lst_steps.Location = Point(X, y); self.lst_steps.Size = Size(W, 120)
        self.lst_steps.Font = CTRL_FONT; self.lst_steps.MultiColumn = False
        self.Controls.Add(self.lst_steps); y += 132

        ok = _btn("Run", 90, 30, True); ok.Location = Point(X + W - 186, y)
        ok.Click += self._ok; self.Controls.Add(ok)
        cn = _btn("Cancel", 90, 30); cn.Location = Point(X + W - 90, y)
        cn.DialogResult = DialogResult.Cancel
        self.Controls.Add(cn); self.CancelButton = cn
        self.ClientSize = Size(W + 2 * X, y + 44)

        self.cmb_an.SelectedIndexChanged += self._on_analysis
        self.cmb_an.SelectedIndex = idx
        self._on_analysis()

    def _add_all_clear(self, x, y, getter):
        ba = _btn("All", 52, 22); ba.Location = Point(x, y)
        bc = _btn("Clear", 56, 22); bc.Location = Point(x + 56, y)
        def _set(state):
            lst = getter()
            for i in range(lst.Items.Count): lst.SetItemChecked(i, state)
        ba.Click += lambda s, e: _set(True)
        bc.Click += lambda s, e: _set(False)
        self.Controls.Add(ba); self.Controls.Add(bc)

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
        self.lst_steps.Items.Clear(); self._steps = []
        self.ds = None; self.dpf_model = None; self.rst_path = path
        self.txt_rst.Text = path if path else "(no result file - solve or browse)"
        if not path: return
        try:
            self.ds, self.dpf_model = dpf_load(path)
            self._steps = get_load_step_info(self.dpf_model)
        except Exception as ex:
            MessageBox.Show("RST could not be read:\n" + safe_str(ex), "Error")
            self.ds = None; self.dpf_model = None; return
        for ls, sid, t in self._steps:
            self.lst_steps.Items.Add("Step {0}   (t = {1:.6g} s,  set {2})".format(ls, t, sid))
        if self._steps:
            self.lst_steps.SetItemChecked(len(self._steps) - 1, True)

    def _ok(self, s, e):
        errs = []
        if self.ds is None: errs.append("- No valid RST file.")
        sel_ns = [self._ns_list[i] for i in range(self.lst_ns.Items.Count)
                  if self.lst_ns.GetItemChecked(i)]
        sel_st = [self._steps[i] for i in range(self.lst_steps.Items.Count)
                  if self.lst_steps.GetItemChecked(i)]
        if not sel_ns: errs.append("- Select at least one Named Selection.")
        if not sel_st: errs.append("- Select at least one step.")
        if errs:
            MessageBox.Show("\n".join(errs), "Missing Input",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning); return
        self.analysis  = self._analyses[self.cmb_an.SelectedIndex]
        self.sel_ns    = sel_ns
        self.sel_steps = sel_st
        self.stype     = [STRESS_TYPES[i] for i, rb in enumerate(self._rbs) if rb.Checked][0]
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
    key, stxt = dlg.stype
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
            pct, avg, una = compute_pct(eids, edata, mesh, key, conn_cache)
            if not pct:
                errors.append("{0} | Step {1}: no matching nodes".format(ns_name, ls)); continue

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
            line = ("{0} | Step {1}: Nodes={2}  Max%={3:.2f} @ Node {4}  "
                    "(Avg={5:.3f}, Unavg={6:.3f})").format(
                        ns_name, ls, len(pct), pct[nmax], nmax, avg[nmax], una[nmax])
            results.append((pr, line))

    if results:
        prog.update("Evaluating Python Results...", "{0} result(s)".format(len(results)))
        try: sol.EvaluateAllResults()
        except: pass
        for pr, line in results:
            try:
                c = pr.AddComment(); c.Text = stxt + "  |  " + line
            except: pass
        try: results[-1][0].Activate()
        except: pass

    prog.Close()
    msg = ["Averaged / Unaveraged % Check  —  " + stxt,
           "% = abs((Unaveraged - Averaged) / Unaveraged) * 100\n"]
    if results: msg += ["Completed:"] + [l for _p, l in results]
    if errors:  msg += ["\nERRORS:"] + errors
    msg.append("\nDuration: " + fmt_duration(time.time() - t0))
    MessageBox.Show("\n".join(msg), "Averaged / Unaveraged %",
                    MessageBoxButtons.OK, MessageBoxIcon.Information)


run()
