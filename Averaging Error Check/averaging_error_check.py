# -*- coding: utf-8 -*-
# ============================================================
# Nodal Averaging Error Check - ANSYS Mechanical (IronPython)
# Averaged vs Unaveraged gerilme farkini node bazinda % olarak
# hesaplar ve DPF Python Result ile plot eder.
#
#   pct(node) = max_i |s_unavg,i - s_avg| / max(|s_avg|, floor) * 100
#   floor     = DENOM_FLOOR * max|s_avg|  (NS icinde)
# ============================================================
import os
import clr
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
from System.Windows.Forms import (
    Form, ComboBox, Button, Label, RadioButton, TextBox, GroupBox,
    DialogResult, FormStartPosition, ComboBoxStyle, FormBorderStyle,
    MessageBox, MessageBoxButtons, MessageBoxIcon
)
from System.Drawing import Size, Point, Color, Font, FontStyle

import mech_dpf
import Ans.DataProcessing as dpf
mech_dpf.setExtAPI(ExtAPI)

CTRL_FONT = Font("Segoe UI", 9)
HDR_FONT  = Font("Segoe UI", 9, FontStyle.Bold)
BTN_BLUE  = Color.FromArgb(0, 120, 215)

DENOM_FLOOR   = 0.01    # payda tabani: NS icindeki max|avg| degerinin %1'i
LIMIT_DEFAULT = 10.0    # %

# (key, UI text, DPF result operator)
STRESS_TYPES = [
    ("EQV", "Equivalent (von Mises)", "stress_von_mises"),
    ("S1",  "Maximum Principal",      "stress_principal_1"),
    ("S3",  "Minimum Principal",      "stress_principal_3"),
]


# =====================================================================
# HELPERS
# =====================================================================
def _btn(text, w, h, blue=False):
    b = Button(); b.Text = text; b.Size = Size(w, h); b.Font = CTRL_FONT
    if blue:
        b.Font = HDR_FONT; b.BackColor = BTN_BLUE
        b.ForeColor = Color.White; b.FlatStyle = 0
    return b


def _cmb(x, y, w):
    c = ComboBox(); c.Location = Point(x, y); c.Size = Size(w, 24)
    c.DropDownStyle = ComboBoxStyle.DropDownList; c.Font = CTRL_FONT
    def _block(s, e):
        try: e.Handled = True
        except: pass
    c.MouseWheel += _block
    return c


def _lbl(text, x, y, w, bold=False):
    l = Label(); l.Text = text; l.Location = Point(x, y)
    l.Size = Size(w, 20); l.Font = HDR_FONT if bold else CTRL_FONT
    return l


def get_sets(analysis):
    """[(label, set_id)] — son set varsayilan."""
    out = []
    try:
        rst = analysis.ResultFileName
        if rst and os.path.exists(rst):
            tf = dpf.Model(dpf.DataSources(rst)).Metadata.TimeFreqSupport
            times = list(tf.TimeFreqs.Data)
            for i in range(tf.NumberSets):
                out.append(("Set {0}  (T={1:.6g})".format(i + 1, times[i]), i + 1))
    except: pass
    if not out:
        out = [("Last Set", -1)]
    return out


def get_elem_ids(analysis, ns):
    """Body NS -> mesh eleman ID'leri."""
    md = analysis.MeshData
    ids = set()
    for gid in ns.Ids:
        try:
            reg = md.MeshRegionById(gid)
            for e in reg.ElementIds: ids.add(int(e))
        except: pass
    return sorted(ids)


# =====================================================================
# POPUP
# =====================================================================
class AvgErrDialog(Form):
    def __init__(self, analyses, ns_list):
        self._analyses = analyses
        self._ns_list  = ns_list
        self._sets     = []
        self.Text = "Nodal Averaging Error Check"
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False; self.MinimizeBox = False
        self.BackColor = Color.White
        W = 420; y = 12

        self.Controls.Add(_lbl("Analysis:", 12, y, W, True)); y += 22
        self.cmb_an = _cmb(12, y, W)
        default_idx = 0
        for i, a in enumerate(analyses):
            atype = str(a.AnalysisType)
            self.cmb_an.Items.Add("[{0}] {1}  ({2})".format(i + 1, a.Name, atype))
            if default_idx == 0 and "static" in atype.lower(): default_idx = i
        self.Controls.Add(self.cmb_an); y += 34

        self.Controls.Add(_lbl("Named Selection (Body):", 12, y, W, True)); y += 22
        self.cmb_ns = _cmb(12, y, W)
        for ns in ns_list: self.cmb_ns.Items.Add(ns.Name)
        if self.cmb_ns.Items.Count > 0: self.cmb_ns.SelectedIndex = 0
        self.Controls.Add(self.cmb_ns); y += 36

        grp = GroupBox(); grp.Text = "Stress Type"; grp.Font = HDR_FONT
        grp.Location = Point(12, y); grp.Size = Size(W, 104)
        self._rbs = []
        for i, (_k, txt, _op) in enumerate(STRESS_TYPES):
            rb = RadioButton(); rb.Text = txt; rb.Font = CTRL_FONT
            rb.Location = Point(14, 22 + i * 26); rb.Size = Size(W - 30, 22)
            rb.Checked = (i == 0)
            grp.Controls.Add(rb); self._rbs.append(rb)
        self.Controls.Add(grp); y += 114

        self.Controls.Add(_lbl("Result Set:", 12, y, W, True)); y += 22
        self.cmb_set = _cmb(12, y, W)
        self.Controls.Add(self.cmb_set); y += 34

        self.Controls.Add(_lbl("Limit (%):", 12, y + 3, 80, True))
        self.txt_lim = TextBox(); self.txt_lim.Text = str(LIMIT_DEFAULT)
        self.txt_lim.Location = Point(95, y); self.txt_lim.Size = Size(70, 24)
        self.txt_lim.Font = CTRL_FONT
        self.Controls.Add(self.txt_lim); y += 40

        ok = _btn("Run", 90, 30, True); ok.Location = Point(W - 178, y)
        ok.Click += self._ok; self.Controls.Add(ok)
        cn = _btn("Cancel", 90, 30); cn.Location = Point(W - 78, y)
        cn.DialogResult = DialogResult.Cancel
        self.Controls.Add(cn); self.CancelButton = cn
        self.ClientSize = Size(W + 24, y + 44)

        self.cmb_an.SelectedIndexChanged += self._reload_sets
        self.cmb_an.SelectedIndex = default_idx
        self._reload_sets()

    def _reload_sets(self, s=None, e=None):
        a = self._analyses[self.cmb_an.SelectedIndex]
        self._sets = get_sets(a)
        self.cmb_set.Items.Clear()
        for lbl, _sid in self._sets: self.cmb_set.Items.Add(lbl)
        self.cmb_set.SelectedIndex = len(self._sets) - 1

    def _ok(self, s, e):
        try:
            lim = float(self.txt_lim.Text.replace(",", "."))
            if lim <= 0: raise ValueError()
        except:
            MessageBox.Show("Limit pozitif bir sayi olmali.", "Hata",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning); return
        if self.cmb_ns.SelectedIndex < 0:
            MessageBox.Show("Named Selection secin.", "Hata",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning); return
        self.analysis   = self._analyses[self.cmb_an.SelectedIndex]
        self.ns         = self._ns_list[self.cmb_ns.SelectedIndex]
        self.stype      = [STRESS_TYPES[i] for i, rb in enumerate(self._rbs) if rb.Checked][0]
        self.set_label, self.set_id = self._sets[self.cmb_set.SelectedIndex]
        self.limit      = lim
        self.DialogResult = DialogResult.OK; self.Close()


# =====================================================================
# DPF COMPUTE  (ElementalNodal -> averaged + % fark)
# =====================================================================
def compute_avg_error(rst, elem_ids, op_name, set_id):
    ds = dpf.DataSources(rst)

    scop = dpf.Scoping(); scop.Location = "Elemental"; scop.Ids = elem_ids

    op = getattr(dpf.operators.result, op_name)()
    op.inputs.data_sources.Connect(ds)
    op.inputs.mesh_scoping.Connect(scop)
    op.inputs.requested_location.Connect("ElementalNodal")
    if set_id > 0:
        ts = dpf.Scoping(); ts.Location = "TimeFreq_sets"; ts.Ids = [set_id]
        op.inputs.time_scoping.Connect(ts)

    field = op.outputs.fields_container.GetData()[0]
    mesh  = dpf.Model(ds).Metadata.MeshedRegion

    # node -> [unaveraged degerler (her eleman icin bir tane)]
    node_vals = {}
    for eid in field.ScopingIds:
        vals = list(field.GetEntityDataById(eid))
        nids = list(mesh.ElementById(eid).NodeIds)
        # Quadratic elemanlarda rst sadece corner node degerlerini tutar;
        # node listesinde corner'lar once geldigi icin ilk len(vals) yeterli.
        for nid, v in zip(nids[:len(vals)], vals):
            node_vals.setdefault(nid, []).append(v)

    if not node_vals:
        raise Exception("DPF bos sonuc dondurdu (eleman scoping / result kontrol edin).")

    avg = dict((n, sum(l) / len(l)) for n, l in node_vals.items())
    amax  = max(abs(v) for v in avg.values()) or 1.0
    floor = DENOM_FLOOR * amax

    pct = {}
    for nid, lst in node_vals.items():
        a = avg[nid]
        dev = max(abs(u - a) for u in lst)
        pct[nid] = dev / max(abs(a), floor) * 100.0
    return pct, avg, node_vals


# =====================================================================
# PYTHON RESULT (forward ile plot) + LEGEND
# =====================================================================
def build_dpf_code(node_vals, result_name):
    pairs = list(node_vals.items()); chunk = 150; lines = []
    for i in range(0, len(pairs), chunk):
        seg = ", ".join("({0}, {1!r})".format(k, v) for k, v in pairs[i:i + chunk])
        lines.append("    for _k, _v in [{0}]: node_vals[_k] = _v".format(seg))
    wf_key = "avgerr_" + result_name[:40].replace(" ", "_").replace("-", "_")
    code = [
        "# Averaging Error % - " + result_name,
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


def apply_legend(py_result, pmax, limit):
    """Yesil: 0..limit/2 | Sari: limit/2..limit | Kirmizi: > limit  (renkler BGR)"""
    try:
        py_result.Activate(); ExtAPI.Graphics.Redraw()
        lg = Ansys.Mechanical.Graphics.Tools.CurrentLegendSettings()
        half = limit / 2.0
        top  = max(pmax, limit * 1.001)
        bands = [(0.0, half, 0x00FF00), (half, limit, 0x00FFFF), (limit, top, 0x0000FF)]
        lg.NumberOfBands = 3
        for i, (lo, hi, col) in enumerate(bands):
            lg.SetLowerBound(i, Quantity(float(lo), ""))
            lg.SetUpperBound(i, Quantity(float(hi), ""))
            lg.SetBandColorAuto(i, False); lg.SetBandColor(i, col)
        ExtAPI.Graphics.Redraw()
        return None
    except Exception as ex:
        return str(ex)


# =====================================================================
# ENTRY POINT
# =====================================================================
def run():
    model    = ExtAPI.DataModel.Project.Model
    analyses = list(model.Analyses)
    ns_list  = list(model.NamedSelections.Children) if model.NamedSelections else []
    if not analyses or not ns_list:
        MessageBox.Show("Analiz veya Named Selection bulunamadi.", "Hata",
                        MessageBoxButtons.OK, MessageBoxIcon.Error); return

    dlg = AvgErrDialog(analyses, ns_list)
    if dlg.ShowDialog() != DialogResult.OK: return

    analysis = dlg.analysis; ns = dlg.ns
    key, stxt, op_name = dlg.stype
    rst = analysis.ResultFileName
    if not rst or not os.path.exists(rst):
        MessageBox.Show("Result dosyasi yok — analizi cozun.", "Hata",
                        MessageBoxButtons.OK, MessageBoxIcon.Error); return

    elem_ids = get_elem_ids(analysis, ns)
    if not elem_ids:
        MessageBox.Show("NS icin eleman bulunamadi (NS body scoped olmali).", "Hata",
                        MessageBoxButtons.OK, MessageBoxIcon.Error); return

    try:
        pct, avg, nv = compute_avg_error(rst, elem_ids, op_name, dlg.set_id)
    except Exception as ex:
        MessageBox.Show("DPF hatasi:\n" + str(ex), "Hata",
                        MessageBoxButtons.OK, MessageBoxIcon.Error); return

    sol = analysis.Solution
    set_tag = "Set{0}".format(dlg.set_id) if dlg.set_id > 0 else "Last"
    rname = "AvgErr%_{0}_{1}_{2}".format(key, ns.Name[:25], set_tag)
    for ch in list(sol.Children):
        if ch.Name == rname:
            try: ch.Delete()
            except: pass

    pr = sol.AddPythonResult()
    pr.Name = rname
    pr.Text = build_dpf_code(pct, rname)
    try: pr.Connect()
    except: pass
    try: pr.EvaluateAllResults()
    except:
        try: sol.EvaluateAllResults()
        except: pass

    pmax_nid = max(pct, key=pct.get); pmax = pct[pmax_nid]
    n_exc = sum(1 for v in pct.values() if v > dlg.limit)
    warn = apply_legend(pr, pmax, dlg.limit)

    summary = [
        "Averaging Error Check  —  {0}".format(stxt),
        "NS: {0}   |   {1}".format(ns.Name, dlg.set_label),
        "Elements: {0}   Nodes: {1}".format(len(elem_ids), len(pct)),
        "Max %: {0:.2f}  @ Node {1}".format(pmax, pmax_nid),
        "  Averaged   : {0:.4f}".format(avg[pmax_nid]),
        "  Unaveraged : {0}".format(", ".join("{0:.4f}".format(v) for v in nv[pmax_nid])),
        "Nodes > {0:.1f}% : {1}".format(dlg.limit, n_exc),
    ]
    try:
        c = pr.AddComment(); c.Text = "  |  ".join(summary)
    except: pass
    if warn: summary.append("\nLegend uyarisi: " + warn)
    summary.append("\nYesil < {0:.1f}%  |  Sari < {1:.1f}%  |  Kirmizi > {1:.1f}%".format(
        dlg.limit / 2.0, dlg.limit))
    MessageBox.Show("\n".join(summary), "Averaging Error Result",
                    MessageBoxButtons.OK, MessageBoxIcon.Information)


run()
