# -*- coding: utf-8 -*-
# ============================================================
# Averaged / Unaveraged % Check - ANSYS Mechanical (IronPython)
# Ayni sonucu Display Option = Averaged ve Unaveraged olarak
# hesaplar, node bazinda  % = abs((Unavg - Avg) / Unavg) * 100
# degerini DPF Python Result ile plot eder.
# ============================================================
import os
import clr
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
from System.Windows.Forms import (
    Form, ComboBox, Button, Label, RadioButton, GroupBox,
    DialogResult, FormStartPosition, ComboBoxStyle, FormBorderStyle,
    MessageBox, MessageBoxButtons, MessageBoxIcon
)
from System.Drawing import Size, Point, Color, Font, FontStyle

CTRL_FONT = Font("Segoe UI", 9)
HDR_FONT  = Font("Segoe UI", 9, FontStyle.Bold)
BTN_BLUE  = Color.FromArgb(0, 120, 215)
AVG_ENUM  = Ansys.Mechanical.DataModel.Enums.ResultAveragingType

# (key, UI text, Solution add-method)
STRESS_TYPES = [
    ("EQV", "Equivalent (von Mises)", "AddEquivalentStress"),
    ("S1",  "Maximum Principal",      "AddMaximumPrincipalStress"),
    ("S3",  "Minimum Principal",      "AddMinimumPrincipalStress"),
]


# =====================================================================
# POPUP
# =====================================================================
def _cmb(x, y, w):
    c = ComboBox(); c.Location = Point(x, y); c.Size = Size(w, 24)
    c.DropDownStyle = ComboBoxStyle.DropDownList; c.Font = CTRL_FONT
    def _block(s, e):
        try: e.Handled = True
        except: pass
    c.MouseWheel += _block
    return c

def _lbl(text, x, y, w):
    l = Label(); l.Text = text; l.Location = Point(x, y)
    l.Size = Size(w, 20); l.Font = HDR_FONT
    return l


class AvgDialog(Form):
    def __init__(self, analyses, ns_list):
        self._analyses = analyses; self._ns_list = ns_list
        self.Text = "Averaged / Unaveraged % Check"
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False; self.MinimizeBox = False
        self.BackColor = Color.White
        W = 400; y = 12

        self.Controls.Add(_lbl("Analysis:", 12, y, W)); y += 22
        self.cmb_an = _cmb(12, y, W)
        idx = 0
        for i, a in enumerate(analyses):
            at = str(a.AnalysisType)
            self.cmb_an.Items.Add("[{0}] {1}  ({2})".format(i + 1, a.Name, at))
            if idx == 0 and "static" in at.lower(): idx = i
        self.cmb_an.SelectedIndex = idx
        self.Controls.Add(self.cmb_an); y += 34

        self.Controls.Add(_lbl("Named Selection (Body):", 12, y, W)); y += 22
        self.cmb_ns = _cmb(12, y, W)
        for ns in ns_list: self.cmb_ns.Items.Add(ns.Name)
        self.cmb_ns.SelectedIndex = 0
        self.Controls.Add(self.cmb_ns); y += 36

        grp = GroupBox(); grp.Text = "Stress Type"; grp.Font = HDR_FONT
        grp.Location = Point(12, y); grp.Size = Size(W, 104)
        self._rbs = []
        for i, (_k, txt, _m) in enumerate(STRESS_TYPES):
            rb = RadioButton(); rb.Text = txt; rb.Font = CTRL_FONT
            rb.Location = Point(14, 22 + i * 26); rb.Size = Size(W - 30, 22)
            rb.Checked = (i == 0)
            grp.Controls.Add(rb); self._rbs.append(rb)
        self.Controls.Add(grp); y += 116

        ok = Button(); ok.Text = "Run"; ok.Size = Size(90, 30)
        ok.Font = HDR_FONT; ok.BackColor = BTN_BLUE; ok.ForeColor = Color.White
        ok.FlatStyle = 0; ok.Location = Point(W - 178, y); ok.Click += self._ok
        self.Controls.Add(ok)
        cn = Button(); cn.Text = "Cancel"; cn.Size = Size(90, 30); cn.Font = CTRL_FONT
        cn.Location = Point(W - 78, y); cn.DialogResult = DialogResult.Cancel
        self.Controls.Add(cn); self.CancelButton = cn
        self.ClientSize = Size(W + 24, y + 44)

    def _ok(self, s, e):
        self.analysis = self._analyses[self.cmb_an.SelectedIndex]
        self.ns       = self._ns_list[self.cmb_ns.SelectedIndex]
        self.stype    = [STRESS_TYPES[i] for i, rb in enumerate(self._rbs) if rb.Checked][0]
        self.DialogResult = DialogResult.OK; self.Close()


# =====================================================================
# HELPERS
# =====================================================================
def get_temp_path(name):
    tmp = os.environ.get("TEMP") or os.environ.get("TMP") or "C:\\Temp"
    if not os.path.exists(tmp): os.makedirs(tmp)
    return os.path.join(tmp, name)


def parse_export(filepath):
    """{node_id: [deger, ...]} — unaveraged export'ta ayni node birden
    fazla satirda (her eleman icin) gelebilir, hepsi toplanir."""
    data = {}
    try:
        with open(filepath, "r") as f: lines = f.readlines()
    except: return data
    for line in lines[1:]:
        parts = line.strip().split("\t")
        if len(parts) < 2: continue
        try:
            nid = int(parts[0].strip())
            val = float(parts[-1].strip().replace(",", "."))
        except: continue
        data.setdefault(nid, []).append(val)
    return data


def add_result(sol, add_method, ns, name, display_option):
    for ch in list(sol.Children):
        if ch.Name == name:
            try: ch.Delete()
            except: pass
    r = getattr(sol, add_method)()
    r.Name = name
    r.Location = ns
    r.DisplayOption = display_option
    try: r.CalculateTimeHistory = False
    except: pass
    return r


def build_dpf_code(node_vals, result_name):
    pairs = list(node_vals.items()); chunk = 150; lines = []
    for i in range(0, len(pairs), chunk):
        seg = ", ".join("({0}, {1!r})".format(k, v) for k, v in pairs[i:i + chunk])
        lines.append("    for _k, _v in [{0}]: node_vals[_k] = _v".format(seg))
    wf_key = "avgpct_" + result_name[:40].replace(" ", "_").replace("-", "_").replace("%", "")
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

    dlg = AvgDialog(analyses, ns_list)
    if dlg.ShowDialog() != DialogResult.OK: return

    sol = dlg.analysis.Solution; ns = dlg.ns
    key, stxt, add_method = dlg.stype
    base = "{0}_{1}".format(key, ns.Name[:25])

    # 1) Averaged + Unaveraged sonuc objeleri
    r_avg = add_result(sol, add_method, ns, base + "_Averaged",   AVG_ENUM.Averaged)
    r_una = add_result(sol, add_method, ns, base + "_Unaveraged", AVG_ENUM.Unaveraged)
    sol.EvaluateAllResults()

    # 2) Export + parse
    p_avg = get_temp_path("avgpct_avg_{0}.txt".format(key))
    p_una = get_temp_path("avgpct_una_{0}.txt".format(key))
    r_avg.ExportToTextFile(p_avg); r_una.ExportToTextFile(p_una)
    d_avg = parse_export(p_avg); d_una = parse_export(p_una)
    for p in (p_avg, p_una):
        try: os.remove(p)
        except: pass
    if not d_avg or not d_una:
        MessageBox.Show("Export bos geldi.", "Hata",
                        MessageBoxButtons.OK, MessageBoxIcon.Error); return

    # 3) Node bazinda % = abs((Unaveraged - Averaged) / Unaveraged) * 100
    #    Unaveraged'da node basina birden fazla deger varsa |max| olan alinir
    pct = {}
    for nid, avals in d_avg.items():
        uvals = d_una.get(nid)
        if not uvals: continue
        a = avals[0]
        u = max(uvals, key=abs)
        if abs(u) < 1e-12: continue
        pct[nid] = abs((u - a) / u) * 100.0
    if not pct:
        MessageBox.Show("Eslesen node bulunamadi.", "Hata",
                        MessageBoxButtons.OK, MessageBoxIcon.Error); return

    # 4) Python Result ile plot
    rname = base + "_Avg_Unavg_%"
    for ch in list(sol.Children):
        if ch.Name == rname:
            try: ch.Delete()
            except: pass
    pr = sol.AddPythonResult()
    pr.Name = rname
    pr.Text = build_dpf_code(pct, rname)
    try: pr.Connect()
    except: pass
    try: sol.EvaluateAllResults()
    except: pass
    try: pr.Activate()
    except: pass

    nmin = min(pct, key=pct.get)
    MessageBox.Show(
        "{0}  |  NS: {1}\nNodes: {2}\nMin %: {3:.2f}  @ Node {4}\nMax %: {5:.2f}".format(
            stxt, ns.Name, len(pct), pct[nmin], nmin, max(pct.values())),
        "Averaged / Unaveraged %", MessageBoxButtons.OK, MessageBoxIcon.Information)


run()
