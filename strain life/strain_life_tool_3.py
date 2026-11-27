# -*- coding: utf-8 -*-
# ============================================================
# Strain-Life Tool - ANSYS Mechanical (IronPython 2.7)
# v3: step-by-step windows + paste box (Strain vs Cycle)
#     Excel built hidden + busy-retry (fixes 0x800AC472), step info in errors
#
# Flow (wizard):
#   Step 1 window : Excel inputs, one per row  (B1, B2, B5, B6)
#   Step 2 window : Temperature, Serie, R ratio + paste box for
#                   Total Strain / Cycle (Ctrl+V, any number of rows)
#                   [< Back] [+ Add Another] [Create Excel]
#   Excel: one block per dataset on a single sheet
#        A1:B13  calculation block   (inputs in column B)
#        D1:E7   Temperature / Serie / R + power fit (A, aaa, R2)
#        D9:E..  Total Strain / Cycle data
#        chart   log-log scatter (X = Total Strain, Y = Cycle)
#                + Power trendline with equation (15 decimals)
#      B13 = A * (ABS(B12) ^ aaa)   (live formula)
#   Excel is left open - save it yourself.
# ============================================================

import os
import math
import clr
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

from System.Windows.Forms import (
    Form, Button, Label, TextBox, ScrollBars,
    DialogResult, FormStartPosition, MessageBox, MessageBoxButtons,
    MessageBoxIcon, FormBorderStyle, Cursor, Cursors, Application
)
from System.Drawing import Size, Point, Color, Font, FontStyle

# =====================================================================
# STYLE CONSTANTS  (same look as the other tools)
# =====================================================================
LBL_H     = 20
CTRL_FONT = Font("Segoe UI", 9)
HDR_FONT  = Font("Segoe UI", 9, FontStyle.Bold)
BTN_BLUE  = Color.FromArgb(0, 120, 215)
BTN_RED   = Color.FromArgb(180, 30, 30)

BLOCK_STRIDE = 18      # columns between dataset blocks on the sheet
DATA_ROW0    = 10      # first data row of the Total Strain / Cycle table
SHEET_NAME   = "Strain Life"


def _blue_btn(text, w, h):
    b = Button()
    b.Text = text; b.Size = Size(w, h)
    b.Font = HDR_FONT; b.BackColor = BTN_BLUE
    b.ForeColor = Color.White; b.FlatStyle = 0
    return b

def _grey_btn(text, w, h):
    b = Button()
    b.Text = text; b.Size = Size(w, h); b.Font = CTRL_FONT
    return b

def _red_btn(text, w, h):
    b = Button()
    b.Text = text; b.Size = Size(w, h)
    b.Font = CTRL_FONT; b.BackColor = BTN_RED
    b.ForeColor = Color.White; b.FlatStyle = 0
    return b


# =====================================================================
# PURE HELPERS  (no GUI / no Excel)
# =====================================================================
def to_float(s):
    """Parse '0,0123' / '1.5' / '1,234.5' / '1.234,5' -> float."""
    s = str(s).strip().replace(" ", "")
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):      # 1.234,5 (european)
            s = s.replace(".", "").replace(",", ".")
        else:                                 # 1,234.5
            s = s.replace(",", "")
    else:
        s = s.replace(",", ".")
    return float(s)


def parse_pairs(text):
    """Parse pasted text (Excel copy: tab separated) into [(x, y), ...].
    Lines that are not two numbers (headers etc.) are skipped.
    Returns (rows, skipped_line_numbers)."""
    rows = []; skipped = []
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    for i, line in enumerate(lines):
        if not line.strip():
            continue
        if "\t" in line:
            parts = line.split("\t")
        else:
            parts = line.replace(";", " ").split()
        parts = [p.strip() for p in parts if p.strip() != ""]
        if len(parts) < 2:
            skipped.append(i + 1); continue
        try:
            x = to_float(parts[0]); y = to_float(parts[1])
        except Exception:
            skipped.append(i + 1); continue
        rows.append((x, y))
    return rows, skipped


def power_fit(xs, ys):
    """Least squares on ln(y) vs ln(x)  ==  Excel power trendline.
    y = A * x^b.  Returns (A, b, r2)."""
    lx = [math.log(x) for x in xs]
    ly = [math.log(y) for y in ys]
    n = len(lx)
    mx = sum(lx) / n
    my = sum(ly) / n
    sxx = sum((a - mx) ** 2 for a in lx)
    sxy = sum((a - mx) * (c - my) for a, c in zip(lx, ly))
    syy = sum((c - my) ** 2 for c in ly)
    b = sxy / sxx
    A = math.exp(my - b * mx)
    r2 = (sxy * sxy) / (sxx * syy) if syy > 0 else 1.0
    return A, b, r2


def col_letter(n):
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def rgb(r, g, b):
    return r + g * 256 + b * 65536


def num_or_text(s):
    try:
        return to_float(s)
    except Exception:
        return str(s)


def calc_block(ds, A, b):
    """Python mirror of the sheet formulas (used for the summary and tests).
    Returns dict with B3,B7,B8,B10,B11,B12,B13 (None if undefined)."""
    B1 = ds.walker; B2 = ds.r_ratio; B5 = ds.max_strain; B6 = ds.min_strain
    out = {}
    try:
        B3 = B6 / B5
        B7 = (B5 - B6) / 2.0
        B8 = B5 - B6
        out["B3"] = B3; out["B7"] = B7; out["B8"] = B8
        B10 = ((B7 * (2.0 / (1.0 - B3))) * ((1.0 - B3) ** B1)) / ((1.0 - B2) ** B1)
        B11 = B2 * B10
        B12 = B10 - B11
        B13 = A * (abs(B12) ** b)
        out["B10"] = B10; out["B11"] = B11; out["B12"] = B12; out["B13"] = B13
    except Exception:
        pass
    return out


class DataSet(object):
    def __init__(self, temp, serie, r_ratio, walker, max_strain, min_strain, rows):
        self.temp = temp
        self.serie = serie
        self.r_ratio = r_ratio          # -> B2 (Available R Ratio, test)
        self.walker = walker            # -> B1
        self.max_strain = max_strain    # -> B5
        self.min_strain = min_strain    # -> B6
        self.rows = rows                # [(total_strain, cycle), ...]

    def label(self):
        return "T={0} | {1} | R={2:g}".format(self.temp, self.serie, self.r_ratio)


# =====================================================================
# EXCEL BUILDER  (COM)
# =====================================================================
_STEP = [""]      # last step name, shown in the error message


def _is_busy(ex):
    t = str(ex).upper()
    return ("800AC472" in t) or ("80010001" in t) or ("8001010A" in t)


def _pause(ms):
    try:
        from System.Threading import Thread
        Thread.Sleep(ms)
        Application.DoEvents()
    except Exception:
        pass


def _put(obj, attr, value, tries=8):
    """setattr on a COM object; retries when Excel answers 'busy'."""
    for i in range(tries):
        try:
            setattr(obj, attr, value)
            return
        except Exception as ex:
            if i == tries - 1 or not _is_busy(ex):
                raise
            _pause(300)


def _safe(warns, what, fn):
    """Run a cosmetic step; a failure is reported as a warning only."""
    try:
        fn()
    except Exception as ex:
        warns.append("{0}: {1}".format(what, str(ex).split("\n")[0]))


def _write_block(ws, k, ds):
    """Write dataset #k as a block. Returns (A, b, r2, warnings)."""
    import System
    warnings = []
    c0 = 1 + k * BLOCK_STRIDE
    B = col_letter(c0 + 1)      # value column  (B)
    E = col_letter(c0 + 4)      # info value col (E)

    def cell(r, c):
        return ws.Cells[r, c]

    def put(r, c, attr, value):
        _put(cell(r, c), attr, value)

    # ---- calculation block A1:B13 (core: values + formulas) -----------
    _STEP[0] = "block {0}: labels".format(k + 1)
    labels = {
        1: "Walker Coefficient",
        2: "Available R Ratio (Total Strain - Test)",
        3: "Demanded R Ratio (Total Strain)",
        5: "Available Max Strain",
        6: "Available Min Strain",
        7: "Available Alt. Strain",
        8: "Analysis Strain Range",
        10: "Calculated Max Strain",
        11: "Calculated Min Strain",
        12: "Calculated Range (Demanded R Ratio)",
        13: "Calculated Cycle",
    }
    for r, t in labels.items():
        put(r, c0, "Value2", t)

    _STEP[0] = "block {0}: inputs + formulas".format(k + 1)
    put(1, c0 + 1, "Value2", float(ds.walker))
    put(2, c0 + 1, "Value2", float(ds.r_ratio))
    put(5, c0 + 1, "Value2", float(ds.max_strain))
    put(6, c0 + 1, "Value2", float(ds.min_strain))

    put(3, c0 + 1, "Formula", "={b}6/{b}5".format(b=B))
    put(7, c0 + 1, "Formula", "=({b}5-{b}6)/2".format(b=B))
    put(8, c0 + 1, "Formula", "={b}5-{b}6".format(b=B))
    put(10, c0 + 1, "Formula",
        "=(({b}7*(2/(1-{b}3)))*((1-{b}3)^{b}1))/((1-{b}2)^{b}1)".format(b=B))
    put(11, c0 + 1, "Formula", "={b}2*{b}10".format(b=B))
    put(12, c0 + 1, "Formula", "={b}10-{b}11".format(b=B))
    put(13, c0 + 1, "Formula", "={e}5*(ABS({b}12)^{e}6)".format(b=B, e=E))

    # ---- info + fit results (D:E) --------------------------------------
    _STEP[0] = "block {0}: info + fit".format(k + 1)
    for r, t in ((1, "Temperature"), (2, "Serie"), (3, "R")):
        put(r, c0 + 3, "Value2", t)
    put(1, c0 + 4, "Value2", num_or_text(ds.temp))
    put(2, c0 + 4, "Value2", num_or_text(ds.serie))
    put(3, c0 + 4, "Formula", "={b}2".format(b=B))

    xs = [x for x, y in ds.rows]
    ys = [y for x, y in ds.rows]
    A, b, r2 = power_fit(xs, ys)

    put(5, c0 + 3, "Value2", "A  (power coefficient)")
    put(6, c0 + 3, "Value2", "aaa  (power exponent)")
    put(7, c0 + 3, "Value2", "R2  (fit)")
    put(5, c0 + 4, "Value2", A)
    put(6, c0 + 4, "Value2", b)
    put(7, c0 + 4, "Value2", r2)

    # ---- data table D9:E.. -------------------------------------------
    _STEP[0] = "block {0}: data table".format(k + 1)
    put(DATA_ROW0 - 1, c0 + 3, "Value2", "Total Strain")
    put(DATA_ROW0 - 1, c0 + 4, "Value2", "Cycle")
    n = len(ds.rows)
    arr = System.Array.CreateInstance(System.Object, n, 2)
    for i, (x, y) in enumerate(ds.rows):
        arr[i, 0] = float(x)
        arr[i, 1] = float(y)
    r_last = DATA_ROW0 + n - 1
    _put(ws.Range(cell(DATA_ROW0, c0 + 3), cell(r_last, c0 + 4)), "Value2", arr)

    # ---- cosmetics (failure -> warning only) ---------------------------
    _STEP[0] = "block {0}: formatting".format(k + 1)

    def _widths():
        for off, w in ((0, 38), (1, 14), (2, 3), (3, 22), (4, 16), (5, 3)):
            _put(ws.Range(cell(1, c0 + off), cell(1, c0 + off)).EntireColumn,
                 "ColumnWidth", w)

    def _numfmt():
        _put(ws.Range(cell(1, c0 + 1), cell(13, c0 + 1)), "NumberFormat", "0.0000")
        _put(cell(13, c0 + 1), "NumberFormat", "0.00")
        _put(cell(5, c0 + 4), "NumberFormat", "0.000000000000000")
        _put(cell(6, c0 + 4), "NumberFormat", "0.000000000000000")
        _put(cell(7, c0 + 4), "NumberFormat", "0.0000")
        _put(ws.Range(cell(DATA_ROW0, c0 + 3), cell(r_last, c0 + 3)),
             "NumberFormat", "0.0000")
        _put(ws.Range(cell(DATA_ROW0, c0 + 4), cell(r_last, c0 + 4)),
             "NumberFormat", "#,##0")

    def _borders():
        for a1, a2 in (((1, c0), (13, c0 + 1)),
                       ((1, c0 + 3), (3, c0 + 4)),
                       ((5, c0 + 3), (7, c0 + 4))):
            _put(ws.Range(cell(*a1), cell(*a2)).Borders, "LineStyle", 1)

    def _fills():
        for r in (1, 2, 5, 6):                      # input cells: light blue
            _put(cell(r, c0 + 1).Interior, "Color", rgb(221, 235, 247))
        for r in (1, 2, 3):                         # info labels: yellow
            _put(cell(r, c0 + 3).Interior, "Color", rgb(255, 255, 0))
            _put(cell(r, c0 + 3).Font, "Bold", True)
        hdr = ws.Range(cell(DATA_ROW0 - 1, c0 + 3), cell(DATA_ROW0 - 1, c0 + 4))
        _put(hdr.Font, "Bold", True)
        _put(hdr.Interior, "Color", rgb(217, 217, 217))

    _safe(warnings, "column widths", _widths)
    _safe(warnings, "number formats", _numfmt)
    _safe(warnings, "borders", _borders)
    _safe(warnings, "fills/fonts", _fills)

    # ---- chart ---------------------------------------------------------
    _STEP[0] = "block {0}: chart".format(k + 1)
    try:
        left = cell(1, c0 + 6).Left
        top  = cell(1, c0 + 6).Top
        co = ws.ChartObjects().Add(left, top, 480, 330)
        ch = co.Chart
        _put(ch, "ChartType", -4169)              # xlXYScatter (markers only)
        s = ch.SeriesCollection().NewSeries()
        _put(s, "XValues", ws.Range(cell(DATA_ROW0, c0 + 3), cell(r_last, c0 + 3)))
        _put(s, "Values",  ws.Range(cell(DATA_ROW0, c0 + 4), cell(r_last, c0 + 4)))
        _put(s, "Name", "{0} / {1}".format(ds.temp, ds.serie))
        try:
            _put(s, "MarkerStyle", 8); _put(s, "MarkerSize", 6)   # circle
        except Exception:
            pass
        _put(ch, "HasLegend", False)
        _put(ch, "HasTitle", True)
        _put(ch.ChartTitle, "Text", "Temp: {0} | Serie: {1} | R: {2:g}".format(
            ds.temp, ds.serie, ds.r_ratio))
        for axis_id, title in ((1, "Total Strain"), (2, "Cycle")):
            ax = ch.Axes(axis_id)
            _put(ax, "ScaleType", -4133)          # xlScaleLogarithmic
            _put(ax, "HasTitle", True)
            _put(ax.AxisTitle, "Text", title)
            _put(ax, "HasMajorGridlines", True)
            _put(ax, "HasMinorGridlines", True)
        _STEP[0] = "block {0}: trendline".format(k + 1)
        try:
            tl = s.Trendlines().Add(4)            # xlPower
            _put(tl, "DisplayEquation", True)
            _put(tl, "DisplayRSquared", False)
            try:
                _put(tl.DataLabel, "NumberFormat", "0.000000000000000")   # 15 decimals
                _put(tl.DataLabel.Font, "Size", 9)
            except Exception as ex2:
                warnings.append("trendline label format: " + str(ex2).split("\n")[0])
        except Exception as ex1:
            warnings.append("trendline: " + str(ex1).split("\n")[0])
    except Exception as ex:
        warnings.append("chart: " + str(ex).split("\n")[0])

    return A, b, r2, warnings


def build_excel(datasets):
    """Create a new Excel workbook with one block per dataset.
    Excel is built hidden and shown at the end (avoids 'Excel busy' errors).
    Returns list of (dataset, A, b, r2, calc_dict, warnings)."""
    import System
    _STEP[0] = "starting Excel"
    excel_type = System.Type.GetTypeFromProgID("Excel.Application")
    if excel_type is None:
        raise Exception("Excel is not installed (ProgID Excel.Application not found).")
    app = System.Activator.CreateInstance(excel_type)
    ok = False
    try:
        try:
            app.Visible = False
            app.DisplayAlerts = False
        except Exception:
            pass
        # wait until Excel is ready (max ~10 s)
        for i in range(50):
            try:
                if bool(app.Ready):
                    break
            except Exception:
                break
            _pause(200)

        _STEP[0] = "creating workbook"
        wb = None
        for i in range(8):
            try:
                wb = app.Workbooks.Add()
                break
            except Exception as ex:
                if i == 7 or not _is_busy(ex):
                    raise
                _pause(400)
        ws = wb.Sheets[1]
        try:
            ws.Name = SHEET_NAME
        except Exception:
            pass

        out = []
        for k, ds in enumerate(datasets):
            A, b, r2, warns = _write_block(ws, k, ds)
            out.append((ds, A, b, r2, calc_block(ds, A, b), warns))

        _STEP[0] = "showing Excel"
        try:
            app.Calculate()
        except Exception:
            pass
        try:
            app.DisplayAlerts = True
        except Exception:
            pass
        try:
            app.UserControl = True
        except Exception:
            pass
        _put(app, "Visible", True)
        try:
            ws.Activate()
        except Exception:
            pass
        ok = True
        return out
    finally:
        if not ok:                      # do not leave a hidden Excel behind
            try:
                app.Quit()
            except Exception:
                pass


# =====================================================================
# STEP 1 WINDOW - Excel inputs (column B), one per row
# =====================================================================
class InputsDialog(Form):
    FIELDS = [("walker", "Walker Coefficient",                    "B1"),
              ("r",      "Available R Ratio (Total Strain - Test)", "B2"),
              ("max",    "Available Max Strain",                  "B5"),
              ("min",    "Available Min Strain",                  "B6")]

    def __init__(self, n, prev, n_done):
        self.action = "cancel"
        self.values = None
        self.Text            = "Strain-Life  -  Step 1/2  -  Dataset #{0}".format(n)
        self.StartPosition   = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox     = False; self.MinimizeBox = False
        self.ClientSize      = Size(480, 285); self.BackColor = Color.White

        lbl = Label()
        lbl.Text = "Excel inputs  (column B)  -  Dataset #{0}".format(n)
        lbl.Location = Point(20, 15); lbl.Size = Size(440, LBL_H)
        lbl.Font = HDR_FONT; self.Controls.Add(lbl)

        self._tbs = {}
        y = 50
        for key, text, cell in self.FIELDS:
            l = Label()
            l.Text = "{0}   ({1})".format(text, cell)
            l.Location = Point(20, y + 3); l.Size = Size(310, LBL_H)
            l.Font = CTRL_FONT; self.Controls.Add(l)
            tb = TextBox()
            tb.Location = Point(340, y); tb.Size = Size(120, 24); tb.Font = CTRL_FONT
            if prev is not None and key in prev:
                tb.Text = repr(prev[key])
            self.Controls.Add(tb)
            self._tbs[key] = tb
            y += 38

        btn_next = _blue_btn("Next  ->", 90, 32)
        btn_next.Location = Point(140, 235); btn_next.Click += self._next
        self.Controls.Add(btn_next)
        self.AcceptButton = btn_next

        if n_done > 0:
            btn_fin = _grey_btn("Create Excel ({0})".format(n_done), 125, 32)
            btn_fin.Location = Point(237, 235); btn_fin.Click += self._finish
            self.Controls.Add(btn_fin)

        btn_cancel = _grey_btn("Cancel", 85, 32)
        btn_cancel.Location = Point(375, 235)
        btn_cancel.DialogResult = DialogResult.Cancel
        self.Controls.Add(btn_cancel)
        self.CancelButton = btn_cancel

    def _finish(self, s, e):
        self.action = "finish"
        self.DialogResult = DialogResult.OK; self.Close()

    def _next(self, s, e):
        errs = []; vals = {}
        for key, text, cell in self.FIELDS:
            try:
                vals[key] = to_float(self._tbs[key].Text)
            except Exception:
                errs.append("- {0} ({1}) is not a valid number.".format(text, cell))
        if not errs:
            if vals["r"] == 1.0:
                errs.append("- R ratio (B2) cannot be 1 (division by zero).")
            if vals["max"] == 0.0:
                errs.append("- Max strain (B5) cannot be 0 (B3 = B6 / B5).")
            if vals["max"] == vals["min"]:
                errs.append("- Max strain and Min strain are equal (Demanded R = 1).")
        if errs:
            MessageBox.Show("\n".join(errs), "Check inputs",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        self.values = vals
        self.action = "next"
        self.DialogResult = DialogResult.OK; self.Close()


# =====================================================================
# STEP 2 WINDOW - Temperature / Serie / R ratio + paste box
# =====================================================================
class DataDialog(Form):
    def __init__(self, n, inputs, state):
        self.action  = "cancel"
        self.dataset = None
        self._inputs = inputs
        self._state  = state
        self.Text            = "Strain-Life  -  Step 2/2  -  Dataset #{0}".format(n)
        self.StartPosition   = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox     = False; self.MinimizeBox = False
        self.ClientSize      = Size(560, 560); self.BackColor = Color.White

        self._tb_temp  = self._field("Temperature:", 15, 15, 95, 120, state.get("temp", ""))
        self._tb_serie = self._field("Serie:",      200, 15, 60, 120, state.get("serie", ""))
        self._tb_r     = self._field("R ratio:",    385, 15, 60, 90, repr(inputs["r"]))

        lbl = Label()
        lbl.Text = ("Copy the 2 columns  (Total Strain | Cycle)  from Excel, click in the "
                    "box below and press Ctrl+V.  Any number of rows.")
        lbl.Location = Point(15, 55); lbl.Size = Size(530, 34)
        lbl.Font = CTRL_FONT; lbl.ForeColor = Color.FromArgb(0, 100, 200)
        self.Controls.Add(lbl)

        self._box = TextBox()
        self._box.Multiline  = True
        self._box.WordWrap   = False
        self._box.ScrollBars = ScrollBars.Both
        self._box.AcceptsReturn = True
        self._box.MaxLength  = 100000000
        self._box.Font       = Font("Consolas", 9)
        self._box.Location   = Point(15, 92); self._box.Size = Size(530, 385)
        self._box.Text       = state.get("text", "")
        self._box.TextChanged += self._update_count
        self.Controls.Add(self._box)

        self._lbl_count = Label()
        self._lbl_count.Location = Point(15, 483); self._lbl_count.Size = Size(440, LBL_H)
        self._lbl_count.Font = CTRL_FONT; self.Controls.Add(self._lbl_count)

        btn_clr = _red_btn("Clear", 70, 24)
        btn_clr.Location = Point(475, 481); btn_clr.Click += self._clear
        self.Controls.Add(btn_clr)

        btn_back = _grey_btn("<  Back", 80, 32)
        btn_back.Location = Point(15, 517); btn_back.Click += self._back
        self.Controls.Add(btn_back)

        btn_add = _grey_btn("+ Add Another", 115, 32)
        btn_add.Location = Point(210, 517); btn_add.Click += self._add
        self.Controls.Add(btn_add)

        btn_fin = _blue_btn("Create Excel", 120, 32)
        btn_fin.Location = Point(332, 517); btn_fin.Click += self._finish
        self.Controls.Add(btn_fin)

        btn_cancel = _grey_btn("Cancel", 80, 32)
        btn_cancel.Location = Point(465, 517)
        btn_cancel.DialogResult = DialogResult.Cancel
        self.Controls.Add(btn_cancel)
        self.CancelButton = btn_cancel

        self._update_count(None, None)

    def _field(self, text, x, y, label_w, tb_w, default=""):
        lbl = Label(); lbl.Text = text
        lbl.Location = Point(x, y + 3); lbl.Size = Size(label_w, LBL_H)
        lbl.Font = CTRL_FONT; self.Controls.Add(lbl)
        tb = TextBox(); tb.Location = Point(x + label_w, y)
        tb.Size = Size(tb_w, 24); tb.Font = CTRL_FONT; tb.Text = default
        self.Controls.Add(tb)
        return tb

    def _update_count(self, s, e):
        rows, skipped = parse_pairs(self._box.Text)
        txt = "Valid rows: {0}".format(len(rows))
        if skipped:
            txt += "     Skipped lines (header / not numeric): {0}".format(len(skipped))
        self._lbl_count.Text = txt

    def _clear(self, s, e):
        self._box.Text = ""

    def _save_state(self):
        self._state["temp"]  = self._tb_temp.Text
        self._state["serie"] = self._tb_serie.Text
        self._state["text"]  = self._box.Text

    def _close_with(self, action):
        self.action = action
        self.DialogResult = DialogResult.OK; self.Close()

    def _back(self, s, e):
        self._save_state()
        self._close_with("back")

    def _build(self):
        """Validate and build the DataSet. Returns True on success."""
        errs = []
        temp  = self._tb_temp.Text.strip()
        serie = self._tb_serie.Text.strip()
        if not temp:  errs.append("- Temperature is empty.")
        if not serie: errs.append("- Serie is empty.")
        r = None
        try:
            r = to_float(self._tb_r.Text)
            if r == 1.0:
                errs.append("- R ratio cannot be 1 (division by zero).")
        except Exception:
            errs.append("- R ratio is not a valid number.")

        rows, skipped = parse_pairs(self._box.Text)
        bad = [i + 1 for i, (x, y) in enumerate(rows) if x <= 0 or y <= 0]
        if bad:
            errs.append("- {0} row(s) have Strain or Cycle <= 0 (first: pasted row {1}). "
                        "Log-log / power fit needs values > 0.".format(len(bad), bad[0]))
        if len(rows) < 2:
            errs.append("- At least 2 valid (Total Strain, Cycle) rows are needed.")
        elif len(set(x for x, y in rows)) < 2:
            errs.append("- Total Strain values are all identical - cannot fit.")
        if errs:
            MessageBox.Show("\n".join(errs), "Check inputs",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return False
        self.dataset = DataSet(temp, serie, r, self._inputs["walker"],
                               self._inputs["max"], self._inputs["min"], rows)
        return True

    def _add(self, s, e):
        if self._build():
            self._close_with("add")

    def _finish(self, s, e):
        if self._build():
            self._close_with("finish")


# =====================================================================
# SUMMARY + MAIN FLOW
# =====================================================================
def show_summary(res):
    lines = ["Excel created (not saved - please save it yourself).\n"]
    for ds, A, b, r2, calc, warns in res:
        lines.append(ds.label())
        lines.append("  y = {0:.15f} x^{1:.15f}   (R2 = {2:.4f})".format(A, b, r2))
        if "B13" in calc:
            lines.append("  Calculated Cycle (B13) = {0:.2f}".format(calc["B13"]))
        else:
            lines.append("  Calculated Cycle: undefined (check inputs)")
        for w in warns:
            lines.append("  WARNING: " + w)
    MessageBox.Show("\n".join(lines), "Strain-Life",
                    MessageBoxButtons.OK, MessageBoxIcon.Information)


def run_wizard():
    """Returns list of DataSet, or None if cancelled."""
    datasets = []
    inputs = None
    state = {}
    step = 1
    while True:
        if step == 1:
            d1 = InputsDialog(len(datasets) + 1, inputs, len(datasets))
            d1.ShowDialog()
            if d1.action == "next":
                inputs = d1.values; step = 2
            elif d1.action == "finish":
                return datasets
            else:
                return None
        else:
            d2 = DataDialog(len(datasets) + 1, inputs, state)
            d2.ShowDialog()
            if d2.action == "back":
                step = 1
            elif d2.action == "add":
                datasets.append(d2.dataset); state = {}; step = 1
            elif d2.action == "finish":
                datasets.append(d2.dataset)
                return datasets
            else:
                return None


def main():
    datasets = run_wizard()
    if not datasets:
        return
    try:
        Cursor.Current = Cursors.WaitCursor
        res = build_excel(datasets)
        Cursor.Current = Cursors.Default
    except Exception as ex:
        Cursor.Current = Cursors.Default
        tail = ""
        try:
            import traceback
            tail = "\n".join(traceback.format_exc().strip().split("\n")[-5:])
        except Exception:
            pass
        MessageBox.Show("Excel error  (step: {0}):\n{1}\n\n{2}".format(
            _STEP[0], str(ex).split("\n")[0], tail), "Error",
            MessageBoxButtons.OK, MessageBoxIcon.Error)
        return
    show_summary(res)


if not os.environ.get("STRAIN_LIFE_NO_RUN"):
    main()
