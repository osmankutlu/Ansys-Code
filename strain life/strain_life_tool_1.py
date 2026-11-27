# -*- coding: utf-8 -*-
# ============================================================
# Strain-Life Tool - ANSYS Mechanical (IronPython 2.7)
# v1: Walker equivalent strain + power fit (Strain vs Cycle)
#
# Flow:
#   1) Main window: list of datasets  (Add / Edit / Remove)
#   2) Dataset popup: Temperature, Serie, R ratio, Walker coef,
#      Max/Min strain + Total Strain / Cycle table (Ctrl+V paste,
#      unlimited rows)
#   3) "Create Excel": one block per dataset on a single sheet
#        A1:B13  calculation block   (inputs in column B)
#        D1:E7   Temperature / Serie / R + power fit (A, aaa, R2)
#        D9:E..  Total Strain / Cycle data
#        chart   log-log scatter (X = Total Strain, Y = Cycle)
#                + Power trendline with equation (15 decimals)
#      B13 = A * ABS(B12) ^ aaa   (live formula)
#   Excel is left open - save it yourself.
# ============================================================

import os
import math
import clr
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

from System.Windows.Forms import (
    Form, Button, Label, TextBox, ListView, ListViewItem, View,
    DataGridView, DataGridViewTextBoxColumn,
    DataGridViewAutoSizeColumnsMode,
    DialogResult, FormStartPosition, MessageBox, MessageBoxButtons,
    MessageBoxIcon, FormBorderStyle, Clipboard, Keys, Cursor, Cursors
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
# DATASET POPUP  (inputs + paste table)
# =====================================================================
class DataSetDialog(Form):
    def __init__(self, ds=None):
        self.result = None
        self.Text            = "Add / Edit Dataset"
        self.StartPosition   = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox     = False; self.MinimizeBox = False
        self.ClientSize      = Size(500, 565); self.BackColor = Color.White

        def val(attr, default=""):
            return str(getattr(ds, attr)) if ds is not None else default

        self._tb_temp   = self._field("Temperature:",       15, 15, 110, 105, val("temp"))
        self._tb_serie  = self._field("Serie:",            265, 15, 110, 105, val("serie"))
        self._tb_r      = self._field("R ratio (B2):",      15, 45, 110, 105, val("r_ratio", "0"))
        self._tb_walker = self._field("Walker coef. (B1):", 265, 45, 110, 105, val("walker"))
        self._tb_max    = self._field("Max strain (B5):",   15, 75, 110, 105, val("max_strain"))
        self._tb_min    = self._field("Min strain (B6):",  265, 75, 110, 105, val("min_strain"))

        lbl = Label()
        lbl.Text = "Total Strain  /  Cycle   (Ctrl+V: paste 2 columns from Excel, any number of rows)"
        lbl.Location = Point(15, 112); lbl.Size = Size(470, LBL_H)
        lbl.Font = CTRL_FONT; lbl.ForeColor = Color.FromArgb(0, 100, 200)
        self.Controls.Add(lbl)

        self._grid = DataGridView()
        self._grid.Location              = Point(15, 136)
        self._grid.Size                  = Size(470, 330)
        self._grid.AllowUserToAddRows    = False
        self._grid.AllowUserToDeleteRows = False
        self._grid.RowHeadersVisible     = False
        self._grid.AutoSizeColumnsMode   = DataGridViewAutoSizeColumnsMode.Fill
        self._grid.Font                  = CTRL_FONT
        col_x = DataGridViewTextBoxColumn()
        col_x.Name = "strain"; col_x.HeaderText = "Total Strain"
        col_y = DataGridViewTextBoxColumn()
        col_y.Name = "cycle"; col_y.HeaderText = "Cycle"
        self._grid.Columns.Add(col_x); self._grid.Columns.Add(col_y)
        if ds is not None and ds.rows:
            for x, y in ds.rows:
                self._grid.Rows.Add(repr(x), repr(y))
        else:
            self._grid.Rows.Add("", "")
        self._grid.KeyDown += self._on_keydown
        self.Controls.Add(self._grid)

        btn_add = _grey_btn("+ Add Row", 95, 26)
        btn_add.Location = Point(15, 472)
        btn_add.Click += self._add_row
        self.Controls.Add(btn_add)

        btn_del = _grey_btn("- Delete Row", 105, 26)
        btn_del.Location = Point(116, 472)
        btn_del.Click += self._del_row
        self.Controls.Add(btn_del)

        btn_clr = _red_btn("Clear All", 85, 26)
        btn_clr.Location = Point(227, 472)
        btn_clr.Click += self._clear_all
        self.Controls.Add(btn_clr)

        self._lbl_count = Label()
        self._lbl_count.Location = Point(330, 477); self._lbl_count.Size = Size(155, LBL_H)
        self._lbl_count.Font = CTRL_FONT; self._lbl_count.TextAlign = 64  # MiddleRight
        self.Controls.Add(self._lbl_count)
        self._update_count()

        btn_ok = _blue_btn("OK", 90, 32)
        btn_ok.Location = Point(300, 520); btn_ok.Click += self._ok
        self.Controls.Add(btn_ok)

        btn_cancel = _grey_btn("Cancel", 90, 32)
        btn_cancel.Location = Point(395, 520)
        btn_cancel.DialogResult = DialogResult.Cancel
        self.Controls.Add(btn_cancel)
        self.CancelButton = btn_cancel

    # ---- small UI helpers --------------------------------------------
    def _field(self, text, x, y, label_w, tb_w, default=""):
        lbl = Label(); lbl.Text = text
        lbl.Location = Point(x, y + 3); lbl.Size = Size(label_w, LBL_H)
        lbl.Font = CTRL_FONT; self.Controls.Add(lbl)
        tb = TextBox(); tb.Location = Point(x + label_w, y)
        tb.Size = Size(tb_w, 24); tb.Font = CTRL_FONT; tb.Text = default
        self.Controls.Add(tb)
        return tb

    def _update_count(self):
        n = 0
        for i in range(self._grid.Rows.Count):
            a = self._grid.Rows[i].Cells["strain"].Value
            b = self._grid.Rows[i].Cells["cycle"].Value
            if (a is not None and str(a).strip()) or (b is not None and str(b).strip()):
                n += 1
        self._lbl_count.Text = "Rows: {0}".format(n)

    def _add_row(self, s, e):
        self._grid.Rows.Add("", "")
        self._update_count()

    def _del_row(self, s, e):
        if self._grid.SelectedRows.Count > 0:
            self._grid.Rows.RemoveAt(self._grid.SelectedRows[0].Index)
        elif self._grid.CurrentCell is not None:
            self._grid.Rows.RemoveAt(self._grid.CurrentCell.RowIndex)
        elif self._grid.Rows.Count > 0:
            self._grid.Rows.RemoveAt(self._grid.Rows.Count - 1)
        self._update_count()

    def _clear_all(self, s, e):
        self._grid.Rows.Clear()
        self._grid.Rows.Add("", "")
        self._update_count()

    def _on_keydown(self, sender, ev):
        if not (ev.Control and ev.KeyCode == Keys.V):
            return
        try:
            clip = Clipboard.GetText()
            if not clip:
                return
            lines = clip.replace("\r\n", "\n").replace("\r", "\n").split("\n")
            rows = []
            for line in lines:
                if not line.strip():
                    continue
                parts = line.split("\t")
                if len(parts) < 2:
                    parts = line.replace(";", " ").split()
                if len(parts) < 2:
                    continue
                try:                       # header lines are skipped
                    x = to_float(parts[0]); y = to_float(parts[1])
                except Exception:
                    continue
                rows.append((x, y))
            ev.Handled = True; ev.SuppressKeyPress = True
            if not rows:
                return
            self._grid.SuspendLayout()
            self._grid.Rows.Clear()
            for x, y in rows:
                self._grid.Rows.Add(repr(x), repr(y))
            self._grid.ResumeLayout()
            self._update_count()
        except Exception as ex:
            MessageBox.Show("Paste error: " + str(ex), "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

    # ---- validation ----------------------------------------------------
    def _read_rows(self, errs):
        rows = []
        for i in range(self._grid.Rows.Count):
            a = self._grid.Rows[i].Cells["strain"].Value
            b = self._grid.Rows[i].Cells["cycle"].Value
            sa = "" if a is None else str(a).strip()
            sb = "" if b is None else str(b).strip()
            if not sa and not sb:
                continue
            try:
                x = to_float(sa); y = to_float(sb)
            except Exception:
                errs.append("- Table row {0}: value is not a number.".format(i + 1))
                continue
            if x <= 0 or y <= 0:
                errs.append("- Table row {0}: Strain and Cycle must be > 0 "
                            "(log-log / power fit).".format(i + 1))
                continue
            rows.append((x, y))
        return rows

    def _ok(self, s, e):
        errs = []
        temp  = self._tb_temp.Text.strip()
        serie = self._tb_serie.Text.strip()
        if not temp:  errs.append("- Temperature is empty.")
        if not serie: errs.append("- Serie is empty.")

        nums = {}
        for key, tb, name in [("r",      self._tb_r,      "R ratio"),
                              ("walker", self._tb_walker, "Walker coefficient"),
                              ("max",    self._tb_max,    "Max strain"),
                              ("min",    self._tb_min,    "Min strain")]:
            try:
                nums[key] = to_float(tb.Text)
            except Exception:
                errs.append("- {0} is not a valid number.".format(name))

        if "r" in nums and nums["r"] == 1.0:
            errs.append("- R ratio cannot be 1 (division by zero).")
        if "max" in nums and nums["max"] == 0.0:
            errs.append("- Max strain cannot be 0 (B3 = B6 / B5).")
        if "max" in nums and "min" in nums and nums["max"] == nums["min"]:
            errs.append("- Max strain and Min strain are equal (Demanded R = 1).")

        rows = self._read_rows(errs)
        if len(rows) < 2:
            errs.append("- At least 2 valid (Total Strain, Cycle) rows are needed.")
        elif len(set(x for x, y in rows)) < 2:
            errs.append("- Total Strain values are all identical - cannot fit.")

        if errs:
            MessageBox.Show("\n".join(errs), "Check inputs",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        self.result = DataSet(temp, serie, nums["r"], nums["walker"],
                              nums["max"], nums["min"], rows)
        self.DialogResult = DialogResult.OK
        self.Close()


# =====================================================================
# EXCEL BUILDER  (COM)
# =====================================================================
def _write_block(ws, k, ds):
    """Write dataset #k as a block. Returns (A, b, r2, warnings)."""
    import System
    warnings = []
    c0 = 1 + k * BLOCK_STRIDE
    L = col_letter(c0)          # label column  (A)
    B = col_letter(c0 + 1)      # value column  (B)
    E = col_letter(c0 + 4)      # info value col (E)

    def cell(r, c):
        return ws.Cells[r, c]

    # ---- column widths ---------------------------------------------
    for off, w in ((0, 38), (1, 14), (2, 3), (3, 22), (4, 16), (5, 3)):
        ws.Range(cell(1, c0 + off), cell(1, c0 + off)).EntireColumn.ColumnWidth = w

    # ---- calculation block A1:B13 ------------------------------------
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
        cell(r, c0).Value2 = t

    # inputs (column B)
    cell(1, c0 + 1).Value2 = float(ds.walker)
    cell(2, c0 + 1).Value2 = float(ds.r_ratio)
    cell(5, c0 + 1).Value2 = float(ds.max_strain)
    cell(6, c0 + 1).Value2 = float(ds.min_strain)

    # formulas
    cell(3, c0 + 1).Formula  = "={b}6/{b}5".format(b=B)
    cell(7, c0 + 1).Formula  = "=({b}5-{b}6)/2".format(b=B)
    cell(8, c0 + 1).Formula  = "={b}5-{b}6".format(b=B)
    cell(10, c0 + 1).Formula = ("=(({b}7*(2/(1-{b}3)))*((1-{b}3)^{b}1))/((1-{b}2)^{b}1)"
                                .format(b=B))
    cell(11, c0 + 1).Formula = "={b}2*{b}10".format(b=B)
    cell(12, c0 + 1).Formula = "={b}10-{b}11".format(b=B)
    cell(13, c0 + 1).Formula = "={e}5*(ABS({b}12)^{e}6)".format(b=B, e=E)

    rng_vals = ws.Range(cell(1, c0 + 1), cell(13, c0 + 1))
    rng_vals.NumberFormat = "0.0000"
    cell(13, c0 + 1).NumberFormat = "0.00"
    ws.Range(cell(1, c0), cell(13, c0 + 1)).Borders.LineStyle = 1
    for r in (1, 2, 5, 6):                      # input cells: light blue
        cell(r, c0 + 1).Interior.Color = rgb(221, 235, 247)

    # ---- info + fit results (D:E) --------------------------------------
    for r, t in ((1, "Temperature"), (2, "Serie"), (3, "R")):
        cell(r, c0 + 3).Value2 = t
        cell(r, c0 + 3).Interior.Color = rgb(255, 255, 0)
        cell(r, c0 + 3).Font.Bold = True
    cell(1, c0 + 4).Value2 = num_or_text(ds.temp)
    cell(2, c0 + 4).Value2 = num_or_text(ds.serie)
    cell(3, c0 + 4).Formula = "={b}2".format(b=B)

    xs = [x for x, y in ds.rows]
    ys = [y for x, y in ds.rows]
    A, b, r2 = power_fit(xs, ys)

    cell(5, c0 + 3).Value2 = "A  (power coefficient)"
    cell(6, c0 + 3).Value2 = "aaa  (power exponent)"
    cell(7, c0 + 3).Value2 = "R2  (fit)"
    cell(5, c0 + 4).Value2 = A
    cell(6, c0 + 4).Value2 = b
    cell(7, c0 + 4).Value2 = r2
    cell(5, c0 + 4).NumberFormat = "0.000000000000000"
    cell(6, c0 + 4).NumberFormat = "0.000000000000000"
    cell(7, c0 + 4).NumberFormat = "0.0000"
    ws.Range(cell(1, c0 + 3), cell(3, c0 + 4)).Borders.LineStyle = 1
    ws.Range(cell(5, c0 + 3), cell(7, c0 + 4)).Borders.LineStyle = 1

    # ---- data table D9:E.. ---------------------------------------------
    cell(DATA_ROW0 - 1, c0 + 3).Value2 = "Total Strain"
    cell(DATA_ROW0 - 1, c0 + 4).Value2 = "Cycle"
    ws.Range(cell(DATA_ROW0 - 1, c0 + 3), cell(DATA_ROW0 - 1, c0 + 4)).Font.Bold = True
    ws.Range(cell(DATA_ROW0 - 1, c0 + 3), cell(DATA_ROW0 - 1, c0 + 4)).Interior.Color = rgb(217, 217, 217)

    n = len(ds.rows)
    arr = System.Array.CreateInstance(System.Object, n, 2)
    for i, (x, y) in enumerate(ds.rows):
        arr[i, 0] = float(x)
        arr[i, 1] = float(y)
    r_last = DATA_ROW0 + n - 1
    ws.Range(cell(DATA_ROW0, c0 + 3), cell(r_last, c0 + 4)).Value2 = arr
    ws.Range(cell(DATA_ROW0, c0 + 3), cell(r_last, c0 + 3)).NumberFormat = "0.0000"
    ws.Range(cell(DATA_ROW0, c0 + 4), cell(r_last, c0 + 4)).NumberFormat = "#,##0"

    # ---- chart ---------------------------------------------------------
    try:
        left = cell(1, c0 + 6).Left
        top  = cell(1, c0 + 6).Top
        co = ws.ChartObjects().Add(left, top, 480, 330)
        ch = co.Chart
        ch.ChartType = -4169                      # xlXYScatter (markers only)
        s = ch.SeriesCollection().NewSeries()
        s.XValues = ws.Range(cell(DATA_ROW0, c0 + 3), cell(r_last, c0 + 3))
        s.Values  = ws.Range(cell(DATA_ROW0, c0 + 4), cell(r_last, c0 + 4))
        s.Name    = "{0} / {1}".format(ds.temp, ds.serie)
        try:
            s.MarkerStyle = 8; s.MarkerSize = 6   # circle
        except Exception:
            pass
        ch.HasLegend = False
        ch.HasTitle = True
        ch.ChartTitle.Text = "Temp: {0} | Serie: {1} | R: {2:g}".format(
            ds.temp, ds.serie, ds.r_ratio)
        for axis_id, title in ((1, "Total Strain"), (2, "Cycle")):
            ax = ch.Axes(axis_id)
            ax.ScaleType = -4133                  # xlScaleLogarithmic
            ax.HasTitle = True
            ax.AxisTitle.Text = title
            ax.HasMajorGridlines = True
            ax.HasMinorGridlines = True
        try:
            tl = s.Trendlines().Add(4)            # xlPower
            tl.DisplayEquation = True
            tl.DisplayRSquared = False
            try:
                tl.DataLabel.NumberFormat = "0.000000000000000"   # 15 decimals
                tl.DataLabel.Font.Size = 9
            except Exception as ex2:
                warnings.append("trendline label format: " + str(ex2))
        except Exception as ex1:
            warnings.append("trendline: " + str(ex1))
    except Exception as ex:
        warnings.append("chart: " + str(ex))

    return A, b, r2, warnings


def build_excel(datasets):
    """Create a new Excel workbook with one block per dataset.
    Returns list of (dataset, A, b, r2, calc_dict, warnings)."""
    import System
    excel_type = System.Type.GetTypeFromProgID("Excel.Application")
    if excel_type is None:
        raise Exception("Excel is not installed (ProgID Excel.Application not found).")
    app = System.Activator.CreateInstance(excel_type)
    app.Visible = True
    wb = app.Workbooks.Add()
    ws = wb.Sheets[1]
    try:
        ws.Name = SHEET_NAME
    except Exception:
        pass
    app.ScreenUpdating = False
    out = []
    try:
        for k, ds in enumerate(datasets):
            A, b, r2, warns = _write_block(ws, k, ds)
            out.append((ds, A, b, r2, calc_block(ds, A, b), warns))
        try:
            app.Calculate()
        except Exception:
            pass
    finally:
        app.ScreenUpdating = True
    try:
        ws.Activate()
        ws.Range("A1").Select()
    except Exception:
        pass
    return out


# =====================================================================
# MAIN WINDOW
# =====================================================================
class StrainLifeForm(Form):
    def __init__(self):
        self._datasets = []
        self.Text            = "Strain-Life Tool  (Walker / Power Fit)"
        self.StartPosition   = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox     = False; self.MinimizeBox = False
        self.ClientSize      = Size(640, 400); self.BackColor = Color.White

        lbl = Label()
        lbl.Text = "Datasets  (each = one Temperature / Serie / R ratio table):"
        lbl.Location = Point(15, 12); lbl.Size = Size(610, LBL_H)
        lbl.Font = HDR_FONT; self.Controls.Add(lbl)

        self._lv = ListView()
        self._lv.Location = Point(15, 38); self._lv.Size = Size(610, 270)
        self._lv.View = View.Details
        self._lv.FullRowSelect = True; self._lv.GridLines = True
        self._lv.MultiSelect = False; self._lv.Font = CTRL_FONT
        for text, w in (("#", 35), ("Temperature", 110), ("Serie", 170),
                        ("R ratio", 80), ("Walker", 80), ("Points", 70)):
            self._lv.Columns.Add(text, w)
        self._lv.DoubleClick += self._on_edit
        self.Controls.Add(self._lv)

        btn_add = _blue_btn("+ Add", 90, 30)
        btn_add.Location = Point(15, 318); btn_add.Click += self._on_add
        self.Controls.Add(btn_add)

        btn_edit = _grey_btn("Edit", 80, 30)
        btn_edit.Location = Point(112, 318); btn_edit.Click += self._on_edit
        self.Controls.Add(btn_edit)

        btn_rem = _red_btn("Remove", 80, 30)
        btn_rem.Location = Point(199, 318); btn_rem.Click += self._on_remove
        self.Controls.Add(btn_rem)

        btn_create = _blue_btn("Create Excel", 130, 34)
        btn_create.Location = Point(365, 355); btn_create.Click += self._on_create
        self.Controls.Add(btn_create)

        btn_close = _grey_btn("Close", 90, 34)
        btn_close.Location = Point(505, 355)
        btn_close.DialogResult = DialogResult.Cancel
        self.Controls.Add(btn_close)
        self.CancelButton = btn_close

        self._refresh()

    def _refresh(self):
        self._lv.Items.Clear()
        for i, ds in enumerate(self._datasets):
            it = ListViewItem(str(i + 1))
            it.SubItems.Add(str(ds.temp))
            it.SubItems.Add(str(ds.serie))
            it.SubItems.Add("{0:g}".format(ds.r_ratio))
            it.SubItems.Add("{0:g}".format(ds.walker))
            it.SubItems.Add(str(len(ds.rows)))
            self._lv.Items.Add(it)

    def _selected_index(self):
        if self._lv.SelectedIndices.Count == 0:
            return -1
        return self._lv.SelectedIndices[0]

    def _on_add(self, s, e):
        dlg = DataSetDialog()
        if dlg.ShowDialog() == DialogResult.OK and dlg.result is not None:
            self._datasets.append(dlg.result)
            self._refresh()

    def _on_edit(self, s, e):
        i = self._selected_index()
        if i < 0:
            return
        dlg = DataSetDialog(self._datasets[i])
        if dlg.ShowDialog() == DialogResult.OK and dlg.result is not None:
            self._datasets[i] = dlg.result
            self._refresh()

    def _on_remove(self, s, e):
        i = self._selected_index()
        if i < 0:
            return
        r = MessageBox.Show("Remove dataset #{0}?".format(i + 1), "Confirm",
                            MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        if r == DialogResult.Yes:
            del self._datasets[i]
            self._refresh()

    def _on_create(self, s, e):
        if not self._datasets:
            MessageBox.Show("Add at least one dataset first.", "Info",
                            MessageBoxButtons.OK, MessageBoxIcon.Information)
            return
        try:
            Cursor.Current = Cursors.WaitCursor
            res = build_excel(self._datasets)
            Cursor.Current = Cursors.Default
        except Exception as ex:
            Cursor.Current = Cursors.Default
            MessageBox.Show("Excel error:\n" + str(ex), "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
            return

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


def main():
    frm = StrainLifeForm()
    frm.ShowDialog()


if not os.environ.get("STRAIN_LIFE_NO_RUN"):
    main()
