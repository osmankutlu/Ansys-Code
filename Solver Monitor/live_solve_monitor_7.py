# -*- coding: utf-8 -*-
"""
live_solve_monitor_5.py  -  Ansys Mechanical (IronPython 2.7)

Live solve monitor window (non-modal):
    top-left     : Displacement Convergence
    top-right    : Force Convergence
    bottom-left  : tabs -> Solver Output | Substeps & Timing | Load Steps
    bottom-right : Time (s) vs Cumulative Iteration
Charts show vertical markers like Mechanical:
    green dashed : Substep converged
    blue  dashed : Load step converged
    red   dashed : Bisection occurred
Data is read incrementally from solve.out while the solver is running.

Wall-clock times (per substep / load step) are measured by this monitor
when it sees the lines appear in solve.out, so they are only available for
events that happen while the monitor is running (history that already
existed in the file when the monitor attached shows "n/a").

Usage: run the script from Mechanical (Automation > Scripting / Run Script,
or a toolbar button) before or during the solve. Re-running it closes the
old window and opens a new one.
"""

import clr
import re
import os
import time
import math

clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms.DataVisualization")

from System import AppDomain, Array, Byte, Int32
from System.Globalization import CultureInfo
from System.IO import File, FileStream, FileMode, FileAccess, FileShare, SeekOrigin
from System.Text import Encoding
from System.Drawing import (Color, Font, FontStyle, Size, ContentAlignment, SolidBrush, Pen,
                            StringFormat, StringFormatFlags, StringAlignment, StringTrimming, Rectangle,
                            SystemColors)
from System.Reflection import BindingFlags
from System.Windows.Forms import (Form, SplitContainer, DockStyle, Orientation, TextBox,
                                  ScrollBars, FlowLayoutPanel, Label, ComboBox,
                                  ComboBoxStyle, Button, CheckBox, Timer, OpenFileDialog,
                                  SaveFileDialog, DialogResult, FormStartPosition,
                                  TabControl, TabPage, ListView, ListViewItem, View,
                                  HorizontalAlignment)
from System.Windows.Forms.DataVisualization.Charting import (Chart, ChartArea, Series,
                                                             Legend, Docking,
                                                             SeriesChartType, MarkerStyle,
                                                             StripLine, ChartDashStyle,
                                                             LabelMarkStyle, GridTickTypes,
                                                             LabelAutoFitStyles)

# ----------------------------------------------------------------------------
# Settings
# ----------------------------------------------------------------------------
POLL_MS = 1000                      # file polling interval [ms]
MAX_BYTES_PER_TICK = 4 * 1024 * 1024
MAX_OUT_LINES = 1500                # lines kept in the Solver Output pane
SOLVING_TIMEOUT_S = 5.0             # no file growth for this long -> "idle"
MAX_MARKERS = 1500                  # max vertical marker lines per chart
AXIS_TICKS = 11                     # number of y-axis labels (like Mechanical)
LABEL_LO = 0.1                      # |value| below this -> scientific notation
LABEL_HI_FORCE = 1e5                # |value| from this on -> scientific notation (force)
LABEL_HI_DISP = 1e4                 # same for displacement
AXIS_FONT_PT = 10.0                 # axis number / title font size (displacement, force)
GRID_COLOR = Color.DarkGray         # dashed grid lines
FORM_KEY = "LiveSolveMonitor_Form"

# marker kind -> (color, line width)
MARKERS = {
    "sub": (Color.Lime, 1),
    "ls": (Color.RoyalBlue, 2),
    "bis": (Color.Red, 1),
}

SUB_COLS = [("LS", 40, True), ("Substep", 65, True), ("Result", 120, False),
            ("Cumulative Iteration", 135, True), ("Iters", 50, True),
            ("Sim Time", 90, True), ("Time Inc", 90, True),
            ("Substep Time", 100, True), ("Cumulative Time", 110, True),
            ("Est. time", 85, True)]
LS_COLS = [("Load Step", 70, True), ("Substeps", 70, True), ("Bisections", 80, True),
           ("Iterations", 80, True), ("End Sim Time", 100, True),
           ("Wall Time", 90, True), ("Est. time", 85, True), ("Status", 90, False)]

NUM = r'[-+]?(?:\d+\.?\d*|\.\d+)(?:[EeDd][-+]?\d+)?'
RE_ITER = re.compile(r'EQUIL\s+ITER\s+(\d+)\s+COMPLETED')
RE_DISP = re.compile(r'DISP\w*\s+CONVERGENCE\s+VALUE\s*=\s*(%s)\s*CRITERION\s*=\s*(%s)' % (NUM, NUM))
RE_FORCE = re.compile(r'FORCE\s+CONVERGENCE\s+VALUE\s*=\s*(%s)\s*CRITERION\s*=\s*(%s)' % (NUM, NUM))
RE_SUB = re.compile(r'LOAD\s+STEP\s+(\d+)\s+SUBSTEP\s+(\d+)\s+(NOT\s+)?COMPLETED\.\s+CUM(?:ULATIVE)?\s+ITER\s*=\s*(\d+)')
RE_TIME = re.compile(r'\*\*\*\s+TIME\s*=\s*(%s)\s+TIME\s+INC\s*=\s*(%s)' % (NUM, NUM))
RE_BIS = re.compile(r'BEGIN\s+BISECTION\s+NUMBER\s+(\d+)\s+NEW\s+TIME\s+INCREMENT\s*=\s*(%s)' % NUM)
RE_LSNUM = re.compile(r'LOAD\s+STEP\s+NUMBER[\s.]*(\d+)')
RE_CPT = re.compile(r'CP\s+Time\s+\(sec\)\s*=\s*(%s)' % NUM)
RE_ELT = re.compile(r'Elapsed\s+Time\s+\(sec\)\s*=\s*(%s)' % NUM)
RE_ENDCLK = re.compile(r'Time\s*=\s*(\d{1,2}:\d{2}:\d{2})')
RE_DATE = re.compile(r'Date\s*=\s*(\S+)')


def to_float(s):
    return float(s.upper().replace("D", "E"))


def fmt_dur(sec):
    """seconds -> mm:ss (h:mm:ss above one hour); None -> n/a"""
    if sec is None:
        return "n/a"
    sec = int(round(sec))
    h = sec // 3600
    m = (sec % 3600) // 60
    s = sec % 60
    if h:
        return "%d:%02d:%02d" % (h, m, s)
    return "%02d:%02d" % (m, s)


def fmt_hms(sec):
    """seconds -> '2 h 28 min 56 s'"""
    sec = int(round(sec))
    return "%d h %02d min %02d s" % (sec // 3600, (sec % 3600) // 60, sec % 60)


def _dec_sep():
    try:
        sep = CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator
        if sep is not None and len(sep) == 1:
            return sep
    except Exception:
        pass
    return "."


def fmt_val(v, hi):
    """Mechanical-like axis label: plain number with 5 significant digits when
    LABEL_LO <= |v| < hi (125,65 / 8281,3 / 67230), otherwise scientific
    notation (1,9156e+5 / 5e-2)."""
    a = abs(v)
    if a == 0:
        return "0"
    s = None
    if LABEL_LO <= a < hi:
        s = "%.5g" % v
        if "e" in s:                    # rounding pushed it over the limit
            s = None
    if s is None:
        m, e = ("%.4e" % v).split("e")
        if "." in m:
            m = m.rstrip("0").rstrip(".")
        s = "%se%+d" % (m, int(e))
    return s.replace(".", _dec_sep())


def make_ticks(lo, up, log, n=AXIS_TICKS):
    """n ticks spread evenly (geometrically on a log axis) from lo to up."""
    if n < 2 or up <= lo:
        return [lo]
    ticks = []
    if log:
        r = (up / lo) ** (1.0 / (n - 1))
        for i in range(n):
            ticks.append(lo * (r ** i))
    else:
        step = (up - lo) / (n - 1)
        for i in range(n):
            ticks.append(lo + step * i)
    ticks[-1] = up
    return ticks


# ----------------------------------------------------------------------------
# Chart / list helpers
# ----------------------------------------------------------------------------
def _make_series(name, color):
    s = Series(name)
    s.ChartType = SeriesChartType.Line
    s.ChartArea = "ca"
    s.Legend = "lg"
    s.Color = color
    s.BorderWidth = 2
    s.MarkerStyle = MarkerStyle.Circle
    s.MarkerSize = 5
    return s


def _legend_series(name, color, width):
    """Empty series that only exists to put a marker entry in the legend."""
    s = Series(name)
    s.ChartType = SeriesChartType.Line
    s.ChartArea = "ca"
    s.Legend = "lg"
    s.Color = color
    s.BorderWidth = width
    s.BorderDashStyle = ChartDashStyle.Dash
    return s


def make_chart(title, ytitle):
    ch = Chart()
    ch.Dock = DockStyle.Fill
    ch.BackColor = Color.White

    ca = ChartArea("ca")
    ca.AxisX.Title = "Cumulative Iteration"
    ca.AxisY.Title = ytitle
    ca.AxisX.LabelStyle.Format = "0"
    ca.AxisX.MajorGrid.LineColor = GRID_COLOR
    ca.AxisX.MajorGrid.LineDashStyle = ChartDashStyle.Dash
    # Y axis: always a linear axis. In log mode the monitor plots log10(value) and
    # writes the real values as custom labels; horizontal dashed grid lines are
    # strip lines placed at the same positions (see _refresh_axes).
    ca.AxisY.MajorGrid.Enabled = False
    ca.AxisX.IsStartedFromZero = False
    ca.AxisY.IsStartedFromZero = False
    for ax_ in (ca.AxisX, ca.AxisY):            # bigger, non-shrinking axis numbers
        ax_.LabelStyle.Font = Font("Segoe UI", AXIS_FONT_PT)
        ax_.IsLabelAutoFit = False
        ax_.LabelAutoFitStyle = getattr(LabelAutoFitStyles, "None")
        ax_.TitleFont = Font("Segoe UI", AXIS_FONT_PT, FontStyle.Bold)
    ca.CursorX.IsUserSelectionEnabled = True     # drag to zoom
    ca.CursorY.IsUserSelectionEnabled = True
    ch.ChartAreas.Add(ca)

    lg = Legend("lg")
    lg.Docking = Docking.Top
    ch.Legends.Add(lg)

    t = ch.Titles.Add(title)
    t.Font = Font("Segoe UI", 10.0, FontStyle.Bold)

    s_val = _make_series("Convergence", Color.Purple)
    s_crit = _make_series("Criterion", Color.DeepSkyBlue)
    ch.Series.Add(s_val)
    ch.Series.Add(s_crit)

    # marker legend entries: hidden until that kind of event really occurs
    leg = {}
    for (kind, name) in (("sub", "Substep Converged"), ("ls", "Load Step Converged"),
                         ("bis", "Bisection Occurred")):
        ls_ = _legend_series(name, MARKERS[kind][0], MARKERS[kind][1])
        ls_.IsVisibleInLegend = False
        ch.Series.Add(ls_)
        leg[kind] = ls_

    def _reset_zoom(sender, args):          # double click = reset zoom
        ca.AxisX.ScaleView.ZoomReset(0)
        ca.AxisY.ScaleView.ZoomReset(0)
    ch.DoubleClick += _reset_zoom
    return ch, ca, s_val, s_crit, leg


def make_time_chart():
    """Simulation time vs cumulative iteration."""
    ch = Chart()
    ch.Dock = DockStyle.Fill
    ch.BackColor = Color.White

    ca = ChartArea("ca")
    ca.AxisX.Title = "Cumulative Iteration"
    ca.AxisY.Title = "Time (s)"
    ca.AxisX.LabelStyle.Format = "0"
    ca.AxisY.LabelStyle.Format = "0.###"
    for ax_ in (ca.AxisX, ca.AxisY):
        ax_.MajorGrid.LineColor = GRID_COLOR
        ax_.MajorGrid.LineDashStyle = ChartDashStyle.Dash
    ca.AxisX.IsStartedFromZero = False
    ca.AxisY.Minimum = 0.0
    ca.CursorX.IsUserSelectionEnabled = True
    ca.CursorY.IsUserSelectionEnabled = True
    ch.ChartAreas.Add(ca)

    t = ch.Titles.Add("Time vs Cumulative Iteration")
    t.Font = Font("Segoe UI", 10.0, FontStyle.Bold)

    s = Series("Time")
    s.ChartType = SeriesChartType.Line
    s.ChartArea = "ca"
    s.Color = Color.Red
    s.BorderWidth = 2
    s.MarkerStyle = MarkerStyle.Circle
    s.MarkerSize = 4
    s.ToolTip = "iter #VALX{0}, time = #VALY{G6} s"
    ch.Series.Add(s)

    def _reset_zoom(sender, args):
        ca.AxisX.ScaleView.ZoomReset(0)
        ca.AxisY.ScaleView.ZoomReset(0)
    ch.DoubleClick += _reset_zoom
    return ch, ca, s


def make_listview(cols):
    """Owner-drawn ListView: every cell centered, bold column headers."""
    lv = ListView()
    lv.Dock = DockStyle.Fill
    lv.View = View.Details
    lv.FullRowSelect = True
    lv.GridLines = True
    lv.HideSelection = False
    lv.OwnerDraw = True
    for (text, w, right) in cols:
        lv.Columns.Add(text, w, HorizontalAlignment.Center)
    try:                                        # avoid flicker
        pi = lv.GetType().GetProperty("DoubleBuffered",
                                      BindingFlags.Instance | BindingFlags.NonPublic)
        pi.SetValue(lv, True, None)
    except Exception:
        pass

    fmt = StringFormat()
    fmt.Alignment = StringAlignment.Center
    fmt.LineAlignment = StringAlignment.Center
    fmt.Trimming = StringTrimming.EllipsisCharacter
    fmt.FormatFlags = fmt.FormatFlags | StringFormatFlags.NoWrap
    hdr_font = Font(lv.Font, FontStyle.Bold)
    hdr_back = SolidBrush(Color.FromArgb(240, 240, 240))
    line_pen = Pen(Color.FromArgb(200, 200, 200))
    txt_brush = SolidBrush(Color.Black)

    def _draw_header(sender, e):
        g = e.Graphics
        b = e.Bounds
        g.FillRectangle(hdr_back, b)
        g.DrawLine(line_pen, b.Right - 1, b.Top, b.Right - 1, b.Bottom - 1)
        g.DrawLine(line_pen, b.Left, b.Bottom - 1, b.Right, b.Bottom - 1)
        g.DrawString(e.Header.Text, hdr_font, txt_brush,
                     Rectangle(b.X + 2, b.Y, b.Width - 4, b.Height), fmt)

    def _draw_item(sender, e):
        pass                                    # cells are drawn in DrawSubItem

    def _draw_sub(sender, e):
        g = e.Graphics
        b = e.Bounds
        if e.ColumnIndex == 0:                  # bounds of column 0 can span the row
            b = Rectangle(b.X, b.Y, sender.Columns[0].Width, b.Height)
        sel = e.Item.Selected
        if sel:
            g.FillRectangle(SystemBrushes_Highlight(), b)
            col = SystemColors.HighlightText
        else:
            g.FillRectangle(SolidBrush(e.Item.BackColor), b)
            col = e.Item.ForeColor
        g.DrawLine(line_pen, b.Right - 1, b.Top, b.Right - 1, b.Bottom - 1)
        g.DrawLine(line_pen, b.Left, b.Bottom - 1, b.Right, b.Bottom - 1)
        br = SolidBrush(col)
        try:
            g.DrawString(e.SubItem.Text, sender.Font, br,
                         Rectangle(b.X + 2, b.Y, b.Width - 4, b.Height), fmt)
        finally:
            br.Dispose()

    lv.DrawColumnHeader += _draw_header
    lv.DrawItem += _draw_item
    lv.DrawSubItem += _draw_sub
    return lv


def SystemBrushes_Highlight():
    return SolidBrush(SystemColors.Highlight)


# ----------------------------------------------------------------------------
# Mechanical helpers
# ----------------------------------------------------------------------------
def get_analyses():
    res = []
    try:
        for a in ExtAPI.DataModel.Project.Model.Analyses:
            res.append(a)
    except Exception:
        pass
    return res


def get_workdir(a):
    try:
        wd = a.WorkingDir
        if wd:
            return str(wd)
    except Exception:
        pass
    return None


def default_analysis_index(analyses):
    """Analysis that owns the currently selected tree object (else 0)."""
    try:
        o = ExtAPI.DataModel.Tree.ActiveObjects[0]
        for _ in range(8):
            oid = getattr(o, "ObjectId", None)
            for i, a in enumerate(analyses):
                if oid is not None and a.ObjectId == oid:
                    return i
            o = o.Parent
    except Exception:
        pass
    return 0


# ----------------------------------------------------------------------------
# Monitor
# ----------------------------------------------------------------------------
class SolveMonitor(object):

    def __init__(self):
        self.path = None
        self.ctime = None
        self.pos = 0
        self.partial = ""
        self.fresh = False
        self.waited = False
        self.last_growth = 0.0
        self.analyses = []
        self.workdirs = []
        self._init_state()
        self._build_ui()
        self._fill_analyses()

        self.timer = Timer()
        self.timer.Interval = POLL_MS
        self.timer.Tick += self._on_tick
        self.timer.Start()

    def _init_state(self):
        """Parsed-data state (everything that is rebuilt from solve.out)."""
        self.pos = 0
        self.partial = ""
        self._hist = False
        self.iter_cum = 0
        self.accept = False             # True between 'EQUIL ITER n' and end of that iteration
        self.disp = []                  # (x, value, criterion)
        self.force = []
        self.tail = []
        self.out_count = 0
        # events
        self.cur_ls = None
        self.loadsteps = {}
        self.ls_dirty = set()
        self.last_row = None
        self.last_cum_row = 0
        self.last_ok_x = None
        self.pending_time_ev = None
        self.pending_items = []
        self.n_markers = 0
        self.n_ok = 0
        self.n_bis = 0
        # timing
        self.t_prev_row = None
        self.t_prev_ok = None
        self.sum_ok_wall = 0.0
        self.n_ok_timed = 0
        self.last_ok_tot = None
        self.cp_total = None
        self.elapsed_total = None
        self.end_clock = None           # 'hh:mm:ss' printed by the solver at the end
        self.end_date = None
        self.t_start_live = None        # wall-clock start seen live by the monitor
        self.rows = []                  # substep rows (for time estimates)
        self.pend_iters = []            # iterations whose simulation time is not known yet
        self.last_time = 0.0            # sim time at the last converged substep
        self.axes_dirty = False
        self.legend_on = set()          # marker kinds whose legend entry is visible

    # ---------------- UI ----------------
    def _build_ui(self):
        f = Form()
        f.Text = "Live Solve Monitor"
        f.Size = Size(1400, 900)
        f.StartPosition = FormStartPosition.CenterScreen
        self.form = f

        self.main = SplitContainer()
        self.main.Dock = DockStyle.Fill
        self.main.Orientation = Orientation.Horizontal      # top / bottom
        self.top = SplitContainer()
        self.top.Dock = DockStyle.Fill
        self.top.Orientation = Orientation.Vertical         # left / right
        self.main.Panel1.Controls.Add(self.top)

        (self.ch_disp, self.ca_disp, self.sd_val, self.sd_crit, leg_d) = \
            make_chart("Displacement Convergence", "Displacement")
        (self.ch_force, self.ca_force, self.sf_val, self.sf_crit, leg_f) = \
            make_chart("Force Convergence", "Force")
        self.legend_series = dict((k, [leg_d[k], leg_f[k]]) for k in leg_d)
        (self.ch_time, self.ca_time, self.s_time) = make_time_chart()
        self.time_series = [self.s_time]
        self.top.Panel1.Controls.Add(self.ch_disp)          # left
        self.top.Panel2.Controls.Add(self.ch_force)         # right
        self.all_series = [self.sd_val, self.sd_crit, self.sf_val, self.sf_crit]

        # bottom tabs
        ob = TextBox()
        ob.Multiline = True
        ob.ReadOnly = True
        ob.WordWrap = False
        ob.ScrollBars = ScrollBars.Both
        ob.Font = Font("Consolas", 9.0)
        ob.BackColor = Color.White
        ob.Dock = DockStyle.Fill
        ob.MaxLength = Int32.MaxValue
        self.out_box = ob

        self.lv_sub = make_listview(SUB_COLS)
        self.lv_ls = make_listview(LS_COLS)

        tabs = TabControl()
        tabs.Dock = DockStyle.Fill
        for (title, ctrl) in (("Solver Output", ob),
                              ("Substeps & Timing", self.lv_sub),
                              ("Load Steps", self.lv_ls)):
            tp = TabPage(title)
            tp.Controls.Add(ctrl)
            tabs.TabPages.Add(tp)
        self.bottom = SplitContainer()
        self.bottom.Dock = DockStyle.Fill
        self.bottom.Orientation = Orientation.Vertical      # left / right
        self.bottom.Panel1.Controls.Add(tabs)               # bottom-left: output / tables
        self.bottom.Panel2.Controls.Add(self.ch_time)       # bottom-right: time chart
        self.main.Panel2.Controls.Add(self.bottom)

        # toolbar
        bar = FlowLayoutPanel()
        bar.Dock = DockStyle.Top
        bar.Height = 34
        bar.WrapContents = False

        lbl = Label()
        lbl.Text = "Analysis:"
        lbl.AutoSize = False
        lbl.Size = Size(60, 24)
        lbl.TextAlign = ContentAlignment.MiddleLeft
        bar.Controls.Add(lbl)

        cb = ComboBox()
        cb.DropDownStyle = ComboBoxStyle.DropDownList
        cb.Width = 260
        self.combo = cb
        bar.Controls.Add(cb)

        b1 = Button()
        b1.Text = "Browse solve.out..."
        b1.Size = Size(130, 24)
        b1.Click += self._on_browse
        bar.Controls.Add(b1)

        b2 = Button()
        b2.Text = "Reload"
        b2.Size = Size(70, 24)
        b2.Click += self._on_reload
        bar.Controls.Add(b2)

        b3 = Button()
        b3.Text = "Export CSV..."
        b3.Size = Size(95, 24)
        b3.Click += self._on_export
        bar.Controls.Add(b3)

        self.log_chk = CheckBox()
        self.log_chk.Text = "Log scale"
        self.log_chk.Checked = True
        self.log_chk.AutoSize = True
        self.log_chk.CheckedChanged += self._on_log
        bar.Controls.Add(self.log_chk)

        self.follow_chk = CheckBox()
        self.follow_chk.Text = "Follow tables"
        self.follow_chk.Checked = True
        self.follow_chk.AutoSize = True
        bar.Controls.Add(self.follow_chk)

        self.pause_chk = CheckBox()
        self.pause_chk.Text = "Pause"
        self.pause_chk.AutoSize = True
        bar.Controls.Add(self.pause_chk)

        self.top_chk = CheckBox()
        self.top_chk.Text = "Always on top"
        self.top_chk.AutoSize = True
        self.top_chk.CheckedChanged += self._on_topmost
        bar.Controls.Add(self.top_chk)

        self.status = Label()
        self.status.Dock = DockStyle.Bottom
        self.status.Height = 22
        self.status.TextAlign = ContentAlignment.MiddleLeft
        self.status.Font = Font("Segoe UI", 9.0, FontStyle.Bold)

        self.status2 = Label()
        self.status2.Dock = DockStyle.Bottom
        self.status2.Height = 22
        self.status2.TextAlign = ContentAlignment.MiddleLeft
        self.status2.Font = Font("Segoe UI", 9.0, FontStyle.Bold)

        self.status3 = Label()
        self.status3.Dock = DockStyle.Bottom
        self.status3.Height = 22
        self.status3.TextAlign = ContentAlignment.MiddleLeft
        self.status3.Font = Font("Segoe UI", 9.0, FontStyle.Bold)

        self.path_lbl = Label()
        self.path_lbl.Dock = DockStyle.Bottom
        self.path_lbl.Height = 20
        self.path_lbl.TextAlign = ContentAlignment.MiddleLeft

        # Fill first, then edge-docked controls
        f.Controls.Add(self.main)
        f.Controls.Add(self.path_lbl)
        f.Controls.Add(self.status3)
        f.Controls.Add(self.status2)
        f.Controls.Add(self.status)
        f.Controls.Add(bar)

        f.Shown += self._on_shown
        f.FormClosed += self._on_closed

    def _on_shown(self, sender, args):
        try:
            self.main.SplitterDistance = int(self.main.Height * 0.55)
            self.top.SplitterDistance = int(self.top.Width * 0.5)
            self.bottom.SplitterDistance = int(self.bottom.Width * 0.5)
        except Exception:
            pass

    def _on_closed(self, sender, args):
        try:
            self.timer.Stop()
            self.timer.Dispose()
        except Exception:
            pass
        try:
            if AppDomain.CurrentDomain.GetData(FORM_KEY) is self.form:
                AppDomain.CurrentDomain.SetData(FORM_KEY, None)
        except Exception:
            pass

    def _on_topmost(self, sender, args):
        self.form.TopMost = self.top_chk.Checked

    # ---------------- analysis / file selection ----------------
    def _fill_analyses(self):
        self.analyses = get_analyses()
        self.workdirs = [get_workdir(a) for a in self.analyses]
        for i, a in enumerate(self.analyses):
            try:
                nm = str(a.Name)
            except Exception:
                nm = "Analysis"
            self.combo.Items.Add("%d: %s" % (i + 1, nm))
        if self.analyses:
            idx = default_analysis_index(self.analyses)
            self.combo.SelectedIndex = idx
            self._use_analysis(idx)
        else:
            self._set_path(None)
        self.combo.SelectedIndexChanged += self._on_combo

    def _on_combo(self, sender, args):
        idx = self.combo.SelectedIndex
        if idx >= 0:
            self._use_analysis(idx)

    def _use_analysis(self, idx):
        wd = get_workdir(self.analyses[idx])
        self.workdirs[idx] = wd
        self._set_path(os.path.join(wd, "solve.out") if wd else None)

    def _on_browse(self, sender, args):
        dlg = OpenFileDialog()
        dlg.Title = "Select solve.out"
        dlg.Filter = "Solver output (*.out)|*.out|All files (*.*)|*.*"
        try:
            if self.path and os.path.isdir(os.path.dirname(self.path)):
                dlg.InitialDirectory = os.path.dirname(self.path)
        except Exception:
            pass
        if dlg.ShowDialog() == DialogResult.OK:
            self._set_path(dlg.FileName)

    def _on_reload(self, sender, args):
        self._reset_data()
        self.fresh = True           # re-read existing content: no wall-clock times

    def _set_path(self, p):
        self.path = p
        self.ctime = None
        self.path_lbl.Text = "  File: %s" % (p if p else "(none - select an analysis or Browse)")
        self._reset_data()
        self.fresh = True           # existing content is history

    # ---------------- reset ----------------
    def _reset_data(self):
        self._init_state()
        self.fresh = False
        for s in self.all_series + self.time_series:
            s.Points.Clear()
        for ca in (self.ca_disp, self.ca_force):
            ca.AxisX.StripLines.Clear()
            ax = ca.AxisY
            ax.CustomLabels.Clear()
            ax.StripLines.Clear()
            ax.Minimum = float("nan")
            ax.Maximum = float("nan")
        for lst in self.legend_series.values():
            for s_ in lst:
                s_.IsVisibleInLegend = False
        self.out_box.Clear()
        self.lv_sub.Items.Clear()
        self.lv_ls.Items.Clear()

    # ---------------- charts ----------------
    def _pv(self, v):
        # log axis cannot show values <= 0
        if self.log_chk.Checked and v <= 0.0:
            return 1e-12
        return v

    def _ty(self, v):
        """value -> plotted y coordinate (log10 in log mode)."""
        v = self._pv(v)
        if self.log_chk.Checked:
            return math.log10(v)
        return v

    def _add_pt(self, ser, x, raw):
        idx = ser.Points.AddXY(float(x), float(self._ty(raw)))
        try:
            ser.Points[idx].ToolTip = "%s: iter %d, value %s" % (ser.Name, int(x), "%.5g" % raw)
        except Exception:
            pass

    def _plot(self, s_val, s_crit, x, v, c):
        self._add_pt(s_val, x, v)
        self._add_pt(s_crit, x, c)
        self.axes_dirty = True

    def _on_log(self, sender, args):
        for s in self.all_series:
            s.Points.Clear()
        for ca in (self.ca_disp, self.ca_force):
            ca.AxisY.CustomLabels.Clear()
            ca.AxisY.StripLines.Clear()
            ca.AxisY.Minimum = float("nan")
            ca.AxisY.Maximum = float("nan")
        for (x, v, c) in self.disp:
            self._plot(self.sd_val, self.sd_crit, x, v, c)
        for (x, v, c) in self.force:
            self._plot(self.sf_val, self.sf_crit, x, v, c)
        self._refresh_axes()

    def _refresh_axes(self):
        """Mechanical-style y axis: ~11 labels spread evenly between the data minimum and
        maximum (plain number or scientific notation, see fmt_val) with horizontal dashed
        grid lines at the same positions."""
        self.axes_dirty = False
        log = self.log_chk.Checked
        for (data, ca, hi) in ((self.disp, self.ca_disp, LABEL_HI_DISP),
                               (self.force, self.ca_force, LABEL_HI_FORCE)):
            if not data:
                continue
            vals = []
            for (x, v, c) in data:
                vals.append(self._pv(v))
                vals.append(self._pv(c))
            lo = min(vals)
            up = max(vals)
            if up <= lo:
                up = lo * 1.1 if lo > 0 else lo + 1.0
            ax = ca.AxisY
            try:
                ylo = self._ty(lo)
                yup = self._ty(up)
                pad = (yup - ylo) * 0.04
                ax.Minimum = ylo - pad
                ax.Maximum = yup + pad
                ax.MajorGrid.Enabled = False
                ax.MajorTickMark.Enabled = False
                ax.CustomLabels.Clear()
                ax.StripLines.Clear()
                w = (yup - ylo) * 0.002
                for t in make_ticks(lo, up, log):
                    pos = self._ty(t)
                    ax.CustomLabels.Add(pos - w, pos + w, fmt_val(t, hi), 0,
                                        getattr(LabelMarkStyle, "None"), getattr(GridTickTypes, "None"))
                    sl = StripLine()
                    sl.IntervalOffset = pos
                    sl.StripWidth = 0.0
                    sl.BorderColor = GRID_COLOR
                    sl.BorderWidth = 1
                    sl.BorderDashStyle = ChartDashStyle.Dash
                    ax.StripLines.Add(sl)
            except Exception:
                pass

    def _plot_time(self, xs, t):
        for x in xs:
            self.s_time.Points.AddXY(float(x), float(t))

    def _add_marker(self, x, kind, tip):
        if self.n_markers >= MAX_MARKERS:
            return
        if kind not in self.legend_on:              # first event of this kind -> show legend entry
            self.legend_on.add(kind)
            for s_ in self.legend_series[kind]:
                s_.IsVisibleInLegend = True
        color, width = MARKERS[kind]
        for ca in (self.ca_disp, self.ca_force):
            sl = StripLine()
            sl.IntervalOffset = float(x)
            sl.StripWidth = 0.0
            sl.BorderColor = color
            sl.BorderWidth = width
            sl.BorderDashStyle = ChartDashStyle.Dash
            try:
                sl.ToolTip = tip
            except Exception:
                pass
            ca.AxisX.StripLines.Add(sl)
        self.n_markers += 1

    # ---------------- tables ----------------
    def _new_item(self, cells, red=False):
        it = ListViewItem(cells[0])
        for c in cells[1:]:
            it.SubItems.Add(c)
        if red:
            it.ForeColor = Color.Red
        return it

    def _ls_get(self, ls):
        L = self.loadsteps.get(ls)
        if L is None:
            it = self._new_item([str(ls)] + [""] * (len(LS_COLS) - 1))
            L = {"ls": ls, "nsub": 0, "nbis": 0, "iters": 0, "wall": 0.0, "timed": 0,
                 "partial": False, "sim_end": None, "done": False, "est": "", "item": it}
            self.loadsteps[ls] = L
            self.lv_ls.Items.Add(it)
            self.ls_dirty.add(ls)
        return L

    def _refresh_ls_row(self, L):
        if L["timed"] == 0:
            wall = "n/a"
        else:
            wall = fmt_dur(L["wall"]) + ("*" if L["partial"] else "")
        sim = ("%.6g" % L["sim_end"]) if L["sim_end"] is not None else ""
        vals = [str(L["ls"]), str(L["nsub"]), str(L["nbis"]), str(L["iters"]), sim, wall,
                L["est"], "done" if L["done"] else "running"]
        for i, v in enumerate(vals):
            L["item"].SubItems[i].Text = v

    def _flush_tables(self):
        if self.pending_items:
            self.lv_sub.BeginUpdate()
            try:
                self.lv_sub.Items.AddRange(Array[ListViewItem](self.pending_items))
                if self.follow_chk.Checked:
                    self.pending_items[-1].EnsureVisible()
            finally:
                self.lv_sub.EndUpdate()
            self.pending_items = []
        for ls in list(self.ls_dirty):
            L = self.loadsteps.get(ls)
            if L is not None:
                self._refresh_ls_row(L)
        self.ls_dirty = set()

    def _on_export(self, sender, args):
        dlg = SaveFileDialog()
        dlg.Title = "Export tables"
        dlg.Filter = "CSV (*.csv)|*.csv"
        dlg.FileName = "solve_monitor.csv"
        if dlg.ShowDialog() != DialogResult.OK:
            return
        try:
            f = open(dlg.FileName, "w")
            try:
                f.write("Substeps\n")
                f.write(";".join([c[0] for c in SUB_COLS]) + "\n")
                for it in self.lv_sub.Items:
                    f.write(";".join([s.Text for s in it.SubItems]) + "\n")
                f.write("\nLoad Steps\n")
                f.write(";".join([c[0] for c in LS_COLS]) + "\n")
                for it in self.lv_ls.Items:
                    f.write(";".join([s.Text for s in it.SubItems]) + "\n")
            finally:
                f.close()
            self.status2.Text = "  Exported: %s" % dlg.FileName
        except Exception as ex:
            self.status2.Text = "  Export error: %s" % ex

    # ---------------- parsing / events ----------------
    def _cum_time(self, t):
        """Wall time since the start of the (live) solve, or None if unknown."""
        if t is None or self.t_start_live is None:
            return None
        return t - self.t_start_live

    def _stamp_row(self, ok):
        """Wall-clock bookkeeping for one table row -> (t, attempt, substep_total)."""
        if self._hist:
            return None, None, None
        t = time.time()
        dur = (t - self.t_prev_row) if self.t_prev_row is not None else None
        tot = None
        if ok:
            if self.t_prev_ok is not None:
                tot = t - self.t_prev_ok
            self.t_prev_ok = t
        self.t_prev_row = t
        return t, dur, tot

    @staticmethod
    def _acc_wall(L, dur):
        if dur is None:
            L["partial"] = True
        else:
            L["wall"] += dur
            L["timed"] += 1

    def _set_ls(self, ls):
        if self.cur_ls is not None and ls != self.cur_ls:
            self._end_loadstep(self.cur_ls)
        self.cur_ls = ls

    def _end_loadstep(self, ls):
        L = self.loadsteps.get(ls)
        if L is None or L["done"] or self.last_ok_x is None:
            return
        L["done"] = True
        self._add_marker(self.last_ok_x, "ls",
                         "Load step %d converged (cum iter %d)" % (ls, self.last_ok_x))
        self.ls_dirty.add(ls)

    def _on_substep(self, ls, sub, notc, cum):
        ok = not notc
        self._set_ls(ls)
        L = self._ls_get(ls)
        t, dur, tot = self._stamp_row(ok)
        iters = cum - self.last_cum_row
        self.last_cum_row = cum
        # Substep Time: how long this substep took (incl. cutback attempts);
        # Cumulative Time: running total since the start of the solve
        sub_t = tot if ok else dur
        cells = [str(ls), str(sub), "Converged" if ok else "NOT converged", str(cum),
                 str(iters), "", "", fmt_dur(sub_t), fmt_dur(self._cum_time(t)), ""]
        item = self._new_item(cells, red=not ok)
        ev = {"ok": ok, "item": item, "bis": None, "ls": ls, "iters": iters}
        self.pending_items.append(item)
        self.rows.append(ev)
        self.last_row = ev
        L["iters"] += iters
        self._acc_wall(L, dur)
        if ok:
            L["nsub"] += 1
            self.n_ok += 1
            self.last_ok_x = cum
            if tot is not None:
                self.sum_ok_wall += tot
                self.n_ok_timed += 1
                self.last_ok_tot = tot
            self._add_marker(cum, "sub", "LS %d substep %d converged (cum iter %d)" % (ls, sub, cum))
            self.pending_time_ev = ev
        else:
            self.pending_time_ev = None
        self.ls_dirty.add(ls)
        self.accept = False

    def _on_time(self, simt, dt):
        ev = self.pending_time_ev
        if ev is None:
            return
        ev["item"].SubItems[5].Text = "%.6g" % simt
        ev["item"].SubItems[6].Text = "%.6g" % dt
        # all iterations of this substep ran at its (end) time
        self._plot_time(self.pend_iters, simt)
        self.pend_iters = []
        self.last_time = simt
        L = self._ls_get(ev["ls"])
        L["sim_end"] = simt
        self.ls_dirty.add(ev["ls"])
        self.pending_time_ev = None

    def _on_bisection(self, n, newdt):
        ls = self.cur_ls if self.cur_ls is not None else 1
        L = self._ls_get(ls)
        L["nbis"] += 1
        self.n_bis += 1
        self.ls_dirty.add(ls)
        self._add_marker(self.iter_cum, "bis",
                         "Bisection #%d (cum iter %d), new time inc = %.6g" % (n, self.iter_cum, newdt))
        # the failed attempt ran at last time + old increment (= 2 x new one for the default 0.5 cutback)
        self._plot_time(self.pend_iters, self.last_time + 2.0 * newdt)
        self.pend_iters = []
        last = self.last_row
        if last is not None and not last["ok"] and last["bis"] is None:
            # the NOT COMPLETED row that triggered this bisection
            last["bis"] = n
            last["item"].SubItems[2].Text = "Bisection #%d" % n
            last["item"].SubItems[6].Text = "%.6g" % newdt
        else:
            t, dur, tot = self._stamp_row(False)
            cells = [str(ls), "", "Bisection #%d" % n, "", "", "", "%.6g" % newdt,
                     fmt_dur(dur), fmt_dur(self._cum_time(t)), ""]
            item = self._new_item(cells, red=True)
            self.pending_items.append(item)
            self._acc_wall(L, dur)
            self.last_row = {"ok": False, "item": item, "bis": n, "ls": ls}
        self.pending_time_ev = None
        self.accept = False

    def _on_run_completed(self):
        if self.cur_ls is not None and self.last_row is not None and self.last_row["ok"]:
            self._end_loadstep(self.cur_ls)

    def _apply_estimates(self):
        """Once the solver reports its total elapsed time, split it over rows and
        load steps in proportion to the iterations (estimate, marked with ~)."""
        total = self.last_cum_row
        if not self.elapsed_total or total <= 0:
            return
        for ev in self.rows:
            ev["item"].SubItems[9].Text = "~" + fmt_dur(self.elapsed_total * ev["iters"] / float(total))
        for ls, L in self.loadsteps.items():
            L["est"] = "~" + fmt_dur(self.elapsed_total * L["iters"] / float(total))
            self.ls_dirty.add(ls)

    def _feed(self, lines):
        for ln in lines:
            if "EQUIL ITER" in ln:
                if RE_ITER.search(ln):
                    self.iter_cum += 1
                    self.pend_iters.append(self.iter_cum)
                    self.accept = True      # convergence values below belong to this iteration
                continue
            if "CONVERGENCE VALUE" in ln:
                # Lines printed before the first 'EQUIL ITER' of a substep are a
                # pre-check that Mechanical does not plot -> ignore them.
                if not self.accept:
                    continue
                m = RE_DISP.search(ln)
                if m:
                    v, c = to_float(m.group(1)), to_float(m.group(2))
                    self.disp.append((self.iter_cum, v, c))
                    self._plot(self.sd_val, self.sd_crit, self.iter_cum, v, c)
                    continue
                m = RE_FORCE.search(ln)
                if m:
                    v, c = to_float(m.group(1)), to_float(m.group(2))
                    self.force.append((self.iter_cum, v, c))
                    self._plot(self.sf_val, self.sf_crit, self.iter_cum, v, c)
                continue
            if "SUBSTEP" in ln and "COMPLETED" in ln:
                m = RE_SUB.search(ln)
                if m:
                    self._on_substep(int(m.group(1)), int(m.group(2)),
                                     m.group(3) is not None, int(m.group(4)))
                continue
            if "BISECTION" in ln:
                m = RE_BIS.search(ln)
                if m:
                    self._on_bisection(int(m.group(1)), to_float(m.group(2)))
                continue
            if "TIME INC" in ln:
                m = RE_TIME.search(ln)
                if m:
                    self._on_time(to_float(m.group(1)), to_float(m.group(2)))
                continue
            if "LOAD STEP NUMBER" in ln:
                m = RE_LSNUM.search(ln)
                if m:
                    self._set_ls(int(m.group(1)))
                continue
            if "SOLUTION CONVERGED" in ln:
                self.accept = False
                continue
            if "RUN COMPLETED" in ln:
                self._on_run_completed()
                continue
            if "(sec)" in ln:
                m = RE_CPT.search(ln)
                if m:
                    self.cp_total = to_float(m.group(1))
                m2 = RE_ENDCLK.search(ln)
                if m2 and "CP" in ln:
                    self.end_clock = m2.group(1)
                m = RE_ELT.search(ln)
                if m:
                    self.elapsed_total = to_float(m.group(1))
                    m3 = RE_DATE.search(ln)
                    if m3:
                        d = m3.group(1).split("/")      # solver prints mm/dd/yyyy
                        self.end_date = ("%s.%s.%s" % (d[1], d[0], d[2])) if len(d) == 3 else m3.group(1)
                    self._apply_estimates()
        self._append_output(lines)

    def _append_output(self, lines):
        if not lines:
            return
        lines = lines[-MAX_OUT_LINES:]
        self.tail.extend(lines)
        self.out_count += len(lines)
        if self.out_count > MAX_OUT_LINES * 1.5:
            self.tail = self.tail[-MAX_OUT_LINES:]
            self.out_box.Text = "\r\n".join(self.tail) + "\r\n"
            self.out_count = len(self.tail)
        else:
            self.out_box.AppendText("\r\n".join(lines) + "\r\n")
        self.out_box.SelectionStart = self.out_box.TextLength
        self.out_box.ScrollToCaret()

    # ---------------- polling ----------------
    def _on_tick(self, sender, args):
        if self.pause_chk.Checked:
            return
        try:
            self._poll()
        except Exception as ex:
            self.status.Text = "  Error: %s" % ex

    def _poll(self):
        p = self.path
        if not p or not File.Exists(p):
            # working dir may change once the solve starts - re-query it
            idx = self.combo.SelectedIndex
            if idx >= 0 and idx < len(self.analyses):
                wd = get_workdir(self.analyses[idx])
                if wd:
                    newp = os.path.join(wd, "solve.out")
                    if newp != p:
                        self._set_path(newp)
                        p = newp
            if not p or not File.Exists(p):
                self.fresh = False      # file will appear with the new solve: it is live
                self.waited = True
                self.status.Text = "  Waiting for solve.out ..."
                return

        if self.waited:                 # file appeared after we started waiting -> live
            self.waited = False
            self.fresh = False
            self.t_prev_row = time.time()
            self.t_prev_ok = self.t_prev_row
            self.t_start_live = self.t_prev_row

        ct = File.GetCreationTimeUtc(p)
        text = ""
        hist = self.fresh
        fs = FileStream(p, FileMode.Open, FileAccess.Read,
                        FileShare.ReadWrite | FileShare.Delete)
        try:
            length = fs.Length
            if (self.ctime is not None and ct != self.ctime) or length < self.pos:
                # a new solve overwrote the file: start over, times are live
                self._reset_data()
                hist = False
                self.t_prev_row = time.time()
                self.t_prev_ok = self.t_prev_row
                self.t_start_live = self.t_prev_row
            self.ctime = ct
            todo = int(min(length - self.pos, MAX_BYTES_PER_TICK))
            if todo > 0:
                fs.Seek(self.pos, SeekOrigin.Begin)
                buf = Array.CreateInstance(Byte, todo)
                n = fs.Read(buf, 0, todo)
                self.pos += n
                text = Encoding.ASCII.GetString(buf, 0, n)
            if length > 0 and self.pos >= length:
                self.fresh = False      # caught up with the file
        finally:
            fs.Close()

        if text:
            self.last_growth = time.time()
            text = self.partial + text
            lines = text.split("\n")
            self.partial = lines.pop()      # incomplete last line
            lines = [l.rstrip("\r") for l in lines]
            self._hist = hist
            self._feed(lines)

        self._flush_tables()
        if self.axes_dirty:
            self._refresh_axes()
        self._update_status()

    def _update_status(self):
        now = time.time()
        solving = (now - self.last_growth) < SOLVING_TIMEOUT_S
        parts = ["Iter: %d" % self.iter_cum]
        if self.disp:
            x, v, c = self.disp[-1]
            parts.append("Disp %.4g / %.4g %s" % (v, c, "OK" if v <= c else "NOT conv."))
        if self.force:
            x, v, c = self.force[-1]
            parts.append("Force %.4g / %.4g %s" % (v, c, "OK" if v <= c else "NOT conv."))
        parts.append("SOLVING" if solving else "idle")
        self.status.Text = "  " + "   |   ".join(parts)

        p2 = []
        if self.cur_ls is not None:
            L = self.loadsteps.get(self.cur_ls)
            p2.append("LS %d: %d substeps" % (self.cur_ls, L["nsub"] if L else 0))
        p2.append("Total substeps: %d" % self.n_ok)
        p2.append("Bisections: %d" % self.n_bis)
        if self.last_ok_tot is not None:
            avg = self.sum_ok_wall / self.n_ok_timed
            p2.append("Last substep: %s (avg %s)" % (fmt_dur(self.last_ok_tot), fmt_dur(avg)))
        if solving and self.t_prev_row is not None:
            p2.append("Running since last substep: %s" % fmt_dur(now - self.t_prev_row))
        if self.elapsed_total is not None:
            p2.append("Solver elapsed: %s" % fmt_dur(self.elapsed_total))
        if self.cp_total is not None:
            p2.append("CP: %s" % fmt_dur(self.cp_total))
        self.status2.Text = "  " + "   |   ".join(p2)
        self._update_summary(now, solving)

    def _update_summary(self, now, solving):
        """Start / end clock and total duration of the analysis."""
        start_txt = None
        end_txt = None
        if self.end_clock is not None and self.elapsed_total is not None:
            try:
                h, mi, s = [int(v) for v in self.end_clock.split(":")]
                st = h * 3600 + mi * 60 + s - int(round(self.elapsed_total))
                prev_day = st < 0
                st = st % 86400
                start_txt = "%02d:%02d:%02d" % (st // 3600, (st % 3600) // 60, st % 60)
                if prev_day:
                    start_txt += " (previous day)"
                elif self.end_date:
                    start_txt = "%s %s" % (self.end_date, start_txt)
                end_txt = ("%s %s" % (self.end_date, self.end_clock)) if self.end_date else self.end_clock
            except Exception:
                start_txt = None
        if start_txt is None and self.t_start_live is not None:
            start_txt = time.strftime("%d.%m.%Y %H:%M:%S", time.localtime(self.t_start_live))
        parts = []
        if start_txt:
            parts.append("Start: %s" % start_txt)
        if end_txt:
            parts.append("End: %s" % end_txt)
        if self.elapsed_total is not None:
            parts.append("Total: %s" % fmt_hms(self.elapsed_total))
        elif solving and self.t_start_live is not None:
            parts.append("Running: %s" % fmt_hms(now - self.t_start_live))
        self.status3.Text = ("  " + "   |   ".join(parts)) if parts else ""


# ----------------------------------------------------------------------------
# Entry point
# ----------------------------------------------------------------------------
try:
    _old = AppDomain.CurrentDomain.GetData(FORM_KEY)
    if _old is not None:
        _old.Close()
except Exception:
    pass

_mon = SolveMonitor()
AppDomain.CurrentDomain.SetData(FORM_KEY, _mon.form)
_mon.form.Show()
