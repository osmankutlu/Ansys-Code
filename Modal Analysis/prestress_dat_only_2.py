# -*- coding: utf-8 -*-
# ============================================================
# Prestress Modal DAT Generator - ANSYS Mechanical
# Static Structural + Modal -> merged file.dat
# + Optional Specified Harmonic Index (cycopt,hindex command + Cyclic Controls)
# ============================================================

import os
import clr
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

from System.Windows.Forms import (
    Form, Button, Label, CheckBox,
    DialogResult, FormStartPosition,
    MessageBox, MessageBoxButtons, MessageBoxIcon,
    FormBorderStyle, OpenFileDialog, Application,
    TextBox, ListBox, SelectionMode
)
from System.Drawing import Size, Point, Color, Font, FontStyle

# =====================================================================
# STYLE CONSTANTS
# =====================================================================
FORM_W    = 520
LBL_H     = 20
ROW_H     = 28
CTRL_FONT = Font("Segoe UI", 9)
HDR_FONT  = Font("Segoe UI", 9, FontStyle.Bold)
BTN_BLUE  = Color.FromArgb(0, 120, 215)

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


# =====================================================================
# HARMONIC INDEX PICKER - split panel (1..200)
# =====================================================================
class HarmonicIndexPickerDialog(Form):
    def __init__(self, preselected=None):
        self.selected_indices = list(preselected) if preselected else []

        self.Text            = "Select Harmonic Indices"
        self.StartPosition   = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox     = False
        self.MinimizeBox     = False
        self.ClientSize      = Size(500, 420)
        self.BackColor       = Color.White

        lbl_left = Label()
        lbl_left.Text     = "Available Indices (0-200)"
        lbl_left.Location = Point(15, 12)
        lbl_left.Size     = Size(180, 20)
        lbl_left.Font     = Font("Segoe UI", 9, FontStyle.Bold)
        self.Controls.Add(lbl_left)

        self.lst_avail = ListBox()
        self.lst_avail.Location      = Point(15, 35)
        self.lst_avail.Size          = Size(180, 320)
        self.lst_avail.SelectionMode = SelectionMode.MultiExtended
        self.Controls.Add(self.lst_avail)

        btn_add = Button()
        btn_add.Text      = "Add >"
        btn_add.Location  = Point(210, 150)
        btn_add.Size      = Size(80, 32)
        btn_add.ForeColor = Color.FromArgb(0, 100, 200)
        btn_add.Click    += self._on_add
        self.Controls.Add(btn_add)

        btn_del = Button()
        btn_del.Text      = "< Remove"
        btn_del.Location  = Point(210, 192)
        btn_del.Size      = Size(80, 32)
        btn_del.ForeColor = Color.FromArgb(200, 0, 0)
        btn_del.Click    += self._on_del
        self.Controls.Add(btn_del)

        lbl_right = Label()
        lbl_right.Text     = "Selected Indices"
        lbl_right.Location = Point(305, 12)
        lbl_right.Size     = Size(180, 20)
        lbl_right.Font     = Font("Segoe UI", 9, FontStyle.Bold)
        self.Controls.Add(lbl_right)

        self.lst_sel = ListBox()
        self.lst_sel.Location      = Point(305, 35)
        self.lst_sel.Size          = Size(180, 320)
        self.lst_sel.SelectionMode = SelectionMode.MultiExtended
        self.Controls.Add(self.lst_sel)

        btn_ok = _blue_btn("OK", 80, 30)
        btn_ok.Location = Point(305, 368)
        btn_ok.Click   += self._on_ok
        self.Controls.Add(btn_ok)

        btn_cancel = _grey_btn("Cancel", 90, 30)
        btn_cancel.Location     = Point(395, 368)
        btn_cancel.DialogResult = DialogResult.Cancel
        self.Controls.Add(btn_cancel)
        self.CancelButton = btn_cancel

        self._populate()

    def _populate(self):
        sel_set = set(self.selected_indices)
        self.lst_avail.Items.Clear()
        self.lst_sel.Items.Clear()
        for i in range(0, 201):
            if i not in sel_set:
                self.lst_avail.Items.Add(str(i))
        for i in sorted(sel_set):
            self.lst_sel.Items.Add(str(i))

    def _on_add(self, sender, args):
        to_add = [int(str(self.lst_avail.Items[i]))
                  for i in range(self.lst_avail.Items.Count)
                  if self.lst_avail.GetSelected(i)]
        for v in to_add:
            if v not in self.selected_indices:
                self.selected_indices.append(v)
        self.selected_indices.sort()
        self._populate()

    def _on_del(self, sender, args):
        to_del = set(int(str(self.lst_sel.Items[i]))
                     for i in range(self.lst_sel.Items.Count)
                     if self.lst_sel.GetSelected(i))
        self.selected_indices = [v for v in self.selected_indices if v not in to_del]
        self._populate()

    def _on_ok(self, sender, args):
        if not self.selected_indices:
            MessageBox.Show("Please select at least one harmonic index.", "Warning")
            return
        self.DialogResult = DialogResult.OK
        self.Close()


# =====================================================================
# PROGRESS DIALOG
# =====================================================================
class ProgressDialog(Form):
    def __init__(self):
        self.cancelled = False
        self.Text            = "Processing..."
        self.StartPosition   = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedSingle
        self.MaximizeBox     = False; self.MinimizeBox = True
        self.ClientSize      = Size(440, 110); self.BackColor = Color.White
        self.ControlBox      = True; self.TopMost = True

        self._lbl_title = Label()
        self._lbl_title.Text     = "Please wait..."
        self._lbl_title.Location = Point(15, 16)
        self._lbl_title.Size     = Size(410, 18)
        self._lbl_title.Font     = HDR_FONT
        self.Controls.Add(self._lbl_title)

        self._lbl_detail = Label()
        self._lbl_detail.Text      = ""
        self._lbl_detail.Location  = Point(15, 40)
        self._lbl_detail.Size      = Size(410, 18)
        self._lbl_detail.Font      = CTRL_FONT
        self._lbl_detail.ForeColor = Color.FromArgb(80, 80, 80)
        self.Controls.Add(self._lbl_detail)

        btn_cancel = _grey_btn("Cancel", 90, 28)
        btn_cancel.Location = Point(335, 72)
        btn_cancel.Click   += self._on_cancel
        self.Controls.Add(btn_cancel)
        self.FormClosing += self._on_closing

    def _on_closing(self, s, e):
        if not self.cancelled:
            e.Cancel = True
            self._on_cancel(s, e)

    def _on_cancel(self, s, e):
        self.cancelled = True
        self._lbl_title.Text  = "Cancelling..."
        self._lbl_detail.Text = "Waiting for current operation to finish."

    def update(self, title, detail=""):
        self._lbl_title.Text  = title
        self._lbl_detail.Text = detail
        self.Refresh()
        Application.DoEvents()

    def close_safe(self):
        try:
            self.cancelled = True
            self.Close()
        except: pass


# =====================================================================
# ANA FORM - Harmonic Index secimi + Output Folder
# =====================================================================
class MainForm(Form):
    def __init__(self, static_name, modal_name):
        self.out_folder       = None
        self.harmonic_indices = []

        self.Text            = "Prestress Modal DAT Generator"
        self.StartPosition   = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox     = False; self.MinimizeBox = False
        self.BackColor       = Color.White

        y = 14

        lbl_title = Label()
        lbl_title.Text      = "Prestress Modal Analysis"
        lbl_title.Location  = Point(15, y)
        lbl_title.Size      = Size(FORM_W - 30, 22)
        lbl_title.Font      = Font("Segoe UI", 11, FontStyle.Bold)
        lbl_title.ForeColor = Color.FromArgb(0, 80, 160)
        self.Controls.Add(lbl_title)
        y += 38

        # -- Specified Harmonic Index --
        self.chk_harmonic = CheckBox()
        self.chk_harmonic.Text     = "Specified Harmonic Index"
        self.chk_harmonic.Location = Point(15, y)
        self.chk_harmonic.Size     = Size(220, 22)
        self.chk_harmonic.Font     = HDR_FONT
        self.chk_harmonic.CheckedChanged += self._on_harmonic_toggle
        self.Controls.Add(self.chk_harmonic)

        self.btn_harmonic = _grey_btn("Select Indices...", 140, 24)
        self.btn_harmonic.Location = Point(245, y - 1)
        self.btn_harmonic.Enabled  = False
        self.btn_harmonic.Click   += self._on_select_harmonic
        self.Controls.Add(self.btn_harmonic)

        self.lbl_harmonic_sel = Label()
        self.lbl_harmonic_sel.Text      = "(none selected)"
        self.lbl_harmonic_sel.Location  = Point(15, y + 26)
        self.lbl_harmonic_sel.Size      = Size(FORM_W - 30, 20)
        self.lbl_harmonic_sel.Font      = CTRL_FONT
        self.lbl_harmonic_sel.ForeColor = Color.Gray
        self.Controls.Add(self.lbl_harmonic_sel)

        y += 26 + 26

        # -- Output Folder --
        lbl_out = Label()
        lbl_out.Text     = "Output Folder:"
        lbl_out.Location = Point(15, y)
        lbl_out.Size     = Size(FORM_W - 30, LBL_H)
        lbl_out.Font     = HDR_FONT
        self.Controls.Add(lbl_out)
        y += LBL_H + 2

        self.txt_folder = TextBox()
        self.txt_folder.Location  = Point(15, y)
        self.txt_folder.Size      = Size(FORM_W - 60, 24)
        self.txt_folder.Font      = CTRL_FONT
        self.txt_folder.ReadOnly  = True
        self.txt_folder.BackColor = Color.White
        self.Controls.Add(self.txt_folder)

        btn_browse = _grey_btn("...", 34, 24)
        btn_browse.Location = Point(FORM_W - 40, y)
        btn_browse.Click   += self._on_browse
        self.Controls.Add(btn_browse)
        y += ROW_H + 14

        # -- Run / Cancel --
        btn_run = _blue_btn("Run  ->", 100, 32)
        btn_run.Location = Point(FORM_W - 208, y)
        btn_run.Click   += self._on_run
        self.Controls.Add(btn_run)

        btn_cancel = _grey_btn("Cancel", 90, 32)
        btn_cancel.Location     = Point(FORM_W - 100, y)
        btn_cancel.DialogResult = DialogResult.Cancel
        self.Controls.Add(btn_cancel)
        self.CancelButton = btn_cancel

        y += 46
        self.ClientSize = Size(FORM_W, y)

    def _on_harmonic_toggle(self, s, e):
        enabled = self.chk_harmonic.Checked
        self.btn_harmonic.Enabled = enabled
        if not enabled:
            self.harmonic_indices = []
            self.lbl_harmonic_sel.Text = "(none selected)"
            self.lbl_harmonic_sel.ForeColor = Color.Gray

    def _on_select_harmonic(self, s, e):
        dlg = HarmonicIndexPickerDialog(preselected=self.harmonic_indices)
        if dlg.ShowDialog() == DialogResult.OK:
            self.harmonic_indices = dlg.selected_indices
            nums = ", ".join(str(v) for v in self.harmonic_indices)
            if len(nums) > 60:
                nums = nums[:57] + "..."
            self.lbl_harmonic_sel.Text = nums
            self.lbl_harmonic_sel.ForeColor = Color.Black

    def _on_browse(self, s, e):
        dlg = OpenFileDialog()
        dlg.Title           = "Select output folder  -  open any file inside it"
        dlg.Filter          = "All Files (*.*)|*.*"
        dlg.CheckFileExists = False
        dlg.FileName        = "Select folder"
        if dlg.ShowDialog() == DialogResult.OK:
            self.txt_folder.Text = os.path.dirname(dlg.FileName)

    def _on_run(self, s, e):
        folder = self.txt_folder.Text.strip()
        if not folder:
            MessageBox.Show("Please select an output folder.", "Missing Input",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        if self.chk_harmonic.Checked and not self.harmonic_indices:
            MessageBox.Show("Please select at least one harmonic index, or uncheck "
                            "'Specified Harmonic Index'.", "Missing Input",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        try:
            if not os.path.isdir(folder):
                os.makedirs(folder)
        except Exception as ex:
            MessageBox.Show("Cannot create folder:\n" + str(ex),
                            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            return

        if not self.chk_harmonic.Checked:
            self.harmonic_indices = []

        self.out_folder    = folder
        self.DialogResult  = DialogResult.OK
        self.Close()


# =====================================================================
# ANALIZ BULMA
# =====================================================================
def _find_static(analyses):
    for a in analyses:
        try:
            atype = str(a.AnalysisType).replace(" ","").replace("_","").lower()
            if "static" in atype: return a
        except: pass
    raise Exception("No Static Structural analysis found.")

def _find_modal(analyses):
    for a in analyses:
        try:
            atype = str(a.AnalysisType).replace(" ","").replace("_","").lower()
            if "modal" in atype: return a
        except: pass
    raise Exception("No Modal analysis found.")


# =====================================================================
# HARMONIC INDEX - Commands objesi + Cyclic Controls API
# =====================================================================
def _add_harmonic_index_commands(modal_analysis, harmonic_indices):
    """
    Modal analizine bir Commands (APDL) objesi ekler. Icerigi:
      cycopt,hindex,4,4,1
      cycopt,hindex,6,6,1
      cycopt,hindex,9,9,1
      ...
    Her secilen harmonic index icin bir satir.
    """
    lines = []
    for hi in sorted(harmonic_indices):
        lines.append("cycopt,hindex,{0},{0},1".format(hi))
    cmd_text = "\n".join(lines)

    cmd_obj = modal_analysis.AddCommandSnippet()
    cmd_obj.Input = cmd_text
    try:
        cmd_obj.Name = "Harmonic Index Selection"
    except: pass
    return cmd_obj


def _set_cyclic_harmonic_index_range(modal_analysis, min_index):
    """
    Modal analizinin Analysis Settings -> Cyclic Controls -> Harmonic Index
    Range -> Manual -> Minimum/Maximum degerlerini API uzerinden ayarlar.
    Hem Minimum hem Maximum 'min_index' (en kucuk secilen harmonic index)
    olarak girilir.

    Surume gore property adlari degisebildigi icin birden fazla olasi
    isim/yol denenir.
    """
    settings = modal_analysis.AnalysisSettings

    # HarmonicIndexRange enum: Program Controlled / Manual.
    # Manual genelde enum degeri 1'dir; calismazsa string ile de denenir.
    range_set = False
    for manual_val in [1, "Manual"]:
        try:
            settings.HarmonicIndexRange = manual_val
            range_set = True
            break
        except: pass

    # Once Maximum, sonra Minimum ayarla (bazi surumlerde Max once
    # ayarlanirsa Min'i otomatik kisitlayip yanlis deger biraktigi
    # gorulmustur; once buyuk degeri yazip sonra Min'i kesin olarak
    # tekrar yazmak en guvenilir sira).
    max_set = False
    try:
        settings.MaximumHarmonicIndex = int(min_index)
        max_set = True
    except: pass

    min_set = False
    try:
        settings.MinimumHarmonicIndex = int(min_index)
        min_set = True
    except: pass

    # Emniyet: Minimum'u son bir kez daha dogrula/yaz (bazi surumlerde
    # Maximum atamasi Minimum'u geri degistirebiliyor)
    try:
        if int(settings.MinimumHarmonicIndex) != int(min_index):
            settings.MinimumHarmonicIndex = int(min_index)
            min_set = True
    except: pass

    return range_set, min_set, max_set


# =====================================================================
# BIRLESTIRME
# =====================================================================
def _merge_dat_files(static_dat, modal_dat, output_dat, insert_token, prog):
    prog.update("Reading Static DAT...")
    with open(static_dat, "r") as f:
        static_lines = f.readlines()

    prog.update("Reading Modal DAT...")
    with open(modal_dat, "r") as f:
        modal_lines = f.readlines()

    if not static_lines:
        raise Exception("Static DAT is empty.")

    merged = list(static_lines[:-1])
    merged.append(insert_token + "\n")
    merged.append(static_lines[-1])
    if modal_lines and not static_lines[-1].endswith("\n"):
        merged.append("\n")
    merged.extend(modal_lines)

    prog.update("Writing file.dat...", output_dat)
    with open(output_dat, "w") as f:
        f.writelines(merged)


# =====================================================================
# SURE FORMATI
# =====================================================================
def _format_duration(seconds):
    seconds = int(seconds)
    if seconds < 60:
        return "{0} sec".format(seconds)
    mins = seconds // 60
    secs = seconds % 60
    if mins < 60:
        return "{0} min {1} sec".format(mins, secs)
    hrs  = mins // 60
    mins = mins % 60
    return "{0} hr {1} min {2} sec".format(hrs, mins, secs)


# =====================================================================
# ENTRY POINT
# =====================================================================
def run():
    import datetime

    model    = ExtAPI.DataModel.Project.Model
    analyses = list(model.Analyses)

    if not analyses:
        MessageBox.Show("No analyses found!", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error)
        return

    try:
        static_a = _find_static(analyses)
        modal_a  = _find_modal(analyses)
    except Exception as ex:
        MessageBox.Show(str(ex), "Analysis Not Found",
                        MessageBoxButtons.OK, MessageBoxIcon.Error)
        return

    frm = MainForm(static_a.Name, modal_a.Name)
    if frm.ShowDialog() != DialogResult.OK:
        return

    out_folder        = frm.out_folder
    harmonic_indices  = frm.harmonic_indices

    # Specified Harmonic Index secildiyse: Commands objesi ekle ve
    # Cyclic Controls Harmonic Index Range Min/Max'i API ile ayarla
    if harmonic_indices:
        try:
            _add_harmonic_index_commands(modal_a, harmonic_indices)
        except Exception as ex:
            MessageBox.Show(
                "Could not add the Harmonic Index command snippet to the "
                "Modal analysis.\n\nError: {0}".format(str(ex)),
                "Harmonic Index Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            return

        min_hindex = min(harmonic_indices)
        range_set, min_set, max_set = _set_cyclic_harmonic_index_range(modal_a, min_hindex)
        if not (min_set and max_set):
            MessageBox.Show(
                "The Harmonic Index command was added successfully, but the "
                "Cyclic Controls -> Harmonic Index Range (Manual) Minimum/Maximum "
                "fields could not be set automatically via the API for this Ansys "
                "version.\n\nPlease set them manually in:\n"
                "Modal -> Analysis Settings -> Cyclic Controls -> Harmonic Index Range "
                "-> Manual -> Minimum = {0}, Maximum = {0}".format(min_hindex),
                "Cyclic Controls", MessageBoxButtons.OK, MessageBoxIcon.Warning)

    static_tmp = os.path.join(out_folder, "static_input.dat")
    modal_tmp  = os.path.join(out_folder, "modal_input.dat")
    merged_dat = os.path.join(out_folder, "file.dat")

    prog = ProgressDialog()
    prog.Show()
    Application.DoEvents()

    start_time = datetime.datetime.now()

    try:
        prog.update("Writing Static DAT...", static_tmp)
        static_a.WriteInputFile(static_tmp)
        if not os.path.exists(static_tmp):
            raise Exception("Static DAT not created:\n" + static_tmp)
        if prog.cancelled: prog.close_safe(); return

        prog.update("Writing Modal DAT...", modal_tmp)
        modal_a.WriteInputFile(modal_tmp)
        if not os.path.exists(modal_tmp):
            raise Exception("Modal DAT not created:\n" + modal_tmp)
        if prog.cancelled: prog.close_safe(); return

        _merge_dat_files(static_tmp, modal_tmp, merged_dat,
                         "/COPY,file,rst,,static,rst", prog)

        try: os.remove(static_tmp)
        except: pass
        try: os.remove(modal_tmp)
        except: pass

        end_time = datetime.datetime.now()
        elapsed  = (end_time - start_time).total_seconds()
        duration = _format_duration(elapsed)

        prog.close_safe()

        harmonic_str = ", ".join(str(v) for v in sorted(harmonic_indices)) if harmonic_indices else "(not specified)"

        MessageBox.Show(
            "Done!\n\n"
            "file.dat        ->  {fd}\n"
            "Harmonic Index  :  {hi}\n\n"
            "Duration  :  {dur}".format(
                fd=merged_dat,
                hi=harmonic_str,
                dur=duration),
            "Complete", MessageBoxButtons.OK, MessageBoxIcon.Information)

    except Exception as ex:
        prog.close_safe()
        MessageBox.Show("An error occurred:\n\n" + str(ex),
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)


run()
