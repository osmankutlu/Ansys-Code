# -*- coding: utf-8 -*-
# ============================================================
# Strain-Life Tool - ANSYS Mechanical (IronPython 2.7)
# v6: per-node cycle from an RST file + contour on the model
#
# Flow (wizard windows):
#   1) Walker window     : Serie + any number of Temperature / Walker pairs (Add / Remove)
#   2) Data window(s)    : one per temperature -> R ratio + pasted Total Strain / Cycle
#   3) Run window        : Analysis, RST file, Named Selection, cycle upper limit
#   Then (Named Selection -> nodes, same resolver as the Cut BC tool):
#     * Equivalent Total Strain per node over the LOAD STEPS (time 0 excluded,
#       last set of every load step, no substeps)  -> max strain (B5) / min strain (B6)
#     * max temperature per node (structural BFE temperature in the RST)
#     * for EVERY given temperature: node cycle with that temperature's Walker,
#       R ratio and power fit  (B3,B7,B10,B11,B12 -> Cycle = A*|B12|^aaa)
#     * node cycle = LINEAR interpolation (or extrapolation with the nearest two)
#       between those cycles at the node's max temperature
#     * infinite life (B12 = 0 / result above the limit) -> upper limit (editable)
#     * result is the cycle value itself (not log) and is shown on the model as a
#       DPF Python Result (same mechanism as the averaging error tool)
#   Results window shows the fit summary + node table; "Export to Excel" is optional.
#
# Strain operator in DPF: found by probing a list of candidates (see log).
# ============================================================

import os
import math
import math as _math
import clr
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

from System.Windows.Forms import (
    Form, Button, Label, TextBox, ScrollBars, ComboBox, ComboBoxStyle,
    ListView, ListViewItem, View, ColumnHeader, OpenFileDialog,
    DataGridView, DataGridViewTextBoxColumn, DataGridViewAutoSizeColumnsMode,
    DialogResult, FormStartPosition, MessageBox, MessageBoxButtons,
    MessageBoxIcon, FormBorderStyle, Cursor, Cursors, Application
)
from System.Drawing import Size, Point, Color, Font, FontStyle
from System import Int32
from System.Collections.Generic import List as _NetList

try:
    import mech_dpf
    import Ans.DataProcessing as dpf
except Exception:          # allows the pure logic to be tested outside Mechanical
    mech_dpf = None
    dpf = None

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
NODE_SHEET   = "Node Results"
DEFAULT_CAP  = 1e10
GRID_MIN_ROWS = 14
MAX_TABLE_ROWS = 2000  # rows shown in the results window (Excel gets all)


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

def _lbl(text, x, y, w, bold=True):
    l = Label()
    l.Text = text; l.Location = Point(x, y); l.Size = Size(w, LBL_H)
    l.Font = HDR_FONT if bold else CTRL_FONT
    return l



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


def trim_num(v, decimals=15):
    """Fixed notation with up to `decimals` decimals, no trailing zeros."""
    s = ("%." + str(decimals) + "f") % v
    if "." in s:
        s = s.rstrip("0").rstrip(".")
    if s in ("", "-0"):
        s = "0"
    return s


def num_fmt_auto(v):
    """Excel number format without trailing zeros ('0' for whole numbers)."""
    try:
        if abs(v - round(v)) < 1e-12:
            return "0"
    except Exception:
        pass
    return "0.###############"


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


# =====================================================================
# RST / DPF BACKEND  (from the Cut BC tool)
# =====================================================================
def contains_non_ascii(name):
    """
    Named Selection isminde ASCII disi karakter (Turkce ı,ğ,ü,ş,ö,ç vb. dahil)
    var mi kontrol eder. .NET string <-> Python str gecislerinde bu karakterler
    "ascii codec can't decode" hatasina yol acabildigi icin, export baslamadan
    once kullaniciya acik bir uyari vermek amaciyla kullanilir.
    """
    if not name:
        return False
    try:
        for ch in name:
            if ord(ch) > 127:
                return True
        return False
    except:
        return True

def safe_str(net_obj):
    """
    .NET String'i Python str'e GUVENLE cevirir.
    Duz str(net_obj) cagrisi, icinde Turkce gibi ASCII-disi karakter olan
    .NET String'lerde IronPython'da "'ascii' codec can't decode byte..." hatasi
    firlatip programi unhandled-exception ekranina dusurebilir.
    Bu fonksiyon ASLA exception firlatmaz; basarisiz olursa, en azindan
    ASCII-disi karakterleri '?' ile degistirilmis bir metin dondurur, boylece
    kod akisi devam edebilir ve contains_non_ascii() ile sorunu tespit edip
    kullaniciya temiz bir uyari gosterebiliriz.
    """
    if net_obj is None:
        return ""
    try:
        return str(net_obj)
    except:
        pass
    try:
        u = unicode(net_obj)
        out = []
        for ch in u:
            out.append(ch if ord(ch) <= 127 else "?")
        return "".join(out)
    except:
        return "<unreadable name>"

def attach_ns_name_warning(combo):
    """
    Combo'da bir Named Selection secildigi anda ismini kontrol eder.
    ASCII-disi (Turkce vb.) karakter icerirse aninda temiz bir Ingilizce
    uyari gosterir ve secimi "-- Select --"'a geri alir, boylece bozuk
    isimli bir NS ile export adimina hic gecilmez.
    """
    def _handler(sender, args):
        try:
            if sender.SelectedIndex <= 0:
                return
            name = safe_str(sender.SelectedItem)
            if contains_non_ascii(name):
                MessageBox.Show(
                    "The Named Selection '{0}' contains non-English characters "
                    "(e.g. Turkish letters). Please rename it in Mechanical using "
                    "only standard English letters/numbers, then select it again.".format(name),
                    "Invalid Named Selection Name",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning)
                sender.SelectedIndex = 0
        except:
            # Bu kontrolun kendisi asla ana akisi bozmamali
            pass
    combo.SelectedIndexChanged += _handler

def _to_net_int_list(py_ids):
    lst = _NetList[Int32]()
    for i in py_ids:
        lst.Add(int(i))
    return lst

# =======================================================================
# DPF BACKEND

def is_thermal_file(path):
    """True for a thermal result file (.rth)."""
    try:
        return os.path.splitext(str(path))[1].lower() == ".rth"
    except:
        return False

def dpf_load(rst_path):
    ds = dpf.DataSources(rst_path)
    model = dpf.Model(ds)
    return ds, model

def get_load_step_info(dpf_model):
    """
    RST'deki load step'leri ve her load step'in son result set ID'sini dondurur.
    Donen: [(load_step_no, last_set_id, time_val), ...]

    Substep yoksa: [(1, 1, t1), (2, 2, t2), ...]
    Substep varsa: [(1, 5, t5), (2, 10, t10), ...] -- her load step'in SON set'i
    Bu Mechanical'in davranisiyla eslesir (Step N = load step N'nin son substep'i).
    """
    tf = dpf_model.TimeFreqSupport
    n_sets = int(tf.NumberSets)

    try:
        time_vals = list(tf.TimeFreqs.Data)
    except:
        time_vals = [0.0] * n_sets

    # Yontem 1: StepIds field'i (her setin hangi load step'e ait oldugunu verir)
    step_ids = None
    try:
        sfield = tf.StepIds
        step_ids = [int(x) for x in sfield.Data]
    except:
        pass

    if step_ids and len(step_ids) == n_sets:
        # Load step bazinda grupla; her load step icin SON seti al
        seen = {}
        for i in range(n_sets):
            ls = step_ids[i]
            set_id = i + 1
            t = float(time_vals[i]) if i < len(time_vals) else 0.0
            seen[ls] = (set_id, t)  # ayni load step'te sonraki override eder
        return sorted([(ls, sid, t) for ls, (sid, t) in seen.items()])

    # Yontem 2: NumberSetsPerStep field'i
    try:
        nps = [int(x) for x in tf.NumberSetsPerStep.Data]
        result = []
        cumul = 0
        for ls_idx, count in enumerate(nps):
            cumul += count
            last_sid = cumul
            t = float(time_vals[last_sid - 1]) if last_sid - 1 < len(time_vals) else 0.0
            result.append((ls_idx + 1, last_sid, t))
        return result
    except:
        pass

    # Fallback: her set = ayri load step (substep yok veya okunamadi)
    return [(i + 1, i + 1, float(time_vals[i]) if i < len(time_vals) else 0.0)
            for i in range(n_sets)]

def get_num_sets(dpf_model):
    """Geriye donuk uyumluluk icin."""
    try:
        return int(dpf_model.TimeFreqSupport.NumberSets)
    except:
        return 1

def fc_count(fc):
    try:
        return fc.FieldCount
    except:
        try:
            return len(fc)
        except:
            return 1

def fc_get_field(fc, i):
    try:
        return fc[i]
    except:
        return fc.GetFieldByIndex(i)

# ------------------- MESH + COORDINATES (v2 fallback chain) -----------
def get_mesh(ds, dpf_model, log):
    mesh = None
    try:
        mesh = dpf_model.Mesh
        log.append("Mesh: model.Mesh OK")
    except Exception as e:
        log.append("Mesh: model.Mesh basarisiz -> {0}".format(str(e)))
    if mesh is None:
        try:
            mop = dpf.operators.mesh.mesh_provider()
            mop.inputs.data_sources.Connect(ds)
            mesh = mop.outputs.mesh.GetData()
            log.append("Mesh: mesh_provider OK")
        except Exception as e:
            log.append("Mesh: mesh_provider basarisiz -> {0}".format(str(e)))
    if mesh is None:
        try:
            mop = dpf.Operator("mapdl::rst::MeshProvider")
            mop.Connect(4, ds)  # pin 4 = data_sources
            mesh = mop.GetOutputAsMeshedRegion(0)
            log.append("Mesh: generic MeshProvider OK")
        except Exception as e:
            log.append("Mesh: generic MeshProvider basarisiz -> {0}".format(str(e)))
    return mesh

def _coords_from_field(cfield, coords):
    ids = list(cfield.ScopingIds)
    data = list(cfield.Data)
    for k in range(len(ids)):
        coords[ids[k]] = (data[k * 3], data[k * 3 + 1], data[k * 3 + 2])

def build_coords_dict(ds, dpf_model, mesh, log):
    """NodeID -> (x, y, z). Donen: (coords_dict, length_unit)"""
    coords = {}
    unit = ""
    if mesh is None:
        log.append("KOORDINAT: Mesh yok, koordinat okunamadi.")
        return coords, unit
    try:
        unit = str(mesh.Unit)
    except:
        unit = ""

    # Yontem 1: Generic operator
    try:
        op = dpf.Operator("mesh::node_coordinates")
        op.Connect(7, mesh)
        cfield = op.GetOutputAsField(0)
        _coords_from_field(cfield, coords)
        try:
            u2 = str(cfield.Unit)
            if u2: unit = u2
        except: pass
        if len(coords) > 0:
            log.append("KOORDINAT: Yontem 1 (generic) OK, {0} node".format(len(coords)))
            return coords, unit
    except Exception as e:
        log.append("KOORDINAT: Yontem 1 basarisiz -> {0}".format(str(e)))

    # Yontem 2: Wrapper operator
    try:
        cop = dpf.operators.mesh.node_coordinates()
        cop.inputs.mesh.Connect(mesh)
        try:
            cfield = cop.outputs.coordinates.GetData()
        except:
            cfield = cop.GetOutputAsField(0)
        _coords_from_field(cfield, coords)
        try:
            u2 = str(cfield.Unit)
            if u2: unit = u2
        except: pass
        if len(coords) > 0:
            log.append("KOORDINAT: Yontem 2 (wrapper) OK, {0} node".format(len(coords)))
            return coords, unit
    except Exception as e:
        log.append("KOORDINAT: Yontem 2 basarisiz -> {0}".format(str(e)))

    # Yontem 3: NodesCoordinatesField
    try:
        cfield = mesh.NodesCoordinatesField
        _coords_from_field(cfield, coords)
        if len(coords) > 0:
            log.append("KOORDINAT: Yontem 3 (NodesCoordinatesField) OK, {0} node".format(len(coords)))
            return coords, unit
    except Exception as e:
        log.append("KOORDINAT: Yontem 3 basarisiz -> {0}".format(str(e)))

    # Yontem 4: mesh.Nodes dongusu
    try:
        cnt = 0
        for n in mesh.Nodes:
            try:
                nid = n.Id
            except:
                continue
            try:
                coords[nid] = (n.X, n.Y, n.Z)
                cnt += 1
            except:
                try:
                    c = n.Coordinates
                    coords[nid] = (c[0], c[1], c[2])
                    cnt += 1
                except:
                    continue
        if cnt > 0:
            log.append("KOORDINAT: Yontem 4 (mesh.Nodes) OK, {0} node".format(cnt))
            return coords, unit
    except Exception as e:
        log.append("KOORDINAT: Yontem 4 basarisiz -> {0}".format(str(e)))

    log.append("KOORDINAT: Hicbir yontem calismadi.")
    return coords, unit

# ------------------- NAMED SELECTION RESOLUTION ------------------------
_ns_cache = {}

def _rst_ns_scoping(ds, ns_name, location):
    """RST icindeki component'ten scoping dener. location: 'Nodal'/'Elemental'."""
    for nm in [ns_name, ns_name.upper()]:
        # Wrapper
        try:
            op = dpf.operators.scoping.on_named_selection()
            op.inputs.requested_location.Connect(location)
            op.inputs.named_selection_name.Connect(nm)
            op.inputs.data_sources.Connect(ds)
            sc = op.outputs.mesh_scoping.GetData()
            ids = list(sc.Ids)
            if len(ids) > 0:
                return ids
        except:
            pass
        # Generic
        try:
            op = dpf.Operator("scoping_provider_by_ns")
            op.Connect(0, location)
            op.Connect(1, nm)
            op.Connect(4, ds)
            sc = op.GetOutputAsScoping(0)
            ids = list(sc.Ids)
            if len(ids) > 0:
                return ids
        except:
            pass
    return None

def _find_model_ns(ns_name):
    """Acik modeldeki NamedSelection objesini ismiyle bulur."""
    try:
        for item in ExtAPI.DataModel.Project.Model.NamedSelections.Children:
            if item.Name == ns_name:
                return item
    except:
        pass
    return None

def _ns_node_ids_from_mesh(ns):
    """
    Herhangi bir NS'i (mesh-node / mesh-element / geometrik edge-face-body)
    acik modelin mesh'i uzerinden NODE id listesine cevirir.
    Sonradan olusturulup henuz rst'ye yazilmamis NS'ler icin de calisir.
    """
    try:
        loc = ns.Location
        st = str(loc.SelectionType)
    except:
        return None

    out = []
    seen = {}

    def _add(i):
        if i not in seen:
            seen[i] = True
            out.append(i)

    md = None
    try:
        md = ExtAPI.DataModel.MeshDataByName("Global")
    except:
        md = None

    # 1) Zaten node bazli NS
    try:
        if "MeshNode" in st:
            for i in loc.Ids:
                _add(i)
            if out:
                return out
    except:
        pass

    # 2) Element bazli NS -> elemanlarin node'lari
    try:
        if "MeshElement" in st and md is not None:
            for eid in loc.Ids:
                try:
                    el = md.ElementById(eid)
                    for nid in el.NodeIds:
                        _add(nid)
                except:
                    continue
            if out:
                return out
    except:
        pass

    # 3) Geometrik NS (edge / face / body) -> mesh region node'lari
    if md is not None:
        try:
            for gid in loc.Ids:
                try:
                    reg = md.MeshRegionById(gid)
                except:
                    continue
                try:
                    for nid in reg.NodeIds:
                        _add(nid)
                except:
                    pass
            if out:
                return out
        except:
            pass

    return out if out else None

def _ui_node_coords_for_ns(ns):
    """
    Geometrik/element/node NS'i acik modelin mesh'i (MeshDataByName) uzerinden
    {ui_node_id: (x, y, z)} sozlugune cevirir.
    Bu koordinatlar UI mesh'inden gelir; rst node ID'lerine eslemek icin kullanilir.
    """
    ui_ids = _ns_node_ids_from_mesh(ns)
    if not ui_ids:
        return None
    try:
        md = ExtAPI.DataModel.MeshDataByName("Global")
    except:
        return None

    coords = {}
    for nid in ui_ids:
        try:
            nd = md.NodeById(nid)
            coords[nid] = (nd.X, nd.Y, nd.Z)
        except:
            continue
    return coords if coords else None

# Koordinat -> rst node ID haritasi (DPF mesh'inden, bir kere kurulur)
_coord_key_map = {"built": False, "grid": {}, "cell": 1.0, "rst_maxabs": 1.0}

def _ui_mean_abs(ui_coords):
    s = 0.0
    n = 0
    for (x, y, z) in ui_coords.values():
        s += abs(x) + abs(y) + abs(z)
        n += 3
        if n >= 600:
            break
    return (s / n) if n else 0.0

def _rst_mean_abs(coords):
    s = 0.0
    n = 0
    for (x, y, z) in coords.values():
        s += abs(x) + abs(y) + abs(z)
        n += 3
        if n >= 600:
            break
    return (s / n) if n else 0.0

def build_coord_to_rstid(coords):
    """
    DPF coords ({rst_nid: (x,y,z)}) uzerinden uzamsal hucre (grid) haritasi kurar.
    Hucre boyutu modelin tipik boyutuna gore secilir, boylece en-yakin-node
    aramasi yapilabilir. Birim farki UI tarafinda olceklenerek giderilir.
    """
    if _coord_key_map["built"]:
        return

    maxabs = 0.0
    for (x, y, z) in coords.values():
        m = max(abs(x), abs(y), abs(z))
        if m > maxabs:
            maxabs = m
    if maxabs <= 0:
        maxabs = 1.0

    # Hucre boyutu: model boyutunun ~1e-4'u (komsu node'lardan kucuk olmali)
    cell = maxabs * 1.0e-4
    if cell <= 0:
        cell = 1.0e-6

    grid = {}
    for rid, (x, y, z) in coords.items():
        gk = (int(_math.floor(x / cell)),
              int(_math.floor(y / cell)),
              int(_math.floor(z / cell)))
        grid.setdefault(gk, []).append((rid, x, y, z))

    _coord_key_map["grid"] = grid
    _coord_key_map["cell"] = cell
    _coord_key_map["rst_maxabs"] = maxabs
    _coord_key_map["built"] = True

def _nearest_rst_id(x, y, z, tol2):
    """Verilen (x,y,z)'ye en yakin rst node ID'sini grid'den bulur.
    tol2 = izin verilen maksimum mesafe^2. Yoksa None."""
    grid = _coord_key_map["grid"]
    cell = _coord_key_map["cell"]
    gx = int(_math.floor(x / cell))
    gy = int(_math.floor(y / cell))
    gz = int(_math.floor(z / cell))
    best_id = None
    best_d2 = None
    for dx in (-1, 0, 1):
        for dy in (-1, 0, 1):
            for dz in (-1, 0, 1):
                bucket = grid.get((gx + dx, gy + dy, gz + dz))
                if not bucket:
                    continue
                for (rid, rx, ry, rz) in bucket:
                    d2 = (rx - x) ** 2 + (ry - y) ** 2 + (rz - z) ** 2
                    if best_d2 is None or d2 < best_d2:
                        best_d2 = d2
                        best_id = rid
    if best_id is not None and best_d2 is not None and best_d2 <= tol2:
        return best_id
    return None

def _detect_ui_scale(ui_coords, coords):
    """
    UI koordinatlari ile DPF coords arasindaki olcek faktorunu tahmin eder.
    Donen: ui_value * scale ~= rst_value olacak sekilde scale.
    mm(UI) vs m(rst) gibi durumlar icin tipik degerler: 1, 0.001, 1000.
    """
    ui_m = _ui_mean_abs(ui_coords)
    rst_m = _rst_mean_abs(coords)
    if ui_m <= 0 or rst_m <= 0:
        return 1.0
    raw = rst_m / ui_m
    # En yakin "makul" olcege oturt
    candidates = [1.0, 0.001, 1000.0, 0.01, 100.0, 0.0254, 1.0/0.0254]
    best = 1.0
    best_err = None
    for c in candidates:
        err = abs(_math.log(raw) - _math.log(c)) if (raw > 0 and c > 0) else 1e9
        if best_err is None or err < best_err:
            best_err = err
            best = c
    # Eger raw zaten 1'e cok yakinsa olcek yok
    if 0.8 < raw < 1.25:
        return 1.0
    return best

def map_ui_ns_to_rst_ids(ns, coords, log, return_detail=False):
    """
    Geometrik (rst'de OLMAYAN) NS'in node'larini UI koordinatlari uzerinden
    DPF mesh'indeki rst node ID'lerine esler (en-yakin-node + otomatik olcek).
    """
    ui_coords = _ui_node_coords_for_ns(ns)
    if not ui_coords:
        if return_detail:
            return None, "UI mesh'inden NS node koordinati alinamadi."
        return None

    build_coord_to_rstid(coords)
    cell = _coord_key_map["cell"]
    scale = _detect_ui_scale(ui_coords, coords)

    # Tolerans: hucre boyutunun yarisi (ayni mesh oldugu icin cok kucuk olmali)
    tol = cell * 0.5
    tol2 = tol * tol

    rst_ids = []
    seen = {}
    miss = 0
    for ui_nid, (x, y, z) in ui_coords.items():
        sx, sy, sz = x * scale, y * scale, z * scale
        rid = _nearest_rst_id(sx, sy, sz, tol2)
        if rid is None:
            miss += 1
            continue
        if rid not in seen:
            seen[rid] = True
            rst_ids.append(rid)

    detail = ("UI node: {0}, eslesen rst node: {1}, eslesmeyen: {2}, "
              "olcek(ui->rst): {3}, hucre: {4:.3e}".format(
                  len(ui_coords), len(rst_ids), miss, scale, cell))
    log.append("NS koordinat eslemesi -> " + detail)

    result = rst_ids if rst_ids else None
    if return_detail:
        return result, detail
    return result

def _ns_elem_ids_from_mesh(ns):
    """NS'i acik modelin mesh'i uzerinden ELEMENT id listesine cevirir."""
    try:
        loc = ns.Location
        st = str(loc.SelectionType)
    except:
        return None

    out = []
    seen = {}

    def _add(i):
        if i not in seen:
            seen[i] = True
            out.append(i)

    # 1) Zaten element bazli NS
    try:
        if "MeshElement" in st:
            for i in loc.Ids:
                _add(i)
            if out:
                return out
    except:
        pass

    # 2) Geometrik NS -> mesh region elementleri
    try:
        md = ExtAPI.DataModel.MeshDataByName("Global")
        for gid in loc.Ids:
            try:
                reg = md.MeshRegionById(gid)
            except:
                continue
            try:
                for eid in reg.ElementIds:
                    _add(eid)
            except:
                pass
        if out:
            return out
    except:
        pass

    return out if out else None

def _model_ns_ids(ns_name, want_elements):
    """Acik Mechanical modelindeki NS'ten node/element ID'leri toplar
    (edge/face/body NS'leri mesh uzerinden donusturerek)."""
    ns = _find_model_ns(ns_name)
    if ns is None:
        return None
    if want_elements:
        return _ns_elem_ids_from_mesh(ns)
    else:
        return _ns_node_ids_from_mesh(ns)

def resolve_ns_ids(ds, ns_name, want_elements, log, coords=None):
    """
    NS -> RST node/element id listesi. Cache'li.

    NODE icin (want_elements=False):
      1) rst component (NS rst'de tanimliysa) -> ID'ler birebir rst uzayinda.
      2) NS rst'de yoksa: UI mesh'inden NS node KOORDINATLARINI al, DPF mesh
         koordinatlariyla eslestirip RST node ID'lerine cevir. (UI node ID'leri
         ile rst node ID'leri farkli olabildigi icin dogrudan ID kullanmiyoruz.)
    ELEMENT icin (want_elements=True):
      1) rst component, 2) UI mesh element ID'leri (genelde rst ile ortustugu
         varsayilir; ENF zaten eleman scoping + node connectivity ile filtrelenir).
    """
    key = (ns_name, want_elements)
    if key in _ns_cache:
        return _ns_cache[key]
    location = "Elemental" if want_elements else "Nodal"

    # 1) rst icindeki component (varsa en kesin)
    ids = _rst_ns_scoping(ds, ns_name, location)
    if ids:
        log.append("NS '{0}' ({1}): RST component'ten {2} id".format(ns_name, location, len(ids)))
        _ns_cache[key] = ids
        return ids

    # 2) NS rst'de yok -> acik modelden coz
    ns = _find_model_ns(ns_name)
    if ns is None:
        log.append("NS '{0}' ({1}): COZULEMEDI (modelde yok)".format(ns_name, location))
        _ns_cache[key] = None
        return None

    if want_elements:
        ids = _ns_elem_ids_from_mesh(ns)
        if ids:
            log.append("NS '{0}' (Elemental): Model mesh'inden {1} eleman".format(ns_name, len(ids)))
            _ns_cache[key] = ids
            return ids
    else:
        # Koordinat eslemesi (UI -> rst). NS rst'de yok, UI ID'leri rst ile
        # ayni olmayabilir; bu yuzden DAIMA koordinattan eslenir.
        if coords:
            ids = map_ui_ns_to_rst_ids(ns, coords, log)
            if ids:
                _ns_cache[key] = ids
                return ids
            log.append("NS '{0}' (Nodal): koordinat eslemesi BASARISIZ "
                       "(birim/olcek uyumsuz olabilir).".format(ns_name))
        else:
            log.append("NS '{0}' (Nodal): coords yok, eslenemez.".format(ns_name))

    log.append("NS '{0}' ({1}): COZULEMEDI".format(ns_name, location))
    _ns_cache[key] = None
    return None

def _is_solver_generated(name):
    """
    Solver'in otomatik olusturdugu NS'leri tespit eder:
    - Alt cizgi ile baslayan (_FIXEDSU, _INTERNAL vb.)
    - Tamamen buyuk harf + rakam, alt cizgi veya bosluk YOK
      (FIXEDSU, CONTACTSF, SYMBC, ABC123 vb.)
    Kullanici NS'leri genelde kucuk harf icerir veya alt cizgi/bosluk iceren
    karisik formattadir (NS_Edge_1, Force_Boundary, deneme_edge vb.).
    """
    s = name.strip()
    if not s:
        return True
    if s.startswith("_"):
        return True
    # Tamamen buyuk harf + rakam, IC BOSLUK veya ALT CIZGI YOK
    # (FIXEDSU, CONTACTSF, SYMBC, ABC123 gibi)
    stripped = s.replace("-", "")
    if stripped == stripped.upper() and stripped.isalnum():
        return True
    return False

def collect_ns_names(ds, dpf_model):
    """
    Combo listesi: oncelikle model agacindaki kullanici NS'leri.
    RST'den gelen solver-generated NS'ler (FIXEDSU vb.) filtrelenir.
    Model agaci bos ise DPF NS'leri fallback olarak kullanilir
    (yine de solver-generated olanlar atlanir).
    """
    names = []
    seen = {}

    # 1) Model agaci: kullanicinin olusturdugu NS'ler
    try:
        for item in ExtAPI.DataModel.Project.Model.NamedSelections.Children:
            nm = item.Name
            if nm not in seen and not _is_solver_generated(nm):
                seen[nm] = True
                names.append(nm)
    except:
        pass

    # 2) DPF fallback: sadece model agaci bossa ve solver-generated degilse
    if not names:
        try:
            for nm in dpf_model.AvailableNamedSelections:
                s = str(nm)
                if s not in seen and not _is_solver_generated(s):
                    seen[s] = True
                    names.append(s)
        except:
            pass

    return names


def _element_node_ids(mesh, eid):
    el = None
    try:
        el = mesh.ElementById(eid)
    except:
        el = None
    if el is None:
        return None
    try:
        return list(el.NodeIds)
    except:
        pass
    try:
        return [n.Id for n in el.Nodes]
    except:
        pass
    try:
        return list(el.CornerNodeIds)
    except:
        pass
    return None

def _make_temp_op(thermal):
    """Thermal (.rth) -> result.temperature; structural (.rst) -> BFE structural_temperature."""
    if thermal:
        try:
            return dpf.operators.result.temperature()
        except:
            pass
    return dpf.operators.result.structural_temperature()

def get_temp_node_map(ds, mesh, set_id, node_id_set, log, thermal=False):
    """
    BFE / structural temperature -> {nid: T}.
    Once Nodal lokasyon denenir; olmazsa ElementalNodal + connectivity.
    """
    # Deneme 1: Nodal
    try:
        op = _make_temp_op(thermal)
        op.inputs.data_sources.Connect(ds)
        op.inputs.time_scoping.Connect([set_id])
        try:
            op.inputs.requested_location.Connect("Nodal")
        except:
            pass
        fc = op.outputs.fields_container.GetData()
        out = {}
        nfields = fc_count(fc)
        for fi in range(nfields):
            field = fc_get_field(fc, fi)
            ids = list(field.ScopingIds)
            data = list(field.Data)
            try:
                comp = int(field.ComponentCount)
            except:
                comp = 1
            for k in range(len(ids)):
                nid = ids[k]
                if nid in node_id_set and nid not in out:
                    out[nid] = data[k * comp]
        if len(out) > 0:
            log.append("TEMP set {0}: Nodal lokasyon OK, {1} node".format(set_id, len(out)))
            return out
        log.append("TEMP set {0}: Nodal lokasyon 0 node dondu, ElementalNodal denenecek".format(set_id))
    except Exception as e:
        log.append("TEMP set {0}: Nodal basarisiz ({1}), ElementalNodal denenecek".format(set_id, str(e)))

    # Deneme 2: ElementalNodal + connectivity (orijinal parse_temp_file gibi
    # node basina son gelen deger yazilir)
    out = {}
    try:
        op = _make_temp_op(thermal)
        op.inputs.data_sources.Connect(ds)
        op.inputs.time_scoping.Connect([set_id])
        fc = op.outputs.fields_container.GetData()
        nfields = fc_count(fc)
        for fi in range(nfields):
            field = fc_get_field(fc, fi)
            try:
                comp = int(field.ComponentCount)
            except:
                comp = 1
            try:
                field_elem_ids = list(field.ScopingIds)
            except:
                field_elem_ids = []
            for eid in field_elem_ids:
                try:
                    vals = list(field.GetEntityDataById(eid))
                except:
                    continue
                nn = len(vals) // comp if comp > 0 else 0
                node_ids = _element_node_ids(mesh, eid)
                if node_ids is None:
                    continue
                limit = min(nn, len(node_ids))
                for j in range(limit):
                    nid = node_ids[j]
                    if nid in node_id_set:
                        out[nid] = vals[j * comp]
        log.append("TEMP set {0}: ElementalNodal ile {1} node".format(set_id, len(out)))
    except Exception as e:
        log.append("TEMP set {0}: ElementalNodal de basarisiz -> {1}".format(set_id, str(e)))
    return out

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

    put(5, c0 + 3, "Value2", "Trendline Coefficient")
    put(6, c0 + 3, "Value2", "Trendline Exponent")
    put(7, c0 + 3, "Value2", "R-squared (fit quality)")
    put(5, c0 + 4, "Value2", A)
    put(6, c0 + 4, "Value2", b)
    put(7, c0 + 4, "Value2", r2)
    put(8, c0 + 3, "Value2", "Trendline Equation")
    put(8, c0 + 4, "Value2", "y = {0}x^{1}".format(trim_num(A), trim_num(b)))

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
        for off, w in ((0, 38), (1, 18), (2, 3), (3, 24), (4, 40), (5, 3)):
            _put(ws.Range(cell(1, c0 + off), cell(1, c0 + off)).EntireColumn,
                 "ColumnWidth", w)

    def _numfmt():
        _put(ws.Range(cell(1, c0 + 1), cell(13, c0 + 1)), "NumberFormat", "General")
        _put(cell(5, c0 + 4), "NumberFormat", num_fmt_auto(A))
        _put(cell(6, c0 + 4), "NumberFormat", num_fmt_auto(b))
        _put(cell(7, c0 + 4), "NumberFormat", num_fmt_auto(r2))
        _put(ws.Range(cell(DATA_ROW0, c0 + 3), cell(r_last, c0 + 4)),
             "NumberFormat", "General")

    def _borders():
        for a1, a2 in (((1, c0), (13, c0 + 1)),
                       ((1, c0 + 3), (3, c0 + 4)),
                       ((5, c0 + 3), (8, c0 + 4))):
            _put(ws.Range(cell(*a1), cell(*a2)).Borders, "LineStyle", 1)

    def _fills():
        for r in (1, 2, 5, 6):                      # input cells: light blue
            _put(cell(r, c0 + 1).Interior, "Color", rgb(221, 235, 247))
        for r in (1, 2, 3):                         # info labels: yellow
            _put(cell(r, c0 + 3).Interior, "Color", rgb(255, 255, 0))
            _put(cell(r, c0 + 3).Font, "Bold", True)
        _put(cell(8, c0 + 4), "HorizontalAlignment", -4131)   # left
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
        _put(ch, "ChartType", 74)                 # xlXYScatterLines (markers + line)
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
            _put(ax, "HasTitle", True)
            _put(ax.AxisTitle, "Text", title)
            _put(ax, "HasMajorGridlines", True)
        _STEP[0] = "block {0}: trendline".format(k + 1)
        try:
            tl = s.Trendlines().Add(4)            # xlPower
            _put(tl, "DisplayEquation", True)
            _put(tl, "DisplayRSquared", False)
            try:
                _put(tl.DataLabel, "NumberFormat", "0.###############")   # up to 15 decimals
                _put(tl.DataLabel.Font, "Size", 9)
            except Exception as ex2:
                warnings.append("trendline label format: " + str(ex2).split("\n")[0])
        except Exception as ex1:
            warnings.append("trendline: " + str(ex1).split("\n")[0])
    except Exception as ex:
        warnings.append("chart: " + str(ex).split("\n")[0])

    return A, b, r2, warnings


def build_excel(datasets, progress=None, node_table=None):
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
            if progress:
                progress("Writing dataset {0} / {1} ...".format(k + 1, len(datasets)))
            A, b, r2, warns = _write_block(ws, k, ds)
            note = getattr(ds, "note", None)
            if note:
                c0 = 1 + k * BLOCK_STRIDE
                _safe(warns, "note", lambda: _put(ws.Cells[4, c0 + 3], "Value2", note))
            out.append((ds, A, b, r2, calc_block(ds, A, b), warns))

        if node_table:
            _STEP[0] = "node results sheet"
            _write_node_sheet(app, wb, node_table[0], node_table[1], progress)

        _STEP[0] = "showing Excel"
        if progress:
            progress("Showing Excel ...")
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


class WaitForm(Form):
    def __init__(self):
        self.Text            = "Strain-Life"
        self.StartPosition   = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.ControlBox      = False
        self.MaximizeBox     = False; self.MinimizeBox = False
        self.ClientSize      = Size(380, 80); self.BackColor = Color.White
        self.TopMost         = True
        self._lbl = Label()
        self._lbl.Text = "Creating Excel, please wait ..."
        self._lbl.Location = Point(20, 28); self._lbl.Size = Size(340, LBL_H)
        self._lbl.Font = HDR_FONT
        self.Controls.Add(self._lbl)

    def update(self, text):
        self._lbl.Text = text
        self.Refresh(); Application.DoEvents()


# =====================================================================

# =====================================================================
# STRAIN-LIFE LOGIC  (pure python - unit tested)
# =====================================================================
class TempFit(object):
    """Everything needed to get a node cycle at one given temperature."""
    def __init__(self, temp, walker, r_ratio, rows):
        self.temp = temp
        self.walker = walker
        self.r_ratio = r_ratio
        self.rows = rows
        xs = [r[0] for r in rows]; ys = [r[1] for r in rows]
        self.A, self.b, self.r2 = power_fit(xs, ys)


def node_cycle(fit, smax, smin, cap):
    """Calculated Cycle of one node at the temperature of `fit`.
    Same formulas as the sheet: B3=B6/B5, B7=(B5-B6)/2, B10=..., B11=B2*B10,
    B12=B10-B11, Cycle=A*ABS(B12)^aaa.  Infinite / undefined life -> cap."""
    try:
        if smax < smin:
            smax, smin = smin, smax
        if smax == 0.0:
            return cap
        B1 = fit.walker; B2 = fit.r_ratio
        B3 = smin / smax
        B7 = (smax - smin) / 2.0
        if B3 == 1.0 or B7 == 0.0:
            return cap
        base = 1.0 - B2
        if base <= 0.0:
            return cap
        B10 = ((B7 * (2.0 / (1.0 - B3))) * ((1.0 - B3) ** B1)) / (base ** B1)
        B11 = B2 * B10
        B12 = B10 - B11
        if B12 == 0.0 or abs(B12) < 1e-300:
            return cap
        c = fit.A * (abs(B12) ** fit.b)
    except (ZeroDivisionError, OverflowError, ValueError):
        return cap
    if c != c or c > cap:           # NaN or above the limit
        return cap
    return max(c, 0.0)


def interp_cycle(temps, cycles, t, cap):
    """Linear interpolation of cycles(temps) at t.  Outside the range: linear
    extrapolation with the nearest two points.  One point: that value.
    temps must be sorted ascending.  Result clamped to [0, cap]."""
    n = len(temps)
    if n == 1:
        return min(max(cycles[0], 0.0), cap)
    if t <= temps[0]:
        i = 0
    elif t >= temps[-1]:
        i = n - 2
    else:
        i = 0
        while i < n - 2 and t > temps[i + 1]:
            i += 1
    t0 = temps[i]; t1 = temps[i + 1]
    c0 = cycles[i]; c1 = cycles[i + 1]
    c = c0 + (c1 - c0) * (t - t0) / (t1 - t0)
    return min(max(c, 0.0), cap)


def compute_nodes(fits, node_strain, node_tmax, cap):
    """fits: TempFit list.  node_strain {nid: (smax, smin)}, node_tmax {nid: T}.
    Returns (rows, temps) ; rows = [(nid, smax, smin, tmax, [cycle@each temp], final)]
    sorted by final cycle ascending."""
    fits = sorted(fits, key=lambda f: f.temp)
    temps = [f.temp for f in fits]
    rows = []
    for nid, (smax, smin) in node_strain.items():
        if len(fits) > 1 and nid not in node_tmax:
            continue
        tmax = node_tmax.get(nid, temps[0])
        cyc = [node_cycle(f, smax, smin, cap) for f in fits]
        fin = interp_cycle(temps, cyc, tmax, cap)
        rows.append((nid, smax, smin, tmax, cyc, fin))
    rows.sort(key=lambda r: (r[5], r[0]))
    return rows, temps


def fmt_cycle(v):
    return "{0:.6g}".format(v)


def ascii_tag(name, maxlen=25):
    s = "".join(c if (c.isalnum() and ord(c) < 128) else "_" for c in safe_str(name))
    return s[:maxlen]


def build_dpf_code(node_vals, result_name):
    """Python Result script (same mechanism as the averaging error tool)."""
    pairs = list(node_vals.items()); chunk = 200; lines = []
    for i in range(0, len(pairs), chunk):
        seg = ", ".join("({0}, {1:.6g})".format(k, v) for k, v in pairs[i:i + chunk])
        lines.append("    for _k, _v in [{0}]: node_vals[_k] = _v".format(seg))
    wf_key = "strainlife_" + ascii_tag(result_name, 45)
    code = [
        "# Strain-Life Cycle - " + result_name,
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
# STRAIN / TEMPERATURE FROM THE RST  (DPF)
# =====================================================================
# Equivalent Total Strain candidates (tested in order).  Names from the DPF operator
# list: wrapper result.total_strain_eqv (= operator "EL_EQV"); older Mechanical DPF has
# no wrapper, but the generic operator name "EL_EQV" may still work.
TOTAL_WRAPPERS = ["total_strain_eqv"]
TOTAL_GENERICS = ["EL_EQV"]
# last resort: elastic eqv + plastic eqv (approximation, see log)
STRAIN_REQUIRED = "elastic_strain_eqv"
STRAIN_OPTIONAL = ["plastic_strain_eqv"]


def _read_fc_nodal(fc, node_id_set, out):
    """fields container -> out {nid: value} (first component)."""
    for fi in range(fc_count(fc)):
        field = fc_get_field(fc, fi)
        ids = list(field.ScopingIds)
        data = list(field.Data)
        try:
            comp = int(field.ComponentCount)
        except Exception:
            comp = 1
        for k in range(len(ids)):
            nid = ids[k]
            if node_id_set is None or nid in node_id_set:
                out[nid] = data[k * comp]
    return out


def _run_wrapper(name, ds, set_id, node_id_set):
    op = getattr(dpf.operators.result, name)()
    op.inputs.data_sources.Connect(ds)
    op.inputs.time_scoping.Connect([set_id])
    try:
        op.inputs.requested_location.Connect("Nodal")
    except Exception:
        pass
    fc = op.outputs.fields_container.GetData()
    return _read_fc_nodal(fc, node_id_set, {})


def _run_generic(name, ds, set_id, node_id_set):
    op = dpf.Operator(name)
    op.Connect(4, ds)
    sc = dpf.Scoping()
    sc.Ids = [set_id]
    op.Connect(0, sc)
    try:
        op.Connect(9, "Nodal")
    except Exception:
        pass
    fc = op.GetOutputAsFieldsContainer(0)
    return _read_fc_nodal(fc, node_id_set, {})


def read_one(kind, name, ds, set_id, node_id_set):
    if kind == "wrapper":
        return _run_wrapper(name, ds, set_id, node_id_set)
    return _run_generic(name, ds, set_id, node_id_set)


def detect_strain_method(ds, set_id, node_id_set, log):
    """Returns ('wrapper'|'generic', name) for a true Equivalent Total Strain operator,
    or ('sum', [elastic, plastic...]) as a fallback."""
    cands = [("wrapper", n) for n in TOTAL_WRAPPERS] + [("generic", n) for n in TOTAL_GENERICS]
    for kind, nm in cands:
        try:
            if kind == "wrapper" and not hasattr(dpf.operators.result, nm):
                log.append("STRAIN: result.{0} not available".format(nm))
                continue
            v = read_one(kind, nm, ds, set_id, node_id_set)
            if len(v) > 0:
                log.append("STRAIN: Equivalent Total Strain = {0} '{1}' ({2} node)".format(
                    kind, nm, len(v)))
                return (kind, nm)
            log.append("STRAIN: {0} '{1}' returned no node".format(kind, nm))
        except Exception as e:
            log.append("STRAIN: {0} '{1}' failed -> {2}".format(kind, nm, str(e).split("\n")[0]))
    parts = [STRAIN_REQUIRED]
    v = _run_wrapper(STRAIN_REQUIRED, ds, set_id, node_id_set)
    if len(v) == 0:
        raise Exception("result.{0} returned no node.".format(STRAIN_REQUIRED))
    for nm in STRAIN_OPTIONAL:
        try:
            if len(_run_wrapper(nm, ds, set_id, node_id_set)) > 0:
                parts.append(nm)
        except Exception:
            pass
    log.append("STRAIN: FALLBACK (approximation) = " + " + ".join(parts))
    return ("sum", parts)


def read_strain_set(method, ds, set_id, node_id_set):
    kind, what = method
    if kind != "sum":
        return read_one(kind, what, ds, set_id, node_id_set)
    tot = {}
    for nm in what:
        try:
            vals = _run_wrapper(nm, ds, set_id, node_id_set)
        except Exception:
            if nm == STRAIN_REQUIRED:
                raise
            continue
        for nid, v in vals.items():
            tot[nid] = tot.get(nid, 0.0) + v
    return tot


def load_step_sets(dpf_model):
    """[(load_step, set_id, time)] - time 0 excluded, last set of each load step."""
    return [(ls, sid, t) for ls, sid, t in get_load_step_info(dpf_model) if t > 0.0]


def read_node_extremes(ds, dpf_model, mesh, node_ids, log, rst_path,
                       progress=None, method=None):
    """Returns (node_strain {nid:(max,min)}, node_tmax {nid:T}, method)."""
    node_set = set(node_ids)
    steps = load_step_sets(dpf_model)
    if not steps:
        raise Exception("RST'de time > 0 olan load step yok.")
    log.append("Load step set'leri (time 0 haric): " +
               ", ".join("LS{0}->set{1}(t={2:g})".format(a, b, c) for a, b, c in steps))
    if method is None:
        method = detect_strain_method(ds, steps[0][1], node_set, log)
    smax = {}; smin = {}
    for i, (ls, sid, t) in enumerate(steps):
        if progress:
            progress("Strain: load step {0} ({1}/{2})".format(ls, i + 1, len(steps)))
        vals = read_strain_set(method, ds, sid, node_set)
        log.append("Strain LS{0} set{1}: {2} node".format(ls, sid, len(vals)))
        for nid, v in vals.items():
            if nid not in smax or v > smax[nid]:
                smax[nid] = v
            if nid not in smin or v < smin[nid]:
                smin[nid] = v
    node_strain = dict((n, (smax[n], smin[n])) for n in smax)

    thermal = is_thermal_file(rst_path)
    tmax = {}
    for i, (ls, sid, t) in enumerate(steps):
        if progress:
            progress("Temperature: load step {0} ({1}/{2})".format(ls, i + 1, len(steps)))
        tm = get_temp_node_map(ds, mesh, sid, node_set, log, thermal)
        for nid, v in tm.items():
            if nid not in tmax or v > tmax[nid]:
                tmax[nid] = v
    log.append("Max temperature: {0} node".format(len(tmax)))
    return node_strain, tmax, method


# =====================================================================
# WINDOWS
# =====================================================================
def _make_list(x, y, w, h, cols):
    lv = ListView()
    lv.Location = Point(x, y); lv.Size = Size(w, h)
    lv.View = View.Details; lv.FullRowSelect = True; lv.GridLines = True
    lv.HideSelection = False; lv.Font = CTRL_FONT
    for text, cw in cols:
        lv.Columns.Add(text, cw)
    return lv


def _list_add(lv, values):
    it = ListViewItem(str(values[0]))
    for v in values[1:]:
        it.SubItems.Add(str(v))
    lv.Items.Add(it)
    return it


class LogForm(Form):
    def __init__(self, title, text):
        self.Text = title
        self.StartPosition = FormStartPosition.CenterScreen
        self.ClientSize = Size(760, 480); self.BackColor = Color.White
        tb = TextBox()
        tb.Multiline = True; tb.ReadOnly = True; tb.WordWrap = False
        tb.ScrollBars = ScrollBars.Both; tb.Font = Font("Consolas", 9)
        tb.Location = Point(10, 10); tb.Size = Size(740, 420)
        tb.Text = text.replace("\n", "\r\n")
        self.Controls.Add(tb)
        b = _grey_btn("Close", 90, 30); b.Location = Point(660, 438)
        b.DialogResult = DialogResult.Cancel
        self.Controls.Add(b); self.CancelButton = b


class ProgressDialog(Form):
    def __init__(self):
        self.Text = "Strain-Life - processing ..."; self.TopMost = True
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedSingle
        self.MaximizeBox = False; self.ControlBox = False
        self.ClientSize = Size(460, 72); self.BackColor = Color.White
        self.l1 = _lbl("Starting ...", 15, 14, 430)
        self.l2 = _lbl("", 15, 38, 430, bold=False)
        self.l2.ForeColor = Color.DimGray
        self.Controls.Add(self.l1); self.Controls.Add(self.l2)

    def update(self, a, b=""):
        self.l1.Text = a; self.l2.Text = b
        Application.DoEvents()


# ---------------------------------------------------------------------
# PROFILES (saved in the registry, like the Yield table profiles)
# ---------------------------------------------------------------------
_REG_KEY = r"Software\StrainLifeTool\Profiles"


def encode_profile(serie, entries):
    """entries {temp: {walker, r, rows}} -> one string."""
    parts = []
    for t in sorted(entries):
        en = entries[t]
        data = ";".join("{0!r}:{1!r}".format(x, y) for x, y in en["rows"])
        parts.append("{0!r}|{1!r}|{2!r}|{3}".format(t, en["walker"], en["r"], data))
    return serie.replace("#", " ").replace("|", " ") + "@" + "#".join(parts)


def decode_profile(raw):
    """-> (serie, entries)"""
    serie, _, body = raw.partition("@")
    entries = {}
    for item in body.split("#"):
        f = item.split("|")
        if len(f) != 4:
            continue
        rows = []
        for pr in f[3].split(";"):
            xy = pr.split(":")
            if len(xy) == 2:
                rows.append((float(xy[0]), float(xy[1])))
        entries[float(f[0])] = {"walker": float(f[1]), "r": float(f[2]), "rows": rows}
    return serie, entries


try:
    import Microsoft.Win32 as Win32

    def _reg_names():
        try:
            key = Win32.Registry.CurrentUser.OpenSubKey(_REG_KEY)
            if key is None:
                return []
            names = [str(n) for n in key.GetValueNames()]; key.Close()
            return sorted(names)
        except Exception:
            return []

    def _reg_save(name, raw):
        try:
            key = Win32.Registry.CurrentUser.CreateSubKey(_REG_KEY)
            key.SetValue(name, raw); key.Close()
            return True
        except Exception:
            return False

    def _reg_load(name):
        try:
            key = Win32.Registry.CurrentUser.OpenSubKey(_REG_KEY)
            if key is None:
                return None
            raw = key.GetValue(name); key.Close()
            return str(raw) if raw else None
        except Exception:
            return None

    def _reg_delete(name):
        try:
            key = Win32.Registry.CurrentUser.OpenSubKey(_REG_KEY, True)
            if key:
                key.DeleteValue(name, False); key.Close()
        except Exception:
            pass
except Exception:
    def _reg_names(): return []
    def _reg_save(n, r): return False
    def _reg_load(n): return None
    def _reg_delete(n): pass


def _input_dialog(title, prompt):
    dlg = Form()
    dlg.Text = title; dlg.StartPosition = FormStartPosition.CenterScreen
    dlg.FormBorderStyle = FormBorderStyle.FixedDialog
    dlg.MaximizeBox = False; dlg.MinimizeBox = False
    dlg.ClientSize = Size(360, 110); dlg.BackColor = Color.White
    dlg.Controls.Add(_lbl(prompt, 15, 12, 330, bold=False))
    tb = TextBox(); tb.Font = CTRL_FONT
    tb.Location = Point(15, 38); tb.Size = Size(330, 24)
    dlg.Controls.Add(tb)
    ok = _blue_btn("OK", 80, 28); ok.Location = Point(175, 72)
    ok.DialogResult = DialogResult.OK
    dlg.Controls.Add(ok); dlg.AcceptButton = ok
    ca = _grey_btn("Cancel", 80, 28); ca.Location = Point(265, 72)
    ca.DialogResult = DialogResult.Cancel
    dlg.Controls.Add(ca); dlg.CancelButton = ca
    if dlg.ShowDialog() == DialogResult.OK:
        return tb.Text.strip()
    return None


# ---------------------------------------------------------------------
# WINDOW 1 - Serie + per temperature: Walker, R ratio, Strain/Cycle grid
# ---------------------------------------------------------------------
class SetupDialog(Form):
    def __init__(self, serie, entries):
        self.action = "cancel"
        self.serie = serie
        self.entries = dict(entries)       # temp -> dict(walker, r, rows)
        self.Text            = "Strain-Life  -  Temperature data"
        self.StartPosition   = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox     = False; self.MinimizeBox = False
        self.ClientSize      = Size(700, 690); self.BackColor = Color.White

        self.Controls.Add(_lbl("Profile:", 15, 18, 55))
        self._cmb = ComboBox(); self._cmb.DropDownStyle = ComboBoxStyle.DropDownList
        self._cmb.Font = CTRL_FONT
        self._cmb.Location = Point(75, 15); self._cmb.Size = Size(200, 24)
        self._cmb.SelectedIndexChanged += self._on_profile
        self.Controls.Add(self._cmb)
        b_save = _blue_btn("Save", 60, 26); b_save.Location = Point(285, 14)
        b_save.Click += self._on_save
        self.Controls.Add(b_save)
        b_sas = _grey_btn("Save As", 70, 26); b_sas.Location = Point(351, 14)
        b_sas.Click += self._on_save_as
        self.Controls.Add(b_sas)
        b_dp = _red_btn("Delete", 65, 26); b_dp.Location = Point(427, 14)
        b_dp.Click += self._on_delete_profile
        self.Controls.Add(b_dp)
        self._fill_profiles(None)

        self.Controls.Add(_lbl("Serie:", 15, 58, 45, bold=False))
        self._tb_serie = TextBox(); self._tb_serie.Font = CTRL_FONT
        self._tb_serie.Location = Point(62, 55); self._tb_serie.Size = Size(110, 24)
        self._tb_serie.Text = serie
        self.Controls.Add(self._tb_serie)
        self.Controls.Add(_lbl("Temperature:", 190, 58, 85, bold=False))
        self._tb_t = TextBox(); self._tb_t.Font = CTRL_FONT
        self._tb_t.Location = Point(278, 55); self._tb_t.Size = Size(75, 24)
        self.Controls.Add(self._tb_t)
        self.Controls.Add(_lbl("Walker:", 368, 58, 55, bold=False))
        self._tb_w = TextBox(); self._tb_w.Font = CTRL_FONT
        self._tb_w.Location = Point(425, 55); self._tb_w.Size = Size(70, 24)
        self.Controls.Add(self._tb_w)
        self.Controls.Add(_lbl("R ratio:", 510, 58, 55, bold=False))
        self._tb_r = TextBox(); self._tb_r.Font = CTRL_FONT
        self._tb_r.Location = Point(568, 55); self._tb_r.Size = Size(70, 24)
        self.Controls.Add(self._tb_r)

        self.Controls.Add(_lbl("Total Strain and Cycle:", 15, 92, 400))
        lbl2 = Label()
        lbl2.Text = "Paste here: copy the 2 columns (Total Strain | Cycle) in Excel, click the first cell and press Ctrl+V."
        lbl2.Location = Point(15, 112); lbl2.Size = Size(670, LBL_H)
        lbl2.Font = CTRL_FONT; lbl2.ForeColor = Color.FromArgb(0, 100, 200)
        self.Controls.Add(lbl2)

        self._grid = DataGridView()
        self._grid.Location              = Point(15, 135)
        self._grid.Size                  = Size(670, 230)
        self._grid.AllowUserToAddRows    = False
        self._grid.AllowUserToDeleteRows = False
        self._grid.RowHeadersVisible     = False
        self._grid.AutoSizeColumnsMode   = DataGridViewAutoSizeColumnsMode.Fill
        self._grid.Font                  = CTRL_FONT
        self._grid.BackgroundColor       = Color.White
        self._grid.GridColor             = Color.FromArgb(200, 200, 200)
        c1 = DataGridViewTextBoxColumn(); c1.Name = "strain"; c1.HeaderText = "Total Strain"
        c2 = DataGridViewTextBoxColumn(); c2.Name = "cycle";  c2.HeaderText = "Cycle"
        self._grid.Columns.Add(c1); self._grid.Columns.Add(c2)
        self._grid.KeyDown += self._on_grid_key
        self.Controls.Add(self._grid)
        self._set_grid([])
        self.Shown += self._on_shown

        b_ar = _grey_btn("+ Add Row", 100, 26); b_ar.Location = Point(15, 370)
        b_ar.Click += lambda s, e: self._grid.Rows.Add("", "")
        self.Controls.Add(b_ar)
        b_dr = _grey_btn("- Delete Row", 110, 26); b_dr.Location = Point(123, 370)
        b_dr.Click += self._del_row
        self.Controls.Add(b_dr)
        b_cl = _red_btn("Clear All", 90, 26); b_cl.Location = Point(241, 370)
        b_cl.Click += lambda s, e: self._set_grid([])
        self.Controls.Add(b_cl)
        b_pa = _blue_btn("Paste from Clipboard", 150, 26); b_pa.Location = Point(338, 370)
        b_pa.Click += lambda s, e: self._paste()
        self.Controls.Add(b_pa)
        b_add = _blue_btn("Add / Update", 190, 30); b_add.Location = Point(495, 368)
        b_add.Click += self._on_add
        self.Controls.Add(b_add)

        self.Controls.Add(_lbl("Added temperatures  (click a row to edit it)", 15, 412, 400))
        self._lv = _make_list(15, 435, 670, 150,
                              [("Temperature", 150), ("Walker", 150), ("R ratio", 150), ("Rows", 150)])
        self._lv.SelectedIndexChanged += self._on_select
        self.Controls.Add(self._lv)
        self._refresh_list()

        b_rem = _red_btn("Remove selected", 130, 30); b_rem.Location = Point(15, 592)
        b_rem.Click += self._on_remove
        self.Controls.Add(b_rem)
        b_next = _blue_btn("Next  ->", 100, 32); b_next.Location = Point(480, 645)
        b_next.Click += self._on_next
        self.Controls.Add(b_next)
        b_cancel = _grey_btn("Cancel", 90, 32); b_cancel.Location = Point(595, 645)
        b_cancel.DialogResult = DialogResult.Cancel
        self.Controls.Add(b_cancel); self.CancelButton = b_cancel

    # ---------------- grid ----------------
    def _on_shown(self, s, e):
        try:
            self._grid.Focus()
            self._grid.CurrentCell = self._grid.Rows[0].Cells[0]
        except Exception:
            pass

    def _paste(self):
        import System.Windows.Forms as WF
        try:
            rows, skipped = parse_pairs(WF.Clipboard.GetText() or "")
        except Exception as ex:
            MessageBox.Show("Paste error: " + str(ex), "Strain-Life")
            return False
        if not rows:
            MessageBox.Show("The clipboard has no two-column numeric data.\n"
                            "Copy the Total Strain and Cycle columns in Excel first.",
                            "Strain-Life", MessageBoxButtons.OK, MessageBoxIcon.Information)
            return False
        self._set_grid(rows)
        return True

    def _on_grid_key(self, sender, ev):
        import System.Windows.Forms as WF
        if ev.Control and ev.KeyCode == WF.Keys.V:
            self._paste()
            ev.Handled = True

    def _del_row(self, s, e):
        if self._grid.SelectedRows.Count > 0:
            self._grid.Rows.RemoveAt(self._grid.SelectedRows[0].Index)
        elif self._grid.CurrentCell is not None:
            self._grid.Rows.RemoveAt(self._grid.CurrentCell.RowIndex)
        elif self._grid.Rows.Count > 0:
            self._grid.Rows.RemoveAt(self._grid.Rows.Count - 1)

    def _grid_rows(self):
        rows = []
        for i in range(self._grid.Rows.Count):
            try:
                a = self._grid.Rows[i].Cells[0].Value
                b = self._grid.Rows[i].Cells[1].Value
                if a is None or b is None:
                    continue
                rows.append((to_float(str(a)), to_float(str(b))))
            except Exception:
                pass
        return rows

    def _set_grid(self, rows):
        self._grid.Rows.Clear()
        for x, y in rows:
            self._grid.Rows.Add(trim_num(x), trim_num(y))
        for i in range(max(0, GRID_MIN_ROWS - len(rows))):     # empty white rows
            self._grid.Rows.Add("", "")

    # ---------------- list ----------------
    def _temps(self):
        return sorted(self.entries.keys())

    def _refresh_list(self):
        self._lv.Items.Clear()
        for t in self._temps():
            en = self.entries[t]
            _list_add(self._lv, [trim_num(t), trim_num(en["walker"]),
                                 trim_num(en["r"]), len(en["rows"])])

    def _on_select(self, s, e):
        idx = [int(i) for i in self._lv.SelectedIndices]
        if not idx:
            return
        t = self._temps()[idx[0]]
        en = self.entries[t]
        self._tb_t.Text = trim_num(t); self._tb_w.Text = trim_num(en["walker"])
        self._tb_r.Text = trim_num(en["r"]); self._set_grid(en["rows"])

    def _on_add(self, s, e):
        errs = []
        t = w = r = None
        try:
            t = to_float(self._tb_t.Text)
        except Exception:
            errs.append("- Temperature is not a valid number.")
        try:
            w = to_float(self._tb_w.Text)
        except Exception:
            errs.append("- Walker is not a valid number.")
        try:
            r = to_float(self._tb_r.Text)
            if r >= 1.0:
                errs.append("- R ratio must be less than 1.")
        except Exception:
            errs.append("- R ratio is not a valid number.")
        rows = self._grid_rows()
        if len(rows) < 2:
            errs.append("- At least 2 valid (Total Strain, Cycle) rows are needed.")
        else:
            try:
                power_fit([x[0] for x in rows], [x[1] for x in rows])
            except Exception as ex:
                errs.append("- Power fit failed: " + str(ex).split("\n")[0])
        if errs:
            MessageBox.Show("\n".join(errs), "Check inputs",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        self.entries[t] = {"walker": w, "r": r, "rows": rows}
        self._refresh_list()
        self._tb_t.Text = ""; self._tb_w.Text = ""; self._tb_r.Text = ""
        self._set_grid([])
        self._tb_t.Focus()

    def _on_remove(self, s, e):
        idx = [int(i) for i in self._lv.SelectedIndices]
        temps = self._temps()
        for i in idx:
            del self.entries[temps[i]]
        self._refresh_list()

    # ---------------- profiles ----------------
    def _fill_profiles(self, select):
        self._loading = True
        self._cmb.Items.Clear()
        self._cmb.Items.Add("-- New --")
        sel = 0
        for i, n in enumerate(_reg_names()):
            self._cmb.Items.Add(n)
            if n == select:
                sel = i + 1
        self._cmb.SelectedIndex = sel
        self._loading = False

    def _on_profile(self, s, e):
        if getattr(self, "_loading", False) or self._cmb.SelectedIndex <= 0:
            return
        raw = _reg_load(str(self._cmb.SelectedItem))
        if not raw:
            return
        try:
            serie, entries = decode_profile(raw)
        except Exception:
            MessageBox.Show("Profile could not be read.", "Strain-Life")
            return
        self.entries = entries
        if serie:
            self._tb_serie.Text = serie
        self._refresh_list()
        self._set_grid([])
        self._tb_t.Text = ""; self._tb_w.Text = ""; self._tb_r.Text = ""

    def _save_profile(self, name):
        if not self.entries:
            MessageBox.Show("Add at least one temperature first (Add / Update).",
                            "Strain-Life", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        if _reg_save(name, encode_profile(self._tb_serie.Text.strip() or "Serie 1", self.entries)):
            self._fill_profiles(name)
            MessageBox.Show("Profile '{0}' saved.".format(name), "Saved",
                            MessageBoxButtons.OK, MessageBoxIcon.Information)
        else:
            MessageBox.Show("Profile could not be saved.", "Strain-Life")

    def _on_save(self, s, e):
        if self._cmb.SelectedIndex <= 0:
            self._on_save_as(s, e)
        else:
            self._save_profile(str(self._cmb.SelectedItem))

    def _on_save_as(self, s, e):
        name = _input_dialog("Save As", "New profile name:")
        if name:
            self._save_profile(name)

    def _on_delete_profile(self, s, e):
        if self._cmb.SelectedIndex <= 0:
            MessageBox.Show("No profile selected.", "Strain-Life")
            return
        name = str(self._cmb.SelectedItem)
        if MessageBox.Show("Delete profile '{0}'?".format(name), "Confirm",
                           MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes:
            _reg_delete(name)
            self._fill_profiles(None)

    # ---------------- next ----------------
    def _on_next(self, s, e):
        if not self.entries:
            MessageBox.Show("Add at least one temperature (fill the fields, paste the data, "
                            "press 'Add / Update').", "Check inputs",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        self.serie = self._tb_serie.Text.strip() or "Serie 1"
        self.action = "next"
        self.DialogResult = DialogResult.OK; self.Close()


# ---------------------------------------------------------------------
# WINDOW 3 - analysis / RST / Named Selection / cap
# ---------------------------------------------------------------------
class RunDialog(Form):
    def __init__(self, analyses, prev_cap):
        self.action = "cancel"
        self._analyses = analyses
        self.ds = None; self.dpf_model = None; self.rst_path = ""
        self.analysis = None; self.ns_name = None; self.cap = prev_cap
        self.Text            = "Strain-Life  -  Analysis / RST / Named Selection"
        self.StartPosition   = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox     = False; self.MinimizeBox = False
        self.ClientSize      = Size(580, 340); self.BackColor = Color.White

        self.Controls.Add(_lbl("Analysis:", 15, 18, 140))
        self._cmb_an = ComboBox(); self._cmb_an.DropDownStyle = ComboBoxStyle.DropDownList
        self._cmb_an.Font = CTRL_FONT
        self._cmb_an.Location = Point(165, 15); self._cmb_an.Size = Size(400, 24)
        sel = 0
        for i, a in enumerate(analyses):
            nm = safe_str(getattr(a, "Name", "Analysis {0}".format(i + 1)))
            self._cmb_an.Items.Add(nm)
            try:
                if "static" in safe_str(a.AnalysisType).lower() and sel == 0:
                    sel = i
            except Exception:
                pass
        if analyses:
            self._cmb_an.SelectedIndex = sel
        self._cmb_an.SelectedIndexChanged += self._on_analysis
        self.Controls.Add(self._cmb_an)

        self.Controls.Add(_lbl("RST file:", 15, 58, 140))
        self._tb_rst = TextBox(); self._tb_rst.Font = CTRL_FONT; self._tb_rst.ReadOnly = True
        self._tb_rst.Location = Point(165, 55); self._tb_rst.Size = Size(300, 24)
        self.Controls.Add(self._tb_rst)
        b_br = _grey_btn("Browse ...", 85, 26); b_br.Location = Point(472, 54)
        b_br.Click += self._on_browse
        self.Controls.Add(b_br)

        self.Controls.Add(_lbl("Name Selection:", 15, 98, 140))
        self._cmb_ns = ComboBox(); self._cmb_ns.DropDownStyle = ComboBoxStyle.DropDownList
        self._cmb_ns.Font = CTRL_FONT
        self._cmb_ns.Location = Point(165, 95); self._cmb_ns.Size = Size(400, 24)
        self.Controls.Add(self._cmb_ns)
        attach_ns_name_warning(self._cmb_ns)

        self.Controls.Add(_lbl("Cycle upper limit:", 15, 138, 140))
        self._tb_cap = TextBox(); self._tb_cap.Font = CTRL_FONT
        self._tb_cap.Location = Point(165, 135); self._tb_cap.Size = Size(100, 24)
        self._tb_cap.Text = "{0:g}".format(prev_cap)
        self.Controls.Add(self._tb_cap)
        note = Label()
        note.Text = ("Infinite life, or any result above this limit, is set to the limit value. "
                     "The contour shows the cycle value itself (not log).")
        note.Location = Point(275, 130); note.Size = Size(290, 52)
        note.Font = CTRL_FONT; note.ForeColor = Color.DimGray
        self.Controls.Add(note)

        info = Label()
        info.Text = ("Equivalent Total Strain and temperature are read from the RST for every "
                     "load step (time 0 excluded, last set of each load step); the max / min "
                     "strain and the max temperature of each node are used.")
        info.Location = Point(15, 195); info.Size = Size(550, 60)
        info.Font = CTRL_FONT; info.ForeColor = Color.FromArgb(0, 100, 200)
        self.Controls.Add(info)

        b_back = _grey_btn("<- Back", 85, 32); b_back.Location = Point(15, 290)
        b_back.Click += self._on_back
        self.Controls.Add(b_back)
        b_run = _blue_btn("Run", 90, 32); b_run.Location = Point(380, 290)
        b_run.Click += self._on_run
        self.Controls.Add(b_run)
        b_cancel = _grey_btn("Cancel", 85, 32); b_cancel.Location = Point(480, 290)
        b_cancel.DialogResult = DialogResult.Cancel
        self.Controls.Add(b_cancel); self.CancelButton = b_cancel

        self._fill_ns()
        if analyses:
            self._on_analysis(None, None)

    def _fill_ns(self):
        self._cmb_ns.Items.Clear()
        self._cmb_ns.Items.Add("-- Select --")
        try:
            names = collect_ns_names(self.ds, self.dpf_model)
        except Exception:
            names = []
        for n in names:
            self._cmb_ns.Items.Add(n)
        self._cmb_ns.SelectedIndex = 0

    def _on_analysis(self, s, e):
        if self._cmb_an.SelectedIndex < 0:
            return
        a = self._analyses[self._cmb_an.SelectedIndex]
        try:
            rst = safe_str(a.ResultFileName)
        except Exception:
            rst = ""
        self._load_rst(rst if rst and os.path.exists(rst) else "")

    def _on_browse(self, s, e):
        dlg = OpenFileDialog()
        dlg.Filter = "Ansys Result File (*.rst;*.rth)|*.rst;*.rth|All files (*.*)|*.*"
        if self.rst_path and os.path.exists(self.rst_path):
            dlg.InitialDirectory = os.path.dirname(self.rst_path)
        if dlg.ShowDialog() == DialogResult.OK:
            self._load_rst(dlg.FileName)

    def _load_rst(self, path):
        self.ds = None; self.dpf_model = None; self.rst_path = path
        self._tb_rst.Text = path if path else "(no result file - solve or browse)"
        if path:
            try:
                self.ds, self.dpf_model = dpf_load(path)
            except Exception as ex:
                MessageBox.Show("RST could not be read:\n" + safe_str(ex), "Error")
                self.ds = None; self.dpf_model = None
        self._fill_ns()

    def _on_back(self, s, e):
        self.action = "back"
        self.DialogResult = DialogResult.OK; self.Close()

    def _on_run(self, s, e):
        errs = []
        if self.ds is None:
            errs.append("- No valid RST file.")
        if self._cmb_ns.SelectedIndex <= 0:
            errs.append("- Select a Name Selection.")
        try:
            cap = to_float(self._tb_cap.Text)
            if cap <= 0:
                errs.append("- Cycle upper limit must be > 0.")
        except Exception:
            errs.append("- Cycle upper limit is not a number.")
            cap = DEFAULT_CAP
        if errs:
            MessageBox.Show("\n".join(errs), "Missing input",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        self.analysis = self._analyses[self._cmb_an.SelectedIndex]
        self.ns_name = safe_str(self._cmb_ns.SelectedItem)
        self.cap = cap
        self.action = "run"
        self.DialogResult = DialogResult.OK; self.Close()


# ---------------------------------------------------------------------
# RESULTS WINDOW
# ---------------------------------------------------------------------
class ResultsForm(Form):
    def __init__(self, serie, fits, temps, rows, cap, info, log_text, export_fn):
        self._export_fn = export_fn
        self._log_text = log_text
        self.Text            = "Strain-Life  -  Results"
        self.StartPosition   = FormStartPosition.CenterScreen
        self.ClientSize      = Size(900, 640); self.BackColor = Color.White
        self.MinimizeBox = False

        self.Controls.Add(_lbl(info, 15, 10, 870, bold=True))

        lv1 = _make_list(15, 42, 870, 28 + 20 * min(len(fits), 6),
                         [("Serie", 150), ("Temperature", 120), ("Walker", 120), ("R ratio", 110),
                          ("Trendline Coefficient", 190), ("Trendline Exponent", 170)])
        for f in sorted(fits, key=lambda f: f.temp):
            _list_add(lv1, [serie, trim_num(f.temp), trim_num(f.walker), trim_num(f.r_ratio),
                            trim_num(f.A), trim_num(f.b)])
        self.Controls.Add(lv1)

        y2 = 42 + lv1.Height + 12
        shown = rows[:MAX_TABLE_ROWS]
        title = "Node results  (lowest cycle first"
        if len(rows) > len(shown):
            title += ";  first {0} of {1} shown - Export has all".format(len(shown), len(rows))
        title += ")"
        self.Controls.Add(_lbl(title, 15, y2, 870))
        cols = [("Node", 80), ("Max strain", 100), ("Min strain", 100), ("Max temp", 80)]
        for t in temps:
            cols.append(("Cycle @ " + trim_num(t), 105))
        cols.append(("Final cycle", 120))
        lv2 = _make_list(15, y2 + 24, 870, 640 - (y2 + 24) - 55, cols)
        for r in shown:
            vals = [r[0], "{0:.6g}".format(r[1]), "{0:.6g}".format(r[2]),
                    "{0:.4g}".format(r[3])]
            vals += [fmt_cycle(c) for c in r[4]]
            vals.append(fmt_cycle(r[5]))
            _list_add(lv2, vals)
        self.Controls.Add(lv2)

        b_x = _blue_btn("Export to Excel", 140, 32); b_x.Location = Point(15, 640 - 45)
        b_x.Click += self._on_export
        self.Controls.Add(b_x)
        b_l = _grey_btn("Log", 80, 32); b_l.Location = Point(165, 640 - 45)
        b_l.Click += self._on_log
        self.Controls.Add(b_l)
        b_c = _grey_btn("Close", 90, 32); b_c.Location = Point(795, 640 - 45)
        b_c.DialogResult = DialogResult.Cancel
        self.Controls.Add(b_c); self.CancelButton = b_c

    def _on_log(self, s, e):
        LogForm("Strain-Life log", self._log_text).ShowDialog()

    def _on_export(self, s, e):
        self._export_fn()


# =====================================================================
# EXPORT TO EXCEL  (optional)
# =====================================================================
def _write_node_sheet(app, wb, header, table, progress=None):
    import System
    if progress:
        progress("Writing node results ...")
    ws2 = wb.Sheets.Add(System.Reflection.Missing.Value, wb.Sheets[wb.Sheets.Count])
    try:
        ws2.Name = NODE_SHEET
    except Exception:
        pass
    nr = len(table) + 1; nc = len(header)
    arr = System.Array.CreateInstance(System.Object, nr, nc)
    for j, h in enumerate(header):
        arr[0, j] = h
    for i, row in enumerate(table):
        for j, v in enumerate(row):
            arr[i + 1, j] = v
    rng = ws2.Range(ws2.Cells[1, 1], ws2.Cells[nr, nc])
    rng.Value2 = arr
    try:
        ws2.Range(ws2.Cells[1, 1], ws2.Cells[1, nc]).Font.Bold = True
        ws2.Columns.AutoFit()
    except Exception:
        pass


def make_export(fits, temps, rows, cap, serie):
    """Returns a function that builds the Excel (per-temperature blocks + node sheet)."""
    def export():
        if not rows:
            return
        fits_sorted = sorted(fits, key=lambda f: f.temp)
        datasets = []
        for k, f in enumerate(fits_sorted):
            # critical node = lowest cycle at this temperature
            best = min(rows, key=lambda r: (r[4][k], r[0]))
            ds = DataSet(f.temp, serie, f.r_ratio, f.walker, best[1], best[2], f.rows)
            ds.note = "Critical node (lowest cycle at this T): {0}".format(best[0])
            datasets.append(ds)
        header = ["Node ID", "Max Strain", "Min Strain", "Max Temp"]
        header += ["Cycle @ " + trim_num(t) for t in temps] + ["Final Cycle"]
        table = [[r[0], r[1], r[2], r[3]] + list(r[4]) + [r[5]] for r in rows]
        wf = WaitForm()
        try:
            wf.Show()
            Cursor.Current = Cursors.WaitCursor
            res = build_excel(datasets, wf.update, node_table=(header, table))
        except Exception as ex:
            try:
                wf.Close()
            except Exception:
                pass
            Cursor.Current = Cursors.Default
            MessageBox.Show("Excel error  (step: {0}):\n{1}".format(
                _STEP[0], str(ex).split("\n")[0]), "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error)
            return
        try:
            wf.Close()
        except Exception:
            pass
        Cursor.Current = Cursors.Default
    return export


# =====================================================================
# MECHANICAL PYTHON RESULT
# =====================================================================
def reverse_legend():
    """Red = minimum life, blue = maximum life (Mechanical default is the opposite).
    1) reverse the band colours one by one, 2) fallback: Reverse Rainbow scheme.
    Returns '' on success or a short message."""
    try:
        leg = Ansys.Mechanical.Graphics.Tools.CurrentLegendSettings()
    except Exception as ex:
        return "Legend not available ({0}); set Legend > Color Scheme > Reverse Rainbow.".format(
            str(ex).split("\n")[0])
    try:
        n = int(leg.NumberOfBands)
        cols = [leg.GetBandColor(i) for i in range(n)]
        for i in range(n):
            leg.SetBandColor(i, cols[n - 1 - i])
        return ""
    except Exception as ex1:
        try:
            import System
            et = leg.ColorScheme.GetType()
            leg.ColorScheme = System.Enum.Parse(et, "ReverseRainbow")
            return ""
        except Exception as ex2:
            return "Legend could not be reversed ({0}); set Legend > Color Scheme > Reverse Rainbow.".format(
                str(ex2).split("\n")[0])


def add_contour(analysis, rows, ns_name, summary):
    sol = analysis.Solution
    rname = "StrainLife_Cycle_{0}".format(ascii_tag(ns_name))
    for ch in list(sol.Children):
        try:
            if ch.Name == rname:
                ch.Delete()
        except Exception:
            pass
    pr = sol.AddPythonResult()
    pr.Name = rname
    pr.Text = build_dpf_code(dict((r[0], r[5]) for r in rows), rname)
    try:
        pr.Connect()
    except Exception:
        pass
    try:
        sol.EvaluateAllResults()
    except Exception:
        pass
    try:
        c = pr.AddComment(); c.Text = summary
    except Exception:
        pass
    try:
        pr.Activate()
    except Exception:
        pass
    msg = reverse_legend()
    # force a full redraw so that the node points pick up the new colours too
    try:
        ExtAPI.DataModel.Tree.Activate([sol])
    except Exception:
        pass
    try:
        pr.Activate()
    except Exception:
        pass
    try:
        ExtAPI.Graphics.Redraw()
    except Exception:
        pass
    return pr, msg


# =====================================================================
# WIZARD + MAIN
# =====================================================================
def run_wizard(analyses):
    """Returns (serie, fits, run_dialog) or None."""
    serie = "Serie 1"; entries = {}
    cap = DEFAULT_CAP
    step = 0
    while True:
        if step == 0:
            d = SetupDialog(serie, entries)
            d.ShowDialog()
            if d.action != "next":
                return None
            serie = d.serie; entries = d.entries
            step = 1
        else:
            d = RunDialog(analyses, cap)
            d.ShowDialog()
            if d.action == "back":
                step = 0; cap = d.cap
                continue
            if d.action != "run":
                return None
            fits = [TempFit(t, entries[t]["walker"], entries[t]["r"], entries[t]["rows"])
                    for t in sorted(entries)]
            return serie, fits, d


def run_calculation(fits, d, prog, log):
    ds, model = d.ds, d.dpf_model
    prog.update("Reading mesh ...")
    mesh = get_mesh(ds, model, log)
    coords, unit = build_coords_dict(ds, model, mesh, log)
    prog.update("Resolving Named Selection ...", d.ns_name)
    ids = resolve_ns_ids(ds, d.ns_name, False, log, coords)
    if not ids:
        raise Exception("Named Selection '{0}' could not be converted to RST nodes.".format(d.ns_name))
    log.append("NS nodes: {0}".format(len(ids)))
    node_strain, tmax, method = read_node_extremes(
        ds, model, mesh, ids, log, d.rst_path,
        progress=lambda t: prog.update(t, d.ns_name))
    if not node_strain:
        raise Exception("No strain values were read for the Named Selection nodes.")
    if len(fits) > 1 and not tmax:
        raise Exception("No node temperatures could be read from the RST "
                        "(needed for interpolation).")
    prog.update("Computing cycles ...", "{0} nodes".format(len(node_strain)))
    rows, temps = compute_nodes(fits, node_strain, tmax, d.cap)
    if not rows:
        raise Exception("No node has both strain and temperature.")
    log.append("Method: {0}  |  nodes computed: {1}".format(method, len(rows)))
    return rows, temps, method


def main():
    try:
        model = ExtAPI.DataModel.Project.Model
        analyses = list(model.Analyses)
    except Exception:
        analyses = []
    if not analyses:
        MessageBox.Show("No analysis found in the project.", "Strain-Life",
                        MessageBoxButtons.OK, MessageBoxIcon.Error)
        return
    w = run_wizard(analyses)
    if w is None:
        return
    serie, fits, d = w
    log = []
    prog = ProgressDialog()
    prog.Show()
    try:
        Cursor.Current = Cursors.WaitCursor
        rows, temps, method = run_calculation(fits, d, prog, log)
        nmin = rows[0]
        cap_n = len([r for r in rows if r[5] >= d.cap])
        summary = ("Strain-Life cycle | NS={0} | nodes={1} | min cycle={2} @ node {3} "
                   "| nodes at the upper limit ({4:g}): {5} of {1}").format(
                       d.ns_name, len(rows), fmt_cycle(nmin[5]), nmin[0], d.cap, cap_n)
        prog.update("Creating Python Result ...")
        _pr, legend_msg = add_contour(d.analysis, rows, d.ns_name, summary)
        if legend_msg:
            log.append(legend_msg)
    except Exception as ex:
        try:
            prog.Close()
        except Exception:
            pass
        Cursor.Current = Cursors.Default
        log.append("HATA: " + safe_str(ex))
        LogForm("Strain-Life - error (copy this log)", "\n".join(log)).ShowDialog()
        return
    try:
        prog.Close()
    except Exception:
        pass
    Cursor.Current = Cursors.Default
    ResultsForm(serie, fits, temps, rows, d.cap, summary, "\n".join(log),
                make_export(fits, temps, rows, d.cap, serie)).ShowDialog()


if not os.environ.get("STRAIN_LIFE_NO_RUN"):
    main()
