# -*- coding: utf-8 -*-
# ============================================================
# Strain read test 2 - ANSYS Mechanical (IronPython 2.7)
# Pick an analysis (RST), ONE Named Selection and a load step -> shows, per node, the
# different strain values side by side so they can be compared with Mechanical's
# "Equivalent Total" strain result:
#   Elastic eqv | Plastic eqv | Thermal eqv | Total (generic operator EL_EQV) |
#   Total (result.total_strain_eqv) | Elastic+Plastic | Elastic+Plastic+Thermal
# (columns whose operator does not exist in this Mechanical are left out and listed)
# "Copy all to clipboard" -> paste into Excel.
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



# =====================================================================
# WINDOW
# =====================================================================
def _make_list(x, y, w, h, cols):
    lv = ListView()
    lv.Location = Point(x, y); lv.Size = Size(w, h)
    lv.View = View.Details; lv.FullRowSelect = True; lv.GridLines = True
    lv.HideSelection = False; lv.Font = CTRL_FONT
    for text, cw in cols:
        lv.Columns.Add(text, cw)
    return lv


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


CANDIDATES = [
    ("Elastic eqv",  [("wrapper", "elastic_strain_eqv"), ("generic", "EPEL_EQV")]),
    ("Plastic eqv",  [("wrapper", "plastic_strain_eqv"), ("generic", "EPPL_EQV")]),
    ("Thermal eqv",  [("wrapper", "thermal_strains_eqv"), ("generic", "ETH_EQV")]),
    ("Total (EL_EQV)",  [("generic", "EL_EQV")]),
    ("Total (total_strain_eqv)", [("wrapper", "total_strain_eqv")]),
]


class StrainTestForm(Form):
    def __init__(self, analyses):
        self._analyses = analyses
        self.ds = None; self.dpf_model = None; self.rst_path = ""
        self._log = []; self._table = []; self._header = []
        self.Text            = "Strain read test  -  Named Selection -> Equivalent Total Strain"
        self.StartPosition   = FormStartPosition.CenterScreen
        self.ClientSize      = Size(940, 640); self.BackColor = Color.White

        self.Controls.Add(_lbl("Analysis:", 15, 18, 100))
        self._cmb_an = ComboBox(); self._cmb_an.DropDownStyle = ComboBoxStyle.DropDownList
        self._cmb_an.Font = CTRL_FONT
        self._cmb_an.Location = Point(120, 15); self._cmb_an.Size = Size(500, 24)
        sel = 0
        for i, a in enumerate(analyses):
            self._cmb_an.Items.Add(safe_str(getattr(a, "Name", "Analysis {0}".format(i + 1))))
            try:
                if "static" in safe_str(a.AnalysisType).lower() and sel == 0:
                    sel = i
            except Exception:
                pass
        if analyses:
            self._cmb_an.SelectedIndex = sel
        self._cmb_an.SelectedIndexChanged += self._on_analysis
        self.Controls.Add(self._cmb_an)

        self.Controls.Add(_lbl("RST file:", 15, 52, 100))
        self._tb_rst = TextBox(); self._tb_rst.Font = CTRL_FONT; self._tb_rst.ReadOnly = True
        self._tb_rst.Location = Point(120, 49); self._tb_rst.Size = Size(410, 24)
        self.Controls.Add(self._tb_rst)
        b_br = _grey_btn("Browse ...", 85, 26); b_br.Location = Point(535, 48)
        b_br.Click += self._on_browse
        self.Controls.Add(b_br)

        self.Controls.Add(_lbl("Named Selection:", 15, 86, 105))
        self._cmb_ns = ComboBox(); self._cmb_ns.DropDownStyle = ComboBoxStyle.DropDownList
        self._cmb_ns.Font = CTRL_FONT
        self._cmb_ns.Location = Point(120, 83); self._cmb_ns.Size = Size(500, 24)
        self.Controls.Add(self._cmb_ns)
        attach_ns_name_warning(self._cmb_ns)

        self.Controls.Add(_lbl("Load step:", 640, 86, 75))
        self._cmb_step = ComboBox(); self._cmb_step.DropDownStyle = ComboBoxStyle.DropDownList
        self._cmb_step.Font = CTRL_FONT
        self._cmb_step.Location = Point(718, 83); self._cmb_step.Size = Size(205, 24)
        self.Controls.Add(self._cmb_step)
        self._steps = []

        b_run = _blue_btn("Read strain", 120, 40); b_run.Location = Point(650, 30)
        b_run.Click += self._on_run
        self.Controls.Add(b_run)

        self._lbl_info = _lbl("", 15, 118, 910, bold=False)
        self._lbl_info.ForeColor = Color.FromArgb(0, 100, 200)
        self.Controls.Add(self._lbl_info)

        self._lv = _make_list(15, 145, 910, 440, [("Node", 90)])
        self.Controls.Add(self._lv)

        b_copy = _grey_btn("Copy all to clipboard", 160, 32); b_copy.Location = Point(15, 595)
        b_copy.Click += self._on_copy
        self.Controls.Add(b_copy)
        b_log = _grey_btn("Log", 80, 32); b_log.Location = Point(185, 595)
        b_log.Click += lambda s, e: LogForm("Log", "\n".join(self._log)).ShowDialog()
        self.Controls.Add(b_log)
        b_close = _grey_btn("Close", 90, 32); b_close.Location = Point(835, 595)
        b_close.DialogResult = DialogResult.Cancel
        self.Controls.Add(b_close); self.CancelButton = b_close

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
        dlg.Filter = "Ansys Result File (*.rst)|*.rst|All files (*.*)|*.*"
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
        self._fill_ns()
        self._fill_steps()

    def _on_copy(self, s, e):
        if not self._table:
            return
        import System.Windows.Forms as WF
        lines = ["\t".join(self._header)]
        for r in self._table:
            lines.append("\t".join(str(v) for v in r))
        try:
            WF.Clipboard.SetText("\r\n".join(lines))
            MessageBox.Show("{0} rows copied.".format(len(self._table)), "Strain read test")
        except Exception as ex:
            MessageBox.Show("Copy failed: " + safe_str(ex), "Strain read test")

    def _on_run(self, s, e):
        if self.ds is None:
            MessageBox.Show("Select a valid RST file.", "Strain read test"); return
        if self._cmb_ns.SelectedIndex <= 0:
            MessageBox.Show("Select a Named Selection.", "Strain read test"); return
        if self._cmb_step.SelectedIndex < 0:
            MessageBox.Show("No load step (time > 0) found in the RST.", "Strain read test"); return
        ns_name = safe_str(self._cmb_ns.SelectedItem)
        ls, sid, tt = self._steps[self._cmb_step.SelectedIndex]
        log = []; self._log = log
        self._lv.Items.Clear(); self._table = []
        Cursor.Current = Cursors.WaitCursor
        try:
            self._lbl_info.Text = "Reading mesh / Named Selection ..."; Application.DoEvents()
            mesh = get_mesh(self.ds, self.dpf_model, log)
            coords, unit = build_coords_dict(self.ds, self.dpf_model, mesh, log)
            ids = resolve_ns_ids(self.ds, ns_name, False, log, coords)
            if not ids:
                raise Exception("Named Selection could not be converted to RST nodes.")
            node_set = set(ids)
            log.append("NS nodes: {0}   |   LS{1} -> set {2} (t={3:g})".format(len(ids), ls, sid, tt))
            data = []          # [(label, {nid: v} or None)]
            for label, alts in CANDIDATES:
                self._lbl_info.Text = "Reading: " + label; Application.DoEvents()
                vals = None
                for kind, nm in alts:
                    try:
                        if kind == "wrapper" and not hasattr(dpf.operators.result, nm):
                            log.append("{0}: result.{1} not available".format(label, nm))
                            continue
                        v = read_one(kind, nm, self.ds, sid, node_set)
                        if len(v) > 0:
                            vals = v
                            log.append("{0}: {1} '{2}' OK ({3} node)".format(label, kind, nm, len(v)))
                            break
                        log.append("{0}: {1} '{2}' returned no node".format(label, kind, nm))
                    except Exception as ex:
                        log.append("{0}: {1} '{2}' failed -> {3}".format(
                            label, kind, nm, safe_str(ex).split("\n")[0]))
                data.append((label, vals))
        except Exception as ex:
            Cursor.Current = Cursors.Default
            log.append("ERROR: " + safe_str(ex))
            self._lbl_info.Text = "Error - see Log"
            LogForm("Strain read test - error", "\n".join(log)).ShowDialog()
            return
        Cursor.Current = Cursors.Default

        dd = dict(data)
        el = dd.get("Elastic eqv"); pl = dd.get("Plastic eqv"); th = dd.get("Thermal eqv")
        cols = [lab for lab, v in data if v is not None]
        extra = []
        if el is not None and pl is not None:
            extra.append("Elastic + Plastic")
        if el is not None and pl is not None and th is not None:
            extra.append("Elastic + Plastic + Thermal")
        self._header = ["Node"] + cols + extra
        for nid in sorted(node_set):
            row = [nid]
            for lab in cols:
                v = dd[lab].get(nid)
                row.append("%.8g" % v if v is not None else "")
            if "Elastic + Plastic" in extra:
                e0 = el.get(nid); p0 = pl.get(nid)
                row.append("%.8g" % (e0 + p0) if e0 is not None and p0 is not None else "")
            if "Elastic + Plastic + Thermal" in extra:
                e0 = el.get(nid); p0 = pl.get(nid); h0 = th.get(nid)
                row.append("%.8g" % (e0 + p0 + h0)
                           if None not in (e0, p0, h0) else "")
            self._table.append(row)
        self._lv.Columns.Clear()
        self._lv.Columns.Add("Node", 80)
        for h in self._header[1:]:
            self._lv.Columns.Add(h, 130)
        for r in self._table[:5000]:
            it = ListViewItem(str(r[0]))
            for v in r[1:]:
                it.SubItems.Add(str(v))
            self._lv.Items.Add(it)
        missing = [lab for lab, v in data if v is None]
        self._lbl_info.Text = ("{0}  |  LS{1} (t={2:g})  |  {3} nodes  |  not available: {4}{5}".format(
            ns_name, ls, tt, len(self._table), ", ".join(missing) if missing else "none",
            "  |  first 5000 shown (Copy has all)" if len(self._table) > 5000 else ""))


def main():
    try:
        analyses = list(ExtAPI.DataModel.Project.Model.Analyses)
    except Exception:
        analyses = []
    if not analyses:
        MessageBox.Show("No analysis found in the project.", "Strain read test",
                        MessageBoxButtons.OK, MessageBoxIcon.Error)
        return
    StrainTestForm(analyses).ShowDialog()


if not os.environ.get("STRAIN_LIFE_NO_RUN"):
    main()
