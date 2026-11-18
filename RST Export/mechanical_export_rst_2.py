# -*- coding: utf-8 -*-
# =============================================================================
# Ansys Mechanical 2024 - Displacement, Force & Temperature Export (RST/DPF)
#
# Orijinal aracin RST tabanli versiyonu:
#   - Sonuc objesi EKLENMEZ, EvaluateAllResults CAGRILMAZ.
#   - Tum degerler secilen .rst dosyasindan DPF ile dogrudan okunur.
#   - Popup duzeni orijinalle ayni, sadece en uste "RST File" alani eklendi.
#   - Cikti formati birebir ayni (tab'li TXT: NodeID, X_Loc, Y_Loc, Step kolonlari).
#
# Notlar:
#   - "Step N" = rst icindeki result SET N. (Store Results At = Last Time Points
#     ile cozulmus modellerde set sayisi = step sayisidir.)
#   - Named Selection ID'leri once rst icindeki component'lerden okunur (hizli),
#     bulunamazsa acik Mechanical modelindeki NS agacindan alinir.
#   - Degerler rst'nin solver birim sistemindedir.
# =============================================================================

import os
import clr
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

from System.Windows.Forms import (Form, Label, ComboBox, Button, FolderBrowserDialog,
                                   TextBox, DialogResult, FormStartPosition,
                                   ComboBoxStyle, MessageBox, CheckBox,
                                   FormBorderStyle, GroupBox, FlatStyle,
                                   Application, OpenFileDialog,
                                   MessageBoxButtons, MessageBoxIcon)
from System.Drawing import Size, Point, Font, FontStyle, Color
from System import Type as _NetType
from System import Int32
from System.Collections.Generic import List as _NetList

import time
import math as _math
import Microsoft.Win32 as Win32

import mech_dpf
import Ans.DataProcessing as dpf

mech_dpf.setExtAPI(ExtAPI)


# -----------------------------------------------------------------------
# COMBOBOX MOUSE WHEEL SCROLL BLOCKER
# Prevents accidental selection change when scrolling the form
# -----------------------------------------------------------------------
def _no_scroll_cmb(cmb):
    """Prevent mouse-wheel from changing a ComboBox selection."""
    def _block(s, e):
        try: e.Handled = True
        except: pass
    cmb.MouseWheel += _block
    return cmb

def _fmt_duration(seconds):
    """Format duration: under 60s -> '53 sec', over 60s -> '3 min 14 sec'"""
    s = int(seconds)
    if s < 60:
        return "{0} sec".format(s)
    m = s // 60
    r = s % 60
    return "{0} min {1} sec".format(m, r)

# -----------------------------------------------------------------------
# REGISTRY HELPERS  (tek bir shared output dir + son rst klasoru)
# -----------------------------------------------------------------------
_REG_KEY = r"Software\NodeExport\Shared"

def _reg_load(name):
    try:
        key = Win32.Registry.CurrentUser.OpenSubKey(_REG_KEY)
        if key is None:
            return ""
        val = key.GetValue(name)
        key.Close()
        return str(val) if val else ""
    except:
        return ""

def _reg_save(name, path):
    try:
        key = Win32.Registry.CurrentUser.CreateSubKey(_REG_KEY)
        key.SetValue(name, path)
        key.Close()
    except:
        pass

def load_shared_dir():
    return _reg_load("LastOutputDir")

def save_shared_dir(path):
    _reg_save("LastOutputDir", path)

def load_rst_dir():
    return _reg_load("LastRstDir")

def save_rst_dir(path):
    _reg_save("LastRstDir", path)

# Eski isimleri kullanan çağrılar için compat shim (kullanılmıyor ama kalsın)
def load_last_dir(key_suffix):
    return load_shared_dir()

def save_last_dir(key_suffix, path):
    save_shared_dir(path)

# -----------------------------------------------------------------------
# GENERAL HELPERS
# -----------------------------------------------------------------------
def fmt(val):
    return "{0:.20f}".format(val).rstrip("0").rstrip(".")

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
# =======================================================================
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

# ------------------- RESULT FIELDS -------------------------------------
def get_disp_node_map(ds, set_id, node_id_set):
    """Bir set icin {nid: (ux, uy)} dondurur (sadece istenen node'lar)."""
    op = dpf.operators.result.displacement()
    op.inputs.data_sources.Connect(ds)
    op.inputs.time_scoping.Connect([set_id])
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
            comp = 3
        for k in range(len(ids)):
            nid = ids[k]
            if nid in node_id_set and nid not in out:
                ux = data[k * comp + 0]
                uy = data[k * comp + 1] if comp > 1 else 0.0
                out[nid] = (ux, uy)
    return out

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

def get_enf_node_map(ds, mesh, set_id, elem_ids, edge_node_set, log):
    """
    ENF (element nodal forces) -> edge node'lara dusen FX/FY toplami.
    Ayni node'a birden fazla elemandan gelen katkilar TOPLANIR
    (orijinal parse_enf_file davranisi).
    Donen: {nid: (fx_sum, fy_sum)}
    """
    op = dpf.operators.result.element_nodal_forces()
    op.inputs.data_sources.Connect(ds)
    op.inputs.time_scoping.Connect([set_id])

    # Once eleman scoping'i ile dene (hizli)
    scoped = False
    try:
        sc = dpf.Scoping()
        sc.Location = "Elemental"
        try:
            sc.Ids = _to_net_int_list(elem_ids)
        except:
            sc.Ids = elem_ids
        op.inputs.mesh_scoping.Connect(sc)
        scoped = True
    except Exception as e:
        log.append("ENF set {0}: scoping baglanamadi ({1}), tum alan okunacak".format(set_id, str(e)))

    fc = op.outputs.fields_container.GetData()

    result = {}
    nfields = fc_count(fc)
    elem_set = set(elem_ids)
    matched_elems = 0

    for fi in range(nfields):
        field = fc_get_field(fc, fi)
        try:
            comp = int(field.ComponentCount)
        except:
            comp = 3
        try:
            field_elem_ids = list(field.ScopingIds)
        except:
            field_elem_ids = []

        for eid in field_elem_ids:
            if eid not in elem_set:
                continue
            try:
                vals = list(field.GetEntityDataById(eid))
            except:
                continue
            if comp <= 0:
                continue
            nn = len(vals) // comp
            node_ids = _element_node_ids(mesh, eid)
            if node_ids is None:
                continue
            matched_elems += 1
            limit = min(nn, len(node_ids))
            for j in range(limit):
                nid = node_ids[j]
                if nid not in edge_node_set:
                    continue
                fx = vals[j * comp + 0]
                fy = vals[j * comp + 1] if comp > 1 else 0.0
                if nid in result:
                    ofx, ofy = result[nid]
                    result[nid] = (ofx + fx, ofy + fy)
                else:
                    result[nid] = (fx, fy)

    log.append("ENF set {0}: {1} eleman islendi, {2} edge node eslesti (scoped={3})".format(
        set_id, matched_elems, len(result), scoped))
    return result

def get_temp_node_map(ds, mesh, set_id, node_id_set, log):
    """
    BFE / structural temperature -> {nid: T}.
    Once Nodal lokasyon denenir; olmazsa ElementalNodal + connectivity.
    """
    # Deneme 1: Nodal
    try:
        op = dpf.operators.result.structural_temperature()
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
        op = dpf.operators.result.structural_temperature()
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

# -----------------------------------------------------------------------
# UI HELPERS
# -----------------------------------------------------------------------
ROW_H = 28  # height per NS combo row

STEP_PICKER_THRESHOLD = 10  # use picker dialog if steps exceed this


class StepPickerDialog(Form):
    """Split-panel step selector for large step counts."""
    def __init__(self, num_steps, preselected=None, step_labels=None):
        self.selected_steps = list(preselected) if preselected is not None else []
        self._step_labels   = step_labels or {}  # {step_no: label_str}

        self.Text            = "Select Steps to Export"
        self.StartPosition   = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox     = False
        self.MinimizeBox     = False
        self.ClientSize      = Size(500, 420)

        from System.Windows.Forms import ListBox, SelectionMode
        self._ListBox = ListBox
        self._SelectionMode = SelectionMode

        lbl_left = Label()
        lbl_left.Text     = "Available Steps"
        lbl_left.Location = Point(15, 12)
        lbl_left.Size     = Size(180, 20)
        lbl_left.Font     = Font(self.Font, FontStyle.Bold)
        self.Controls.Add(lbl_left)

        self.lst_avail = ListBox()
        self.lst_avail.Location      = Point(15, 35)
        self.lst_avail.Size          = Size(180, 320)
        self.lst_avail.SelectionMode = SelectionMode.MultiExtended
        self.Controls.Add(self.lst_avail)

        btn_add = Button()
        btn_add.Text     = "Add >"
        btn_add.Location = Point(210, 150)
        btn_add.Size     = Size(80, 32)
        btn_add.ForeColor = Color.FromArgb(0, 100, 200)
        btn_add.Click   += self._on_add
        self.Controls.Add(btn_add)

        btn_del = Button()
        btn_del.Text      = "< Remove"
        btn_del.Location  = Point(210, 192)
        btn_del.Size      = Size(80, 32)
        btn_del.ForeColor = Color.FromArgb(200, 0, 0)
        btn_del.Click    += self._on_del
        self.Controls.Add(btn_del)

        lbl_right = Label()
        lbl_right.Text     = "Selected Steps"
        lbl_right.Location = Point(305, 12)
        lbl_right.Size     = Size(180, 20)
        lbl_right.Font     = Font(self.Font, FontStyle.Bold)
        self.Controls.Add(lbl_right)

        self.lst_sel = ListBox()
        self.lst_sel.Location      = Point(305, 35)
        self.lst_sel.Size          = Size(180, 320)
        self.lst_sel.SelectionMode = SelectionMode.MultiExtended
        self.Controls.Add(self.lst_sel)

        btn_ok = Button()
        btn_ok.Text     = "OK"
        btn_ok.Location = Point(305, 368)
        btn_ok.Size     = Size(80, 30)
        btn_ok.Click   += self._on_ok
        self.Controls.Add(btn_ok)

        btn_cancel = Button()
        btn_cancel.Text         = "Cancel"
        btn_cancel.Location     = Point(395, 368)
        btn_cancel.Size         = Size(90, 30)
        btn_cancel.DialogResult = DialogResult.Cancel
        self.Controls.Add(btn_cancel)
        self.CancelButton = btn_cancel

        self._num_steps = num_steps
        self._populate()

    def _populate(self):
        sel_set = set(self.selected_steps)
        self.lst_avail.Items.Clear()
        self.lst_sel.Items.Clear()
        for s in range(1, self._num_steps + 1):
            lbl = self._step_labels.get(s, "Step {0}".format(s))
            if s not in sel_set:
                self.lst_avail.Items.Add(lbl)
        for s in sorted(sel_set):
            lbl = self._step_labels.get(s, "Step {0}".format(s))
            self.lst_sel.Items.Add(lbl)

    def _parse_step(self, item_str):
        # "Step 3  (t=1.5s)" -> 3
        s = str(item_str).strip()
        # step_labels sozlugunden ters arama
        for step_no, lbl in self._step_labels.items():
            if lbl == s:
                return step_no
        # Fallback: "Step N" formatini parse et
        try:
            return int(s.split()[1])
        except:
            return int("".join(c for c in s if c.isdigit()) or "1")

    def _on_add(self, sender, args):
        to_add = [self._parse_step(self.lst_avail.Items[i])
                  for i in range(self.lst_avail.Items.Count)
                  if self.lst_avail.GetSelected(i)]
        for s in to_add:
            if s not in self.selected_steps:
                self.selected_steps.append(s)
        self.selected_steps.sort()
        self._populate()

    def _on_del(self, sender, args):
        to_del = set(self._parse_step(self.lst_sel.Items[i])
                     for i in range(self.lst_sel.Items.Count)
                     if self.lst_sel.GetSelected(i))
        self.selected_steps = [s for s in self.selected_steps if s not in to_del]
        self._populate()

    def _on_ok(self, sender, args):
        if not self.selected_steps:
            MessageBox.Show("Please select at least one step.", "Warning")
            return
        self.DialogResult = DialogResult.OK
        self.Close()


def make_step_checks(num_steps, step_labels=None):
    """Returns (controls_list, step_checks_list, block_height)
    controls_list: [(ctrl, rel_x, rel_y), ...] - caller positions them.
    step_labels: {step_no: label_str} -- opsiyonel, varsa checkbox metnini override eder.
    Steps default to UNCHECKED."""
    step_checks = []
    controls    = []
    labels      = step_labels or {}

    lbl = Label()
    lbl.Text = "Steps to Evaluate:"
    lbl.Size = Size(150, 20)
    controls.append((lbl, 8, 4))

    if num_steps <= STEP_PICKER_THRESHOLD:
        def select_all(sender, args):
            for c in step_checks:
                c.Checked = True
        def clear_all(sender, args):
            for c in step_checks:
                c.Checked = False

        btn_all = Button()
        btn_all.Text  = "All"
        btn_all.Size  = Size(50, 22)
        btn_all.Click += select_all
        controls.append((btn_all, 165, 0))

        btn_none = Button()
        btn_none.Text  = "Clear"
        btn_none.Size  = Size(50, 22)
        btn_none.Click += clear_all
        controls.append((btn_none, 220, 0))

        # Substep varsa label daha genis, satir basi 2 checkbox
        has_labels = any(labels.get(i + 1, "") != "" for i in range(num_steps))
        if has_labels:
            cols = 2; col_w = 230; row_h = 24
            cb_w = 228
        else:
            cols = 6; col_w = 80; row_h = 24
            cb_w = 78

        for i in range(num_steps):
            s   = i + 1
            cb  = CheckBox()
            cb.Text    = labels.get(s, "Step {0}".format(s))
            cb.Checked = False
            cb.Size    = Size(cb_w, 20)
            cb.Tag     = str(s)
            rx = 8 + (i % cols) * col_w
            ry = 26 + (i // cols) * row_h
            controls.append((cb, rx, ry))
            step_checks.append(cb)

        rows_used = (num_steps + cols - 1) // cols
        block_h   = 26 + rows_used * row_h + 8

    else:
        btn_pick = Button()
        btn_pick.Text = "Select Steps...  (none selected)"
        btn_pick.Size = Size(220, 26)
        btn_pick.Tag  = []

        def open_picker(sender, args):
            dlg = StepPickerDialog(num_steps, preselected=btn_pick.Tag,
                                   step_labels=labels)
            if dlg.ShowDialog() == DialogResult.OK:
                btn_pick.Tag  = dlg.selected_steps
                count = len(dlg.selected_steps)
                if count == 0:
                    btn_pick.Text = "Select Steps...  (none selected)"
                elif count == num_steps:
                    btn_pick.Text = "Select Steps...  (all {0} selected)".format(num_steps)
                else:
                    btn_pick.Text = "Select Steps...  ({0} selected)".format(count)

        btn_pick.Click += open_picker
        controls.append((btn_pick, 165, 0))
        block_h = 32

    return controls, step_checks, block_h

def get_selected_steps(step_checks, step_ctrls):
    if step_checks:
        return [int(cb.Tag) for cb in step_checks if cb.Checked]
    else:
        for ctrl, rx, ry in step_ctrls:
            if hasattr(ctrl, "Tag") and isinstance(ctrl.Tag, list):
                return list(ctrl.Tag)
        return []

def validate_steps(step_checks, step_ctrls):
    return len(get_selected_steps(step_checks, step_ctrls)) > 0

def make_dir_controls(reg_suffix):
    """Output folder control - shared across all panels, remembers last dir."""
    controls = []

    lbl = Label()
    lbl.Text = "Output Folder:"
    lbl.Size = Size(100, 20)
    controls.append((lbl, 8, 3))

    txt = TextBox()
    txt.Size      = Size(310, 22)
    txt.ReadOnly  = False
    txt.BackColor = Color.White
    txt.Text      = load_shared_dir()
    controls.append((txt, 114, 0))

    def _validate(s, e):
        val = txt.Text.strip()
        if not val:
            txt.BackColor = Color.White
            return
        if os.path.isdir(val):
            txt.BackColor = Color.White
            save_shared_dir(val)
        else:
            txt.BackColor = Color.FromArgb(255, 220, 220)

    def _on_key(s, e):
        from System.Windows.Forms import Keys
        if e.KeyCode == Keys.Return:
            _validate(s, e)

    txt.Leave   += _validate
    txt.KeyDown += _on_key

    def browse(sender, args):
        dlg = FolderBrowserDialog()
        dlg.Description         = "Select Output Folder"
        dlg.ShowNewFolderButton = True
        cur = txt.Text.strip()
        if cur and os.path.exists(cur):
            dlg.SelectedPath = cur
        if dlg.ShowDialog() == DialogResult.OK:
            folder = dlg.SelectedPath
            txt.Text = folder
            save_shared_dir(folder)

    btn = Button()
    btn.Text      = "..."
    btn.Size      = Size(40, 24)
    btn.BackColor = Color.FromArgb(0, 120, 215)
    btn.ForeColor = Color.White
    btn.FlatStyle = FlatStyle.Flat
    btn.Click    += browse
    controls.append((btn, 430, 0))

    return controls, txt, 36

def place_controls(grp, controls_list, y_offset):
    for ctrl, rx, ry in controls_list:
        ctrl.Location = Point(rx, y_offset + ry)

def add_controls_to_grp(grp, controls_list):
    for ctrl, rx, ry in controls_list:
        grp.Controls.Add(ctrl)

# -----------------------------------------------------------------------
# ELEMENT SIDE DIALOG
# Two reference images + warning + Yes/No question
# -----------------------------------------------------------------------
# !!! ONEMLI: Asagidaki satira ORIJINAL scriptinizdeki _IMG_ELEM_2D_B64
# degerini birebir kopyalayip yapistirin (binary blob oldugu icin elle
# yeniden uretilemez). Bos birakilirsa dialog calisir, sadece 2D referans
# gorseli yerine metin gosterilir.

class ElementSideDialog(Form):
    """
    Cut BC yon secimi - 2 sayfali wizard:
      Sayfa 1: 3D/2D taraf secimi (resme tiklayinca secilir, kapanmaz)
      Sayfa 2: Number of Sectors / Tangential Node Count (resimli aciklama ile)
    Sonuc: elements_on_3d_side, num_sectors, num_nodes (Finish'e basinca dolar)
    """
    _B64_3D     = "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"
    _B64_2D     = "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"
    _B64_SECTOR = "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"
    _B64_NODES  = "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"

    def __init__(self):
        import System
        from System.IO import MemoryStream
        from System.Windows.Forms import PictureBox, PictureBoxSizeMode, Panel as _WPanel
        from System.Drawing import Image as DImage

        self.elements_on_3d_side = None
        self.num_sectors = 1
        self.num_nodes   = 1
        self._page = 1

        self.Text            = "Cut BC - Element Side"
        self.StartPosition   = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox     = False
        self.MinimizeBox     = False

        IMG_W  = 320
        IMG_H  = 220
        MARGIN = 12
        GAP    = 12
        BORDER = 4
        total_w  = MARGIN + IMG_W + GAP + IMG_W + MARGIN
        ICON_SZ  = 64
        WARN_H   = 70
        INSTR_H  = 22
        BTN_H    = 32
        NAV_H    = 32
        P2_ROW_H = 56   # page2: label + numeric spinner (stacked)

        page1_h = (WARN_H + MARGIN + INSTR_H + 8 + (IMG_H + BORDER * 2) + 8 + BTN_H)
        page2_h = (INSTR_H + 8 + (IMG_H + BORDER * 2) + 8 + P2_ROW_H)

        self._total_w        = total_w
        self._page1_total_h  = MARGIN + page1_h + GAP + NAV_H + MARGIN
        self._page2_total_h  = MARGIN + page2_h + GAP + NAV_H + MARGIN
        self._nav_y_page1    = MARGIN + page1_h + GAP
        self._nav_y_page2    = MARGIN + page2_h + GAP

        self.ClientSize = Size(total_w, self._page1_total_h)

        def load_b64(b64str):
            ms = MemoryStream(System.Convert.FromBase64String(b64str))
            return DImage.FromStream(ms)

        # =====================================================================
        # PAGE 1: 3D / 2D taraf secimi
        # =====================================================================
        self.pnl_page1 = _WPanel()
        self.pnl_page1.Location = Point(0, 0)
        self.pnl_page1.Size    = Size(total_w, MARGIN + page1_h)
        self.pnl_page1.BackColor = self.BackColor
        self.Controls.Add(self.pnl_page1)

        y = MARGIN

        icon_y = y + (WARN_H - ICON_SZ) // 2
        tri = _WPanel()
        tri.Location  = Point(MARGIN, icon_y)
        tri.Size      = Size(ICON_SZ, ICON_SZ)
        tri.BackColor = self.BackColor
        def _tri(s, e):
            from System.Drawing import SolidBrush, PointF
            g = e.Graphics
            SZ = float(ICON_SZ)
            pts = System.Array.CreateInstance(PointF, 3)
            pts[0] = PointF(SZ * 0.50, SZ * 0.02)
            pts[1] = PointF(SZ * 0.02, SZ * 0.96)
            pts[2] = PointF(SZ * 0.98, SZ * 0.96)
            b = SolidBrush(Color.FromArgb(200, 30, 30))
            g.FillPolygon(b, pts); b.Dispose()
            w = SolidBrush(Color.White)
            bar_w = SZ * 0.12
            g.FillRectangle(w, SZ * 0.5 - bar_w / 2.0, SZ * 0.22, bar_w, SZ * 0.40)
            dot_sz = SZ * 0.12
            g.FillEllipse(w, SZ * 0.5 - dot_sz / 2.0, SZ * 0.68, dot_sz, dot_sz)
            w.Dispose()
        tri.Paint += _tri
        self.pnl_page1.Controls.Add(tri)

        # Hanging-indent madde isaretleri: "-" ve cumle ayri label'larda,
        # boylece satir kaydiginda cumlenin ilk harfiyle hizali kalir.
        DASH_W  = 14
        TEXT_X  = MARGIN + ICON_SZ + 8
        SENT_X  = TEXT_X + DASH_W
        SENT_W  = total_w - MARGIN - SENT_X
        ROW_H_B = (WARN_H - 2) // 2   # iki madde arasinda 2px bosluk

        warn_font = Font("Segoe UI", 9.5, FontStyle.Bold)
        warn_color = Color.FromArgb(200, 30, 30)

        lbl_dash1 = Label()
        lbl_dash1.Text      = "-"
        lbl_dash1.Location  = Point(TEXT_X, y)
        lbl_dash1.Size      = Size(DASH_W, ROW_H_B)
        lbl_dash1.Font      = warn_font
        lbl_dash1.ForeColor = warn_color
        self.pnl_page1.Controls.Add(lbl_dash1)

        lbl_sent1 = Label()
        lbl_sent1.Text      = ("The selected Edge and Elements in the name selections must be adjacent. "
                               "Elements must share nodes with the edge.")
        lbl_sent1.Location  = Point(SENT_X, y)
        lbl_sent1.Size      = Size(SENT_W, ROW_H_B)
        lbl_sent1.Font      = warn_font
        lbl_sent1.ForeColor = warn_color
        self.pnl_page1.Controls.Add(lbl_sent1)

        y_row2 = y + ROW_H_B + 2

        lbl_dash2 = Label()
        lbl_dash2.Text      = "-"
        lbl_dash2.Location  = Point(TEXT_X, y_row2)
        lbl_dash2.Size      = Size(DASH_W, ROW_H_B)
        lbl_dash2.Font      = warn_font
        lbl_dash2.ForeColor = warn_color
        self.pnl_page1.Controls.Add(lbl_dash2)

        lbl_sent2 = Label()
        lbl_sent2.Text      = ("The images below are representative; however, the working logic "
                               "is the same for all Cut BC regions.")
        lbl_sent2.Location  = Point(SENT_X, y_row2)
        lbl_sent2.Size      = Size(SENT_W, ROW_H_B)
        lbl_sent2.Font      = warn_font
        lbl_sent2.ForeColor = warn_color
        self.pnl_page1.Controls.Add(lbl_sent2)

        y += WARN_H + MARGIN

        lbl_instr = Label()
        lbl_instr.Text      = "Select the side where your elements are located relative to the 3D submodel"
        lbl_instr.Location  = Point(MARGIN, y)
        lbl_instr.Size      = Size(total_w - MARGIN * 2, INSTR_H)
        lbl_instr.Font      = Font("Segoe UI", 9.5, FontStyle.Bold)
        lbl_instr.ForeColor = Color.FromArgb(0, 70, 150)
        self.pnl_page1.Controls.Add(lbl_instr)

        y += INSTR_H + 8

        try:
            img3d = load_b64(self._B64_3D)
            img2d = load_b64(self._B64_2D)
        except:
            img3d = img2d = None

        DEFAULT_FRAME = Color.FromArgb(225, 225, 225)
        self._FRAME_3D_ON = Color.FromArgb(20, 90, 160)
        self._FRAME_2D_ON = Color.FromArgb(40, 130, 90)

        self.pnl_3d_frame = _WPanel()
        self.pnl_3d_frame.Location  = Point(MARGIN, y)
        self.pnl_3d_frame.Size      = Size(IMG_W + BORDER * 2, IMG_H + BORDER * 2)
        self.pnl_3d_frame.BackColor = DEFAULT_FRAME
        self.pnl_page1.Controls.Add(self.pnl_3d_frame)

        pb3d = PictureBox()
        pb3d.Location  = Point(BORDER, BORDER)
        pb3d.Size      = Size(IMG_W, IMG_H)
        pb3d.SizeMode  = PictureBoxSizeMode.Zoom
        pb3d.BackColor = Color.FromArgb(245, 245, 245)
        if img3d: pb3d.Image = img3d
        self.pnl_3d_frame.Controls.Add(pb3d)

        self.pnl_2d_frame = _WPanel()
        self.pnl_2d_frame.Location  = Point(MARGIN + IMG_W + GAP, y)
        self.pnl_2d_frame.Size      = Size(IMG_W + BORDER * 2, IMG_H + BORDER * 2)
        self.pnl_2d_frame.BackColor = DEFAULT_FRAME
        self.pnl_page1.Controls.Add(self.pnl_2d_frame)

        pb2d = PictureBox()
        pb2d.Location  = Point(BORDER, BORDER)
        pb2d.Size      = Size(IMG_W, IMG_H)
        pb2d.SizeMode  = PictureBoxSizeMode.Zoom
        pb2d.BackColor = Color.FromArgb(245, 245, 245)
        if img2d: pb2d.Image = img2d
        self.pnl_2d_frame.Controls.Add(pb2d)

        y += IMG_H + BORDER * 2 + 8

        self.btn_3d = Button()
        self.btn_3d.Text      = "Select this side"
        self.btn_3d.Size      = Size(IMG_W, BTN_H)
        self.btn_3d.Location  = Point(MARGIN, y)
        self.btn_3d.BackColor = Color.FromArgb(20, 90, 160)
        self.btn_3d.ForeColor = Color.White
        self.btn_3d.FlatStyle = FlatStyle.Flat
        self.btn_3d.Click    += self._on_select_3d
        self.pnl_page1.Controls.Add(self.btn_3d)

        self.btn_2d = Button()
        self.btn_2d.Text      = "Select this side"
        self.btn_2d.Size      = Size(IMG_W, BTN_H)
        self.btn_2d.Location  = Point(MARGIN + IMG_W + GAP, y)
        self.btn_2d.BackColor = Color.FromArgb(40, 130, 90)
        self.btn_2d.ForeColor = Color.White
        self.btn_2d.FlatStyle = FlatStyle.Flat
        self.btn_2d.Click    += self._on_select_2d
        self.pnl_page1.Controls.Add(self.btn_2d)

        # =====================================================================
        # PAGE 2: Number of Sectors / Tangential Node Count
        # =====================================================================
        self.pnl_page2 = _WPanel()
        self.pnl_page2.Location  = Point(0, 0)
        self.pnl_page2.Size      = Size(total_w, MARGIN + page2_h)
        self.pnl_page2.BackColor = self.BackColor
        self.pnl_page2.Visible   = False
        self.Controls.Add(self.pnl_page2)
        y2 = MARGIN

        lbl_instr2 = Label()
        lbl_instr2.Text      = "Specify the number of sectors and tangential nodes represented by this Cut BC"
        lbl_instr2.Location  = Point(MARGIN, y2)
        lbl_instr2.Size      = Size(total_w - MARGIN * 2, INSTR_H)
        lbl_instr2.Font      = Font("Segoe UI", 9.5, FontStyle.Bold)
        lbl_instr2.ForeColor = Color.FromArgb(0, 70, 150)
        self.pnl_page2.Controls.Add(lbl_instr2)

        y2 += INSTR_H + 8

        try:
            img_sector = load_b64(self._B64_SECTOR)
            img_nodes  = load_b64(self._B64_NODES)
        except:
            img_sector = img_nodes = None

        pnl_sector_frame = _WPanel()
        pnl_sector_frame.Location  = Point(MARGIN, y2)
        pnl_sector_frame.Size      = Size(IMG_W + BORDER * 2, IMG_H + BORDER * 2)
        pnl_sector_frame.BackColor = Color.FromArgb(225, 225, 225)
        self.pnl_page2.Controls.Add(pnl_sector_frame)

        pb_sector = PictureBox()
        pb_sector.Location  = Point(BORDER, BORDER)
        pb_sector.Size      = Size(IMG_W, IMG_H)
        pb_sector.SizeMode  = PictureBoxSizeMode.Zoom
        pb_sector.BackColor = Color.White
        if img_sector: pb_sector.Image = img_sector
        pnl_sector_frame.Controls.Add(pb_sector)

        pnl_nodes_frame = _WPanel()
        pnl_nodes_frame.Location  = Point(MARGIN + IMG_W + GAP, y2)
        pnl_nodes_frame.Size      = Size(IMG_W + BORDER * 2, IMG_H + BORDER * 2)
        pnl_nodes_frame.BackColor = Color.FromArgb(225, 225, 225)
        self.pnl_page2.Controls.Add(pnl_nodes_frame)

        pb_nodes = PictureBox()
        pb_nodes.Location  = Point(BORDER, BORDER)
        pb_nodes.Size      = Size(IMG_W, IMG_H)
        pb_nodes.SizeMode  = PictureBoxSizeMode.Zoom
        pb_nodes.BackColor = Color.White
        if img_nodes: pb_nodes.Image = img_nodes
        pnl_nodes_frame.Controls.Add(pb_nodes)

        from System.Windows.Forms import NumericUpDown
        from System.Drawing import ContentAlignment
        from System.Windows.Forms import HorizontalAlignment

        y2 += IMG_H + BORDER * 2 + 8

        col_w = IMG_W + BORDER * 2
        LBL_H = 22
        INPUT_W = 90
        INPUT_H = 28

        lbl_font = Font("Segoe UI", 11.0, FontStyle.Bold)

        lbl_sec = Label()
        lbl_sec.Text      = "Number of Sectors:"
        lbl_sec.Location  = Point(MARGIN, y2)
        lbl_sec.Size      = Size(col_w, LBL_H)
        lbl_sec.Font      = lbl_font
        lbl_sec.TextAlign = ContentAlignment.MiddleCenter
        self.pnl_page2.Controls.Add(lbl_sec)

        self.txt_sector_w = TextBox()
        self.txt_sector_w.Text     = "1"
        self.txt_sector_w.Font     = Font("Segoe UI", 11.0, FontStyle.Bold)
        self.txt_sector_w.TextAlign = HorizontalAlignment.Center
        self.txt_sector_w.Size     = Size(INPUT_W, INPUT_H)
        self.txt_sector_w.Location = Point(MARGIN + (col_w - INPUT_W) // 2, y2 + LBL_H + 4)
        self.pnl_page2.Controls.Add(self.txt_sector_w)

        lbl_nod = Label()
        lbl_nod.Text      = "Tangential Node Count:"
        lbl_nod.Location  = Point(MARGIN + IMG_W + GAP, y2)
        lbl_nod.Size      = Size(col_w, LBL_H)
        lbl_nod.Font      = lbl_font
        lbl_nod.TextAlign = ContentAlignment.MiddleCenter
        self.pnl_page2.Controls.Add(lbl_nod)

        self.txt_nodes_w = NumericUpDown()
        self.txt_nodes_w.Minimum  = 1
        self.txt_nodes_w.Maximum  = 999
        self.txt_nodes_w.Value    = 1
        self.txt_nodes_w.Font     = Font("Segoe UI", 11.0, FontStyle.Bold)
        self.txt_nodes_w.TextAlign = HorizontalAlignment.Center
        self.txt_nodes_w.Size     = Size(INPUT_W, INPUT_H)
        self.txt_nodes_w.Location = Point(MARGIN + IMG_W + GAP + (col_w - INPUT_W) // 2, y2 + LBL_H + 4)
        self.pnl_page2.Controls.Add(self.txt_nodes_w)

        # =====================================================================
        # NAV: Back / Next (sayfa 1: sadece Next; sayfa 2: Back + Finish)
        # =====================================================================
        nav_y = self._nav_y_page1

        self.btn_back = Button()
        self.btn_back.Text      = "Back"
        self.btn_back.Size      = Size(100, NAV_H)
        self.btn_back.Location  = Point(MARGIN, nav_y)
        self.btn_back.FlatStyle = FlatStyle.Flat
        self.btn_back.Visible   = False
        self.btn_back.Click    += self._on_back
        self.Controls.Add(self.btn_back)

        self.btn_next = Button()
        self.btn_next.Text      = "Next"
        self.btn_next.Size      = Size(120, NAV_H)
        self.btn_next.Location  = Point(total_w - MARGIN - 120, nav_y)
        self.btn_next.BackColor = Color.FromArgb(0, 120, 215)
        self.btn_next.ForeColor = Color.White
        self.btn_next.FlatStyle = FlatStyle.Flat
        self.btn_next.Enabled   = False
        self.btn_next.Click    += self._on_next_or_finish
        self.Controls.Add(self.btn_next)

    def _select_side(self, on_3d_side):
        self.elements_on_3d_side = on_3d_side
        if on_3d_side:
            self.pnl_3d_frame.BackColor = self._FRAME_3D_ON
            self.pnl_2d_frame.BackColor = Color.FromArgb(225, 225, 225)
        else:
            self.pnl_2d_frame.BackColor = self._FRAME_2D_ON
            self.pnl_3d_frame.BackColor = Color.FromArgb(225, 225, 225)
        self.btn_next.Enabled = True

    def _on_select_3d(self, sender, args):
        self._select_side(True)

    def _on_select_2d(self, sender, args):
        self._select_side(False)

    def _on_back(self, sender, args):
        self._page = 1
        self.pnl_page2.Visible = False
        self.pnl_page1.Visible = True
        self.btn_back.Visible  = False
        self.btn_next.Text     = "Next"
        self.btn_next.Enabled  = True
        self.ClientSize = Size(self._total_w, self._page1_total_h)
        self.btn_back.Location = Point(12, self._nav_y_page1)
        self.btn_next.Location = Point(self._total_w - 12 - 120, self._nav_y_page1)

    def _on_next_or_finish(self, sender, args):
        if self._page == 1:
            if self.elements_on_3d_side is None:
                MessageBox.Show("Please select a side (3D or 2D) before continuing.", "Warning")
                return
            self._page = 2
            self.pnl_page1.Visible = False
            self.pnl_page2.Visible = True
            self.btn_back.Visible  = True
            self.btn_next.Text     = "Finish"
            self.ClientSize = Size(self._total_w, self._page2_total_h)
            self.btn_back.Location = Point(12, self._nav_y_page2)
            self.btn_next.Location = Point(self._total_w - 12 - 120, self._nav_y_page2)
        else:
            sec_txt = str(self.txt_sector_w.Text).strip()
            if "," in sec_txt:
                MessageBox.Show(
                    "Please use a dot (.) as the decimal separator, not a comma.\n"
                    "Example: 12.1",
                    "Number of Sectors", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                return
            try:
                sec_val = float(sec_txt)
            except:
                sec_val = None
            if sec_val is None or sec_val != sec_val or sec_val <= 0 or sec_val > 999999:
                MessageBox.Show(
                    "Number of Sectors must be a positive number (use a dot for decimals, e.g. 12.1).",
                    "Number of Sectors", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                return
            self.num_sectors = sec_val
            self.num_nodes   = int(self.txt_nodes_w.Value)
            self.DialogResult = DialogResult.OK
            self.Close()


class DispPanel(object):
    """GroupBox with dynamic NS rows + steps + dir, all repositioned on reflow."""
    def __init__(self, ns_names, num_steps, step_labels=None):
        self.ns_names   = ns_names
        self.num_steps  = num_steps
        self.step_labels = step_labels or {}

        self.grp = GroupBox()
        self.grp.Text  = ""
        self.grp.Width = 580

        self.ns_rows = []
        self._form  = None
        self.step_ctrls, self.step_checks, self.step_h = make_step_checks(num_steps, self.step_labels)
        self.dir_ctrls,  self.txt_dir,     self.dir_h  = make_dir_controls("Disp")

        add_controls_to_grp(self.grp, self.step_ctrls)
        add_controls_to_grp(self.grp, self.dir_ctrls)

        self._apply_btn = Button()
        self._apply_btn.Text      = "Sync Steps >>"
        self._apply_btn.Size      = Size(128, 22)
        self._apply_btn.BackColor = Color.FromArgb(230, 230, 230)
        self._apply_btn.FlatStyle = FlatStyle.Flat
        self._apply_btn.Click    += self._on_apply_to_all
        self.grp.Controls.Add(self._apply_btn)

        self._add_ns_row()
        self.reflow()

    def set_steps(self, steps):
        """Verilen step listesini bu panele programatik olarak uygular."""
        steps_set = set(steps)
        if self.step_checks:
            for cb in self.step_checks:
                cb.Checked = int(cb.Tag) in steps_set
        else:
            for ctrl, rx, ry in self.step_ctrls:
                if hasattr(ctrl, 'Tag') and isinstance(ctrl.Tag, list):
                    ctrl.Tag = sorted([s for s in steps if 1 <= s <= self.num_steps])
                    n = len(ctrl.Tag)
                    if n == 0:
                        ctrl.Text = "Select Steps...  (none selected)"
                    elif n == self.num_steps:
                        ctrl.Text = "Select Steps...  (all {0} selected)".format(self.num_steps)
                    else:
                        ctrl.Text = "Select Steps...  ({0} selected)".format(n)
                    break

    def _on_apply_to_all(self, sender, args):
        steps = get_selected_steps(self.step_checks, self.step_ctrls)
        if not steps:
            MessageBox.Show("No step is selected to apply.", "Apply to All")
            return
        frm = self._form
        if frm is None:
            return
        targets = []
        if hasattr(frm, '_disp') and frm._disp is not self and frm.chk_disp.Checked:
            targets.append(frm._disp)
        if hasattr(frm, '_force') and frm._force is not self and frm.chk_force.Checked:
            targets.append(frm._force)
        if hasattr(frm, '_temp') and frm._temp is not self and frm.chk_temp.Checked:
            targets.append(frm._temp)
        if not targets:
            MessageBox.Show("No other active panel found.\nPlease check other export types.", "Apply to All")
            return
        for p in targets:
            p.set_steps(steps)

    def _add_ns_row(self):
        idx      = len(self.ns_rows)
        can_rem  = idx > 0

        lbl = Label()
        lbl.Text = "Named Selection {0}:".format(idx + 1)
        lbl.Size = Size(150, 20)
        self.grp.Controls.Add(lbl)

        combo = ComboBox()
        combo.Size          = Size(218, 22)
        combo.DropDownStyle = ComboBoxStyle.DropDownList
        combo.Items.Add("-- Select --")
        for n in self.ns_names:
            combo.Items.Add(n)
        combo.SelectedIndex = 0
        _no_scroll_cmb(combo)
        attach_ns_name_warning(combo)
        self.grp.Controls.Add(combo)

        btn_add = Button()
        btn_add.Text      = "+"
        btn_add.Size      = Size(26, 22)
        btn_add.BackColor = Color.FromArgb(0, 120, 215)
        btn_add.ForeColor = Color.White
        btn_add.FlatStyle = FlatStyle.Flat
        btn_add.Click    += self._on_add
        self.grp.Controls.Add(btn_add)

        btn_rem = Button()
        btn_rem.Text      = "-"
        btn_rem.Size      = Size(26, 22)
        btn_rem.Enabled   = can_rem
        btn_rem.BackColor = Color.FromArgb(200, 0, 0)
        btn_rem.ForeColor = Color.White
        btn_rem.FlatStyle = FlatStyle.Flat
        btn_rem.Click    += self._on_rem
        self.grp.Controls.Add(btn_rem)

        self.ns_rows.append((lbl, combo, btn_add, btn_rem))

    def _remove_last_ns_row(self):
        if len(self.ns_rows) <= 1:
            return
        lbl, combo, btn_add, btn_rem = self.ns_rows.pop()
        for c in [lbl, combo, btn_add, btn_rem]:
            self.grp.Controls.Remove(c)

    def _on_add(self, sender, args):
        self._add_ns_row()
        self.reflow()

    def _on_rem(self, sender, args):
        self._remove_last_ns_row()
        self.reflow()

    def reflow(self):
        y = 10
        x_lbl = 8; x_combo = 165; x_plus = 512; x_minus = 541
        for i, (lbl, combo, btn_add, btn_rem) in enumerate(self.ns_rows):
            lbl.Location     = Point(x_lbl,   y + 4)
            combo.Location   = Point(x_combo, y)
            btn_add.Location = Point(x_plus,  y)
            btn_rem.Location = Point(x_minus, y)
            y += ROW_H
        y += 6
        place_controls(self.grp, self.step_ctrls, y)
        self._apply_btn.Location = Point(432, y)
        y += self.step_h + 4
        place_controls(self.grp, self.dir_ctrls, y)
        y += self.dir_h
        self.grp.Height = y + 10
        try: self._form._relayout()
        except: pass

    def refresh_steps(self, num_steps, step_labels):
        """RST degisince step kontrollerini yerinde guncelle (dialog kapanmaz)."""
        for ctrl, rx, ry in self.step_ctrls:
            try: self.grp.Controls.Remove(ctrl)
            except: pass
        self.num_steps   = num_steps
        self.step_labels = step_labels or {}
        self.step_ctrls, self.step_checks, self.step_h = make_step_checks(
            num_steps, self.step_labels)
        add_controls_to_grp(self.grp, self.step_ctrls)
        self.reflow()

    def get_selected(self):
        return [safe_str(combo.SelectedItem) if combo.SelectedIndex > 0 else ""
                for lbl, combo, btn_add, btn_rem in self.ns_rows]



class ForcePanel(object):
    """Force panel with per-row direction (Cut BC) selection."""
    def __init__(self, ns_names, num_steps, step_labels=None):
        self.ns_names   = ns_names
        self.num_steps  = num_steps
        self.step_labels = step_labels or {}

        self.grp = GroupBox()
        self.grp.Text  = ""
        self.grp.Width = 580

        self._form  = None
        self.ns_rows = []

        self.step_ctrls, self.step_checks, self.step_h = make_step_checks(num_steps, self.step_labels)
        self.dir_ctrls,  self.txt_dir,     self.dir_h  = make_dir_controls("Force")
        add_controls_to_grp(self.grp, self.step_ctrls)
        add_controls_to_grp(self.grp, self.dir_ctrls)

        self._apply_btn = Button()
        self._apply_btn.Text      = "Sync Steps >>"
        self._apply_btn.Size      = Size(128, 22)
        self._apply_btn.BackColor = Color.FromArgb(230, 230, 230)
        self._apply_btn.FlatStyle = FlatStyle.Flat
        self._apply_btn.Click    += self._on_apply_to_all
        self.grp.Controls.Add(self._apply_btn)

        # side_selections[row_index]  = True (3D taraf) / False (2D taraf)
        # sector_selections[row_index] = Number of Sectors (int)
        # node_selections[row_index]   = Tangential Node Count (int)
        self._side_selections   = {}
        self._sector_selections = {}
        self._node_selections    = {}

        self._add_ns_row()
        self.reflow()

    def set_steps(self, steps):
        """Verilen step listesini bu panele programatik olarak uygular."""
        steps_set = set(steps)
        if self.step_checks:
            for cb in self.step_checks:
                cb.Checked = int(cb.Tag) in steps_set
        else:
            for ctrl, rx, ry in self.step_ctrls:
                if hasattr(ctrl, 'Tag') and isinstance(ctrl.Tag, list):
                    ctrl.Tag = sorted([s for s in steps if 1 <= s <= self.num_steps])
                    n = len(ctrl.Tag)
                    if n == 0:
                        ctrl.Text = "Select Steps...  (none selected)"
                    elif n == self.num_steps:
                        ctrl.Text = "Select Steps...  (all {0} selected)".format(self.num_steps)
                    else:
                        ctrl.Text = "Select Steps...  ({0} selected)".format(n)
                    break

    def _on_apply_to_all(self, sender, args):
        steps = get_selected_steps(self.step_checks, self.step_ctrls)
        if not steps:
            MessageBox.Show("No step is selected to apply.", "Apply to All")
            return
        frm = self._form
        if frm is None:
            return
        targets = []
        if hasattr(frm, '_disp') and frm._disp is not self and frm.chk_disp.Checked:
            targets.append(frm._disp)
        if hasattr(frm, '_force') and frm._force is not self and frm.chk_force.Checked:
            targets.append(frm._force)
        if hasattr(frm, '_temp') and frm._temp is not self and frm.chk_temp.Checked:
            targets.append(frm._temp)
        if not targets:
            MessageBox.Show("No other active panel found.\nPlease check other export types.", "Apply to All")
            return
        for p in targets:
            p.set_steps(steps)

    def _add_ns_row(self):
        idx     = len(self.ns_rows)
        can_rem = idx > 0

        lbl_e = Label()
        lbl_e.Text = "Edge Name Selection {0}:".format(idx + 1)
        lbl_e.Size = Size(90, 20)
        self.grp.Controls.Add(lbl_e)

        combo_e = ComboBox()
        combo_e.Size          = Size(175, 22)
        combo_e.DropDownStyle = ComboBoxStyle.DropDownList
        combo_e.Items.Add("-- Select --")
        for n in self.ns_names:
            combo_e.Items.Add(n)
        combo_e.SelectedIndex = 0
        _no_scroll_cmb(combo_e)
        attach_ns_name_warning(combo_e)
        self.grp.Controls.Add(combo_e)

        lbl_el = Label()
        lbl_el.Text = "Element Name Selection {0}:".format(idx + 1)
        lbl_el.Size = Size(80, 20)
        self.grp.Controls.Add(lbl_el)

        combo_el = ComboBox()
        combo_el.Size          = Size(175, 22)
        combo_el.DropDownStyle = ComboBoxStyle.DropDownList
        combo_el.Items.Add("-- Select --")
        for n in self.ns_names:
            combo_el.Items.Add(n)
        combo_el.SelectedIndex = 0
        _no_scroll_cmb(combo_el)
        attach_ns_name_warning(combo_el)
        self.grp.Controls.Add(combo_el)

        # Direction (Cut BC) selection button - opens the dialog
        btn_dir = Button()
        btn_dir.Text      = "Select Direction..."
        btn_dir.Size      = Size(118, 24)
        btn_dir.BackColor = Color.FromArgb(240, 240, 240)
        btn_dir.Click    += self._on_dir_click
        self.grp.Controls.Add(btn_dir)

        btn_add = Button()
        btn_add.Text      = "+"
        btn_add.Size      = Size(26, 22)
        btn_add.BackColor = Color.FromArgb(0, 120, 215)
        btn_add.ForeColor = Color.White
        btn_add.FlatStyle = FlatStyle.Flat
        btn_add.Click    += self._on_add
        self.grp.Controls.Add(btn_add)

        btn_rem = Button()
        btn_rem.Text      = "-"
        btn_rem.Size      = Size(26, 22)
        btn_rem.Enabled   = can_rem
        btn_rem.BackColor = Color.FromArgb(200, 0, 0)
        btn_rem.ForeColor = Color.White
        btn_rem.FlatStyle = FlatStyle.Flat
        btn_rem.Click    += self._on_rem
        self.grp.Controls.Add(btn_rem)

        self.ns_rows.append((lbl_e, combo_e, lbl_el, combo_el, btn_dir, btn_add, btn_rem))

    def _remove_last_ns_row(self):
        if len(self.ns_rows) <= 1:
            return
        row = self.ns_rows.pop()
        idx = len(self.ns_rows)
        self._side_selections.pop(idx, None)
        self._sector_selections.pop(idx, None)
        self._node_selections.pop(idx, None)
        for c in row:
            self.grp.Controls.Remove(c)

    def _on_add(self, sender, args):
        self._add_ns_row()
        self.reflow()

    def _on_rem(self, sender, args):
        self._remove_last_ns_row()
        self.reflow()

    def _on_dir_click(self, sender, args):
        row_idx = None
        for i, row in enumerate(self.ns_rows):
            if row[4] is sender:   # btn_dir is at index 4
                row_idx = i
                break
        if row_idx is None:
            return
        dlg = ElementSideDialog()
        if dlg.ShowDialog() != DialogResult.OK:
            return
        self._side_selections[row_idx]   = dlg.elements_on_3d_side
        self._sector_selections[row_idx] = dlg.num_sectors
        self._node_selections[row_idx]   = dlg.num_nodes
        if dlg.elements_on_3d_side:
            sender.Text      = "Direction: 3D"
            sender.BackColor = Color.FromArgb(200, 220, 245)   # mavi
        else:
            sender.Text      = "Direction: 2D"
            sender.BackColor = Color.FromArgb(200, 235, 200)   # yesil

    def reflow(self):
        y = 10
        for i, (lbl_e, combo_e, lbl_el, combo_el, btn_dir, btn_add, btn_rem) in enumerate(self.ns_rows):
            lbl_e.Location    = Point(8,   y)
            lbl_e.Size        = Size(160, 18)
            combo_e.Location  = Point(8,   y + 18)
            combo_e.Size      = Size(175, 22)
            lbl_el.Location   = Point(195, y)
            lbl_el.Size       = Size(180, 18)
            combo_el.Location = Point(195, y + 18)
            combo_el.Size     = Size(175, 22)
            btn_dir.Location  = Point(382, y + 18)
            btn_add.Location  = Point(512, y + 18)
            btn_rem.Location  = Point(541, y + 18)
            y += ROW_H + 18
        y += 6

        place_controls(self.grp, self.step_ctrls, y)
        self._apply_btn.Location = Point(432, y)
        y += self.step_h + 4
        place_controls(self.grp, self.dir_ctrls, y)
        y += self.dir_h + 6
        self.grp.Height = y + 10
        try: self._form._relayout()
        except: pass

    def refresh_steps(self, num_steps, step_labels):
        """RST degisince step kontrollerini yerinde guncelle (dialog kapanmaz)."""
        for ctrl, rx, ry in self.step_ctrls:
            try: self.grp.Controls.Remove(ctrl)
            except: pass
        self.num_steps   = num_steps
        self.step_labels = step_labels or {}
        self.step_ctrls, self.step_checks, self.step_h = make_step_checks(
            num_steps, self.step_labels)
        add_controls_to_grp(self.grp, self.step_ctrls)
        self.reflow()

    def get_pairs(self):
        """Returns list of (edge_ns, elem_ns, elements_on_3d_side, num_sectors, num_nodes).
        elements_on_3d_side/num_sectors/num_nodes are None if direction wizard
        has not been completed for that row yet."""
        pairs = []
        for i, (lbl_e, combo_e, lbl_el, combo_el, btn_dir, btn_add, btn_rem) in enumerate(self.ns_rows):
            e  = safe_str(combo_e.SelectedItem)  if combo_e.SelectedIndex  > 0 else ""
            el = safe_str(combo_el.SelectedItem) if combo_el.SelectedIndex > 0 else ""
            on_3d   = self._side_selections.get(i, None)
            n_sec   = self._sector_selections.get(i, None)
            n_node  = self._node_selections.get(i, None)
            pairs.append((e, el, on_3d, n_sec, n_node))
        return pairs


class TempPanel(object):
    def __init__(self, ns_names, num_steps, step_labels=None):
        self.ns_names   = ns_names
        self.num_steps  = num_steps
        self.step_labels = step_labels or {}

        self.grp = GroupBox()
        self.grp.Text  = ""
        self.grp.Width = 580

        self._form  = None
        self.ns_rows = []

        self.step_ctrls, self.step_checks, self.step_h = make_step_checks(num_steps, self.step_labels)
        self.dir_ctrls,  self.txt_dir,     self.dir_h  = make_dir_controls("Temp")
        add_controls_to_grp(self.grp, self.step_ctrls)
        add_controls_to_grp(self.grp, self.dir_ctrls)

        self._apply_btn = Button()
        self._apply_btn.Text      = "Sync Steps >>"
        self._apply_btn.Size      = Size(128, 22)
        self._apply_btn.BackColor = Color.FromArgb(230, 230, 230)
        self._apply_btn.FlatStyle = FlatStyle.Flat
        self._apply_btn.Click    += self._on_apply_to_all
        self.grp.Controls.Add(self._apply_btn)

        self._add_ns_row()
        self.reflow()

    def set_steps(self, steps):
        """Verilen step listesini bu panele programatik olarak uygular."""
        steps_set = set(steps)
        if self.step_checks:
            for cb in self.step_checks:
                cb.Checked = int(cb.Tag) in steps_set
        else:
            for ctrl, rx, ry in self.step_ctrls:
                if hasattr(ctrl, 'Tag') and isinstance(ctrl.Tag, list):
                    ctrl.Tag = sorted([s for s in steps if 1 <= s <= self.num_steps])
                    n = len(ctrl.Tag)
                    if n == 0:
                        ctrl.Text = "Select Steps...  (none selected)"
                    elif n == self.num_steps:
                        ctrl.Text = "Select Steps...  (all {0} selected)".format(self.num_steps)
                    else:
                        ctrl.Text = "Select Steps...  ({0} selected)".format(n)
                    break

    def _on_apply_to_all(self, sender, args):
        steps = get_selected_steps(self.step_checks, self.step_ctrls)
        if not steps:
            MessageBox.Show("No step is selected to apply.", "Apply to All")
            return
        frm = self._form
        if frm is None:
            return
        targets = []
        if hasattr(frm, '_disp') and frm._disp is not self and frm.chk_disp.Checked:
            targets.append(frm._disp)
        if hasattr(frm, '_force') and frm._force is not self and frm.chk_force.Checked:
            targets.append(frm._force)
        if hasattr(frm, '_temp') and frm._temp is not self and frm.chk_temp.Checked:
            targets.append(frm._temp)
        if not targets:
            MessageBox.Show("No other active panel found.\nPlease check other export types.", "Apply to All")
            return
        for p in targets:
            p.set_steps(steps)

    def _add_ns_row(self):
        idx     = len(self.ns_rows)
        can_rem = idx > 0

        lbl = Label()
        lbl.Text = "Named Selection {0}:".format(idx + 1)
        lbl.Size = Size(155, 20)
        self.grp.Controls.Add(lbl)

        combo = ComboBox()
        combo.Size          = Size(218, 22)
        combo.DropDownStyle = ComboBoxStyle.DropDownList
        combo.Items.Add("-- Select --")
        for n in self.ns_names:
            combo.Items.Add(n)
        combo.SelectedIndex = 0
        _no_scroll_cmb(combo)
        attach_ns_name_warning(combo)
        self.grp.Controls.Add(combo)

        btn_add = Button()
        btn_add.Text      = "+"
        btn_add.Size      = Size(26, 22)
        btn_add.BackColor = Color.FromArgb(0, 120, 215)
        btn_add.ForeColor = Color.White
        btn_add.FlatStyle = FlatStyle.Flat
        btn_add.Click    += self._on_add
        self.grp.Controls.Add(btn_add)

        btn_rem = Button()
        btn_rem.Text      = "-"
        btn_rem.Size      = Size(26, 22)
        btn_rem.Enabled   = can_rem
        btn_rem.BackColor = Color.FromArgb(200, 0, 0)
        btn_rem.ForeColor = Color.White
        btn_rem.FlatStyle = FlatStyle.Flat
        btn_rem.Click    += self._on_rem
        self.grp.Controls.Add(btn_rem)

        self.ns_rows.append((lbl, combo, btn_add, btn_rem))

    def _remove_last_ns_row(self):
        if len(self.ns_rows) <= 1:
            return
        lbl, combo, btn_add, btn_rem = self.ns_rows.pop()
        for c in [lbl, combo, btn_add, btn_rem]:
            self.grp.Controls.Remove(c)

    def _on_add(self, sender, args):
        self._add_ns_row()
        self.reflow()

    def _on_rem(self, sender, args):
        self._remove_last_ns_row()
        self.reflow()

    def reflow(self):
        y = 10
        for i, (lbl, combo, btn_add, btn_rem) in enumerate(self.ns_rows):
            lbl.Location     = Point(8,   y + 4)
            combo.Location   = Point(165, y)
            btn_add.Location = Point(512, y)
            btn_rem.Location = Point(541, y)
            y += ROW_H
        y += 6
        place_controls(self.grp, self.step_ctrls, y)
        self._apply_btn.Location = Point(432, y)
        y += self.step_h + 4
        place_controls(self.grp, self.dir_ctrls, y)
        y += self.dir_h
        self.grp.Height = y + 10
        try: self._form._relayout()
        except: pass

    def refresh_steps(self, num_steps, step_labels):
        """RST degisince step kontrollerini yerinde guncelle (dialog kapanmaz)."""
        for ctrl, rx, ry in self.step_ctrls:
            try: self.grp.Controls.Remove(ctrl)
            except: pass
        self.num_steps   = num_steps
        self.step_labels = step_labels or {}
        self.step_ctrls, self.step_checks, self.step_h = make_step_checks(
            num_steps, self.step_labels)
        add_controls_to_grp(self.grp, self.step_ctrls)
        self.reflow()

    def get_selected(self):
        return [safe_str(combo.SelectedItem) if combo.SelectedIndex > 0 else ""
                for lbl, combo, btn_add, btn_rem in self.ns_rows]

# -----------------------------------------------------------------------
# PROGRESS DIALOG
# -----------------------------------------------------------------------
class ProgressDialog(Form):
    """Progress window - TopMost + DoEvents pattern."""
    def __init__(self):
        self.Text            = "Exporting..."
        self.StartPosition   = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedSingle
        self.MaximizeBox     = False
        self.MinimizeBox     = True
        self.ShowInTaskbar   = True
        self.TopMost         = True
        self.ClientSize      = Size(420, 80)

        self.lbl_stage = Label()
        self.lbl_stage.Text     = "Starting..."
        self.lbl_stage.Location = Point(15, 14)
        self.lbl_stage.Size     = Size(390, 22)
        self.lbl_stage.Font     = Font(self.Font, FontStyle.Bold)
        self.Controls.Add(self.lbl_stage)

        self.lbl_detail = Label()
        self.lbl_detail.Text      = ""
        self.lbl_detail.Location  = Point(15, 40)
        self.lbl_detail.Size      = Size(390, 22)
        self.lbl_detail.ForeColor = Color.DimGray
        self.Controls.Add(self.lbl_detail)

    def update(self, stage, detail=""):
        self.lbl_stage.Text  = stage
        self.lbl_detail.Text = detail
        Application.DoEvents()

    def done(self):
        self.Close()

# -----------------------------------------------------------------------
# MAIN COMBINED DIALOG  (orijinalle ayni + en ustte RST alani)
# -----------------------------------------------------------------------
FORM_W     = 620
CHECKBOX_X = 10

class CombinedExportDialog(Form):
    def __init__(self, ns_names, num_steps, rst_path,
                 step_to_set_id=None, step_labels=None,
                 ds=None, dpf_model=None):
        # step_to_set_id: {load_step_no: last_set_id_in_rst}
        # step_labels:    {load_step_no: "Step N (t=X.XXs)"}
        self.step_to_set_id = step_to_set_id or {i: i for i in range(1, num_steps + 1)}
        self.step_labels    = step_labels or {}
        self.ns_names  = ns_names
        self.num_steps = num_steps
        self.rst_path  = rst_path
        self.new_rst_path = None
        self.ds           = ds
        self.dpf_model    = dpf_model

        self.export_disp  = False
        self.export_force = False
        self.export_temp  = False

        self.disp_ns_list      = []
        self.disp_dir          = None
        self.disp_steps        = []

        self.force_pairs       = []
        self.force_dir         = None
        self.force_steps       = []

        self.temp_ns_list      = []
        self.temp_dir          = None
        self.temp_steps        = []

        self.Text            = "Ansys Export Tool (RST / DPF)"
        self.StartPosition   = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox     = False
        self.MinimizeBox     = True
        self.AutoScroll      = True

        self._build_ui()

    def _build_ui(self):
        # ---- RST FILE PANEL (belirgin secim alani) ----
        self.grp_rst = GroupBox()
        self.grp_rst.Text      = "  Select RST Result File  "
        self.grp_rst.Font      = Font("Segoe UI", 9.5, FontStyle.Bold)
        self.grp_rst.ForeColor = Color.FromArgb(0, 90, 180)
        self.grp_rst.Size      = Size(FORM_W - 30, 64)
        self.Controls.Add(self.grp_rst)

        self.txt_rst = TextBox()
        self.txt_rst.Size      = Size(FORM_W - 200, 24)
        self.txt_rst.ReadOnly  = True
        self.txt_rst.Font      = Font("Segoe UI", 9.0, FontStyle.Regular)
        self.txt_rst.BackColor = Color.White
        self.txt_rst.Text      = self.rst_path if self.rst_path else ""
        self.grp_rst.Controls.Add(self.txt_rst)

        self.btn_rst = Button()
        self.btn_rst.Text      = "Browse RST..."
        self.btn_rst.Size      = Size(120, 26)
        self.btn_rst.BackColor = Color.FromArgb(0, 120, 215)
        self.btn_rst.ForeColor = Color.White
        self.btn_rst.Font      = Font("Segoe UI", 9.0, FontStyle.Bold)
        self.btn_rst.FlatStyle = FlatStyle.Flat
        self.btn_rst.Click    += self._on_rst_browse
        self.grp_rst.Controls.Add(self.btn_rst)


        # ---- Orijinal kontroller ----
        self.chk_disp = CheckBox()
        self.chk_disp.Text            = "Displacement (UX / UY)"
        self.chk_disp.Size            = Size(300, 22)
        self.chk_disp.Font            = Font(self.Font, FontStyle.Bold)
        self.chk_disp.Checked         = False
        self.chk_disp.CheckedChanged += self._on_toggle
        self.Controls.Add(self.chk_disp)

        self._disp  = DispPanel(self.ns_names, self.num_steps, step_labels=self.step_labels)
        self._disp.grp.Visible = False
        self.Controls.Add(self._disp.grp)

        self.chk_force = CheckBox()
        self.chk_force.Text            = "Nodal Force (FX / FY)"
        self.chk_force.Size            = Size(300, 22)
        self.chk_force.Font            = Font(self.Font, FontStyle.Bold)
        self.chk_force.Checked         = False
        self.chk_force.CheckedChanged += self._on_toggle
        self.Controls.Add(self.chk_force)

        self._force = ForcePanel(self.ns_names, self.num_steps, step_labels=self.step_labels)
        self._force.grp.Visible = False
        self.Controls.Add(self._force.grp)

        self.chk_temp = CheckBox()
        self.chk_temp.Text            = "Temperature (BFE / Nodal)"
        self.chk_temp.Size            = Size(300, 22)
        self.chk_temp.Font            = Font(self.Font, FontStyle.Bold)
        self.chk_temp.Checked         = False
        self.chk_temp.CheckedChanged += self._on_toggle
        self.Controls.Add(self.chk_temp)

        self._temp  = TempPanel(self.ns_names, self.num_steps, step_labels=self.step_labels)
        self._temp.grp.Visible = False
        self.Controls.Add(self._temp.grp)

        self.btn_export = Button()
        self.btn_export.Text      = "Export"
        self.btn_export.Size      = Size(90, 30)
        self.btn_export.BackColor = Color.FromArgb(0, 120, 215)
        self.btn_export.ForeColor = Color.White
        self.btn_export.FlatStyle = FlatStyle.Flat
        self.btn_export.Click    += self._on_export
        self.Controls.Add(self.btn_export)

        self.btn_cancel = Button()
        self.btn_cancel.Text         = "Cancel"
        self.btn_cancel.Size         = Size(80, 30)
        self.btn_cancel.DialogResult = DialogResult.Cancel
        self.Controls.Add(self.btn_cancel)
        self.CancelButton = self.btn_cancel

        # Give each panel a reference to the form so reflow also triggers _relayout
        self._disp._form  = self
        self._force._form = self
        self._temp._form  = self

        self._relayout()

    def _on_rst_browse(self, sender, args):
        dlg = OpenFileDialog()
        dlg.Filter = "Ansys Result File (*.rst)|*.rst|All files (*.*)|*.*"
        dlg.Title  = "Select RST File"
        last = load_rst_dir()
        if last and os.path.isdir(last):
            dlg.InitialDirectory = last
        if dlg.ShowDialog() != DialogResult.OK:
            return
        new_path = dlg.FileName
        save_rst_dir(os.path.dirname(new_path))
        # RST'yi yukle (dialog kapanmaz, step panelleri yerinde guncellenir)
        try:
            new_ds, new_model = dpf_load(new_path)
        except Exception as e:
            MessageBox.Show("RST could not be read:\n{0}".format(str(e)), "Error")
            return
        # Basarili: guncelle
        self.rst_path    = new_path
        self.new_rst_path = new_path
        self.txt_rst.Text = new_path
        self.ds          = new_ds
        self.dpf_model   = new_model
        # Step bilgisini guncelle
        load_step_info = get_load_step_info(new_model)
        new_num_steps  = len(load_step_info)
        self.step_to_set_id = {ls: sid for (ls, sid, t) in load_step_info}
        new_labels = {ls: "Step {0}  (t={1:.4g}s)".format(ls, t)
                      for (ls, sid, t) in load_step_info}
        self.step_labels = new_labels
        # Panel step kontrollerini yerinde guncelle (dialog kapanmaz)
        self._disp.refresh_steps(new_num_steps, new_labels)
        self._force.refresh_steps(new_num_steps, new_labels)
        self._temp.refresh_steps(new_num_steps, new_labels)
        self._relayout()

    def _relayout(self):
        GAP = 8
        y   = 12

        # RST panel (GroupBox)
        self.grp_rst.Location = Point(CHECKBOX_X, y)
        self.txt_rst.Location = Point(12, 26)
        self.btn_rst.Location = Point(self.grp_rst.Width - 134, 25)
        y += self.grp_rst.Height + 4


        self.chk_disp.Location = Point(CHECKBOX_X, y); y += 26
        self._disp.grp.Location = Point(CHECKBOX_X + 10, y)
        y += (self._disp.grp.Height + GAP) if self._disp.grp.Visible else (GAP - 4)

        self.chk_force.Location = Point(CHECKBOX_X, y); y += 26
        self._force.grp.Location = Point(CHECKBOX_X + 10, y)
        y += (self._force.grp.Height + GAP) if self._force.grp.Visible else (GAP - 4)

        self.chk_temp.Location = Point(CHECKBOX_X, y); y += 26
        self._temp.grp.Location = Point(CHECKBOX_X + 10, y)
        y += (self._temp.grp.Height + GAP) if self._temp.grp.Visible else (GAP - 4)

        btn_y = y + 6
        self.btn_export.Location = Point(FORM_W - 200, btn_y)
        self.btn_cancel.Location = Point(FORM_W - 102, btn_y)
        self.ClientSize = Size(FORM_W, min(btn_y + 50, 800))

    def _on_toggle(self, sender, args):
        self._disp.grp.Visible  = self.chk_disp.Checked
        self._force.grp.Visible = self.chk_force.Checked
        self._temp.grp.Visible  = self.chk_temp.Checked
        self._relayout()

    def _on_export(self, sender, args):
        """
        Guvenlik sarmalayicisi: _on_export_impl icinde beklenmedik bir
        exception (ornegin Turkce/ASCII-disi karakterlerden kaynaklanan
        bir encoding sorunu) olursa, .NET'in cirkin "Unhandled exception"
        ekrani yerine temiz, anlasilir bir Ingilizce MessageBox gosterilir.
        """
        try:
            self._on_export_impl(sender, args)
        except Exception as ex:
            try:
                ex_detail = safe_str(ex)
            except:
                ex_detail = "(unable to display error detail)"
            MessageBox.Show(
                "An unexpected error occurred while preparing the export.\n\n"
                "This is often caused by Named Selection names that contain "
                "non-English characters (e.g. Turkish letters i, g, u, s, o, c "
                "with accents). Please rename any such Named Selections using "
                "only standard English letters/numbers and try again.\n\n"
                "Technical detail: {0}".format(ex_detail),
                "Export Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

    def _on_export_impl(self, sender, args):
        if not self.rst_path:
            MessageBox.Show("Please select an RST file first (use the '...' button next to 'RST File').",
                            "Warning")
            return

        if not self.chk_disp.Checked and not self.chk_force.Checked and not self.chk_temp.Checked:
            MessageBox.Show("Please select at least one export type.", "Warning")
            return

        errors = []

        if self.chk_disp.Checked:
            sel = self._disp.get_selected()
            if all(s == "" for s in sel):
                errors.append("[Displacement] At least one Named Selection must be selected.")
            if not self._disp.txt_dir.Text.strip():
                errors.append("[Displacement] Output folder is not selected.")
            if not validate_steps(self._disp.step_checks, self._disp.step_ctrls):
                errors.append("[Displacement] No steps selected.")
            for s in sel:
                if s and contains_non_ascii(s):
                    errors.append("[Displacement] Named Selection '{0}' contains non-English "
                                  "characters (e.g. Turkish letters). Please rename it using only "
                                  "standard English letters/numbers and try again.".format(s))

        if self.chk_force.Checked:
            pairs = self._force.get_pairs()
            for i, (e, el, on_3d, n_sec, n_node) in enumerate(pairs):
                if e  == "": errors.append("[Force] Edge NS {0} not selected.".format(i+1))
                if el == "": errors.append("[Force] Element NS {0} not selected.".format(i+1))
                if on_3d is None or n_sec is None or n_node is None:
                    errors.append("[Force] Direction/Sector/Node settings not completed for row {0}. "
                                  "Click 'Select Direction...' and finish the wizard.".format(i+1))
                if e and contains_non_ascii(e):
                    errors.append("[Force] Named Selection '{0}' contains non-English characters "
                                  "(e.g. Turkish letters). Please rename it using only standard "
                                  "English letters/numbers and try again.".format(e))
                if el and contains_non_ascii(el):
                    errors.append("[Force] Named Selection '{0}' contains non-English characters "
                                  "(e.g. Turkish letters). Please rename it using only standard "
                                  "English letters/numbers and try again.".format(el))
            if not self._force.txt_dir.Text.strip():
                errors.append("[Force] Output folder is not selected.")
            if not validate_steps(self._force.step_checks, self._force.step_ctrls):
                errors.append("[Force] No steps selected.")

        if self.chk_temp.Checked:
            sel = self._temp.get_selected()
            if all(s == "" for s in sel):
                errors.append("[Temperature] At least one Named Selection must be selected.")
            if not self._temp.txt_dir.Text.strip():
                errors.append("[Temperature] Output folder is not selected.")
            if not validate_steps(self._temp.step_checks, self._temp.step_ctrls):
                errors.append("[Temperature] No steps selected.")
            for s in sel:
                if s and contains_non_ascii(s):
                    errors.append("[Temperature] Named Selection '{0}' contains non-English "
                                  "characters (e.g. Turkish letters). Please rename it using only "
                                  "standard English letters/numbers and try again.".format(s))

        if errors:
            MessageBox.Show("Please fix the following:\n\n" + "\n".join(errors), "Warning")
            return

        if self.chk_disp.Checked:
            self.export_disp   = True
            self.disp_ns_list  = [s for s in self._disp.get_selected() if s != ""]
            self.disp_dir      = self._disp.txt_dir.Text.strip()
            self.disp_steps    = get_selected_steps(self._disp.step_checks, self._disp.step_ctrls)

        if self.chk_force.Checked:
            self.export_force = True
            self.force_pairs  = [(e, el, on_3d, n_sec, n_node)
                                 for e, el, on_3d, n_sec, n_node in self._force.get_pairs()
                                 if e and el and on_3d is not None
                                 and n_sec is not None and n_node is not None]
            self.force_dir    = self._force.txt_dir.Text.strip()
            self.force_steps  = get_selected_steps(self._force.step_checks, self._force.step_ctrls)

        if self.chk_temp.Checked:
            self.export_temp   = True
            self.temp_ns_list  = [s for s in self._temp.get_selected() if s != ""]
            self.temp_dir      = self._temp.txt_dir.Text.strip()
            self.temp_steps    = get_selected_steps(self._temp.step_checks, self._temp.step_ctrls)

        self.DialogResult = DialogResult.OK
        self.Close()

# -----------------------------------------------------------------------
# UNIQUE OUTPUT FILENAME HELPERS
# -----------------------------------------------------------------------
def _steps_suffix(steps):
    """
    10 veya daha az step: '1-2-5-10'
    10'dan fazla step:    'S1-65_n7'  (ilk-son_kactane)
    """
    s = sorted(steps)
    if len(s) <= 10:
        return "-".join(str(x) for x in s)
    return "S{0}-{1}_n{2}".format(s[0], s[-1], len(s))

def _make_name(prefix, parts, steps, max_len=150, ext=".txt"):
    """
    prefix: 'Disp', 'Force', 'Temp'
    parts:  NS isim parcalari listesi
    max_len: dosya adi karakter siniri (uzanti haric)
    """
    step_str = _steps_suffix(steps)
    base = "_".join([prefix] + parts)
    full = "{0}_{1}".format(base, step_str)
    if len(full) > max_len:
        available = max_len - len(base) - 1
        if available >= 6:
            s = sorted(steps)
            short = "S{0}-{1}_n{2}".format(s[0], s[-1], len(s))
            full = "{0}_{1}".format(base, short)
        if len(full) > max_len:
            full = full[:max_len]
    return full + ext

def _unique(candidate, used_paths):
    """
    Hem bu export oturumunda kullanilan yolları (used_paths) hem de
    diskte mevcut olan dosyaları kontrol eder.
    Cakisma varsa sonuna _2, _3, _4 ... ekler.
    """
    def _taken(p):
        return p in used_paths or os.path.exists(p)

    if not _taken(candidate):
        return candidate
    base, ext = os.path.splitext(candidate)
    n = 2
    while True:
        c = "{0}_{1}{2}".format(base, n, ext)
        if not _taken(c):
            return c
        n += 1

def make_disp_output_path(output_dir, ns_name, steps, used_paths):
    name = _make_name("Disp", [ns_name], steps)
    return _unique(os.path.join(output_dir, name), used_paths)

def make_force_output_path(output_dir, edge_ns, elem_ns, steps, used_paths):
    name = _make_name("Force", [edge_ns, elem_ns], steps)
    return _unique(os.path.join(output_dir, name), used_paths)

def make_temp_output_path(output_dir, ns_name, steps, used_paths):
    name = _make_name("Temp", [ns_name], steps)
    return _unique(os.path.join(output_dir, name), used_paths)

# -----------------------------------------------------------------------
# EXPORT FUNCTIONS (DPF tabanli)
# -----------------------------------------------------------------------
def run_disp_export(ds, coords, ns_name, output_dir, selected_steps, log,
                    step_to_set_id=None, output_file=None):
    """selected_steps: load step numaralari. step_to_set_id: load_step->last_set_id."""
    s2sid = step_to_set_id or {s: s for s in selected_steps}
    node_ids_list = resolve_ns_ids(ds, ns_name, False, log, coords)
    if not node_ids_list:
        MessageBox.Show("Named Selection '" + ns_name + "' node IDs could not be resolved.\n\n"
                        "If this name contains Turkish characters, please rename it in "
                        "Mechanical using only standard English letters/numbers, then try again.",
                        "Error")
        return None, 0
    node_id_set = set(node_ids_list)
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    if output_file is None:
        output_file = os.path.join(output_dir, "All_{0}_node_displacements.txt".format(ns_name))
    print("=== DISPLACEMENT (DPF): {0} ===".format(ns_name))

    step_data = {}
    for s in selected_steps:
        set_id = s2sid.get(s, s)
        dmap = get_disp_node_map(ds, set_id, node_id_set)
        step_data[s] = dmap

    node_ids = [nid for nid in node_id_set if nid in coords]
    if not node_ids:
        MessageBox.Show("No coordinates found for nodes in NS: " + ns_name, "Error")
        return None, 0
    node_coords = {nid: (coords[nid][0], coords[nid][1]) for nid in node_ids}

    header = ["NodeID", "X_Loc", "Y_Loc"]
    for s in selected_steps:
        header += ["Step{0}_UX".format(s), "Step{0}_UY".format(s)]
    node_ids = sorted(node_ids, key=lambda nid: node_coords[nid][1], reverse=True)

    with open(output_file, "w") as f:
        f.write("\t".join(header) + "\n")
        for nid in node_ids:
            cx, cy = node_coords[nid]
            row = [str(nid), fmt(cx), fmt(cy)]
            for s in selected_steps:
                d = step_data[s].get(nid, (0.0, 0.0))
                row += [fmt(d[0]), fmt(d[1])]
            f.write("\t".join(row) + "\n")

    print("DONE: " + output_file)
    return output_file, len(node_ids)


def run_force_export(ds, mesh, coords, edge_ns_name, elem_ns_name, output_dir,
                     selected_steps, log, num_sectors=1, num_nodes=1,
                     elements_on_3d_side=True, step_to_set_id=None, output_file=None):
    s2sid = step_to_set_id or {s: s for s in selected_steps}
    edge_ids = resolve_ns_ids(ds, edge_ns_name, False, log, coords)
    if not edge_ids:
        MessageBox.Show("Edge Named Selection '" + edge_ns_name + "' node IDs could not be resolved.\n\n"
                        "If this name contains Turkish characters, please rename it in "
                        "Mechanical using only standard English letters/numbers, then try again.",
                        "Error")
        return None, 0
    elem_ids = resolve_ns_ids(ds, elem_ns_name, True, log)
    if not elem_ids:
        MessageBox.Show("Element Named Selection '" + elem_ns_name + "' element IDs could not be resolved.\n\n"
                        "If this name contains Turkish characters, please rename it in "
                        "Mechanical using only standard English letters/numbers, then try again.",
                        "Error")
        return None, 0
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    if output_file is None:
        output_file = os.path.join(output_dir, "Force_{0}_{1}.txt".format(edge_ns_name, elem_ns_name))
    print("=== FORCE (DPF/ENF): {0}/{1} ===".format(edge_ns_name, elem_ns_name))

    edge_node_set = set(edge_ids)
    coord_data = {}
    for nid in edge_node_set:
        if nid in coords:
            coord_data[nid] = (coords[nid][0], coords[nid][1])
    if not coord_data:
        MessageBox.Show("No coordinates found for Edge NS nodes: " + edge_ns_name, "Error")
        return None, 0

    sign = -1.0 if elements_on_3d_side else 1.0
    print("  elements_on_3d_side={0} -> sign={1}".format(elements_on_3d_side, sign))

    step_data = {}
    for s in selected_steps:
        set_id = s2sid.get(s, s)
        enf_map = get_enf_node_map(ds, mesh, set_id, elem_ids, edge_node_set, log)
        if not enf_map:
            MessageBox.Show("No edge nodes matched in ENF data!\nSet: {0}".format(set_id), "Error")
            return None, 0
        sd = {}
        for nid in edge_node_set:
            sd[nid] = enf_map.get(nid, (0.0, 0.0))
        step_data[s] = sd

    header = ["NodeID", "X_Loc", "Y_Loc"]
    for s in selected_steps:
        header += ["Step{0}_FX".format(s), "Step{0}_FY".format(s)]
    sorted_nodes = sorted(coord_data.keys(), key=lambda nid: coord_data[nid][1], reverse=True)
    divisor = float(num_sectors) * float(num_nodes)

    with open(output_file, "w") as f:
        f.write("\t".join(header) + "\n")
        for nid in sorted_nodes:
            cx, cy = coord_data[nid]
            row = [str(nid), fmt(cx), fmt(cy)]
            for s in selected_steps:
                fx, fy = step_data[s][nid]
                row += [fmt((fx / divisor) * sign),
                        fmt((fy / divisor) * sign)]
            f.write("\t".join(row) + "\n")

    print("DONE: " + output_file)
    return output_file, len(sorted_nodes)


def run_temp_export(ds, mesh, coords, ns_name, output_dir, selected_steps, log, output_file=None):
    node_ids_list = resolve_ns_ids(ds, ns_name, False, log, coords)
    if not node_ids_list:
        MessageBox.Show("Named Selection '" + ns_name + "' node IDs could not be resolved.\n\n"
                        "If this name contains Turkish characters, please rename it in "
                        "Mechanical using only standard English letters/numbers, then try again.",
                        "Error")
        return None, 0
    node_id_set = set(node_ids_list)
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    if output_file is None:
        output_file = os.path.join(output_dir, "All_{0}_node_temperatures.txt".format(ns_name))
    print("=== TEMPERATURE (DPF/BFE): {0} ===".format(ns_name))

    step_data = {}
    for s in selected_steps:
        tmap = get_temp_node_map(ds, mesh, s, node_id_set, log)
        step_data[s] = tmap

    if not step_data[selected_steps[0]]:
        MessageBox.Show("Could not read temperature data from RST for NS: " + ns_name, "Error")
        return None, 0

    node_ids = sorted([nid for nid in node_id_set if nid in coords])
    if not node_ids:
        MessageBox.Show("No coordinates found for nodes in NS: " + ns_name, "Error")
        return None, 0
    node_coords = {nid: (coords[nid][0], coords[nid][1]) for nid in node_ids}

    header = ["NodeID", "X_Loc", "Y_Loc"] + ["Step{0}_T".format(s) for s in selected_steps]

    with open(output_file, "w") as f:
        f.write("\t".join(header) + "\n")
        for nid in node_ids:
            cx, cy = node_coords[nid]
            row = [str(nid), fmt(cx), fmt(cy)] + [fmt(step_data[s].get(nid, 0.0)) for s in selected_steps]
            f.write("\t".join(row) + "\n")

    print("DONE: " + output_file)
    return output_file, len(node_ids)

# -----------------------------------------------------------------------
# RST SELECTION
# -----------------------------------------------------------------------
def pick_rst_file():
    dlg = OpenFileDialog()
    dlg.Filter = "Ansys Result File (*.rst)|*.rst|All files (*.*)|*.*"
    dlg.Title  = "Select RST File"
    last = load_rst_dir()
    if last and os.path.isdir(last):
        dlg.InitialDirectory = last
    if dlg.ShowDialog() == DialogResult.OK:
        save_rst_dir(os.path.dirname(dlg.FileName))
        return dlg.FileName
    return None

def show_topmost_message(msg, title):
    """
    Normal MessageBox.Show bazen Mechanical penceresinin altinda acilip
    fark edilmiyor. Gecici, gorunmez, TopMost bir owner form ile
    MessageBox'i en one tasir.
    """
    owner = Form()
    try:
        owner.TopMost      = True
        owner.ShowInTaskbar = False
        owner.StartPosition = FormStartPosition.CenterScreen
        owner.Size    = Size(1, 1)
        owner.Opacity = 0
        owner.Show()
        owner.Activate()
        MessageBox.Show(owner, msg, title)
    except:
        # Herhangi bir sorunda normal davranisa geri don
        MessageBox.Show(msg, title)
    finally:
        try:
            owner.Close()
        except:
            pass

# -----------------------------------------------------------------------
# NODAL FORCE OUTPUT CONTROL KONTROLU
# -----------------------------------------------------------------------
def find_static_structural_analysis(analyses):
    """
    Returns the first Static Structural analysis in the list (thermal,
    modal, transient etc. are ignored). Returns None if none is found.
    """
    for a in analyses:
        try:
            atype = str(a.AnalysisType).strip().lower()
        except:
            atype = ""
        try:
            ptype = str(a.PhysicsType).strip().lower()
        except:
            ptype = ""
        if atype and ptype:
            if "static" in atype and "structural" in ptype:
                return a
            continue
        try:
            nm = str(a.Name).strip().lower()
        except:
            nm = ""
        if "static structural" in nm:
            return a
    return None

def check_nodal_force_output():
    """
    Analysis Settings > Output Controls > Nodal Forces ayarini okumaya
    calisir. Donen: True (Yes), False (No), None (okunamadi/belirsiz).
    Sadece kesin "No" durumunda kullanici uyarilir; okunamazsa arac
    normal akisina devam eder (surum farkliliklarinda kilitlenmesin).
    """
    try:
        analyses = list(ExtAPI.DataModel.Project.Model.Analyses)
    except:
        return None
    if not analyses:
        return None

    a = find_static_structural_analysis(analyses)
    if a is None:
        return None

    for a in [a]:
        val = None
        try:
            val = a.AnalysisSettings.NodalForces
        except:
            pass
        if val is None:
            try:
                val = a.AnalysisSettings.CalculateNodalForces
            except:
                pass
        if val is None:
            try:
                val = a.AnalysisSettings.OutputControlsNodalForces
            except:
                pass
        if val is not None:
            s = str(val).strip().lower()
            if s in ("no", "false", "0"):
                return False
            if s in ("yes", "true", "1"):
                return True
    return None

def warn_if_nodal_force_off():
    """
    Shows a warning if the Nodal Forces output control is confirmed 'No',
    and returns True (caller should terminate the script in that case).
    Returns False if undetermined/unreadable (does not block).
    """
    result = check_nodal_force_output()
    if result is False:
        MessageBox.Show(
            "The Nodal Force output control is set to 'No'.\n\n"
            "For Force export to work, before solving the analysis:\n"
            "Analysis Settings > Output Controls > Nodal Forces = Yes\n"
            "must be set, and the analysis must be re-solved with this setting.\n\n"
            "Please enable this setting, re-solve the analysis, then run this tool again.",
            "Nodal Force Output Control",
            MessageBoxButtons.OK, MessageBoxIcon.Warning)
        return True
    return False

# -----------------------------------------------------------------------
# MAIN
# -----------------------------------------------------------------------
def main():
    if warn_if_nodal_force_off():
        return

    rst_path = None   # Acilista RST sectirilmez; kullanici dialog icinden secer

    global _ns_cache, _coord_key_map
    _ns_cache = {}
    _coord_key_map = {"built": False, "grid": {}, "cell": 1.0, "rst_maxabs": 1.0}
    if True:  # (while dongusunun yerini tutar, girintiler korunuyor)

        ds = None
        dpf_model = None
        num_steps = 1

        if rst_path:
            try:
                ds, dpf_model = dpf_load(rst_path)
            except Exception as e:
                MessageBox.Show("RST could not be read:\n{0}".format(str(e)), "Error")
                rst_path = None
                ds = None
                dpf_model = None

        if dpf_model is not None:
            load_step_info = get_load_step_info(dpf_model)
            num_steps = len(load_step_info)
            # load_step_no -> last_set_id haritasi
            step_to_set_id = {ls: sid for (ls, sid, t) in load_step_info}
            # Checkbox etiketleri: "Step N (t=X.XXs)"
            step_labels = {ls: "Step {0}  (t={1:.4g}s)".format(ls, t)
                           for (ls, sid, t) in load_step_info}
        else:
            # RST yokken model analizinden tahmin et
            try:
                analyses = list(ExtAPI.DataModel.Project.Model.Analyses)
                if analyses:
                    num_steps = int(analyses[0].AnalysisSettings.NumberOfSteps)
                else:
                    num_steps = 1
            except:
                num_steps = 1
            if num_steps < 1:
                num_steps = 1
            step_to_set_id = {i: i for i in range(1, num_steps + 1)}
            step_labels = {}
            load_step_info = []

        ns_names = collect_ns_names(ds, dpf_model)
        if not ns_names:
            MessageBox.Show("No Named Selections found in the model tree.", "Error")
            return

        dlg = CombinedExportDialog(ns_names, num_steps, rst_path,
                                    step_to_set_id=step_to_set_id,
                                    step_labels=step_labels,
                                    ds=ds, dpf_model=dpf_model)
        res = dlg.ShowDialog()
        if res != DialogResult.OK:
            return
        # Dialog icinde RST yuklendi, oradan al
        rst_path  = dlg.rst_path
        ds        = dlg.ds
        dpf_model = dlg.dpf_model
        if not rst_path or ds is None:
            MessageBox.Show("Please select an RST file first.", "Error")
            return

    summary_lines = []
    error_lines   = []
    t_start       = time.time()

    prog = ProgressDialog()
    prog.Show()

    step_to_set_id = dlg.step_to_set_id  # dialog icinde RST degismis olabilir
    # Save shared output dir (use whichever was set last)
    last_dir = dlg.disp_dir or dlg.force_dir or dlg.temp_dir
    if last_dir:
        save_shared_dir(last_dir)

    # ---------------- MESH + COORDINATES (bir kere) ----------------
    log = []
    prog.update("Reading RST...", "Loading mesh and node coordinates from RST")
    mesh = get_mesh(ds, dpf_model, log)
    coords, coord_unit = build_coords_dict(ds, dpf_model, mesh, log)
    if len(coords) == 0:
        prog.done()
        MessageBox.Show("Node coordinates could not be read from RST!\n\nLog:\n" + "\n".join(log), "Error")
        return
    print("Coordinate unit: {0}".format(coord_unit))
    for line in log:
        print(line)

    # ---------------- DISPLACEMENT ----------------
    if dlg.export_disp:
        used_paths = set()
        disp_paths = []
        for ns_name in dlg.disp_ns_list:
            p = make_disp_output_path(dlg.disp_dir, ns_name, dlg.disp_steps, used_paths)
            used_paths.add(p)
            disp_paths.append(p)

        total = len(dlg.disp_ns_list)
        for i, ns_name in enumerate(dlg.disp_ns_list):
            prog.update(
                "Displacement ({0}/{1}): {2}".format(i + 1, total, ns_name),
                "Reading displacement fields from RST...")
            try:
                out_file, n = run_disp_export(ds, coords, ns_name, dlg.disp_dir,
                                              dlg.disp_steps, log,
                                              output_file=disp_paths[i])
                if out_file:
                    steps_str = ", ".join(str(s) for s in dlg.disp_steps)
                    summary_lines.append("Displacement\n  NS: {0}\n  Steps: {1}\n-> {2}".format(ns_name, steps_str, out_file))
                    prog.update(
                        "Displacement ({0}/{1}): {2}".format(i + 1, total, ns_name),
                        "Done. {0} nodes exported.".format(n))
            except Exception as ex:
                error_lines.append("Displacement [{0}] failed:\n{1}".format(ns_name, str(ex)))

    # ---------------- FORCE ----------------
    if dlg.export_force:
        used_paths  = set()
        force_paths = []
        for edge_ns, elem_ns, on_3d, n_sec, n_node in dlg.force_pairs:
            p = make_force_output_path(dlg.force_dir, edge_ns, elem_ns, dlg.force_steps, used_paths)
            used_paths.add(p)
            force_paths.append(p)

        total = len(dlg.force_pairs)
        for i, (edge_ns, elem_ns, on_3d, n_sec, n_node) in enumerate(dlg.force_pairs):
            prog.update(
                "Force ({0}/{1}): {2}".format(i + 1, total, edge_ns),
                "Reading ENF from RST and filtering edge nodes...")
            try:
                out_file, n = run_force_export(ds, mesh, coords, edge_ns, elem_ns,
                                               dlg.force_dir, dlg.force_steps, log,
                                               n_sec, n_node,
                                               elements_on_3d_side=on_3d,
                                               output_file=force_paths[i])
                if out_file:
                    steps_str = ", ".join(str(s) for s in dlg.force_steps)
                    summary_lines.append("Force\n  Edge NS: {0}\n  Elem NS: {1}\n  Sectors: {2}  Nodes: {3}\n  Steps: {4}\n-> {5}".format(
                        edge_ns, elem_ns, n_sec, n_node, steps_str, out_file))
                    prog.update(
                        "Force ({0}/{1}): {2}".format(i + 1, total, edge_ns),
                        "Done. {0} nodes exported.".format(n))
            except Exception as ex:
                error_lines.append("Force [{0}] failed:\n{1}".format(edge_ns, str(ex)))

    # ---------------- TEMPERATURE ----------------
    if dlg.export_temp:
        used_paths = set()
        temp_paths = []
        for ns_name in dlg.temp_ns_list:
            p = make_temp_output_path(dlg.temp_dir, ns_name, dlg.temp_steps, used_paths)
            used_paths.add(p)
            temp_paths.append(p)

        total = len(dlg.temp_ns_list)
        for i, ns_name in enumerate(dlg.temp_ns_list):
            prog.update(
                "Temperature ({0}/{1}): {2}".format(i + 1, total, ns_name),
                "Reading BFE temperatures from RST...")
            try:
                out_file, n = run_temp_export(ds, mesh, coords, ns_name, dlg.temp_dir,
                                              dlg.temp_steps, log,
                                              output_file=temp_paths[i])
                if out_file:
                    steps_str = ", ".join(str(s) for s in dlg.temp_steps)
                    summary_lines.append("Temperature\n  NS: {0}\n  Steps: {1}\n-> {2}".format(ns_name, steps_str, out_file))
                    prog.update(
                        "Temperature ({0}/{1}): {2}".format(i + 1, total, ns_name),
                        "Done. {0} nodes exported.".format(n))
            except Exception as ex:
                error_lines.append("Temperature [{0}] failed:\n{1}".format(ns_name, str(ex)))

    prog.done()

    duration = _fmt_duration(time.time() - t_start)
    msg = ""
    if summary_lines:
        msg += "Export complete!\n\n" + "\n\n".join(summary_lines)
    if error_lines:
        if msg: msg += "\n\n"
        msg += "ERRORS:\n" + "\n".join(error_lines)
    if msg:
        msg += "\n\n----------------------------------\nDuration: {0}".format(duration)
        show_topmost_message(msg, "Export Summary")

main()