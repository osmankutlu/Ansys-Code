# -*- coding: utf-8 -*-
# =====================================================================
#  Contact Normal Check v3.7   -   Viewer / Flip   (Ansys Mechanical 2D, IronPython 2.7)
#
#  Sayfa 1 : kontak listesi (agac)
#  Sayfa 2 : sol -> secili kontaklar (check = flip, Side = contact/target/both)
#            sag -> interaktif model gorunumu (mesh + contact/target eleman normalleri)
#  Normal yonleri mumkunse solver input (ds.dat) icindeki gercek contact/target
#  elemanlarindan okunur; okunamazsa mesh'ten tahmin edilir (ekranda yazar).
#  Goruntu: -Z'den bakis, X yukari, Y saga
#  Mouse : tekerlek = zoom (imlec uzerinde) | orta tus (Ctrl ile de) surukle = pan
#          sol surukle = pan | tikla = eleman sec/birak | Shift+surukle = kutu secim
#          "Hide body" acikken body'ye tikla = gizle (Undo / Show all)
#          Gizlenen body'nin contact/target elemanlari da gizlenir
#          Orta pencerenin solunda body agaci (Ansys tree isimleri): isaret kaldir = gizle | Ctrl+Z = Undo
#          Yesil kontak = ContactNormalFlip komutu uygulanmis ve ters ok kalmamis
# =====================================================================
import clr
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
import System
import re
import time

TITLE = u"Contact Normal Check"
FORM_ID = u"CNC_v3_7"                  # hidden window name: tells the versions apart (not shown anywhere)
VERSION = 3.7         # v3.7 = v3.6 + plain window title, capitalised texts

# ==== CORE START ====
import math
import os

CONTACT_ET = (171, 172, 173, 174)
TARGET_ET = (169, 170)
ASSIGN = re.compile(r"^([a-z_]\w*)\s*=\s*([-+]?\d+)\s*(!.*)?$")


NAME_RE = re.compile(r"^cnc_([ct])_(\d+)$", re.I)


def flag_minority(segs):
    """With solver data every element is a real contact/target element, so 'reversed' means:
    opposite to the majority direction on its side."""
    n = 0
    for x in segs:
        if x.get("cp"):
            x["rev"] = False
    for side in ("contact", "target"):
        grp = [x for x in segs if x["side"] == side and not x.get("cp")]
        rx = sum(x["nx"] for x in grp)
        ry = sum(x["ny"] for x in grp)
        if math.hypot(rx, ry) < 1e-9:
            for x in grp:
                x["rev"] = False
            continue
        for x in grp:
            x["rev"] = (x["nx"] * rx + x["ny"] * ry) < 0
            n += 1 if x["rev"] else 0
    return n


def flag_reference(segs):
    """Orange = an element whose normal is NOT the exact opposite of the other side.
    contact and target normals must face each other, so the correct direction of one side is the reverse of
    the other side. Which side is the reference (= the one that is right):
    - the side whose normals all point the same way (the more uniform side, mean normal length R closest to 1);
      the other side is then checked against it: an element is reversed when it points the SAME way as its
      nearest element on the reference side (it should point the opposite way).
      Elements of the reference side that go against its own majority are reversed as well.
    - when both sides are equally uniform and conflict (e.g. one whole side reversed) the side whose elements
      mostly point INTO their own body is the wrong one (when that can be told), otherwise both sides are
      marked and the Flip drop-down decides.
    Falls back to the plain majority rule (flag_minority) when one side is missing."""
    prim = [x for x in segs if not x.get("cp")]
    for x in segs:
        x["rev"] = False
    side = {"contact": [x for x in prim if x["side"] == "contact"],
            "target": [x for x in prim if x["side"] == "target"]}
    if not side["contact"] or not side["target"] or len(side["contact"]) * len(side["target"]) > 4000000:
        return flag_minority(segs)
    info = {}
    for sd, g in side.items():
        vx = sum(x["nx"] for x in g)
        vy = sum(x["ny"] for x in g)
        L = math.hypot(vx, vy)
        info[sd] = (L / len(g), (vx / L, vy / L) if L > 1e-9 else None)
    ref = None
    if abs(info["contact"][0] - info["target"][0]) > 0.05:
        ref = "contact" if info["contact"][0] > info["target"][0] else "target"
    only = None
    if ref is None:
        wrong = []
        for sd, g in side.items():
            known = [x for x in g if x.get("inward") is not None]
            if known and sum(1 for x in known if x["inward"]) > 0.5 * len(known):
                wrong.append(sd)
        if len(wrong) == 1:
            only = wrong[0]
    other = {"contact": "target", "target": "contact"}

    def partners(sd):
        g = side[sd]
        R, d = info[sd]
        if sd == ref and d is not None and R > 0.5:
            good = [x for x in g if x["nx"] * d[0] + x["ny"] * d[1] >= 0]
            if good:
                return good
        return g

    for sd in ("contact", "target"):
        R, d = info[sd]
        if sd == ref:
            if d is not None and R > 0.5:
                for x in side[sd]:
                    if x["nx"] * d[0] + x["ny"] * d[1] < 0:
                        x["rev"] = True
            continue
        if only is not None and sd != only:
            continue
        cand = partners(other[sd])
        for x in side[sd]:
            best, bd = None, None
            for y in cand:
                dd = (x["mx"] - y["mx"]) ** 2 + (x["my"] - y["my"]) ** 2
                if bd is None or dd < bd:
                    bd, best = dd, y
            if x["nx"] * best["nx"] + x["ny"] * best["ny"] > 0.5:
                x["rev"] = True
    return sum(1 for x in prim if x["rev"])


def edge_owners(bodies, nodes):
    """(lo, hi) -> [(body, a, c)] for every element edge a->c (in element node order) on `nodes`."""
    out = {}
    for b in bodies:
        for e in b["elems"]:
            corner = e[1]
            n = len(corner)
            for k in range(n):
                a, c = corner[k], corner[(k + 1) % n]
                if a != c and a in nodes and c in nodes:
                    out.setdefault((min(a, c), max(a, c)), []).append((b, a, c))
    return out


def direction_counts(segs):
    """Counts of arrows per side by screen direction (view from -Z: X up, Y right)."""
    res = {"contact": {}, "target": {}}
    for sg in segs:
        if sg.get("cp"):
            continue
        d = res[sg["side"]]
        k = dir_label(sg)
        d[k] = d.get(k, 0) + 1
    return res


def dat_segments(recs, xy, node2bodies, dummy, edge_map=None):
    """Segments straight from the solver elements: direction = node order of the element."""
    segs = []
    for side in ("contact", "target"):
        for rec in recs.get(side, []):
            i, j, enum = rec[0], rec[1], rec[2]
            pa, pb = xy.get(i), xy.get(j)
            if pa is None or pb is None:
                continue
            tx, ty = pb[0] - pa[0], pb[1] - pa[1]
            L = math.hypot(tx, ty)
            if L == 0:
                continue
            bi, bj = node2bodies.get(i, []), node2bodies.get(j, [])
            common = [b for b in bi if b in bj] or (bi + bj) or [dummy]
            inward = None                       # True: normal points into the body it sits on; None: cannot tell
            if edge_map:
                cands = edge_map.get((min(i, j), max(i, j)))
                if cands:
                    same = [b for (b, a, c) in cands if a == i and c == j]
                    common = [same[0] if same else cands[0][0]]   # the element this solver element sits on
                    if len(cands) == 1:         # a boundary edge of exactly one solid element: its direction decides
                        inward = not same
            third = rec[4] if len(rec) > 4 else None
            segs.append({"side": side, "eid": enum, "enum": enum, "a": i, "b": j, "m": third,
                         "ax": pa[0], "ay": pa[1], "bx": pb[0], "by": pb[1],
                         "mx": (pa[0] + pb[0]) / 2.0, "my": (pa[1] + pb[1]) / 2.0,
                         "ex": (pa[0] + pb[0]) / 2.0, "ey": (pa[1] + pb[1]) / 2.0,
                         "nx": ty / L, "ny": -tx / L, "len": L, "rev": False, "known": True,
                         "body": common[0], "bodies": common, "inward": inward})
    return segs


def inward_info(segs):
    """Information only (does not change the orange marking): how many primary elements of each side have a
    normal pointing INTO the body they sit on. Only returned when the solid-element orientation can be trusted
    (at least one side has >= 90 % of its elements pointing outwards), otherwise None."""
    prim = [x for x in segs if not x.get("cp") and x.get("inward") is not None]
    res, ok = {}, False
    for side in ("contact", "target"):
        g = [x for x in prim if x["side"] == side]
        res[side] = sum(1 for x in g if x["inward"])
        if g and sum(1 for x in g if not x["inward"]) >= 0.9 * len(g):
            ok = True
    return res if ok else None


def seg_visible(sg):
    b = sg.get("bodies")
    return any(x["visible"] for x in b) if b else sg["body"]["visible"]


def dump_snippet(key):
    """TEMPORARY command snippet for a contact: writes the contact (type cid) and target (type tid)
    elements of THIS contact (element number, node numbers, coordinates of the first two nodes)
    into cnc_<key>.txt in the solver working directory. The last line has tag -1 (= file complete)."""
    fmt = "(5(1X,F10.0),4(1X,E20.10),3(1X,F8.0))"
    vw = "*vwrite,cnc_k,cnc_e,cnc_n1,cnc_n2,cnc_n3,cnc_x1,cnc_y1,cnc_x2,cnc_y2,cnc_ty,cnc_rl,cnc_mt"
    L = ["! ContactNormalCheck - TEMPORARY command (removed automatically)",
         "*cfclos",
         "allsel,all",
         "cnc_c = cid",
         "cnc_t = tid",
         "*cfopen,cnc_%d,txt" % key,
         "cnc_k = 0",                                   # header line (tag 0): contact id / target id of this contact
         "cnc_e = 0",
         "cnc_n1 = cnc_c",
         "cnc_n2 = cnc_t",
         "cnc_n3 = 0",
         "cnc_x1 = 0",
         "cnc_y1 = 0",
         "cnc_x2 = 0",
         "cnc_y2 = 0",
         "cnc_ty = 0",
         "cnc_rl = 0",
         "cnc_mt = 0",
         vw, fmt,
         "*do,cnc_k,1,2",
         "*if,cnc_k,eq,1,then",
         "esel,s,type,,cnc_c",
         "*else",
         "esel,s,type,,cnc_t",
         "*endif",
         "*get,cnc_ne,elem,0,count",
         "cnc_e = 0",
         "*do,cnc_i,1,cnc_ne",
         "cnc_e = elnext(cnc_e)",
         "cnc_n1 = nelem(cnc_e,1)",
         "cnc_n2 = nelem(cnc_e,2)",
         "cnc_n3 = nelem(cnc_e,3)",
         "cnc_x1 = nx(cnc_n1)",
         "cnc_y1 = ny(cnc_n1)",
         "cnc_x2 = nx(cnc_n2)",
         "cnc_y2 = ny(cnc_n2)",
         "*get,cnc_ty,elem,cnc_e,attr,type",
         "*get,cnc_rl,elem,cnc_e,attr,real",
         "*get,cnc_mt,elem,cnc_e,attr,mat",
         vw, fmt,
         "*enddo",
         "*enddo",
         "cnc_k = -1",
         "cnc_e = 0",
         "cnc_n1 = 0",
         "cnc_n2 = 0",
         "cnc_n3 = 0",
         "cnc_x1 = 0",
         "cnc_y1 = 0",
         "cnc_x2 = 0",
         "cnc_y2 = 0",
         "cnc_ty = 0",
         "cnc_rl = 0",
         "cnc_mt = 0",
         vw, fmt,
         "*cfclos",
         "allsel,all"]
    return "\n".join(L) + "\n"


EXIT_TEXT = ("! ContactNormalCheck - TEMPORARY command: ends the solver run right after the model has been read\n"
             "/exit,nosave\n")


def parse_dump(lines):
    """-> ({'contact': [items], 'target': [items]}, complete)  item = {'e', 'nodes', 'p1', 'p2'}"""
    out = {"contact": [], "target": []}
    done = False
    for l in lines:
        t = l.split()
        if len(t) < 9:
            continue
        try:
            v = [float(x) for x in t[:9]]
        except ValueError:
            continue
        tag = int(round(v[0]))
        if tag < 0:
            done = True
            continue
        if tag == 0:                                   # header: contact id, target id
            out["hdr"] = (int(round(v[2])), int(round(v[3])))
            continue
        extra = [int(round(float(x))) for x in t[9:12]] if len(t) >= 12 else [0, 0, 0]
        try:
            ty, rl, mt = extra
        except ValueError:
            ty = rl = mt = 0
        out["contact" if tag == 1 else "target"].append(
            {"e": int(round(v[1])),
             "nodes": [int(round(v[2])), int(round(v[3])), int(round(v[4]))],
             "p1": (v[5], v[6]), "p2": (v[7], v[8]), "type": ty, "real": rl, "mat": mt})
    return out, done


def map_dump_nodes(dump_nodes, mesh_xy):
    """Solver node id -> mesh node id by coordinates, inside ONE candidate set (the nodes of the edges of
    that side), so coincident nodes of other edges cannot be mixed up. The unit scale between the solver
    and the Mechanical API is detected automatically. Returns (map, scale)."""
    if not dump_nodes or not mesh_xy:
        return {}, None
    xs = [p[0] for p in mesh_xy.values()]
    ys = [p[1] for p in mesh_xy.values()]
    ext = max(max(xs) - min(xs), max(ys) - min(ys), 1e-30)
    tol = ext * 1e-5
    cell = tol * 2.0
    grid = {}
    for mid, (x, y) in mesh_xy.items():
        grid.setdefault((int(math.floor(x / cell)), int(math.floor(y / cell))), []).append((mid, x, y))
    best, best_sc = {}, None
    for sc in (1.0, 1e-3, 1e3, 1e-6, 1e6, 0.0254, 39.3701):
        m = {}
        for sid, (x, y) in dump_nodes.items():
            X, Y = x * sc, y * sc
            cx, cy = int(math.floor(X / cell)), int(math.floor(Y / cell))
            bd, bm = None, None
            for dx in (-1, 0, 1):
                for dy in (-1, 0, 1):
                    for (mid, mx, my) in grid.get((cx + dx, cy + dy), ()):
                        d = math.hypot(mx - X, my - Y)
                        if d <= tol and (bd is None or d < bd):
                            bd, bm = d, mid
            if bm is not None:
                m[sid] = bm
        if len(m) > len(best):
            best, best_sc = m, sc
        if len(m) >= 0.98 * len(dump_nodes):
            break
    return best, best_sc


def dump_to_recs(items, xy_side, xy_union):
    """Dump items of one side -> solver-record tuples with MESH node ids; returns (recs, nodes, mapped, scale)."""
    nodes = {}
    for it in items:
        nodes[it["nodes"][0]] = it["p1"]
        nodes[it["nodes"][1]] = it["p2"]
    m, sc = map_dump_nodes(nodes, xy_side)
    miss = dict((n, p) for n, p in nodes.items() if n not in m)
    if miss:
        m2, sc2 = map_dump_nodes(miss, xy_union)
        m.update(m2)
        sc = sc or sc2
    recs = []
    for it in items:
        a, b = it["nodes"][0], it["nodes"][1]
        if a in m and b in m and m[a] != m[b]:
            recs.append((m[a], m[b], it["e"], 3 if it["nodes"][2] else 2, None))
    return recs, len(nodes), len(m), sc, m


def number_runs(nums):
    """Sorted element numbers -> list of runs of consecutive numbers."""
    runs = []
    for n in sorted(nums):
        if runs and n == runs[-1][-1] + 1:
            runs[-1].append(n)
        else:
            runs.append([n])
    return runs


def split_symmetric(recs_by_side, S, learn=None):
    """A symmetric contact has TWO pairs: contact + target on the scoped edges (the pair Ansys draws) and a
    counterpart pair with the roles swapped (contact elements on the target edge, target elements on the
    contact edge). Returns (element numbers of the counterpart pair, text describing how it was decided).
    Each pair is written as a block of consecutive element numbers, so per element type:
    1. two runs of consecutive numbers: the run lying on this side's own edge nodes is the primary pair.
    2. one run (e.g. all contact elements written back to back): the other type's two runs give the sizes
       (primary contact count = counterpart target count and vice versa), so the run is cut into primary
       and counterpart parts; which end is primary follows from the edge nodes.
    3. when the edge nodes cannot tell (shared edges) the pattern learned from the other contacts decides
       (`learn`, filled while the other contacts are analysed).
    4. otherwise: twins on one node pair -> the lowest number is primary; an element lying only on the
       other side's edge nodes is a counterpart element."""
    if learn is None:
        learn = {}
    cp = set()
    info = {}
    sizes = {}
    own_of = {"contact": (S["contact"], S["target"]), "target": (S["target"], S["contact"])}
    runs_of = {}
    nodes_of = {}
    for side in ("contact", "target"):
        recs = recs_by_side.get(side, [])
        runs_of[side] = number_runs([r[2] for r in recs])
        nodes_of[side] = dict((r[2], (r[0], r[1])) for r in recs)

    def score(side, elems):
        own, other = own_of[side]
        sc = 0
        for e in elems:
            a_, b_ = nodes_of[side][e]
            on_own = a_ in own and b_ in own
            on_oth = a_ in other and b_ in other
            if on_own and not on_oth:
                sc += 1
            elif on_oth and not on_own:
                sc -= 1
        return sc

    def label(side):
        return "%s: runs %s" % (side, ", ".join("%d-%d (%d)" % (r[0], r[-1], len(r)) for r in runs_of[side][:6]) or "-")

    # step 0: the solver's own real constants. The two pairs of a symmetric contact carry different real
    # constant sets, so each element type splits into two groups; the edge nodes say which group is the
    # scoped (primary) pair, and where they cannot tell (shared nodes) the other type / the contact id decides.
    meta = recs_by_side.get("meta") or {}
    real_of = meta.get("real") or {}
    hdr = meta.get("hdr")
    groups_of = {}
    for side in ("contact", "target"):
        g = {}
        for r in recs_by_side.get(side, []):
            g.setdefault(real_of.get(r[2], 0), []).append(r[2])
        if len(g) == 2 and 0 not in g:
            groups_of[side] = sorted(g.items())
    pair_real = None
    for rnd in (0, 1):                                         # round 0: edge nodes decide, round 1: the rest
        for side in ("contact", "target"):
            if side in sizes or side not in groups_of:
                continue
            (ka, ea), (kb, eb) = groups_of[side]
            sa, sb = score(side, ea), score(side, eb)
            if rnd == 0 and sa == sb:
                continue
            if sa != sb:
                prim = ka if sa > sb else kb
                how = "by real constant (%d / %d), edge nodes decide" % (ka, kb)
                learn[side + "_real_low"] = (prim == ka)
            elif pair_real in (ka, kb):
                prim = pair_real
                how = "by real constant (%d / %d), same pair as the other type" % (ka, kb)
            elif hdr and hdr[0] in (ka, kb):
                prim = hdr[0]
                how = "by real constant (%d / %d), edge nodes cannot tell: contact id" % (ka, kb)
            else:
                prim = ka if learn.get(side + "_real_low", True) else kb
                how = "by real constant (%d / %d), edge nodes cannot tell: learned pattern" % (ka, kb)
            if pair_real is None:
                pair_real = prim
            pe, ce = (ea, eb) if prim == ka else (eb, ea)
            cp.update(ce)
            sizes[side] = (len(pe), len(ce))
            info[side] = label(side) + " -> " + how
    for side in ("contact", "target"):                         # step 1: two runs
        runs = runs_of[side]
        if side in sizes or len(runs) != 2:
            continue
        sc = [score(side, r) for r in runs]
        key = side + "_cp_run"
        if sc[0] != sc[1]:
            cpi = 0 if sc[0] < sc[1] else 1
            learn[key] = cpi
            how = "by numbering"
        else:
            cpi = learn.get(key, 0)
            how = "by numbering, shared edge: learned pattern"
        cp.update(runs[cpi])
        sizes[side] = (len(runs[1 - cpi]), len(runs[cpi]))
        info[side] = label(side) + " -> " + how
    for side in ("contact", "target"):                         # step 1b: several runs, one vote per run
        runs = runs_of[side]
        if side in sizes or len(runs) < 3:
            continue
        votes = [score(side, r) for r in runs]
        if any(v == 0 for v in votes):
            continue
        for r, v in zip(runs, votes):
            if v < 0:
                cp.update(r)
        sizes[side] = (sum(len(r) for r, v in zip(runs, votes) if v > 0),
                       sum(len(r) for r, v in zip(runs, votes) if v < 0))
        info[side] = label(side) + " -> by numbering, one vote per run"
    for side, other in (("contact", "target"), ("target", "contact")):      # step 2: one run
        if side in sizes:
            continue
        runs = runs_of[side]
        if len(runs) == 1 and other in sizes and sizes[other][0] + sizes[other][1] == len(runs[0]):
            ps = sizes[other][1]                               # primary here = counterpart count of the other type
            run = runs[0]
            first, last = run[:ps], run[len(run) - ps:]
            sf, sl = score(side, first), score(side, last)
            key = side + "_primary_last"
            if sf != sl:
                prim_last = sl > sf
                learn[key] = prim_last
                how = "by the other type's run sizes"
            else:
                prim_last = learn.get(key, True)
                how = "by the other type's run sizes, shared edge: learned pattern"
            cp.update(run[:len(run) - ps] if prim_last else run[ps:])
            sizes[side] = (ps, len(run) - ps)
            info[side] = label(side) + " -> " + how
    for side in ("contact", "target"):                         # step 4: fallback
        if side in sizes:
            continue
        own, other = own_of[side]
        by_pair = {}
        for r in recs_by_side.get(side, []):
            by_pair.setdefault((min(r[0], r[1]), max(r[0], r[1])), []).append(r)
        for key_, lst in by_pair.items():
            lst.sort(key=lambda r: r[2])
            for r in lst[1:]:
                cp.add(r[2])
            r0 = lst[0]
            if (r0[0] in other and r0[1] in other) and not (r0[0] in own and r0[1] in own):
                cp.add(r0[2])
        info[side] = label(side) + " -> by node position"
    return cp, "; ".join(info[k] for k in ("contact", "target") if k in info)


def dir_label(sg):
    """Screen direction of an arrow (view from -Z: X up, Y right)."""
    dx, dy = sg["ny"], -sg["nx"]
    if abs(dx) >= abs(dy):
        return "right" if dx > 0 else "left"
    return "down" if dy > 0 else "up"


def describe_dump(label, items, maps, segs, cp=()):
    """Readable table of what the solver wrote for one contact (for checking / sending)."""
    dir_of = dict(((sg["side"], sg["enum"]), dir_label(sg)) for sg in segs)
    L = [label, "",
         "side     element  type  real   mat | solver nodes (n1 n2 n3)   | mesh nodes     | arrow  | pair"]
    for side in ("contact", "target"):
        m = maps.get(side, {})
        for it in items.get(side, []):
            n = it["nodes"]
            L.append("%-8s %7d %5d %5d %5d | %7d %7d %7d | %6s %6s | %-6s | %s" % (
                side, it["e"], it.get("type", 0), it.get("real", 0), it.get("mat", 0), n[0], n[1], n[2],
                m.get(n[0], "-"), m.get(n[1], "-"), dir_of.get((side, it["e"]), "(none)"),
                "counterpart" if it["e"] in cp else "primary"))
    return "\n".join(L)


def build_flip_whole(name, side):
    L = ["! ContactNormalCheck - normal flip (ESURF,,REVERSE)",
         "! Contact: %s" % name,
         "! FLIP whole side=%s" % side]
    if side in ("both", "contact"):
        L += ["esel,s,type,,cid", "esurf,,reverse"]
    if side in ("both", "target"):
        L += ["esel,s,type,,tid", "esurf,,reverse"]
    L.append("allsel,all")
    return "\n".join(L) + "\n"


def ranges(nums):
    nums = sorted(set(nums))
    out, i = [], 0
    while i < len(nums):
        j = i
        while j + 1 < len(nums) and nums[j + 1] == nums[j] + 1:
            j += 1
        out.append((nums[i], nums[j]))
        i = j + 1
    return out


def build_flip_elems(name, segs):
    """Flip only the given segments with ESURF,,REVERSE. Elements are selected by ELEMENT NUMBER
    (from the solver input); segments without an element number fall back to node-based selection."""
    L = ["! ContactNormalCheck - normal flip of selected elements (ESURF,,REVERSE)",
         "! Contact: %s" % name,
         "! If the mesh changes, re-run the tool and update this command."]
    for kind, typ in (("contact", "cid"), ("target", "tid")):
        group = [x for x in segs if x["side"] == kind]
        if not group:
            continue
        nums = [x["enum"] for x in group if x.get("enum")]
        by_nodes = [x for x in group if not x.get("enum")]
        L.append("! FLIP elems side=%s n=%d (by element number: %d, by nodes: %d)"
                 % (kind, len(group), len(nums), len(by_nodes)))
        L.append("esel,none")
        for a, b in ranges(nums):
            L.append("esel,a,elem,,%d,%d,1" % (a, b))
        for x in by_nodes:
            L.append("nsel,s,node,,%d" % x["a"])
            L.append("nsel,a,node,,%d" % x["b"])
            if x["m"]:
                L.append("nsel,a,node,,%d" % x["m"])
            L.append("esln,a,1")
        if by_nodes:
            L.append("esel,r,type,,%s" % typ)       # node based selection needs the type filter
        L.append("esurf,,reverse")
        L.append("allsel,all")
    return "\n".join(L) + "\n"


def corner_of(alln, api_corner=None):
    n = len(alln)
    if n in (3, 4):
        return list(alln)
    if n == 6:
        return list(alln[:3])
    if n == 8:
        return list(alln[:4])
    return list(api_corner) if api_corner else list(alln)


def point_in_poly(x, y, pts):
    inside = False
    n = len(pts)
    j = n - 1
    for i in range(n):
        xi, yi = pts[i]
        xj, yj = pts[j]
        if (yi > y) != (yj > y) and x < (xj - xi) * (y - yi) / (yj - yi) + xi:
            inside = not inside
        j = i
    return inside


def zoom_about(px, py, mx, my, W, H, r):
    """New pan offset so the point under (mx, my) stays fixed when the scale is multiplied by r."""
    ox, oy = mx - W / 2.0, my - H / 2.0
    return ox - (ox - px) * r, oy - (oy - py) * r


def arrow_px(median_seg_px):
    return max(16.0, min(44.0, 0.8 * median_seg_px))
def dump_all_snippet(chunk=2500):
    """ONE temporary command for the whole model (placed on the analysis): writes every contact (CONTA171-174)
    and target (TARGE169/170) element of ALL contacts and ends the solver run.
    The output goes into small files cnc_all1.txt, cnc_all2.txt, ... (a new file after `chunk` lines): a single file
    stopped at 2 MiB (12408 lines) on the first try.
    Tag 0 = header (first line: number of contact elements, number of target elements), 1 = contact element,
    2 = target element, -1 = end marker (number of elements written, number of files)."""
    fmt = "(5(1X,F10.0),4(1X,E20.10),3(1X,F8.0))"
    vw = "*vwrite,cnc_k,cnc_e,cnc_n1,cnc_n2,cnc_n3,cnc_x1,cnc_y1,cnc_x2,cnc_y2,cnc_ty,cnc_rl,cnc_mt"
    zero = ["cnc_n3 = 0", "cnc_x1 = 0", "cnc_y1 = 0", "cnc_x2 = 0", "cnc_y2 = 0",
            "cnc_ty = 0", "cnc_rl = 0", "cnc_mt = 0"]
    L = ["! ContactNormalCheck - TEMPORARY command (removed automatically)",
         "*cfclos",
         "allsel,all",
         "esel,s,ename,,171,174",
         "*get,cnc_n1c,elem,0,count",
         "esel,s,ename,,169,170",
         "*get,cnc_n2c,elem,0,count",
         "cnc_f = 1",
         "cnc_cnt = 0",
         "*cfopen,cnc_all%cnc_f%,txt",
         "cnc_k = 0",
         "cnc_e = cnc_n1c",
         "cnc_n1 = cnc_n2c",
         "cnc_n2 = 0"] + zero + [vw, fmt,
         "*do,cnc_k,1,2",
         "*if,cnc_k,eq,1,then",
         "esel,s,ename,,171,174",
         "*else",
         "esel,s,ename,,169,170",
         "*endif",
         "*get,cnc_ne,elem,0,count",
         "cnc_e = 0",
         "*if,cnc_ne,gt,0,then",
         "*do,cnc_i,1,cnc_ne",
         "cnc_e = elnext(cnc_e)",
         "cnc_n1 = nelem(cnc_e,1)",
         "cnc_n2 = nelem(cnc_e,2)",
         "cnc_n3 = nelem(cnc_e,3)",
         "cnc_x1 = 0",                              # single-node elements (pilot node targets) have no node 2;
         "cnc_y1 = 0",                              # nx(0) is an ERROR and stopped the whole command
         "cnc_x2 = 0",
         "cnc_y2 = 0",
         "*if,cnc_n1,gt,0,then",
         "cnc_x1 = nx(cnc_n1)",
         "cnc_y1 = ny(cnc_n1)",
         "*endif",
         "*if,cnc_n2,gt,0,then",
         "cnc_x2 = nx(cnc_n2)",
         "cnc_y2 = ny(cnc_n2)",
         "*endif",
         "*get,cnc_ty,elem,cnc_e,attr,type",
         "*get,cnc_rl,elem,cnc_e,attr,real",
         "*get,cnc_mt,elem,cnc_e,attr,mat",
         vw, fmt,
         "cnc_cnt = cnc_cnt + 1",
         "*if,cnc_cnt,ge,%d,then" % chunk,
         "*cfclos",
         "cnc_f = cnc_f + 1",
         "cnc_cnt = 0",
         "*cfopen,cnc_all%cnc_f%,txt",
         "*endif",
         "*enddo",
         "*endif",
         "*enddo",
         "cnc_k = -1",
         "cnc_e = cnc_n1c + cnc_n2c",
         "cnc_n1 = cnc_f",
         "cnc_n2 = 0"] + zero + [vw, fmt,
         "*cfclos",
         "allsel,all"]
    return "\n".join(L) + "\n"


EXIT_ALL_TEXT = ("! ContactNormalCheck - TEMPORARY command: closes the output file and ends the solver run\n"
                 "*cfclos\n"
                 "/exit,nosave\n")


def parse_dump_all(lines):
    """-> (items, header, complete, single). header = (contact elements, target elements) announced by the solver or
    None; complete = the end marker was seen; single = number of single-node elements (pilot node targets), which
    are counted but not returned. item = {'side', 'e', 'nodes', 'p1', 'p2', 'type', 'real', 'mat'}"""
    items, hdr, done, single = [], None, False, 0
    for l in lines:
        t = l.split()
        if len(t) < 12:
            continue
        try:
            v = [float(x) for x in t[:12]]
        except ValueError:
            continue
        tag = int(round(v[0]))
        if tag < 0:
            done = True
            continue
        if tag == 0:
            hdr = (int(round(v[1])), int(round(v[2])))
            continue
        if tag not in (1, 2):
            continue
        if int(round(v[3])) == 0:                      # no second node: a pilot node, not a line element
            single += 1
            continue
        items.append({"side": "contact" if tag == 1 else "target", "e": int(round(v[1])),
                      "nodes": [int(round(v[2])), int(round(v[3])), int(round(v[4]))],
                      "p1": (v[5], v[6]), "p2": (v[7], v[8]),
                      "type": int(round(v[9])), "real": int(round(v[10])), "mat": int(round(v[11]))})
    return items, hdr, done, single


def group_contacts(items):
    """Elements of ALL contacts -> {(contact type, target type): {'contact': [items], 'target': [items]}}.
    A contact pair = the contact and target elements that share one real constant set; the two pairs of a
    symmetric contact use the same two element types, so one group = one contact region."""
    by_real = {}
    for it in items:
        d = by_real.setdefault(it["real"], {"contact": set(), "target": set()})
        d[it["side"]].add(it["type"])
    key_of = {}
    for r, d in by_real.items():
        key_of[r] = (min(d["contact"]) if d["contact"] else 0, min(d["target"]) if d["target"] else 0)
    groups = {}
    for it in items:
        g = groups.setdefault(key_of[it["real"]], {"contact": [], "target": []})
        g[it["side"]].append(it)
    return groups


def match_groups(groups, scopes, xy):
    """Assign the element groups of the solver (one per contact region) to the selected contacts by POSITION.
    groups : group_contacts() result
    scopes : {contact key: {'contact': set(mesh nodes of the source edges), 'target': set(... target edges)}}
    xy     : mesh node -> (x, y)
    A group belongs to the contact whose edge nodes carry its element nodes. Score: fraction of the group's element
    nodes that lie on the contact's edges (>= 0.9 needed), then the fraction that lies on the expected side
    (primary pair: contact elements on the source edges, target elements on the target edges; the primary pair
    has real = contact element type), then the smaller contact (fits tighter). Every contact and every group is
    used once. Returns (assign {contact key: group key}, info)."""
    info = {"scale": None, "unmatched": [], "candidates": 0}
    if not groups or not scopes or not xy:
        return {}, info
    xs = [p[0] for p in xy.values()]
    ys = [p[1] for p in xy.values()]
    ext = max(max(xs) - min(xs), max(ys) - min(ys), 1e-30)
    tol = ext * 1e-5
    cell = tol * 2.0
    sample = {}
    for g in groups.values():
        for side in ("contact", "target"):
            for it in g[side][:20]:
                sample[it["nodes"][0]] = it["p1"]
        if len(sample) > 300:
            break
    m, scale = map_dump_nodes(sample, xy)
    if scale is None:
        scale = 1.0
    info["scale"] = scale
    index = {}
    size = {}
    for ck, sc in scopes.items():
        size[ck] = len(sc["contact"]) + len(sc["target"])
        for role, nodes in ((0, sc["contact"]), (1, sc["target"])):
            for n in nodes:
                p = xy.get(n)
                if p is None:
                    continue
                index.setdefault((int(math.floor(p[0] / cell)), int(math.floor(p[1] / cell))), []).append(
                    (ck, role, p[0], p[1]))

    def lookup(X, Y):
        cx, cy = int(math.floor(X / cell)), int(math.floor(Y / cell))
        out = []
        for dx in (-1, 0, 1):
            for dy in (-1, 0, 1):
                for (ck, role, x, y) in index.get((cx + dx, cy + dy), ()):
                    if abs(x - X) <= tol and abs(y - Y) <= tol:
                        out.append((ck, role))
        return out

    cands = []
    for gk, g in groups.items():
        n = 0
        any_h, role_h = {}, {}
        for side, want in (("contact", 0), ("target", 1)):
            for it in g[side]:
                prim = (it["real"] == gk[0])
                for (X, Y) in (it["p1"], it["p2"]):
                    n += 1
                    seen_any, seen_role = set(), set()
                    for ck, role in lookup(X * scale, Y * scale):
                        seen_any.add(ck)
                        if prim and role == want:
                            seen_role.add(ck)
                    for ck in seen_any:
                        any_h[ck] = any_h.get(ck, 0) + 1
                    for ck in seen_role:
                        role_h[ck] = role_h.get(ck, 0) + 1
        if n == 0:
            continue
        for ck, h in any_h.items():
            fa = float(h) / n
            if fa >= 0.9:
                cands.append((round(fa, 3), round(float(role_h.get(ck, 0)) / n, 3), -size[ck], gk, ck))
    cands.sort(reverse=True)
    info["candidates"] = len(cands)
    assign, used = {}, set()
    for fa, fr, ns, gk, ck in cands:
        if ck in assign or gk in used:
            continue
        assign[ck] = gk
        used.add(gk)
    info["unmatched"] = [gk for gk in groups if gk not in used]
    return assign, info
# ---------------------------------------------------------------------------------------------------------------------
#  Flip commands by POSITION (coordinate windows) + a record in the command that tells whether the mesh changed
# ---------------------------------------------------------------------------------------------------------------------
def f10(x):
    return "%.10g" % x


def item_mid(it):
    return ((it["p1"][0] + it["p2"][0]) / 2.0, (it["p1"][1] + it["p2"][1]) / 2.0)


def item_len(it):
    return math.hypot(it["p2"][0] - it["p1"][0], it["p2"][1] - it["p1"][1])


def split_components(items):
    """Chains of elements that share corner nodes (solver node numbers)."""
    parent = {}

    def find(a):
        while parent[a] != a:
            parent[a] = parent[parent[a]]
            a = parent[a]
        return a
    for it in items:
        a, b = it["nodes"][0], it["nodes"][1]
        parent.setdefault(a, a)
        parent.setdefault(b, b)
        ra, rb = find(a), find(b)
        if ra != rb:
            parent[ra] = rb
    groups = {}
    for it in items:
        groups.setdefault(find(it["nodes"][0]), []).append(it)
    out = [sorted(g, key=lambda x: x["e"]) for g in groups.values()]
    out.sort(key=lambda g: g[0]["e"])
    return out


def run_box(items):
    """Coordinate window around a chain of elements: its nodes plus 10 % of the mean element length."""
    xs = [p[0] for it in items for p in (it["p1"], it["p2"])]
    ys = [p[1] for it in items for p in (it["p1"], it["p2"])]
    m = 0.10 * sum(item_len(it) for it in items) / float(len(items))
    return (min(xs) - m, max(xs) + m, min(ys) - m, max(ys) + m)


def in_box(it, box):
    mx, my = item_mid(it)
    return box[0] <= mx <= box[1] and box[2] <= my <= box[3]


def select_box(pool, box):
    return [it for it in pool if in_box(it, box)]


def pool_items(d, side, real):
    """The elements ESEL type + real would select on this side (real None = every real constant)."""
    return [it for it in d.get(side, []) if real is None or it["real"] == real]


def real_token(r, hdr):
    if r is None:
        return "-"
    if hdr and r == hdr[0]:
        return "cid"
    if hdr and r == hdr[1]:
        return "tid"
    return str(r)


def resolve_real(tok, hdr):
    if tok in (None, "-", ""):
        return None
    if tok == "cid":
        return hdr[0]
    if tok == "tid":
        return hdr[1]
    return int(tok)


def mesh_signature(ids):
    """Fingerprint of the MESH nodes of a scoped edge (taken from the Mechanical mesh, no solver needed)."""
    ids = list(ids)
    if not ids:
        return {"nodes": 0, "idmin": 0, "idmax": 0, "idsum": 0}
    return {"nodes": len(ids), "idmin": min(ids), "idmax": max(ids), "idsum": sum(ids)}


def make_runs(chosen, pool):
    """One run per chain of chosen elements. A run is a coordinate window when that window selects exactly these
    elements among `pool` (what the command would select), otherwise it falls back to element numbers."""
    runs = []
    for comp in split_components(chosen):
        box = run_box(comp)
        same = set(it["e"] for it in select_box(pool, box)) == set(it["e"] for it in comp)
        runs.append({"items": comp, "box": box if same else None})
    return runs


def _id_ranges(nums):
    return ";".join(("%d-%d" % ab) if ab[0] != ab[1] else str(ab[0]) for ab in ranges(nums))


def record_lines(contact_id, created, revised, unit, sigs, runs):
    """The RECORD of a flip command ('! CNC-...' comment lines). unit: mesh coordinate = solver coordinate * unit.
    runs: [{'side', 'real' (token), 'n', 'box' (solver units) or None, 'ids' (element numbers) or None,
            'nodes': [(mesh node id, x, y), ...]}]"""
    L = ["! CNC-REC 2 id=%s created=%s revised=%s unit=%s" % (contact_id, created, revised or "-",
                                                             f10(unit) if unit else "-")]
    for sd in ("contact", "target"):
        if sd in sigs:
            g = sigs[sd]
            L.append("! CNC-SIG side=%s nodes=%d idmin=%d idmax=%d idsum=%d" % (sd, g["nodes"], g["idmin"], g["idmax"], g["idsum"]))
    for k, run in enumerate(runs):
        if run["box"] is not None:
            L.append("! CNC-RUN k=%d side=%s real=%s n=%d box=%s" % (k, run["side"], run["real"], run["n"], ",".join(f10(v) for v in run["box"])))
        else:
            L.append("! CNC-RUN k=%d side=%s real=%s n=%d ids=%s" % (k, run["side"], run["real"], run["n"], _id_ranges(run["ids"])))
        for (nid, x, y) in run["nodes"]:
            L.append("! CNC-FN k=%d id=%d x=%s y=%s" % (k, nid, f10(x), f10(y)))
    return L


def rewrite_record(text, rec_lines):
    """Replaces the record lines of a command, everything else (the commands) stays."""
    out, done = [], False
    for l in text.splitlines():
        if l.strip().startswith("! CNC-"):
            if not done:
                out.extend(rec_lines)
                done = True
            continue
        out.append(l)
    if not done:
        out = list(rec_lines) + out
    return "\n".join(out) + "\n"


def build_flip_positional(name, contact_id, plans, hdr, sigs, created, revised, unit):
    """The flip command by coordinate windows (runs that a window cannot describe exactly use element numbers)."""
    rruns = []
    for pl in plans:
        rt = real_token(pl["real"], hdr)
        for run in pl["runs"]:
            rruns.append({"side": pl["side"], "real": rt, "n": len(run["items"]), "box": run["box"],
                          "ids": [it["e"] for it in run["items"]] if run["box"] is None else None,
                          "nodes": run.get("nodes", [])})
    L = ["! ContactNormalCheck - normal flip by POSITION (ESURF,,REVERSE): a coordinate window keeps working "
         "after a mesh change",
         "! Contact: %s" % name]
    L += record_lines(contact_id, created, revised, unit, sigs, rruns)
    L.append("allsel,all")
    for pl in plans:
        typ = "cid" if pl["side"] == "contact" else "tid"
        rt = real_token(pl["real"], hdr)
        for run in pl["runs"]:
            if run["box"] is not None:
                b = run["box"]
                L.append("esel,s,type,,%s" % typ)
                if pl["real"] is not None:
                    L.append("esel,r,real,,%s" % rt)
                L.append("esel,r,cent,x,%s,%s" % (f10(b[0]), f10(b[1])))
                L.append("esel,r,cent,y,%s,%s" % (f10(b[2]), f10(b[3])))
            else:
                L.append("esel,none")
                for lo, hi in ranges([it["e"] for it in run["items"]]):
                    L.append("esel,a,elem,,%d,%d,1" % (lo, hi))
                L.append("esel,r,type,,%s" % typ)
                if pl["real"] is not None:
                    L.append("esel,r,real,,%s" % rt)
            L.append("esurf,,reverse")
    L.append("allsel,all")
    return "\n".join(L) + "\n"


def parse_cnc_record(text):
    """Reads the record back. None when the command has no (version 2) record."""
    rec = None
    for l in text.splitlines():
        l = l.strip()
        if not l.startswith("! CNC-"):
            continue
        parts = l[2:].split()
        kind = parts[0]
        kv = dict(p.split("=", 1) for p in parts[1:] if "=" in p)
        try:
            if kind == "CNC-REC":
                if len(parts) < 2 or parts[1] != "2":
                    return None
                u = kv.get("unit", "-")
                rec = {"id": kv.get("id"), "created": kv.get("created"), "unit": float(u) if u != "-" else None,
                       "sigs": {}, "runs": []}
            elif rec is None:
                continue
            elif kind == "CNC-SIG":
                rec["sigs"][kv["side"]] = {"nodes": int(kv["nodes"]), "idmin": int(kv["idmin"]),
                                           "idmax": int(kv["idmax"]), "idsum": int(kv["idsum"])}
            elif kind == "CNC-RUN":
                box = tuple(float(v) for v in kv["box"].split(",")) if "box" in kv else None
                rec["runs"].append({"side": kv["side"], "real": kv.get("real", "-"), "n": int(kv["n"]),
                                    "box": box, "fn": []})
            elif kind == "CNC-FN":
                rec["runs"][int(kv["k"])]["fn"].append((int(kv["id"]), float(kv["x"]), float(kv["y"])))
        except (KeyError, ValueError, IndexError):
            continue
    return rec


def check_mesh_record(rec, cur_ids, node_xy):
    """Did the mesh change since the command was written? No solver needed.
    cur_ids: {side: ids of the nodes on the scoped edges now}; node_xy(side, id) -> (x, y) or None.
    -> list of problems (empty = still the same mesh). Checks node count and numbers of each edge and every flipped node."""
    problems = []
    for sd, sg in sorted(rec["sigs"].items()):
        cur = mesh_signature(cur_ids.get(sd, []))
        if cur["nodes"] != sg["nodes"]:
            problems.append(u"%s edge: %d nodes, was %d" % (sd.capitalize(), cur["nodes"], sg["nodes"]))
        elif (cur["idmin"], cur["idmax"], cur["idsum"]) != (sg["idmin"], sg["idmax"], sg["idsum"]):
            problems.append(u"%s edge: node numbers changed" % sd.capitalize())
    missing = moved = total = 0
    for run in rec["runs"]:
        for (nid, x, y) in run["fn"]:
            total += 1
            p = node_xy(run["side"], nid)
            if p is None:
                missing += 1
            elif math.hypot(p[0] - x, p[1] - y) > 1e-6 * max(1.0, abs(x), abs(y)):
                moved += 1
    if missing or moved:
        problems.append(u"Flipped nodes: %d of %d missing, %d moved" % (missing, total, moved))
    return problems


def remap_mesh_record(rec, cur_xy):
    """New record data from the NEW mesh. The coordinate windows of the command stay (they are positions); the nodes
    inside each window are listed again and the edge fingerprints are renewed.
    cur_xy: {side: {node id: (x, y)}} of the scoped edges now. -> (runs, sigs, report) or (None, None, reason)."""
    unit = rec.get("unit") or 1.0
    runs, old_n, new_n = [], 0, 0
    for run in rec["runs"]:
        sd = run["side"]
        if run["box"] is None:
            return None, None, u"A part was written by element numbers (no window) - it cannot be moved to a new mesh"
        b = [v * unit for v in run["box"]]
        nodes = [(nid, p[0], p[1]) for nid, p in sorted(cur_xy.get(sd, {}).items())
                 if b[0] <= p[0] <= b[1] and b[2] <= p[1] <= b[3]]
        if not nodes:
            return None, None, u"The window of the %s side contains no node of the edge any more" % sd
        old_n += len(run["fn"])
        new_n += len(nodes)
        runs.append({"side": sd, "real": run["real"], "n": max(len(nodes) - 1, 1), "box": run["box"], "ids": None,
                     "nodes": nodes})
    sides = set(rec["sigs"].keys()) | set(r["side"] for r in runs)
    sigs = dict((sd, mesh_signature(list(cur_xy.get(sd, {}).keys()))) for sd in sides)
    return runs, sigs, u"%d -> %d nodes in the window(s)" % (old_n, new_n)
# ==== CORE END ====

import os
from System.Windows.Forms import (Form, Button, Label, MessageBox, MessageBoxButtons,
                                  MessageBoxIcon, FormStartPosition, FormBorderStyle,
                                  DialogResult, GroupBox, FlatStyle, TextBox, TreeView,
                                  TreeNode, Application, ComboBox, ComboBoxStyle,
                                  ListView, ListViewItem, View, ScrollBars, Panel,
                                  Cursor, Cursors, Control, Keys, MouseButtons, Clipboard,
                                  Timer, IWin32Window, BorderStyle, AnchorStyles, Screen, ToolTip,
                                  SplitContainer, Orientation, DockStyle, FixedPanel, OpenFileDialog)
from System.Drawing import Size, Point, Font, FontStyle, Color, PointF, Pen, SolidBrush, Bitmap, Graphics
from System.Drawing.Drawing2D import Matrix, SmoothingMode, DashStyle

import Microsoft.Win32 as Win32

REG_KEY = r"Software\ContactNormalCheck"


SNIPPET_NAME = u"ContactNormalFlip"
PALETTE = [(150, 165, 182), (155, 180, 148), (232, 220, 200), (190, 195, 225),
           (232, 190, 190), (180, 225, 225), (220, 205, 160), (200, 180, 215),
           (170, 200, 170), (215, 175, 150), (160, 190, 215), (205, 205, 205)]


# ---------------------------------------------------------------
# Contact helpers
# ---------------------------------------------------------------
def obj_id(c):
    try:
        return int(c.ObjectId)
    except Exception:
        return None


def label_of(c):
    oid = obj_id(c)
    return c.Name + (u"  [ID %d]" % oid if oid is not None else u"")


def ukey_of(c, idx):
    oid = obj_id(c)
    return oid if oid is not None else 1000000 + idx


def build_contact_tree(connections):
    tree = []
    for conn_group in connections.Children:
        group_contacts = []
        for contact in conn_group.Children:
            if "ContactRegion" in contact.GetType().ToString():
                group_contacts.append(contact)
        if group_contacts:
            tree.append((conn_group.Name, group_contacts))
    return tree


def find_flip_snippets(contact):
    """ALL flip commands this tool wrote under a contact ("ContactNormalFlip", "ContactNormalFlip 2", ...), in order."""
    out = []
    for child in contact.Children:
        try:
            if "CommandSnippet" in child.GetType().ToString() and child.Name.startswith(SNIPPET_NAME):
                out.append(child)
        except Exception:
            pass
    return out


def find_flip_snippet(contact):
    for child in contact.Children:
        try:
            if "CommandSnippet" in child.GetType().ToString() and child.Name == SNIPPET_NAME:
                return child
        except Exception:
            pass
    return None


# ---------------------------------------------------------------
# Model scene (from mesh data) - every mesh element is drawn
# ---------------------------------------------------------------
def tree_body_labels(log):
    """GeoBody id -> (part name, body name) exactly as the Mechanical tree (Geometry) shows them."""
    out = {}
    try:
        for part in ExtAPI.DataModel.Project.Model.Geometry.Children:
            try:
                kids = list(part.Children)
            except Exception:
                continue
            for b in kids:
                gid = None
                try:
                    gid = b.GetGeoBody().Id
                except Exception:
                    try:
                        gid = b.GeoBody.Id
                    except Exception:
                        gid = None
                if gid is not None:
                    out[int(gid)] = (part.Name, b.Name)
    except Exception as ex:
        log(u"Body names from the Mechanical tree are not available (%s) - geometry data names are used" % str(ex))
    return out


def load_scene(log):
    mesh = ExtAPI.DataModel.MeshDataByName("Global")
    owner = {}
    bodies = []
    body_part = []
    body_lab = []
    tl = tree_body_labels(log)
    try:
        pidx = -1
        for asm in ExtAPI.DataModel.GeoData.Assemblies:
            for part in asm.Parts:
                pidx += 1
                for body in part.Bodies:
                    name = u"%s\\%s [%d]" % (part.Name, body.Name, body.Id)
                    bodies.append(name)
                    body_part.append(pidx)
                    body_lab.append(tl.get(int(body.Id), (part.Name, body.Name)))
                    for face in body.Faces:
                        for eid in list(mesh.MeshRegionById(face.Id).ElementIds):
                            owner[eid] = len(bodies) - 1
    except Exception as ex:
        log(u"Body/face lookup failed (%s)" % str(ex))
    names = list(bodies) + [u"Other elements"]
    body_part.append(max(body_part) + 1 if body_part else 0)
    body_lab.append((u"", u"Other elements"))
    groups = dict((i, []) for i in range(len(names)))
    for el in mesh.Elements:
        groups[owner.get(el.Id, len(names) - 1)].append(el.Id)
    xy = {}
    scene = {"bodies": [], "xy": xy, "mesh": mesh, "eid2body": {}}
    drawn = skipped = 0
    for bi, name in enumerate(names):
        elems = []
        for eid in groups[bi]:
            el = mesh.ElementById(eid)
            alln = list(el.NodeIds)
            try:
                api_c = list(el.CornerNodeIds)
            except Exception:
                api_c = None
            corner = corner_of(alln, api_c)
            for nid in alln:
                if nid not in xy:
                    nd = mesh.NodeById(nid)
                    xy[nid] = (nd.X, nd.Y)
            pts = [(xy[n][1], -xy[n][0]) for n in corner if n in xy]
            if len(set(pts)) < 3:
                skipped += 1
                continue
            xs = [p[0] for p in pts]
            ys = [p[1] for p in pts]
            poly = System.Array[PointF]([PointF(p[0], p[1]) for p in pts])
            elems.append((eid, corner, alln, poly, pts, (min(xs), max(xs), min(ys), max(ys))))
            drawn += 1
        if not elems:
            continue
        body = {"name": name, "elems": elems, "visible": True,
                "part": body_lab[bi][0], "label": body_lab[bi][1],
                "color": Color.FromArgb(*PALETTE[body_part[bi] % len(PALETTE)])}
        scene["bodies"].append(body)
        for e in elems:
            scene["eid2body"][e[0]] = body
    xs = [p[0] for p in xy.values()]
    ys = [p[1] for p in xy.values()]
    scene["box"] = (min(ys), max(ys), -max(xs), -min(xs))      # draw coords (X=y, Y=-x)
    log(u"Model view: %d bodies, %d area elements drawn, %d line/degenerate elements skipped; "
        u"%d body names read from the Mechanical tree" % (len(scene["bodies"]), drawn, skipped, len(tl)))
    return scene


def contact_node_sets(scene, c, log):
    """Mesh nodes of the contact edges and of the target edges of a contact."""
    mesh = scene["mesh"]
    out = {}
    for side, loc in (("contact", c.SourceLocation), ("target", c.TargetLocation)):
        nodes = set()
        for gid in list(loc.Ids):
            try:
                nodes.update(list(mesh.MeshRegionById(gid).NodeIds))
            except Exception as ex:
                log(u"%s: edge %s mesh lookup failed (%s)" % (label_of(c), gid, str(ex)))
        out[side] = nodes
    return out


# ---------------------------------------------------------------
# Solver input (ds.dat): real contact/target element orientation
# ---------------------------------------------------------------


DUMP_NAME = u"CNC_Dump"
DUMPALL_NAME = u"CNC_DumpAll"
EXIT_NAME = u"CNC_Exit"
try:
    from Ansys.Mechanical.DataModel.Enums import SequenceSelectionType
except Exception:
    SequenceSelectionType = None


def find_analysis(model):
    for a in model.Analyses:
        if "static" in str(a.AnalysisType).lower():
            return a
    return model.Analyses[0]


def find_snippet(contact, name):
    for child in contact.Children:
        try:
            if "CommandSnippet" in child.GetType().ToString() and child.Name == name:
                return child
        except Exception:
            pass
    return None


def read_text_lines(path):
    """Read a file that the solver may still be writing."""
    fs = System.IO.FileStream(path, System.IO.FileMode.Open, System.IO.FileAccess.Read,
                              System.IO.FileShare.ReadWrite)
    try:
        sr = System.IO.StreamReader(fs)
        return sr.ReadToEnd().splitlines()
    finally:
        fs.Close()


def solver_report(wd, n=16):
    """What the solver left in its working directory: tail of solve.out and the newest files."""
    out = []
    try:
        for nm in (u"solve.out", u"file.out", u"ds.out"):
            p = System.IO.Path.Combine(wd, nm)
            if not System.IO.File.Exists(p):
                continue
            fs = System.IO.FileStream(p, System.IO.FileMode.Open, System.IO.FileAccess.Read,
                                      System.IO.FileShare.ReadWrite)
            try:
                size = fs.Length
                fs.Seek(max(0, size - 20000), System.IO.SeekOrigin.Begin)
                text = System.IO.StreamReader(fs).ReadToEnd()
            finally:
                fs.Close()
            ls = [l for l in text.splitlines() if l.strip()][-n:]
            out.append(u"%s (%d KB, modified %s) - last lines:" % (
                nm, size // 1024, System.IO.File.GetLastWriteTime(p).ToString("HH:mm:ss")))
            out.extend(u"   " + l.strip()[:100] for l in ls)
            break
        for nm in (u"file0.err", u"file1.err", u"file.err"):
            p = System.IO.Path.Combine(wd, nm)
            if not System.IO.File.Exists(p):
                continue
            fs = System.IO.FileStream(p, System.IO.FileMode.Open, System.IO.FileAccess.Read,
                                      System.IO.FileShare.ReadWrite)
            try:
                size = fs.Length
                fs.Seek(max(0, size - 6000), System.IO.SeekOrigin.Begin)
                text = System.IO.StreamReader(fs).ReadToEnd()
            finally:
                fs.Close()
            ls = [l for l in text.splitlines() if l.strip()][-12:]
            out.append(u"%s (%d KB, modified %s) - last lines:" % (
                nm, size // 1024, System.IO.File.GetLastWriteTime(p).ToString("HH:mm:ss")))
            out.extend(u"   " + l.strip()[:110] for l in ls)
            break
        files = sorted(System.IO.Directory.GetFiles(wd), key=lambda f: System.IO.File.GetLastWriteTime(f),
                       reverse=True)[:6]
        out.append(u"newest files: " + u", ".join(
            u"%s (%s)" % (System.IO.Path.GetFileName(f), System.IO.File.GetLastWriteTime(f).ToString("HH:mm:ss"))
            for f in files))
    except Exception as ex:
        out.append(u"(could not read the solver directory: %s)" % str(ex))
    return u"\n".join(out) if out else u"(no solver output found)"


class WaitForm(Form):
    def __init__(self):
        self.Text = TITLE
        self.ClientSize = Size(520, 90)
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.ControlBox = False
        self.lbl = Label()
        self.lbl.Location = Point(16, 16)
        self.lbl.Size = Size(490, 60)
        self.Controls.Add(self.lbl)

    def set_text(self, t):
        self.lbl.Text = t
        Application.DoEvents()


# ---------------------------------------------------------------
# Interactive view
# ---------------------------------------------------------------
def _dist_pt_seg(px, py, ax, ay, bx, by):
    dx, dy = bx - ax, by - ay
    d2 = dx * dx + dy * dy
    t = 0.0 if d2 == 0 else max(0.0, min(1.0, ((px - ax) * dx + (py - ay) * dy) / d2))
    return math.hypot(px - (ax + t * dx), py - (ay + t * dy))


class SceneView(Panel):
    def __init__(self, owner):
        Panel.__init__(self)
        self.owner = owner
        self.DoubleBuffered = True
        self.BackColor = Color.White
        self.scene = None
        self.segs = []
        self.sel = set()
        self.note = u""
        self.mesh_on = True
        self.hide_mode = False
        self.hidden = []
        self.s = 1.0
        self.px = 0.0
        self.py = 0.0
        self.cx = 0.0
        self.cy = 0.0
        self._down = None
        self._pan = None
        self._drag = False
        self._rb = None
        self._err = u""
        self.warn = u""
        self.show_cp = False
        self.local_on = True
        self.local_ids = None
        self.local_box = None
        self._fit_pending = False
        self._last = None
        self.Paint += self._paint
        self.MouseDown += self._mdown
        self.MouseMove += self._mmove
        self.MouseUp += self._mup
        self.MouseWheel += self._wheel
        self.MouseEnter += lambda s, e: self.Focus()
        self.Resize += self._on_resize

    def _on_resize(self, sender, e):
        """Panel resized (splitter / window): scale the drawing so the same model portion stays visible."""
        W, H = self.Width, self.Height
        if (self._last is not None and not self._fit_pending and W > 50 and H > 50
                and self._last[0] > 50 and self._last[1] > 50):
            r = min(float(W) / self._last[0], float(H) / self._last[1])
            self.s *= r
            self.px *= r
            self.py *= r
        if W > 0 and H > 0:
            self._last = (W, H)
        self.Invalidate()

    # draw coords: X = model y, Y = -model x   (view from -Z, X up, Y right).
    # Everything is converted to device pixels in Python (float64) - no GDI+ world transform.
    def _ofs(self):
        return (self.Width / 2.0 + self.px - self.cx * self.s,
                self.Height / 2.0 + self.py - self.cy * self.s)

    def _to_screen(self, pts):
        ox, oy = self._ofs()
        s = self.s
        return [(x * s + ox, y * s + oy) for (x, y) in pts]

    def _to_draw(self, mx, my):
        return (self.cx + (mx - self.Width / 2.0 - self.px) / self.s,
                self.cy + (my - self.Height / 2.0 - self.py) / self.s)

    def _visible(self):
        return [(i, g) for i, g in enumerate(self.segs)
                if seg_visible(g) and (self.show_cp or not g.get("cp"))]

    def _arrow_len(self):
        """Arrow length in screen pixels (relative to the element size, clamped)."""
        vis = self._visible()
        if not vis:
            return 30.0
        lens = sorted(g["len"] * self.s for i, g in vis)
        return arrow_px(lens[len(lens) // 2])

    # ----- fit / zoom -----
    def _fit_box(self, box):
        x0, x1, y0, y1 = box
        bw, bh = x1 - x0, y1 - y0
        big = max(bw, bh, 1e-9)
        bw, bh = max(bw, 0.25 * big), max(bh, 0.25 * big)
        self.cx, self.cy = (x0 + x1) / 2.0, (y0 + y1) / 2.0
        self.s = min((self.Width - 40.0) / (bw * 1.5), (self.Height - 40.0) / (bh * 1.5))
        self.px = self.py = 0.0

    def fit(self):
        vis = self._visible()
        if vis:
            xs, ys = [], []
            for i, g in vis:
                for (x, y) in ((g["ax"], g["ay"]), (g["bx"], g["by"])):
                    xs.append(y)
                    ys.append(-x)
            self._fit_box((min(xs), max(xs), min(ys), max(ys)))
            self.Invalidate()
        else:
            self.fit_all()

    def fit_all(self):
        if self.scene is not None:
            self._fit_box(self.local_box if (self.local_on and self.local_box) else self.scene["box"])
        self.Invalidate()

    def _drawn(self, b):
        """Body is visible and (in local mode) touches the current contact's edges."""
        if not b["visible"]:
            return False
        return not (self.local_on and self.local_ids is not None and id(b) not in self.local_ids)

    def zoom(self, f, mx=None, my=None):
        if mx is None:
            mx, my = self.Width / 2.0, self.Height / 2.0
        self.px, self.py = zoom_about(self.px, self.py, mx, my, self.Width, self.Height, f)
        self.s = max(1e-9, self.s * f)
        self.Invalidate()

    # ----- painting -----
    def center_on(self, X, Y):
        """Center the view on a point given in draw coordinates (used by the overview)."""
        self.cx, self.cy = X, Y
        self.px = self.py = 0.0
        self.Invalidate()

    def request_fit(self):
        """Fit on the next paint, when the control already has its final size."""
        self._fit_pending = True
        self.Invalidate()

    def _paint(self, sender, e):
        g = e.Graphics
        if self._fit_pending and self.Width > 60 and self.Height > 60:
            self._fit_pending = False
            self.fit()
        g.Clear(Color.White)
        g.SmoothingMode = SmoothingMode.AntiAlias
        self._err = u""
        if self.scene is not None:
            try:
                self._draw_geometry(g)
                self._draw_segments(g)
            except Exception as ex:
                self._err = u"Draw error: %s" % str(ex)
        self._draw_overlay(g)
        ov = getattr(self.owner, "overview", None)
        if ov is not None:
            ov.Invalidate()

    def _draw_geometry(self, g):
        s = self.s
        ox, oy = self._ofs()
        W, H = float(self.Width), float(self.Height)
        BIG = 1.0e6
        for b in self.scene["bodies"]:
            if not self._drawn(b):
                continue
            br = SolidBrush(b["color"])
            pen = Pen(Color.FromArgb(90, 90, 90) if self.mesh_on else b["color"], 1.0)
            for el in b["elems"]:
                x0, x1, y0, y1 = el[5]
                if x1 * s + ox < -2 or x0 * s + ox > W + 2 or y1 * s + oy < -2 or y0 * s + oy > H + 2:
                    continue
                pts = System.Array[PointF]([PointF(max(-BIG, min(BIG, p[0] * s + ox)),
                                                   max(-BIG, min(BIG, p[1] * s + oy))) for p in el[4]])
                g.FillPolygon(br, pts)
                if self.mesh_on and (x1 - x0) * s > 3.0:
                    g.DrawPolygon(pen, pts)

    def _draw_segments(self, g):
        base = {"contact": Color.FromArgb(220, 20, 40), "target": Color.FromArgb(30, 70, 220)}
        orange = Color.FromArgb(255, 120, 0)
        pale = {"contact": Color.FromArgb(255, 165, 175), "target": Color.FromArgb(150, 170, 255)}
        glow1 = Pen(Color.FromArgb(0, 210, 120), 12.0)
        glow2 = Pen(Color.FromArgb(0, 210, 120), 9.0)
        pens, shafts, fills = {}, {}, {}
        for k in base:
            for rv in (False, True):
                c = orange if rv else base[k]
                pens[(k, rv)] = Pen(c, 4.0)
                shafts[(k, rv)] = Pen(c, 2.4)
                fills[(k, rv)] = SolidBrush(c)
            pens[(k, "cp")] = Pen(pale[k], 3.0)
            shafts[(k, "cp")] = Pen(pale[k], 1.6)
            fills[(k, "cp")] = SolidBrush(pale[k])
        L = self._arrow_len()
        off = 3.0
        hl = min(L * 0.45, 13.0)
        r = 3.0
        s = self.s
        ox, oy = self._ofs()
        for i, sg in self._visible():
            key = (sg["side"], "cp") if sg.get("cp") else (sg["side"], sg["rev"])
            dx, dy = sg["ny"], -sg["nx"]                       # unit normal in draw/screen coords
            A = PointF(sg["ay"] * s + ox + dx * off, -sg["ax"] * s + oy + dy * off)
            B = PointF(sg["by"] * s + ox + dx * off, -sg["bx"] * s + oy + dy * off)
            tl = math.hypot(B.X - A.X, B.Y - A.Y)
            lat = (5.0 if sg["side"] == "contact" else -5.0)
            tx_, ty_ = ((B.X - A.X) / tl, (B.Y - A.Y) / tl) if tl > 1e-9 else (0.0, 0.0)
            M = PointF((A.X + B.X) / 2.0 + tx_ * lat, (A.Y + B.Y) / 2.0 + ty_ * lat)
            T = PointF(M.X + dx * L, M.Y + dy * L)
            if i in self.sel:
                g.DrawLine(glow1, A, B)
                g.DrawLine(glow2, M, T)
            g.DrawLine(pens[key], A, B)
            g.DrawLine(shafts[key], M, T)
            qx, qy = -dy, dx
            head = System.Array[PointF]([T,
                                         PointF(T.X - dx * hl + qx * hl * 0.5, T.Y - dy * hl + qy * hl * 0.5),
                                         PointF(T.X - dx * hl - qx * hl * 0.5, T.Y - dy * hl - qy * hl * 0.5)])
            g.FillPolygon(fills[key], head)
            g.FillEllipse(fills[key], M.X - r, M.Y - r, 2 * r, 2 * r)

    def _draw_overlay(self, g):
        f = Font("Segoe UI", 8.5)
        dim = SolidBrush(Color.DimGray)
        vis = self._visible()
        nc = sum(1 for i, x in vis if x["side"] == "contact")
        nr = sum(1 for i, x in vis if x["rev"])
        g.DrawString(u"Contact %d / Target %d Segments | To Flip %d | Selected %d    %s"
                     % (nc, len(vis) - nc, nr, len(self.sel), self.note), f, dim, 6.0, 4.0)
        y = self.Height - 20.0
        x = 8.0
        for col, txt, w in ((Color.FromArgb(220, 20, 40), u"Contact Normal", 100.0),
                            (Color.FromArgb(30, 70, 220), u"Target Normal", 98.0),
                            (Color.FromArgb(255, 120, 0), u"Should Be Flipped", 120.0),
                            (Color.FromArgb(0, 210, 120), u"Selected", 62.0)):
            g.FillRectangle(SolidBrush(col), x, y + 3, 14.0, 8.0)
            g.DrawString(txt, f, dim, x + 18.0, y)
            x += w + 22.0
        if self.warn:
            g.DrawString(self.warn, f, SolidBrush(Color.Firebrick), 6.0, 22.0)
        if getattr(self, "_err", u""):
            g.DrawString(self._err, f, SolidBrush(Color.Firebrick), 6.0, 22.0)
        if self.hide_mode:
            g.DrawString(u"HIDE BODY: Click A Body", f, SolidBrush(Color.DarkOrange), self.Width - 170.0, 4.0)
        if self._rb is not None:
            x0, y0, x1, y1 = self._rb
            pen = Pen(Color.DarkOrange, 1.0)
            pen.DashStyle = DashStyle.Dash
            g.DrawRectangle(pen, min(x0, x1), min(y0, y1), abs(x1 - x0), abs(y1 - y0))

    # ----- interaction -----
    def _wheel(self, sender, e):
        self.zoom(1.2 if e.Delta > 0 else 1.0 / 1.2, e.X, e.Y)

    def _mdown(self, sender, e):
        if e.Button == MouseButtons.Left:
            self._down = (e.X, e.Y, self.px, self.py)
            self._drag = False
        elif e.Button == MouseButtons.Middle:
            self._pan = (e.X, e.Y, self.px, self.py)
            self.Cursor = Cursors.SizeAll

    def _mmove(self, sender, e):
        if self._pan is not None:
            x0, y0, p0x, p0y = self._pan
            self.px = p0x + (e.X - x0)
            self.py = p0y + (e.Y - y0)
            self.Invalidate()
            return
        if self._down is None:
            return
        x0, y0, p0x, p0y = self._down
        if not self._drag and math.hypot(e.X - x0, e.Y - y0) > 4:
            self._drag = True
        if not self._drag:
            return
        if (Control.ModifierKeys & Keys.Shift) == Keys.Shift and not self.hide_mode:
            self._rb = (x0, y0, e.X, e.Y)
        else:
            self.px = p0x + (e.X - x0)
            self.py = p0y + (e.Y - y0)
        self.Invalidate()

    def _mup(self, sender, e):
        if e.Button == MouseButtons.Middle:
            self._pan = None
            self.Cursor = Cursors.Cross if self.hide_mode else Cursors.Default
            return
        if self._down is None:
            return
        x0, y0 = self._down[0], self._down[1]
        if self.hide_mode:
            if not self._drag:
                self._hide_at(x0, y0)
        elif self._drag and self._rb is not None:
            xa, xb = sorted((x0, e.X))
            ya, yb = sorted((y0, e.Y))
            vis = self._visible()
            pts = self._to_screen([(g["my"], -g["mx"]) for i, g in vis])
            for (i, g), (sx, sy) in zip(vis, pts):
                if xa <= sx <= xb and ya <= sy <= yb:
                    self.sel.add(i)
            self.owner.selection_changed()
        elif not self._drag:
            self._pick(x0, y0)
        self._down = None
        self._drag = False
        self._rb = None
        self.Invalidate()

    def _pick(self, mx, my):
        best, bi = 14.0, None
        L = self._arrow_len() / self.s
        for i, g in self._visible():
            dx, dy = g["ny"], -g["nx"]
            pts = self._to_screen([(g["ay"], -g["ax"]), (g["by"], -g["bx"]), (g["my"], -g["mx"]),
                                   (g["my"] + dx * L, -g["mx"] + dy * L)])
            d = min(_dist_pt_seg(mx, my, pts[0][0], pts[0][1], pts[1][0], pts[1][1]),
                    _dist_pt_seg(mx, my, pts[2][0], pts[2][1], pts[3][0], pts[3][1]))
            if d < best:
                best, bi = d, i
        if bi is not None:
            if bi in self.sel:
                self.sel.discard(bi)
            else:
                self.sel.add(bi)
            self.owner.selection_changed()

    def select_reversed(self):
        for i, g in self._visible():
            if g["rev"]:
                self.sel.add(i)
        self.owner.selection_changed()
        self.Invalidate()

    # ----- body hide / undo (hidden body's contact/target segments are hidden too) -----
    def _hide_at(self, mx, my):
        if self.scene is None:
            return
        X, Y = self._to_draw(mx, my)
        for b in self.scene["bodies"]:
            if not self._drawn(b):
                continue
            for el in b["elems"]:
                x0, x1, y0, y1 = el[5]
                if x0 <= X <= x1 and y0 <= Y <= y1 and point_in_poly(X, Y, el[4]):
                    self.hide_body(b)
                    return

    def hide_body(self, b):
        if b["visible"]:
            b["visible"] = False
            self.hidden.append(b)
            self.owner.note(u"Hidden: %s" % b["name"])
        self.owner.sync_tree()
        self.Invalidate()

    def _drop_from_stack(self, b):
        """Remove a body from the undo stack (entries are bodies or groups of bodies)."""
        new = []
        for x in self.hidden:
            if isinstance(x, list):
                y = [z for z in x if z is not b]
                if y:
                    new.append(y)
            elif x is not b:
                new.append(x)
        self.hidden = new

    def show_body(self, b):
        b["visible"] = True
        self._drop_from_stack(b)
        self.owner.note(u"Shown: %s" % b["name"])
        self.owner.sync_tree()
        self.Invalidate()

    def isolate(self, keep):
        """Only the bodies in `keep` stay visible; the rest is hidden as ONE undo step."""
        group = []
        for b in self.scene["bodies"]:
            if any(b is k for k in keep):
                if not b["visible"]:
                    b["visible"] = True
                    self._drop_from_stack(b)
            elif b["visible"]:
                b["visible"] = False
                group.append(b)
        if group:
            self.hidden.append(group)
        self.owner.note(u"Isolated: kept %s; %d other bodies hidden (Undo / Ctrl+Z / Show All brings them back)"
                        % (u", ".join(k["name"] for k in keep), len(group)))
        self.owner.sync_tree()
        self.Invalidate()

    def undo_hide(self):
        if self.hidden:
            x = self.hidden.pop()
            bs = x if isinstance(x, list) else [x]
            for b in bs:
                b["visible"] = True
            self.owner.note(u"Restored: %s" % (u"%d bodies" % len(bs) if isinstance(x, list) else x["name"]))
            self.owner.sync_tree()
            self.Invalidate()

    def show_all(self):
        while self.hidden:
            x = self.hidden.pop()
            for b in (x if isinstance(x, list) else [x]):
                b["visible"] = True
        self.owner.sync_tree()
        self.Invalidate()

    def clear_hidden(self):
        """All bodies visible again, undo stack empty (no messages)."""
        for b in self.scene["bodies"]:
            b["visible"] = True
        self.hidden = []

    def put_hidden(self, stack):
        """Shows the hidden state of a contact: `stack` = its saved undo stack (bodies / groups of bodies, None = nothing
        hidden). Every contact has its own hidden bodies; they come back as they were left."""
        self.clear_hidden()
        keep = []
        for x in (stack or []):
            for b in (x if isinstance(x, list) else [x]):
                b["visible"] = False
            keep.append(x)
        self.hidden = keep

    def set_hide_mode(self, on):
        self.hide_mode = on
        self.Cursor = Cursors.Cross if on else Cursors.Default
        self.Invalidate()


class OverviewView(Panel):
    """Whole model on the right; the current contact is marked and the area shown in the main view
    is outlined. Click / drag inside it to move the main view there."""
    def __init__(self, owner):
        Panel.__init__(self)
        self.owner = owner
        self.DoubleBuffered = True
        self.BackColor = Color.White
        self.Cursor = Cursors.Hand
        self.scene = None
        self.segs = []
        self.main = None
        self._cache = None
        self._key = None
        self.Paint += self._paint
        self.MouseDown += self._mouse
        self.MouseMove += self._mouse
        self.Resize += lambda s, e: self.Invalidate()

    def _geom(self):
        x0, x1, y0, y1 = self.scene["box"]
        bw, bh = max(x1 - x0, 1e-12), max(y1 - y0, 1e-12)
        W, H = float(self.Width), float(self.Height)
        sc = min((W - 20.0) / bw, (H - 20.0) / bh)
        return sc, W / 2.0 - sc * (x0 + x1) / 2.0, H / 2.0 - sc * (y0 + y1) / 2.0

    def _render(self, sc, ox, oy):
        bmp = Bitmap(max(1, self.Width), max(1, self.Height))
        g = Graphics.FromImage(bmp)
        try:
            g.Clear(Color.White)
            for b in self.scene["bodies"]:
                br = SolidBrush(b["color"])                  # hidden bodies stay visible here
                pen = Pen(b["color"], 1.0)
                for el in b["elems"]:
                    pts = System.Array[PointF]([PointF(p[0] * sc + ox, p[1] * sc + oy) for p in el[4]])
                    g.FillPolygon(br, pts)
                    g.DrawPolygon(pen, pts)
        finally:
            g.Dispose()
        return bmp

    def _paint(self, sender, e):
        g = e.Graphics
        g.Clear(Color.White)
        if self.scene is None or self.Width < 30 or self.Height < 30:
            return
        try:
            sc, ox, oy = self._geom()
            key = (self.Width, self.Height, id(self.scene))
            if key != self._key or self._cache is None:
                if self._cache is not None:
                    self._cache.Dispose()
                self._cache = self._render(sc, ox, oy)
                self._key = key
            g.DrawImage(self._cache, 0, 0)
            g.SmoothingMode = SmoothingMode.AntiAlias
            vis = [x for x in self.segs if not x.get("cp")]
            if vis:
                pens = {"contact": Pen(Color.FromArgb(220, 20, 40), 3.0),
                        "target": Pen(Color.FromArgb(30, 70, 220), 3.0)}
                xs, ys = [], []
                for sg in vis:
                    a = (sg["ay"] * sc + ox, -sg["ax"] * sc + oy)
                    b = (sg["by"] * sc + ox, -sg["bx"] * sc + oy)
                    g.DrawLine(pens[sg["side"]], PointF(a[0], a[1]), PointF(b[0], b[1]))
                    xs += [a[0], b[0]]
                    ys += [a[1], b[1]]
                pad = 10.0
                g.DrawRectangle(Pen(Color.FromArgb(255, 140, 0), 2.0), min(xs) - pad, min(ys) - pad,
                                max(xs) - min(xs) + 2 * pad, max(ys) - min(ys) + 2 * pad)
            m = self.main
            if m is not None and m.s > 0 and m.Width > 0:
                hw, hh = (m.Width / 2.0) / m.s, (m.Height / 2.0) / m.s
                Xc, Yc = m.cx - m.px / m.s, m.cy - m.py / m.s
                x0, x1 = (Xc - hw) * sc + ox, (Xc + hw) * sc + ox
                y0, y1 = (Yc - hh) * sc + oy, (Yc + hh) * sc + oy
                pen = Pen(Color.FromArgb(120, 0, 160), 1.5)
                pen.DashStyle = DashStyle.Dash
                g.DrawRectangle(pen, x0, y0, max(2.0, x1 - x0), max(2.0, y1 - y0))
        except Exception as ex:
            g.DrawString(u"Overview error: %s" % str(ex), Font("Segoe UI", 8.5),
                         SolidBrush(Color.Firebrick), 6.0, 6.0)

    def _mouse(self, sender, e):
        if e.Button != MouseButtons.Left or self.scene is None or self.main is None:
            return
        sc, ox, oy = self._geom()
        self.main.center_on((e.X - ox) / sc, (e.Y - oy) / sc)


def bring_to_front(form):
    """Raises the window above Mechanical once; TopMost is switched off again right away so
    other windows can still be put on top of it later."""
    try:
        form.TopMost = True
        form.Activate()
        form.BringToFront()
        form.TopMost = False
    except Exception:
        pass


def schedule_front(form, times=4, interval=300):
    """Mechanical may grab the focus again right after our code returns; raise the window a few
    more times shortly afterwards (it never stays topmost)."""
    try:
        t = Timer()
        t.Interval = interval
        state = {"n": 0}

        def tick(sender, e):
            state["n"] += 1
            bring_to_front(form)
            if state["n"] >= times:
                t.Stop()
        t.Tick += tick
        form._front_timer = t
        t.Start()
    except Exception:
        pass


class _HwndOwner(IWin32Window):
    """Makes Mechanical's main window the owner of our dialog, so the dialog always stays above it."""
    def __init__(self, h):
        self._h = h

    def get_Handle(self):
        return self._h
    Handle = property(get_Handle)


def mechanical_owner():
    try:
        h = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle
        if h != System.IntPtr.Zero:
            return _HwndOwner(h)
    except Exception:
        pass
    return None


class TextForm(Form):
    def __init__(self, title, text):
        self.Text = title
        self.ClientSize = Size(760, 560)
        self.StartPosition = FormStartPosition.CenterScreen
        self.Shown += lambda s, e: bring_to_front(self)
        tb = TextBox()
        tb.Multiline = True
        tb.ReadOnly = True
        tb.ScrollBars = ScrollBars.Both
        tb.WordWrap = False
        tb.Font = Font("Consolas", 9.0)
        tb.Location = Point(10, 10)
        tb.Size = Size(740, 500)
        tb.Text = text
        self.Controls.Add(tb)
        b = Button()
        b.Text = u"Copy"
        b.Location = Point(10, 520)
        b.Size = Size(90, 28)
        b.Click += lambda s, e: Clipboard.SetText(text)
        self.Controls.Add(b)
        c = Button()
        c.Text = u"Close"
        c.Location = Point(660, 520)
        c.Size = Size(90, 28)
        c.DialogResult = DialogResult.OK
        self.Controls.Add(c)


# ---------------------------------------------------------------
# UI
# ---------------------------------------------------------------
class NoWheelCombo(ComboBox):
    """Drop-down whose selection changes ONLY by clicking: the mouse wheel over it does nothing to the
    selection (it scrolls the contact list instead)."""
    scroll_target = None

    def OnMouseWheel(self, e):
        try:
            e.Handled = True
        except Exception:
            pass
        t = self.scroll_target
        if t is not None:
            try:
                pos = t.AutoScrollPosition
                t.AutoScrollPosition = Point(-pos.X, max(0, -pos.Y - e.Delta))
            except Exception:
                pass


class NormalForm(Form):
    def __init__(self, contact_tree, model):
        self.Text = TITLE
        self.Name = FORM_ID
        wa = Screen.PrimaryScreen.WorkingArea
        self.ClientSize = Size(int(0.82 * min(1700, wa.Width - 30)), int(0.82 * min(900, wa.Height - 70)))
        self.MinimumSize = Size(900, 560)
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.Sizable
        self.MaximizeBox = True
        self.MinimizeBox = True
        self._panel = 0
        self._tree_updating = False
        self._body_nodes = []
        self._tree_collapsed = False
        self._ov_collapsed = False
        self._tt = ToolTip()
        self.KeyPreview = True
        self.KeyDown += self._on_key
        self.FormClosed += self._on_closed
        self.Resize += lambda s, e: self._layout()
        self.Shown += lambda s, e: (bring_to_front(self), schedule_front(self))
        self._tree = contact_tree
        self._model = model
        self._sel_keys = set()
        self._cds = []
        self._scene = None
        self._diag = None
        self._dump_info = {}
        self._results_ready = False
        self._last_idx = 0
        self._mesh_stale = []                               # contacts whose mesh changed since their flip command
        self._has_flip = set()                              # keys of contacts that carry a flip command
        base = Form().Font
        self.base_font = base
        self.font_bold = Font(base, FontStyle.Bold)
        self.font_title = Font(base.FontFamily, 9, FontStyle.Bold)
        self._build_select_panel()

    def _btn(self, text, x, y, w, h, handler, bold=False):
        b = Button()
        b.Text = text
        b.Font = self.font_bold if bold else self.base_font
        b.FlatStyle = FlatStyle.Standard
        b.Location = Point(x, y)
        b.Size = Size(w, h)
        if handler is not None:
            b.Click += handler
        return b

    # ---------------- panel 1 ----------------
    def _build_select_panel(self):
        self.Controls.Clear()
        self._panel = 1
        self.Text = TITLE + u"  -  Select Contacts"
        W, H = self.ClientSize.Width, self.ClientSize.Height
        ALL = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
        stale, has_flip = self._scan_mesh_changes()          # quick check, no solver: only contacts with a flip command
        self._mesh_stale = stale
        self._has_flip = has_flip
        stale_keys = set(x["key"] for x in stale)
        RW_ = 320 if stale else 0
        grp = GroupBox()
        grp.Text = (u"Contacts  (check the ones whose normals you want to see;  green = has a flip command, "
                    u"red = mesh changed since)")
        grp.Font = self.font_title
        grp.SetBounds(14, 10, W - 28 - RW_, H - 96)
        grp.Anchor = ALL
        self.Controls.Add(grp)
        if stale:
            rp = GroupBox()
            rp.Text = u"Mesh Check"
            rp.Font = self.font_title
            rp.SetBounds(W - 14 - 308, 10, 308, H - 96)
            rp.Anchor = AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Bottom
            self.Controls.Add(rp)
            tb = TextBox()
            tb.Multiline = True
            tb.ReadOnly = True
            tb.ScrollBars = ScrollBars.Vertical
            tb.Font = self.base_font
            tb.SetBounds(10, 24, 288, H - 96 - 24 - 56)
            tb.Anchor = ALL
            ll = [u"MESH CHANGE DETECTED", u"",
                  u"%d contact(s) with a flip command differ from the current mesh (red in the list):" % len(stale), u""]
            for x in stale:
                ll.append(label_of(x["c"]))
                for (sn, rec, pr) in x["items"]:
                    for p_ in pr:
                        ll.append(u"   - " + p_)
            ll += [u"", u"'Apply Changes' rewrites the flip commands of these contacts for the new mesh (the old "
                        u"record is replaced). Then run Show Normals: a contact turns green when it is right."]
            tb.Text = u"\r\n".join(ll)
            rp.Controls.Add(tb)
            bap = self._btn(u"Apply Changes", 10, H - 96 - 44, 288, 32, self._on_apply_mesh_changes, True)
            bap.BackColor = Color.FromArgb(255, 214, 200)
            bap.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
            rp.Controls.Add(bap)
        self.tv = TreeView()
        self.tv.SetBounds(10, 24, W - 48 - RW_, H - 96 - 34)
        self.tv.Anchor = ALL
        self.tv.CheckBoxes = True
        self.tv.ShowLines = True
        self.tv.AfterCheck += self._on_node_check
        grp.Controls.Add(self.tv)
        self._ordered, self._groups = [], []
        for (gname, contacts) in self._tree:
            gn = TreeNode(gname + u"  [" + str(len(contacts)) + u" contacts]")
            gn.Tag = u"GROUP"
            gn.NodeFont = self.font_bold
            self._groups.append(gn)
            for c in contacts:
                key = ukey_of(c, len(self._ordered))
                ch = TreeNode(label_of(c) + (u"   [Mesh Changed]" if key in stale_keys else u""))
                ch.Tag = u"CONTACT"
                ch.NodeFont = self.base_font
                if key in stale_keys:
                    ch.ForeColor = Color.FromArgb(190, 30, 30)
                elif key in has_flip:                                 # already carries a flip command
                    ch.ForeColor = Color.FromArgb(20, 120, 40)
                ch.Checked = (key in self._sel_keys)
                self._ordered.append((ch, c))
                gn.Nodes.Add(ch)
            self.tv.Nodes.Add(gn)
            gn.Expand()
        self._updating = False

        LB = AnchorStyles.Bottom | AnchorStyles.Left
        RB = AnchorStyles.Bottom | AnchorStyles.Right
        hint = Label()
        hint.Text = (u"Show Normals adds ONE temporary command to the analysis and starts the solution; the command writes "
                     u"all contact/target elements of the model and ends the solver run, then it is removed.")
        hint.Font = self.base_font
        hint.ForeColor = Color.DimGray
        hint.AutoSize = True
        hint.Location = Point(16, H - 80)
        hint.Anchor = LB
        self.Controls.Add(hint)

        yb = H - 40
        b1 = self._btn(u"Select All", 14, yb, 100, 30, lambda s, e: self._set_all(True))
        b2 = self._btn(u"Clear", 124, yb, 80, 30, lambda s, e: self._set_all(False))
        b4 = self._btn(u"Refresh List", 214, yb, 110, 30, self._on_refresh_list)
        b7 = self._btn(u"Select Flipped", 334, yb, 120, 30, self._on_select_flipped)
        self._tt.SetToolTip(b7, u"Check ONLY the contacts that have a flip command (green and red ones)")
        bc = self._btn(u"Cancel", W - 14 - 124 - 10 - 100, yb, 100, 30, lambda s, e: self.Close())
        b3 = self._btn(u"Show Normals >", W - 14 - 124, yb, 124, 30, self._on_show, True)
        b5 = self._btn(u"Last Results >", W - 14 - 124 - 10 - 100 - 10 - 120, yb, 120, 30, self._on_last_results)
        b5.Enabled = bool(self._results_ready and self._cds)
        for b, an in ((b1, LB), (b2, LB), (b4, LB), (b7, LB), (bc, RB), (b3, RB), (b5, RB)):
            b.Anchor = an
            self.Controls.Add(b)
        bc.DialogResult = DialogResult.Cancel
        self.CancelButton = bc


    def _edge_node_ids(self, mesh, c):
        """Mesh node numbers on the scoped edges of a contact: {'contact': set, 'target': set}."""
        out = {}
        for side, loc in (("contact", c.SourceLocation), ("target", c.TargetLocation)):
            ids = set()
            for gid in list(loc.Ids):
                ids.update(list(mesh.MeshRegionById(gid).NodeIds))
            out[side] = ids
        return out

    def _edge_node_xy(self, mesh, c):
        """{'contact': {node: (x, y)}, 'target': {...}} of the scoped edges (mesh coordinates)."""
        out = {}
        for side, ids in self._edge_node_ids(mesh, c).items():
            d = {}
            for nid in ids:
                nd = mesh.NodeById(nid)
                d[nid] = (nd.X, nd.Y)
            out[side] = d
        return out

    def _scan_mesh_changes(self):
        """Quick check WITHOUT the solver (runs whenever page 1 is built): only contacts that carry a flip command with
        a mesh record are compared with the current mesh (node count / numbers of the edges, flipped nodes).
        -> (list of changed contacts, set of keys of contacts that have a flip command)."""
        stale, has_flip, mesh, idx = [], set(), None, 0
        for (gname, contacts) in self._tree:
            for c in contacts:
                key = ukey_of(c, idx)
                idx += 1
                try:
                    snips = find_flip_snippets(c)
                    if not snips:
                        continue
                    has_flip.add(key)
                    recs = [(sn, parse_cnc_record(sn.Input)) for sn in snips]
                    recs = [(sn, r) for (sn, r) in recs if r is not None]
                    if not recs:
                        continue
                    if mesh is None:
                        mesh = ExtAPI.DataModel.MeshDataByName("Global")
                    cur = self._edge_node_ids(mesh, c)
                    items = []
                    for sn, rec in recs:
                        def node_xy(side, nid, cur=cur, mesh=mesh):
                            if nid not in cur.get(side, ()):
                                return None
                            nd = mesh.NodeById(nid)
                            return (nd.X, nd.Y)
                        pr = check_mesh_record(rec, cur, node_xy)
                        if pr:
                            items.append((sn, rec, pr))
                    if items:
                        stale.append({"c": c, "key": key, "items": items})
                except Exception:
                    pass
        return stale, has_flip

    def _on_apply_mesh_changes(self, s, e):
        """Rewrites the flip commands of the contacts whose mesh changed: the coordinate windows of the command are
        applied to the NEW mesh, the nodes inside them are listed again and the mesh record is REPLACED by the new one
        (old record out, new record in). Afterwards Show Normals shows whether the result is right (green)."""
        stamp = System.DateTime.Now.ToString("yyyy-MM-dd HH:mm").replace(" ", "T")
        ok, bad = [], []
        try:
            mesh = ExtAPI.DataModel.MeshDataByName("Global")
        except Exception as ex:
            MessageBox.Show(u"The mesh could not be read: %s" % str(ex), u"Apply Changes",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        for x in self._mesh_stale:
            c = x["c"]
            try:
                cur_xy = self._edge_node_xy(mesh, c)
            except Exception as ex:
                bad.append(u"%s: %s" % (label_of(c), str(ex)))
                continue
            for (sn, rec, pr) in x["items"]:
                try:
                    runs, sigs, report = remap_mesh_record(rec, cur_xy)
                    if runs is None:
                        bad.append(u"%s ('%s'): %s" % (label_of(c), sn.Name, report))
                        continue
                    sn.Input = rewrite_record(sn.Input, record_lines(rec["id"], rec["created"] or stamp, stamp,
                                                                     rec.get("unit"), sigs, runs))
                    ok.append(u"%s ('%s'): %s" % (label_of(c), sn.Name, report))
                except Exception as ex:
                    bad.append(u"%s ('%s'): %s" % (label_of(c), sn.Name, str(ex)))
        msg = u""
        if ok:
            msg += u"Applied to %d command(s):\n  " % len(ok) + u"\n  ".join(ok[:12]) + (u"\n  ..." if len(ok) > 12 else u"")
        if bad:
            msg += (u"\n\n" if msg else u"") + u"NOT applied (%d):\n  " % len(bad) + u"\n  ".join(bad[:12])
        if ok and not bad:
            msg += u"\n\nRun Show Normals to check the result."
        MessageBox.Show(msg or u"Nothing to apply.", u"Apply Changes", MessageBoxButtons.OK,
                        MessageBoxIcon.Warning if bad else MessageBoxIcon.Information)
        self._build_select_panel()                                # re-scan: applied contacts are no longer red

    def _on_select_flipped(self, s, e):
        """Checks ONLY the contacts that carry a flip command (green = fine / red = mesh changed), so Show Normals
        can verify exactly those, e.g. after 'Apply changes'."""
        keys = []
        for i, (n, c) in enumerate(self._ordered):
            if ukey_of(c, i) in self._has_flip:
                keys.append(n)
        if not keys:
            MessageBox.Show(u"No contact has a flip command written by this tool.", u"Select Flipped",
                            MessageBoxButtons.OK, MessageBoxIcon.Information)
            return
        self._set_all(False)
        self._updating = True
        for n in keys:
            n.Checked = True
        for g in self._groups:
            g.Checked = all(ch.Checked for ch in g.Nodes)
        self._updating = False

    def _on_last_results(self, s, e):
        """Back to the last results WITHOUT calculating again (arrows, selections, flip actions and hidden bodies
        are kept). Use Show Normals to read the solver state again."""
        if not self._results_ready or not self._cds or self._scene is None:
            return
        self._build_result_panel(self._lines)

    def _on_refresh_list(self, s, e):
        """Re-read the contacts of the model (the window stays open while you keep working in Mechanical)."""
        keep = set(self._sel_keys)
        for i, (n, c) in enumerate(self._ordered):
            if n.Checked:
                keep.add(ukey_of(c, i))
        self._tree = build_contact_tree(self._model.Connections)
        self._sel_keys = keep
        self._build_select_panel()

    def _on_node_check(self, sender, e):
        if self._updating:
            return
        self._updating = True
        node = e.Node
        if str(node.Tag) == u"GROUP":
            for ch in node.Nodes:
                ch.Checked = node.Checked
        elif node.Parent is not None:
            node.Parent.Checked = all(ch.Checked for ch in node.Parent.Nodes)
        self._updating = False

    def _set_all(self, v):
        self._updating = True
        for g in self._groups:
            g.Checked = v
            for ch in g.Nodes:
                ch.Checked = v
        self._updating = False

    def _solver_dump(self, sel, lines, S_by_key):
        """ONE temporary command on the analysis writes all contact / target elements of the model and ends the
        solver run (no command under the contacts). The element groups are then assigned to the selected contacts
        by position (match_groups). The temporary command is removed afterwards (also when something fails).
        Returns ({key: {'contact': [recs], 'target': [recs], 'meta': ...}} with MESH node ids, [keys without data])."""
        an = find_analysis(self._model)
        sol = an.Solution
        wd = an.WorkingDir
        lines.append(u"Solver working directory: %s" % wd)
        keys = [k for k, c in sel]
        try:                                   # leftovers of an interrupted run (a stray exit command ends real solves!)
            for ch in list(an.Children):
                if "CommandSnippet" in ch.GetType().ToString() and ch.Name in (EXIT_NAME, DUMPALL_NAME):
                    ch.Delete()
        except Exception:
            pass
        for old_f in System.IO.Directory.GetFiles(wd, u"cnc_all*.txt"):
            try:
                System.IO.File.Delete(old_f)
            except Exception:
                pass
        temp = []
        wf = WaitForm()
        solve_msg = u""
        t1 = time.time()
        try:
            for nm_, text_ in ((DUMPALL_NAME, dump_all_snippet()), (EXIT_NAME, EXIT_ALL_TEXT)):
                sn = an.AddCommandSnippet()
                temp.append((sn, nm_))
                sn.Name = nm_
                sn.Input = text_
                if SequenceSelectionType is not None:
                    try:
                        sn.StepSelectionMode = SequenceSelectionType.ByNumber
                        sn.StepNumber = 1
                    except Exception:
                        pass
            t2 = time.time()
            lines.append(u"Timing: 2 temporary commands (dump + exit) added to the analysis in %.1f s" % (t2 - t1))
            wf.Show(self)
            wf.set_text(u"The solver is running and stops right after the model has been read.\n"
                        u"Please wait ...")
            try:
                sol.Solve(True)                      # blocks; the command ends the solver run
            except Exception as ex:
                solve_msg = str(ex)                  # an ended solver run is reported as an error
                lines.append(u"Solve reported: %s" % solve_msg)
            lines.append(u"Timing: solver run %.1f s" % (time.time() - t2))
        finally:
            try:
                wf.Close()
            except Exception:
                pass
            t3 = time.time()
            for sn_, nm_ in temp:
                try:
                    sn_.Delete()
                except Exception as ex:
                    lines.append(u"!! could not remove the temporary analysis command '%s' (%s) - delete it "
                                 u"by hand, otherwise real solves end at once" % (nm_, str(ex)))
            lines.append(u"Temporary commands removed (%.1f s)" % (time.time() - t3))
        files = []
        while True:
            pth = System.IO.Path.Combine(wd, u"cnc_all%d.txt" % (len(files) + 1))
            if not System.IO.File.Exists(pth):
                break
            files.append(pth)
        if not files:
            try:
                status = u", solution status: %s" % str(sol.Status)
            except Exception:
                status = u""
            raise Exception(u"The solver wrote no cnc_all1.txt%s%s\n\n%s" % (
                (u" - Solve reported: %s" % solve_msg) if solve_msg else u"", status, solver_report(wd)))
        t4 = time.time()
        all_lines = []
        sizes = []
        for pth in files:
            all_lines.extend(read_text_lines(pth))
            sizes.append(System.IO.FileInfo(pth).Length // 1024)
        items, hdr, done, single = parse_dump_all(all_lines)
        expected = (hdr[0] + hdr[1]) if hdr else None
        complete = (len(items) + single == expected) if expected is not None else done
        if not complete:
            tail = u" | ".join(l.strip()[:60] for l in all_lines[-2:])
            raise Exception(u"The solver output is incomplete: %d elements read (+ %d single-node), %s announced, "
                            u"end marker %s; %d file(s) of %s KB; last lines: %s\n\n%s" % (
                                len(items), single, expected if expected is not None else u"?",
                                u"seen" if done else u"missing", len(files),
                                u"/".join(str(x) for x in sizes[:6]), tail, solver_report(wd)))
        lines.append(u"Timing: read %d elements from %d file(s) in %.1f s%s%s" % (
            len(items), len(files), time.time() - t4,
            u" (%d single-node pilot elements skipped)" % single if single else u"",
            u"" if done else u" (end marker missing, but all elements are there)"))
        t5 = time.time()
        xy = self._scene["xy"]
        groups = group_contacts(items)
        assign, minfo = match_groups(groups, dict((k, S_by_key[k]) for k in keys), xy)
        lines.append(u"%d contact regions found in the solver; %d of %d selected contacts matched by position "
                     u"(unit scale %s, %d unmatched group(s): %s) - %.1f s" % (
                         len(groups), len(assign), len(keys), minfo["scale"], len(minfo["unmatched"]),
                         u", ".join(u"types %d/%d" % g for g in minfo["unmatched"][:8]) or u"-",
                         time.time() - t5))
        if not assign:
            raise Exception(u"None of the %d selected contacts could be matched to the %d contact regions the solver "
                            u"wrote (compared by position of the edge nodes)" % (len(keys), len(groups)))
        out, missing = {}, []
        self._dump_info = {}
        for k, c in sel:
            if k not in assign:
                missing.append(k)
                continue
            gk = assign[k]
            d = {"contact": groups[gk]["contact"], "target": groups[gk]["target"], "hdr": gk}
            S = S_by_key[k]
            union = dict((n, xy[n]) for n in (S["contact"] | S["target"]) if n in xy)
            res = {}
            maps = {}
            scs = []
            for side in ("contact", "target"):
                xy_side = dict((n, xy[n]) for n in S[side] if n in xy)
                recs, n_nodes, n_map, sc, maps[side] = dump_to_recs(d[side], xy_side, union)
                res[side] = recs
                scs.append(sc)
                lines.append(u"%s %s: %d solver elements, %d/%d nodes matched to the mesh (unit scale %s)"
                             % (label_of(c), side, len(d[side]), n_map, n_nodes, sc))
            d["unit"] = ([v for v in scs if v] or [None])[0]        # mesh coordinate = solver coordinate * unit
            real_of = {}
            rc = {"contact": {}, "target": {}}
            for side in ("contact", "target"):
                for it in d[side]:
                    real_of[it["e"]] = it.get("real", 0)
                    rc[side][it.get("real", 0)] = rc[side].get(it.get("real", 0), 0) + 1
            res["meta"] = {"real": real_of, "hdr": gk}
            lines.append(u"%s real constants: contact %s | target %s | contact type / target type = %s" % (
                label_of(c), u" ".join(u"%d x%d" % kv for kv in sorted(rc["contact"].items())) or u"-",
                u" ".join(u"%d x%d" % kv for kv in sorted(rc["target"].items())) or u"-", gk))
            out[k] = res
            self._dump_info[k] = (d, maps)
        return out, missing

    def _on_show(self, s, e):
        sel = []
        for i, (n, c) in enumerate(self._ordered):
            if n.Checked:
                sel.append((ukey_of(c, i), c))
        if not sel:
            MessageBox.Show(u"Please select at least one contact.", u"Warning",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        self._sel_keys = set(k for k, c in sel)
        lines = []
        self._cds = []
        self._diag = None
        self._results_ready = False
        self._last_idx = 0
        Cursor.Current = Cursors.WaitCursor
        t_start = time.time()
        try:
            self._scene = load_scene(lines.append)
            lines.append(u"Timing: model view loaded in %.1f s" % (time.time() - t_start))
            t_ns = time.time()
            S_by_key = dict((k, contact_node_sets(self._scene, c, lines.append)) for k, c in sel)
            lines.append(u"Timing: edge node sets of %d contacts in %.1f s" % (len(sel), time.time() - t_ns))
            t_dump = time.time()
            # 1) the solver itself tells us the contact / target elements and their node order
            try:
                dump, failed = self._solver_dump(sel, lines, S_by_key)
                for k_ in failed:
                    dump[k_] = {"contact": [], "target": []}
                if failed:
                    lines.append(u"The solver wrote nothing for %d contact(s): %s" % (
                        len(failed), u", ".join(str(k_) for k_ in failed)))
            except Exception as ex:
                Cursor.Current = Cursors.Default
                MessageBox.Show(u"The contact elements could not be read from the solver:\n\n%s" % str(ex),
                                u"Solver", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                return
            lines.append(u"Timing: solver part in total %.1f s" % (time.time() - t_dump))
            t_proc = time.time()
            node2b = {}
            for b_ in self._scene["bodies"]:
                ns = set()
                for e_ in b_["elems"]:
                    ns.update(e_[2])
                for n_ in ns:
                    node2b.setdefault(n_, []).append(b_)
            dummy = {"visible": True, "name": u"?"}
            for i, (k, c) in enumerate(sel):
                self._cds.append({"c": c, "key": k, "segs": [], "sel": set(), "side": u"none",
                                  "idx": i, "note": u"", "warn": u"", "lids": None, "lbox": None,
                                  "fixed": bool(find_flip_snippets(c))})
            learn = {}
            for k_ in S_by_key:                                    # learn the numbering pattern from all contacts first
                if dump.get(k_):
                    split_symmetric(dump[k_], S_by_key[k_], learn)
            for cd in sorted(self._cds, key=lambda x: x["key"]):
                c = cd["c"]
                S = S_by_key[cd["key"]]
                inv_b = []
                for side in ("contact", "target"):
                    for n_ in S[side]:
                        for b_ in node2b.get(n_, []):
                            if all(b_ is not x_ for x_ in inv_b):
                                inv_b.append(b_)
                cd["lids"] = set(id(b_) for b_ in inv_b) or None
                if inv_b:
                    bxs = [e_[5] for b_ in inv_b for e_ in b_["elems"]]
                    cd["lbox"] = (min(q[0] for q in bxs), max(q[1] for q in bxs),
                                  min(q[2] for q in bxs), max(q[3] for q in bxs))
                emap = edge_owners(inv_b, S["contact"] | S["target"])
                segs = dat_segments(dump[cd["key"]], self._scene["xy"], node2b, dummy, emap)
                cps, cpinfo = split_symmetric(dump[cd["key"]], S, learn)
                for sg in segs:
                    sg["cp"] = sg["enum"] in cps
                cd["cps"] = cps
                cd["S"] = S
                flag_reference(segs)
                n_cp = sum(1 for sg in segs if sg["cp"])
                cd["note"] = u"| exact: read from the solver" + (
                    u" | %d symmetric counterpart element(s) not drawn" % n_cp if n_cp else u"")
                inw = inward_info(segs)
                if inw and (inw["contact"] or inw["target"]):
                    cd["note"] += (u" | normal points into its own body: contact %d, target %d"
                                   % (inw["contact"], inw["target"]))
                if not segs:
                    cd["warn"] = (u"The solver wrote nothing for this contact (suppressed, or no contact "
                                  u"elements?)" if cd["key"] in failed
                                  else u"NO contact/target elements were read for this contact")
                cd["segs"] = segs
                dc = direction_counts(segs)
                dtxt = u"; ".join(
                    u"%s: %s" % (sd, u" ".join(u"%s %d" % (k_, v_) for k_, v_ in sorted(dc[sd].items())))
                    for sd in ("contact", "target") if dc[sd])
                if dtxt:
                    cd["note"] += u" | directions (screen) " + dtxt
                nc = sum(1 for x in segs if x["side"] == "contact")
                lines.append(u"%s: %d contact / %d target elements in the solver (%d symmetric counterpart hidden), "
                             u"%d to flip (not the opposite of the other side); %s"
                             % (label_of(c), nc, len(segs) - nc, n_cp, sum(1 for x in segs if x["rev"]), dtxt))
                lines.append(u"%s: counterpart split - %s" % (label_of(c), cpinfo))
            lines.append(u"Timing: processing the results %.1f s" % (time.time() - t_proc))
        finally:
            Cursor.Current = Cursors.Default
        lines.append(u"Timing: TOTAL %.1f s" % (time.time() - t_start))
        self._build_result_panel(lines)
        bring_to_front(self)
        schedule_front(self)

    # ---------------- panel 2 ----------------
    def _build_result_panel(self, lines):
        old_cur = getattr(self, "_cur", None)                 # rebuilt from the cache ("Last results"): keep the hidden
        if old_cur is not None and getattr(self, "view", None) is not None:      # bodies of the contact that was shown
            old_cur["hid"] = list(self.view.hidden)
        if self._scene is not None:
            for b_ in self._scene["bodies"]:
                b_["visible"] = True                           # each contact restores its own hidden bodies when selected
        self.Controls.Clear()
        self._panel = 0                                   # layout enabled at the end
        self._split_ready = False
        self.Text = TITLE + u"  -  Normals"

        # three resizable columns: contact list | contact view | whole model
        self.area = Panel()
        self.Controls.Add(self.area)
        self.split1 = SplitContainer()
        self.split1.Orientation = Orientation.Vertical
        self.split1.Dock = DockStyle.Fill
        self.split1.SplitterWidth = 8
        self.split1.FixedPanel = FixedPanel.Panel1
        self.area.Controls.Add(self.split1)
        self.split2 = SplitContainer()
        self.split2.Orientation = Orientation.Vertical
        self.split2.Dock = DockStyle.Fill
        self.split2.SplitterWidth = 8
        self.split1.Panel2.Controls.Add(self.split2)
        self.strip_ov = self._btn(u"\u00ab", 0, 0, 24, 30, lambda s, e: self._set_ov_collapsed(False), True)
        self.strip_ov.Dock = DockStyle.Right                  # added after split2, so it docks first
        self.strip_ov.Visible = False
        self._tt.SetToolTip(self.strip_ov, u"Show the whole-model view")
        self.split1.Panel2.Controls.Add(self.strip_ov)

        g1 = GroupBox()
        g1.Text = u"Contacts"
        g1.Font = self.font_title
        g1.Dock = DockStyle.Fill
        self.g1 = g1
        self.split1.Panel1.Controls.Add(g1)
        self.hdr = []
        for txt in (u"Contact", u"Flip Action  v", u"Sel", u"To Flip"):
            hl = Label()
            hl.Text = txt
            hl.Font = self.font_bold
            hl.AutoSize = True
            g1.Controls.Add(hl)
            self.hdr.append(hl)
        self.rows = Panel()
        self.rows.AutoScroll = True
        self.rows.BorderStyle = BorderStyle.FixedSingle
        self.rows.BackColor = Color.White
        g1.Controls.Add(self.rows)
        self._rows = []
        self._cur = None
        for i, cd in enumerate(self._cds):
            self._rows.append(self._make_row(i, cd))
        n_flag = sum(1 for cd in self._cds if cd.get("n_rev", 0) > 0)
        n_fix = sum(1 for cd in self._cds if cd.get("fixed") and cd.get("n_rev", 0) == 0)
        g1.Text = (u"Contacts" + (u"   -   %d with arrows to flip" % n_flag if n_flag else u"")
                   + (u"   -   %d fixed (green)" % n_fix if n_fix else u""))
        g1.Resize += lambda s, e: self._layout_g1()

        g2 = GroupBox()
        g2.Text = u"Contact View  (contact = red, target = blue, should be flipped = orange)"
        g2.Font = self.font_title
        g2.Dock = DockStyle.Fill
        self.g2 = g2
        self.split2.Panel1.Controls.Add(g2)
        self.split3 = SplitContainer()                    # body tree | contact view
        self.split3.Orientation = Orientation.Vertical
        self.split3.SplitterWidth = 6
        self.split3.FixedPanel = FixedPanel.Panel1
        g2.Controls.Add(self.split3)
        self._split3_ready = False
        self.bodytree = TreeView()
        self.bodytree.Dock = DockStyle.Fill
        self.bodytree.CheckBoxes = True
        self.bodytree.ShowLines = True
        self.bodytree.HideSelection = False
        self.split3.Panel1.Controls.Add(self.bodytree)
        tcap = Panel()
        tcap.Dock = DockStyle.Top
        tcap.Height = 24
        tlab = Label()
        tlab.Text = u"Bodies  (uncheck = hide)"
        tlab.Font = self.font_bold
        tlab.AutoSize = False
        tlab.SetBounds(0, 4, 150, 18)
        tcap.Controls.Add(tlab)
        self.btn_tree_close = self._btn(u"\u00ab", 0, 0, 26, 22, lambda s, e: self._set_tree_collapsed(True), True)
        self._tt.SetToolTip(self.btn_tree_close, u"Hide the body tree")
        tcap.Controls.Add(self.btn_tree_close)
        tcap.Resize += lambda s, e: self.btn_tree_close.SetBounds(max(0, tcap.Width - 28), 0, 26, 22)
        self.split3.Panel1.Controls.Add(tcap)
        self.strip_tree = self._btn(u"\u00bb", 10, 24, 24, 100, lambda s, e: self._set_tree_collapsed(False), True)
        self.strip_tree.Visible = False
        self._tt.SetToolTip(self.strip_tree, u"Show the body tree")
        g2.Controls.Add(self.strip_tree)
        self.view = SceneView(self)
        self.view.scene = self._scene
        self.view.hidden = []
        self.view.Dock = DockStyle.Fill
        self.split3.Panel2.Controls.Add(self.view)
        self._build_body_tree()
        self.bodytree.AfterCheck += self._on_tree_check
        v = self.view
        row1 = [(u"Fit", lambda s, e: v.fit(), 70), (u"Fit All", lambda s, e: v.fit_all(), 80),
                (u"Local: ON", self._on_local_toggle, 90),
                (u"Clear Sel", self._on_clear_sel, 90),
                (u"Select To Flip", lambda s, e: v.select_reversed(), 110)]
        row2 = [(u"Hide Body", self._on_hide_toggle, 116), (u"Undo", lambda s, e: v.undo_hide(), 74),
                (u"Show All", lambda s, e: v.show_all(), 84),
                (u"Isolate Wrong", self._on_isolate_wrong, 106)]
        self._row_btns = ([], [])
        for ri, row in enumerate((row1, row2)):
            x = 10
            for t, h, w in row:
                bt = self._btn(t, x, 0, w, 30, h)
                bt.Anchor = AnchorStyles.Bottom | AnchorStyles.Left
                g2.Controls.Add(bt)
                self._row_btns[ri].append(bt)
                if t == u"Hide Body":
                    self.btn_hide = bt
                if t == u"Local: ON":
                    self.btn_local = bt
                if t == u"Isolate Wrong":
                    self._tt.SetToolTip(bt, u"Keep only the bodies of the side (target / contact) that has wrong "
                                            u"(orange) normals - the whole side, all its arrows - and hide all "
                                            u"other bodies")
                x += w + 6
        g2.Resize += lambda s, e: self._layout_g2()

        g3 = GroupBox()
        g3.Text = u"Whole Model  (click / drag = move the view)"
        g3.Font = self.font_title
        g3.Dock = DockStyle.Fill
        self.g3 = g3
        self.split2.Panel2.Controls.Add(g3)
        self.overview = OverviewView(self)
        self.overview.scene = self._scene
        self.overview.main = self.view
        g3.Controls.Add(self.overview)
        self.btn_ov_close = self._btn(u"\u00bb", 0, 0, 26, 20, lambda s, e: self._set_ov_collapsed(True), True)
        self._tt.SetToolTip(self.btn_ov_close, u"Hide the whole-model view")
        g3.Controls.Add(self.btn_ov_close)
        g3.Resize += lambda s, e: self._layout_g3()

        self._lines = list(lines)
        self.status = Label()
        self.status.Font = self.base_font
        self.status.AutoEllipsis = True
        self.status.Cursor = Cursors.Hand
        self.status.ForeColor = Color.DimGray
        self.status.Text = u"%d log messages - click here to open the full log" % len(self._lines)
        self.status.Click += lambda s, e: TextForm(u"Log", u"\r\n".join(self._lines)).ShowDialog(self)
        self.Controls.Add(self.status)

        self.btn_back = self._btn(u"< Back", 14, 0, 90, 30, lambda s, e: self._build_select_panel())
        self.btn_close = self._btn(u"Close", 0, 0, 90, 30, lambda s, e: self.Close())
        self.btn_close.DialogResult = DialogResult.Cancel
        self.CancelButton = self.btn_close
        self.btn_flip = self._btn(u"Flip Selected", 0, 0, 124, 30, self._on_flip, True)
        for b in (self.btn_back, self.btn_close, self.btn_flip):
            self.Controls.Add(b)

        self._panel = 2
        self._results_ready = True
        self._layout()
        if self._tree_collapsed:                           # remembered from before ("Last results", new calculation)
            self._set_tree_collapsed(True)
        if self._ov_collapsed:
            self._set_ov_collapsed(True)
        self._update_hide_button()
        if self._cds:
            self._select_row(self._last_idx if 0 <= self._last_idx < len(self._cds) else 0)

    def _layout(self):
        if self._panel != 2:
            return
        W, H = self.ClientSize.Width, self.ClientSize.Height
        self.area.SetBounds(14, 10, W - 28, H - 60)
        if not self._split_ready and self.area.Width > 700:
            try:
                self.split1.Panel1MinSize = 260
                self.split1.Panel2MinSize = 420
                self.split1.SplitterDistance = 380
                self.split2.Panel1MinSize = 260
                self.split2.Panel2MinSize = 200
                self.split2.SplitterDistance = int(self.split2.Width * 0.58)
                self._split_ready = True
            except Exception:
                pass
        yb = H - 40
        self.btn_back.SetBounds(14, yb, 90, 30)
        self.btn_flip.SetBounds(W - 14 - 124, yb, 124, 30)
        self.btn_close.SetBounds(W - 14 - 124 - 10 - 90, yb, 90, 30)
        self.status.SetBounds(112, yb + 6, max(100, W - 14 - 124 - 10 - 90 - 10 - 112), 22)
        self._layout_g1()
        self._layout_g2()
        self._layout_g3()

    def _layout_g1(self):
        W, H = self.g1.Width, self.g1.Height
        self.rows.SetBounds(10, 48, max(120, W - 20), max(50, H - 58))
        RW = self.rows.ClientSize.Width
        for r in self._rows:
            r["row"].Width = RW
            r["name"].Width = max(40, RW - 214)
            r["cb"].Left = max(46, RW - 204)
            r["cnt"].Left = max(100, RW - 94)
            r["rev"].Left = max(140, RW - 46)
        left = self.rows.Left
        self.hdr[0].Location = Point(left + 4, 28)
        self.hdr[1].Location = Point(left + max(46, RW - 204), 28)
        self.hdr[2].Location = Point(left + max(100, RW - 94), 28)
        self.hdr[3].Location = Point(left + max(140, RW - 46), 28)

    def _layout_g2(self):
        W, H = self.g2.Width, self.g2.Height
        hs = max(50, H - 24 - 80)
        x0 = 10
        if self._tree_collapsed:
            self.strip_tree.SetBounds(10, 24, 24, hs)
            self.strip_tree.Visible = True
            x0 = 10 + 24 + 4
        else:
            self.strip_tree.Visible = False
        self.split3.SetBounds(x0, 24, max(50, W - x0 - 10), hs)
        if (not self._tree_collapsed) and not self._split3_ready and self.split3.Width > 520:
            try:
                self.split3.Panel1MinSize = 110
                self.split3.Panel2MinSize = 240
                self.split3.SplitterDistance = 200
                self._split3_ready = True
            except Exception:
                pass
        for ri, y in ((0, H - 74), (1, H - 38)):
            for bt in self._row_btns[ri]:
                bt.Top = y

    def _layout_g3(self):
        W, H = self.g3.Width, self.g3.Height
        self.overview.SetBounds(10, 24, max(50, W - 20), max(50, H - 34))
        self.btn_ov_close.SetBounds(max(10, W - 32), 0, 26, 20)
        self.btn_ov_close.BringToFront()

    FLIP_CHOICES = (u"none", u"both", u"contact", u"target", u"element", u"remove")

    def _make_row(self, i, cd):
        row = Panel()
        row.Location = Point(0, i * 36)
        row.Size = Size(340, 34)
        row.BackColor = Color.White
        name = Label()
        name.Text = label_of(cd["c"])
        name.AutoEllipsis = True
        name.Font = self.base_font
        name.Location = Point(6, 7)
        name.Size = Size(172, 20)
        cb = NoWheelCombo()
        cb.scroll_target = self.rows
        cb.DropDownStyle = ComboBoxStyle.DropDownList
        cb.FlatStyle = FlatStyle.Standard
        cb.Font = self.base_font
        for t in self.FLIP_CHOICES:
            cb.Items.Add(t.capitalize())                       # shown as None / Both / Contact / ...
        cb.SelectedItem = cd["side"].capitalize()
        cb.Location = Point(182, 5)
        cb.Size = Size(102, 24)
        cnt = Label()
        cnt.Text = unicode(len(cd["sel"]))
        cnt.Font = self.base_font
        cnt.Location = Point(250, 7)
        cnt.Size = Size(40, 20)
        n_rev = sum(1 for x in cd["segs"] if x["rev"])
        cd["n_rev"] = n_rev
        rev = Label()
        fixed = bool(cd.get("fixed")) and not n_rev
        rev.Text = unicode(n_rev) if n_rev else (u"OK" if fixed else u"-")
        rev.Font = self.font_bold if (n_rev or fixed) else self.base_font
        rev.ForeColor = (Color.FromArgb(200, 80, 0) if n_rev else
                         (Color.FromArgb(20, 130, 50) if fixed else Color.Gray))
        rev.Location = Point(296, 7)
        rev.Size = Size(40, 20)
        for ctl in (name, cb, cnt, rev):
            row.Controls.Add(ctl)

        def pick(sender, e, idx=i):
            self._select_row(idx)

        def commit(sender, e, idx=i):
            cd2 = self._cds[idx]
            cd2["side"] = unicode(cb.SelectedItem).lower()
            self._style_row(idx)
            self._select_row(idx)

        for ctl in (row, name, cnt, rev):
            ctl.Click += pick
        cb.DropDown += pick
        cb.SelectionChangeCommitted += commit
        self.rows.Controls.Add(row)
        r = {"row": row, "name": name, "cb": cb, "cnt": cnt, "rev": rev}
        self._rows_style(r, cd)
        return r

    def _rows_style(self, r, cd):
        active = cd["side"] != u"none"
        r["cb"].BackColor = Color.FromArgb(255, 226, 170) if active else Color.White
        n_rev = cd.get("n_rev", 0)
        flagged = n_rev > 0
        pending = bool(cd.get("pending"))
        fixed = bool(cd.get("fixed")) and not flagged
        r["name"].Font = self.font_bold if (active or flagged or fixed or pending) else self.base_font
        r["name"].ForeColor = (Color.FromArgb(100, 40, 160) if pending else
                               (Color.FromArgb(170, 60, 0) if flagged else
                                (Color.FromArgb(20, 110, 40) if fixed else Color.Black)))
        r["rev"].Text = u"NEW" if pending else (unicode(n_rev) if n_rev else (u"OK" if fixed else u"-"))
        r["rev"].Font = self.font_bold if (pending or flagged or fixed) else self.base_font
        r["rev"].ForeColor = (Color.FromArgb(110, 40, 170) if pending else
                              (Color.FromArgb(200, 80, 0) if flagged else
                               (Color.FromArgb(20, 130, 50) if fixed else Color.Gray)))

    def _row_color(self, idx, selected):
        """Violet = a flip command was written but is not reflected yet; orange = reversed arrows; green = has a flip
        command and nothing is left to flip. The selected row is darker."""
        cd = self._cds[idx]
        if cd.get("pending"):
            return Color.FromArgb(200, 172, 235) if selected else Color.FromArgb(228, 212, 246)
        if cd.get("n_rev", 0) > 0:
            return Color.FromArgb(250, 200, 140) if selected else Color.FromArgb(255, 228, 192)
        if cd.get("fixed"):
            return Color.FromArgb(160, 222, 150) if selected else Color.FromArgb(214, 240, 208)
        return Color.FromArgb(205, 225, 248) if selected else Color.White

    def _style_row(self, idx):
        self._rows_style(self._rows[idx], self._cds[idx])

    def _select_row(self, idx):
        cd = self._cds[idx]
        self._last_idx = idx
        for k, r in enumerate(self._rows):
            r["row"].BackColor = self._row_color(k, k == idx)
        if cd is self._cur:
            return
        if self._cur is not None:
            self._cur["hid"] = list(self.view.hidden)        # this contact keeps ITS hidden bodies ...
        self.view.put_hidden(cd.get("hid"))                  # ... the next one shows its own (nothing hidden by default)
        self.sync_tree()
        self._cur = cd
        self.view.segs = cd["segs"]
        self.view.sel = cd["sel"]
        self.view.note = cd["note"]
        self.view.warn = self._warn_of(cd)
        self.view.local_ids = cd.get("lids")
        self.view.local_box = cd.get("lbox")
        self.overview.segs = cd["segs"]
        self._style_tree()
        self.view.request_fit()

    def selection_changed(self):
        cd = self._cur
        if cd is not None:
            self._rows[cd["idx"]]["cnt"].Text = str(len(cd["sel"]))

    def note(self, text):
        t = text.strip()
        try:
            self._lines.append(t)
        except Exception:
            pass
        try:
            if self._panel == 2:                              # the status line exists on page 2 only
                self.status.Text = t
        except Exception:
            pass

    # ---------------- body tree (names as in the Mechanical tree) ----------------
    def _build_body_tree(self):
        tv = self.bodytree
        tv.Nodes.Clear()
        self._body_nodes = []
        order, by_part = [], {}
        for b in self._scene["bodies"]:
            pn = b.get("part") or u""
            if pn not in by_part:
                by_part[pn] = []
                order.append(pn)
            by_part[pn].append(b)
        self._tree_updating = True
        try:
            for pn in order:
                bl = by_part[pn]
                if len(bl) == 1:
                    b = bl[0]
                    nm = b.get("label") or b["name"]
                    n = TreeNode(nm if (not pn or pn == nm) else u"%s \\ %s" % (pn, nm))
                    n.Tag = b
                    n.Checked = b["visible"]
                    tv.Nodes.Add(n)
                    self._body_nodes.append((n, b))
                else:
                    pnode = TreeNode(pn or u"(no part)")
                    for b in bl:
                        n = TreeNode(b.get("label") or b["name"])
                        n.Tag = b
                        n.Checked = b["visible"]
                        pnode.Nodes.Add(n)
                        self._body_nodes.append((n, b))
                    pnode.Checked = all(b["visible"] for b in bl)
                    tv.Nodes.Add(pnode)
                    pnode.Expand()
        finally:
            self._tree_updating = False

    def _on_tree_check(self, sender, e):
        if self._tree_updating:
            return
        self._tree_updating = True
        try:
            node = e.Node
            b = node.Tag
            if b is None:                                  # a part: all its bodies
                for ch in node.Nodes:
                    ch.Checked = node.Checked
                    self._set_body(ch.Tag, node.Checked)
            else:
                self._set_body(b, node.Checked)
                if node.Parent is not None:
                    node.Parent.Checked = all(ch.Checked for ch in node.Parent.Nodes)
        finally:
            self._tree_updating = False
        self._update_hide_button()
        self.view.Invalidate()

    def _set_body(self, b, visible):
        if b is None:
            return
        if visible and not b["visible"]:
            self.view.show_body(b)
        elif (not visible) and b["visible"]:
            self.view.hide_body(b)

    def sync_tree(self):
        """Tick marks follow the bodies' visibility (hidden by click, Undo, Show all, Ctrl+Z)."""
        if self._tree_updating:
            return
        self._tree_updating = True
        try:
            for n, b in self._body_nodes:
                n.Checked = b["visible"]
                if n.Parent is not None:
                    n.Parent.Checked = all(ch.Checked for ch in n.Parent.Nodes)
        finally:
            self._tree_updating = False
        self._update_hide_button()

    def _update_hide_button(self):
        """The 'Hide body' button is orange and shows the count while any body is hidden (so a hidden view is obvious)."""
        try:
            n = sum(1 for (nd, b) in self._body_nodes if not b["visible"])
            base = u"Hide: ON" if self.view.hide_mode else u"Hide Body"
            self.btn_hide.Text = base + (u" (%d)" % n if n else u"")
            if n:
                self.btn_hide.BackColor = Color.FromArgb(255, 190, 110)
            else:
                self.btn_hide.UseVisualStyleBackColor = True            # back to the normal look of the other buttons
        except Exception:
            pass

    def _style_tree(self):
        """Bodies that carry the selected contact are shown in blue."""
        lids = self._cur.get("lids") if self._cur is not None else None
        for n, b in self._body_nodes:
            n.ForeColor = Color.FromArgb(0, 70, 170) if (lids is not None and id(b) in lids) else Color.Black

    def _set_tree_collapsed(self, c):
        """Body tree fully hidden to the left (a thin strip with an arrow brings it back)."""
        self._tree_collapsed = bool(c)
        try:
            self.split3.Panel1Collapsed = self._tree_collapsed
        except Exception:
            pass
        self.strip_tree.Visible = self._tree_collapsed
        self._layout_g2()
        self.view.Invalidate()

    def _set_ov_collapsed(self, c):
        """Whole-model view fully hidden to the right (a thin strip with an arrow brings it back)."""
        self._ov_collapsed = bool(c)
        try:
            self.split2.Panel2Collapsed = self._ov_collapsed
        except Exception:
            pass
        self.strip_ov.Visible = self._ov_collapsed
        self._layout()
        self.view.Invalidate()

    def _on_closed(self, sender, e):
        try:
            t = getattr(self, "_front_timer", None)
            if t is not None:
                t.Stop()
        except Exception:
            pass

    def _on_key(self, sender, e):
        """Ctrl+Z = Undo (restores the last hidden body)."""
        if self._panel == 2 and e.Control and e.KeyCode == Keys.Z:
            self.view.undo_hide()
            e.Handled = True
            e.SuppressKeyPress = True

    def _on_isolate_wrong(self, s, e):
        """Keeps the bodies of the WHOLE side (target or contact) that has wrong (orange) normals - not only the bodies
        of the orange elements - and hides all other bodies. Undo / Ctrl+Z / Show all bring them back."""
        cd = self._cur
        if cd is None:
            return
        wrong = set(sg["side"] for sg in cd["segs"] if sg.get("rev") and not sg.get("cp"))
        if not wrong:
            self.note(u"%s: no wrong normals - nothing to isolate" % label_of(cd["c"]))
            return
        keep = []
        for sg in cd["segs"]:
            b = sg.get("body")
            if (sg["side"] in wrong and not sg.get("cp") and b is not None
                    and any(b is x for x in self._scene["bodies"]) and all(b is not k for k in keep)):
                keep.append(b)
        self.view.isolate(keep)
        self.view.fit()

    def _on_local_toggle(self, s, e):
        v = self.view
        v.local_on = not v.local_on
        self.btn_local.Text = u"Local: ON" if v.local_on else u"Local: OFF"
        v.fit()

    def _on_hide_toggle(self, s, e):
        on = not self.view.hide_mode
        self.view.set_hide_mode(on)
        self._update_hide_button()

    def _on_clear_sel(self, s, e):
        cd = self._cur
        if cd is not None:
            cd["sel"].clear()
            self.selection_changed()
            self.view.Invalidate()

    def _on_flip(self, s, e):
        """Writes a NEW flip command ("ContactNormalFlip", "ContactNormalFlip 2", ...) under every contact whose Flip
        action is not 'none'. Commands that already exist are never changed. The view shows the solver state AFTER the
        existing flip commands, so a new command flips relative to that state (an element flipped twice is back to
        normal). 'none' does nothing; 'remove' deletes the NEWEST flip command of the contact. A contact that already
        got a command in this session (violet 'NEW') asks for confirmation first."""
        n = 0
        stamp = System.DateTime.Now.ToString("yyyy-MM-dd HH:mm")
        clock = System.DateTime.Now.ToString("HH:mm")
        for cd in self._cds:
            c = cd["c"]
            olds = find_flip_snippets(c)
            side = cd["side"]
            if side == u"none":
                continue
            if side != u"remove" and cd.get("pending"):
                r_ = MessageBox.Show(
                    u"%s\n\nA flip command was already written for this contact (%s) and is NOT yet reflected in the "
                    u"view. A new flip is calculated relative to the OLD view, so flipping the same elements again "
                    u"would undo them.\n\nWrite another flip command anyway?" % (label_of(c), cd["pending"]),
                    u"Flip Already Pending", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
                if r_ != DialogResult.Yes:
                    continue
            if side == u"remove":
                if olds:
                    nm = olds[-1].Name
                    olds[-1].Delete()
                    self.note(u"Flip command removed (newest: '%s'): %s\r\n" % (nm, label_of(c)))
                    if nm in cd.get("added", []):                  # written in this session: the view never showed it
                        cd["added"].remove(nm)
                    else:                                          # it was part of the view: removing it is a change
                        cd["removed_old"] = cd.get("removed_old", 0) + 1
                    self._update_pending(cd, clock)
                else:
                    self.note(u"%s: there is no flip command to remove\r\n" % label_of(c))
                self._reset_action(cd)
                continue
            if side == u"element":
                chosen = [cd["segs"][i] for i in sorted(cd["sel"]) if i < len(cd["segs"])]
                if not chosen:
                    self.note(u"%s: Flip = element but no element is selected\r\n" % label_of(c))
                    continue
                text, what = None, u""
                info = self._dump_info.get(cd["key"])
                if info is not None:
                    try:
                        text, what = self._positional_flip_text(cd, c, chosen, info[0], stamp)
                    except Exception as ex:
                        self.note(u"%s: command by position failed (%s) - element numbers are used\r\n"
                                  % (label_of(c), str(ex)))
                        text = None
                if text is None:
                    text = build_flip_elems(c.Name, chosen)
                    n_num = sum(1 for x in chosen if x.get("enum"))
                    what = u"%d selected element(s): %d by element number, %d by nodes" % (
                        len(chosen), n_num, len(chosen) - n_num)
            else:                                              # both / contact / target
                text = build_flip_whole(c.Name, side)
                what = u"all %s elements" % side
            names = [x.Name for x in olds]
            nm = SNIPPET_NAME
            k = 1
            while nm in names:
                k += 1
                nm = u"%s %d" % (SNIPPET_NAME, k)
            sn = c.AddCommandSnippet()
            sn.Name = nm
            sn.Input = u"! created %s\n" % stamp + text
            self.note(u"New flip command '%s' written (%s): %s\r\n" % (nm, what, label_of(c)))
            self._reset_action(cd)
            cd.setdefault("added", []).append(nm)
            self._update_pending(cd, clock)
            n += 1
        if n:
            self.note(u"Done. Flip is applied at the next solve; it is not reflected in the view - run "
                      u"Show Normals again to check it (the contact turns green when nothing is left to flip "
                      u"any more). Uses cid / tid of each contact.\r\n")

    def _update_pending(self, cd, clock):
        """Violet only while flip commands are written / removed that the view does not show yet. Commands written in
        this session and removed again cancel out (the row goes back to orange / green); with two or more written
        commands removing one keeps the contact violet."""
        added, gone = cd.get("added", []), cd.get("removed_old", 0)
        if added or gone:
            parts = []
            if added:
                parts.append(added[0] if len(added) == 1 else u"%d commands" % len(added))
            if gone:
                parts.append(u"%d removed" % gone)
            cd["pending"] = u"%s - %s" % (clock, u", ".join(parts))
        else:
            cd["pending"] = None
        self._style_row(cd["idx"])
        if cd is self._cur:
            self.view.warn = self._warn_of(cd)
            self.view.Invalidate()
        for k, r in enumerate(self._rows):
            r["row"].BackColor = self._row_color(k, self._cds[k] is self._cur)

    def _warn_of(self, cd):
        if cd.get("pending"):
            return u"Flip commands changed (%s) - NOT yet reflected in this view: run Show Normals" % cd["pending"]
        return cd.get("warn", u"")

    def _positional_flip_text(self, cd, c, chosen, d, stamp):
        """Flip command by coordinate windows + the mesh record. -> (text, description) or (None, u"")."""
        look = {}
        for sd in ("contact", "target"):
            for it in d.get(sd, []):
                look[(sd, it["e"])] = it
        seg_by = {}
        by_side = {}
        for sg in chosen:
            it = look.get((sg["side"], sg["enum"]))
            if it is None:
                return None, u""
            seg_by[(sg["side"], it["e"])] = sg
            by_side.setdefault(sg["side"], []).append(it)
        hdr = d.get("hdr")
        xy = self._scene["xy"]
        S = cd.get("S", {})
        plans, sigs = [], {}
        n_win = n_ids = 0
        for sd in ("contact", "target"):
            its = by_side.get(sd)
            if not its:
                continue
            reals = set(it["real"] for it in its)
            r = list(reals)[0] if len(reals) == 1 else None
            runs = make_runs(sorted(its, key=lambda x: x["e"]), pool_items(d, sd, r))
            for run in runs:                                  # mesh nodes of the run (for the quick mesh check later)
                ids = set()
                for it in run["items"]:
                    sg = seg_by[(sd, it["e"])]
                    ids.add(sg["a"])
                    ids.add(sg["b"])
                run["nodes"] = [(nid, xy[nid][0], xy[nid][1]) for nid in sorted(ids) if nid in xy]
            plans.append({"side": sd, "real": r, "runs": runs})
            sigs[sd] = mesh_signature(S.get(sd, []))
            n_win += sum(1 for ru in runs if ru["box"] is not None)
            n_ids += sum(1 for ru in runs if ru["box"] is None)
        text = build_flip_positional(c.Name, obj_id(c), plans, hdr, sigs, stamp.replace(" ", "T"), None, d.get("unit"))
        return text, u"%d selected element(s) by position: %d run(s) as coordinate window, %d as element numbers" % (
            len(chosen), n_win, n_ids)

    def _reset_action(self, cd):
        """After the command is written the action goes back to 'none', so pressing Flip Selected again
        cannot flip the same elements a second time."""
        cd["side"] = u"none"
        cd["sel"].clear()
        r = self._rows[cd["idx"]]
        r["cb"].SelectedItem = u"None"
        r["cnt"].Text = u"0"
        self._style_row(cd["idx"])
        self.view.Invalidate()


# ---------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------
MODAL = False        # True: the window blocks Mechanical until it is closed (old behaviour)

model = ExtAPI.DataModel.Project.Model
contact_tree = build_contact_tree(model.Connections)
if sum(len(v) for (_, v) in contact_tree) == 0:
    MessageBox.Show(u"No contacts found in the model.", u"Info",
                    MessageBoxButtons.OK, MessageBoxIcon.Information)
else:
    _open = None
    for _f in Application.OpenForms:                      # run the script again: reuse the open window
        try:
            if _f.Name == FORM_ID:
                _open = _f
                break
        except Exception:
            pass
    if _open is not None:
        bring_to_front(_open)
    else:
        _frm = NormalForm(contact_tree, model)
        if MODAL:
            _owner = mechanical_owner()
            if _owner is not None:
                _frm.ShowDialog(_owner)
            else:
                _frm.ShowDialog()
        else:
            _owner = mechanical_owner()                   # owned by Mechanical's main window: stays above it
            if _owner is not None:                        # (other windows opened later can still cover it)
                _frm.Show(_owner)
            else:
                _frm.Show()                               # Mechanical stays usable
