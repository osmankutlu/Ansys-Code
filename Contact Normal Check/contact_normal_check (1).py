# -*- coding: utf-8 -*-
# =====================================================================
#  Contact Normal Viewer / Flip   (Ansys Mechanical 2D, IronPython 2.7)
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
# =====================================================================
import clr
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
import System
import re

# ==== CORE START ====
import math
import os

CONTACT_ET = (171, 172, 173, 174)
TARGET_ET = (169, 170)
ASSIGN = re.compile(r"^([a-z_]\w*)\s*=\s*([-+]?\d+)\s*(!.*)?$")


def compute_segments(elems, xy, edge_nodes, side):
    """elems: [(eid, corner_ids, all_ids)]; xy: {nid: (x, y)}.
    Direction = element node order; normal = right of the edge direction = (ty, -tx)."""
    segs = []
    for eid, corner, alln in elems:
        n = len(corner)
        if n < 2:
            continue
        pts = [xy[c] for c in corner if c in xy]
        if not pts:
            continue
        ecx = sum(p[0] for p in pts) / float(len(pts))
        ecy = sum(p[1] for p in pts) / float(len(pts))
        for k in range(n):
            a, b = corner[k], corner[(k + 1) % n]
            if a == b or a not in edge_nodes or b not in edge_nodes:
                continue
            pa, pb = xy.get(a), xy.get(b)
            if pa is None or pb is None:
                continue
            tx, ty = pb[0] - pa[0], pb[1] - pa[1]
            L = math.hypot(tx, ty)
            if L == 0:
                continue
            mid = alln[n + k] if len(alln) == 2 * n else None
            segs.append({"side": side, "eid": eid, "a": a, "b": b, "m": mid,
                         "ax": pa[0], "ay": pa[1], "bx": pb[0], "by": pb[1],
                         "mx": (pa[0] + pb[0]) / 2.0, "my": (pa[1] + pb[1]) / 2.0,
                         "ex": ecx, "ey": ecy,
                         "nx": ty / L, "ny": -tx / L, "len": L, "rev": False, "known": False, "enum": None})
    return segs


def iter_dat_lines(path, depth=0):
    base = os.path.dirname(path)
    f = open(path, "r")
    try:
        for raw in f:
            low = raw.strip().lower()
            if depth < 3 and (low.startswith("/input,") or low.startswith("cdread,")):
                p = [x.strip() for x in raw.strip().split(",")]
                try:
                    if low.startswith("/input,"):
                        name, ext, d = p[1], (p[2] if len(p) > 2 else ""), (p[3] if len(p) > 3 else "")
                    else:
                        name, ext, d = p[2], (p[3] if len(p) > 3 else "cdb"), ""
                    fn = name + (("." + ext) if ext else "")
                    cand = os.path.join(d if d else base, fn)
                    if not os.path.isfile(cand):
                        cand = os.path.join(base, fn)
                    if os.path.isfile(cand):
                        for l2 in iter_dat_lines(cand, depth + 1):
                            yield l2
                        continue
                except Exception:
                    pass
            yield raw
    finally:
        f.close()


def _strip0(v):
    v = list(v)
    while v and v[-1] == 0:
        v.pop()
    return v


def learn_node_map(solid_recs, eid2nodes):
    """Solver node id -> mesh node id. Solver and mesh elements with the SAME element number are
    aligned node by node, so it works even when the node numbering of the solver input differs
    (extra nodes, renumbering) - no coordinates are used."""
    votes = {}
    used = 0
    for enum, nodes in solid_recs:
        mn = eid2nodes.get(enum)
        if mn is None or len(mn) != len(nodes):
            continue
        used += 1
        for d, m in zip(nodes, mn):
            v = votes.get(d)
            if v is None:
                votes[d] = {m: 1}
            else:
                v[m] = v.get(m, 0) + 1
    nmap, conflicts = {}, 0
    for d, v in votes.items():
        nmap[d] = max(v.items(), key=lambda kv: kv[1])[0]
        if len(v) > 1:
            conflicts += 1
    diff = sum(1 for d, m in nmap.items() if d != m)
    return nmap, {"elements": used, "nodes": len(nmap), "conflicts": conflicts, "different": diff}


def parse_dat_contacts(it, mesh_elems=None):
    """Reads 2/3-node (contact/target-like) element blocks from solver-input lines.
    COMPACT EBLOCK records are: element number, node numbers.  SOLID records are typed.
    mesh_elems = {element number: [mesh node ids]} is used to translate solver node numbers.
    Returns (blocks, counts, info); block = {'kind','mat','order','pairs': {(lo,hi): (i,j,enum,n,third)}}."""
    params, ets = {}, {}
    cur = {"type": 0, "mat": 0}
    raw_blocks, solid_recs, info = [], [], []
    en_acc, en_meta = {}, {}
    order = [0]

    def num(tok):
        t = tok.split("!")[0].strip().lower()
        if not t:
            return None
        try:
            return int(float(t))
        except Exception:
            return params.get(t)

    def kind_of(typ):
        et = ets.get(typ)
        return "contact" if et in CONTACT_ET else ("target" if et in TARGET_ET else None)

    def new_block(kind, mat):
        order[0] += 1
        b = {"kind": kind, "mat": mat, "order": order[0], "recs": []}
        raw_blocks.append(b)
        return b

    for raw in it:
        s = raw.strip()
        if not s or s[0] == "!":
            continue
        low = s.lower()
        if low.startswith("/com"):
            continue
        m = ASSIGN.match(low)
        if m:
            params[m.group(1)] = int(m.group(2))
            continue
        if low.startswith("*set,"):
            p = [x.strip() for x in low.split(",")]
            if len(p) >= 3 and num(p[2]) is not None:
                params[p[1]] = num(p[2])
        elif low.startswith("et,"):
            p = low.split(",")
            if len(p) >= 3:
                t, e = num(p[1]), num(p[2])
                if t is not None and e is not None:
                    ets[t] = e
        elif low.startswith("type,") or low.startswith("mat,"):
            v = num(low.split(",")[1])
            if v is not None:
                cur["type" if low.startswith("type,") else "mat"] = v
        elif low.startswith("en,"):
            p = [x.strip() for x in low.split(",")]
            try:
                enum = int(p[1])
                nds = [int(x) for x in p[2:] if x]
            except Exception:
                continue
            if enum not in en_acc:
                en_acc[enum] = []
                en_meta[enum] = (cur["type"], cur["mat"])
            en_acc[enum].extend(nds)
        elif low.startswith("eblock"):
            hdr = [x.strip() for x in low.split(",")]
            try:
                nf = int(hdr[1])
            except Exception:
                nf = 19
            is_solid = ("solid" in low) or (nf >= 19 and "compact" not in low)
            next(it)                                   # format line
            crecs, by_kind = [], {}
            for l in it:
                t = l.split()
                if not t or t[0] == "-1":
                    break
                v = [int(x) for x in t]
                if is_solid:
                    if len(v) < 11:
                        continue
                    typ, numn, enum = v[1], v[8], v[10]
                    nodes = v[11:]
                    while len(nodes) < numn:
                        nodes.extend([int(x) for x in next(it).split()])
                    nodes = nodes[:numn]
                    k = kind_of(typ)
                    if k and 2 <= len(nodes) <= 3:
                        b = by_kind.get(k) or new_block(k, cur["mat"])
                        by_kind[k] = b
                        b["recs"].append((enum, nodes))
                    elif len(nodes) >= 3:
                        solid_recs.append((enum, nodes))
                else:
                    crecs.append(v)
            if crecs:
                parsed = [(r[0], _strip0(r[1:])) for r in crecs]
                short = [(e, n) for e, n in parsed if 2 <= len(n) <= 3]
                if len(short) >= 0.8 * len(parsed):
                    new_block(None, cur["mat"])["recs"] = short
                else:
                    solid_recs.extend((e, n) for e, n in parsed if len(n) >= 3)
    for enum, nds in en_acc.items():
        typ, mat = en_meta[enum]
        k = kind_of(typ)
        if k and 2 <= len(nds) <= 3:
            b = None
            for x in raw_blocks:
                if x["kind"] == k and x["mat"] == mat and x.get("en"):
                    b = x
            if b is None:
                b = new_block(k, mat)
                b["en"] = True
            b["recs"].append((enum, nds))
        elif len(nds) >= 3:
            solid_recs.append((enum, nds))

    nmap = None
    if mesh_elems and solid_recs:
        nmap, st = learn_node_map(solid_recs, mesh_elems)
        if st["elements"] == 0:
            info.append("no solver element has the same element number as a mesh element - the "
                        "file does not belong to this mesh")
            return [], {"records": 0, "blocks": 0}, info
        elif st["different"] == 0:
            info.append("solver node numbers = mesh node numbers (checked with %d elements)" % st["elements"])
            nmap = None
        else:
            info.append("solver node numbers differ from the mesh numbers: translated through %d elements "
                        "with the same element number (%d nodes, %d different, %d conflicting)"
                        % (st["elements"], st["nodes"], st["different"], st["conflicts"]))
    elif mesh_elems:
        info.append("no solid elements found in the file to check the numbering")

    blocks, total = [], 0
    for rb in raw_blocks:
        pairs = {}
        for enum, nodes in rb["recs"]:
            if nmap is not None:
                if any(n not in nmap for n in nodes):
                    continue
                nodes = [nmap[n] for n in nodes]
            i, j = nodes[0], nodes[1]
            pairs[(min(i, j), max(i, j))] = (i, j, enum, len(nodes), nodes[2] if len(nodes) == 3 else None)
        if pairs:
            blocks.append({"kind": rb["kind"], "mat": rb["mat"], "order": rb["order"], "pairs": pairs})
            total += len(pairs)
    return blocks, {"records": total, "blocks": len(blocks)}, info


def assign_blocks(side_keys, blocks, used):
    """Chooses, for ONE contact, the solver-input block of its contact elements and the block of
    its target elements - independently, by how many of the mesh edge node pairs each block covers.
    side_keys = {'contact': set(pair keys), 'target': set(pair keys)}; used = set of block orders
    already given to other contacts (duplicates of the same edges get the next block).
    Returns ({'contact': pairs|None, 'target': pairs|None}, note)."""
    res = {"contact": None, "target": None}
    picked = {}
    taken = set()
    for side in ("target", "contact"):            # target first: it comes first in the input
        keys = side_keys[side]
        if not keys:
            continue
        cand = []
        for blk in blocks:
            if blk["kind"] and blk["kind"] != side:
                continue
            if blk["order"] in taken:
                continue
            sc = len(keys & set(blk["pairs"].keys()))
            if sc < 1:
                continue
            pen = 0.5 if blk["order"] in used else 0.0
            cand.append((-(sc - pen), blk["order"], blk))
        if cand:
            cand.sort(key=lambda c: (c[0], c[1]))
            blk = cand[0][2]
            taken.add(blk["order"])
            used.add(blk["order"])
            res[side] = blk["pairs"]
            picked[side] = blk
    note = ""
    if len(picked) == 2 and set(picked["contact"]["pairs"]) == set(picked["target"]["pairs"]):
        note = "contact and target elements share nodes; assumed the target block comes first"
    elif len(picked) < 2:
        missing = [k for k in ("contact", "target") if k not in picked and side_keys[k]]
        if missing:
            note = "no solver-input block found for: %s" % ", ".join(missing)
    return res, note


def loose_pairs(blocks, keys):
    """Solver records for the given node pairs taken from ANY block (first block in file order)."""
    out = {}
    for blk in sorted(blocks, key=lambda b: b["order"]):
        for k, rec in blk["pairs"].items():
            if k in keys and k not in out:
                out[k] = rec
    return out


def apply_dat(segs, sides, xy):
    """Takes direction and element number from the solver-input element with the same node pair.
    Duplicate segments (the same edge seen from both neighbouring mesh elements) are merged and
    segments without a solver element are dropped when that side has solver data.
    Returns (matched, dropped)."""
    matched = dropped = 0
    keep, seen = [], set()
    side_has = dict((k, bool(sides.get(k))) for k in ("contact", "target"))
    for sg in segs:
        key = (min(sg["a"], sg["b"]), max(sg["a"], sg["b"]))
        recs = sides.get(sg["side"]) or {}
        rec = recs.get(key)
        ok = rec is not None and rec[0] in xy and rec[1] in xy
        if ok and len(rec) > 4 and rec[4] is not None and rec[4] in xy:
            p, q, t3 = xy[rec[0]], xy[rec[1]], xy[rec[4]]
            mid = ((p[0] + q[0]) / 2.0, (p[1] + q[1]) / 2.0)
            if math.hypot(t3[0] - mid[0], t3[1] - mid[1]) > 0.25 * math.hypot(q[0] - p[0], q[1] - p[1]):
                ok = False
        if ok:
            if (sg["side"], key) in seen:
                dropped += 1
                continue
            seen.add((sg["side"], key))
            p, q = xy[rec[0]], xy[rec[1]]
            tx, ty = q[0] - p[0], q[1] - p[1]
            L = math.hypot(tx, ty)
            if L > 0:
                sg["nx"], sg["ny"] = ty / L, -tx / L
                sg["known"] = True
                sg["enum"] = rec[2]
                matched += 1
            keep.append(sg)
        elif side_has[sg["side"]]:
            dropped += 1
        else:
            keep.append(sg)
    segs[:] = keep
    return matched, dropped


def flag_minority(segs):
    """With solver data every element is a real contact/target element, so 'reversed' means:
    opposite to the majority direction on its side."""
    n = 0
    for side in ("contact", "target"):
        grp = [x for x in segs if x["side"] == side]
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


def _on_nodes(rec, nodes):
    """Solver record (i, j, enum, nnodes, third) lies on the given edge nodes."""
    if rec[0] not in nodes or rec[1] not in nodes:
        return False
    return len(rec) < 5 or rec[4] is None or rec[4] in nodes


def collect_elements(blocks, S, used):
    """Contact/target line elements of the solver input that lie on one contact's edges, found
    directly from the node numbers (no matching with mesh segments).
    S = {'contact': set(nodes), 'target': set(nodes)}; used = block orders already given to other
    contacts (identical duplicate contacts get the next block when there is one).
    Returns ({'contact': [rec], 'target': [rec]}, note)."""
    cand = {"contact": [], "target": [], "both": []}
    for blk in sorted(blocks, key=lambda b: b["order"]):
        rc = [r for r in blk["pairs"].values() if _on_nodes(r, S["contact"])]
        rt = [r for r in blk["pairs"].values() if _on_nodes(r, S["target"])]
        if rc and rt:
            cand["both"].append((blk, rc, rt))
        elif rc:
            cand["contact"].append((blk, rc))
        elif rt:
            cand["target"].append((blk, rt))
    out = {"contact": {}, "target": {}}
    info = []
    for side in ("contact", "target"):
        lst = cand[side]
        use = [x for x in lst if x[0]["order"] not in used] or lst
        for blk, recs in use:
            used.add(blk["order"])
            for r in recs:
                out[side].setdefault((min(r[0], r[1]), max(r[0], r[1])), r)
        if use:
            info.append("%s: %d solver elements (block %s)"
                        % (side, len(out[side]), ",".join("#%d" % x[0]["order"] for x in use)))
    both = [x for x in cand["both"] if x[0]["order"] not in used] or cand["both"]
    if both:
        if not out["target"]:
            blk, rc, rt = both[0]
            used.add(blk["order"])
            for r in rt:
                out["target"].setdefault((min(r[0], r[1]), max(r[0], r[1])), r)
            both = both[1:]
            info.append("target: %d solver elements (block #%d, shares nodes with the contact side; "
                        "assumed the target block comes first)" % (len(out["target"]), blk["order"]))
        if not out["contact"] and both:
            blk, rc, rt = both[0]
            used.add(blk["order"])
            for r in rc:
                out["contact"].setdefault((min(r[0], r[1]), max(r[0], r[1])), r)
            info.append("contact: %d solver elements (block #%d, shares nodes)"
                        % (len(out["contact"]), blk["order"]))
    for side in ("contact", "target"):
        if not out[side] and S[side]:
            info.append("%s: NO solver elements found on its edges" % side)
    return dict((k, list(v.values())) for k, v in out.items()), "; ".join(info)


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
        dx, dy = sg["ny"], -sg["nx"]
        if abs(dx) >= abs(dy):
            k = "right" if dx > 0 else "left"
        else:
            k = "down" if dy > 0 else "up"
        d = res[sg["side"]]
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
            if edge_map:
                cands = edge_map.get((min(i, j), max(i, j)))
                if cands:
                    same = [b for (b, a, c) in cands if a == i and c == j]
                    common = [same[0] if same else cands[0][0]]   # the element this solver element sits on
            third = rec[4] if len(rec) > 4 else None
            segs.append({"side": side, "eid": enum, "enum": enum, "a": i, "b": j, "m": third,
                         "ax": pa[0], "ay": pa[1], "bx": pb[0], "by": pb[1],
                         "mx": (pa[0] + pb[0]) / 2.0, "my": (pa[1] + pb[1]) / 2.0,
                         "ex": (pa[0] + pb[0]) / 2.0, "ey": (pa[1] + pb[1]) / 2.0,
                         "nx": ty / L, "ny": -tx / L, "len": L, "rev": False, "known": True,
                         "body": common[0], "bodies": common})
    return segs


def seg_visible(sg):
    b = sg.get("bodies")
    return any(x["visible"] for x in b) if b else sg["body"]["visible"]


def dat_id_stats(path):
    """Node / element id ranges of a solver input (NBLOCK and COMPACT/SOLID EBLOCK)."""
    nodes = []
    elems = []
    if True:
        it = iter_dat_lines(path)
        for raw in it:
            t = raw.strip().lower()
            if t.startswith("nblock"):
                next(it)
                for l in it:
                    w = l.split()
                    if not w or w[0] == "-1":
                        break
                    nodes.append(int(w[0]))
            elif t.startswith("eblock"):
                solid = ("solid" in t) or "compact" not in t
                next(it)
                for l in it:
                    w = l.split()
                    if not w or w[0] == "-1":
                        break
                    v = [int(x) for x in w]
                    if solid:
                        if len(v) >= 11:
                            elems.append((v[10], len(v) - 11))
                    else:
                        nn = len(_strip0(v[1:]))
                        elems.append((v[0], nn))
    area = [e for e, n in elems if n >= 3]
    return {"n_nodes": len(nodes), "n_min": min(nodes) if nodes else 0, "n_max": max(nodes) if nodes else 0,
            "n_area": len(area), "e_min": min(area) if area else 0, "e_max": max(area) if area else 0,
            "n_line": len(elems) - len(area)}


def check_dat(path):
    """Quick look at a file: is it a solver input (APDL .dat with NBLOCK / EBLOCK)?"""
    n_el = n_nb = n_et = n_cp = n_lines = 0
    try:
        for raw in iter_dat_lines(path):
            n_lines += 1
            t = raw.strip().lower()
            if t.startswith("eblock"):
                n_el += 1
                if "compact" in t:
                    n_cp += 1
            elif t.startswith("nblock"):
                n_nb += 1
            elif t.startswith("et,"):
                n_et += 1
    except Exception as ex:
        return False, "cannot read the file: %s" % str(ex)
    if n_el == 0 or n_nb == 0:
        return False, ("does not look like a solver input file: %d EBLOCK, %d NBLOCK, %d ET lines "
                       "(%d lines)" % (n_el, n_nb, n_et, n_lines))
    return True, ("looks like a solver input: %d element blocks (%d compact), %d node blocks, "
                  "%d ET lines, %d lines" % (n_el, n_cp, n_nb, n_et, n_lines))


def flag_inward(segs):
    """rev = normal points into its own underlying element."""
    n = 0
    for s in segs:
        d = s["nx"] * (s["mx"] - s["ex"]) + s["ny"] * (s["my"] - s["ey"])
        s["rev"] = d < 0
        if s["rev"]:
            n += 1
    return n


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
# ==== CORE END ====

import os
from System.Windows.Forms import (Form, Button, Label, MessageBox, MessageBoxButtons,
                                  MessageBoxIcon, FormStartPosition, FormBorderStyle,
                                  DialogResult, GroupBox, FlatStyle, TextBox, TreeView,
                                  TreeNode, Application, ComboBox, ComboBoxStyle,
                                  ListView, ListViewItem, View, ScrollBars, Panel,
                                  Cursor, Cursors, Control, Keys, MouseButtons, Clipboard,
                                  Timer, IWin32Window, BorderStyle, AnchorStyles, Screen,
                                  SplitContainer, Orientation, DockStyle, FixedPanel, OpenFileDialog)
from System.Drawing import Size, Point, Font, FontStyle, Color, PointF, Pen, SolidBrush, Bitmap, Graphics
from System.Drawing.Drawing2D import Matrix, SmoothingMode, DashStyle

import Microsoft.Win32 as Win32

REG_KEY = r"Software\ContactNormalCheck"


def reg_get(name, default=u""):
    try:
        k = Win32.Registry.CurrentUser.OpenSubKey(REG_KEY)
        v = k.GetValue(name, default) if k else default
        return v if v else default
    except Exception:
        return default


def reg_set(name, val):
    try:
        k = Win32.Registry.CurrentUser.CreateSubKey(REG_KEY)
        k.SetValue(name, val)
    except Exception:
        pass


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
def load_scene(log):
    mesh = ExtAPI.DataModel.MeshDataByName("Global")
    owner = {}
    bodies = []
    body_part = []
    try:
        pidx = -1
        for asm in ExtAPI.DataModel.GeoData.Assemblies:
            for part in asm.Parts:
                pidx += 1
                for body in part.Bodies:
                    name = u"%s\\%s [%d]" % (part.Name, body.Name, body.Id)
                    bodies.append(name)
                    body_part.append(pidx)
                    for face in body.Faces:
                        for eid in list(mesh.MeshRegionById(face.Id).ElementIds):
                            owner[eid] = len(bodies) - 1
    except Exception as ex:
        log(u"body/face lookup failed (%s)" % str(ex))
    names = list(bodies) + [u"Other elements"]
    body_part.append(max(body_part) + 1 if body_part else 0)
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
                "color": Color.FromArgb(*PALETTE[body_part[bi] % len(PALETTE)])}
        scene["bodies"].append(body)
        for e in elems:
            scene["eid2body"][e[0]] = body
    xs = [p[0] for p in xy.values()]
    ys = [p[1] for p in xy.values()]
    scene["box"] = (min(ys), max(ys), -max(xs), -min(xs))      # draw coords (X=y, Y=-x)
    log(u"model view: %d bodies, %d area elements drawn, %d line/degenerate elements skipped"
        % (len(scene["bodies"]), drawn, skipped))
    return scene


def dedupe_predicted(segs, dummy):
    """Prediction only: keep ONE arrow per node pair (lowest element number) and make it independent
    of body visibility, so hiding bodies never changes what is drawn."""
    best = {}
    for sg in segs:
        key = (sg["side"], min(sg["a"], sg["b"]), max(sg["a"], sg["b"]))
        if key not in best or sg["eid"] < best[key]["eid"]:
            best[key] = sg
    out = []
    for sg in segs:
        key = (sg["side"], min(sg["a"], sg["b"]), max(sg["a"], sg["b"]))
        if best[key] is sg:
            sg["body"] = dummy
            sg["bodies"] = [dummy]
            out.append(sg)
    return out


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


def contact_segments(scene, c, log):
    mesh = scene["mesh"]
    out = []
    for side, loc in (("contact", c.SourceLocation), ("target", c.TargetLocation)):
        nodes = set()
        for gid in list(loc.Ids):
            try:
                nodes.update(list(mesh.MeshRegionById(gid).NodeIds))
            except Exception as ex:
                log(u"%s: edge %s mesh lookup failed (%s)" % (label_of(c), gid, str(ex)))
        for b in scene["bodies"]:
            segs = compute_segments([(e[0], e[1], e[2]) for e in b["elems"]],
                                    scene["xy"], nodes, side)
            for sg in segs:
                sg["body"] = b
                sg["bodies"] = [b]
            out.extend(segs)
    return out


# ---------------------------------------------------------------
# Solver input (ds.dat): real contact/target element orientation
# ---------------------------------------------------------------
def dat_summary(path):
    rows = []
    counts = {}
    f = open(path, "r")
    try:
        for raw in f:
            rows.append(raw.rstrip())
            t = raw.strip().lower()
            if t and t[0] not in "!-0123456789 ":
                k = re.split(r"[,\s=]", t)[0]
                counts[k] = counts.get(k, 0) + 1
    finally:
        f.close()
    out = [u"file: %s  (%d lines)" % (path, len(rows)), u""]
    out.append(u"=== ET / parameter lines ===")
    for r in rows:
        t = r.strip().lower()
        if t.startswith("et,") or t.startswith("*set,") or t.startswith("keyo,") and False:
            out.append(u"  " + r[:100])
    out.append(u"")
    out.append(u"=== first EBLOCKs (header, format, first 3 records) ===")
    shown = 0
    for n, r in enumerate(rows):
        if r.strip().lower().startswith("eblock"):
            out.append(u"--- line %d ---" % (n + 1))
            for x in rows[n:n + 5]:
                out.append(u"  " + x[:120])
            shown += 1
            if shown >= 4:
                break
    out.append(u"")
    out.append(u"=== lines mentioning esurf / gcgen / cmblock / /input / cdread ===")
    k = 0
    for n, r in enumerate(rows):
        t = r.strip().lower()
        if t.startswith(("esurf", "gcgen", "cmblock", "/input", "cdread", "nread", "eread")):
            out.append(u"  %d: %s" % (n + 1, r[:100]))
            k += 1
            if k >= 15:
                break
    out.append(u"")
    out.append(u"=== command counts (top 30) ===")
    for key, v in sorted(counts.items(), key=lambda kv: -kv[1])[:30]:
        out.append(u"  %s : %d" % (key, v))
    return u"\r\n".join(out)


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
        return [(i, g) for i, g in enumerate(self.segs) if seg_visible(g)]

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
                self._err = u"draw error: %s" % str(ex)
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
        glow1 = Pen(Color.FromArgb(0, 210, 120), 12.0)
        glow2 = Pen(Color.FromArgb(0, 210, 120), 9.0)
        pens, shafts, fills = {}, {}, {}
        for k in base:
            for rv in (False, True):
                c = orange if rv else base[k]
                pens[(k, rv)] = Pen(c, 4.0)
                shafts[(k, rv)] = Pen(c, 2.4)
                fills[(k, rv)] = SolidBrush(c)
        L = self._arrow_len()
        off = 3.0
        hl = min(L * 0.45, 13.0)
        r = 3.0
        s = self.s
        ox, oy = self._ofs()
        for i, sg in self._visible():
            key = (sg["side"], sg["rev"])
            dx, dy = sg["ny"], -sg["nx"]                       # unit normal in draw/screen coords
            A = PointF(sg["ay"] * s + ox + dx * off, -sg["ax"] * s + oy + dy * off)
            B = PointF(sg["by"] * s + ox + dx * off, -sg["bx"] * s + oy + dy * off)
            M = PointF((A.X + B.X) / 2.0, (A.Y + B.Y) / 2.0)
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
        g.DrawString(u"contact %d / target %d segments | reversed %d | selected %d    %s"
                     % (nc, len(vis) - nc, nr, len(self.sel), self.note), f, dim, 6.0, 4.0)
        y = self.Height - 20.0
        x = 8.0
        for col, txt, w in ((Color.FromArgb(220, 20, 40), u"contact normal", 100.0),
                            (Color.FromArgb(30, 70, 220), u"target normal", 98.0),
                            (Color.FromArgb(255, 120, 0), u"opposite to the side's majority", 190.0),
                            (Color.FromArgb(0, 210, 120), u"selected", 60.0)):
            g.FillRectangle(SolidBrush(col), x, y + 3, 14.0, 8.0)
            g.DrawString(txt, f, dim, x + 18.0, y)
            x += w + 22.0
        if self.warn:
            g.DrawString(self.warn, f, SolidBrush(Color.Firebrick), 6.0, 22.0)
        if getattr(self, "_err", u""):
            g.DrawString(self._err, f, SolidBrush(Color.Firebrick), 6.0, 22.0)
        if self.hide_mode:
            g.DrawString(u"HIDE BODY: click a body", f, SolidBrush(Color.DarkOrange), self.Width - 170.0, 4.0)
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
                    b["visible"] = False
                    self.hidden.append(b)
                    self.owner.note(u"Hidden: %s" % b["name"])
                    return

    def undo_hide(self):
        if self.hidden:
            b = self.hidden.pop()
            b["visible"] = True
            self.owner.note(u"Restored: %s" % b["name"])
            self.Invalidate()

    def show_all(self):
        while self.hidden:
            self.hidden.pop()["visible"] = True
        self.Invalidate()

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
                if not b["visible"]:
                    continue
                br = SolidBrush(b["color"])
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
            key = (self.Width, self.Height, id(self.scene),
                   tuple(b["visible"] for b in self.scene["bodies"]))
            if key != self._key or self._cache is None:
                if self._cache is not None:
                    self._cache.Dispose()
                self._cache = self._render(sc, ox, oy)
                self._key = key
            g.DrawImage(self._cache, 0, 0)
            g.SmoothingMode = SmoothingMode.AntiAlias
            vis = [x for x in self.segs if seg_visible(x)]
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
            g.DrawString(u"overview error: %s" % str(ex), Font("Segoe UI", 8.5),
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
class NormalForm(Form):
    def __init__(self, contact_tree, model):
        self.Text = u"Contact Normal Check"
        wa = Screen.PrimaryScreen.WorkingArea
        self.ClientSize = Size(min(1700, wa.Width - 30), min(900, wa.Height - 70))
        self.MinimumSize = Size(1000, 620)
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.Sizable
        self.MaximizeBox = True
        self.MinimizeBox = True
        self._panel = 0
        self.Resize += lambda s, e: self._layout()
        self.Shown += lambda s, e: (bring_to_front(self), schedule_front(self))
        self._tree = contact_tree
        self._model = model
        self._sel_keys = set()
        self._cds = []
        self._scene = None
        self._diag = None
        self._dat_path = u""
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
        self.Text = u"Contact Normal Check  -  Select Contacts"
        W, H = self.ClientSize.Width, self.ClientSize.Height
        ALL = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
        grp = GroupBox()
        grp.Text = u"Contacts  (check the ones whose normals you want to see)"
        grp.Font = self.font_title
        grp.SetBounds(14, 10, W - 28, H - 144)
        grp.Anchor = ALL
        self.Controls.Add(grp)
        self.tv = TreeView()
        self.tv.SetBounds(10, 24, W - 48, H - 144 - 34)
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
                ch = TreeNode(label_of(c))
                ch.Tag = u"CONTACT"
                ch.NodeFont = self.base_font
                ch.Checked = (ukey_of(c, len(self._ordered)) in self._sel_keys)
                self._ordered.append((ch, c))
                gn.Nodes.Add(ch)
            self.tv.Nodes.Add(gn)
            gn.Expand()
        self._updating = False

        LB = AnchorStyles.Bottom | AnchorStyles.Left
        RB = AnchorStyles.Bottom | AnchorStyles.Right
        LRB = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        yd = H - 128
        lbl = Label()
        lbl.Text = u"Solver input (ds.dat):"
        lbl.Font = self.font_bold
        lbl.AutoSize = True
        lbl.Location = Point(16, yd + 5)
        lbl.Anchor = LB
        self.Controls.Add(lbl)
        self.txt_dat = TextBox()
        self.txt_dat.Font = self.base_font
        self.txt_dat.SetBounds(180, yd + 2, W - 14 - 180 - 100 - 8, 24)
        self.txt_dat.Anchor = LRB
        self.txt_dat.Text = self._dat_path or reg_get(u"LastDat", u"")
        self.Controls.Add(self.txt_dat)
        bb = self._btn(u"Browse...", W - 14 - 100, yd, 100, 28, self._on_browse)
        bb.Anchor = RB
        self.Controls.Add(bb)
        self.lbl_dat = Label()
        self.lbl_dat.Font = self.base_font
        self.lbl_dat.AutoSize = True
        self.lbl_dat.Location = Point(180, yd + 31)
        self.lbl_dat.Anchor = LB
        self.Controls.Add(self.lbl_dat)
        hint = Label()
        hint.Text = (u"Use the complete ds.dat of this analysis (solver files directory, or Solution > Tools > "
                     u"Write Input Files). It must contain NBLOCK / EBLOCK and match the current mesh.")
        hint.Font = self.base_font
        hint.ForeColor = Color.DimGray
        hint.AutoSize = True
        hint.Location = Point(180, yd + 51)
        hint.Anchor = LB
        self.Controls.Add(hint)
        self._dat_status()

        yb = H - 40
        b1 = self._btn(u"Select All", 14, yb, 100, 30, lambda s, e: self._set_all(True))
        b2 = self._btn(u"Clear", 124, yb, 80, 30, lambda s, e: self._set_all(False))
        bc = self._btn(u"Cancel", W - 14 - 124 - 10 - 100, yb, 100, 30, None)
        b3 = self._btn(u"Show Normals >", W - 14 - 124, yb, 124, 30, self._on_show, True)
        for b, an in ((b1, LB), (b2, LB), (bc, RB), (b3, RB)):
            b.Anchor = an
            self.Controls.Add(b)
        bc.DialogResult = DialogResult.Cancel
        self.CancelButton = bc

    def _dat_status(self):
        p = self.txt_dat.Text.strip()
        if not p:
            self.lbl_dat.ForeColor = Color.Firebrick
            self.lbl_dat.Text = u"No file selected."
            return
        if not System.IO.File.Exists(p):
            self.lbl_dat.ForeColor = Color.Firebrick
            self.lbl_dat.Text = u"File not found."
            return
        ok, msg = check_dat(p)
        self.lbl_dat.ForeColor = Color.ForestGreen if ok else Color.Firebrick
        self.lbl_dat.Text = msg

    def _on_browse(self, s, e):
        dlg = OpenFileDialog()
        dlg.Title = u"Select the solver input file (ds.dat)"
        dlg.Filter = u"Solver input (*.dat)|*.dat|All files (*.*)|*.*"
        try:
            cur = self.txt_dat.Text.strip()
            if cur and System.IO.Directory.Exists(System.IO.Path.GetDirectoryName(cur)):
                dlg.InitialDirectory = System.IO.Path.GetDirectoryName(cur)
        except Exception:
            pass
        if dlg.ShowDialog() == DialogResult.OK:
            self.txt_dat.Text = dlg.FileName
            reg_set(u"LastDat", dlg.FileName)
            self._dat_status()

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

    def _on_show(self, s, e):
        sel = []
        for i, (n, c) in enumerate(self._ordered):
            if n.Checked:
                sel.append((ukey_of(c, i), c))
        if not sel:
            MessageBox.Show(u"Please select at least one contact.", u"Warning",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        self._dat_path = self.txt_dat.Text.strip()
        if not self._dat_path or not System.IO.File.Exists(self._dat_path):
            MessageBox.Show(u"Please select the solver input file (ds.dat) first.", u"Warning",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        reg_set(u"LastDat", self._dat_path)
        self._sel_keys = set(k for k, c in sel)
        lines = []
        self._cds = []
        self._diag = None
        Cursor.Current = Cursors.WaitCursor
        try:
            if self._scene is None:
                self._scene = load_scene(lines.append)
            blocks, dpath, reason = None, None, u""
            try:
                dpath = self._dat_path
                lines.append(u"solver input file: %s" % dpath)
                if dpath:
                    mesh_elems = {}
                    for b in self._scene["bodies"]:
                        for e in b["elems"]:
                            mesh_elems[e[0]] = list(e[2])
                    blocks, cnt, pinfo = parse_dat_contacts(iter_dat_lines(dpath), mesh_elems)
                    for t in pinfo:
                        lines.append(t)
                    lines.append(u"solver input: %d contact/target-like element records in %d blocks"
                                 % (cnt["records"], cnt["blocks"]))
                    if cnt["records"] == 0:
                        blocks = None
                        reason = u"; ".join(pinfo) or (u"no 2/3-node contact/target elements were found "
                                                       u"in the file")
                        smp = []
                        for b in self._scene["bodies"][:2]:
                            for e in b["elems"][:3]:
                                smp.append(u"  element %s nodes %s" % (e[0], list(e[2])))
                        self._diag = (dat_summary(dpath) + u"\r\n\r\n=== mesh element samples (from the "
                                      u"Mechanical API) ===\r\n" + u"\r\n".join(smp))
                        lines.append(u"no contact elements usable in the solver input - see the "
                                     u"diagnostics window (Copy, then send it)")
                else:
                    lines.append(u"solver input not available")
            except Exception as ex:
                reason = u"reading failed: %s" % str(ex)
                lines.append(u"solver input read failed: %s" % str(ex))
            if not blocks and dpath:
                try:
                    st = dat_id_stats(dpath)
                    xy_ = self._scene["xy"]
                    ids_e = [e_[0] for b_ in self._scene["bodies"] for e_ in b_["elems"]]
                    reason += (u"\n\nSolver input: %d nodes (ids %d..%d), %d area elements (ids %d..%d)"
                               u"\nMechanical mesh: %d nodes (ids %d..%d), %d area elements (ids %d..%d)"
                               % (st["n_nodes"], st["n_min"], st["n_max"], st["n_area"], st["e_min"], st["e_max"],
                                  len(xy_), min(xy_), max(xy_), len(ids_e), min(ids_e), max(ids_e)))
                    if abs(st["n_area"] - len(ids_e)) <= 0.02 * max(1, len(ids_e)):
                        reason += (u"\n\nSame number of elements but different ids: the solver numbering "
                                   u"differs from the mesh numbering (renumbering, e.g. Mesh Numbering).")
                    else:
                        reason += (u"\n\nThe element counts differ: the file was not written from the "
                                   u"current mesh (or belongs to another model). Select the complete ds.dat of this "
                                   u"analysis.")
                except Exception as ex:
                    reason += u"\n(could not compare with the mesh: %s)" % str(ex)
                lines.append(reason)
            if not blocks:
                Cursor.Current = Cursors.Default
                ans = MessageBox.Show(
                    u"The solver input could not be used:\n\n%s\n\nWithout it the arrows are only "
                    u"PREDICTED from the mesh and can differ from Ansys.\n\nContinue with the "
                    u"prediction anyway?" % (reason or u"unknown reason"),
                    u"Solver input", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
                if ans != DialogResult.Yes:
                    if self._diag:
                        TextForm(u"Solver input diagnostics", self._diag).ShowDialog(self)
                    return
                Cursor.Current = Cursors.WaitCursor
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
                                  "idx": i, "note": u"", "warn": u"", "lids": None, "lbox": None})
            used = set()
            for cd in sorted(self._cds, key=lambda x: x["key"]):
                c = cd["c"]
                S = contact_node_sets(self._scene, c, lines.append)
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
                if blocks:
                    recs, note = collect_elements(blocks, S, used)
                    emap = edge_owners(inv_b, S["contact"] | S["target"])
                    segs = dat_segments(recs, self._scene["xy"], node2b, dummy, emap)
                    flag_minority(segs)
                    cd["note"] = u"| solver input: %d contact/target elements" % len(segs)
                    if not segs:
                        cd["warn"] = u"NO solver-input elements found on this contact's edges"
                    lines.append(u"%s: %s" % (label_of(c), note))
                else:
                    try:
                        segs = contact_segments(self._scene, c, lines.append)
                    except Exception as ex:
                        segs = []
                        lines.append(u"%s: segment extraction failed (%s)" % (label_of(c), str(ex)))
                    segs = dedupe_predicted(segs, dummy)
                    flag_inward(segs)
                    cd["warn"] = (u"PREDICTED from the mesh (solver input not available) - "
                                  u"may differ from Ansys")
                cd["segs"] = segs
                dc = direction_counts(segs)
                dtxt = u"; ".join(
                    u"%s: %s" % (sd, u" ".join(u"%s %d" % (k_, v_) for k_, v_ in sorted(dc[sd].items())))
                    for sd in ("contact", "target") if dc[sd])
                if dtxt:
                    cd["note"] += u" | directions (screen) " + dtxt
                nc = sum(1 for x in segs if x["side"] == "contact")
                lines.append(u"%s: %d contact / %d target elements, %d in the minority direction; %s"
                             % (label_of(c), nc, len(segs) - nc, sum(1 for x in segs if x["rev"]), dtxt))
            try:
                bx = self._scene["box"]
                lines.append(u"geometry check: scene box X[%.5g, %.5g] Y[%.5g, %.5g] (draw coords), "
                             u"visible bodies %d" % (bx[0], bx[1], bx[2], bx[3],
                                                     sum(1 for b in self._scene["bodies"] if b["visible"])))
                if self._cds and self._cds[0]["segs"]:
                    g0 = self._cds[0]["segs"][0]
                    p0 = self._scene["bodies"][0]["elems"][0][4][0]
                    lines.append(u"geometry check: first segment at (%.5g, %.5g), first element point "
                                 u"at (%.5g, %.5g)" % (g0["ay"], -g0["ax"], p0[0], p0[1]))
            except Exception as ex:
                lines.append(u"geometry check failed: %s" % str(ex))
        finally:
            Cursor.Current = Cursors.Default
        self._build_result_panel(lines)
        bring_to_front(self)
        schedule_front(self)
        if self._diag:
            TextForm(u"Solver input diagnostics", self._diag).ShowDialog(self)

    # ---------------- panel 2 ----------------
    def _build_result_panel(self, lines):
        self.Controls.Clear()
        self._panel = 0                                   # layout enabled at the end
        self._split_ready = False
        self.Text = u"Contact Normal Check  -  Normals"

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

        g1 = GroupBox()
        g1.Text = u"Contacts"
        g1.Font = self.font_title
        g1.Dock = DockStyle.Fill
        self.g1 = g1
        self.split1.Panel1.Controls.Add(g1)
        self.hdr = []
        for txt in (u"Contact", u"Flip action  v", u"Sel"):
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
        g1.Resize += lambda s, e: self._layout_g1()

        g2 = GroupBox()
        g2.Text = u"Contact view  (contact = red, target = blue, minority direction = orange)"
        g2.Font = self.font_title
        g2.Dock = DockStyle.Fill
        self.g2 = g2
        self.split2.Panel1.Controls.Add(g2)
        self.view = SceneView(self)
        self.view.scene = self._scene
        g2.Controls.Add(self.view)
        v = self.view
        row1 = [(u"Fit", lambda s, e: v.fit(), 100), (u"Fit all", lambda s, e: v.fit_all(), 100),
                (u"Mesh", self._on_mesh, 100), (u"Local: ON", self._on_local_toggle, 100),
                (u"Clear sel", self._on_clear_sel, 100),
                (u"Select reversed", lambda s, e: v.select_reversed(), 130)]
        row2 = [(u"Hide body", self._on_hide_toggle, 120), (u"Undo", lambda s, e: v.undo_hide(), 100),
                (u"Show all", lambda s, e: v.show_all(), 100)]
        self._row_btns = ([], [])
        for ri, row in enumerate((row1, row2)):
            x = 10
            for t, h, w in row:
                bt = self._btn(t, x, 0, w, 30, h)
                bt.Anchor = AnchorStyles.Bottom | AnchorStyles.Left
                g2.Controls.Add(bt)
                self._row_btns[ri].append(bt)
                if t == u"Hide body":
                    self.btn_hide = bt
                if t == u"Local: ON":
                    self.btn_local = bt
                x += w + 6
        g2.Resize += lambda s, e: self._layout_g2()

        g3 = GroupBox()
        g3.Text = u"Whole model  (click or drag to move the contact view)"
        g3.Font = self.font_title
        g3.Dock = DockStyle.Fill
        self.g3 = g3
        self.split2.Panel2.Controls.Add(g3)
        self.overview = OverviewView(self)
        self.overview.scene = self._scene
        self.overview.main = self.view
        g3.Controls.Add(self.overview)
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
        self.btn_close = self._btn(u"Close", 0, 0, 90, 30, None)
        self.btn_close.DialogResult = DialogResult.Cancel
        self.CancelButton = self.btn_close
        self.btn_flip = self._btn(u"Flip Selected", 0, 0, 124, 30, self._on_flip, True)
        for b in (self.btn_back, self.btn_close, self.btn_flip):
            self.Controls.Add(b)

        self._panel = 2
        self._layout()
        if self._cds:
            self._select_row(0)

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
            r["name"].Width = max(40, RW - 160)
            r["cb"].Left = max(46, RW - 154)
            r["cnt"].Left = max(100, RW - 46)
        left = self.rows.Left
        self.hdr[0].Location = Point(left + 4, 28)
        self.hdr[1].Location = Point(left + max(46, RW - 154), 28)
        self.hdr[2].Location = Point(left + max(100, RW - 46), 28)

    def _layout_g2(self):
        W, H = self.g2.Width, self.g2.Height
        self.view.SetBounds(10, 24, max(50, W - 20), max(50, H - 24 - 80))
        for ri, y in ((0, H - 74), (1, H - 38)):
            for bt in self._row_btns[ri]:
                bt.Top = y

    def _layout_g3(self):
        W, H = self.g3.Width, self.g3.Height
        self.overview.SetBounds(10, 24, max(50, W - 20), max(50, H - 34))

    FLIP_CHOICES = (u"none", u"both", u"contact", u"target", u"element")

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
        cb = ComboBox()
        cb.DropDownStyle = ComboBoxStyle.DropDownList
        cb.FlatStyle = FlatStyle.Standard
        cb.Font = self.base_font
        for t in self.FLIP_CHOICES:
            cb.Items.Add(t)
        cb.SelectedItem = cd["side"]
        cb.Location = Point(182, 5)
        cb.Size = Size(102, 24)
        cnt = Label()
        cnt.Text = u"0"
        cnt.Font = self.base_font
        cnt.Location = Point(294, 7)
        cnt.Size = Size(40, 20)
        for ctl in (name, cb, cnt):
            row.Controls.Add(ctl)

        def pick(sender, e, idx=i):
            self._select_row(idx)

        def commit(sender, e, idx=i):
            cd2 = self._cds[idx]
            cd2["side"] = cb.SelectedItem
            self._style_row(idx)
            self._select_row(idx)

        for ctl in (row, name, cnt):
            ctl.Click += pick
        cb.DropDown += pick
        cb.SelectionChangeCommitted += commit
        self.rows.Controls.Add(row)
        r = {"row": row, "name": name, "cb": cb, "cnt": cnt}
        self._rows_style(r, cd)
        return r

    def _rows_style(self, r, cd):
        active = cd["side"] != u"none"
        r["cb"].BackColor = Color.FromArgb(255, 226, 170) if active else Color.White
        r["name"].Font = self.font_bold if active else self.base_font

    def _style_row(self, idx):
        self._rows_style(self._rows[idx], self._cds[idx])

    def _select_row(self, idx):
        cd = self._cds[idx]
        for k, r in enumerate(self._rows):
            r["row"].BackColor = Color.FromArgb(205, 225, 248) if k == idx else Color.White
        if cd is self._cur:
            return
        self._cur = cd
        self.view.segs = cd["segs"]
        self.view.sel = cd["sel"]
        self.view.note = cd["note"]
        self.view.warn = cd.get("warn", u"")
        self.view.local_ids = cd.get("lids")
        self.view.local_box = cd.get("lbox")
        self.overview.segs = cd["segs"]
        self.view.request_fit()

    def selection_changed(self):
        cd = self._cur
        if cd is not None:
            self._rows[cd["idx"]]["cnt"].Text = str(len(cd["sel"]))

    def note(self, text):
        t = text.strip()
        self._lines.append(t)
        self.status.Text = t

    def _on_local_toggle(self, s, e):
        v = self.view
        v.local_on = not v.local_on
        self.btn_local.Text = u"Local: ON" if v.local_on else u"Local: OFF"
        v.fit()

    def _on_mesh(self, s, e):
        self.view.mesh_on = not self.view.mesh_on
        self.view.Invalidate()

    def _on_hide_toggle(self, s, e):
        on = not self.view.hide_mode
        self.view.set_hide_mode(on)
        self.btn_hide.Text = u"Hide body: ON" if on else u"Hide body"

    def _on_clear_sel(self, s, e):
        cd = self._cur
        if cd is not None:
            cd["sel"].clear()
            self.selection_changed()
            self.view.Invalidate()

    def _on_flip(self, s, e):
        n = 0
        for cd in self._cds:
            c = cd["c"]
            old = find_flip_snippet(c)
            side = cd["side"]
            if side == u"element":
                chosen = [cd["segs"][i] for i in sorted(cd["sel"]) if i < len(cd["segs"])]
                if not chosen:
                    self.note(u"%s: Flip = element but no element is selected\r\n" % label_of(c))
                    if old is not None:
                        old.Delete()
                        self.note(u"Flip command removed: %s\r\n" % label_of(c))
                    continue
                text = build_flip_elems(c.Name, chosen)
                n_num = sum(1 for x in chosen if x.get("enum"))
                what = u"%d selected element(s): %d by element number, %d by nodes" % (
                    len(chosen), n_num, len(chosen) - n_num)
            elif side in (u"both", u"contact", u"target"):
                text = build_flip_whole(c.Name, side)
                what = u"all %s elements" % side
            else:                                              # none
                if old is not None:
                    old.Delete()
                    self.note(u"Flip command removed: %s\r\n" % label_of(c))
                continue
            if old is not None:
                old.Delete()
            sn = c.AddCommandSnippet()
            sn.Name = SNIPPET_NAME
            sn.Input = text
            self.note(u"Flip command written (%s): %s\r\n" % (what, label_of(c)))
            n += 1
        if n:
            self.note(u"Done. Flip is applied at the next solve; it is not reflected "
                                u"in the view. Uses cid / tid of each contact.\r\n")


# ---------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------
model = ExtAPI.DataModel.Project.Model
contact_tree = build_contact_tree(model.Connections)
if sum(len(v) for (_, v) in contact_tree) == 0:
    MessageBox.Show(u"No contacts found in the model.", u"Info",
                    MessageBoxButtons.OK, MessageBoxIcon.Information)
else:
    _frm = NormalForm(contact_tree, model)
    _owner = mechanical_owner()
    if _owner is not None:
        _frm.ShowDialog(_owner)
    else:
        _frm.ShowDialog()
