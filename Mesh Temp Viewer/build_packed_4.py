# -*- coding: utf-8 -*-
"""Okunabilir kaynak (SC03_DOMAIN_TEST_N_source.py) -> hizli acilan tek dosya (SC03_DOMAIN_TEST_N.py).
Kullanim: python3 build_packed.py kaynak.py cikti.py
Cikti: kucuk bir baslangic kodu (giris penceresi hemen acilir) + sikistirilmis ana kod (arka planda derlenir)."""
import ast, sys, zlib, base64

src_path, out_path = sys.argv[1], sys.argv[2]
src = open(src_path, "rb").read().decode("utf-8")
tree = ast.parse(src)
lines = src.split("\n")

def seg(n):
    return "\n".join(lines[n.lineno - 1:n.end_lineno])

top = {}
for n in tree.body:
    if isinstance(n, (ast.FunctionDef, ast.ClassDef)):
        top[n.name] = n
    elif isinstance(n, ast.Assign):
        for tg in n.targets:
            if isinstance(tg, ast.Name):
                top[tg.id] = n

def names(n):
    return {x.id for x in ast.walk(n) if isinstance(x, ast.Name)}

need, stack = set(), ["StartForm"]
while stack:
    k = stack.pop()
    if k in need or k not in top:
        continue
    need.add(k)
    stack += list(names(top[k]))
parts = sorted({id(top[k]): top[k] for k in need}.values(), key=lambda n: n.lineno)
used = set()
for n in parts:
    used |= names(n)

# import satirlari (sadece gerekenler)
imps = []
for n in tree.body:
    if isinstance(n, ast.ImportFrom):
        keep = [a for a in n.names if (a.asname or a.name) in used]
        if keep:
            imps.append("from %s import %s" % (n.module, ", ".join(a.name for a in keep)))
stub_head = "import clr\nclr.AddReference(\"System.Windows.Forms\")\nclr.AddReference(\"System.Drawing\")\nimport System\nimport threading\nimport zlib\nimport base64\nimport traceback\n" + "\n".join(imps)

# govde: main() ve sondaki main() cagrisi cikarilir
body_lines = list(lines)
for n in tree.body:
    if isinstance(n, ast.FunctionDef) and n.name == "main":
        for i in range(n.lineno - 1, n.end_lineno):
            body_lines[i] = ""
    if isinstance(n, ast.Expr) and isinstance(n.value, ast.Call) and getattr(n.value.func, "id", "") == "main":
        body_lines[n.lineno - 1] = ""
body = "\n".join(body_lines)
body = "".join(c if ord(c) < 128 else "?" for c in body)
ast.parse(body)
pack = base64.b64encode(zlib.compress(body.encode("ascii"), 9)).decode("ascii")
chunks = [pack[i:i + 100] for i in range(0, len(pack), 100)]

header = "# =====================================================================\n#  SC03 CONVERTER (Ansys Mechanical, IronPython 2.7) - hizli acilisli paket\n#  Okunabilir kaynak: SC03_DOMAIN_TEST_N_source.py (degisiklik notlari orada). Bu dosyayi duzenleme, kaynaktan yeniden uretilir.\n# ====================================================================="
LOADER = '''

# ---------------------------------------------------------------
# Hizli acilis: bu kisim kucuktur, giris penceresi hemen gelir. Asil kod (asagidaki sikistirilmis blok) pencere acikken
# arka planda derlenir; Load'a basildiginda hazirdir.
# ---------------------------------------------------------------
_BODY = {"ns": None, "err": None}


def _load_body():
    try:
        code = zlib.decompress(base64.b64decode("".join(_PACK)))
        ns = {"__name__": "sc03_main"}
        exec(compile(code, "SC03_BODY", "exec"), ns)
        _BODY["ns"] = ns
        try:
            ns["native_cloud_type"]()                 # C# motoru da simdiden derlensin
        except Exception:
            pass
    except Exception:
        _BODY["err"] = traceback.format_exc()


def main():
    th = threading.Thread(target=_load_body)
    th.daemon = True
    th.start()
    holder = []

    def start(m, t, d, prt=u""):
        if th.is_alive():
            Cursors_ = System.Windows.Forms.Cursors
            System.Windows.Forms.Cursor.Current = Cursors_.WaitCursor
            th.join()                                 # kod henuz derleniyorsa bitmesini bekle
            System.Windows.Forms.Cursor.Current = Cursors_.Default
        if _BODY["err"] or _BODY["ns"] is None:
            MessageBox.Show(u"Kod yuklenemedi:\\n\\n%s" % (_BODY["err"],), TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error)
            return False
        if not holder:
            holder.append(_BODY["ns"]["ViewerForm"]())
        f = holder[0]
        ok = f.load_files(m, t, d, prt)
        if ok and not f.Visible:
            f.Show()                                  # sahipsiz, modeless: Mechanical gorev cubugundan one getirilebilir
        if ok:
            try:
                f.bring_front()                       # giris penceresi kapaninca odak Mechanical'a kaymasin
            except Exception:
                pass
        return ok

    sf = StartForm(u"", u"", u"", start)
    sf.Show()


'''
out = ["# -*- coding: utf-8 -*-", header, "", stub_head, ""]
out += [seg(n) + "\n" for n in parts]
out.append(LOADER)
out.append("_PACK = (\n" + "\n".join('    "%s"' % c for c in chunks) + "\n)\n\n\nmain()\n")
open(out_path, "wb").write("\n".join(out).encode("utf-8"))
ast.parse("\n".join(out))
print("stub parts:", sorted(need), "| body bytes:", len(body), "| packed:", len(pack), "| out lines:", len("\n".join(out).split("\n")))
