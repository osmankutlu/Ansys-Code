# Ansys Mechanical IronPython araclari icin HIZLI ACILIS yontemi (referans)

Bu not SC03 CONVERTER projesinde bulunan ve ise yarayan yontemi anlatir. Buyuk (binlerce satirlik) IronPython 2.7 araclarinda
"Run'a bas -> ilk pencere gelsin" suresini kisaltmak icin baska projelerde de kullanilabilir.

## 1. Sorun
- Arac ~6200 satirdi. Mechanical'da calistirinca ilk pencere ~8 saniye sonra geliyordu.
- Kodun icine koydugumuz sureolcer 0.1 s gosterdi: yani kayip **kodun ilk satiri calismadan once** oluyordu.
- Cok kucuk bir test dosyasi (bos Form acan, ~15 satir) ayni Mechanical'da neredeyse aninda acildi.
- Sonuc: zaman Mechanical'in baslatma maliyeti degil, **buyuk dosyanin tamaminin okunup derlenmesi**. Dosya buyudukce ilk pencere gecikir,
  cunku ilk satir ancak tum dosya derlenince calisir.

## 2. Teshis nasil yapildi (tekrarlanabilir)
1. Dosyanin basina `t0 = time.time()` koy, giris penceresinin baslik cubuguna `time.time() - t0` yaz.
2. Kronometre ile "Run -> pencere" suresini olc.
3. Iki sayi cok farkliysa (ornek: 0.1 s'ye karsi 8 s) kayip ilk satirdan ONCE = dosya derleme maliyeti.
4. Kesinlestirmek icin ayni ortamda 15 satirlik bos Form dosyasini calistir; anindaysa neden dosya boyutudur.

## 3. Cozum: iki asamali (baslangic + paketli govde)
Tek dosya kalir, ama icerigi ikiye bolunur:
- **Baslangic kismi (kucuk, okunabilir):** sadece ilk pencereyi (ornek: dosya secme penceresi) ve ona gerekenleri icerir. Hemen derlenir, pencere hemen acilir.
- **Govde (buyuk):** gerisi (ana pencere, cizim, hesaplar ...) **sikistirilmis (zlib + base64) bir metin** olarak dosyada durur. Metin sabiti cok hizli
  okunur; derleme **pencere acildiktan sonra, arka plan thread'inde** `exec(compile(...))` ile yapilir.
- Kullanici ilk penceredeki dosyalari secerken govde hazirlanir. Kullanici cok hizliysa `thread.join()` ile (bekleme imleciyle) beklenir.
- Govdede agir bir sey varsa (ornek: C# motorunun `CSharpCodeProvider` ile derlenmesi) ayni arka plan thread'inde, govdeden hemen sonra yapilir.

Akis:
```
Run -> baslangic derlenir (kucuk) -> thread: govdeyi coz + derle (arka plan)
                                  -> StartForm hemen gorunur
Load'a basildi -> govde hazirsa devam, degilse join() ile bekle -> ana pencere: ns["ViewerForm"]() ...
```

## 4. Dosya duzeni
- `X_source.py`: **okunabilir ana kaynak** (tum degisiklikler burada yapilir; surum notlari burada).
- `X.py`: Mechanical'da calistirilan **paketli dosya**; `build_packed.py` ile kaynaktan otomatik uretilir. Elle duzenlenmez.
- Kaynakta `main()` fonksiyonu ve sondaki `main()` cagrisi bulunabilir; paketleyici govdeden bunlari cikarir ve kendi `main()` baslangicini yazar.

## 5. Baska projeye uygularken
1. Kaynak dosyada ilk pencerenin sinifini bul (bizde `StartForm`). `build_packed.py` icindeki `"StartForm"` ismini degistir.
   Paketleyici bu sinifin bagimli oldugu fonksiyon / sabit / import'lari AST ile kendisi bulup baslangica koyar
   (bizde sadece `StartForm`, `TITLE`, `pick_open`, `force_taskbar` ciktı: ~100 satir).
2. Ilk pencere bir "callback" ile ana pencereyi acar. Paketleyicideki `start(m, t, d)` fonksiyonunu kendi akisina uyarla:
   govdedeki siniflara `_BODY["ns"]["SinifAdi"]` ile ulasilir.
3. Govde `exec` ile ayri bir sozluge (namespace) yuklenir; govdenin kendi `import`'lari orada calisir.
4. Bir sey bozulursa govde yuklenemezse hata metni bir MessageBox ile gosterilir.

## 6. Dikkat edilecekler / sinirlar
- Govdedeki hata satir numaralari `SC03_BODY` dosyasina gore gorunur; bunlar `X_source.py` satirlariyla aynidir (govde, `main` cikarilmis haliyle ayni satir duzenindedir).
- Govdede ASCII disi karakterler `?` ile degistirilir (paketleyici yapar); Turkce metinleri `u"\u011f"` gibi kacis dizileriyle yaz.
- Govde ayri thread'de tanimlanir: modul seviyesinde pencere / UI nesnesi olusturma, sadece sinif ve fonksiyon tanimla.
- `__file__` kullanma (govdede yok).
- Tam bu yontem sadece "ilk satira kadar gecen sure" sorunu icindir. Yukleme / hesap sureleri baska konudur.
- Ilk pencereyi baska nedenlerle de yavas yapma: onu olusturmadan once agir is (dosya okuma, derleme) yapma; ana pencereyi (cok kontrollu) **ilk ihtiyacta** kur, baslangicta degil.

## 7. Bu projede ayrica yapilanlar (hizli acilisa katkisi)
- Ana pencere (`ViewerForm`) giris penceresi gosterildikten sonra, dosya secilince kurulur.
- C# motorunun derlenmesi (CSharpCodeProvider) arka planda, giris penceresi acikken baslar.

## 8. Paketleyici script (build_packed.py)
Kullanim: `python3 build_packed.py X_source.py X.py`

```python
# -*- coding: utf-8 -*-
"""Okunabilir kaynak (SC03_CONVERTER_N_source.py) -> hizli acilan tek dosya (SC03_CONVERTER_N.py).
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

header = "# =====================================================================\n#  SC03 CONVERTER (Ansys Mechanical, IronPython 2.7) - hizli acilisli paket\n#  Okunabilir kaynak: SC03_CONVERTER_N_source.py (degisiklik notlari orada). Bu dosyayi duzenleme, kaynaktan yeniden uretilir.\n# ====================================================================="
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

    def start(m, t, d):
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
        ok = f.load_files(m, t, d)
        if ok and not f.Visible:
            f.Show()                                  # sahipsiz, modeless: Mechanical gorev cubugundan one getirilebilir
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
```
