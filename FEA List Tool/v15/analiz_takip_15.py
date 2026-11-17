# -*- coding: utf-8 -*-
# =============================================================================
#  ANALYSIS TRACKER  (analiz_takip_15.py)
#  Ansys Mechanical IronPython 2.7 + WinForms.  No PowerShell / Excel / macro needed.
#  Run:  Mechanical > Automation > Scripting > Run Script > select this file.
#  Data: one JSON file per record in a shared folder (see VERI_KLASORU below
#        or Admin > Change data folder ...).
# =============================================================================
# ============================================================================
# PART 1 - DATA LAYER  (one JSON file per record in a shared folder)
#   <Data>\Kayitlar\AN-0001.json   <Data>\Silinenler   <Data>\Kullanicilar
#   <Data>\Bildirimler             <Data>\Yedek        <Data>\Listeler.json
# Multi-user safety: per-record lock file + revision (conflict) check.
# ============================================================================
import clr
import os
import re
import json
import uuid
import time
import sys
from datetime import datetime, timedelta

clr.AddReference("System")

from System import Environment, UnauthorizedAccessException, Array, Byte
from System.IO import (File, FileStream, FileMode, FileAccess, FileShare, FileOptions,
                       Directory, Path, DirectoryInfo, IOException)
from System.Text import Encoding, UTF8Encoding

_UTF8 = UTF8Encoding(False)
_SHARE_NONE = getattr(FileShare, "None")      # 'None' is a Python keyword

KAPALI = ["Completed", "Cancelled"]
ROLLER = ["Leader", "Design Engineer", "Structural Analysis Engineer"]

# old (Turkish) values found in old records -> English
DEGER_MAP = {
    u'Beklemede': 'Pending', u'Model Hazırlanıyor': 'Model Preparation',
    u'Analiz Devam Ediyor': 'In Progress', u'İnceleme / Kontrol': 'Review / Check',
    u'Askıda': 'On Hold', u'İptal': 'Cancelled', u'Yüksek': 'High', u'Orta': 'Medium',
    u'Düşük': 'Low', u'Elastik': 'Elastic', u'Elastoplastik': 'Elastoplastic',
    u'Alpha On': 'On', u'Alpha Off': 'Off', u'Bağımsız': 'Standalone',
    u'Gerekmiyor': 'Not Required', u'Bekleniyor': 'Waiting', u'Hazır': 'Ready',
    u'Diğer': 'Other', u'Statik Yapısal': 'Static Structural',
    u'Yorulma (LCF/HCF)': 'Fatigue (LCF/HCF)', u'Sürünme': 'Creep',
    u'Termal-Yapısal': 'Thermal-Structural', u'Kontak / Cıvata': 'Contact / Bolt',
}


def convert_deger(s):
    if not s:
        return s
    parcalar = [DEGER_MAP.get(p, p) for p in re.split(r'\s*,\s*', s)]
    if len(parcalar) > 1:
        return ", ".join(parcalar)
    return DEGER_MAP.get(s, s)


class G(object):
    """module state"""
    Kok = ''
    KayitDir = ''
    SilinenDir = ''
    KullaniciDir = ''
    BildirimDir = ''
    YedekDir = ''
    Klasor = ''                 # program folder (AdminAccounts.txt, settings)
    Ben = None                  # current user (Kullanici)
    HesapOverride = None
    BildirimOnbellek = {}


class Kayit(dict):
    """record: dict with attribute access"""
    def __getattr__(self, k):
        try:
            return self[k]
        except KeyError:
            raise AttributeError(k)

    def __setattr__(self, k, v):
        self[k] = v


class Obj(object):
    def __init__(self, **kw):
        self.__dict__.update(kw)


def _alan(K, L, T, **kw):
    d = {'K': K, 'L': L, 'T': T}
    d.update(kw)
    return d


# Form / record field definitions (the UI builds the form from these)
# Miras: copied from the master model to a sub-analysis.  Bagli: text box next to a combo.
# SadeceDuzenle: not asked in a new request, visible when the record is opened later.
ALANLAR = [
    _alan('ID', 'ID (auto)', 'ro'),
    _alan('Proje', 'Project / Program', 'text', Miras=True),
    _alan('ParcaAdi', 'Part Name(s) *', 'multitext', Miras=True),
    _alan('Layout', 'Layout', 'text', Miras=True),
    _alan('ModelNo', 'Model ID', 'text', Miras=True),
    _alan('ModelRolu', 'Model Role', 'combo', List='ModelRolu', Strict=True),
    _alan('AnaModelID', 'Linked Master Model', 'master'),
    _alan('ModelTipi', 'Model Type', 'combo', List='ModelTipi', Strict=True, Miras=True),
    _alan('AnalizTipi', 'Material Model *', 'combo', List='AnalizTipi', Strict=True, Miras=True),
    _alan('Alpha', 'Alpha', 'combo', List='Alpha', Strict=True, Miras=True),
    _alan('Malzeme', 'Material(s)', 'multitext', List='Malzeme', Miras=True),
    _alan('Amac', 'Purpose', 'memo'),
    _alan('Kosul', 'Condition', 'combo', List='Kosul'),
    _alan('Hiz', 'Speed', 'text'),
    _alan('Aero', 'Aero', 'combo', List='Girdi', Strict=True, Miras=True),
    _alan('AeroAdi', 'Aero Name', 'text', Bagli='Aero', Miras=True),
    _alan('SAS', 'SAS', 'combo', List='Girdi', Strict=True, Miras=True),
    _alan('SASAdi', 'SAS Name', 'text', Bagli='SAS', Miras=True),
    _alan('Termal', 'Thermal', 'combo', List='Girdi', Strict=True, Miras=True),
    _alan('TermalAdi', 'Thermal Name', 'text', Bagli='Termal', Miras=True),
    _alan('AnalizTuru', 'Analysis Type', 'combo', List='AnalizTuru', Miras=True),
    _alan('Oncelik', 'Priority', 'combo', List='Oncelik', Strict=True),
    _alan('TalepEden', 'Requester', 'combo', List='TalepEden', Strict=True),
    _alan('AnalizYapan', 'Analyst', 'combo', List='AnalizYapan', Strict=True),
    _alan('TalepTarihi', 'Request Date', 'date'),
    _alan('Deadline', 'Due Date', 'date'),
    _alan('TalepNotu', 'Requester Notes', 'memo'),
    _alan('Durum', 'Status', 'combo', List='Durum', Strict=True, SadeceDuzenle=True),
    _alan('YapanNotu', 'Analyst Notes', 'memo', SadeceDuzenle=True),
    _alan('ModelYolu', 'Model File Path', 'text', SadeceDuzenle=True),
    _alan('SonucYolu', 'Results / Report Path', 'text', SadeceDuzenle=True),
    _alan('Tamamlanma', 'Completion Date', 'date', SadeceDuzenle=True),
    _alan('Gecmis', 'Change History', 'hist', SadeceDuzenle=True),
]
# fields filled by the analyst (a Design Engineer cannot edit them)
for _a in ALANLAR:
    if _a['K'] in ('Durum', 'YapanNotu', 'ModelYolu', 'SonucYolu', 'Tamamlanma'):
        _a['Analizci'] = True
# form section headers (the first field of a section carries the header)
BOLUMLER = {'Proje': 'Project and part', 'ModelRolu': 'Model structure',
            'ModelTipi': 'Model and analysis definition', 'Amac': 'Purpose and load condition',
            'Aero': 'Input data', 'AnalizTuru': 'Request details', 'Durum': 'Analyst section'}
for _a in ALANLAR:
    if _a['K'] in BOLUMLER:
        _a['Bolum'] = BOLUMLER[_a['K']]

DEGER_ALANLARI = ('Durum', 'Oncelik', 'ModelRolu', 'Aero', 'SAS', 'Termal', 'Malzeme', 'Kosul',
                  'AnalizTuru', 'AnalizTipi', 'Alpha', 'ModelTipi')


def alan(k):
    for a in ALANLAR:
        if a['K'] == k:
            return a
    return None


def get_varsayilan_listeler():
    return [
        ('Durum', ['Pending', 'Model Preparation', 'In Progress', 'Review / Check', 'On Hold', 'Completed', 'Cancelled']),
        ('Oncelik', ['High', 'Medium', 'Low']),
        ('ModelTipi', ['2D', '3D']),
        ('AnalizTipi', ['Elastic', 'Elastoplastic']),
        ('Alpha', ['On', 'Off']),
        ('ModelRolu', ['Standalone', 'Master Model', 'Sub-Analysis']),
        ('Girdi', ['Not Required', 'Waiting', 'Ready']),
        ('Malzeme', ['IN718', 'Hastelloy X', 'Waspaloy', 'Ti-6Al-4V', '17-4PH', 'Other']),
        ('AnalizTuru', ['Static Structural', 'Prestressed Modal', 'Campbell / SAFE', 'Fatigue (LCF/HCF)', 'Creep',
                        'Thermal-Structural', 'Contact / Bolt', 'Cyclic Symmetry', 'Other']),
        ('Kosul', ['Take-off', 'Max Continuous', 'Cruise', 'Idle', 'Overspeed', 'Burst', 'Transient', 'Other']),
        ('TalepEden', []),
        ('AnalizYapan', []),
    ]


# ---------------------------------------------------------------- file helpers
def now_s():
    return datetime.now().strftime('%Y-%m-%dT%H:%M:%S')


def read_text(path):
    return File.ReadAllText(path, Encoding.UTF8)


def write_dosya_atomik(yol, metin):
    tmp = yol + '.' + uuid.uuid4().hex + '.tmp'
    File.WriteAllText(tmp, metin, _UTF8)
    if File.Exists(yol):
        File.Replace(tmp, yol, None)
    else:
        File.Move(tmp, yol)


def to_json(obj):
    return json.dumps(obj, indent=2, ensure_ascii=False)


def enter_kilit(ad):
    """per-record lock: FileShare.None + DeleteOnClose (released even if the process dies)"""
    kp = os.path.join(G.KayitDir, ad + '.lock')
    for i in range(60):
        try:
            return FileStream(kp, FileMode.CreateNew, FileAccess.Write, _SHARE_NONE, 1, FileOptions.DeleteOnClose)
        except (IOException, UnauthorizedAccessException):
            time.sleep(0.1)
    raise Exception('The record is currently being saved by another user. Please try again in a few seconds.')


def initialize_veri(kok):
    G.Kok = kok
    G.KayitDir = os.path.join(kok, 'Kayitlar')
    G.SilinenDir = os.path.join(kok, 'Silinenler')
    G.KullaniciDir = os.path.join(kok, 'Kullanicilar')
    G.BildirimDir = os.path.join(kok, 'Bildirimler')
    G.YedekDir = os.path.join(kok, 'Yedek')
    for d in (kok, G.KayitDir, G.SilinenDir, G.KullaniciDir, G.BildirimDir, G.YedekDir):
        if not Directory.Exists(d):
            Directory.CreateDirectory(d)
    if not File.Exists(os.path.join(kok, 'Listeler.json')):
        save_listeler(dict(get_varsayilan_listeler()))


# ---------------------------------------------------------------- lists
def get_listeler():
    vars_ = get_varsayilan_listeler()
    sonuc = {}
    for k, v in vars_:
        sonuc[k] = list(v)
    yol = os.path.join(G.Kok, 'Listeler.json')
    try:
        o = json.loads(read_text(yol))
        for k, v in vars_:
            if k in o and o[k] is not None:
                if isinstance(o[k], list):
                    sonuc[k] = [convert_deger(x) for x in o[k]]
                else:
                    sonuc[k] = [convert_deger(o[k])]
    except Exception:
        pass
    d = dict(vars_)
    for k in ('ModelTipi', 'AnalizTipi', 'Alpha', 'ModelRolu', 'Girdi'):
        sonuc[k] = list(d[k])
    for kp in KAPALI:
        if kp not in sonuc['Durum']:
            sonuc['Durum'] = sonuc['Durum'] + [kp]
    return sonuc


def save_listeler(listeler):
    h = {}
    for k in listeler:
        h[k] = [unicode(x).strip() for x in listeler[k] if x is not None and unicode(x).strip() != '']
    write_dosya_atomik(os.path.join(G.Kok, 'Listeler.json'), to_json(h))


# ---------------------------------------------------------------- record object
def _s(v):
    if v is None:
        return ''
    if isinstance(v, datetime):
        return v.strftime('%Y-%m-%d')
    if isinstance(v, bool):
        return 'True' if v else 'False'
    return v if isinstance(v, basestring) else unicode(v)


def convert_to_kayit(o):
    """o: dict-like (or None) -> normalized Kayit"""
    if o is None:
        o = {}
    h = Kayit()
    for a in ALANLAR:
        if a['T'] != 'hist':
            v = _s(o.get(a['K']))
            if a['K'] in DEGER_ALANLARI:
                v = convert_deger(v)
            h[a['K']] = v
    h['Olusturan'] = _s(o.get('Olusturan'))
    try:
        h['Revizyon'] = int(o.get('Revizyon') or 0)
    except Exception:
        h['Revizyon'] = 0
    h['SonGuncelleyen'] = _s(o.get('SonGuncelleyen'))
    h['SonGuncelleme'] = _s(o.get('SonGuncelleme'))
    g = o.get('Gecmis') or []
    if isinstance(g, basestring):
        g = [g]
    h['Gecmis'] = [_s(x) for x in g]
    return h


def new_bos_kayit():
    return convert_to_kayit({})


def kayit_json(rec):
    sira = [a['K'] for a in ALANLAR if a['T'] != 'hist'] + ['Olusturan', 'Revizyon', 'SonGuncelleyen', 'SonGuncelleme', 'Gecmis']
    od = []
    for k in sira:
        od.append((k, rec.get(k)))
    # ordered dump
    parts = []
    for k, v in od:
        parts.append(json.dumps(k) + ': ' + json.dumps(v, ensure_ascii=False))
    return '{\n  ' + ',\n  '.join(parts) + '\n}'


def read_kayit(yol):
    for i in range(4):
        try:
            return convert_to_kayit(json.loads(read_text(yol)))
        except Exception:
            time.sleep(0.12)
    return None


def get_tum_kayitlar():
    liste = []
    for f in Directory.GetFiles(G.KayitDir, '*.json'):
        r = read_kayit(f)
        if r is not None and r['ID']:
            liste.append(r)
    liste.sort(key=lambda r: r['ID'])
    return liste


def get_veri_imzasi():
    """did anything change in the folder? (count + last write times)"""
    di = DirectoryInfo(G.KayitDir)
    n = 0
    t = 0
    for f in di.GetFiles('*.json'):
        n += 1
        t += f.LastWriteTimeUtc.Ticks % 1000000007
    return '%d|%d' % (n, t)


# ---------------------------------------------------------------- dates / account
def convert_to_tarih(s):
    if not s:
        return None
    try:
        return datetime.strptime(s, '%Y-%m-%d')
    except Exception:
        return None


def convert_from_tarih(d):
    if d is None:
        return ''
    return d.strftime('%Y-%m-%d')


def bugun():
    n = datetime.now()
    return datetime(n.year, n.month, n.day)


def test_kapali(r):
    return r['Durum'] in KAPALI


def get_kalan_gun(r):
    if test_kapali(r):
        return None
    d = convert_to_tarih(r['Deadline'])
    if d is None:
        return None
    return (d - bugun()).days


def test_gecikti(r):
    k = get_kalan_gun(r)
    return k is not None and k < 0


def get_degisiklikler(eski, yeni):
    satirlar = []
    for a in ALANLAR:
        if a['T'] in ('ro', 'hist'):
            continue
        e = _s(eski.get(a['K']))
        y = _s(yeni.get(a['K']))
        if e != y:
            if len(e) > 50:
                e = e[:50] + '...'
            if len(y) > 50:
                y = y[:50] + '...'
            satirlar.append(u'%s: "%s" -> "%s"' % (a['L'].replace(' *', ''), e, y))
    return satirlar


def get_kullanici_adi():
    if G.Ben is not None:
        return G.Ben['Ad']
    return Environment.UserName


def get_hesap_adi():
    if G.HesapOverride:
        return unicode(G.HesapOverride).lower()
    return (u'%s\\%s' % (Environment.UserDomainName, Environment.UserName)).lower()


# ---------------------------------------------------------------- users / roles / permissions
def get_kullanici_dosyasi(hesap):
    g = re.sub(r'[^a-z0-9._-]', '_', hesap.lower())
    return os.path.join(G.KullaniciDir, g + '.json')


def convert_to_kullanici(o):
    adm = False
    if o.get('Admin') is not None:
        try:
            adm = bool(o.get('Admin'))
        except Exception:
            adm = False
    k = Kayit()
    k['Hesap'] = _s(o.get('Hesap')).lower()
    k['Ad'] = _s(o.get('Ad'))
    k['Rol'] = _s(o.get('Rol'))
    k['Admin'] = adm
    k['Bilgisayar'] = _s(o.get('Bilgisayar'))
    k['Kayit'] = _s(o.get('Kayit'))
    k['SonGiris'] = _s(o.get('SonGiris'))
    k['SonGorulen'] = _s(o.get('SonGorulen'))
    return k


def kullanici_json(k):
    return to_json(dict((x, k[x]) for x in k))


def get_kullanici(hesap):
    """None if the file does not exist; raises if it exists but cannot be read
    (a temporary network problem must not look like 'user deleted')"""
    if not Directory.Exists(G.KullaniciDir):
        raise Exception('The user folder cannot be reached.')
    yol = get_kullanici_dosyasi(hesap)
    if not File.Exists(yol):
        return None
    son = None
    for i in range(4):
        try:
            k = convert_to_kullanici(json.loads(read_text(yol)))
            if k['Hesap'] != hesap.lower():
                return None
            return k
        except Exception as e:
            son = e
            time.sleep(0.12)
    raise son


def get_tum_kullanicilar():
    liste = []
    for f in Directory.GetFiles(G.KullaniciDir, '*.json'):
        try:
            k = convert_to_kullanici(json.loads(read_text(f)))
            if k['Hesap'] and k['Ad']:
                liste.append(k)
        except Exception:
            pass
    liste.sort(key=lambda k: k['Ad'].lower())
    return liste


def save_kullanici(k):
    write_dosya_atomik(get_kullanici_dosyasi(k['Hesap']), kullanici_json(k))


def remove_kullanici(hesap):
    yol = get_kullanici_dosyasi(hesap)
    if File.Exists(yol):
        File.Delete(yol)


def test_ad_kullanimda(ad, haric_hesap=''):
    for k in get_tum_kullanicilar():
        if k['Hesap'] != haric_hesap.lower() and k['Ad'].lower() == ad.strip().lower():
            return True
    return False


def new_kullanici(hesap, ad, rol):
    """first-run registration. The very first person becomes the administrator."""
    ad = (ad or '').strip()
    if len(ad) < 2:
        raise Exception('Please enter your name (at least 2 characters).')
    if rol not in ROLLER:
        raise Exception('Please select your role.')
    if test_ad_kullanimda(ad, hesap):
        raise Exception("The name '%s' is already registered by another user. Please enter your full name." % ad)
    ilk = (len(get_tum_kullanicilar()) == 0)
    simdi = now_s()
    k = convert_to_kullanici({'Hesap': hesap.lower(), 'Ad': ad, 'Rol': rol, 'Admin': ilk,
                              'Bilgisayar': Environment.MachineName, 'Kayit': simdi, 'SonGiris': simdi, 'SonGorulen': simdi})
    save_kullanici(k)
    return k


def test_admin(k):
    """admin: Admin flag in the user file, or AdminAccounts.txt next to the program (one Windows account per line)"""
    if k is None:
        return False
    if k['Admin']:
        return True
    try:
        if G.Klasor:
            p = os.path.join(G.Klasor, 'AdminAccounts.txt')
            if File.Exists(p):
                for s in File.ReadAllLines(p):
                    t = s.strip().lower()
                    if t and not t.startswith('#') and t == k['Hesap']:
                        return True
    except Exception:
        pass
    return False


def test_tam_yetki(k):
    """Leader, Structural Analysis Engineer and Admin can edit every record"""
    if k is None:
        return False
    return test_admin(k) or k['Rol'] == 'Leader' or k['Rol'] == 'Structural Analysis Engineer'


def get_yetki(rec, k):
    """Mod: Full (all fields) | Talep (request fields only, own record) | Gor (read only)"""
    y = Obj(Mod='Gor', Sil=False)
    if k is None:
        return y
    if test_tam_yetki(k):
        y.Mod = 'Full'
        y.Sil = True
        return y
    benim = (bool(rec['Olusturan']) and rec['Olusturan'] == k['Hesap']) or (bool(rec['TalepEden']) and rec['TalepEden'] == k['Ad'])
    if benim and not test_kapali(rec):
        y.Mod = 'Talep'
        y.Sil = (rec['Durum'] == get_listeler()['Durum'][0])
    return y


# ---------------------------------------------------------------- notifications
def new_bildirim_dosyasi(kayit_id, tur, mesaj, alicilar):
    al = []
    for a in alicilar:
        if a and a not in al:
            al.append(a)
    if not al:
        return
    ben = G.Ben
    o = {'Zaman': now_s(), 'KayitID': kayit_id, 'Tur': tur, 'Mesaj': mesaj,
         'Kimden': ben['Ad'] if ben else '', 'KimdenHesap': ben['Hesap'] if ben else '', 'Alicilar': al}
    ad = '%s_%s.json' % (datetime.now().strftime('%Y%m%d%H%M%S') + '%03d' % (datetime.now().microsecond // 1000), uuid.uuid4().hex[:8])
    write_dosya_atomik(os.path.join(G.BildirimDir, ad), to_json(o))


def send_bildirim(rec, eski, tur):
    """tur: Yeni | Degisti | Silindi
    Rules: Leader never gets notifications; nobody is notified about their own action.
     - Design Engineer acts: the assigned analyst only, or all Structural Analysis Engineers if none assigned
     - Other roles act: the requester; a newly assigned analyst is also notified"""
    ben = G.Ben
    if ben is None:
        return
    kul = get_tum_kullanicilar()
    degis = []
    if tur == 'Degisti':
        degis = get_degisiklikler(eski, rec)
        if not degis:
            return
    analiz = rec['AnalizYapan']
    hedef = []
    if ben['Rol'] == 'Design Engineer':
        if analiz:
            hedef = [k for k in kul if k['Ad'] == analiz]
        if not hedef:
            hedef = [k for k in kul if k['Rol'] == 'Structural Analysis Engineer']
    else:
        if rec['Olusturan']:
            hedef += [k for k in kul if k['Hesap'] == rec['Olusturan']]
        if rec['TalepEden']:
            hedef += [k for k in kul if k['Ad'] == rec['TalepEden']]
        atandi = (tur == 'Yeni') or (eski is not None and eski['AnalizYapan'] != analiz)
        if atandi and analiz:
            hedef += [k for k in kul if k['Ad'] == analiz]
    alicilar = [k['Hesap'] for k in hedef if k['Hesap'] != ben['Hesap'] and k['Rol'] != 'Leader']
    if not alicilar:
        return
    parca = rec['ParcaAdi']
    if len(parca) > 40:
        parca = parca[:40] + '...'
    if tur == 'Yeni':
        m = u'New request %s: %s by %s' % (rec['ID'], parca, ben['Ad'])
    elif tur == 'Silindi':
        m = u'%s (%s) was deleted by %s' % (rec['ID'], parca, ben['Ad'])
    else:
        m = u'%s (%s) updated by %s: ' % (rec['ID'], parca, ben['Ad']) + '; '.join(degis[:2])
        if len(degis) > 2:
            m += '; +%d more' % (len(degis) - 2)
    new_bildirim_dosyasi(rec['ID'], tur, m, alicilar)


def get_benim_bildirimler(k):
    sonuc = []
    for f in Directory.GetFiles(G.BildirimDir, '*.json'):
        ad = Path.GetFileName(f)
        if ad not in G.BildirimOnbellek:
            try:
                o = json.loads(read_text(f))
                al = o.get('Alicilar') or []
                if isinstance(al, basestring):
                    al = [al]
                G.BildirimOnbellek[ad] = Kayit(Dosya=ad, Zaman=_s(o.get('Zaman')), KayitID=_s(o.get('KayitID')),
                                               Tur=_s(o.get('Tur')), Mesaj=_s(o.get('Mesaj')), Kimden=_s(o.get('Kimden')),
                                               Alicilar=list(al))
            except Exception:
                continue
        b = G.BildirimOnbellek[ad]
        if k['Hesap'] in b['Alicilar']:
            sonuc.append(b)
    sonuc.sort(key=lambda b: b['Zaman'], reverse=True)
    out = []
    for b in sonuc:
        c = Kayit(b)
        c['Okunmadi'] = b['Zaman'] > _s(k['SonGorulen'])
        out.append(c)
    return out


def set_bildirim_okundu(k):
    u = get_kullanici(k['Hesap'])
    if u is None:
        return k
    u['SonGorulen'] = now_s()
    save_kullanici(u)
    return u


def remove_eski_bildirimler():
    """delete notification files older than 30 days"""
    sinir = (datetime.now() - timedelta(days=30)).strftime('%Y%m%d%H%M%S')
    for f in Directory.GetFiles(G.BildirimDir, '*.json'):
        ad = Path.GetFileName(f)
        if len(ad) > 14 and ad[:14] < sinir:
            try:
                File.Delete(f)
            except Exception:
                pass


# ---------------------------------------------------------------- admin tools
def update_kayitlarda_ad(eski, yeni, anahtarlar, sadece_acik):
    """bulk rename/reassign a person name in records"""
    ok = 0
    hata = 0
    for r in get_tum_kayitlar():
        if sadece_acik and test_kapali(r):
            continue
        n = convert_to_kayit(r)
        degisti = False
        for k in anahtarlar:
            if n[k] == eski:
                n[k] = yeni
                degisti = True
        if not degisti:
            continue
        try:
            save_kayit(n, r)
            ok += 1
        except Exception:
            hata += 1
    return Obj(Guncellenen=ok, Hatali=hata)


def get_silinen_kayitlar():
    liste = []
    for f in Directory.GetFiles(G.SilinenDir, '*.json'):
        ad = Path.GetFileNameWithoutExtension(f)
        m = re.match(r'^(AN-\d+)_(\d{14})$', ad)
        if m:
            r = read_kayit(f)
            if r is None:
                continue
            z = datetime.strptime(m.group(2), '%Y%m%d%H%M%S')
            liste.append(Obj(Dosya=f, ID=m.group(1), Silinme=z, Parca=r['ParcaAdi'], Proje=r['Proje'], Talep=r['TalepEden']))
    liste.sort(key=lambda x: x.Silinme, reverse=True)
    return liste


def restore_kayit(dosya, id_):
    hedef = os.path.join(G.KayitDir, id_ + '.json')
    if File.Exists(hedef):
        raise Exception('A record with ID %s already exists.' % id_)
    File.Move(dosya, hedef)


def backup_veri():
    """copy records, deleted records, users and lists to Yedek\\<date>; keep the last 10 backups"""
    hedef = os.path.join(G.YedekDir, datetime.now().strftime('%Y%m%d_%H%M%S'))
    Directory.CreateDirectory(hedef)
    n = 0
    for d in ('Kayitlar', 'Silinenler', 'Kullanicilar'):
        kd = os.path.join(hedef, d)
        Directory.CreateDirectory(kd)
        for f in Directory.GetFiles(os.path.join(G.Kok, d), '*.json'):
            try:
                File.Copy(f, os.path.join(kd, Path.GetFileName(f)))
                n += 1
            except Exception:
                pass
    lp = os.path.join(G.Kok, 'Listeler.json')
    if File.Exists(lp):
        File.Copy(lp, os.path.join(hedef, 'Listeler.json'))
    eskiler = sorted([Path.GetFileName(d) for d in Directory.GetDirectories(G.YedekDir)])
    if len(eskiler) > 10:
        for e in eskiler[:len(eskiler) - 10]:
            try:
                Directory.Delete(os.path.join(G.YedekDir, e), True)
            except Exception:
                pass
    return Obj(Yol=hedef, Dosya=n)


# ---------------------------------------------------------------- save / delete
def new_kayit_dosyasi(rec):
    """assign a new ID; the file is created with CreateNew (two people cannot get the same ID)"""
    mx = 0
    for f in Directory.GetFiles(G.KayitDir, '*.json'):
        m = re.match(r'^AN-(\d+)$', Path.GetFileNameWithoutExtension(f))
        if m and int(m.group(1)) > mx:
            mx = int(m.group(1))
    for f in Directory.GetFiles(G.SilinenDir, '*.json'):
        m = re.match(r'^AN-(\d+)_', Path.GetFileName(f))
        if m and int(m.group(1)) > mx:
            mx = int(m.group(1))
    for i in range(1, 101):
        id_ = 'AN-%04d' % (mx + i)
        yol = os.path.join(G.KayitDir, id_ + '.json')
        try:
            fs = FileStream(yol, FileMode.CreateNew, FileAccess.Write, _SHARE_NONE)
        except (IOException, UnauthorizedAccessException):
            continue
        try:
            rec['ID'] = id_
            b = _UTF8.GetBytes(kayit_json(rec))
            fs.Write(b, 0, b.Length)
        finally:
            fs.Dispose()
        return id_
    raise Exception('Could not assign a new ID.')


def save_kayit(rec, eski):
    """new record: eski = None.  existing record: eski = copy read when opened (conflict check)"""
    simdi = now_s()
    kul = get_kullanici_adi()
    rec['SonGuncelleyen'] = kul
    rec['SonGuncelleme'] = simdi
    zaman = datetime.now().strftime('%Y-%m-%d %H:%M')
    if eski is None:
        rec['Olusturan'] = get_hesap_adi()
        rec['Revizyon'] = 1
        rec['Gecmis'] = [u'%s | %s | Record created' % (zaman, kul)]
        return new_kayit_dosyasi(rec)
    yol = os.path.join(G.KayitDir, rec['ID'] + '.json')
    kilit = enter_kilit(rec['ID'])
    try:
        mevcut = read_kayit(yol)
        if mevcut is None:
            raise Exception('Record file not found (someone else may have deleted it).')
        if mevcut['Revizyon'] != eski['Revizyon']:
            raise Exception('CAKISMA')
        rec['Revizyon'] = mevcut['Revizyon'] + 1
        yeni = [u'%s | %s | %s' % (zaman, kul, x) for x in get_degisiklikler(mevcut, rec)]
        g = list(mevcut['Gecmis']) + yeni
        if len(g) > 80:
            g = g[len(g) - 80:]
        rec['Gecmis'] = g
        write_dosya_atomik(yol, kayit_json(rec))
    finally:
        kilit.Dispose()
    return rec['ID']


def remove_kayit(id_):
    yol = os.path.join(G.KayitDir, id_ + '.json')
    kilit = enter_kilit(id_)
    try:
        if not File.Exists(yol):
            return
        hedef = os.path.join(G.SilinenDir, '%s_%s.json' % (id_, datetime.now().strftime('%Y%m%d%H%M%S')))
        File.Move(yol, hedef)
    finally:
        kilit.Dispose()


# ---------------------------------------------------------------- validation
def get_cocuklar(tum, id_):
    return [r for r in tum if r['ModelRolu'] == 'Sub-Analysis' and r['AnaModelID'] == id_]


def test_kayit(rec, tum, eski):
    """list of error messages (empty = OK)"""
    h = []
    if not (rec['ParcaAdi'] or '').strip():
        h.append('Part name cannot be empty.')
    if not (rec['AnalizTipi'] or '').strip():
        h.append('Select a material model (Elastic or Elastoplastic). Use Duplicate to request the same analysis with another option.')
    dt1 = convert_to_tarih(rec['TalepTarihi'])
    dt2 = convert_to_tarih(rec['Deadline'])
    if dt1 is not None and dt2 is not None and dt2 < dt1:
        h.append('Due date cannot be earlier than the request date.')
    if rec['ModelRolu'] == 'Sub-Analysis':
        if not (rec['AnaModelID'] or '').strip():
            h.append('A Sub-Analysis must have a Linked Master Model.')
        elif rec['AnaModelID'] == rec['ID']:
            h.append('A record cannot be its own sub-analysis.')
        else:
            m = [x for x in tum if x['ID'] == rec['AnaModelID']]
            if not m:
                h.append('Linked master model not found: %s' % rec['AnaModelID'])
            elif m[0]['ModelRolu'] != 'Master Model':
                h.append("Record %s does not have the Model Role 'Master Model'." % rec['AnaModelID'])
    if eski is not None and eski['ModelRolu'] == 'Master Model' and rec['ModelRolu'] != 'Master Model':
        c = get_cocuklar(tum, rec['ID'])
        if c:
            h.append('This record is the master model of %d sub-analysis(es). Link them to another master model or delete them before changing its role.' % len(c))
    return h


# ---------------------------------------------------------------- summary / export
def get_ozet(tum):
    acik = [r for r in tum if not test_kapali(r)]
    o = Obj()
    o.Toplam = len(tum)
    o.Acik = len(acik)
    o.Tamam = len([r for r in tum if r['Durum'] == 'Completed'])
    o.Geciken = len([r for r in acik if test_gecikti(r)])
    o.Yakin7 = len([r for r in acik if get_kalan_gun(r) is not None and 0 <= get_kalan_gun(r) <= 7])
    o.AnaModel = len([r for r in tum if r['ModelRolu'] == 'Master Model'])
    o.AltAnaliz = len([r for r in tum if r['ModelRolu'] == 'Sub-Analysis'])
    d = {}
    for r in tum:
        d[r['Durum']] = d.get(r['Durum'], 0) + 1
    o.Durumlar = [Obj(Ad=k, Adet=d[k]) for k in sorted(d)]
    kisi = {}
    for r in acik:
        kisi.setdefault(r['AnalizYapan'], []).append(r)
    o.Kisiler = [Obj(Ad=(k if k else '(unassigned)'), Acik=len(kisi[k]), Geciken=len([x for x in kisi[k] if test_gecikti(x)])) for k in sorted(kisi)]
    yak = [r for r in acik if get_kalan_gun(r) is not None]
    yak.sort(key=lambda r: get_kalan_gun(r))
    o.Yaklasan = yak[:10]
    return o


def export_kayitlar_csv(tum, yol):
    basliklar = []
    for a in ALANLAR:
        if a['T'] == 'hist':
            continue
        basliklar.append(a['L'].replace(' *', '').replace(' (auto)', ''))
    basliklar += ['Days Left', 'Overdue', 'Last Updated By']

    def q(v):
        return '"' + _s(v).replace('"', '""') + '"'
    satirlar = [';'.join(q(b) for b in basliklar)]
    for r in tum:
        v = []
        for a in ALANLAR:
            if a['T'] == 'hist':
                continue
            x = _s(r.get(a['K']))
            if a['T'] == 'date' and x:
                d = convert_to_tarih(x)
                if d:
                    x = d.strftime('%d.%m.%Y')
            v.append(x)
        k = get_kalan_gun(r)
        v.append('' if k is None else str(k))
        v.append('OVERDUE' if test_gecikti(r) else '')
        v.append(r['SonGuncelleyen'])
        satirlar.append(';'.join(q(x) for x in v))
    File.WriteAllText(yol, u'\r\n'.join(satirlar) + u'\r\n', UTF8Encoding(True))


# ---------------------------------------------------------------- sample data
def add_ornek_veri():
    """sample data: 8 master models (2-3 sub-analyses each) + 4 standalone analyses.
    Names come from the registered users (designers / analysts); generic names otherwise."""
    bg = datetime.now()
    kul = get_tum_kullanicilar()
    tas = [k['Ad'] for k in kul if k['Rol'] == 'Design Engineer']
    ana = [k['Ad'] for k in kul if k['Rol'] == 'Structural Analysis Engineer']

    def kisi(liste, n, genel):
        if not liste:
            return '%s %d' % (genel, n)
        return liste[(n - 1) % len(liste)]

    def yeni_ornek(ortak, o):
        r = new_bos_kayit()
        for k in ortak:
            r[k] = ortak[k]
        for k in o:
            if not k.startswith('_'):
                r[k] = o[k]
        r['TalepEden'] = kisi(tas, o['_r'], 'Requester')
        if o['_a'] > 0:
            r['AnalizYapan'] = kisi(ana, o['_a'], 'Engineer')
        r['TalepTarihi'] = convert_from_tarih(bg + timedelta(days=o['_t']))
        r['Deadline'] = convert_from_tarih(bg + timedelta(days=o['_d']))
        if '_b' in o:
            r['Tamamlanma'] = convert_from_tarih(bg + timedelta(days=o['_b']))
        return r

    def S(amac, kosul, hiz, durum, onc, r, a, nt, t, d, b=None, an=''):
        h = {'Amac': amac, 'Kosul': kosul, 'Hiz': hiz, 'Durum': durum, 'Oncelik': onc, '_r': r, '_a': a, 'TalepNotu': nt, '_t': t, '_d': d}
        if b is not None:
            h['_b'] = b
        if an:
            h['YapanNotu'] = an
        return h
    G_ = 'Ready'
    B_ = 'Waiting'
    N_ = 'Not Required'

    def O(**kw):
        return kw
    modeller = [
        dict(Ortak=O(Proje='Sample Engine', ParcaAdi='HPT Stage 1 Disk', Layout='Layout A', ModelNo='HPT-D1-M01', ModelTipi='3D', AnalizTipi='Elastic', Alpha='On', Malzeme='IN718', AnalizTuru='Cyclic Symmetry', Aero=G_, AeroAdi='Aero_Load_RevC', SAS=G_, SASAdi='SAS_Temp_RevB', Termal=G_, TermalAdi='Therm_Map_RevB'),
             Ana=S('Fatigue life verification', 'Take-off', '105% speed', 'In Progress', 'High', 1, 1, 'Aero loads rev C', -20, 16, None, 'Cyclic symmetry model set up'),
             Alt=[S('Overspeed check', 'Overspeed', '120% speed', 'Pending', 'Medium', 1, 2, 'Speed 120%, temperature field unchanged', -2, 30),
                  S('Max continuous strength', 'Max Continuous', '100% speed', 'Completed', 'Medium', 2, 1, 'Only speed and pressure loads changed', -15, -8, -9, 'Results delivered, margin OK'),
                  S('Burst margin', 'Burst', '122% speed', 'Review / Check', 'High', 3, 1, 'Check rim burst margin', -9, 6)]),
        dict(Ortak=O(Proje='Sample Engine', ParcaAdi='LPT Rotor Blade', Layout='Layout B', ModelNo='LPT-B-M02', ModelTipi='3D', AnalizTipi='Elastic', Alpha='Off', Malzeme='Hastelloy X', AnalizTuru='Prestressed Modal', Aero=G_, AeroAdi='LPT_Aero_Pressure_R2', SAS=B_, SASAdi='LPT_SAS_Temp', Termal=N_, TermalAdi=''),
             Ana=S('Modal analysis for Campbell diagram', 'Cruise', '100% speed', 'Model Preparation', 'High', 2, 2, 'Waiting for SAS temperatures', -18, -4),
             Alt=[S('Modal at idle speed', 'Idle', '60% speed', 'Pending', 'Low', 2, 2, '6000 rpm, prestress updated', -3, 20),
                  S('Overspeed modal check', 'Overspeed', '115% speed', 'In Progress', 'Medium', 1, 1, '115% speed, temperature field unchanged', -6, 9)]),
        dict(Ortak=O(Proje='Sample Engine 2', ParcaAdi='Labyrinth Seal; Seal Carrier', Layout='Layout A', ModelNo='LAB-M03', ModelTipi='2D', AnalizTipi='Elastoplastic', Alpha='On', Malzeme='Hastelloy X; IN718', AnalizTuru='Thermal-Structural', Aero=N_, AeroAdi='', SAS=G_, SASAdi='Seal_SAS_Temp', Termal=G_, TermalAdi='Seal_Therm_Map'),
             Ana=S('Thermal stress check', 'Transient', '100% speed', 'In Progress', 'Medium', 1, 1, 'Transient curve attached', -10, 12),
             Alt=[S('Fast acceleration transient', 'Transient', '100% speed', 'Pending', 'Medium', 2, 2, 'Thermal transient curve updated', -1, 25),
                  S('Burst condition', 'Burst', '122% speed', 'Review / Check', 'High', 2, 1, 'Speed 122%, thermal load removed', -12, 5)]),
        dict(Ortak=O(Proje='Sample Engine 2', ParcaAdi='HPC Stage 3 Disk', Layout='Layout C', ModelNo='HPC-D3-M04', ModelTipi='3D', AnalizTipi='Elastic', Alpha='On', Malzeme='Waspaloy', AnalizTuru='Cyclic Symmetry', Aero=G_, AeroAdi='HPC_Aero_Load_R1', SAS=G_, SASAdi='HPC_SAS_R3', Termal=G_, TermalAdi='HPC_Therm_R2'),
             Ana=S('LCF life assessment', 'Take-off', '103% speed', 'In Progress', 'High', 3, 2, 'Use updated thermal map R2', -14, 20, None, 'Mesh convergence study running'),
             Alt=[S('Cruise steady state', 'Cruise', '95% speed', 'Pending', 'Medium', 3, 2, 'Same mesh as master', -2, 28),
                  S('Max continuous strength', 'Max Continuous', '100% speed', 'Completed', 'Medium', 1, 1, 'Margin check only', -25, -12, -13, 'Delivered, report attached'),
                  S('Emergency shutdown thermal shock', 'Shutdown', '0% speed', 'Pending', 'High', 3, 0, 'Analyst to be assigned', -1, 35)]),
        dict(Ortak=O(Proje='Sample Engine 2', ParcaAdi='Turbine Casing', Layout='Layout A', ModelNo='TCS-M05', ModelTipi='3D', AnalizTipi='Elastic', Alpha='Off', Malzeme='Hastelloy X', AnalizTuru='Static Structural', Aero=G_, AeroAdi='TCS_Pressure_R4', SAS=N_, SASAdi='', Termal=G_, TermalAdi='TCS_Therm_R1'),
             Ana=S('Casing containment stiffness', 'Take-off', '105% speed', 'On Hold', 'Low', 2, 1, 'On hold until design freeze', -30, 40),
             Alt=[S('Pressure + thermal combined', 'Transient', '100% speed', 'On Hold', 'Low', 2, 1, 'Linked to master hold', -22, 45)]),
        dict(Ortak=O(Proje='Sample Engine 3', ParcaAdi='Bolted Flange Joint', Layout='Layout A', ModelNo='BFJ-M06', ModelTipi='3D', AnalizTipi='Elastoplastic', Alpha='Off', Malzeme='IN718; A286', AnalizTuru='Contact / Bolt', Aero=N_, AeroAdi='', SAS=N_, SASAdi='', Termal=G_, TermalAdi='Flange_Therm_R1'),
             Ana=S('Bolt preload and gapping', 'Assembly', '0% speed', 'In Progress', 'Medium', 3, 1, 'Preload 70% of yield', -8, 14, None, 'Contact convergence OK'),
             Alt=[S('Operating load case', 'Take-off', '105% speed', 'Pending', 'Medium', 3, 2, 'Add axial load from engine mount', -2, 22),
                  S('Bolt relaxation after hold', 'Creep hold', '100% speed', 'Pending', 'Low', 1, 2, 'Creep data from material group', -1, 33)]),
        dict(Ortak=O(Proje='Sample Engine 3', ParcaAdi='Blisk Stage 1', Layout='Layout B', ModelNo='BLK-M07', ModelTipi='3D', AnalizTipi='Elastic', Alpha='On', Malzeme='Ti-6Al-4V', AnalizTuru='Prestressed Modal', Aero=G_, AeroAdi='Blisk_Aero_R2', SAS=G_, SASAdi='Blisk_SAS_R1', Termal=B_, TermalAdi='Blisk_Therm_Pending'),
             Ana=S('Campbell and SAFE diagram modes', 'Cruise', '100% speed', 'Model Preparation', 'High', 1, 2, 'Need modes up to 20 kHz', -5, 10),
             Alt=[S('Mistuning sensitivity', 'Cruise', '100% speed', 'Pending', 'Medium', 1, 2, 'Random mistuning 3 sigma', -1, 24),
                  S('Idle prestress modes', 'Idle', '60% speed', 'Pending', 'Low', 2, 2, 'Idle speed prestress only', -1, 27)]),
        dict(Ortak=O(Proje='Sample Engine 3', ParcaAdi='Compressor Shaft', Layout='Layout A', ModelNo='SHF-M08', ModelTipi='2D', AnalizTipi='Elastic', Alpha='Off', Malzeme='Waspaloy', AnalizTuru='Static Structural', Aero=N_, AeroAdi='', SAS=N_, SASAdi='', Termal=N_, TermalAdi=''),
             Ana=S('Torque and bending check', 'Max Torque', '100% speed', 'Completed', 'Medium', 2, 1, 'Axisymmetric model', -40, -20, -22, 'Delivered, no issues'),
             Alt=[S('Hard landing bending', 'Hard landing', '0% speed', 'Completed', 'High', 2, 2, 'Extra bending moment case', -30, -15, -16, 'Delivered')]),
    ]
    for m in modeller:
        a1 = yeni_ornek(m['Ortak'], m['Ana'])
        a1['ModelRolu'] = 'Master Model'
        ana_id = save_kayit(a1, None)
        for a in m['Alt']:
            alt = yeni_ornek(m['Ortak'], a)
            alt['ModelRolu'] = 'Sub-Analysis'
            alt['AnaModelID'] = ana_id
            save_kayit(alt, None)
    bagimsiz = [
        dict(Ortak=O(Proje='Sample Engine', ParcaAdi='Fan Blade', Layout='Layout A', ModelNo='FAN-S01', ModelTipi='3D', AnalizTipi='Elastic', Alpha='Off', Malzeme='Ti-6Al-4V', AnalizTuru='Modal', Aero=G_, AeroAdi='Fan_Aero_R1', SAS=N_, SASAdi='', Termal=N_, TermalAdi=''),
             A=S('Free-free modal check', 'Idle', '0% speed', 'Completed', 'Low', 1, 2, 'Test correlation request', -35, -20, -21, 'Correlation within 3%')),
        dict(Ortak=O(Proje='Sample Engine 2', ParcaAdi='Bearing Housing', Layout='Layout B', ModelNo='BRG-S02', ModelTipi='3D', AnalizTipi='Elastic', Alpha='Off', Malzeme='IN718', AnalizTuru='Static Structural', Aero=N_, AeroAdi='', SAS=N_, SASAdi='', Termal=G_, TermalAdi='BRG_Therm_R1'),
             A=S('Oil seal region stress', 'Take-off', '105% speed', 'In Progress', 'Medium', 3, 1, 'Check local yielding', -7, 8)),
        dict(Ortak=O(Proje='Sample Engine 3', ParcaAdi='Heat Shield', Layout='Layout A', ModelNo='HSH-S03', ModelTipi='2D', AnalizTipi='Elastic', Alpha='On', Malzeme='Hastelloy X', AnalizTuru='Thermal-Structural', Aero=N_, AeroAdi='', SAS=G_, SASAdi='HSH_SAS_R2', Termal=G_, TermalAdi='HSH_Therm_R2'),
             A=S('Thermal buckling check', 'Transient', '100% speed', 'Pending', 'High', 2, 0, 'Analyst to be assigned', 0, 18)),
        dict(Ortak=O(Proje='Sample Engine 3', ParcaAdi='Turbine Vane Segment', Layout='Layout C', ModelNo='VAN-S04', ModelTipi='3D', AnalizTipi='Elastoplastic', Alpha='Off', Malzeme='Hastelloy X', AnalizTuru='Fatigue (LCF/HCF)', Aero=G_, AeroAdi='Vane_Aero_R3', SAS=G_, SASAdi='Vane_SAS_R2', Termal=G_, TermalAdi='Vane_Therm_R2'),
             A=S('Trailing edge LCF', 'Take-off', '103% speed', 'Review / Check', 'High', 1, 2, 'Use new cooling scheme', -16, -2, None, 'Overdue: waiting for temperature review')),
    ]
    for m in bagimsiz:
        r = yeni_ornek(m['Ortak'], m['A'])
        r['ModelRolu'] = 'Standalone'
        save_kayit(r, None)
# -*- coding: utf-8 -*-


# ============================================================================
# PART 2 - USER INTERFACE: infrastructure, main window, refresh functions
# ============================================================================
import traceback
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
clr.AddReference("System.Data")

from System import String, DateTime, Int32, DBNull, AppDomain
from System.Reflection import BindingFlags
from System.Diagnostics import Process
from System.Data import DataTable
from System.Drawing import (Color, Font, FontStyle, Point, Size, ContentAlignment, SystemIcons)
from System.Windows.Forms import (
    Application, Form, Label, Button, TextBox, ComboBox, CheckBox, ListBox, Panel,
    FlowLayoutPanel, TableLayoutPanel, ColumnStyle, RowStyle, SizeType, TabControl, TabPage,
    DataGridView, DataGridViewTextBoxColumn, DataGridViewSelectionMode, DataGridViewAutoSizeColumnsMode,
    DataGridViewHitTestType, ListView, ListViewItem, View, ToolStrip, ToolStripSeparator,
    ToolStripDropDownButton, ToolStripStatusLabel, ToolStripGripStyle, StatusStrip,
    MessageBox, MessageBoxButtons, MessageBoxIcon, DialogResult, DockStyle, AnchorStyles, Padding,
    FlowDirection, BorderStyle, FlatStyle, ScrollBars, Keys, Timer, NotifyIcon, ToolTipIcon,
    SaveFileDialog, FolderBrowserDialog, FormBorderStyle, FormStartPosition, Screen, ComboBoxStyle,
    ColumnHeaderStyle)

# ---------------------------------------------------------------- settings you may edit
# Shared data folder. Leave empty to use  <program folder>\Veri  (or the folder chosen with
# Admin > Change data folder ...).  Example:  VERI_KLASORU = r"\\server\share\AnalizTakip\Veri"
VERI_KLASORU = r""
# Only for testing with several users on one PC: act as another Windows account.
TEST_HESAP = ""

FORM_KEY = "AnalizTakip.AnaForm"
_NONE_HIT = getattr(DataGridViewHitTestType, "None")


class U(object):
    """UI state"""
    ana = None
    Tum = []
    Imza = ''
    DlgAcik = False
    Kapaniyor = False
    Gosterilen = {}
    Tray = None
    FiltreKilit = False
    KopyaID = None
    KopyaKaynak = None
    Kisiler = []
    HataLog = ''
    Fontlar = {}
    Zamanlayici = None
    D = None


AY_ADLARI = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']


# ---------------------------------------------------------------- helpers
def exc_msg(e):
    try:
        m = e.Message
        if m is not None:
            return unicode(m)
    except Exception:
        pass
    try:
        return unicode(e)
    except Exception:
        return repr(e)


def write_hata_log(mesaj):
    try:
        File.AppendAllText(U.HataLog, u'%s | %s | %s\r\n' % (now_s(), Environment.UserName, mesaj), UTF8Encoding(True))
    except Exception:
        pass


def mesaj(metin, simge='Information', dugme='OK'):
    """message box; returns 'OK' / 'Yes' / 'No' ..."""
    ic = getattr(MessageBoxIcon, simge)
    bt = getattr(MessageBoxButtons, dugme)
    try:
        if U.ana is not None and not U.ana.IsDisposed and U.ana.Visible:
            return str(MessageBox.Show(U.ana, metin, 'Analysis Tracker', bt, ic))
    except Exception:
        pass
    return str(MessageBox.Show(metin, 'Analysis Tracker', bt, ic))


def guvenli(fn, *args):
    """run fn; show and log errors instead of crashing"""
    try:
        return fn(*args)
    except Exception as e:
        try:
            write_hata_log(traceback.format_exc() + u' | ' + exc_msg(e))
        except Exception:
            write_hata_log(exc_msg(e))
        mesaj(exc_msg(e), 'Error')
        return None


def olay(fn, *args):
    """event handler that runs fn(*args) safely"""
    def h(s, e):
        guvenli(fn, *args)
    return h


def cift_tampon(c):
    """double buffering: stops text flicker while scrolling horizontally"""
    try:
        p = c.GetType().GetProperty('DoubleBuffered', BindingFlags.Instance | BindingFlags.NonPublic)
        if p is not None:
            p.SetValue(c, True, None)
    except Exception:
        pass


def sec_temizle(g):
    """remove the blue selection (click on empty area)"""
    try:
        g.ClearSelection()
        g.CurrentCell = None
    except Exception:
        pass


def renk(r, g, b):
    return Color.FromArgb(r, g, b)


def boyut(w, h):
    return Size(w, h)


def font(punto=9.0, kalin=False):
    anahtar = (float(punto), bool(kalin))
    f = U.Fontlar.get(anahtar)
    if f is None:
        f = Font('Segoe UI', float(punto), FontStyle.Bold if kalin else FontStyle.Regular)
        U.Fontlar[anahtar] = f
    return f


def pad(l, t, r, b):
    return Padding(l, t, r, b)


def tekerlek_durdur(c):
    """mouse wheel must not change a combo box value (click only)"""
    def h(s, e):
        try:
            e.Handled = True
        except Exception:
            pass
    c.MouseWheel += h


def yeni_etiket(metin, w):
    l = Label()
    l.Text = metin
    l.Width = w
    l.Height = 24
    l.TextAlign = ContentAlignment.MiddleLeft
    l.Margin = pad(8, 0, 0, 0)
    return l


def add_kolon(g, baslik, alan_adi, genislik, bicim=''):
    c = DataGridViewTextBoxColumn()
    c.HeaderText = baslik
    c.Name = alan_adi
    c.DataPropertyName = alan_adi
    c.Width = genislik
    if bicim:
        c.DefaultCellStyle.Format = bicim
    g.Columns.Add(c)


def tarih_g(d):
    return d.strftime('%d.%m.%Y') if d else ''


# ---------------------------------------------------------------- program folder / settings
def program_klasoru():
    p = ''
    try:
        p = os.path.dirname(os.path.abspath(__file__))
    except Exception:
        p = ''
    if not p or not Directory.Exists(p):
        try:
            p = os.path.dirname(os.path.abspath(sys.argv[0])) if sys.argv and sys.argv[0] else ''
        except Exception:
            p = ''
    if not p or not Directory.Exists(p):
        p = os.path.join(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), 'AnalizTakip')
        try:
            if not Directory.Exists(p):
                Directory.CreateDirectory(p)
        except Exception:
            pass
    return p


def _ayar_adaylari():
    adaylar = [os.path.join(G.Klasor, 'AnalizTakip.ayar.json')]
    ap = os.path.join(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), 'AnalizTakip', 'AnalizTakip.ayar.json')
    if ap not in adaylar:
        adaylar.append(ap)
    return adaylar


def get_veri_klasoru():
    v = os.path.join(G.Klasor, 'Veri')
    if VERI_KLASORU:
        v = VERI_KLASORU
    for yol in _ayar_adaylari():
        try:
            if File.Exists(yol):
                o = json.loads(read_text(yol))
                if o.get('VeriKlasoru'):
                    return unicode(o['VeriKlasoru'])
        except Exception:
            pass
    return v


def set_veri_klasoru(yol):
    metin = json.dumps({'VeriKlasoru': yol}, ensure_ascii=False, indent=2)
    son = None
    for ayar in _ayar_adaylari():
        try:
            d = os.path.dirname(ayar)
            if not Directory.Exists(d):
                Directory.CreateDirectory(d)
            File.WriteAllText(ayar, metin, UTF8Encoding(False))
            # remove a stale setting file elsewhere so it cannot override the new choice
            for diger in _ayar_adaylari():
                if diger != ayar and File.Exists(diger):
                    try:
                        File.Delete(diger)
                    except Exception:
                        pass
            return
        except Exception as e:
            son = e
    if son is not None:
        raise son


# ---------------------------------------------------------------- main window
def build_main():
    ana = Form()
    U.ana = ana
    ana.Text = 'Analysis Tracker'
    ana.StartPosition = FormStartPosition.CenterScreen
    ana.Size = boyut(1320, 790)
    ana.MinimumSize = boyut(1000, 600)
    ana.Font = font(9)
    cift_tampon(ana)

    tab = TabControl()
    U.tab = tab
    cift_tampon(tab)
    tab.Dock = DockStyle.Fill
    tp1 = TabPage('Analysis List')
    tp2 = TabPage('Master Model Tree')
    tp3 = TabPage('Summary')
    U.tp1, U.tp2, U.tp3 = tp1, tp2, tp3
    tab.TabPages.Add(tp1)
    tab.TabPages.Add(tp2)
    tab.TabPages.Add(tp3)
    for tpp in (tp1, tp2, tp3):
        cift_tampon(tpp)

    # ---- tab 1: list
    dt = DataTable()
    for c in ('ID', 'Proje', 'Parca', 'Layout', 'ModelNo', 'Tip', 'MalzTipi', 'Alpha', 'Hiz', 'Malzeme', 'Kosul',
              'AnalizTuru', 'Rol', 'AnaID', 'Durum', 'Oncelik', 'Yapan', 'Talep', 'Gecikme', 'Guncelleyen'):
        dt.Columns.Add(c, clr.GetClrType(String))
    dt.Columns.Add('Deadline', clr.GetClrType(DateTime))
    dt.Columns.Add('Kalan', clr.GetClrType(Int32))
    dt.Columns.Add('Alt', clr.GetClrType(Int32))
    dt.DefaultView.Sort = 'ID DESC'
    U.dt = dt

    grid = DataGridView()
    U.grid = grid
    cift_tampon(grid)
    grid.Dock = DockStyle.Fill
    grid.ReadOnly = True
    grid.AllowUserToAddRows = False
    grid.AllowUserToDeleteRows = False
    grid.AllowUserToResizeRows = False
    grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect
    grid.MultiSelect = False
    grid.AutoGenerateColumns = False
    grid.RowHeadersVisible = False
    grid.BackgroundColor = Color.White
    grid.ColumnHeadersDefaultCellStyle.Font = font(9, True)
    grid.EnableHeadersVisualStyles = False
    grid.ColumnHeadersDefaultCellStyle.BackColor = renk(31, 56, 100)
    grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
    grid.AlternatingRowsDefaultCellStyle.BackColor = renk(242, 246, 252)
    grid.ColumnHeadersHeight = 30
    for b, a, w, bic in (('ID', 'ID', 70, ''), ('Project', 'Proje', 110, ''), ('Part', 'Parca', 170, ''),
                         ('Layout', 'Layout', 80, ''), ('Model ID', 'ModelNo', 85, ''), ('2D/3D', 'Tip', 60, ''),
                         ('Material Model', 'MalzTipi', 140, ''), ('Alpha', 'Alpha', 120, ''),
                         ('Material', 'Malzeme', 110, ''), ('Condition', 'Kosul', 100, ''), ('Speed', 'Hiz', 90, ''),
                         ('Analysis Type', 'AnalizTuru', 130, ''), ('Role', 'Rol', 100, ''), ('Master', 'AnaID', 75, ''),
                         ('Status', 'Durum', 160, ''), ('Priority', 'Oncelik', 65, ''), ('Analyst', 'Yapan', 110, ''),
                         ('Requester', 'Talep', 100, ''), ('Due Date', 'Deadline', 85, 'dd.MM.yyyy'),
                         ('Days Left', 'Kalan', 65, ''), ('Overdue', 'Gecikme', 75, ''), ('Subs', 'Alt', 40, ''),
                         ('Last Updated By', 'Guncelleyen', 120, '')):
        add_kolon(grid, b, a, w, bic)
    grid.DataSource = dt

    fp = FlowLayoutPanel()
    U.fp = fp
    fp.Dock = DockStyle.Top
    fp.Height = 40
    fp.Padding = pad(6, 6, 6, 0)
    txtAra = TextBox()
    txtAra.Width = 200
    txtAra.Margin = pad(0, 2, 0, 0)
    cmbDurum = ComboBox()
    cmbDurum.DropDownStyle = ComboBoxStyle.DropDownList
    cmbDurum.Width = 150
    cmbDurum.Margin = pad(0, 2, 0, 0)
    cmbKisi = ComboBox()
    cmbKisi.DropDownStyle = ComboBoxStyle.DropDownList
    cmbKisi.Width = 140
    cmbKisi.Margin = pad(0, 2, 0, 0)
    tekerlek_durdur(cmbDurum)
    tekerlek_durdur(cmbKisi)
    chkAcik = CheckBox()
    chkAcik.Text = 'Open only'
    chkAcik.AutoSize = True
    chkAcik.Margin = pad(12, 3, 0, 0)
    chkGec = CheckBox()
    chkGec.Text = 'Overdue only'
    chkGec.AutoSize = True
    chkGec.Margin = pad(8, 3, 0, 0)
    btnTemizle = Button()
    btnTemizle.Text = 'Clear filter'
    btnTemizle.AutoSize = True
    btnTemizle.Margin = pad(12, 0, 0, 0)
    U.txtAra, U.cmbDurum, U.cmbKisi, U.chkAcik, U.chkGec, U.btnTemizle = txtAra, cmbDurum, cmbKisi, chkAcik, chkGec, btnTemizle
    fp.Controls.Add(yeni_etiket('Search:', 50))
    fp.Controls.Add(txtAra)
    fp.Controls.Add(yeni_etiket('Status:', 50))
    fp.Controls.Add(cmbDurum)
    fp.Controls.Add(yeni_etiket('Analyst:', 55))
    fp.Controls.Add(cmbKisi)
    fp.Controls.Add(chkAcik)
    fp.Controls.Add(chkGec)
    fp.Controls.Add(btnTemizle)

    tp1.Controls.Add(grid)
    tp1.Controls.Add(fp)
    pnlYeni = Panel()
    pnlYeni.Dock = DockStyle.Bottom
    pnlYeni.Height = 52
    tbYeni = Button()
    tbYeni.Text = '+  New Analysis'
    tbYeni.Width = 160
    tbYeni.Height = 36
    tbYeni.Location = Point(8, 8)
    tbYeni.FlatStyle = FlatStyle.Flat
    tbYeni.BackColor = renk(84, 130, 53)
    tbYeni.ForeColor = Color.White
    tbYeni.Font = font(10, True)
    pnlYeni.Controls.Add(tbYeni)
    tp1.Controls.Add(pnlYeni)
    U.tbYeni = tbYeni

    # ---- tab 2: master model tree
    lv = ListView()
    U.lv = lv
    cift_tampon(lv)
    lv.Dock = DockStyle.Fill
    lv.View = View.Details
    lv.FullRowSelect = True
    lv.HideSelection = False
    lv.MultiSelect = False
    lv.ShowGroups = False
    lv.GridLines = False
    lv.BorderStyle = getattr(BorderStyle, 'None')
    lv.Font = font(9.5)
    for b, w in (('ID', 80), ('Role', 100), ('Part', 210), ('Condition / Purpose', 300), ('Status', 130),
                 ('Priority', 70), ('Analyst', 130), ('Requester', 130), ('Due Date', 90), ('Days Left', 70)):
        lv.Columns.Add(b, w)
    fpAgac = FlowLayoutPanel()
    U.fpAgac = fpAgac
    fpAgac.Dock = DockStyle.Top
    fpAgac.Height = 40
    fpAgac.Padding = pad(6, 6, 6, 0)
    txtAgac = TextBox()
    txtAgac.Width = 220
    txtAgac.Margin = pad(0, 2, 0, 0)
    chkAgac = CheckBox()
    chkAgac.Text = 'Open only'
    chkAgac.AutoSize = True
    chkAgac.Margin = pad(12, 3, 0, 0)
    lblAgac = Label()
    lblAgac.AutoSize = True
    lblAgac.TextAlign = ContentAlignment.MiddleLeft
    lblAgac.ForeColor = renk(100, 100, 100)
    lblAgac.Margin = pad(24, 4, 0, 0)
    lblAgac.Text = 'blue band: master model   |   red: overdue   |   gray: closed   |   double-click to open'
    fpAgac.Controls.Add(yeni_etiket('Search:', 50))
    fpAgac.Controls.Add(txtAgac)
    fpAgac.Controls.Add(chkAgac)
    fpAgac.Controls.Add(lblAgac)
    U.txtAgac, U.chkAgac = txtAgac, chkAgac
    tp2.Controls.Add(lv)
    tp2.Controls.Add(fpAgac)

    # ---- tab 3: summary (three tables)
    ozetTablo = TableLayoutPanel()
    U.ozetTablo = ozetTablo
    ozetTablo.Dock = DockStyle.Fill
    ozetTablo.ColumnCount = 3
    ozetTablo.RowCount = 1
    for i in range(3):
        ozetTablo.ColumnStyles.Add(ColumnStyle(SizeType.Percent, 33.3))

    def yeni_ozet_grid(basliklar):
        g = DataGridView()
        cift_tampon(g)
        g.Dock = DockStyle.Fill
        g.ReadOnly = True
        g.AllowUserToAddRows = False
        g.RowHeadersVisible = False
        g.BackgroundColor = Color.White
        g.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        g.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        g.EnableHeadersVisualStyles = False
        g.ColumnHeadersDefaultCellStyle.BackColor = renk(31, 56, 100)
        g.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        for b in basliklar:
            g.Columns.Add(b, b)
        return g
    gDurum = yeni_ozet_grid(['Status', 'Count'])
    gKisi = yeni_ozet_grid(['Analyst', 'Open', 'Overdue'])
    gYakin = yeni_ozet_grid(['Days Left', 'ID', 'Part', 'Analyst', 'Due Date'])
    U.gDurum, U.gKisi, U.gYakin = gDurum, gKisi, gYakin
    for g, baslik in ((gDurum, 'By status'), (gKisi, 'By analyst (open items)'), (gYakin, 'Next 10 due dates (open items)')):
        pn = Panel()
        pn.Dock = DockStyle.Fill
        pn.Padding = pad(4, 4, 4, 4)
        lb = Label()
        lb.Text = baslik
        lb.Dock = DockStyle.Top
        lb.Height = 24
        lb.Font = font(9, True)
        pn.Controls.Add(g)
        pn.Controls.Add(lb)
        ozetTablo.Controls.Add(pn)
    tp3.Controls.Add(ozetTablo)

    # ---- tool strip and status strip
    ts = ToolStrip()
    ts.GripStyle = ToolStripGripStyle.Hidden
    ts.Padding = pad(6, 4, 6, 4)
    U.tbDuzenle = ts.Items.Add('Edit')
    U.tbKopya = ts.Items.Add('Duplicate')
    U.tbSil = ts.Items.Add('Delete')
    ts.Items.Add(ToolStripSeparator())
    U.tbYenile = ts.Items.Add('Refresh')
    U.tbExcel = ts.Items.Add('Export to Excel')
    U.tbBildirim = ts.Items.Add('Notifications')
    ts.Items.Add(ToolStripSeparator())
    tbAyar = ToolStripDropDownButton('Settings')
    U.miProfil = tbAyar.DropDownItems.Add('My profile')
    U.miListe = tbAyar.DropDownItems.Add('Edit lists (material, condition, status ...)')
    U.miAc = tbAyar.DropDownItems.Add('Open data folder')
    U.miHak = tbAyar.DropDownItems.Add('About')
    ts.Items.Add(tbAyar)
    tbAdmin = ToolStripDropDownButton('Admin')
    U.tbAdmin = tbAdmin
    U.miKull = tbAdmin.DropDownItems.Add('Manage users ...')
    U.miSilinen = tbAdmin.DropDownItems.Add('Restore deleted records ...')
    U.miDevir = tbAdmin.DropDownItems.Add('Reassign analyst ...')
    U.miYedek = tbAdmin.DropDownItems.Add('Back up data now')
    U.miKlasor = tbAdmin.DropDownItems.Add('Change data folder ...')
    U.miOrnek = tbAdmin.DropDownItems.Add('Add sample data')
    ts.Items.Add(tbAdmin)
    for b in (U.tbDuzenle, U.tbKopya, U.tbSil, U.tbYenile, U.tbExcel, U.tbBildirim):
        b.Padding = pad(6, 2, 6, 2)

    ss = StatusStrip()
    U.sbKayit = ToolStripStatusLabel()
    U.sbKayit.Text = 'Records: 0'
    U.sbVeri = ToolStripStatusLabel()
    U.sbVeri.Spring = True
    U.sbVeri.TextAlign = ContentAlignment.MiddleLeft
    U.sbKul = ToolStripStatusLabel()
    U.sbZaman = ToolStripStatusLabel()
    ss.Items.Add(U.sbKayit)
    ss.Items.Add(U.sbVeri)
    ss.Items.Add(U.sbKul)
    ss.Items.Add(U.sbZaman)

    ana.Controls.Add(tab)
    ana.Controls.Add(ts)
    ana.Controls.Add(ss)
    return ana


# ---------------------------------------------------------------- refresh
def get_secili_id():
    if U.tab.SelectedTab == U.tp2:
        if U.lv.SelectedItems.Count > 0 and U.lv.SelectedItems[0].Tag:
            return unicode(U.lv.SelectedItems[0].Tag)
        return None
    if U.tab.SelectedTab == U.tp1 and U.grid.CurrentRow is not None and U.grid.CurrentRow.DataBoundItem is not None:
        return unicode(U.grid.CurrentRow.DataBoundItem['ID'])
    return None


def yenile_liste():
    U.FiltreKilit = True
    try:
        grid, dt = U.grid, U.dt
        secili = None
        if grid.CurrentRow is not None and grid.CurrentRow.DataBoundItem is not None:
            secili = unicode(grid.CurrentRow.DataBoundItem['ID'])
        ara = (U.txtAra.Text or '').strip().lower()
        fd = U.cmbDurum.SelectedItem
        fk = U.cmbKisi.SelectedItem
        fd = unicode(fd) if fd is not None else ''
        fk = unicode(fk) if fk is not None else ''
        dt.BeginLoadData()
        dt.Rows.Clear()
        for r in U.Tum:
            if fd and fd != 'All' and r['Durum'] != fd:
                continue
            if fk and fk != 'All' and r['AnalizYapan'] != fk:
                continue
            if U.chkAcik.Checked and test_kapali(r):
                continue
            gec = test_gecikti(r)
            if U.chkGec.Checked and not gec:
                continue
            if ara:
                hay = u' '.join([r['ID'], r['ParcaAdi'], r['Proje'], r['Layout'], r['Malzeme'], r['Amac'], r['Kosul'],
                                 r['AnalizTuru'], r['AnalizYapan'], r['TalepEden'], r['ModelNo'], r['TalepNotu'],
                                 r['YapanNotu']]).lower()
                if ara not in hay:
                    continue
            row = dt.NewRow()
            row['ID'] = r['ID']
            row['Parca'] = r['ParcaAdi']
            row['Proje'] = r['Proje']
            row['Layout'] = r['Layout']
            row['ModelNo'] = r['ModelNo']
            row['Alpha'] = r['Alpha']
            row['Hiz'] = r['Hiz']
            row['MalzTipi'] = r['AnalizTipi']
            row['Tip'] = r['ModelTipi']
            row['Malzeme'] = r['Malzeme']
            row['Kosul'] = r['Kosul']
            row['AnalizTuru'] = r['AnalizTuru']
            row['Rol'] = r['ModelRolu']
            row['AnaID'] = r['AnaModelID']
            row['Durum'] = r['Durum']
            row['Oncelik'] = r['Oncelik']
            row['Yapan'] = r['AnalizYapan']
            row['Talep'] = r['TalepEden']
            d = convert_to_tarih(r['Deadline'])
            if d is not None:
                row['Deadline'] = DateTime(d.year, d.month, d.day)
            else:
                row['Deadline'] = DBNull.Value
            k = get_kalan_gun(r)
            if k is not None:
                row['Kalan'] = int(k)
            else:
                row['Kalan'] = DBNull.Value
            row['Gecikme'] = 'OVERDUE' if gec else ''
            if r['ModelRolu'] == 'Master Model':
                row['Alt'] = len([x for x in U.Tum if x['ModelRolu'] == 'Sub-Analysis' and x['AnaModelID'] == r['ID']])
            else:
                row['Alt'] = DBNull.Value
            row['Guncelleyen'] = r['SonGuncelleyen']
            dt.Rows.Add(row)
        dt.EndLoadData()
        if secili:
            for gr in grid.Rows:
                if gr.DataBoundItem is not None and unicode(gr.DataBoundItem['ID']) == secili:
                    grid.CurrentCell = gr.Cells[1]
                    break
    finally:
        U.FiltreKilit = False


def get_durum_rengi(d):
    if d == 'Pending':
        return renk(105, 105, 105)
    if d == 'Model Preparation':
        return renk(191, 120, 0)
    if d == 'In Progress':
        return renk(30, 100, 200)
    if d == 'Review / Check':
        return renk(120, 60, 160)
    if d == 'On Hold':
        return renk(160, 90, 40)
    if d == 'Completed':
        return renk(70, 130, 50)
    return renk(128, 128, 128)


def yeni_agac_satiri(r, rol, ana):
    """one row: master model (blue band, bold) or sub / standalone analysis"""
    dl = convert_to_tarih(r['Deadline'])
    kal = get_kalan_gun(r)
    amac = u' - '.join([x for x in (r['Kosul'], r['Amac']) if x])
    parca = r['ParcaAdi']
    if (not ana) and rol == 'Sub-Analysis':
        parca = unichr(0x2514) + u'  ' + parca
    it = ListViewItem(r['ID'])
    it.Tag = r['ID']
    it.UseItemStyleForSubItems = False
    for t in ('Master' if ana else rol, parca, amac, r['Durum'], r['Oncelik'], r['AnalizYapan'], r['TalepEden'],
              tarih_g(dl), (unicode(kal) if kal is not None else '')):
        it.SubItems.Add(t)
    kapali = test_kapali(r)
    gec = test_gecikti(r)
    kalin = font(9.5, True)
    normal = font(9.5)
    zemin = renk(226, 234, 248) if ana else Color.White
    if kapali:
        yazi = renk(150, 150, 150)
    elif ana:
        yazi = renk(31, 56, 140)
    else:
        yazi = renk(40, 40, 40)
    i = 0
    for si in it.SubItems:
        si.BackColor = zemin
        si.ForeColor = yazi
        si.Font = kalin if ana else normal
        if i == 4 and not kapali:
            si.ForeColor = get_durum_rengi(r['Durum'])
            si.Font = kalin
        if i == 4 and r['Durum'] == 'Completed':
            si.ForeColor = get_durum_rengi('Completed')
            si.Font = kalin
        if i == 5 and r['Oncelik'] == 'High' and not kapali:
            si.ForeColor = renk(192, 0, 0)
            si.Font = kalin
        if (i == 9 or i == 8) and gec:
            si.ForeColor = renk(192, 0, 0)
            si.Font = kalin
        i += 1
    return it


def yenile_agac():
    lv = U.lv
    secili = None
    try:
        if lv.SelectedItems.Count > 0:
            secili = unicode(lv.SelectedItems[0].Tag)
    except Exception:
        pass
    ara = (U.txtAgac.Text or '').strip().lower()
    sadece_acik = bool(U.chkAgac.Checked)

    def uyar(r):
        if not ara:
            return True
        return ara in u' '.join([r['ID'], r['ParcaAdi'], r['Proje'], r['Layout'], r['Malzeme'], r['Amac'], r['Kosul'],
                                 r['AnalizYapan'], r['TalepEden'], r['ModelNo']]).lower()
    lv.BeginUpdate()
    try:
        lv.Items.Clear()
        kullanilan = {}
        for m in [x for x in U.Tum if x['ModelRolu'] == 'Master Model']:
            kullanilan[m['ID']] = True
            alt = sorted([x for x in U.Tum if x['ModelRolu'] == 'Sub-Analysis' and x['AnaModelID'] == m['ID']], key=lambda x: x['ID'])
            for c in alt:
                kullanilan[c['ID']] = True
            gorunen_alt = [x for x in alt if (not sadece_acik or not test_kapali(x)) and (uyar(x) or uyar(m))]
            ana_uyar = uyar(m)
            if ara and not ana_uyar and len(gorunen_alt) == 0:
                continue
            if sadece_acik and test_kapali(m) and len(gorunen_alt) == 0:
                continue
            lv.Items.Add(yeni_agac_satiri(m, 'Master Model', True))
            for c in gorunen_alt:
                lv.Items.Add(yeni_agac_satiri(c, 'Sub-Analysis', False))
        yetim = [x for x in U.Tum if x['ModelRolu'] == 'Sub-Analysis' and x['ID'] not in kullanilan
                 and (not sadece_acik or not test_kapali(x)) and uyar(x)]
        for c in yetim:
            lv.Items.Add(yeni_agac_satiri(c, 'Sub-Analysis', False))
        bag = [x for x in U.Tum if x['ModelRolu'] != 'Master Model' and x['ModelRolu'] != 'Sub-Analysis'
               and (not sadece_acik or not test_kapali(x)) and uyar(x)]
        for c in bag:
            lv.Items.Add(yeni_agac_satiri(c, 'Standalone', False))
        if secili:
            for x in lv.Items:
                if unicode(x.Tag) == secili:
                    x.Selected = True
                    break
    finally:
        lv.EndUpdate()


def yenile_ozet():
    o = get_ozet(U.Tum)
    U.gDurum.Rows.Clear()
    for x in o.Durumlar:
        U.gDurum.Rows.Add(x.Ad, x.Adet)
    U.gKisi.Rows.Clear()
    for x in o.Kisiler:
        U.gKisi.Rows.Add(x.Ad, x.Acik, x.Geciken)
    U.gYakin.Rows.Clear()
    for x in o.Yaklasan:
        dl = convert_to_tarih(x['Deadline'])
        U.gYakin.Rows.Add(get_kalan_gun(x), x['ID'], x['ParcaAdi'], x['AnalizYapan'], tarih_g(dl))
    for gr in U.gYakin.Rows:
        kv = gr.Cells[0].Value
        if kv is not None and int(kv) < 0:
            gr.DefaultCellStyle.ForeColor = renk(192, 0, 0)
        elif kv is not None and int(kv) <= 7:
            gr.DefaultCellStyle.ForeColor = renk(191, 120, 0)
    for gg in (U.gDurum, U.gKisi, U.gYakin):
        sec_temizle(gg)


def yenile_filtreler():
    fd = U.cmbDurum.SelectedItem
    fk = U.cmbKisi.SelectedItem
    fd = unicode(fd) if fd is not None else ''
    fk = unicode(fk) if fk is not None else ''
    U.FiltreKilit = True
    try:
        l = get_listeler()
        U.cmbDurum.Items.Clear()
        U.cmbDurum.Items.Add('All')
        for x in l['Durum']:
            U.cmbDurum.Items.Add(x)
        U.cmbKisi.Items.Clear()
        U.cmbKisi.Items.Add('All')
        for x in sorted(set([r['AnalizYapan'] for r in U.Tum if r['AnalizYapan']])):
            U.cmbKisi.Items.Add(x)
        U.cmbDurum.SelectedItem = fd if (fd and U.cmbDurum.Items.Contains(fd)) else 'All'
        U.cmbKisi.SelectedItem = fk if (fk and U.cmbKisi.Items.Contains(fk)) else 'All'
    finally:
        U.FiltreKilit = False


def yenile_hepsi():
    U.Tum = list(get_tum_kayitlar())
    U.Imza = get_veri_imzasi()
    yenile_filtreler()
    yenile_liste()
    yenile_agac()
    yenile_ozet()
    U.sbKayit.Text = 'Records: %d' % len(U.Tum)
    U.sbVeri.Text = 'Data: ' + G.Kok
    U.Kisiler = list(get_tum_kullanicilar())
    if G.Ben is not None:
        U.sbKul.Text = 'User: %s (%s)' % (G.Ben['Ad'], G.Ben['Rol'])
    U.sbZaman.Text = 'Last refresh: ' + datetime.now().strftime('%H:%M:%S')


def yetki_uygula():
    k = G.Ben
    U.miListe.Visible = test_tam_yetki(k)
    U.tbAdmin.Visible = test_admin(k)
    U.tbBildirim.Visible = (k['Rol'] != 'Leader')
    U.sbKul.Text = 'User: %s (%s)' % (k['Ad'], k['Rol'])
    U.ana.Text = 'Analysis Tracker  -  %s  (%s)' % (k['Ad'], k['Rol'])


def yenile_ben():
    """the administrator may have changed role/name; if the registration is gone the program closes.
    Temporary network errors are ignored."""
    try:
        k = get_kullanici(get_hesap_adi())
    except Exception:
        return True
    if k is None:
        if not U.Kapaniyor:
            U.Kapaniyor = True
            mesaj('Your user registration was removed by the administrator. The program will now close.', 'Warning')
            try:
                U.ana.Close()
            except Exception:
                pass
        return False
    degisti = (G.Ben is None or k['Rol'] != G.Ben['Rol'] or k['Ad'] != G.Ben['Ad'] or test_admin(k) != test_admin(G.Ben))
    G.Ben = k
    if degisti:
        yetki_uygula()
    return True


def yenile_bildirim():
    if G.Ben is None or G.Ben['Rol'] == 'Leader':
        return
    liste = get_benim_bildirimler(G.Ben)
    oku = [b for b in liste if b['Okunmadi']]
    U.tbBildirim.Text = ('Notifications (%d)' % len(oku)) if len(oku) > 0 else 'Notifications'
    U.tbBildirim.ForeColor = renk(192, 0, 0) if len(oku) > 0 else Color.Black
    yeni = [b for b in oku if b['Dosya'] not in U.Gosterilen]
    for b in oku:
        U.Gosterilen[b['Dosya']] = True
    if len(yeni) > 0 and U.Tray is not None:
        if len(yeni) == 1:
            metin = yeni[0]['Mesaj']
        else:
            metin = u'%d new notifications (latest: %s)' % (len(yeni), yeni[0]['Mesaj'])
        try:
            U.Tray.ShowBalloonTip(8000, 'Analysis Tracker', metin, ToolTipIcon.Info)
        except Exception:
            pass


def otomatik_yenile():
    if not yenile_ben():
        return
    yenile_bildirim()
    if U.DlgAcik:
        return
    if get_veri_imzasi() != U.Imza:
        yenile_hepsi()
# -*- coding: utf-8 -*-


# ============================================================================
# PART 3 - USER INTERFACE: record form (modal)
# ============================================================================
def get_secenekler(a, listeler, tum):
    if a['K'] == 'TalepEden':
        return [k['Ad'] for k in get_tum_kullanicilar()]
    if a['K'] == 'AnalizYapan':
        return [''] + [k['Ad'] for k in get_tum_kullanicilar() if k['Rol'] == 'Structural Analysis Engineer']
    liste = []
    if a.get('List') in listeler:
        liste = list(listeler[a['List']])
    if not a.get('Strict'):
        for r in tum:
            e = r.get(a['K'])
            if e and e not in liste:
                liste.append(e)
    return liste


def new_tarih_kontrol(deger, zorunlu, sabit):
    """English day / month / year selection (the Windows calendar popup uses the system language)"""
    p = FlowLayoutPanel()
    p.FlowDirection = FlowDirection.LeftToRight
    p.WrapContents = False
    p.Width = 330
    p.Height = 30
    cb = CheckBox()
    cb.AutoSize = True
    cb.Text = ''
    cb.Width = 20
    cb.Margin = pad(0, 6, 4, 0)
    g = ComboBox()
    g.DropDownStyle = ComboBoxStyle.DropDownList
    g.Width = 52
    a = ComboBox()
    a.DropDownStyle = ComboBoxStyle.DropDownList
    a.Width = 64
    y = ComboBox()
    y.DropDownStyle = ComboBoxStyle.DropDownList
    y.Width = 70
    for x in (g, a, y):
        tekerlek_durdur(x)
        x.Margin = pad(0, 3, 6, 0)
    for i in range(1, 32):
        g.Items.Add(str(i))
    for m in AY_ADLARI:
        a.Items.Add(m)
    y0 = datetime.now().year
    d0 = deger.year if deger is not None else y0
    for i in range(min(y0, d0) - 1, max(y0, d0) + 7):
        y.Items.Add(str(i))
    t = {'Chk': cb, 'Gun': g, 'Ay': a, 'Yil': y}
    p.Tag = t
    if deger is not None:
        g.SelectedItem = str(deger.day)
        a.SelectedItem = AY_ADLARI[deger.month - 1]
        y.SelectedItem = str(deger.year)
        cb.Checked = True
    else:
        cb.Checked = False
    if not zorunlu:
        def degisti(s, e):
            try:
                for k in ('Gun', 'Ay', 'Yil'):
                    t[k].Enabled = s.Checked
            except Exception:
                pass
        cb.CheckedChanged += degisti
    for k in ('Gun', 'Ay', 'Yil'):
        t[k].Enabled = (cb.Checked and not sabit)
    if zorunlu:
        cb.Checked = True
    else:
        p.Controls.Add(cb)
    if sabit:
        cb.Enabled = False
    for x in (g, a, y):
        p.Controls.Add(x)
    return p


def get_tarih_deger(p):
    t = p.Tag
    if not t['Chk'].Checked:
        return None
    if t['Gun'].SelectedItem is None or t['Ay'].SelectedItem is None or t['Yil'].SelectedItem is None:
        return None
    ay = AY_ADLARI.index(unicode(t['Ay'].SelectedItem)) + 1
    yil = int(unicode(t['Yil'].SelectedItem))
    import calendar
    gun = min(int(unicode(t['Gun'].SelectedItem)), calendar.monthrange(yil, ay)[1])
    return datetime(yil, ay, gun)


def _split(deger, ayirac):
    return [x for x in re.split(r'\s*' + ayirac + r'\s*', deger or '') if x]


class FormDurum(object):
    """state of one open record form"""
    pass


def set_alan_degeri(d, a, c, deger):
    t = a['T']
    if t == 'multi':
        secili = _split(deger, ',')
        for cb in d.Chk[a['K']]:
            cb.Checked = (cb.Text in secili)
    elif t == 'multitext':
        lb = d.Lst[a['K']]
        lb.Items.Clear()
        for s in _split(deger, ';'):
            lb.Items.Add(s)
        c.Text = ''
    elif t == 'combo':
        if a.get('Strict'):
            if deger:
                if not c.Items.Contains(deger):
                    c.Items.Add(deger)
                c.SelectedItem = deger
            else:
                c.SelectedItem = None
        else:
            c.Text = deger
    elif t == 'text':
        c.Text = deger


def get_alan_degeri(d, a, c):
    t = a['T']
    if t == 'multi':
        return u', '.join([cb.Text for cb in d.Chk[a['K']] if cb.Checked])
    if t == 'date':
        return convert_from_tarih(get_tarih_deger(c))
    if t == 'multitext':
        liste = [unicode(i) for i in d.Lst[a['K']].Items]
        x = (c.Text or '').strip()
        if x and x not in liste:
            liste.append(x)
        return u'; '.join(liste)
    if t == 'master':
        x = c.Text
        if x:
            return unicode(x).split(' | ')[0]
        return ''
    return (c.Text or '').strip()


def show_kayit_formu(rec, yeni, mod='Full'):
    """opens the record form. Returns True if saved."""
    listeler = get_listeler()
    tum = list(get_tum_kayitlar())
    eski = None
    if not yeni:
        eski = convert_to_kayit(rec)

    dlg = Form()
    if yeni:
        dlg.Text = 'New Analysis'
    elif mod == 'Gor':
        dlg.Text = 'Analysis - ' + rec['ID'] + '  (read only)'
    else:
        dlg.Text = 'Analysis - ' + rec['ID']
    dlg.StartPosition = FormStartPosition.CenterParent
    yuk = 880
    try:
        yuk = min(880, max(500, Screen.PrimaryScreen.WorkingArea.Height - 40))
    except Exception:
        pass
    dlg.Size = boyut(880, yuk)
    dlg.MinimumSize = boyut(700, 520)
    dlg.Font = font(9)
    dlg.MinimizeBox = False
    dlg.MaximizeBox = True
    dlg.ShowInTaskbar = False

    # bottom button bar
    alt = Panel()
    alt.Dock = DockStyle.Bottom
    alt.Height = 58
    alt.BackColor = renk(240, 243, 249)
    btnK = Button()
    btnK.Text = 'Save'
    btnK.Width = 120
    btnK.Height = 32
    btnK.BackColor = renk(31, 56, 100)
    btnK.ForeColor = Color.White
    btnK.FlatStyle = FlatStyle.Flat
    btnK.Font = font(9.5, True)
    btnI = Button()
    btnI.Text = 'Cancel'
    btnI.Width = 100
    btnI.Height = 32
    btnI.FlatStyle = FlatStyle.Flat
    btnI.BackColor = Color.White
    akis = FlowLayoutPanel()
    akis.Dock = DockStyle.Fill
    akis.FlowDirection = FlowDirection.RightToLeft
    akis.Padding = pad(10, 10, 14, 0)
    btnD = Button()
    btnD.Text = 'Save + Duplicate'
    btnD.Width = 140
    btnD.Height = 32
    btnD.FlatStyle = FlatStyle.Flat
    btnD.BackColor = Color.White
    akis.Controls.Add(btnI)
    akis.Controls.Add(btnK)
    akis.Controls.Add(btnD)
    alt.Controls.Add(akis)
    ipucuL = Label()
    ipucuL.Dock = DockStyle.Left
    ipucuL.Width = 330
    ipucuL.TextAlign = ContentAlignment.MiddleLeft
    ipucuL.ForeColor = renk(110, 110, 110)
    ipucuL.Text = '  Read only' if mod == 'Gor' else '  * Required fields'
    alt.Controls.Add(ipucuL)

    # banner
    ban = Panel()
    ban.Dock = DockStyle.Top
    ban.Height = 66
    ban.BackColor = renk(31, 56, 100)
    bt1 = Label()
    bt1.Location = Point(20, 9)
    bt1.Size = boyut(800, 28)
    bt1.Font = font(14, True)
    bt1.ForeColor = Color.White
    bt2 = Label()
    bt2.Location = Point(22, 38)
    bt2.Size = boyut(800, 20)
    bt2.ForeColor = renk(196, 210, 232)
    if yeni:
        bt1.Text = 'Duplicate Analysis Request' if U.KopyaKaynak else 'New Analysis Request'
        if U.KopyaKaynak:
            bt2.Text = 'Copy of ' + U.KopyaKaynak + '  |  Change the option that differs (Material Model, Alpha, Model Type ...) and save.'
        else:
            bt2.Text = 'Requested by ' + G.Ben['Ad'] + '  |  Fill in the fields below. Status, analyst notes and file paths are completed later by the analyst.'
    else:
        bt1.Text = rec['ID'] + '   ' + rec['ParcaAdi']
        bt2.Text = 'Status: ' + rec['Durum'] + '  |  Requester: ' + rec['TalepEden'] + (('  |  Last updated by ' + rec['SonGuncelleyen']) if rec['SonGuncelleyen'] else '')
    ban.Controls.Add(bt1)
    ban.Controls.Add(bt2)

    pnl = Panel()
    pnl.Dock = DockStyle.Fill
    pnl.AutoScroll = True
    pnl.BackColor = Color.White
    tbl = TableLayoutPanel()
    tbl.Dock = DockStyle.Top
    tbl.ColumnCount = 2
    tbl.RowCount = len(ALANLAR) + 10
    tbl.Padding = pad(12, 6, 22, 12)
    tbl.ColumnStyles.Add(ColumnStyle(SizeType.Absolute, 190))
    tbl.ColumnStyles.Add(ColumnStyle(SizeType.Percent, 100))

    d = FormDurum()
    d.Ctl = {}
    d.Chk = {}
    d.Lst = {}
    d.Combos = []
    d.Rec = rec
    d.Eski = eski
    d.Yeni = yeni
    d.Mod = mod
    d.Kaydedildi = False
    d.KopyaIste = False
    d.SonID = ''
    d.Tum = tum
    d.Dlg = dlg
    d.AnaAlan = alan('AnaModelID')
    U.D = d
    ctl, chk, lst, combos = d.Ctl, d.Chk, d.Lst, d.Combos
    satir_yer = {}

    def liste_ekle(k):
        kutu = ctl[k]
        lb = lst[k]
        t = (kutu.Text or '').strip()
        if t:
            var = False
            for i in lb.Items:
                if unicode(i) == t:
                    var = True
            if not var:
                lb.Items.Add(t)
        kutu.Text = ''

    def liste_cikar(k):
        lb = lst[k]
        if lb.SelectedIndex >= 0:
            lb.Items.RemoveAt(lb.SelectedIndex)

    def secim_temizle_olayi(s, e):
        try:
            if not s.Focused:
                s.SelectionLength = 0
        except Exception:
            pass

    satir = 0
    toplam_h = 16
    for a in ALANLAR:
        if a['T'] == 'ro' or a.get('Bagli'):
            continue
        if a.get('SadeceDuzenle') and yeni:
            continue
        if a.get('Bolum'):
            hl = Label()
            hl.Text = a['Bolum'].upper()
            hl.Dock = DockStyle.Fill
            hl.TextAlign = ContentAlignment.MiddleLeft
            hl.Font = font(8.5, True)
            hl.ForeColor = renk(31, 56, 100)
            hl.BackColor = renk(230, 236, 246)
            hl.Padding = pad(8, 0, 0, 0)
            hl.Margin = pad(0, 10, 0, 4)
            tbl.RowStyles.Add(RowStyle(SizeType.Absolute, 40))
            tbl.Controls.Add(hl, 0, satir)
            tbl.SetColumnSpan(hl, 2)
            satir += 1
            toplam_h += 40
        h = 34
        yer = None
        deger = _s(rec.get(a['K']))
        t = a['T']
        c = None
        if t == 'text':
            c = TextBox()
            c.Text = deger
        elif t == 'memo':
            c = TextBox()
            c.Multiline = True
            c.ScrollBars = ScrollBars.Vertical
            c.AcceptsReturn = True
            c.Text = deger.replace('\r\n', '\n').replace('\n', '\r\n')
            h = 66
        elif t == 'combo':
            c = ComboBox()
            tekerlek_durdur(c)
            if not a.get('Strict'):
                combos.append(c)
                c.SizeChanged += secim_temizle_olayi
            sec = get_secenekler(a, listeler, tum)
            if deger and deger not in sec:
                sec = [deger] + sec
            c.DropDownStyle = ComboBoxStyle.DropDownList if a.get('Strict') else ComboBoxStyle.DropDown
            for x in sec:
                c.Items.Add(x)
            if a.get('Strict'):
                if deger:
                    c.SelectedItem = deger
            else:
                c.Text = deger
            # text box next to the combo (Aero / SAS / Thermal name)
            es = [x for x in ALANLAR if x.get('Bagli') == a['K']]
            if len(es) > 0:
                tb = TextBox()
                tb.Text = _s(rec.get(es[0]['K']))
                tb.Dock = DockStyle.Fill
                tb.Margin = pad(8, 3, 0, 3)
                c.Dock = DockStyle.Fill
                yer = TableLayoutPanel()
                yer.ColumnCount = 2
                yer.RowCount = 1
                yer.ColumnStyles.Add(ColumnStyle(SizeType.Absolute, 170))
                yer.ColumnStyles.Add(ColumnStyle(SizeType.Percent, 100))
                yer.RowStyles.Add(RowStyle(SizeType.Percent, 100))
                yer.Controls.Add(c, 0, 0)
                yer.Controls.Add(tb, 1, 0)
                ctl[es[0]['K']] = tb
        elif t == 'multitext':
            # several values (e.g. several part names / materials): type + Add, Remove from the list
            if a.get('List'):
                c = ComboBox()
                c.DropDownStyle = ComboBoxStyle.DropDown
                tekerlek_durdur(c)
                combos.append(c)
                c.SizeChanged += secim_temizle_olayi
                for x in get_secenekler(a, listeler, tum):
                    c.Items.Add(x)
            else:
                c = TextBox()
            c.Dock = DockStyle.Fill
            k = a['K']

            def anahtar_bas(s, e, k=k):
                if e.KeyCode == Keys.Return:
                    e.Handled = True
                    e.SuppressKeyPress = True
                    guvenli(liste_ekle, k)
            c.KeyDown += anahtar_bas
            lb = ListBox()
            lb.Dock = DockStyle.Fill
            lb.IntegralHeight = False
            for s in _split(deger, ';'):
                lb.Items.Add(s)
            lb.DoubleClick += (lambda s, e, k=k: guvenli(liste_cikar, k))
            bE = Button()
            bE.Text = 'Add'
            bE.Dock = DockStyle.Fill
            bE.Click += (lambda s, e, k=k: guvenli(liste_ekle, k))
            bC = Button()
            bC.Text = 'Remove'
            bC.Dock = DockStyle.Top
            bC.Height = 26
            bC.Click += (lambda s, e, k=k: guvenli(liste_cikar, k))
            yer = TableLayoutPanel()
            yer.ColumnCount = 2
            yer.RowCount = 2
            yer.ColumnStyles.Add(ColumnStyle(SizeType.Percent, 100))
            yer.ColumnStyles.Add(ColumnStyle(SizeType.Absolute, 84))
            yer.RowStyles.Add(RowStyle(SizeType.Absolute, 28))
            yer.RowStyles.Add(RowStyle(SizeType.Percent, 100))
            yer.Controls.Add(c, 0, 0)
            yer.Controls.Add(bE, 1, 0)
            yer.Controls.Add(lb, 0, 1)
            yer.Controls.Add(bC, 1, 1)
            lst[a['K']] = lb
            h = 104
        elif t == 'multi':
            c = FlowLayoutPanel()
            c.FlowDirection = FlowDirection.LeftToRight
            c.WrapContents = False
            secili = _split(deger, ',')
            cbs = []
            for s in listeler.get(a.get('List'), []):
                cb = CheckBox()
                cb.Text = s
                cb.AutoSize = True
                cb.Margin = pad(0, 6, 24, 0)
                cb.Checked = (s in secili)
                c.Controls.Add(cb)
                cbs.append(cb)
            chk[a['K']] = cbs
        elif t == 'master':
            c = ComboBox()
            c.DropDownStyle = ComboBoxStyle.DropDownList
            tekerlek_durdur(c)
            c.Items.Add('')
            secili = ''
            for m in [x for x in tum if x['ModelRolu'] == 'Master Model' and x['ID'] != rec['ID']]:
                metin = u'%s | %s | %s | %s' % (m['ID'], m['Proje'], m['ParcaAdi'], m['Layout'])
                c.Items.Add(metin)
                if m['ID'] == deger:
                    secili = metin
            c.SelectedItem = secili
        elif t == 'date':
            d0 = convert_to_tarih(deger)
            bugun_ = bugun()
            if a['K'] == 'TalepTarihi' and yeni:
                c = new_tarih_kontrol(bugun_, True, True)
            else:
                c = new_tarih_kontrol(d0, False, False)
        elif t == 'hist':
            c = TextBox()
            c.Multiline = True
            c.ReadOnly = True
            c.ScrollBars = ScrollBars.Vertical
            c.BackColor = renk(245, 245, 245)
            h = 110
            c.Text = u'\r\n'.join(rec['Gecmis'])
        l = Label()
        l.Text = a['L']
        l.Dock = DockStyle.Fill
        l.TextAlign = ContentAlignment.MiddleLeft
        if t in ('memo', 'hist', 'multitext'):
            l.TextAlign = ContentAlignment.TopLeft
            l.Padding = pad(0, 6, 0, 0)
        l.Font = font(9)
        l.ForeColor = renk(170, 30, 30) if '*' in a['L'] else renk(55, 55, 55)
        if yer is None:
            yer = c
        if t != 'date':
            yer.Dock = DockStyle.Fill
        else:
            c.Anchor = AnchorStyles.Left
            c.Width = 330
        tbl.RowStyles.Add(RowStyle(SizeType.Absolute, h))
        tbl.Controls.Add(l, 0, satir)
        tbl.Controls.Add(yer, 1, satir)
        ctl[a['K']] = c
        satir_yer[a['K']] = yer
        satir += 1
        toplam_h += h

    # lock fields by permission (read only: all; request mode: analyst fields + requester)
    for a in ALANLAR:
        if a['K'] not in satir_yer or a['T'] == 'hist':
            continue
        if mod == 'Gor' or (mod == 'Talep' and (a.get('Analizci') or a['K'] == 'TalepEden')):
            satir_yer[a['K']].Enabled = False
    if mod == 'Gor':
        btnK.Visible = False
        btnD.Visible = False
        btnI.Text = 'Close'
    tbl.Height = toplam_h
    pnl.Controls.Add(tbl)
    dlg.Controls.Add(pnl)
    dlg.Controls.Add(ban)
    dlg.Controls.Add(alt)

    def secimi_temizle(*args):
        for cb in combos:
            try:
                if not cb.Focused:
                    cb.SelectionLength = 0
            except Exception:
                pass
    dlg.Shown += secimi_temizle
    dlg.SizeChanged += secimi_temizle
    dlg.ResizeEnd += secimi_temizle

    # a text box next to Aero / SAS / Thermal is visible only for 'Ready' or 'Waiting'
    def ad_degisti(*args):
        for p in (('Aero', 'AeroAdi'), ('SAS', 'SASAdi'), ('Termal', 'TermalAdi')):
            gerek = (ctl[p[0]].Text == 'Ready' or ctl[p[0]].Text == 'Waiting')
            ctl[p[1]].Visible = gerek
            if not gerek:
                ctl[p[1]].Text = ''

    # 'Sub-Analysis' enables the linked master model; otherwise it is disabled and empty
    def rol_degisti(*args):
        alt_ = (ctl['ModelRolu'].Text == 'Sub-Analysis')
        ctl['AnaModelID'].Enabled = alt_
        if not alt_:
            ctl['AnaModelID'].SelectedItem = ''

    # choosing a master copies the model data from it (new records only)
    def ana_degisti(*args):
        id_ = get_alan_degeri(d, d.AnaAlan, ctl['AnaModelID'])
        if id_ and d.Yeni:
            m = [x for x in d.Tum if x['ID'] == id_]
            if len(m) > 0:
                for a in ALANLAR:
                    if a.get('Miras') and a['K'] in ctl:
                        set_alan_degeri(d, a, ctl[a['K']], _s(m[0].get(a['K'])))
            ad_degisti()
    for k in ('Aero', 'SAS', 'Termal'):
        ctl[k].SelectedIndexChanged += olay(ad_degisti)
    ctl['ModelRolu'].SelectedIndexChanged += olay(rol_degisti)
    ctl['AnaModelID'].SelectedIndexChanged += olay(ana_degisti)
    rol_degisti()
    ad_degisti()

    def iptal(s, e):
        dlg.DialogResult = DialogResult.Cancel
        dlg.Close()
    btnI.Click += iptal

    def kaydet_blok():
        yeni_rec = convert_to_kayit(d.Rec)
        for a in ALANLAR:
            if a['T'] in ('ro', 'hist'):
                continue
            c = ctl.get(a['K'])
            if c is None:
                continue
            yeni_rec[a['K']] = get_alan_degeri(d, a, c)
        if d.Mod == 'Gor':
            return
        if yeni_rec['Deadline']:
            dd = convert_to_tarih(yeni_rec['Deadline'])
            eski_d = convert_to_tarih(d.Eski['Deadline']) if d.Eski else None
            if dd is not None and dd < bugun() and not (eski_d is not None and eski_d == dd):
                mesaj('Due date cannot be earlier than today.', 'Warning')
                return
        if d.Mod == 'Talep' and d.Eski is not None:
            # fields the requester may not edit keep their old values
            for a in ALANLAR:
                if a.get('Analizci') and a['T'] != 'hist':
                    yeni_rec[a['K']] = d.Eski[a['K']]
            yeni_rec['TalepEden'] = d.Eski['TalepEden']
        if yeni_rec['ModelRolu'] != 'Sub-Analysis':
            yeni_rec['AnaModelID'] = ''
        if d.Yeni and not yeni_rec['Durum']:
            yeni_rec['Durum'] = get_listeler()['Durum'][0]
        if yeni_rec['Durum'] == 'Completed' and not yeni_rec['Tamamlanma']:
            yeni_rec['Tamamlanma'] = convert_from_tarih(datetime.now())
        taze = list(get_tum_kayitlar())
        hatalar = test_kayit(yeni_rec, taze, d.Eski)
        if len(hatalar) > 0:
            mesaj(u'\r\n'.join(hatalar), 'Warning')
            return
        try:
            sid = save_kayit(yeni_rec, d.Eski)
            d.SonID = unicode(sid)
            try:
                send_bildirim(yeni_rec, d.Eski, 'Yeni' if d.Yeni else 'Degisti')
            except Exception as e2:
                write_hata_log('Notification: ' + exc_msg(e2))
        except Exception as e:
            if exc_msg(e) == 'CAKISMA':
                mesaj("This record was modified by another user while you were editing it.\r\n\r\nYour changes were NOT saved so that the other user's changes are not overwritten. Close this window, reopen the record and make your change again.", 'Warning')
                dlg.DialogResult = DialogResult.Cancel
                dlg.Close()
                return
            raise
        d.Kaydedildi = True
        dlg.DialogResult = DialogResult.OK
        dlg.Close()

    def kaydet_tikla(s, e):
        d.KopyaIste = False
        guvenli(kaydet_blok)

    def kaydet_kopya_tikla(s, e):
        d.KopyaIste = True
        guvenli(kaydet_blok)
    btnK.Click += kaydet_tikla
    btnD.Click += kaydet_kopya_tikla

    U.DlgAcik = True
    try:
        dlg.ShowDialog(U.ana)
    finally:
        U.DlgAcik = False
    sonuc = d.Kaydedildi
    U.KopyaID = d.SonID if (sonuc and d.KopyaIste) else None
    dlg.Dispose()
    return sonuc
# -*- coding: utf-8 -*-


# ============================================================================
# PART 4 - USER INTERFACE: actions, dialogs, administration, start-up
# ============================================================================
def yeni_analiz():
    if not yenile_ben():
        return
    r = new_bos_kayit()
    l = get_listeler()
    r['TalepEden'] = G.Ben['Ad']
    r['Durum'] = l['Durum'][0]
    r['ModelRolu'] = l['ModelRolu'][0]
    r['TalepTarihi'] = convert_from_tarih(datetime.now())
    mod = 'Full' if test_tam_yetki(G.Ben) else 'Talep'
    if show_kayit_formu(r, True, mod):
        yenile_hepsi()
    if U.KopyaID:
        kopya_ac(U.KopyaID)


def kopya_ac(id_):
    """copy a record; the user picks a different option (Alpha, material model ...) and saves.
    Chained with 'Save + Duplicate'."""
    sayac = 0
    while id_ and sayac < 20:
        sayac += 1
        U.KopyaID = None
        src = read_kayit(os.path.join(G.KayitDir, id_ + '.json'))
        if src is None:
            mesaj('The source record could not be read.', 'Warning')
            return
        if get_yetki(src, G.Ben).Mod == 'Gor':
            mesaj('You can only duplicate your own requests.', 'Warning')
            return
        r = new_bos_kayit()
        l = get_listeler()
        for a in ALANLAR:
            if a['T'] in ('ro', 'hist') or a.get('Analizci'):
                continue
            r[a['K']] = src[a['K']]
        r['TalepEden'] = G.Ben['Ad']
        r['Durum'] = l['Durum'][0]
        r['TalepTarihi'] = convert_from_tarih(datetime.now())
        dd = convert_to_tarih(src['Deadline'])
        if dd is not None and dd < bugun():
            r['Deadline'] = ''
        mod = 'Full' if test_tam_yetki(G.Ben) else 'Talep'
        U.KopyaKaynak = id_
        try:
            ok = show_kayit_formu(r, True, mod)
        finally:
            U.KopyaKaynak = None
        if ok:
            yenile_hepsi()
        id_ = U.KopyaID
    U.KopyaID = None


def kopyala_secili():
    if not yenile_ben():
        return
    id_ = get_secili_id()
    if not id_:
        mesaj('Please select the record to duplicate first.')
        return
    kopya_ac(id_)


def duzenle_secili():
    id_ = get_secili_id()
    if not id_:
        mesaj('Please select a record from the list first.')
        return
    ac_kayit(id_)


def ac_kayit(id_):
    if not yenile_ben():
        return
    r = read_kayit(os.path.join(G.KayitDir, id_ + '.json'))
    if r is None:
        mesaj('The record file could not be read (it may have been deleted). Refreshing the list.', 'Warning')
        yenile_hepsi()
        return
    y = get_yetki(r, G.Ben)
    if show_kayit_formu(r, False, y.Mod):
        yenile_hepsi()
    else:
        otomatik_yenile()
    if U.KopyaID:
        kopya_ac(U.KopyaID)


def sil_secili():
    id_ = get_secili_id()
    if not id_:
        mesaj('Please select the record to delete first.')
        return
    if not yenile_ben():
        return
    rec = read_kayit(os.path.join(G.KayitDir, id_ + '.json'))
    if rec is None:
        mesaj('The record file could not be read (it may have been deleted). Refreshing the list.', 'Warning')
        yenile_hepsi()
        return
    yet = get_yetki(rec, G.Ben)
    if not yet.Sil:
        mesaj('You do not have permission to delete this record.\r\n\r\nDesign engineers can delete only their own requests while the status is still the first one (Pending). Leaders and structural analysts can delete any record.', 'Warning')
        return
    tum = list(get_tum_kayitlar())
    c = get_cocuklar(tum, id_)
    if len(c) > 0:
        mesaj('This record is the master model of %d sub-analysis(es); delete them or link them to another master model first.' % len(c), 'Warning')
        return
    if mesaj("Delete record %s?\r\n(The file is moved to the 'Silinenler' folder and can be restored if needed.)" % id_, 'Question', 'YesNo') != 'Yes':
        return
    remove_kayit(id_)
    try:
        send_bildirim(rec, None, 'Silindi')
    except Exception as e:
        write_hata_log('Notification: ' + exc_msg(e))
    yenile_hepsi()


def aktar_excel():
    sfd = SaveFileDialog()
    sfd.Filter = 'CSV (opens in Excel) (*.csv)|*.csv'
    sfd.FileName = 'AnalysisList_' + datetime.now().strftime('%Y%m%d_%H%M') + '.csv'
    sfd.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
    if str(sfd.ShowDialog(U.ana)) != 'OK':
        return
    export_kayitlar_csv(list(get_tum_kayitlar()), sfd.FileName)
    if mesaj('Saved:\r\n%s\r\n\r\nOpen in Excel?' % sfd.FileName, 'Question', 'YesNo') == 'Yes':
        Process.Start(sfd.FileName)


# ---------------------------------------------------------------- small dialogs
def _mini_form(baslik, w, h, merkez='CenterParent'):
    f = Form()
    f.Text = baslik
    f.StartPosition = getattr(FormStartPosition, merkez)
    f.Size = boyut(w, h)
    f.Font = font(9)
    return f


def _alt_panel(yukseklik=46):
    alt = FlowLayoutPanel()
    alt.Dock = DockStyle.Bottom
    alt.Height = yukseklik
    alt.FlowDirection = FlowDirection.RightToLeft
    alt.Padding = pad(8, 8, 8, 0)
    return alt


def _dugme(metin, w, h=28):
    b = Button()
    b.Text = metin
    b.Width = w
    b.Height = h
    return b


def _duzgun_grid(basliklar, fill=True):
    g = DataGridView()
    cift_tampon(g)
    g.Dock = DockStyle.Fill
    g.ReadOnly = True
    g.AllowUserToAddRows = False
    g.RowHeadersVisible = False
    g.BackgroundColor = Color.White
    g.SelectionMode = DataGridViewSelectionMode.FullRowSelect
    g.MultiSelect = False
    if fill:
        g.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
    g.EnableHeadersVisualStyles = False
    g.ColumnHeadersDefaultCellStyle.BackColor = renk(31, 56, 100)
    g.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
    for b in basliklar:
        g.Columns.Add(b, b)

    def md(s, e):
        try:
            if s.HitTest(e.X, e.Y).Type == _NONE_HIT:
                sec_temizle(s)
        except Exception:
            pass
    g.MouseDown += md
    return g


def duzenle_listeler():
    anahtarlar = [('Status', 'Durum'), ('Priority', 'Oncelik'), ('Material', 'Malzeme'), ('Analysis Type', 'AnalizTuru'),
                  ('Condition', 'Kosul')]
    l = get_listeler()
    f = _mini_form('Edit lists', 560, 520)
    f.ShowInTaskbar = False
    f.MinimizeBox = False
    tc = TabControl()
    tc.Dock = DockStyle.Fill
    kutular = {}
    for a in anahtarlar:
        p = TabPage(a[0])
        t = TextBox()
        t.Multiline = True
        t.ScrollBars = ScrollBars.Vertical
        t.Dock = DockStyle.Fill
        t.AcceptsReturn = True
        t.Text = u'\r\n'.join(l[a[1]])
        p.Controls.Add(t)
        tc.TabPages.Add(p)
        kutular[a[1]] = t
    ipucu = Label()
    ipucu.Dock = DockStyle.Top
    ipucu.Height = 38
    ipucu.Text = "  Each line is one option. The 'Completed' and 'Cancelled' statuses are fixed.\r\n  People (requesters, analysts) come from the registered users."
    alt = Panel()
    alt.Dock = DockStyle.Bottom
    alt.Height = 46
    ok = _dugme('Save', 100)
    iptal = _dugme('Cancel', 100)
    akis = _alt_panel()
    akis.Dock = DockStyle.Fill
    akis.Controls.Add(iptal)
    akis.Controls.Add(ok)
    alt.Controls.Add(akis)
    f.Controls.Add(tc)
    f.Controls.Add(ipucu)
    f.Controls.Add(alt)

    def iptal_tikla(s, e):
        f.DialogResult = DialogResult.Cancel
        f.Close()

    def kaydet():
        yeni = get_listeler()
        for a in anahtarlar:
            satirlar = [x.strip() for x in re.split(r'\r?\n', kutular[a[1]].Text) if x.strip()]
            yeni[a[1]] = satirlar
        save_listeler(yeni)
        f.DialogResult = DialogResult.OK
        f.Close()
    iptal.Click += iptal_tikla
    ok.Click += olay(kaydet)
    U.DlgAcik = True
    try:
        sonuc = str(f.ShowDialog(U.ana))
    finally:
        U.DlgAcik = False
    f.Dispose()
    if sonuc == 'OK':
        yenile_hepsi()


def show_girdi(baslik, etiket, varsayilan='', secenekler=None):
    """small input window: text or (with options) a selection list. Returns None on cancel."""
    f = Form()
    f.Text = baslik
    f.StartPosition = FormStartPosition.CenterParent
    f.Size = boyut(440, 175)
    f.FormBorderStyle = FormBorderStyle.FixedDialog
    f.MinimizeBox = False
    f.MaximizeBox = False
    f.ShowInTaskbar = False
    f.Font = font(9)
    f.Padding = pad(12, 12, 12, 12)
    l = Label()
    l.Text = etiket
    l.Dock = DockStyle.Top
    l.Height = 44
    if secenekler is not None:
        c = ComboBox()
        c.DropDownStyle = ComboBoxStyle.DropDownList
        tekerlek_durdur(c)
        for x in secenekler:
            c.Items.Add(x)
        if varsayilan:
            c.SelectedItem = varsayilan
    else:
        c = TextBox()
        c.Text = varsayilan
    c.Dock = DockStyle.Top
    alt = FlowLayoutPanel()
    alt.Dock = DockStyle.Bottom
    alt.Height = 44
    alt.FlowDirection = FlowDirection.RightToLeft
    ok = _dugme('OK', 90)
    ip = _dugme('Cancel', 90)
    alt.Controls.Add(ip)
    alt.Controls.Add(ok)
    f.Controls.Add(c)
    f.Controls.Add(l)
    f.Controls.Add(alt)
    f.AcceptButton = ok
    f.CancelButton = ip
    sonuc = {'v': None}

    def ok_tikla(s, e):
        sonuc['v'] = (c.Text or '').strip()
        f.DialogResult = DialogResult.OK
        f.Close()

    def ip_tikla(s, e):
        f.DialogResult = DialogResult.Cancel
        f.Close()
    ok.Click += ok_tikla
    ip.Click += ip_tikla
    f.ShowDialog(U.ana)
    f.Dispose()
    return sonuc['v']


def show_kullanici_kayit(hesap):
    """first start: name + role. Returns None when cancelled (the program closes)."""
    ilk = (len(get_tum_kullanicilar()) == 0)
    f = Form()
    f.Text = 'Analysis Tracker - First login'
    f.StartPosition = FormStartPosition.CenterScreen
    f.Size = boyut(520, 450)
    f.FormBorderStyle = FormBorderStyle.FixedDialog
    f.MinimizeBox = False
    f.MaximizeBox = False
    f.Font = font(9)
    l1 = Label()
    l1.Location = Point(16, 14)
    l1.Size = boyut(470, 112)
    l1.Text = u'Welcome. This Windows account is not registered yet:\r\n%s\r\n\r\nPlease enter your name and role. You will be recognised automatically next time.' % hesap
    l2 = Label()
    l2.Text = 'Your name (full name):'
    l2.Location = Point(16, 134)
    l2.Size = boyut(470, 22)
    tAd = TextBox()
    tAd.Location = Point(16, 158)
    tAd.Size = boyut(470, 26)
    l3 = Label()
    l3.Text = 'Your role:'
    l3.Location = Point(16, 198)
    l3.Size = boyut(470, 22)
    cRol = ComboBox()
    cRol.DropDownStyle = ComboBoxStyle.DropDownList
    cRol.Location = Point(16, 222)
    cRol.Size = boyut(470, 26)
    tekerlek_durdur(cRol)
    for x in ROLLER:
        cRol.Items.Add(x)
    nt = Label()
    nt.Location = Point(16, 266)
    nt.Size = boyut(470, 44)
    nt.Text = 'You are the first user, so you become the administrator.' if ilk else 'Only the administrator can change your name or role later.'
    ok = _dugme('Continue', 110, 32)
    ok.Location = Point(262, 336)
    ip = _dugme('Exit', 100, 32)
    ip.Location = Point(382, 336)
    for c in (l1, l2, tAd, l3, cRol, nt, ok, ip):
        f.Controls.Add(c)
    f.AcceptButton = ok
    f.CancelButton = ip
    sonuc = {'v': None}

    def ip_tikla(s, e):
        f.DialogResult = DialogResult.Cancel
        f.Close()

    def ok_tikla(s, e):
        try:
            sonuc['v'] = new_kullanici(hesap, tAd.Text, unicode(cRol.Text))
            f.DialogResult = DialogResult.OK
            f.Close()
        except Exception as ex:
            mesaj(exc_msg(ex), 'Warning')
    ok.Click += ok_tikla
    ip.Click += ip_tikla
    f.ShowDialog()
    f.Dispose()
    return sonuc['v']


def initialize_kimlik():
    hesap = get_hesap_adi()
    k = get_kullanici(hesap)
    if k is None:
        k = show_kullanici_kayit(hesap)
    if k is None:
        return None
    try:
        k['SonGiris'] = now_s()
        k['Bilgisayar'] = Environment.MachineName
        save_kullanici(k)
    except Exception:
        pass
    return k


# ---------------------------------------------------------------- notifications window
def goster_bildirimler():
    if not yenile_ben():
        return
    liste = get_benim_bildirimler(G.Ben)
    f = _mini_form('Notifications', 900, 460)
    f.ShowInTaskbar = False
    f.MinimizeBox = False
    g = _duzgun_grid([], fill=False)
    for b, a, w in (('New', 'Yeni', 45), ('Time', 'Zaman', 130), ('Record', 'Kayit', 75), ('Message', 'Mesaj', 520), ('From', 'Kimden', 100)):
        c = DataGridViewTextBoxColumn()
        c.HeaderText = b
        c.Name = a
        c.Width = w
        g.Columns.Add(c)
    nokta = unichr(0x25CF)
    for b in liste:
        g.Rows.Add(nokta if b['Okunmadi'] else '', b['Zaman'].replace('T', ' ')[:16], b['KayitID'], b['Mesaj'], b['Kimden'])
    kalin = font(9, True)

    def bicim(s, e):
        try:
            if e.RowIndex >= 0 and s.Rows[e.RowIndex].Cells[0].Value:
                e.CellStyle.Font = kalin
        except Exception:
            pass
    g.CellFormatting += bicim
    alt = _alt_panel()
    bKapat = _dugme('Close (marks all as read)', 170)
    bAc = _dugme('Open record', 110)
    alt.Controls.Add(bKapat)
    alt.Controls.Add(bAc)
    f.Controls.Add(g)
    f.Controls.Add(alt)
    durum = {'ac': ''}

    def sec():
        r = g.CurrentRow
        if r is not None:
            id_ = r.Cells[2].Value
            if id_:
                durum['ac'] = unicode(id_)
                f.Close()
    bAc.Click += olay(sec)
    g.CellDoubleClick += (lambda s, e: guvenli(sec) if e.RowIndex >= 0 else None)
    bKapat.Click += (lambda s, e: f.Close())
    U.DlgAcik = True
    try:
        f.ShowDialog(U.ana)
    finally:
        U.DlgAcik = False
    f.Dispose()
    G.Ben = set_bildirim_okundu(G.Ben)
    yenile_bildirim()
    id_ = durum['ac']
    if id_ and File.Exists(os.path.join(G.KayitDir, id_ + '.json')):
        ac_kayit(id_)


# ---------------------------------------------------------------- administrator screens
def yonet_kullanicilar():
    f = _mini_form('Manage users', 960, 460)
    f.ShowInTaskbar = False
    f.MinimizeBox = False
    g = _duzgun_grid(['Name', 'Role', 'Admin', 'Windows account', 'Computer', 'Registered', 'Last login'])
    alt = _alt_panel()
    bs = {}
    for b, w in (('Close', 90), ('Remove', 90), ('Admin on/off', 100), ('Rename', 90), ('Change role', 100)):
        x = _dugme(b, w)
        alt.Controls.Add(x)
        bs[b] = x
    f.Controls.Add(g)
    f.Controls.Add(alt)

    def doldur():
        g.Rows.Clear()
        for k in get_tum_kullanicilar():
            g.Rows.Add(k['Ad'], k['Rol'], 'Yes' if test_admin(k) else '', k['Hesap'], k['Bilgisayar'],
                       k['Kayit'].replace('T', ' '), k['SonGiris'].replace('T', ' '))
        sec_temizle(g)

    def sec():
        r = g.CurrentRow
        if r is None:
            mesaj('Please select a user first.')
            return None
        return get_kullanici(unicode(r.Cells[3].Value))

    def rol_degistir():
        u = sec()
        if u is None:
            return
        r = show_girdi('Change role', 'New role for %s:' % u['Ad'], u['Rol'], ROLLER)
        if r:
            u['Rol'] = r
            save_kullanici(u)
            doldur()

    def yeniden_adlandir():
        u = sec()
        if u is None:
            return
        yeni = show_girdi('Rename user', 'New name for %s:' % u['Ad'], u['Ad'])
        if not yeni or yeni == u['Ad']:
            return
        if len(yeni) < 2:
            mesaj('The name is too short.', 'Warning')
            return
        if test_ad_kullanimda(yeni, u['Hesap']):
            mesaj('This name is already used by another user.', 'Warning')
            return
        s = update_kayitlarda_ad(u['Ad'], yeni, ['TalepEden', 'AnalizYapan'], False)
        eski_ad = u['Ad']
        u['Ad'] = yeni
        save_kullanici(u)
        ek = ('; %d could not be updated (changed by someone else meanwhile)' % s.Hatali) if s.Hatali else ''
        mesaj("Renamed '%s' to '%s'. Updated %d record(s)%s." % (eski_ad, yeni, s.Guncellenen, ek))
        doldur()

    def admin_degistir():
        u = sec()
        if u is None:
            return
        adminler = [x for x in get_tum_kullanicilar() if x['Admin']]
        if u['Admin'] and len(adminler) <= 1:
            mesaj('At least one administrator is required.', 'Warning')
            return
        u['Admin'] = not u['Admin']
        save_kullanici(u)
        doldur()

    def kaldir():
        u = sec()
        if u is None:
            return
        if u['Hesap'] == G.Ben['Hesap']:
            mesaj('You cannot remove your own registration.', 'Warning')
            return
        if mesaj('Remove %s (%s)?\r\nThe user will have to register again at the next start. Their records are kept.' % (u['Ad'], u['Hesap']), 'Question', 'YesNo') != 'Yes':
            return
        remove_kullanici(u['Hesap'])
        doldur()
    bs['Close'].Click += (lambda s, e: f.Close())
    bs['Change role'].Click += olay(rol_degistir)
    bs['Rename'].Click += olay(yeniden_adlandir)
    bs['Admin on/off'].Click += olay(admin_degistir)
    bs['Remove'].Click += olay(kaldir)
    doldur()
    U.DlgAcik = True
    try:
        f.ShowDialog(U.ana)
    finally:
        U.DlgAcik = False
    f.Dispose()
    yenile_hepsi()


def geri_yukle_silinenler():
    f = _mini_form('Restore deleted records', 820, 440)
    f.ShowInTaskbar = False
    f.MinimizeBox = False
    g = _duzgun_grid(['ID', 'Deleted', 'Part', 'Project', 'Requester', 'File'])
    alt = _alt_panel()
    bK = _dugme('Close', 90)
    bR = _dugme('Restore', 100)
    alt.Controls.Add(bK)
    alt.Controls.Add(bR)
    f.Controls.Add(g)
    f.Controls.Add(alt)

    def doldur():
        g.Rows.Clear()
        for s in get_silinen_kayitlar():
            g.Rows.Add(s.ID, s.Silinme.strftime('%Y-%m-%d %H:%M'), s.Parca, s.Proje, s.Talep, s.Dosya)
        sec_temizle(g)

    def geri_yukle():
        r = g.CurrentRow
        if r is None:
            mesaj('Please select a record first.')
            return
        restore_kayit(unicode(r.Cells[5].Value), unicode(r.Cells[0].Value))
        doldur()
    bK.Click += (lambda s, e: f.Close())
    bR.Click += olay(geri_yukle)
    doldur()
    U.DlgAcik = True
    try:
        f.ShowDialog(U.ana)
    finally:
        U.DlgAcik = False
    f.Dispose()
    yenile_hepsi()


def devret_analizci():
    tum = list(get_tum_kayitlar())
    adlar = sorted(set([r['AnalizYapan'] for r in tum if not test_kapali(r) and r['AnalizYapan']]))
    if len(adlar) == 0:
        mesaj('There are no open analyses with an assigned analyst.')
        return
    kimden = show_girdi('Reassign analyst', 'Reassign the OPEN analyses of this analyst:', '', adlar)
    if not kimden:
        return
    kime = [k['Ad'] for k in get_tum_kullanicilar() if k['Rol'] == 'Structural Analysis Engineer' and k['Ad'] != kimden]
    if len(kime) == 0:
        mesaj('There is no other registered structural analysis engineer.', 'Warning')
        return
    hedef = show_girdi('Reassign analyst', 'Reassign the open analyses of %s to:' % kimden, '', kime)
    if not hedef:
        return
    sayi = len([r for r in tum if not test_kapali(r) and r['AnalizYapan'] == kimden])
    if mesaj('Reassign %d open analysis(es) from %s to %s?' % (sayi, kimden, hedef), 'Question', 'YesNo') != 'Yes':
        return
    s = update_kayitlarda_ad(kimden, hedef, ['AnalizYapan'], True)
    h = [k for k in get_tum_kullanicilar() if k['Ad'] == hedef]
    if s.Guncellenen > 0:
        try:
            new_bildirim_dosyasi('', 'Reassigned', u'%d open analysis(es) were reassigned from %s to you by %s' % (s.Guncellenen, kimden, G.Ben['Ad']), [k['Hesap'] for k in h])
        except Exception:
            pass
    ek = ('; %d could not be updated (changed by someone else meanwhile) - please repeat' % s.Hatali) if s.Hatali else ''
    mesaj('Reassigned %d analysis(es)%s.' % (s.Guncellenen, ek))
    yenile_hepsi()


def yedekle_veri():
    s = backup_veri()
    mesaj('Backup created (%d files):\r\n%s\r\n\r\nThe last 10 backups are kept.' % (s.Dosya, s.Yol))


def goster_profil():
    k = G.Ben
    mesaj(u'Name: %s\r\nRole: %s\r\nAdministrator: %s\r\nWindows account: %s\r\nComputer: %s\r\nRegistered: %s\r\n\r\nOnly the administrator can change your name or role.' % (
        k['Ad'], k['Rol'], 'Yes' if test_admin(k) else 'No', k['Hesap'], k['Bilgisayar'], k['Kayit'].replace('T', ' ')))


def degistir_veri_klasoru():
    fb = FolderBrowserDialog()
    fb.Description = 'Select the shared data folder (everyone must use the same folder)'
    fb.SelectedPath = G.Kok
    if str(fb.ShowDialog(U.ana)) != 'OK':
        return
    initialize_veri(fb.SelectedPath)
    set_veri_klasoru(fb.SelectedPath)
    yenile_hepsi()


def ornek_veri_ekle():
    if mesaj('Add sample records for testing?\r\n(8 master models with sub-analyses + 4 standalone analyses)\r\n(To remove them, select the records and press Delete.)', 'Question', 'YesNo') == 'Yes':
        add_ornek_veri()
        yenile_hepsi()


def hakkinda():
    mesaj(u'Analysis Tracker\r\n\r\nData folder: %s\r\nUser: %s (%s)\r\nError log: %s' % (G.Kok, G.Ben['Ad'], G.Ben['Rol'], U.HataLog))


# ---------------------------------------------------------------- events
def olaylari_bagla():
    U.tbYeni.Click += olay(yeni_analiz)
    U.tbDuzenle.Click += olay(duzenle_secili)
    U.tbSil.Click += olay(sil_secili)
    U.tbKopya.Click += olay(kopyala_secili)
    U.tbYenile.Click += olay(yenile_hepsi)
    U.tbExcel.Click += olay(aktar_excel)
    U.miListe.Click += olay(duzenle_listeler)
    U.miKlasor.Click += olay(degistir_veri_klasoru)
    U.miAc.Click += olay(lambda: Process.Start('explorer.exe', '"' + G.Kok + '"'))
    U.miOrnek.Click += olay(ornek_veri_ekle)
    U.miHak.Click += olay(hakkinda)
    U.miProfil.Click += olay(goster_profil)
    U.tbBildirim.Click += olay(goster_bildirimler)
    U.miKull.Click += olay(yonet_kullanicilar)
    U.miSilinen.Click += olay(geri_yukle_silinenler)
    U.miDevir.Click += olay(devret_analizci)
    U.miYedek.Click += olay(yedekle_veri)

    # a click on an empty area removes the blue selection (all tables)
    def grid_md(s, e):
        try:
            if s.HitTest(e.X, e.Y).Type == _NONE_HIT:
                sec_temizle(s)
        except Exception:
            pass
    for gg in (U.grid, U.gDurum, U.gKisi, U.gYakin):
        gg.MouseDown += grid_md

    def lv_md(s, e):
        try:
            if s.GetItemAt(e.X, e.Y) is None:
                s.SelectedItems.Clear()
        except Exception:
            pass
    U.lv.MouseDown += lv_md

    def bos_tikla(s, e):
        for gg in (U.gDurum, U.gKisi, U.gYakin, U.grid):
            sec_temizle(gg)
        try:
            U.lv.SelectedItems.Clear()
        except Exception:
            pass
    for ct in (U.tp1, U.tp2, U.tp3, U.fp, U.ozetTablo, U.fpAgac):
        ct.Click += bos_tikla

    def grid_cift(s, e):
        if e.RowIndex >= 0:
            guvenli(duzenle_secili)
    U.grid.CellDoubleClick += grid_cift

    def grid_tus(s, e):
        if e.KeyCode == Keys.Return:
            e.Handled = True
            e.SuppressKeyPress = True
            guvenli(duzenle_secili)
    U.grid.KeyDown += grid_tus

    kalin9 = font(9, True)

    def grid_bicim(s, e):
        try:
            if e.RowIndex < 0:
                return
            drv = s.Rows[e.RowIndex].DataBoundItem
            if drv is None:
                return
            durum = unicode(drv['Durum'])
            if durum in KAPALI:
                e.CellStyle.ForeColor = Color.Gray
            kol = s.Columns[e.ColumnIndex].Name
            if kol == 'Gecikme' and unicode(drv['Gecikme']):
                e.CellStyle.ForeColor = Color.White
                e.CellStyle.BackColor = renk(192, 0, 0)
            if kol == 'Rol' and unicode(drv['Rol']) == 'Master Model':
                e.CellStyle.Font = kalin9
                e.CellStyle.ForeColor = renk(31, 56, 140)
        except Exception:
            pass
    U.grid.CellFormatting += grid_bicim

    def lv_cift(s, e):
        if U.lv.SelectedItems.Count > 0:
            guvenli(duzenle_secili)
    U.lv.DoubleClick += lv_cift

    def lv_tus(s, e):
        if e.KeyCode == Keys.Return and s.SelectedItems.Count > 0:
            e.Handled = True
            guvenli(duzenle_secili)
    U.lv.KeyDown += lv_tus

    agac_degisti = olay(yenile_agac)
    U.txtAgac.TextChanged += agac_degisti
    U.chkAgac.CheckedChanged += agac_degisti

    def filtre_degisti(s, e):
        if not U.FiltreKilit:
            guvenli(yenile_liste)
    U.txtAra.TextChanged += filtre_degisti
    U.cmbDurum.SelectedIndexChanged += filtre_degisti
    U.cmbKisi.SelectedIndexChanged += filtre_degisti
    U.chkAcik.CheckedChanged += filtre_degisti
    U.chkGec.CheckedChanged += filtre_degisti

    def temizle(s, e):
        U.FiltreKilit = True
        U.txtAra.Text = ''
        U.chkAcik.Checked = False
        U.chkGec.Checked = False
        U.cmbDurum.SelectedItem = 'All'
        U.cmbKisi.SelectedItem = 'All'
        # header-click sorting is reset too: newest request on top
        U.dt.DefaultView.Sort = 'ID DESC'
        try:
            U.grid.ClearSelection()
        except Exception:
            pass
        U.FiltreKilit = False
        guvenli(yenile_liste)
    U.btnTemizle.Click += temizle

    U.ana.KeyPreview = True

    def ana_tus(s, e):
        if e.Control and e.KeyCode == Keys.N:
            e.Handled = True
            guvenli(yeni_analiz)
        elif e.KeyCode == Keys.F5:
            e.Handled = True
            guvenli(yenile_hepsi)
    U.ana.KeyDown += ana_tus

    zaman = Timer()
    zaman.Interval = 20000
    zaman.Tick += olay(otomatik_yenile)
    U.Zamanlayici = zaman

    def goster(s, e):
        guvenli(yenile_hepsi)

        def ornek_sor():
            if len(U.Tum) == 0 and test_admin(G.Ben):
                if mesaj('There are no analysis records yet.\r\n\r\nAdd sample data (8 master models with sub-analyses and 4 standalone analyses) so you can try the program?\r\n(You can delete them later, or add them from Admin > Add sample data.)', 'Question', 'YesNo') == 'Yes':
                    add_ornek_veri()
                    yenile_hepsi()
        guvenli(ornek_sor)
        guvenli(yenile_bildirim)
        zaman.Start()
        U.ana.Activate()
    U.ana.Shown += goster

    def kapandi(s, e):
        try:
            zaman.Stop()
            zaman.Dispose()
        except Exception:
            pass
        try:
            if U.Tray is not None:
                U.Tray.Visible = False
                U.Tray.Dispose()
        except Exception:
            pass
        try:
            AppDomain.CurrentDomain.SetData(FORM_KEY, None)
        except Exception:
            pass
    U.ana.FormClosed += kapandi


# ---------------------------------------------------------------- start
def main():
    # close a previous instance of this window (the script can be started again from Mechanical)
    try:
        eski = AppDomain.CurrentDomain.GetData(FORM_KEY)
        if eski is not None:
            try:
                eski.Close()
            except Exception:
                pass
            AppDomain.CurrentDomain.SetData(FORM_KEY, None)
    except Exception:
        pass
    try:
        G.Klasor = program_klasoru()
        U.HataLog = os.path.join(G.Klasor, 'AnalizTakip.hata.log')
        if TEST_HESAP:
            G.HesapOverride = TEST_HESAP
        build_main()
        veri = get_veri_klasoru()
        try:
            initialize_veri(veri)
        except Exception as e:
            write_hata_log(exc_msg(e))
            MessageBox.Show('Cannot access the data folder:\r\n%s\r\n\r\n%s\r\n\r\nYou can choose another folder in the next window.' % (veri, exc_msg(e)),
                            'Analysis Tracker', MessageBoxButtons.OK, MessageBoxIcon.Warning)
            fb = FolderBrowserDialog()
            fb.Description = 'Select the shared data folder'
            if str(fb.ShowDialog()) != 'OK':
                return
            veri = fb.SelectedPath
            initialize_veri(veri)
            set_veri_klasoru(veri)
        G.Ben = initialize_kimlik()
        if G.Ben is None:
            return
        try:
            remove_eski_bildirimler()
        except Exception:
            pass
        yetki_uygula()
        try:
            tray = NotifyIcon()
            tray.Icon = SystemIcons.Information
            tray.Text = 'Analysis Tracker'
            tray.Visible = True
            tray.BalloonTipClicked += olay(goster_bildirimler)
            U.Tray = tray
        except Exception:
            U.Tray = None
        olaylari_bagla()
        AppDomain.CurrentDomain.SetData(FORM_KEY, U.ana)
        U.ana.Show()      # non-modal: Mechanical keeps its own message loop
    except Exception as e:
        write_hata_log(traceback.format_exc() + u' | ' + exc_msg(e))
        MessageBox.Show('Unexpected error:\r\n\r\n%s\r\n\r\nDetails: %s' % (exc_msg(e), U.HataLog),
                        'Analysis Tracker', MessageBoxButtons.OK, MessageBoxIcon.Error)


main()
