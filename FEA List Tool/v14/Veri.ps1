# =====================================================================
#  Analiz Takip - veri katmani (Veri.ps1)            Windows PowerShell 5.1 uyumlu
#  Her analiz kaydi ortak klasorde ayri bir JSON dosyasidir:
#     <Veri>\Kayitlar\AN-0001.json
#  Ayni anda birden fazla kullanici icin: kayit basina kilit + revizyon (cakisma) kontrolu.
# =====================================================================

$script:Kapali  = @('Completed', 'Cancelled')
# eski (Turkce) kayitlarda kalan degerleri Ingilizce karsiliklarina cevirir (okuma sirasinda)
$script:DegerMap = @{
    'Beklemede' = 'Pending'; 'Model Hazırlanıyor' = 'Model Preparation'; 'Analiz Devam Ediyor' = 'In Progress'; 'İnceleme / Kontrol' = 'Review / Check'
    'Askıda' = 'On Hold'; 'Completed' = 'Completed'; 'İptal' = 'Cancelled'; 'Yüksek' = 'High'; 'Orta' = 'Medium'; 'Düşük' = 'Low'
    'Elastik' = 'Elastic'; 'Elastoplastik' = 'Elastoplastic'; 'Alpha On' = 'On'; 'Alpha Off' = 'Off'
    'Bağımsız' = 'Standalone'; 'Master Model' = 'Master Model'; 'Sub-Analysis' = 'Sub-Analysis'
    'Gerekmiyor' = 'Not Required'; 'Bekleniyor' = 'Waiting'; 'Hazır' = 'Ready'; 'Diğer' = 'Other'
    'Statik Yapısal' = 'Static Structural'; 'Yorulma (LCF/HCF)' = 'Fatigue (LCF/HCF)'; 'Sürünme' = 'Creep'; 'Termal-Yapısal' = 'Thermal-Structural'; 'Kontak / Cıvata' = 'Contact / Bolt'
}
function Convert-Deger {
    param([string]$S)
    if ([string]::IsNullOrEmpty($S)) { return $S }
    $parcalar = @($S -split '\s*,\s*' | ForEach-Object { if ($script:DegerMap.ContainsKey($_)) { $script:DegerMap[$_] } else { $_ } })
    if ($parcalar.Count -gt 1) { return ($parcalar -join ', ') }
    if ($script:DegerMap.ContainsKey($S)) { return $script:DegerMap[$S] }
    return $S
}
$script:Invar   = [Globalization.CultureInfo]::InvariantCulture
$script:Utf8    = New-Object Text.UTF8Encoding($false)
$script:Kok = ''; $script:KayitDir = ''; $script:SilinenDir = ''

# Form / kayit alan tanimlari (arayuz bunlardan form uretir)
$script:Alanlar = @(
    # Miras = $true : ana model secilince alt analize kopyalanir.  Bagli = ust alanin yaninda acilan metin kutusu.
    # SadeceDuzenle = $true : yeni kayitta sorulmaz, kayit sonradan acilinca gorunur.
    # multitext = birden fazla girdi (liste); multi = birden fazla onay kutusu
    @{ K = 'ID';          L = 'ID (auto)';                             T = 'ro' },
    @{ K = 'Proje';       L = 'Project / Program';                     T = 'text'; Miras = $true },
    @{ K = 'ParcaAdi';    L = 'Part Name(s) *';                        T = 'multitext'; Miras = $true },
    @{ K = 'Layout';      L = 'Layout';                                T = 'text'; Miras = $true },
    @{ K = 'ModelNo';     L = 'Model ID';                              T = 'text'; Miras = $true },
    @{ K = 'ModelRolu';   L = 'Model Role';                            T = 'combo'; List = 'ModelRolu'; Strict = $true },
    @{ K = 'AnaModelID';  L = 'Linked Master Model';                   T = 'master' },
    @{ K = 'ModelTipi';   L = 'Model Type';                            T = 'combo'; List = 'ModelTipi'; Strict = $true; Miras = $true },
    @{ K = 'AnalizTipi';  L = 'Material Model *';                      T = 'combo'; List = 'AnalizTipi'; Strict = $true; Miras = $true },
    @{ K = 'Alpha';       L = 'Alpha';                                 T = 'combo'; List = 'Alpha'; Strict = $true; Miras = $true },
    @{ K = 'Malzeme';     L = 'Material(s)';                           T = 'multitext'; List = 'Malzeme'; Miras = $true },
    @{ K = 'Amac';        L = 'Purpose';                               T = 'memo' },
    @{ K = 'Kosul';       L = 'Condition';                             T = 'combo'; List = 'Kosul' },
    @{ K = 'Hiz';         L = 'Speed';                                 T = 'text' },
    @{ K = 'Aero';        L = 'Aero';                                  T = 'combo'; List = 'Girdi'; Strict = $true; Miras = $true },
    @{ K = 'AeroAdi';     L = 'Aero Name';                             T = 'text'; Bagli = 'Aero'; Miras = $true },
    @{ K = 'SAS';         L = 'SAS';                                   T = 'combo'; List = 'Girdi'; Strict = $true; Miras = $true },
    @{ K = 'SASAdi';      L = 'SAS Name';                              T = 'text'; Bagli = 'SAS'; Miras = $true },
    @{ K = 'Termal';      L = 'Thermal';                               T = 'combo'; List = 'Girdi'; Strict = $true; Miras = $true },
    @{ K = 'TermalAdi';   L = 'Thermal Name';                          T = 'text'; Bagli = 'Termal'; Miras = $true },
    @{ K = 'AnalizTuru';  L = 'Analysis Type';                         T = 'combo'; List = 'AnalizTuru'; Miras = $true },
    @{ K = 'Oncelik';     L = 'Priority';                              T = 'combo'; List = 'Oncelik'; Strict = $true },
    @{ K = 'TalepEden';   L = 'Requester';                             T = 'combo'; List = 'TalepEden'; Strict = $true },
    @{ K = 'AnalizYapan'; L = 'Analyst';                               T = 'combo'; List = 'AnalizYapan'; Strict = $true },
    @{ K = 'TalepTarihi'; L = 'Request Date';                          T = 'date' },
    @{ K = 'Deadline';    L = 'Due Date';                              T = 'date' },
    @{ K = 'TalepNotu';   L = 'Requester Notes';                       T = 'memo' },
    @{ K = 'Durum';       L = 'Status';                                T = 'combo'; List = 'Durum'; Strict = $true; SadeceDuzenle = $true },
    @{ K = 'YapanNotu';   L = 'Analyst Notes';                         T = 'memo'; SadeceDuzenle = $true },
    @{ K = 'ModelYolu';   L = 'Model File Path';                       T = 'text'; SadeceDuzenle = $true },
    @{ K = 'SonucYolu';   L = 'Results / Report Path';                 T = 'text'; SadeceDuzenle = $true },
    @{ K = 'Tamamlanma';  L = 'Completion Date';                       T = 'date'; SadeceDuzenle = $true },
    @{ K = 'Gecmis';      L = 'Change History';                        T = 'hist'; SadeceDuzenle = $true }
)
# analizcinin dolduracagi alanlar (tasarimci duzenleyemez)
foreach ($a in $script:Alanlar) { if ($a.K -in @('Durum', 'YapanNotu', 'ModelYolu', 'SonucYolu', 'Tamamlanma')) { $a.Analizci = $true } }
# form bolum basliklari (bolumun ilk alani)
$script:Bolumler = @{ Proje = 'Project and part'; ModelRolu = 'Model structure'; ModelTipi = 'Model and analysis definition'; Amac = 'Purpose and load condition'; Aero = 'Input data'; AnalizTuru = 'Request details'; Durum = 'Analyst section' }
foreach ($a in $script:Alanlar) { if ($script:Bolumler.ContainsKey($a.K)) { $a.Bolum = $script:Bolumler[$a.K] } }

function Get-VarsayilanListeler {
    [ordered]@{
        Durum      = @('Pending', 'Model Preparation', 'In Progress', 'Review / Check', 'On Hold', 'Completed', 'Cancelled')
        Oncelik    = @('High', 'Medium', 'Low')
        ModelTipi  = @('2D', '3D')
        AnalizTipi = @('Elastic', 'Elastoplastic')
        Alpha      = @('On', 'Off')
        ModelRolu  = @('Standalone', 'Master Model', 'Sub-Analysis')
        Girdi      = @('Not Required', 'Waiting', 'Ready')
        Malzeme    = @('IN718', 'Hastelloy X', 'Waspaloy', 'Ti-6Al-4V', '17-4PH', 'Other')
        AnalizTuru = @('Static Structural', 'Prestressed Modal', 'Campbell / SAFE', 'Fatigue (LCF/HCF)', 'Creep', 'Thermal-Structural', 'Contact / Bolt', 'Cyclic Symmetry', 'Other')
        Kosul      = @('Take-off', 'Max Continuous', 'Cruise', 'Idle', 'Overspeed', 'Burst', 'Transient', 'Other')
        TalepEden  = @()
        AnalizYapan = @()
    }
}

# ---------------------------------------------------------------- dosya yardimcilari
function Write-DosyaAtomik {
    param([string]$Yol, [string]$Metin)
    $tmp = $Yol + '.' + [guid]::NewGuid().ToString('N') + '.tmp'
    [IO.File]::WriteAllText($tmp, $Metin, $script:Utf8)
    if (Test-Path -LiteralPath $Yol) { [IO.File]::Replace($tmp, $Yol, [NullString]::Value) }
    else { [IO.File]::Move($tmp, $Yol) }
}

function Enter-Kilit {
    # kayit basina kilit: FileShare.None + DeleteOnClose (surec olse bile birakilir)
    param([string]$Ad)
    $kp = Join-Path $script:KayitDir ($Ad + '.lock')
    for ($i = 0; $i -lt 60; $i++) {
        try {
            return (New-Object IO.FileStream($kp, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None, 1, [IO.FileOptions]::DeleteOnClose))
        } catch [IO.IOException], [UnauthorizedAccessException] {
            # Windows'ta silinmek uzere olan kilit dosyasi UnauthorizedAccess verebilir: beklenip tekrar denenir
            Start-Sleep -Milliseconds 100
        }
    }
    throw 'The record is currently being saved by another user. Please try again in a few seconds.'
}

function Initialize-Veri {
    param([string]$Kok)
    $script:Kok = $Kok
    $script:KayitDir = Join-Path $Kok 'Kayitlar'
    $script:SilinenDir = Join-Path $Kok 'Silinenler'
    $script:KullaniciDir = Join-Path $Kok 'Kullanicilar'
    $script:BildirimDir = Join-Path $Kok 'Bildirimler'
    $script:YedekDir = Join-Path $Kok 'Yedek'
    foreach ($d in @($Kok, $script:KayitDir, $script:SilinenDir, $script:KullaniciDir, $script:BildirimDir, $script:YedekDir)) {
        if (-not (Test-Path -LiteralPath $d)) { [void](New-Item -ItemType Directory -Path $d -Force) }
    }
    if (-not (Test-Path -LiteralPath (Join-Path $Kok 'Listeler.json'))) { Save-Listeler (Get-VarsayilanListeler) }
}

# ---------------------------------------------------------------- listeler
function Get-Listeler {
    $vars = Get-VarsayilanListeler
    $sonuc = @{}
    foreach ($k in $vars.Keys) { $sonuc[$k] = @($vars[$k]) }
    $yol = Join-Path $script:Kok 'Listeler.json'
    try {
        $o = [IO.File]::ReadAllText($yol, [Text.Encoding]::UTF8) | ConvertFrom-Json
        foreach ($k in @($vars.Keys)) {
            $v = $o.$k
            if ($null -ne $v) { $sonuc[$k] = @($v | ForEach-Object { Convert-Deger ([string]$_) }) }
        }
    } catch { }
    # sabit listeler dosyadan degismesin
    $sonuc['ModelTipi'] = @($vars['ModelTipi']); $sonuc['AnalizTipi'] = @($vars['AnalizTipi']); $sonuc['Alpha'] = @($vars['Alpha']); $sonuc['ModelRolu'] = @($vars['ModelRolu']); $sonuc['Girdi'] = @($vars['Girdi'])
    foreach ($kp in $script:Kapali) { if ($sonuc['Durum'] -notcontains $kp) { $sonuc['Durum'] = @($sonuc['Durum']) + $kp } }
    return $sonuc
}

function Save-Listeler {
    param($Listeler)
    $h = [ordered]@{}
    foreach ($k in $Listeler.Keys) { $h[$k] = @($Listeler[$k] | Where-Object { $_ -ne $null -and ([string]$_).Trim() -ne '' } | ForEach-Object { ([string]$_).Trim() }) }
    Write-DosyaAtomik (Join-Path $script:Kok 'Listeler.json') (ConvertTo-Json $h -Depth 4)
}

# ---------------------------------------------------------------- kayit nesnesi
function ConvertTo-Kayit {
    param($O)
    $h = [ordered]@{}
    foreach ($a in $script:Alanlar) {
        if ($a.T -ne 'hist') {
            $v = $O.($a.K)
            if ($v -is [datetime]) { $v = $v.ToString('yyyy-MM-dd', $script:Invar) }
            $h[$a.K] = if ($null -eq $v) { '' } else { [string]$v }
            if ($a.K -in @('Durum', 'Oncelik', 'ModelRolu', 'Aero', 'SAS', 'Termal', 'Malzeme', 'Kosul', 'AnalizTuru', 'AnalizTipi', 'Alpha', 'ModelTipi')) { $h[$a.K] = Convert-Deger $h[$a.K] }
        }
    }
    $h['Olusturan'] = if ($null -eq $O.Olusturan) { '' } else { [string]$O.Olusturan }
    $h['Revizyon'] = if ($null -eq $O.Revizyon) { 0 } else { [int]$O.Revizyon }
    $h['SonGuncelleyen'] = if ($null -eq $O.SonGuncelleyen) { '' } else { [string]$O.SonGuncelleyen }
    $sg = $O.SonGuncelleme
    if ($sg -is [datetime]) { $sg = $sg.ToString('s', $script:Invar) }
    $h['SonGuncelleme'] = if ($null -eq $sg) { '' } else { [string]$sg }
    $g = @()
    if ($null -ne $O.Gecmis) { $g = @($O.Gecmis | ForEach-Object { [string]$_ }) }
    $h['Gecmis'] = $g
    return [pscustomobject]$h
}

function New-BosKayit { return (ConvertTo-Kayit ([pscustomobject]@{})) }

function Read-Kayit {
    param([string]$Yol)
    for ($i = 0; $i -lt 4; $i++) {
        try {
            $t = [IO.File]::ReadAllText($Yol, [Text.Encoding]::UTF8)
            $o = $t | ConvertFrom-Json
            return (ConvertTo-Kayit $o)
        } catch { Start-Sleep -Milliseconds 120 }
    }
    return $null
}

function Get-TumKayitlar {
    $liste = New-Object System.Collections.ArrayList
    foreach ($f in [IO.Directory]::GetFiles($script:KayitDir, '*.json')) {
        $r = Read-Kayit $f
        if ($null -ne $r -and $r.ID) { [void]$liste.Add($r) }
    }
    return ($liste.ToArray() | Sort-Object ID)
}

function Get-VeriImzasi {
    # klasorde bir sey degisti mi? (sayi + son yazma zamanlari)
    $di = New-Object IO.DirectoryInfo($script:KayitDir)
    $n = 0; $t = [int64]0
    foreach ($f in $di.GetFiles('*.json')) { $n++; $t += $f.LastWriteTimeUtc.Ticks % 1000000007 }
    return "$n|$t"
}

# ---------------------------------------------------------------- tarih / hesap
function ConvertTo-Tarih {
    param([string]$S)
    if ([string]::IsNullOrEmpty($S)) { return $null }
    try { return [datetime]::ParseExact($S, 'yyyy-MM-dd', $script:Invar) } catch { return $null }
}
function ConvertFrom-Tarih { param($D) if ($null -eq $D) { return '' } else { return ([datetime]$D).ToString('yyyy-MM-dd', $script:Invar) } }

function Test-Kapali { param($R) return ($script:Kapali -contains $R.Durum) }

function Get-KalanGun {
    param($R)
    if (Test-Kapali $R) { return $null }
    $d = ConvertTo-Tarih $R.Deadline
    if ($null -eq $d) { return $null }
    return [int](($d.Date - (Get-Date).Date).TotalDays)
}

function Test-Gecikti {
    param($R)
    $k = Get-KalanGun $R
    return ($null -ne $k -and $k -lt 0)
}

# ---------------------------------------------------------------- degisiklik gecmisi
function Get-Degisiklikler {
    param($Eski, $Yeni)
    $satirlar = @()
    foreach ($a in $script:Alanlar) {
        if ($a.T -in @('ro', 'hist')) { continue }
        $e = [string]$Eski.($a.K); $y = [string]$Yeni.($a.K)
        if ($e -ne $y) {
            if ($e.Length -gt 50) { $e = $e.Substring(0, 50) + '...' }
            if ($y.Length -gt 50) { $y = $y.Substring(0, 50) + '...' }
            $satirlar += ('{0}: "{1}" -> "{2}"' -f $a.L.Replace(' *', ''), $e, $y)
        }
    }
    return $satirlar
}

function Get-KullaniciAdi { if ($null -ne $script:Ben) { return $script:Ben.Ad } else { return [Environment]::UserName } }

# ---------------------------------------------------------------- kullanicilar / roller / yetki
$script:Roller = @('Leader', 'Design Engineer', 'Structural Analysis Engineer')
$script:Ben = $null
$script:HesapOverride = $null
$script:BildirimOnbellek = @{}

function Get-HesapAdi {
    if ($script:HesapOverride) { return ([string]$script:HesapOverride).ToLower() }
    return ('{0}\{1}' -f [Environment]::UserDomainName, [Environment]::UserName).ToLower()
}

function ConvertTo-ZamanMetni {
    param($V)
    if ($null -eq $V) { return '' }
    if ($V -is [datetime]) { return $V.ToString('s', $script:Invar) }
    return [string]$V
}

function Get-KullaniciDosyasi {
    param([string]$Hesap)
    $g = ($Hesap.ToLower() -replace '[^a-z0-9._-]', '_')
    return (Join-Path $script:KullaniciDir ($g + '.json'))
}

function ConvertTo-Kullanici {
    param($O)
    $adm = $false
    if ($null -ne $O.Admin) { try { $adm = [bool]$O.Admin } catch { $adm = $false } }
    return [pscustomobject]([ordered]@{
        Hesap = ([string]$O.Hesap).ToLower(); Ad = [string]$O.Ad; Rol = [string]$O.Rol; Admin = $adm
        Bilgisayar = [string]$O.Bilgisayar; Kayit = (ConvertTo-ZamanMetni $O.Kayit)
        SonGiris = (ConvertTo-ZamanMetni $O.SonGiris); SonGorulen = (ConvertTo-ZamanMetni $O.SonGorulen)
    })
}

function Get-Kullanici {
    # dosya yoksa $null; dosya var ama okunamiyorsa hata verir (gecici ag sorunu 'kullanici silindi' sanilmasin)
    param([string]$Hesap)
    if (-not (Test-Path -LiteralPath $script:KullaniciDir)) { throw 'The user folder cannot be reached.' }
    $yol = Get-KullaniciDosyasi $Hesap
    if (-not (Test-Path -LiteralPath $yol)) { return $null }
    $son = $null
    for ($i = 0; $i -lt 4; $i++) {
        try {
            $o = [IO.File]::ReadAllText($yol, [Text.Encoding]::UTF8) | ConvertFrom-Json
            $k = ConvertTo-Kullanici $o
            if ($k.Hesap -ne $Hesap.ToLower()) { return $null }
            return $k
        } catch { $son = $_; Start-Sleep -Milliseconds 120 }
    }
    throw $son
}

function Get-TumKullanicilar {
    $liste = New-Object System.Collections.ArrayList
    foreach ($f in [IO.Directory]::GetFiles($script:KullaniciDir, '*.json')) {
        try {
            $o = [IO.File]::ReadAllText($f, [Text.Encoding]::UTF8) | ConvertFrom-Json
            $k = ConvertTo-Kullanici $o
            if ($k.Hesap -and $k.Ad) { [void]$liste.Add($k) }
        } catch { }
    }
    return @($liste.ToArray() | Sort-Object Ad)
}

function Save-Kullanici {
    param($K)
    Write-DosyaAtomik (Get-KullaniciDosyasi $K.Hesap) (ConvertTo-Json $K -Depth 3)
}

function Remove-Kullanici {
    param([string]$Hesap)
    $yol = Get-KullaniciDosyasi $Hesap
    if (Test-Path -LiteralPath $yol) { [IO.File]::Delete($yol) }
}

function Test-AdKullanimda {
    param([string]$Ad, [string]$HaricHesap = '')
    foreach ($k in @(Get-TumKullanicilar)) {
        if ($k.Hesap -ne $HaricHesap.ToLower() -and $k.Ad.ToLower() -eq $Ad.Trim().ToLower()) { return $true }
    }
    return $false
}

function New-Kullanici {
    # ilk acilis kaydi. Hic kullanici yoksa ilk kisi yonetici olur.
    param([string]$Hesap, [string]$Ad, [string]$Rol)
    $Ad = ([string]$Ad).Trim()
    if ($Ad.Length -lt 2) { throw 'Please enter your name (at least 2 characters).' }
    if ($script:Roller -notcontains $Rol) { throw 'Please select your role.' }
    if (Test-AdKullanimda $Ad $Hesap) { throw "The name '$Ad' is already registered by another user. Please enter your full name." }
    $ilk = (@(Get-TumKullanicilar).Count -eq 0)
    $simdi = (Get-Date).ToString('s')
    $k = ConvertTo-Kullanici ([pscustomobject]@{ Hesap = $Hesap.ToLower(); Ad = $Ad; Rol = $Rol; Admin = $ilk; Bilgisayar = [Environment]::MachineName; Kayit = $simdi; SonGiris = $simdi; SonGorulen = $simdi })
    Save-Kullanici $k
    return $k
}

function Test-Admin {
    # yonetici: kullanici dosyasindaki Admin isareti veya program klasorundeki AdminAccounts.txt (her satir bir Windows hesabi)
    param($K)
    if ($null -eq $K) { return $false }
    if ($K.Admin) { return $true }
    try {
        if ($script:Klasor) {
            $p = Join-Path $script:Klasor 'AdminAccounts.txt'
            if (Test-Path -LiteralPath $p) {
                foreach ($s in [IO.File]::ReadAllLines($p)) { $t = $s.Trim().ToLower(); if ($t -and -not $t.StartsWith('#') -and $t -eq $K.Hesap) { return $true } }
            }
        }
    } catch { }
    return $false
}

function Test-TamYetki {
    # Leader, Structural Analysis Engineer ve Admin her kaydi duzenleyebilir
    param($K)
    if ($null -eq $K) { return $false }
    return ((Test-Admin $K) -or $K.Rol -eq 'Leader' -or $K.Rol -eq 'Structural Analysis Engineer')
}

function Get-Yetki {
    # Mod: Full (tum alanlar) | Talep (yalniz talep alanlari, kendi kaydi) | Gor (salt okunur)
    param($Rec, $K)
    $y = [pscustomobject]@{ Mod = 'Gor'; Sil = $false }
    if ($null -eq $K) { return $y }
    if (Test-TamYetki $K) { $y.Mod = 'Full'; $y.Sil = $true; return $y }
    $benim = ((-not [string]::IsNullOrEmpty($Rec.Olusturan) -and $Rec.Olusturan -eq $K.Hesap) -or (-not [string]::IsNullOrEmpty($Rec.TalepEden) -and $Rec.TalepEden -eq $K.Ad))
    if ($benim -and -not (Test-Kapali $Rec)) {
        $y.Mod = 'Talep'
        $y.Sil = ($Rec.Durum -eq (Get-Listeler)['Durum'][0])
    }
    return $y
}

# ---------------------------------------------------------------- bildirimler
function New-BildirimDosyasi {
    param([string]$KayitID, [string]$Tur, [string]$Mesaj, [string[]]$Alicilar)
    $al = @($Alicilar | Where-Object { $_ } | Select-Object -Unique)
    if ($al.Count -eq 0) { return }
    $ben = $script:Ben
    $o = [ordered]@{
        Zaman = (Get-Date).ToString('s'); KayitID = $KayitID; Tur = $Tur; Mesaj = $Mesaj
        Kimden = $(if ($ben) { $ben.Ad } else { '' }); KimdenHesap = $(if ($ben) { $ben.Hesap } else { '' }); Alicilar = $al
    }
    $ad = '{0}_{1}.json' -f (Get-Date).ToString('yyyyMMddHHmmssfff'), [guid]::NewGuid().ToString('N').Substring(0, 8)
    Write-DosyaAtomik (Join-Path $script:BildirimDir $ad) (ConvertTo-Json ([pscustomobject]$o) -Depth 3)
}

function Send-Bildirim {
    # Tur: Yeni | Degisti | Silindi
    # Kurallar: Leader hic bildirim almaz; kimse kendi eylemi icin almaz.
    #  - Tasarimci yaparsa: kayitta analizci varsa yalniz ona, yoksa tum Structural Analysis Engineer'lara
    #  - Diger roller yaparsa: talebi olusturana/talep edene; analizci yeni atandiysa ona
    param($Rec, $Eski, [string]$Tur)
    $ben = $script:Ben
    if ($null -eq $ben) { return }
    $kul = @(Get-TumKullanicilar)
    $degis = @()
    if ($Tur -eq 'Degisti') {
        $degis = @(Get-Degisiklikler $Eski $Rec)
        if ($degis.Count -eq 0) { return }
    }
    $analiz = [string]$Rec.AnalizYapan
    $hedef = @()
    if ($ben.Rol -eq 'Design Engineer') {
        if ($analiz) { $hedef = @($kul | Where-Object { $_.Ad -eq $analiz }) }
        if ($hedef.Count -eq 0) { $hedef = @($kul | Where-Object { $_.Rol -eq 'Structural Analysis Engineer' }) }
    } else {
        if ($Rec.Olusturan) { $hedef += @($kul | Where-Object { $_.Hesap -eq $Rec.Olusturan }) }
        if ($Rec.TalepEden) { $hedef += @($kul | Where-Object { $_.Ad -eq $Rec.TalepEden }) }
        $atandi = ($Tur -eq 'Yeni') -or ($null -ne $Eski -and [string]$Eski.AnalizYapan -ne $analiz)
        if ($atandi -and $analiz) { $hedef += @($kul | Where-Object { $_.Ad -eq $analiz }) }
    }
    $alicilar = @($hedef | Where-Object { $_.Hesap -ne $ben.Hesap -and $_.Rol -ne 'Leader' } | ForEach-Object { $_.Hesap })
    if ($alicilar.Count -eq 0) { return }
    $parca = ([string]$Rec.ParcaAdi); if ($parca.Length -gt 40) { $parca = $parca.Substring(0, 40) + '...' }
    switch ($Tur) {
        'Yeni' { $m = "New request $($Rec.ID): $parca by $($ben.Ad)" }
        'Silindi' { $m = "$($Rec.ID) ($parca) was deleted by $($ben.Ad)" }
        default {
            $m = "$($Rec.ID) ($parca) updated by $($ben.Ad): " + (($degis | Select-Object -First 2) -join '; ')
            if ($degis.Count -gt 2) { $m += "; +$($degis.Count - 2) more" }
        }
    }
    New-BildirimDosyasi $Rec.ID $Tur $m $alicilar
}

function Get-BenimBildirimler {
    param($K)
    $sonuc = New-Object System.Collections.ArrayList
    foreach ($f in [IO.Directory]::GetFiles($script:BildirimDir, '*.json')) {
        $ad = [IO.Path]::GetFileName($f)
        if (-not $script:BildirimOnbellek.ContainsKey($ad)) {
            try {
                $o = [IO.File]::ReadAllText($f, [Text.Encoding]::UTF8) | ConvertFrom-Json
                $script:BildirimOnbellek[$ad] = [pscustomobject]@{ Dosya = $ad; Zaman = (ConvertTo-ZamanMetni $o.Zaman); KayitID = [string]$o.KayitID; Tur = [string]$o.Tur; Mesaj = [string]$o.Mesaj; Kimden = [string]$o.Kimden; Alicilar = @($o.Alicilar) }
            } catch { continue }
        }
        $b = $script:BildirimOnbellek[$ad]
        if ($b.Alicilar -contains $K.Hesap) { [void]$sonuc.Add($b) }
    }
    $liste = @($sonuc.ToArray() | Sort-Object Zaman -Descending)
    foreach ($b in $liste) { $b | Add-Member -NotePropertyName Okunmadi -NotePropertyValue ([string]::CompareOrdinal($b.Zaman, [string]$K.SonGorulen) -gt 0) -Force }
    return $liste
}

function Set-BildirimOkundu {
    param($K)
    $u = Get-Kullanici $K.Hesap
    if ($null -eq $u) { return $K }
    $u.SonGorulen = (Get-Date).ToString('s')
    Save-Kullanici $u
    return $u
}

function Remove-EskiBildirimler {
    # 30 gunden eski bildirim dosyalarini siler
    $sinir = (Get-Date).AddDays(-30).ToString('yyyyMMddHHmmss')
    foreach ($f in [IO.Directory]::GetFiles($script:BildirimDir, '*.json')) {
        $ad = [IO.Path]::GetFileName($f)
        if ($ad.Length -gt 14 -and [string]::CompareOrdinal($ad.Substring(0, 14), $sinir) -lt 0) { try { [IO.File]::Delete($f) } catch { } }
    }
}

# ---------------------------------------------------------------- yonetici islemleri
function Update-KayitlardaAd {
    # kayitlardaki bir kisi adini topluca degistirir (analizci devri, kisi yeniden adlandirma)
    param([string]$Eski, [string]$Yeni, [string[]]$Anahtarlar, [bool]$SadeceAcik)
    $ok = 0; $hata = 0
    foreach ($r in @(Get-TumKayitlar)) {
        if ($SadeceAcik -and (Test-Kapali $r)) { continue }
        $n = ConvertTo-Kayit $r; $degisti = $false
        foreach ($k in $Anahtarlar) { if ($n.$k -eq $Eski) { $n.$k = $Yeni; $degisti = $true } }
        if (-not $degisti) { continue }
        try { [void](Save-Kayit $n $r); $ok++ } catch { $hata++ }
    }
    return [pscustomobject]@{ Guncellenen = $ok; Hatali = $hata }
}

function Get-SilinenKayitlar {
    $liste = New-Object System.Collections.ArrayList
    foreach ($f in [IO.Directory]::GetFiles($script:SilinenDir, '*.json')) {
        $ad = [IO.Path]::GetFileNameWithoutExtension($f)
        if ($ad -match '^(AN-\d+)_(\d{14})$') {
            $r = Read-Kayit $f
            if ($null -eq $r) { continue }
            $z = [datetime]::ParseExact($Matches[2], 'yyyyMMddHHmmss', $script:Invar)
            [void]$liste.Add([pscustomobject]@{ Dosya = $f; ID = $Matches[1]; Silinme = $z; Parca = $r.ParcaAdi; Proje = $r.Proje; Talep = $r.TalepEden })
        }
    }
    return @($liste.ToArray() | Sort-Object Silinme -Descending)
}

function Restore-Kayit {
    param([string]$Dosya, [string]$Id)
    $hedef = Join-Path $script:KayitDir ($Id + '.json')
    if (Test-Path -LiteralPath $hedef) { throw "A record with ID $Id already exists." }
    [IO.File]::Move($Dosya, $hedef)
}

function Backup-Veri {
    # kayitlar, silinenler, kullanicilar ve listeler Yedek\<tarih> altina kopyalanir; son 10 yedek tutulur
    $hedef = Join-Path $script:YedekDir (Get-Date).ToString('yyyyMMdd_HHmmss')
    [void](New-Item -ItemType Directory -Path $hedef -Force)
    $n = 0
    foreach ($d in 'Kayitlar', 'Silinenler', 'Kullanicilar') {
        $kd = Join-Path $hedef $d
        [void](New-Item -ItemType Directory -Path $kd -Force)
        foreach ($f in [IO.Directory]::GetFiles((Join-Path $script:Kok $d), '*.json')) {
            try { [IO.File]::Copy($f, (Join-Path $kd ([IO.Path]::GetFileName($f)))); $n++ } catch { }
        }
    }
    $lp = Join-Path $script:Kok 'Listeler.json'
    if (Test-Path -LiteralPath $lp) { [IO.File]::Copy($lp, (Join-Path $hedef 'Listeler.json')) }
    $eskiler = @(Get-ChildItem -LiteralPath $script:YedekDir -Directory | Sort-Object Name)
    if ($eskiler.Count -gt 10) { foreach ($e in $eskiler[0..($eskiler.Count - 11)]) { try { Remove-Item -LiteralPath $e.FullName -Recurse -Force } catch { } } }
    return [pscustomobject]@{ Yol = $hedef; Dosya = $n }
}

# ---------------------------------------------------------------- kaydet / sil
function New-KayitDosyasi {
    # yeni ID atar; dosyayi CreateNew ile olusturur (iki kisi ayni ID'yi alamaz)
    param($Rec)
    $max = 0
    foreach ($f in [IO.Directory]::GetFiles($script:KayitDir, '*.json')) {
        if ([IO.Path]::GetFileNameWithoutExtension($f) -match '^AN-(\d+)$') { $v = [int]$Matches[1]; if ($v -gt $max) { $max = $v } }
    }
    foreach ($f in [IO.Directory]::GetFiles($script:SilinenDir, '*.json')) {
        if ([IO.Path]::GetFileName($f) -match '^AN-(\d+)_') { $v = [int]$Matches[1]; if ($v -gt $max) { $max = $v } }
    }
    for ($i = 1; $i -le 100; $i++) {
        $id = 'AN-{0:0000}' -f ($max + $i)
        $yol = Join-Path $script:KayitDir ($id + '.json')
        $fs = $null
        try { $fs = New-Object IO.FileStream($yol, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None) }
        catch [IO.IOException], [UnauthorizedAccessException] { continue }
        try {
            $Rec.ID = $id
            $bytes = $script:Utf8.GetBytes((ConvertTo-Json $Rec -Depth 4))
            $fs.Write($bytes, 0, $bytes.Length)
        } finally { $fs.Dispose() }
        return $id
    }
    throw 'Could not assign a new ID.'
}

function Save-Kayit {
    # Yeni kayit: $Eski = $null.  Mevcut kayit: $Eski = acilirken okunan kopya (cakisma kontrolu icin)
    param($Rec, $Eski)
    $simdi = (Get-Date).ToString('s')
    $kul = Get-KullaniciAdi
    $Rec.SonGuncelleyen = $kul
    $Rec.SonGuncelleme = $simdi
    $zaman = (Get-Date).ToString('yyyy-MM-dd HH:mm')
    if ($null -eq $Eski) {
        $Rec.Olusturan = Get-HesapAdi
        $Rec.Revizyon = 1
        $Rec.Gecmis = @("$zaman | $kul | Record created")
        return (New-KayitDosyasi $Rec)
    }
    $yol = Join-Path $script:KayitDir ($Rec.ID + '.json')
    $kilit = Enter-Kilit $Rec.ID
    try {
        $mevcut = Read-Kayit $yol
        if ($null -eq $mevcut) { throw 'Record file not found (someone else may have deleted it).' }
        if ($mevcut.Revizyon -ne $Eski.Revizyon) { throw 'CAKISMA' }
        $Rec.Revizyon = $mevcut.Revizyon + 1
        $yeniSatirlar = @(Get-Degisiklikler $mevcut $Rec | ForEach-Object { "$zaman | $kul | $_" })
        $g = @($mevcut.Gecmis) + $yeniSatirlar
        if ($g.Count -gt 80) { $g = $g[($g.Count - 80)..($g.Count - 1)] }
        $Rec.Gecmis = @($g)
        Write-DosyaAtomik $yol (ConvertTo-Json $Rec -Depth 4)
    } finally { $kilit.Dispose() }
    return $Rec.ID
}

function Remove-Kayit {
    param([string]$Id)
    $yol = Join-Path $script:KayitDir ($Id + '.json')
    $kilit = Enter-Kilit $Id
    try {
        if (-not (Test-Path -LiteralPath $yol)) { return }
        $hedef = Join-Path $script:SilinenDir ('{0}_{1}.json' -f $Id, (Get-Date).ToString('yyyyMMddHHmmss'))
        [IO.File]::Move($yol, $hedef)
    } finally { $kilit.Dispose() }
}

# ---------------------------------------------------------------- dogrulama
function Get-Cocuklar {
    param($Tum, [string]$Id)
    return @($Tum | Where-Object { $_.ModelRolu -eq 'Sub-Analysis' -and $_.AnaModelID -eq $Id })
}

function Test-Kayit {
    # hata mesajlari dizisi dondurur (bos = tamam)
    param($Rec, $Tum, $Eski)
    $h = @()
    if ([string]::IsNullOrWhiteSpace($Rec.ParcaAdi)) { $h += 'Part name cannot be empty.' }
    if ([string]::IsNullOrWhiteSpace($Rec.AnalizTipi)) { $h += 'Select a material model (Elastic or Elastoplastic). Use Duplicate to request the same analysis with another option.' }
    $dt1 = ConvertTo-Tarih $Rec.TalepTarihi; $dt2 = ConvertTo-Tarih $Rec.Deadline
    if ($null -ne $dt1 -and $null -ne $dt2 -and $dt2.Date -lt $dt1.Date) { $h += 'Due date cannot be earlier than the request date.' }
    if ($Rec.ModelRolu -eq 'Sub-Analysis') {
        if ([string]::IsNullOrWhiteSpace($Rec.AnaModelID)) { $h += "A Sub-Analysis must have a Linked Master Model." }
        elseif ($Rec.AnaModelID -eq $Rec.ID) { $h += 'A record cannot be its own sub-analysis.' }
        else {
            $m = @($Tum | Where-Object { $_.ID -eq $Rec.AnaModelID })
            if ($m.Count -eq 0) { $h += "Linked master model not found: $($Rec.AnaModelID)" }
            elseif ($m[0].ModelRolu -ne 'Master Model') { $h += "Record $($Rec.AnaModelID) does not have the Model Role 'Master Model'." }
        }
    }
    if ($null -ne $Eski -and $Eski.ModelRolu -eq 'Master Model' -and $Rec.ModelRolu -ne 'Master Model') {
        $c = Get-Cocuklar $Tum $Rec.ID
        if ($c.Count -gt 0) { $h += "This record is the master model of $($c.Count) sub-analysis(es). Link them to another master model or delete them before changing its role." }
    }
    return $h
}

# ---------------------------------------------------------------- ozet / dis aktarim
function Get-Ozet {
    param($Tum)
    $acik = @($Tum | Where-Object { -not (Test-Kapali $_) })
    $o = [ordered]@{}
    $o.Toplam = $Tum.Count
    $o.Acik = $acik.Count
    $o.Tamam = @($Tum | Where-Object { $_.Durum -eq 'Completed' }).Count
    $o.Geciken = @($acik | Where-Object { Test-Gecikti $_ }).Count
    $o.Yakin7 = @($acik | Where-Object { $k = Get-KalanGun $_; $null -ne $k -and $k -ge 0 -and $k -le 7 }).Count
    $o.AnaModel = @($Tum | Where-Object { $_.ModelRolu -eq 'Master Model' }).Count
    $o.AltAnaliz = @($Tum | Where-Object { $_.ModelRolu -eq 'Sub-Analysis' }).Count
    $o.Durumlar = @($Tum | Group-Object Durum | Sort-Object Name | ForEach-Object { [pscustomobject]@{ Ad = $_.Name; Adet = $_.Count } })
    $o.Kisiler = @($acik | Group-Object AnalizYapan | Sort-Object Name | ForEach-Object {
        [pscustomobject]@{ Ad = $(if ($_.Name) { $_.Name } else { '(unassigned)' }); Acik = $_.Count; Geciken = @($_.Group | Where-Object { Test-Gecikti $_ }).Count }
    })
    $o.Yaklasan = @($acik | Where-Object { $null -ne (Get-KalanGun $_) } | Sort-Object { Get-KalanGun $_ } | Select-Object -First 10)
    return [pscustomobject]$o
}

function Export-KayitlarCsv {
    param($Tum, [string]$Yol)
    $satir = foreach ($r in $Tum) {
        $s = [ordered]@{}
        foreach ($a in $script:Alanlar) {
            if ($a.T -eq 'hist') { continue }
            $v = [string]$r.($a.K)
            if ($a.T -eq 'date' -and $v) { $d = ConvertTo-Tarih $v; if ($d) { $v = $d.ToString('dd.MM.yyyy') } }
            $s[$a.L.Replace(' *', '').Replace(' (auto)', '')] = $v
        }
        $s['Days Left'] = Get-KalanGun $r
        $s['Overdue'] = if (Test-Gecikti $r) { 'OVERDUE' } else { '' }
        $s['Last Updated By'] = $r.SonGuncelleyen
        [pscustomobject]$s
    }
    @($satir) | Export-Csv -Path $Yol -NoTypeInformation -Delimiter ';' -Encoding UTF8
}

function Add-OrnekVeri {
    # Ornek veri: 8 ana model (her birine 2-3 alt analiz) + 4 bagimsiz analiz.
    # Kisi adlari kayitli kullanicilardan alinir (tasarimci / yapisalci); yoksa genel adlar kullanilir.
    $bugun = Get-Date
    $kul = @(Get-TumKullanicilar)
    $tas = @($kul | Where-Object { $_.Rol -eq 'Design Engineer' } | ForEach-Object { $_.Ad })
    $ana = @($kul | Where-Object { $_.Rol -eq 'Structural Analysis Engineer' } | ForEach-Object { $_.Ad })
    function Kisi($liste, $n, $genel) { if ($liste.Count -eq 0) { return "$genel $n" } else { return $liste[($n - 1) % $liste.Count] } }
    function Yeni-Ornek($ortak, $o) {
        $r = New-BosKayit
        foreach ($k in $ortak.Keys) { $r.$k = $ortak[$k] }
        foreach ($k in $o.Keys) { if (-not $k.StartsWith('_')) { $r.$k = $o[$k] } }
        $r.TalepEden = Kisi $tas $o['_r'] 'Requester'
        if ($o['_a'] -gt 0) { $r.AnalizYapan = Kisi $ana $o['_a'] 'Engineer' }
        $r.TalepTarihi = ConvertFrom-Tarih ($bugun.AddDays($o['_t']))
        $r.Deadline = ConvertFrom-Tarih ($bugun.AddDays($o['_d']))
        if ($o.ContainsKey('_b')) { $r.Tamamlanma = ConvertFrom-Tarih ($bugun.AddDays($o['_b'])) }
        return $r
    }
    # kisa tanim: amac, kosul, hiz, durum, oncelik, talep eden no, analizci no, talep notu, talep gunu, termin gunu, (bitis gunu), (analizci notu)
    function S($amac, $kosul, $hiz, $durum, $onc, $r, $a, $not, $t, $d, $b = $null, $an = '') {
        $h = @{ Amac = $amac; Kosul = $kosul; Hiz = $hiz; Durum = $durum; Oncelik = $onc; _r = $r; _a = $a; TalepNotu = $not; _t = $t; _d = $d }
        if ($null -ne $b) { $h['_b'] = $b }
        if ($an) { $h['YapanNotu'] = $an }
        return $h
    }
    $G = 'Ready'; $B = 'Waiting'; $N = 'Not Required'
    $modeller = @(
        @{ Ortak = @{ Proje = 'Sample Engine'; ParcaAdi = 'HPT Stage 1 Disk'; Layout = 'Layout A'; ModelNo = 'HPT-D1-M01'; ModelTipi = '3D'; AnalizTipi = 'Elastic'; Alpha = 'On'; Malzeme = 'IN718'; AnalizTuru = 'Cyclic Symmetry'; Aero = $G; AeroAdi = 'Aero_Load_RevC'; SAS = $G; SASAdi = 'SAS_Temp_RevB'; Termal = $G; TermalAdi = 'Therm_Map_RevB' }
           Ana = (S 'Fatigue life verification' 'Take-off' '105% speed' 'In Progress' 'High' 1 1 'Aero loads rev C' -20 16 $null 'Cyclic symmetry model set up')
           Alt = @((S 'Overspeed check' 'Overspeed' '120% speed' 'Pending' 'Medium' 1 2 'Speed 120%, temperature field unchanged' -2 30),
                   (S 'Max continuous strength' 'Max Continuous' '100% speed' 'Completed' 'Medium' 2 1 'Only speed and pressure loads changed' -15 -8 -9 'Results delivered, margin OK'),
                   (S 'Burst margin' 'Burst' '122% speed' 'Review / Check' 'High' 3 1 'Check rim burst margin' -9 6)) },
        @{ Ortak = @{ Proje = 'Sample Engine'; ParcaAdi = 'LPT Rotor Blade'; Layout = 'Layout B'; ModelNo = 'LPT-B-M02'; ModelTipi = '3D'; AnalizTipi = 'Elastic'; Alpha = 'Off'; Malzeme = 'Hastelloy X'; AnalizTuru = 'Prestressed Modal'; Aero = $G; AeroAdi = 'LPT_Aero_Pressure_R2'; SAS = $B; SASAdi = 'LPT_SAS_Temp'; Termal = $N; TermalAdi = '' }
           Ana = (S 'Modal analysis for Campbell diagram' 'Cruise' '100% speed' 'Model Preparation' 'High' 2 2 'Waiting for SAS temperatures' -18 -4)
           Alt = @((S 'Modal at idle speed' 'Idle' '60% speed' 'Pending' 'Low' 2 2 '6000 rpm, prestress updated' -3 20),
                   (S 'Overspeed modal check' 'Overspeed' '115% speed' 'In Progress' 'Medium' 1 1 '115% speed, temperature field unchanged' -6 9)) },
        @{ Ortak = @{ Proje = 'Sample Engine 2'; ParcaAdi = 'Labyrinth Seal; Seal Carrier'; Layout = 'Layout A'; ModelNo = 'LAB-M03'; ModelTipi = '2D'; AnalizTipi = 'Elastoplastic'; Alpha = 'On'; Malzeme = 'Hastelloy X; IN718'; AnalizTuru = 'Thermal-Structural'; Aero = $N; AeroAdi = ''; SAS = $G; SASAdi = 'Seal_SAS_Temp'; Termal = $G; TermalAdi = 'Seal_Therm_Map' }
           Ana = (S 'Thermal stress check' 'Transient' '100% speed' 'In Progress' 'Medium' 1 1 'Transient curve attached' -10 12)
           Alt = @((S 'Fast acceleration transient' 'Transient' '100% speed' 'Pending' 'Medium' 2 2 'Thermal transient curve updated' -1 25),
                   (S 'Burst condition' 'Burst' '122% speed' 'Review / Check' 'High' 2 1 'Speed 122%, thermal load removed' -12 5)) },
        @{ Ortak = @{ Proje = 'Sample Engine 2'; ParcaAdi = 'HPC Stage 3 Disk'; Layout = 'Layout C'; ModelNo = 'HPC-D3-M04'; ModelTipi = '3D'; AnalizTipi = 'Elastic'; Alpha = 'On'; Malzeme = 'Waspaloy'; AnalizTuru = 'Cyclic Symmetry'; Aero = $G; AeroAdi = 'HPC_Aero_Load_R1'; SAS = $G; SASAdi = 'HPC_SAS_R3'; Termal = $G; TermalAdi = 'HPC_Therm_R2' }
           Ana = (S 'LCF life assessment' 'Take-off' '103% speed' 'In Progress' 'High' 3 2 'Use updated thermal map R2' -14 20 $null 'Mesh convergence study running')
           Alt = @((S 'Cruise steady state' 'Cruise' '95% speed' 'Pending' 'Medium' 3 2 'Same mesh as master' -2 28),
                   (S 'Max continuous strength' 'Max Continuous' '100% speed' 'Completed' 'Medium' 1 1 'Margin check only' -25 -12 -13 'Delivered, report attached'),
                   (S 'Emergency shutdown thermal shock' 'Shutdown' '0% speed' 'Pending' 'High' 3 0 'Analyst to be assigned' -1 35)) },
        @{ Ortak = @{ Proje = 'Sample Engine 2'; ParcaAdi = 'Turbine Casing'; Layout = 'Layout A'; ModelNo = 'TCS-M05'; ModelTipi = '3D'; AnalizTipi = 'Elastic'; Alpha = 'Off'; Malzeme = 'Hastelloy X'; AnalizTuru = 'Static Structural'; Aero = $G; AeroAdi = 'TCS_Pressure_R4'; SAS = $N; SASAdi = ''; Termal = $G; TermalAdi = 'TCS_Therm_R1' }
           Ana = (S 'Casing containment stiffness' 'Take-off' '105% speed' 'On Hold' 'Low' 2 1 'On hold until design freeze' -30 40)
           Alt = @((S 'Pressure + thermal combined' 'Transient' '100% speed' 'On Hold' 'Low' 2 1 'Linked to master hold' -22 45)) },
        @{ Ortak = @{ Proje = 'Sample Engine 3'; ParcaAdi = 'Bolted Flange Joint'; Layout = 'Layout A'; ModelNo = 'BFJ-M06'; ModelTipi = '3D'; AnalizTipi = 'Elastoplastic'; Alpha = 'Off'; Malzeme = 'IN718; A286'; AnalizTuru = 'Contact / Bolt'; Aero = $N; AeroAdi = ''; SAS = $N; SASAdi = ''; Termal = $G; TermalAdi = 'Flange_Therm_R1' }
           Ana = (S 'Bolt preload and gapping' 'Assembly' '0% speed' 'In Progress' 'Medium' 3 1 'Preload 70% of yield' -8 14 $null 'Contact convergence OK')
           Alt = @((S 'Operating load case' 'Take-off' '105% speed' 'Pending' 'Medium' 3 2 'Add axial load from engine mount' -2 22),
                   (S 'Bolt relaxation after hold' 'Creep hold' '100% speed' 'Pending' 'Low' 1 2 'Creep data from material group' -1 33)) },
        @{ Ortak = @{ Proje = 'Sample Engine 3'; ParcaAdi = 'Blisk Stage 1'; Layout = 'Layout B'; ModelNo = 'BLK-M07'; ModelTipi = '3D'; AnalizTipi = 'Elastic'; Alpha = 'On'; Malzeme = 'Ti-6Al-4V'; AnalizTuru = 'Prestressed Modal'; Aero = $G; AeroAdi = 'Blisk_Aero_R2'; SAS = $G; SASAdi = 'Blisk_SAS_R1'; Termal = $B; TermalAdi = 'Blisk_Therm_Pending' }
           Ana = (S 'Campbell and SAFE diagram modes' 'Cruise' '100% speed' 'Model Preparation' 'High' 1 2 'Need modes up to 20 kHz' -5 10)
           Alt = @((S 'Mistuning sensitivity' 'Cruise' '100% speed' 'Pending' 'Medium' 1 2 'Random mistuning 3 sigma' -1 24),
                   (S 'Idle prestress modes' 'Idle' '60% speed' 'Pending' 'Low' 2 2 'Idle speed prestress only' -1 27)) },
        @{ Ortak = @{ Proje = 'Sample Engine 3'; ParcaAdi = 'Compressor Shaft'; Layout = 'Layout A'; ModelNo = 'SHF-M08'; ModelTipi = '2D'; AnalizTipi = 'Elastic'; Alpha = 'Off'; Malzeme = 'Waspaloy'; AnalizTuru = 'Static Structural'; Aero = $N; AeroAdi = ''; SAS = $N; SASAdi = ''; Termal = $N; TermalAdi = '' }
           Ana = (S 'Torque and bending check' 'Max Torque' '100% speed' 'Completed' 'Medium' 2 1 'Axisymmetric model' -40 -20 -22 'Delivered, no issues')
           Alt = @((S 'Hard landing bending' 'Hard landing' '0% speed' 'Completed' 'High' 2 2 'Extra bending moment case' -30 -15 -16 'Delivered')) }
    )
    foreach ($m in $modeller) {
        $a1 = Yeni-Ornek $m.Ortak $m.Ana
        $a1.ModelRolu = 'Master Model'
        $anaId = Save-Kayit $a1 $null
        foreach ($a in $m.Alt) {
            $alt = Yeni-Ornek $m.Ortak $a
            $alt.ModelRolu = 'Sub-Analysis'; $alt.AnaModelID = $anaId
            [void](Save-Kayit $alt $null)
        }
    }
    $bagimsiz = @(
        @{ Ortak = @{ Proje = 'Sample Engine'; ParcaAdi = 'Fan Blade'; Layout = 'Layout A'; ModelNo = 'FAN-S01'; ModelTipi = '3D'; AnalizTipi = 'Elastic'; Alpha = 'Off'; Malzeme = 'Ti-6Al-4V'; AnalizTuru = 'Modal'; Aero = $G; AeroAdi = 'Fan_Aero_R1'; SAS = $N; SASAdi = ''; Termal = $N; TermalAdi = '' }
           A = (S 'Free-free modal check' 'Idle' '0% speed' 'Completed' 'Low' 1 2 'Test correlation request' -35 -20 -21 'Correlation within 3%') },
        @{ Ortak = @{ Proje = 'Sample Engine 2'; ParcaAdi = 'Bearing Housing'; Layout = 'Layout B'; ModelNo = 'BRG-S02'; ModelTipi = '3D'; AnalizTipi = 'Elastic'; Alpha = 'Off'; Malzeme = 'IN718'; AnalizTuru = 'Static Structural'; Aero = $N; AeroAdi = ''; SAS = $N; SASAdi = ''; Termal = $G; TermalAdi = 'BRG_Therm_R1' }
           A = (S 'Oil seal region stress' 'Take-off' '105% speed' 'In Progress' 'Medium' 3 1 'Check local yielding' -7 8) },
        @{ Ortak = @{ Proje = 'Sample Engine 3'; ParcaAdi = 'Heat Shield'; Layout = 'Layout A'; ModelNo = 'HSH-S03'; ModelTipi = '2D'; AnalizTipi = 'Elastic'; Alpha = 'On'; Malzeme = 'Hastelloy X'; AnalizTuru = 'Thermal-Structural'; Aero = $N; AeroAdi = ''; SAS = $G; SASAdi = 'HSH_SAS_R2'; Termal = $G; TermalAdi = 'HSH_Therm_R2' }
           A = (S 'Thermal buckling check' 'Transient' '100% speed' 'Pending' 'High' 2 0 'Analyst to be assigned' 0 18) },
        @{ Ortak = @{ Proje = 'Sample Engine 3'; ParcaAdi = 'Turbine Vane Segment'; Layout = 'Layout C'; ModelNo = 'VAN-S04'; ModelTipi = '3D'; AnalizTipi = 'Elastoplastic'; Alpha = 'Off'; Malzeme = 'Hastelloy X'; AnalizTuru = 'Fatigue (LCF/HCF)'; Aero = $G; AeroAdi = 'Vane_Aero_R3'; SAS = $G; SASAdi = 'Vane_SAS_R2'; Termal = $G; TermalAdi = 'Vane_Therm_R2' }
           A = (S 'Trailing edge LCF' 'Take-off' '103% speed' 'Review / Check' 'High' 1 2 'Use new cooling scheme' -16 -2 $null 'Overdue: waiting for temperature review') }
    )
    foreach ($m in $bagimsiz) {
        $r = Yeni-Ornek $m.Ortak $m.A
        $r.ModelRolu = 'Standalone'
        [void](Save-Kayit $r $null)
    }
}
