# =====================================================================
#  ANALİZ TAKİP  -  ayrı pencereli program (PowerShell + WinForms)
#  Excel'deki butondan başlatılır:  powershell -STA -File AnalizTakip.ps1
#  Veriler: <Veri klasörü>\Kayitlar\AN-xxxx.json  (ortak klasör, çok kullanıcılı)
# =====================================================================
param([string]$VeriKlasoru, [string]$HesapAdi)
$ErrorActionPreference = 'Stop'
$script:Klasor = $PSScriptRoot
if (-not $script:Klasor) { $script:Klasor = Split-Path -Parent $MyInvocation.MyCommand.Path }
$script:HataLog = Join-Path $script:Klasor 'AnalizTakip.hata.log'

function Write-HataLog {
    param([string]$Mesaj)
    try { Add-Content -LiteralPath $script:HataLog -Value ("{0} | {1} | {2}" -f (Get-Date).ToString('s'), [Environment]::UserName, $Mesaj) -Encoding UTF8 } catch { }
}

try {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    try { $en = [Globalization.CultureInfo]::GetCultureInfo('en-US'); [Threading.Thread]::CurrentThread.CurrentCulture = $en; [Threading.Thread]::CurrentThread.CurrentUICulture = $en; [Globalization.CultureInfo]::DefaultThreadCurrentCulture = $en; [Globalization.CultureInfo]::DefaultThreadCurrentUICulture = $en } catch { }
    try { [System.Windows.Forms.Application]::EnableVisualStyles() } catch { }
    try { [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false) } catch { }
    . (Join-Path $script:Klasor 'Veri.ps1')
    if ($HesapAdi) { $script:HesapOverride = $HesapAdi }
} catch {
    Write-HataLog $_.Exception.ToString()
    try { Add-Type -AssemblyName System.Windows.Forms; [void][System.Windows.Forms.MessageBox]::Show("The program could not be started:`r`n`r`n" + $_.Exception.Message + "`r`n`r`nDetails: " + $script:HataLog, 'Analysis Tracker', 'OK', 'Error') } catch { }
    exit 1
}

# ---------------------------------------------------------------- genel yardımcılar
$script:Tum = @()
$script:FontKalin = $null
$script:Imza = ''
$script:DlgAcik = $false
$script:D = $null
$script:Gosterilen = @{}
$script:Tray = $null
$script:Kapaniyor = $false

function Mesaj {
    param([string]$Metin, [string]$Simge = 'Information', [string]$Dugme = 'OK')
    return [System.Windows.Forms.MessageBox]::Show($Metin, 'Analysis Tracker', [System.Windows.Forms.MessageBoxButtons]::$Dugme, [System.Windows.Forms.MessageBoxIcon]::$Simge)
}

function Guvenli {
    param([scriptblock]$Is)
    try { & $Is }
    catch {
        Write-HataLog ($_.Exception.ToString() + "`r`n" + $_.InvocationInfo.PositionMessage + "`r`n" + $_.ScriptStackTrace)
        [void](Mesaj $_.Exception.Message 'Error')
    }
}

function Set-CiftTampon {
    # cift tamponlama: yatay kaydirirken yazilarin titremesini/kaybolmasini onler
    param($C)
    try {
        $p = $C.GetType().GetProperty('DoubleBuffered', [Reflection.BindingFlags]'Instance,NonPublic')
        if ($null -ne $p) { $p.SetValue($C, $true, $null) }
    } catch { }
}

function Sec-Temizle {
    # bos alana tiklaninca mavi secimi kaldirir
    param($G)
    try { $G.ClearSelection(); $G.CurrentCell = $null } catch { }
}

function Get-AyarYolu { return (Join-Path $script:Klasor 'AnalizTakip.ayar.json') }

function Get-VeriKlasoru {
    $v = Join-Path $script:Klasor 'Veri'
    try {
        if (Test-Path -LiteralPath (Get-AyarYolu)) {
            $o = [IO.File]::ReadAllText((Get-AyarYolu), [Text.Encoding]::UTF8) | ConvertFrom-Json
            if ($o.VeriKlasoru) { $v = [string]$o.VeriKlasoru }
        }
    } catch { }
    return $v
}

function Set-VeriKlasoru {
    param([string]$Yol)
    $json = ConvertTo-Json ([ordered]@{ VeriKlasoru = $Yol })
    [IO.File]::WriteAllText((Get-AyarYolu), $json, (New-Object Text.UTF8Encoding($false)))
}

function Get-Renk { param([int]$R, [int]$G, [int]$B) return [System.Drawing.Color]::FromArgb($R, $G, $B) }
function Stop-Tekerlek {
    # acilir listelerde fare tekerleği degeri degistirmesin (yalnizca tiklama)
    param($C)
    $C.Add_MouseWheel({ param($s, $e) $e.Handled = $true })
}

function Get-Boyut { param([int]$W, [int]$H) return (New-Object System.Drawing.Size($W, $H)) }

# ---------------------------------------------------------------- ana pencere
$ana = New-Object System.Windows.Forms.Form
$ana.Text = 'Analysis Tracker'
$ana.StartPosition = 'CenterScreen'
$ana.Size = Get-Boyut 1320 790
$ana.MinimumSize = Get-Boyut 1000 600
$ana.Font = New-Object System.Drawing.Font('Segoe UI', 9)

Set-CiftTampon $ana
$tab = New-Object System.Windows.Forms.TabControl
Set-CiftTampon $tab
$tab.Dock = 'Fill'
$tp1 = New-Object System.Windows.Forms.TabPage('Analysis List')
$tp2 = New-Object System.Windows.Forms.TabPage('Master Model Tree')
$tp3 = New-Object System.Windows.Forms.TabPage('Summary')
[void]$tab.TabPages.Add($tp1); [void]$tab.TabPages.Add($tp2); [void]$tab.TabPages.Add($tp3)
foreach ($tpp in @($tp1, $tp2, $tp3)) { Set-CiftTampon $tpp }

# ---- sekme 1: liste
$dt = New-Object System.Data.DataTable
foreach ($c in 'ID', 'Proje', 'Parca', 'Layout', 'ModelNo', 'Tip', 'MalzTipi', 'Alpha', 'Hiz', 'Malzeme', 'Kosul', 'AnalizTuru', 'Rol', 'AnaID', 'Durum', 'Oncelik', 'Yapan', 'Talep', 'Gecikme', 'Guncelleyen') { [void]$dt.Columns.Add($c, [string]) }
[void]$dt.Columns.Add('Deadline', [datetime])
[void]$dt.Columns.Add('Kalan', [int])
[void]$dt.Columns.Add('Alt', [int])
$dt.DefaultView.Sort = 'ID DESC'

$grid = New-Object System.Windows.Forms.DataGridView
Set-CiftTampon $grid
$grid.Dock = 'Fill'
$grid.ReadOnly = $true
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.AllowUserToResizeRows = $false
$grid.SelectionMode = 'FullRowSelect'
$grid.MultiSelect = $false
$grid.AutoGenerateColumns = $false
$grid.RowHeadersVisible = $false
$grid.BackgroundColor = [System.Drawing.Color]::White
$grid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$grid.EnableHeadersVisualStyles = $false
$grid.ColumnHeadersDefaultCellStyle.BackColor = Get-Renk 31 56 100
$grid.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White
$grid.AlternatingRowsDefaultCellStyle.BackColor = Get-Renk 242 246 252
$grid.ColumnHeadersHeight = 30

function Add-Kolon {
    param($G, [string]$Baslik, [string]$Alan, [int]$Genislik, [string]$Bicim = '')
    $c = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $c.HeaderText = $Baslik; $c.Name = $Alan; $c.DataPropertyName = $Alan; $c.Width = $Genislik
    if ($Bicim) { $c.DefaultCellStyle.Format = $Bicim }
    [void]$G.Columns.Add($c)
}
Add-Kolon $grid 'ID' 'ID' 70
Add-Kolon $grid 'Project' 'Proje' 110
Add-Kolon $grid 'Part' 'Parca' 170
Add-Kolon $grid 'Layout' 'Layout' 80
Add-Kolon $grid 'Model ID' 'ModelNo' 85
Add-Kolon $grid '2D/3D' 'Tip' 60
Add-Kolon $grid 'Material Model' 'MalzTipi' 140
Add-Kolon $grid 'Alpha' 'Alpha' 120
Add-Kolon $grid 'Material' 'Malzeme' 110
Add-Kolon $grid 'Condition' 'Kosul' 100
Add-Kolon $grid 'Speed' 'Hiz' 90
Add-Kolon $grid 'Analysis Type' 'AnalizTuru' 130
Add-Kolon $grid 'Role' 'Rol' 100
Add-Kolon $grid 'Master' 'AnaID' 75
Add-Kolon $grid 'Status' 'Durum' 160
Add-Kolon $grid 'Priority' 'Oncelik' 65
Add-Kolon $grid 'Analyst' 'Yapan' 110
Add-Kolon $grid 'Requester' 'Talep' 100
Add-Kolon $grid 'Due Date' 'Deadline' 85 'dd.MM.yyyy'
Add-Kolon $grid 'Days Left' 'Kalan' 65
Add-Kolon $grid 'Overdue' 'Gecikme' 75
Add-Kolon $grid 'Subs' 'Alt' 40
Add-Kolon $grid 'Last Updated By' 'Guncelleyen' 120
$grid.DataSource = $dt

$fp = New-Object System.Windows.Forms.FlowLayoutPanel
$fp.Dock = 'Top'; $fp.Height = 40; $fp.Padding = New-Object System.Windows.Forms.Padding(6, 6, 6, 0)

function Yeni-Etiket { param([string]$Metin, [int]$W) $l = New-Object System.Windows.Forms.Label; $l.Text = $Metin; $l.Width = $W; $l.Height = 24; $l.TextAlign = 'MiddleLeft'; $l.Margin = New-Object System.Windows.Forms.Padding(8, 0, 0, 0); return $l }

$txtAra = New-Object System.Windows.Forms.TextBox; $txtAra.Width = 200; $txtAra.Margin = New-Object System.Windows.Forms.Padding(0, 2, 0, 0)
$cmbDurum = New-Object System.Windows.Forms.ComboBox; $cmbDurum.DropDownStyle = 'DropDownList'; $cmbDurum.Width = 150; $cmbDurum.Margin = New-Object System.Windows.Forms.Padding(0, 2, 0, 0)
$cmbKisi = New-Object System.Windows.Forms.ComboBox; $cmbKisi.DropDownStyle = 'DropDownList'; $cmbKisi.Width = 140; $cmbKisi.Margin = New-Object System.Windows.Forms.Padding(0, 2, 0, 0)
Stop-Tekerlek $cmbDurum; Stop-Tekerlek $cmbKisi
$chkAcik = New-Object System.Windows.Forms.CheckBox; $chkAcik.Text = 'Open only'; $chkAcik.AutoSize = $true; $chkAcik.Margin = New-Object System.Windows.Forms.Padding(12, 3, 0, 0)
$chkGec = New-Object System.Windows.Forms.CheckBox; $chkGec.Text = 'Overdue only'; $chkGec.AutoSize = $true; $chkGec.Margin = New-Object System.Windows.Forms.Padding(8, 3, 0, 0)
$btnTemizle = New-Object System.Windows.Forms.Button; $btnTemizle.Text = 'Clear filter'; $btnTemizle.AutoSize = $true; $btnTemizle.Margin = New-Object System.Windows.Forms.Padding(12, 0, 0, 0)
$fp.Controls.Add((Yeni-Etiket 'Search:' 50)); $fp.Controls.Add($txtAra)
$fp.Controls.Add((Yeni-Etiket 'Status:' 50)); $fp.Controls.Add($cmbDurum)
$fp.Controls.Add((Yeni-Etiket 'Analyst:' 55)); $fp.Controls.Add($cmbKisi)
$fp.Controls.Add($chkAcik); $fp.Controls.Add($chkGec); $fp.Controls.Add($btnTemizle)

$tp1.Controls.Add($grid)
$tp1.Controls.Add($fp)
$pnlYeni = New-Object System.Windows.Forms.Panel; $pnlYeni.Dock = 'Bottom'; $pnlYeni.Height = 52
$tbYeni = New-Object System.Windows.Forms.Button; $tbYeni.Text = '+  New Analysis'; $tbYeni.Width = 160; $tbYeni.Height = 36
$tbYeni.Location = New-Object System.Drawing.Point(8, 8); $tbYeni.FlatStyle = 'Flat'; $tbYeni.BackColor = Get-Renk 84 130 53; $tbYeni.ForeColor = [System.Drawing.Color]::White
$tbYeni.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$pnlYeni.Controls.Add($tbYeni)
$tp1.Controls.Add($pnlYeni)

# ---- sekme 2: ağaç
$lv = New-Object System.Windows.Forms.ListView
Set-CiftTampon $lv
$lv.Dock = 'Fill'; $lv.View = 'Details'; $lv.FullRowSelect = $true; $lv.HideSelection = $false; $lv.MultiSelect = $false
$lv.ShowGroups = $false; $lv.GridLines = $false; $lv.BorderStyle = 'None'; $lv.Font = New-Object System.Drawing.Font('Segoe UI', 9.5)
foreach ($c in @(@('ID', 80), @('Role', 100), @('Part', 210), @('Condition / Purpose', 300), @('Status', 130), @('Priority', 70), @('Analyst', 130), @('Requester', 130), @('Due Date', 90), @('Days Left', 70))) { [void]$lv.Columns.Add($c[0], $c[1]) }
$fpAgac = New-Object System.Windows.Forms.FlowLayoutPanel
$fpAgac.Dock = 'Top'; $fpAgac.Height = 40; $fpAgac.Padding = New-Object System.Windows.Forms.Padding(6, 6, 6, 0)
$txtAgac = New-Object System.Windows.Forms.TextBox; $txtAgac.Width = 220; $txtAgac.Margin = New-Object System.Windows.Forms.Padding(0, 2, 0, 0)
$chkAgac = New-Object System.Windows.Forms.CheckBox; $chkAgac.Text = 'Open only'; $chkAgac.AutoSize = $true; $chkAgac.Margin = New-Object System.Windows.Forms.Padding(12, 3, 0, 0)
$lblAgac = New-Object System.Windows.Forms.Label
$lblAgac.AutoSize = $true; $lblAgac.TextAlign = 'MiddleLeft'; $lblAgac.ForeColor = Get-Renk 100 100 100; $lblAgac.Margin = New-Object System.Windows.Forms.Padding(24, 4, 0, 0)
$lblAgac.Text = 'blue band: master model   |   red: overdue   |   gray: closed   |   double-click to open'
$fpAgac.Controls.Add((Yeni-Etiket 'Search:' 50)); $fpAgac.Controls.Add($txtAgac); $fpAgac.Controls.Add($chkAgac); $fpAgac.Controls.Add($lblAgac)
$tp2.Controls.Add($lv)
$tp2.Controls.Add($fpAgac)

# ---- sekme 3: özet
$kartPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$kartPanel.Dock = 'Top'; $kartPanel.Height = 96; $kartPanel.Padding = New-Object System.Windows.Forms.Padding(6)
$script:Kartlar = @{}
function Yeni-Kart {
    param([string]$Anahtar, [string]$Baslik, $Renk)
    $p = New-Object System.Windows.Forms.Panel
    $p.Width = 150; $p.Height = 80; $p.Margin = New-Object System.Windows.Forms.Padding(4); $p.BackColor = Get-Renk 242 246 252; $p.BorderStyle = 'FixedSingle'
    $b = New-Object System.Windows.Forms.Label; $b.Text = $Baslik; $b.Dock = 'Top'; $b.Height = 24; $b.TextAlign = 'MiddleCenter'; $b.ForeColor = [System.Drawing.Color]::White; $b.BackColor = Get-Renk 31 56 100
    $v = New-Object System.Windows.Forms.Label; $v.Text = '0'; $v.Dock = 'Fill'; $v.TextAlign = 'MiddleCenter'; $v.Font = New-Object System.Drawing.Font('Segoe UI', 20, [System.Drawing.FontStyle]::Bold); $v.ForeColor = $Renk
    $p.Controls.Add($v); $p.Controls.Add($b)
    $kartPanel.Controls.Add($p)
    $script:Kartlar[$Anahtar] = $v
}
$lacivert = Get-Renk 31 56 100; $kirmizi = Get-Renk 192 0 0; $yesil = Get-Renk 84 130 53
Yeni-Kart 'Toplam' 'Total Analyses' $lacivert
Yeni-Kart 'Acik' 'Open Analyses' $lacivert
Yeni-Kart 'Tamam' 'Completed' $yesil
Yeni-Kart 'Geciken' 'Overdue' $kirmizi
Yeni-Kart 'Yakin7' 'Due within 7 Days' (Get-Renk 191 120 0)
Yeni-Kart 'AnaModel' 'Master Models' $lacivert
Yeni-Kart 'AltAnaliz' 'Sub-Analyses' $lacivert

$ozetTablo = New-Object System.Windows.Forms.TableLayoutPanel
$ozetTablo.Dock = 'Fill'; $ozetTablo.ColumnCount = 3; $ozetTablo.RowCount = 1
for ($i = 0; $i -lt 3; $i++) { [void]$ozetTablo.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 33.3))) }

function Yeni-OzetGrid {
    param([string[]]$Basliklar)
    $g = New-Object System.Windows.Forms.DataGridView
    Set-CiftTampon $g
    $g.Dock = 'Fill'; $g.ReadOnly = $true; $g.AllowUserToAddRows = $false; $g.RowHeadersVisible = $false; $g.BackgroundColor = [System.Drawing.Color]::White
    $g.SelectionMode = 'FullRowSelect'; $g.AutoSizeColumnsMode = 'Fill'; $g.EnableHeadersVisualStyles = $false
    $g.ColumnHeadersDefaultCellStyle.BackColor = Get-Renk 31 56 100; $g.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White
    foreach ($b in $Basliklar) { [void]$g.Columns.Add($b, $b) }
    return $g
}
$gDurum = Yeni-OzetGrid @('Status', 'Count')
$gKisi = Yeni-OzetGrid @('Analyst', 'Open', 'Overdue')
$gYakin = Yeni-OzetGrid @('Days Left', 'ID', 'Part', 'Analyst', 'Due Date')
foreach ($pair in @(@($gDurum, 'By status'), @($gKisi, 'By analyst (open items)'), @($gYakin, 'Next 10 due dates (open items)'))) {
    $pn = New-Object System.Windows.Forms.Panel; $pn.Dock = 'Fill'; $pn.Padding = New-Object System.Windows.Forms.Padding(4)
    $lb = New-Object System.Windows.Forms.Label; $lb.Text = $pair[1]; $lb.Dock = 'Top'; $lb.Height = 24; $lb.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
    $pn.Controls.Add($pair[0]); $pn.Controls.Add($lb)
    $ozetTablo.Controls.Add($pn)
}
$tp3.Controls.Add($ozetTablo)

# ---- araç çubuğu ve durum çubuğu
$ts = New-Object System.Windows.Forms.ToolStrip
$ts.GripStyle = 'Hidden'; $ts.Padding = New-Object System.Windows.Forms.Padding(6, 4, 6, 4)
$tbDuzenle = $ts.Items.Add('Edit')
$tbKopya = $ts.Items.Add('Duplicate')
$tbSil = $ts.Items.Add('Delete')
[void]$ts.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
$tbYenile = $ts.Items.Add('Refresh')
$tbExcel = $ts.Items.Add('Export to Excel')
$tbBildirim = $ts.Items.Add('Notifications')
[void]$ts.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
$tbAyar = New-Object System.Windows.Forms.ToolStripDropDownButton('Settings')
$miProfil = $tbAyar.DropDownItems.Add('My profile')
$miListe = $tbAyar.DropDownItems.Add('Edit lists (material, condition, status ...)')
$miAc = $tbAyar.DropDownItems.Add('Open data folder')
$miHak = $tbAyar.DropDownItems.Add('About')
[void]$ts.Items.Add($tbAyar)
$tbAdmin = New-Object System.Windows.Forms.ToolStripDropDownButton('Admin')
$miKull = $tbAdmin.DropDownItems.Add('Manage users ...')
$miSilinen = $tbAdmin.DropDownItems.Add('Restore deleted records ...')
$miDevir = $tbAdmin.DropDownItems.Add('Reassign analyst ...')
$miYedek = $tbAdmin.DropDownItems.Add('Back up data now')
$miKlasor = $tbAdmin.DropDownItems.Add('Change data folder ...')
$miOrnek = $tbAdmin.DropDownItems.Add('Add sample data')
[void]$ts.Items.Add($tbAdmin)
foreach ($b in @($tbDuzenle, $tbKopya, $tbSil, $tbYenile, $tbExcel, $tbBildirim)) { $b.Padding = New-Object System.Windows.Forms.Padding(6, 2, 6, 2) }

$ss = New-Object System.Windows.Forms.StatusStrip
$sbKayit = New-Object System.Windows.Forms.ToolStripStatusLabel; $sbKayit.Text = 'Records: 0'
$sbVeri = New-Object System.Windows.Forms.ToolStripStatusLabel; $sbVeri.Spring = $true; $sbVeri.TextAlign = 'MiddleLeft'
$sbKul = New-Object System.Windows.Forms.ToolStripStatusLabel
$sbZaman = New-Object System.Windows.Forms.ToolStripStatusLabel
[void]$ss.Items.Add($sbKayit); [void]$ss.Items.Add($sbVeri); [void]$ss.Items.Add($sbKul); [void]$ss.Items.Add($sbZaman)

$ana.Controls.Add($tab)
$ana.Controls.Add($ts)
$ana.Controls.Add($ss)

# ---------------------------------------------------------------- yenileme
$script:FiltreKilit = $false

function Get-SeciliID {
    if ($tab.SelectedTab -eq $tp2) {
        if ($lv.SelectedItems.Count -gt 0 -and $lv.SelectedItems[0].Tag) { return [string]$lv.SelectedItems[0].Tag }
        return $null
    }
    if ($tab.SelectedTab -eq $tp1 -and $grid.CurrentRow -ne $null -and $grid.CurrentRow.DataBoundItem -ne $null) {
        return [string]$grid.CurrentRow.DataBoundItem['ID']
    }
    return $null
}

function Yenile-Liste {
    $script:FiltreKilit = $true
    try {
        $secili = $null
        if ($grid.CurrentRow -ne $null -and $grid.CurrentRow.DataBoundItem -ne $null) { $secili = [string]$grid.CurrentRow.DataBoundItem['ID'] }
        $ara = $txtAra.Text.Trim().ToLower()
        $fd = [string]$cmbDurum.SelectedItem; $fk = [string]$cmbKisi.SelectedItem
        $dt.BeginLoadData()
        $dt.Rows.Clear()
        foreach ($r in $script:Tum) {
            if ($fd -and $fd -ne 'All' -and $r.Durum -ne $fd) { continue }
            if ($fk -and $fk -ne 'All' -and $r.AnalizYapan -ne $fk) { continue }
            if ($chkAcik.Checked -and (Test-Kapali $r)) { continue }
            $gec = Test-Gecikti $r
            if ($chkGec.Checked -and -not $gec) { continue }
            if ($ara) {
                $hay = (($r.ID, $r.ParcaAdi, $r.Proje, $r.Layout, $r.Malzeme, $r.Amac, $r.Kosul, $r.AnalizTuru, $r.AnalizYapan, $r.TalepEden, $r.ModelNo, $r.TalepNotu, $r.YapanNotu) -join ' ').ToLower()
                if (-not $hay.Contains($ara)) { continue }
            }
            $row = $dt.NewRow()
            $row['ID'] = $r.ID; $row['Parca'] = $r.ParcaAdi; $row['Proje'] = $r.Proje; $row['Layout'] = $r.Layout; $row['ModelNo'] = $r.ModelNo; $row['Alpha'] = $r.Alpha; $row['Hiz'] = $r.Hiz; $row['MalzTipi'] = $r.AnalizTipi; $row['Tip'] = $r.ModelTipi; $row['Malzeme'] = $r.Malzeme
            $row['Kosul'] = $r.Kosul; $row['AnalizTuru'] = $r.AnalizTuru; $row['Rol'] = $r.ModelRolu; $row['AnaID'] = $r.AnaModelID
            $row['Durum'] = $r.Durum; $row['Oncelik'] = $r.Oncelik; $row['Yapan'] = $r.AnalizYapan; $row['Talep'] = $r.TalepEden
            $d = ConvertTo-Tarih $r.Deadline
            if ($null -ne $d) { $row['Deadline'] = $d } else { $row['Deadline'] = [DBNull]::Value }
            $k = Get-KalanGun $r
            if ($null -ne $k) { $row['Kalan'] = $k } else { $row['Kalan'] = [DBNull]::Value }
            $row['Gecikme'] = $(if ($gec) { 'OVERDUE' } else { '' })
            if ($r.ModelRolu -eq 'Master Model') { $row['Alt'] = @($script:Tum | Where-Object { $_.ModelRolu -eq 'Sub-Analysis' -and $_.AnaModelID -eq $r.ID }).Count } else { $row['Alt'] = [DBNull]::Value }
            $row['Guncelleyen'] = $r.SonGuncelleyen
            $dt.Rows.Add($row)
        }
        $dt.EndLoadData()
        if ($secili) {
            foreach ($gr in $grid.Rows) {
                if ($gr.DataBoundItem -ne $null -and [string]$gr.DataBoundItem['ID'] -eq $secili) { $grid.CurrentCell = $gr.Cells[1]; break }
            }
        }
    } finally { $script:FiltreKilit = $false }
}

function Get-DurumRengi {
    param([string]$D)
    switch ($D) {
        'Pending' { return (Get-Renk 105 105 105) }
        'Model Preparation' { return (Get-Renk 191 120 0) }
        'In Progress' { return (Get-Renk 30 100 200) }
        'Review / Check' { return (Get-Renk 120 60 160) }
        'On Hold' { return (Get-Renk 160 90 40) }
        'Completed' { return (Get-Renk 70 130 50) }
        default { return (Get-Renk 128 128 128) }
    }
}

function Yeni-AgacSatiri {
    # tek satir: ana model (mavi bant, kalin) veya alt/bagimsiz analiz
    param($R, [string]$Rol, [bool]$Ana)
    $dl = ConvertTo-Tarih $R.Deadline
    $kal = Get-KalanGun $R
    $amac = (@($R.Kosul, $R.Amac) | Where-Object { $_ }) -join ' - '
    $parca = [string]$R.ParcaAdi
    if (-not $Ana -and $Rol -eq 'Sub-Analysis') { $parca = ([string][char]0x2514) + '  ' + $parca }
    $it = New-Object System.Windows.Forms.ListViewItem([string]$R.ID)
    $it.Tag = $R.ID
    $it.UseItemStyleForSubItems = $false
    foreach ($t in @($(if ($Ana) { 'Master' } else { $Rol }), $parca, $amac, [string]$R.Durum, [string]$R.Oncelik, [string]$R.AnalizYapan, [string]$R.TalepEden, $(if ($dl) { $dl.ToString('dd.MM.yyyy') } else { '' }), $(if ($null -ne $kal) { [string]$kal } else { '' }))) { [void]$it.SubItems.Add($t) }
    $kapali = (Test-Kapali $R); $gec = (Test-Gecikti $R)
    $kalin = New-Object System.Drawing.Font('Segoe UI', 9.5, [System.Drawing.FontStyle]::Bold)
    $normal = New-Object System.Drawing.Font('Segoe UI', 9.5)
    $zemin = $(if ($Ana) { Get-Renk 226 234 248 } else { [System.Drawing.Color]::White })
    $yazi = $(if ($kapali) { Get-Renk 150 150 150 } elseif ($Ana) { Get-Renk 31 56 140 } else { Get-Renk 40 40 40 })
    $i = 0
    foreach ($si in $it.SubItems) {
        $si.BackColor = $zemin; $si.ForeColor = $yazi; $si.Font = $(if ($Ana) { $kalin } else { $normal })
        if ($i -eq 4 -and -not $kapali) { $si.ForeColor = (Get-DurumRengi $R.Durum); $si.Font = $kalin }
        if ($i -eq 4 -and $R.Durum -eq 'Completed') { $si.ForeColor = (Get-DurumRengi 'Completed'); $si.Font = $kalin }
        if ($i -eq 5 -and $R.Oncelik -eq 'High' -and -not $kapali) { $si.ForeColor = Get-Renk 192 0 0; $si.Font = $kalin }
        if (($i -eq 9 -or $i -eq 8) -and $gec) { $si.ForeColor = Get-Renk 192 0 0; $si.Font = $kalin }
        $i++
    }
    return $it
}

function Yenile-Agac {
    $secili = $null
    try { if ($lv.SelectedItems.Count -gt 0) { $secili = [string]$lv.SelectedItems[0].Tag } } catch { }
    $ara = ([string]$txtAgac.Text).Trim().ToLower()
    $sadeceAcik = [bool]$chkAgac.Checked
    $uyar = {
        param($r)
        if (-not $ara) { return $true }
        return ((($r.ID, $r.ParcaAdi, $r.Proje, $r.Layout, $r.Malzeme, $r.Amac, $r.Kosul, $r.AnalizYapan, $r.TalepEden, $r.ModelNo) -join ' ').ToLower().Contains($ara))
    }
    $lv.BeginUpdate()
    try {
        $lv.Items.Clear(); $lv.Groups.Clear()
        $dolu = [string][char]0x25CF; $bos = [string][char]0x25CB
        $kullanilan = @{}
        foreach ($m in @($script:Tum | Where-Object { $_.ModelRolu -eq 'Master Model' })) {
            $kullanilan[$m.ID] = $true
            $alt = @($script:Tum | Where-Object { $_.ModelRolu -eq 'Sub-Analysis' -and $_.AnaModelID -eq $m.ID } | Sort-Object ID)
            foreach ($c in $alt) { $kullanilan[$c.ID] = $true }
            $gorunenAlt = @($alt | Where-Object { (-not $sadeceAcik -or -not (Test-Kapali $_)) -and ((& $uyar $_) -or (& $uyar $m)) })
            $anaUyar = (& $uyar $m)
            if ($ara -and -not $anaUyar -and $gorunenAlt.Count -eq 0) { continue }
            if ($sadeceAcik -and (Test-Kapali $m) -and $gorunenAlt.Count -eq 0) { continue }
            $biten = @($alt | Where-Object { Test-Kapali $_ }).Count
            $noktalar = (($alt | ForEach-Object { if (Test-Kapali $_) { $dolu } else { $bos } }) -join '')
            $baslik = '{0}   {1}   |   {2}   |   {3}   |   {4}   {5}  {6}/{7} sub-analyses completed' -f $m.ID, $m.ParcaAdi, $m.Proje, $m.Malzeme, $m.Durum, $noktalar, $biten, $alt.Count
            $g = New-Object System.Windows.Forms.ListViewGroup($baslik)
            [void]$lv.Groups.Add($g)
            $it = Yeni-AgacSatiri $m 'Master Model' $true; $it.Group = $g; [void]$lv.Items.Add($it)
            foreach ($c in $gorunenAlt) { $it = Yeni-AgacSatiri $c 'Sub-Analysis' $false; $it.Group = $g; [void]$lv.Items.Add($it) }
        }
        $yetim = @($script:Tum | Where-Object { $_.ModelRolu -eq 'Sub-Analysis' -and -not $kullanilan.ContainsKey($_.ID) -and (-not $sadeceAcik -or -not (Test-Kapali $_)) -and (& $uyar $_) })
        if ($yetim.Count -gt 0) {
            $g = New-Object System.Windows.Forms.ListViewGroup('SUB-ANALYSES WITHOUT A MASTER MODEL (please check)'); [void]$lv.Groups.Add($g)
            foreach ($c in $yetim) { $it = Yeni-AgacSatiri $c 'Sub-Analysis' $false; $it.Group = $g; [void]$lv.Items.Add($it) }
        }
        $bag = @($script:Tum | Where-Object { $_.ModelRolu -ne 'Master Model' -and $_.ModelRolu -ne 'Sub-Analysis' -and (-not $sadeceAcik -or -not (Test-Kapali $_)) -and (& $uyar $_) })
        if ($bag.Count -gt 0) {
            $g = New-Object System.Windows.Forms.ListViewGroup('STANDALONE ANALYSES  (' + $bag.Count + ')'); [void]$lv.Groups.Add($g)
            foreach ($c in $bag) { $it = Yeni-AgacSatiri $c 'Standalone' $false; $it.Group = $g; [void]$lv.Items.Add($it) }
        }
        if ($secili) { foreach ($x in $lv.Items) { if ([string]$x.Tag -eq $secili) { $x.Selected = $true; break } } }
    } finally { $lv.EndUpdate() }
}

function Yenile-Ozet {
    $o = Get-Ozet $script:Tum
    foreach ($k in 'Toplam', 'Acik', 'Tamam', 'Geciken', 'Yakin7', 'AnaModel', 'AltAnaliz') { $script:Kartlar[$k].Text = [string]$o.$k }
    $gDurum.Rows.Clear(); foreach ($x in $o.Durumlar) { [void]$gDurum.Rows.Add($x.Ad, $x.Adet) }
    $gKisi.Rows.Clear(); foreach ($x in $o.Kisiler) { [void]$gKisi.Rows.Add($x.Ad, $x.Acik, $x.Geciken) }
    $gYakin.Rows.Clear()
    foreach ($x in $o.Yaklasan) {
        $dl = ConvertTo-Tarih $x.Deadline
        [void]$gYakin.Rows.Add((Get-KalanGun $x), $x.ID, $x.ParcaAdi, $x.AnalizYapan, $(if ($dl) { $dl.ToString('dd.MM.yyyy') } else { '' }))
    }
    foreach ($gr in $gYakin.Rows) {
        $kv = $gr.Cells[0].Value
        if ($null -ne $kv -and [int]$kv -lt 0) { $gr.DefaultCellStyle.ForeColor = $kirmizi } elseif ($null -ne $kv -and [int]$kv -le 7) { $gr.DefaultCellStyle.ForeColor = Get-Renk 191 120 0 }
    }
    foreach ($gg in @($gDurum, $gKisi, $gYakin)) { Sec-Temizle $gg }
}

function Yenile-Filtreler {
    $fd = [string]$cmbDurum.SelectedItem; $fk = [string]$cmbKisi.SelectedItem
    $script:FiltreKilit = $true
    try {
        $l = Get-Listeler
        $cmbDurum.Items.Clear(); [void]$cmbDurum.Items.Add('All')
        foreach ($x in $l['Durum']) { [void]$cmbDurum.Items.Add($x) }
        $cmbKisi.Items.Clear(); [void]$cmbKisi.Items.Add('All')
        foreach ($x in @($script:Tum | ForEach-Object { $_.AnalizYapan } | Where-Object { $_ } | Sort-Object -Unique)) { [void]$cmbKisi.Items.Add($x) }
        $cmbDurum.SelectedItem = $(if ($fd -and $cmbDurum.Items.Contains($fd)) { $fd } else { 'All' })
        $cmbKisi.SelectedItem = $(if ($fk -and $cmbKisi.Items.Contains($fk)) { $fk } else { 'All' })
    } finally { $script:FiltreKilit = $false }
}

function Yenile-Hepsi {
    $script:Tum = @(Get-TumKayitlar)
    $script:Imza = Get-VeriImzasi
    Yenile-Filtreler
    Yenile-Liste
    Yenile-Agac
    Yenile-Ozet
    $sbKayit.Text = 'Records: ' + $script:Tum.Count
    $sbVeri.Text = 'Data: ' + $script:Kok
    $script:Kisiler = @(Get-TumKullanicilar)
    if ($null -ne $script:Ben) { $sbKul.Text = 'User: ' + $script:Ben.Ad + ' (' + $script:Ben.Rol + ')' }
    $sbZaman.Text = 'Last refresh: ' + (Get-Date).ToString('HH:mm:ss')
}

function Yetki-Uygula {
    $k = $script:Ben
    $miListe.Visible = (Test-TamYetki $k)
    $tbAdmin.Visible = (Test-Admin $k)
    $tbBildirim.Visible = ($k.Rol -ne 'Leader')
    $sbKul.Text = 'User: ' + $k.Ad + ' (' + $k.Rol + ')'
    $ana.Text = 'Analysis Tracker  -  ' + $k.Ad + '  (' + $k.Rol + ')'
}

function Yenile-Ben {
    # yonetici rol/ad degistirmis olabilir; kayit silinmisse program kapanir. Gecici ag hatalari yok sayilir.
    try { $k = Get-Kullanici (Get-HesapAdi) } catch { return $true }
    if ($null -eq $k) {
        if (-not $script:Kapaniyor) {
            $script:Kapaniyor = $true
            [void](Mesaj 'Your user registration was removed by the administrator. The program will now close.' 'Warning')
            try { [void]$ana.Close() } catch { }
        }
        return $false
    }
    $degisti = ($null -eq $script:Ben -or $k.Rol -ne $script:Ben.Rol -or $k.Ad -ne $script:Ben.Ad -or (Test-Admin $k) -ne (Test-Admin $script:Ben))
    $script:Ben = $k
    if ($degisti) { Yetki-Uygula }
    return $true
}

function Yenile-Bildirim {
    if ($null -eq $script:Ben -or $script:Ben.Rol -eq 'Leader') { return }
    $liste = @(Get-BenimBildirimler $script:Ben)
    $oku = @($liste | Where-Object { $_.Okunmadi })
    $tbBildirim.Text = $(if ($oku.Count -gt 0) { 'Notifications (' + $oku.Count + ')' } else { 'Notifications' })
    $tbBildirim.ForeColor = $(if ($oku.Count -gt 0) { Get-Renk 192 0 0 } else { [System.Drawing.Color]::Black })
    $yeni = @($oku | Where-Object { -not $script:Gosterilen.ContainsKey($_.Dosya) })
    foreach ($b in $oku) { $script:Gosterilen[$b.Dosya] = $true }
    if ($yeni.Count -gt 0 -and $null -ne $script:Tray) {
        $metin = $(if ($yeni.Count -eq 1) { $yeni[0].Mesaj } else { "$($yeni.Count) new notifications (latest: $($yeni[0].Mesaj))" })
        try { $script:Tray.ShowBalloonTip(8000, 'Analysis Tracker', $metin, [System.Windows.Forms.ToolTipIcon]::Info) } catch { }
    }
}

function Otomatik-Yenile {
    if (-not (Yenile-Ben)) { return }
    Yenile-Bildirim
    if ($script:DlgAcik) { return }
    if ((Get-VeriImzasi) -ne $script:Imza) { Yenile-Hepsi }
}

# ---------------------------------------------------------------- kayıt formu (modal)
function Get-Secenekler {
    param($Alan, $Listeler, $Tum)
    if ($Alan.K -eq 'TalepEden') { return @(@(Get-TumKullanicilar) | ForEach-Object { $_.Ad }) }
    if ($Alan.K -eq 'AnalizYapan') { return (@('') + @(@(Get-TumKullanicilar) | Where-Object { $_.Rol -eq 'Structural Analysis Engineer' } | ForEach-Object { $_.Ad })) }
    $liste = @()
    if ($Listeler.ContainsKey($Alan.List)) { $liste = @($Listeler[$Alan.List]) }
    if (-not $Alan.Strict) {
        $ek = @($Tum | ForEach-Object { $_.($Alan.K) } | Where-Object { $_ })
        foreach ($e in $ek) { if ($liste -notcontains $e) { $liste += $e } }
    }
    return $liste
}

$script:AyAdlari = @('Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec')

function New-TarihKontrol {
    # Ingilizce gun / ay / yil secimi (Windows takvim penceresi sistem diliyle acildigi icin kullanilmaz)
    param($Deger, [bool]$Zorunlu, [bool]$Sabit)
    $p = New-Object System.Windows.Forms.FlowLayoutPanel; $p.FlowDirection = 'LeftToRight'; $p.WrapContents = $false; $p.Width = 330; $p.Height = 30
    $cb = New-Object System.Windows.Forms.CheckBox; $cb.AutoSize = $true; $cb.Text = ''; $cb.Width = 20; $cb.Margin = New-Object System.Windows.Forms.Padding(0, 6, 4, 0)
    $g = New-Object System.Windows.Forms.ComboBox; $g.DropDownStyle = 'DropDownList'; $g.Width = 52
    $a = New-Object System.Windows.Forms.ComboBox; $a.DropDownStyle = 'DropDownList'; $a.Width = 64
    $y = New-Object System.Windows.Forms.ComboBox; $y.DropDownStyle = 'DropDownList'; $y.Width = 70
    foreach ($x in @($g, $a, $y)) { [void](Stop-Tekerlek $x); $x.Margin = New-Object System.Windows.Forms.Padding(0, 3, 6, 0) }
    for ($i = 1; $i -le 31; $i++) { [void]$g.Items.Add([string]$i) }
    foreach ($m in $script:AyAdlari) { [void]$a.Items.Add($m) }
    $y0 = (Get-Date).Year
    $d0 = $(if ($null -ne $Deger) { ([datetime]$Deger).Year } else { $y0 })
    for ($i = [Math]::Min($y0, $d0) - 1; $i -le [Math]::Max($y0, $d0) + 6; $i++) { [void]$y.Items.Add([string]$i) }
    $t = @{ Chk = $cb; Gun = $g; Ay = $a; Yil = $y }
    $p.Tag = $t
    $guncelle = { param($s, $e) }
    if ($null -ne $Deger) {
        $d = [datetime]$Deger
        $g.SelectedItem = [string]$d.Day; $a.SelectedItem = $script:AyAdlari[$d.Month - 1]; $y.SelectedItem = [string]$d.Year; $cb.Checked = $true
    } else { $cb.Checked = $false }
    if (-not $Zorunlu) { [void]$cb.Add_CheckedChanged({ param($s, $e) try { $t2 = $s.Parent.Tag; foreach ($k in 'Gun', 'Ay', 'Yil') { $t2[$k].Enabled = $s.Checked } } catch { } }) }
    foreach ($k in 'Gun', 'Ay', 'Yil') { $t[$k].Enabled = ($cb.Checked -and -not $Sabit) }
    if ($Zorunlu) { $cb.Checked = $true } else { [void]$p.Controls.Add($cb) }
    if ($Sabit) { $cb.Enabled = $false }
    foreach ($x in @($g, $a, $y)) { [void]$p.Controls.Add($x) }
    return ,$p
}

function Get-TarihDeger {
    param($P)
    $t = $P.Tag
    if (-not $t.Chk.Checked) { return $null }
    if ($null -eq $t.Gun.SelectedItem -or $null -eq $t.Ay.SelectedItem -or $null -eq $t.Yil.SelectedItem) { return $null }
    $ay = [Array]::IndexOf($script:AyAdlari, [string]$t.Ay.SelectedItem) + 1
    $yil = [int]$t.Yil.SelectedItem
    $gun = [Math]::Min([int]$t.Gun.SelectedItem, [DateTime]::DaysInMonth($yil, $ay))
    return (New-Object DateTime($yil, $ay, $gun))
}

function Set-AlanDegeri {
    param($Alan, $C, [string]$Deger)
    switch ($Alan.T) {
        'multi' {
            $secili = @($Deger -split '\s*,\s*' | Where-Object { $_ })
            foreach ($cb in $script:D.Chk[$Alan.K]) { $cb.Checked = ($secili -contains $cb.Text) }
        }
        'multitext' {
            $lb = $script:D.Lst[$Alan.K]; $lb.Items.Clear()
            foreach ($s in @($Deger -split '\s*;\s*' | Where-Object { $_ })) { [void]$lb.Items.Add($s) }
            $C.Text = ''
        }
        'combo' { if ($Alan.Strict) { if ($Deger) { if (-not $C.Items.Contains($Deger)) { [void]$C.Items.Add($Deger) }; $C.SelectedItem = $Deger } else { $C.SelectedItem = $null } } else { $C.Text = $Deger } }
        'text' { $C.Text = $Deger }
    }
}

function Get-AlanDegeri {
    param($Alan, $C)
    switch ($Alan.T) {
        'multi' { return ((@($script:D.Chk[$Alan.K] | Where-Object { $_.Checked } | ForEach-Object { $_.Text })) -join ', ') }
        'date' { return (ConvertFrom-Tarih (Get-TarihDeger $C)) }
        'multitext' {
            $liste = @(); foreach ($i in $script:D.Lst[$Alan.K].Items) { $liste += [string]$i }
            $t = ([string]$C.Text).Trim()
            if ($t -and ($liste -notcontains $t)) { $liste += $t }
            return ($liste -join '; ')
        }
        'master' { $t = $C.Text; if ($t) { return (($t -split ' \| ')[0]) } else { return '' } }
        default { return ([string]$C.Text).Trim() }
    }
}

function Show-KayitFormu {
    # $Rec: acilacak kayit (yeni ise bos taslak). Kaydedildiyse $true doner.
    param($Rec, [bool]$Yeni, [string]$Mod = 'Full')
    $listeler = Get-Listeler
    $tum = @(Get-TumKayitlar)
    $eski = $null
    if (-not $Yeni) { $eski = ConvertTo-Kayit $Rec }

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = $(if ($Yeni) { 'New Analysis' } elseif ($Mod -eq 'Gor') { 'Analysis - ' + $Rec.ID + '  (read only)' } else { 'Analysis - ' + $Rec.ID })
    $dlg.StartPosition = 'CenterParent'
    # kucuk ekranlarda pencere ekrandan tasmasin
    $yuk = 880
    try { $yuk = [Math]::Min(880, [Math]::Max(500, [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Height - 40)) } catch { }
    $dlg.Size = Get-Boyut 880 $yuk
    $dlg.MinimumSize = Get-Boyut 700 520
    $dlg.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $dlg.MinimizeBox = $false; $dlg.MaximizeBox = $true; $dlg.ShowInTaskbar = $false

    # alt dugme alani: sabit konum / Anchor yok, pencere nasil yeniden boyutlanirsa boyutlansin sag altta kalir
    $alt = New-Object System.Windows.Forms.Panel; $alt.Dock = 'Bottom'; $alt.Height = 58; $alt.BackColor = Get-Renk 240 243 249
    $btnK = New-Object System.Windows.Forms.Button; $btnK.Text = 'Save'; $btnK.Width = 120; $btnK.Height = 32
    $btnK.BackColor = Get-Renk 31 56 100; $btnK.ForeColor = [System.Drawing.Color]::White; $btnK.FlatStyle = 'Flat'; $btnK.Font = New-Object System.Drawing.Font('Segoe UI', 9.5, [System.Drawing.FontStyle]::Bold)
    $btnI = New-Object System.Windows.Forms.Button; $btnI.Text = 'Cancel'; $btnI.Width = 100; $btnI.Height = 32; $btnI.FlatStyle = 'Flat'; $btnI.BackColor = [System.Drawing.Color]::White
    $akis = New-Object System.Windows.Forms.FlowLayoutPanel; $akis.Dock = 'Fill'; $akis.FlowDirection = 'RightToLeft'
    $akis.Padding = New-Object System.Windows.Forms.Padding(10, 10, 14, 0)
    $btnD = New-Object System.Windows.Forms.Button; $btnD.Text = 'Save + Duplicate'; $btnD.Width = 140; $btnD.Height = 32; $btnD.FlatStyle = 'Flat'; $btnD.BackColor = [System.Drawing.Color]::White
    $akis.Controls.Add($btnI); $akis.Controls.Add($btnK); $akis.Controls.Add($btnD)
    $alt.Controls.Add($akis)
    $ipucuL = New-Object System.Windows.Forms.Label; $ipucuL.Dock = 'Left'; $ipucuL.Width = 330; $ipucuL.TextAlign = 'MiddleLeft'; $ipucuL.ForeColor = Get-Renk 110 110 110
    $ipucuL.Text = $(if ($Mod -eq 'Gor') { '  Read only' } else { '  * Required fields' })
    $alt.Controls.Add($ipucuL)

    # ust baslik seridi
    $ban = New-Object System.Windows.Forms.Panel; $ban.Dock = 'Top'; $ban.Height = 66; $ban.BackColor = Get-Renk 31 56 100
    $bt1 = New-Object System.Windows.Forms.Label; $bt1.Location = New-Object System.Drawing.Point(20, 9); $bt1.Size = Get-Boyut 800 28
    $bt1.Font = New-Object System.Drawing.Font('Segoe UI', 14, [System.Drawing.FontStyle]::Bold); $bt1.ForeColor = [System.Drawing.Color]::White
    $bt2 = New-Object System.Windows.Forms.Label; $bt2.Location = New-Object System.Drawing.Point(22, 38); $bt2.Size = Get-Boyut 800 20; $bt2.ForeColor = Get-Renk 196 210 232
    if ($Yeni) {
        $bt1.Text = $(if ($script:KopyaKaynak) { 'Duplicate Analysis Request' } else { 'New Analysis Request' })
        if ($script:KopyaKaynak) { $bt2.Text = 'Copy of ' + $script:KopyaKaynak + '  |  Change the option that differs (Material Model, Alpha, Model Type ...) and save.' } else {
        $bt2.Text = 'Requested by ' + $script:Ben.Ad + '  |  Fill in the fields below. Status, analyst notes and file paths are completed later by the analyst.' }
    } else {
        $bt1.Text = $Rec.ID + '   ' + $Rec.ParcaAdi
        $bt2.Text = 'Status: ' + $Rec.Durum + '  |  Requester: ' + $Rec.TalepEden + $(if ($Rec.SonGuncelleyen) { '  |  Last updated by ' + $Rec.SonGuncelleyen } else { '' })
    }
    $ban.Controls.Add($bt1); $ban.Controls.Add($bt2)

    $pnl = New-Object System.Windows.Forms.Panel; $pnl.Dock = 'Fill'; $pnl.AutoScroll = $true; $pnl.BackColor = [System.Drawing.Color]::White
    $tbl = New-Object System.Windows.Forms.TableLayoutPanel
    $tbl.Dock = 'Top'; $tbl.ColumnCount = 2; $tbl.RowCount = $script:Alanlar.Count + 10; $tbl.Padding = New-Object System.Windows.Forms.Padding(12, 6, 22, 12)
    [void]$tbl.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Absolute', 190)))
    [void]$tbl.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 100)))

    $ctl = @{}; $chk = @{}; $lst = @{}; $combos = @(); $satirYer = @{}
    $satir = 0; $toplamH = 16
    foreach ($a in $script:Alanlar) {
        if ($a.T -eq 'ro' -or $a.Bagli) { continue }
        if ($a.SadeceDuzenle -and $Yeni) { continue }
        if ($a.Bolum) {
            $hl = New-Object System.Windows.Forms.Label; $hl.Text = $a.Bolum.ToUpper(); $hl.Dock = 'Fill'; $hl.TextAlign = 'MiddleLeft'
            $hl.Font = New-Object System.Drawing.Font('Segoe UI', 8.5, [System.Drawing.FontStyle]::Bold); $hl.ForeColor = Get-Renk 31 56 100; $hl.BackColor = Get-Renk 230 236 246
            $hl.Padding = New-Object System.Windows.Forms.Padding(8, 0, 0, 0); $hl.Margin = New-Object System.Windows.Forms.Padding(0, 10, 0, 4)
            [void]$tbl.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 40)))
            $tbl.Controls.Add($hl, 0, $satir); $tbl.SetColumnSpan($hl, 2)
            $satir++; $toplamH += 40
        }
        $h = 34
        $yer = $null
        $deger = [string]$Rec.($a.K)
        switch ($a.T) {
            'text' { $c = New-Object System.Windows.Forms.TextBox; $c.Text = $deger }
            'memo' {
                $c = New-Object System.Windows.Forms.TextBox; $c.Multiline = $true; $c.ScrollBars = 'Vertical'; $c.AcceptsReturn = $true; $c.Text = $deger.Replace("`r`n", "`n").Replace("`n", "`r`n"); $h = 66
            }
            'combo' {
                $c = New-Object System.Windows.Forms.ComboBox; Stop-Tekerlek $c
                if (-not $a.Strict) { $combos += $c; $c.Add_SizeChanged({ param($s, $e) try { if (-not $s.Focused) { $s.SelectionLength = 0 } } catch { } }) }
                $sec = @(Get-Secenekler $a $listeler $tum)
                if ($deger -and ($sec -notcontains $deger)) { $sec = @($deger) + $sec }
                if ($a.Strict) { $c.DropDownStyle = 'DropDownList' } else { $c.DropDownStyle = 'DropDown' }
                [void]$c.Items.AddRange([object[]]$sec)
                if ($a.Strict) { if ($deger) { $c.SelectedItem = $deger } } else { $c.Text = $deger }
                # yaninda metin kutusu acilan alan (Aero / SAS / Termal)
                $es = @($script:Alanlar | Where-Object { $_.Bagli -eq $a.K })
                if ($es.Count -gt 0) {
                    $t = New-Object System.Windows.Forms.TextBox; $t.Text = [string]$Rec.($es[0].K); $t.Dock = 'Fill'
                    $t.Margin = New-Object System.Windows.Forms.Padding(8, 3, 0, 3)
                    $c.Dock = 'Fill'
                    $yer = New-Object System.Windows.Forms.TableLayoutPanel; $yer.ColumnCount = 2; $yer.RowCount = 1
                    [void]$yer.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Absolute', 170)))
                    [void]$yer.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 100)))
                    [void]$yer.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 100)))
                    $yer.Controls.Add($c, 0, 0); $yer.Controls.Add($t, 1, 0)
                    $ctl[$es[0].K] = $t
                }
            }
            'multitext' {
                # birden fazla deger girilebilir (ornegin birden fazla parca adi / malzeme): yaz + Add, listeden Remove
                if ($a.List) {
                    $c = New-Object System.Windows.Forms.ComboBox; $c.DropDownStyle = 'DropDown'; Stop-Tekerlek $c
                    $combos += $c; $c.Add_SizeChanged({ param($s, $e) try { if (-not $s.Focused) { $s.SelectionLength = 0 } } catch { } })
                    [void]$c.Items.AddRange([object[]]@(Get-Secenekler $a $listeler $tum))
                } else { $c = New-Object System.Windows.Forms.TextBox }
                $c.Dock = 'Fill'; $c.Tag = $a.K
                $c.Add_KeyDown({ param($s, $e) if ($e.KeyCode -eq 'Return') { $e.Handled = $true; $e.SuppressKeyPress = $true; & $script:D.ListeEkle $s.Tag } })
                $lb = New-Object System.Windows.Forms.ListBox; $lb.Dock = 'Fill'; $lb.IntegralHeight = $false; $lb.Tag = $a.K
                foreach ($s in @($deger -split '\s*;\s*' | Where-Object { $_ })) { [void]$lb.Items.Add($s) }
                $lb.Add_DoubleClick({ param($s, $e) & $script:D.ListeCikar $s.Tag })
                $bE = New-Object System.Windows.Forms.Button; $bE.Text = 'Add'; $bE.Dock = 'Fill'; $bE.Tag = $a.K
                $bE.Add_Click({ param($s, $e) & $script:D.ListeEkle $s.Tag })
                $bC = New-Object System.Windows.Forms.Button; $bC.Text = 'Remove'; $bC.Dock = 'Top'; $bC.Height = 26; $bC.Tag = $a.K
                $bC.Add_Click({ param($s, $e) & $script:D.ListeCikar $s.Tag })
                $yer = New-Object System.Windows.Forms.TableLayoutPanel; $yer.ColumnCount = 2; $yer.RowCount = 2
                [void]$yer.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 100)))
                [void]$yer.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Absolute', 84)))
                [void]$yer.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 28)))
                [void]$yer.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 100)))
                $yer.Controls.Add($c, 0, 0); $yer.Controls.Add($bE, 1, 0); $yer.Controls.Add($lb, 0, 1); $yer.Controls.Add($bC, 1, 1)
                $lst[$a.K] = $lb; $h = 104
            }
            'multi' {
                # birden fazla secilebilir (ornegin Elastik + Elastoplastik)
                $c = New-Object System.Windows.Forms.FlowLayoutPanel; $c.FlowDirection = 'LeftToRight'; $c.WrapContents = $false
                $secili = @($deger -split '\s*,\s*' | Where-Object { $_ })
                $cbs = @()
                foreach ($s in @($listeler[$a.List])) {
                    $cb = New-Object System.Windows.Forms.CheckBox; $cb.Text = $s; $cb.AutoSize = $true
                    $cb.Margin = New-Object System.Windows.Forms.Padding(0, 6, 24, 0); $cb.Checked = ($secili -contains $s)
                    $c.Controls.Add($cb); $cbs += $cb
                }
                $chk[$a.K] = $cbs
            }
            'master' {
                $c = New-Object System.Windows.Forms.ComboBox; $c.DropDownStyle = 'DropDownList'; Stop-Tekerlek $c
                [void]$c.Items.Add('')
                $secili = ''
                foreach ($m in @($tum | Where-Object { $_.ModelRolu -eq 'Master Model' -and $_.ID -ne $Rec.ID })) {
                    $metin = '{0} | {1} | {2} | {3}' -f $m.ID, $m.Proje, $m.ParcaAdi, $m.Layout
                    [void]$c.Items.Add($metin)
                    if ($m.ID -eq $deger) { $secili = $metin }
                }
                $c.SelectedItem = $secili
            }
            'date' {
                $d0 = ConvertTo-Tarih $deger
                $bugun = (Get-Date).Date
                if ($a.K -eq 'TalepTarihi' -and $Yeni) { $c = New-TarihKontrol $bugun $true $true }
                else { $c = New-TarihKontrol $d0 $false $false }
            }
            'hist' {
                $c = New-Object System.Windows.Forms.TextBox; $c.Multiline = $true; $c.ReadOnly = $true; $c.ScrollBars = 'Vertical'; $c.BackColor = Get-Renk 245 245 245; $h = 110
                $c.Text = (@($Rec.Gecmis) -join "`r`n")
            }
        }
        $l = New-Object System.Windows.Forms.Label; $l.Text = $a.L; $l.Dock = 'Fill'; $l.TextAlign = 'MiddleLeft'
        if ($a.T -in @('memo', 'hist', 'multitext')) { $l.TextAlign = 'TopLeft'; $l.Padding = New-Object System.Windows.Forms.Padding(0, 6, 0, 0) }
        $l.Font = New-Object System.Drawing.Font('Segoe UI', 9); $l.ForeColor = $(if ($a.L -like '*`**') { Get-Renk 170 30 30 } else { Get-Renk 55 55 55 })
        if ($null -eq $yer) { $yer = $c }
        if ($a.T -ne 'date') { $yer.Dock = 'Fill' } else { $c.Anchor = 'Left'; $c.Width = 330 }
        [void]$tbl.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', $h)))
        $tbl.Controls.Add($l, 0, $satir)
        $tbl.Controls.Add($yer, 1, $satir)
        $ctl[$a.K] = $c; $satirYer[$a.K] = $yer
        $satir++; $toplamH += $h
    }
    # yetkiye gore alanlari kilitle (salt okunur: hepsi; talep modu: analizci alanlari + talep eden)
    foreach ($a in $script:Alanlar) {
        if (-not $satirYer.ContainsKey($a.K) -or $a.T -eq 'hist') { continue }
        if ($Mod -eq 'Gor' -or ($Mod -eq 'Talep' -and ($a.Analizci -or $a.K -eq 'TalepEden'))) { $satirYer[$a.K].Enabled = $false }
    }
    if ($Mod -eq 'Gor') { $btnK.Visible = $false; $btnD.Visible = $false; $btnI.Text = 'Close' }
    $tbl.Height = $toplamH
    $pnl.Controls.Add($tbl)
    $dlg.Controls.Add($pnl)
    $dlg.Controls.Add($ban)
    $dlg.Controls.Add($alt)

    $anaAlan = @($script:Alanlar | Where-Object { $_.K -eq 'AnaModelID' })[0]
    $script:D = @{ Dlg = $dlg; Ctl = $ctl; Chk = $chk; Lst = $lst; Combos = $combos; Rec = $Rec; Eski = $eski; Yeni = $Yeni; Mod = $Mod; Kaydedildi = $false; KopyaIste = $false; SonID = ''; Tum = $tum; AnaAlan = $anaAlan }

    # coklu giris listeleri (parca adi, malzeme)
    $script:D.ListeEkle = {
        param($k)
        $kutu = $script:D.Ctl[$k]; $lb = $script:D.Lst[$k]
        $t = ([string]$kutu.Text).Trim()
        if ($t) {
            $var = $false
            foreach ($i in $lb.Items) { if ([string]$i -eq $t) { $var = $true } }
            if (-not $var) { [void]$lb.Items.Add($t) }
        }
        $kutu.Text = ''
    }
    $script:D.ListeCikar = {
        param($k)
        $lb = $script:D.Lst[$k]
        if ($lb.SelectedIndex -ge 0) { $lb.Items.RemoveAt($lb.SelectedIndex) }
    }
    # acilir kutularda tum metnin mavi secili kalmasini onler (pencere boyutu degisince olusan hata)
    $script:D.SecimiTemizle = { foreach ($cb in $script:D.Combos) { try { if (-not $cb.Focused) { $cb.SelectionLength = 0 } } catch { } } }
    $dlg.Add_Shown({ & $script:D.SecimiTemizle })
    $dlg.Add_SizeChanged({ & $script:D.SecimiTemizle })
    $dlg.Add_ResizeEnd({ & $script:D.SecimiTemizle })

    # Aero/SAS/Termal 'Ready' veya 'Waiting' ise yanda metin kutusu gorunur
    $script:D.AdDegisti = {
        $c = $script:D.Ctl
        foreach ($p in @(@('Aero', 'AeroAdi'), @('SAS', 'SASAdi'), @('Termal', 'TermalAdi'))) {
            $gerek = ($c[$p[0]].Text -eq 'Ready' -or $c[$p[0]].Text -eq 'Waiting')
            $c[$p[1]].Visible = $gerek
            if (-not $gerek) { $c[$p[1]].Text = '' }
        }
    }
    foreach ($k in 'Aero', 'SAS', 'Termal') { $ctl[$k].Add_SelectedIndexChanged({ & $script:D.AdDegisti }) }

    # 'Sub-Analysis' secilirse bagli ana model aktif olur; aksi halde kapali ve bos
    $script:D.RolDegisti = {
        $c = $script:D.Ctl
        $alt = ($c['ModelRolu'].Text -eq 'Sub-Analysis')
        $c['AnaModelID'].Enabled = $alt
        if (-not $alt) { $c['AnaModelID'].SelectedItem = '' }
    }
    # ana model secilince model bilgileri ana modelden kopyalanir (yalnizca yeni kayitta)
    $script:D.AnaDegisti = {
        $d = $script:D; $c = $d.Ctl
        $id = Get-AlanDegeri $d.AnaAlan $c['AnaModelID']
        if ($id -and $d.Yeni) {
            $m = @($d.Tum | Where-Object { $_.ID -eq $id })[0]
            if ($null -ne $m) {
                foreach ($a in $script:Alanlar) { if ($a.Miras -and $c.ContainsKey($a.K)) { Set-AlanDegeri $a $c[$a.K] ([string]$m.($a.K)) } }
            }
            & $d.AdDegisti
        }
    }
    $ctl['ModelRolu'].Add_SelectedIndexChanged({ & $script:D.RolDegisti })
    $ctl['AnaModelID'].Add_SelectedIndexChanged({ & $script:D.AnaDegisti })
    & $script:D.RolDegisti
    & $script:D.AdDegisti

    $btnI.Add_Click({ $script:D.Dlg.DialogResult = 'Cancel'; $script:D.Dlg.Close() })
    $script:KaydetBlok = {
        Guvenli {
            $d = $script:D
            $yeniRec = ConvertTo-Kayit $d.Rec
            foreach ($a in $script:Alanlar) {
                if ($a.T -in @('ro', 'hist')) { continue }
                $c = $d.Ctl[$a.K]
                if ($null -eq $c) { continue }
                $yeniRec.($a.K) = Get-AlanDegeri $a $c
            }
            if ($d.Mod -eq 'Gor') { return }
            foreach ($a in $script:Alanlar) {
                if ($a.K -eq 'Deadline' -and $yeniRec.Deadline) {
                    $dd = ConvertTo-Tarih $yeniRec.Deadline; $eskiD = $(if ($d.Eski) { ConvertTo-Tarih $d.Eski.Deadline } else { $null })
                    if ($dd -and $dd.Date -lt (Get-Date).Date -and -not ($eskiD -and $eskiD.Date -eq $dd.Date)) { [void](Mesaj 'Due date cannot be earlier than today.' 'Warning'); return }
                }
            }
            if ($d.Mod -eq 'Talep' -and $null -ne $d.Eski) {
                # yetkisiz alanlar kayitta eski degerinde kalir
                foreach ($a in $script:Alanlar) { if ($a.Analizci -and $a.T -ne 'hist') { $yeniRec.($a.K) = $d.Eski.($a.K) } }
                $yeniRec.TalepEden = $d.Eski.TalepEden
            }
            if ($yeniRec.ModelRolu -ne 'Sub-Analysis') { $yeniRec.AnaModelID = '' }
            if ($d.Yeni -and -not $yeniRec.Durum) { $yeniRec.Durum = (Get-Listeler)['Durum'][0] }
            if ($yeniRec.Durum -eq 'Completed' -and -not $yeniRec.Tamamlanma) { $yeniRec.Tamamlanma = ConvertFrom-Tarih (Get-Date) }
            $taze = @(Get-TumKayitlar)
            $hatalar = @(Test-Kayit $yeniRec $taze $d.Eski)
            if ($hatalar.Count -gt 0) { [void](Mesaj ($hatalar -join "`r`n") 'Warning'); return }
            try {
                $sid = Save-Kayit $yeniRec $d.Eski
                $d.SonID = [string]$sid
                try { Send-Bildirim $yeniRec $d.Eski $(if ($d.Yeni) { 'Yeni' } else { 'Degisti' }) } catch { Write-HataLog ('Notification: ' + $_.Exception.ToString()) }
            } catch {
                if ($_.Exception.Message -eq 'CAKISMA') {
                    [void](Mesaj "This record was modified by another user while you were editing it.`r`n`r`nYour changes were NOT saved so that the other user's changes are not overwritten. Close this window, reopen the record and make your change again." 'Warning')
                    $d.Dlg.DialogResult = 'Cancel'; $d.Dlg.Close()
                    return
                }
                throw
            }
            $d.Kaydedildi = $true
            $d.Dlg.DialogResult = 'OK'; $d.Dlg.Close()
        }
    }
    $btnK.Add_Click({ $script:D.KopyaIste = $false; & $script:KaydetBlok })
    $btnD.Add_Click({ $script:D.KopyaIste = $true; & $script:KaydetBlok })

    $script:DlgAcik = $true
    try { [void]$dlg.ShowDialog($ana) } finally { $script:DlgAcik = $false }
    $sonuc = $script:D.Kaydedildi
    $script:KopyaID = $(if ($sonuc -and $script:D.KopyaIste) { $script:D.SonID } else { $null })
    $dlg.Dispose()
    return $sonuc
}

# ---------------------------------------------------------------- eylemler
function Yeni-Analiz {
    if (-not (Yenile-Ben)) { return }
    $r = New-BosKayit
    $l = Get-Listeler
    $r.TalepEden = $script:Ben.Ad
    $r.Durum = $l['Durum'][0]; $r.ModelRolu = $l['ModelRolu'][0]; $r.TalepTarihi = ConvertFrom-Tarih (Get-Date)
    $mod = $(if (Test-TamYetki $script:Ben) { 'Full' } else { 'Talep' })
    if (Show-KayitFormu $r $true $mod) { Yenile-Hepsi }
    if ($script:KopyaID) { Kopya-Ac $script:KopyaID }
}

function Kopya-Ac {
    # kaydi kopyalar; kullanici farkli secenegi (Alpha, malzeme modeli ...) secip kaydeder. 'Save + Duplicate' ile zincirlenir.
    param([string]$id)
    $sayac = 0
    while ($id -and $sayac -lt 20) {
        $sayac++
        $script:KopyaID = $null
        $src = Read-Kayit (Join-Path $script:KayitDir ($id + '.json'))
        if ($null -eq $src) { [void](Mesaj 'The source record could not be read.' 'Warning'); return }
        if ((Get-Yetki $src $script:Ben).Mod -eq 'Gor') { [void](Mesaj 'You can only duplicate your own requests.' 'Warning'); return }
        $r = New-BosKayit; $l = Get-Listeler
        foreach ($a in $script:Alanlar) { if ($a.T -in @('ro', 'hist') -or $a.Analizci) { continue }; $r.($a.K) = $src.($a.K) }
        $r.TalepEden = $script:Ben.Ad; $r.Durum = $l['Durum'][0]; $r.TalepTarihi = ConvertFrom-Tarih (Get-Date)
        $dd = ConvertTo-Tarih $src.Deadline
        if ($dd -and $dd.Date -lt (Get-Date).Date) { $r.Deadline = '' }
        $mod = $(if (Test-TamYetki $script:Ben) { 'Full' } else { 'Talep' })
        $script:KopyaKaynak = $id
        try { $ok = Show-KayitFormu $r $true $mod } finally { $script:KopyaKaynak = $null }
        if ($ok) { Yenile-Hepsi }
        $id = $script:KopyaID
    }
    $script:KopyaID = $null
}

function Kopyala-Secili {
    if (-not (Yenile-Ben)) { return }
    $id = Get-SeciliID
    if (-not $id) { [void](Mesaj 'Please select the record to duplicate first.'); return }
    Kopya-Ac $id
}

function Duzenle-Secili {
    $id = Get-SeciliID
    if (-not $id) { [void](Mesaj 'Please select a record from the list first.'); return }
    Ac-Kayit $id
}

function Ac-Kayit {
    param([string]$id)
    if (-not (Yenile-Ben)) { return }
    $yol = Join-Path $script:KayitDir ($id + '.json')
    $r = Read-Kayit $yol
    if ($null -eq $r) { [void](Mesaj 'The record file could not be read (it may have been deleted). Refreshing the list.' 'Warning'); Yenile-Hepsi; return }
    $y = Get-Yetki $r $script:Ben
    if (Show-KayitFormu $r $false $y.Mod) { Yenile-Hepsi } else { Otomatik-Yenile }
    if ($script:KopyaID) { Kopya-Ac $script:KopyaID }
}

function Sil-Secili {
    $id = Get-SeciliID
    if (-not $id) { [void](Mesaj 'Please select the record to delete first.'); return }
    if (-not (Yenile-Ben)) { return }
    $rec = Read-Kayit (Join-Path $script:KayitDir ($id + '.json'))
    if ($null -eq $rec) { [void](Mesaj 'The record file could not be read (it may have been deleted). Refreshing the list.' 'Warning'); Yenile-Hepsi; return }
    $yet = Get-Yetki $rec $script:Ben
    if (-not $yet.Sil) {
        [void](Mesaj "You do not have permission to delete this record.`r`n`r`nDesign engineers can delete only their own requests while the status is still the first one (Pending). Leaders and structural analysts can delete any record." 'Warning'); return
    }
    $tum = @(Get-TumKayitlar)
    $c = Get-Cocuklar $tum $id
    if ($c.Count -gt 0) { [void](Mesaj "This record is the master model of $($c.Count) sub-analysis(es); delete them or link them to another master model first." 'Warning'); return }
    if ((Mesaj "Delete record ${id}?`r`n(The file is moved to the 'Silinenler' folder and can be restored if needed.)" 'Question' 'YesNo') -ne 'Yes') { return }
    Remove-Kayit $id
    try { Send-Bildirim $rec $null 'Silindi' } catch { Write-HataLog ('Notification: ' + $_.Exception.ToString()) }
    Yenile-Hepsi
}

function Aktar-Excel {
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = 'CSV (opens in Excel) (*.csv)|*.csv'
    $sfd.FileName = 'AnalysisList_' + (Get-Date).ToString('yyyyMMdd_HHmm') + '.csv'
    $sfd.InitialDirectory = [Environment]::GetFolderPath('MyDocuments')
    if ($sfd.ShowDialog($ana) -ne 'OK') { return }
    Export-KayitlarCsv @(Get-TumKayitlar) $sfd.FileName
    if ((Mesaj "Saved:`r`n$($sfd.FileName)`r`n`r`nOpen in Excel?" 'Question' 'YesNo') -eq 'Yes') { Start-Process $sfd.FileName }
}

function Duzenle-Listeler {
    $anahtarlar = @(
        @('Status', 'Durum'), @('Priority', 'Oncelik'), @('Material', 'Malzeme'), @('Analysis Type', 'AnalizTuru'),
        @('Condition', 'Kosul')
    )
    $l = Get-Listeler
    $f = New-Object System.Windows.Forms.Form
    $f.Text = 'Edit lists'; $f.StartPosition = 'CenterParent'; $f.Size = Get-Boyut 560 520; $f.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $f.ShowInTaskbar = $false; $f.MinimizeBox = $false
    $tc = New-Object System.Windows.Forms.TabControl; $tc.Dock = 'Fill'
    $kutular = @{}
    foreach ($a in $anahtarlar) {
        $p = New-Object System.Windows.Forms.TabPage($a[0])
        $t = New-Object System.Windows.Forms.TextBox; $t.Multiline = $true; $t.ScrollBars = 'Vertical'; $t.Dock = 'Fill'; $t.AcceptsReturn = $true
        $t.Text = (@($l[$a[1]]) -join "`r`n")
        $p.Controls.Add($t); [void]$tc.TabPages.Add($p); $kutular[$a[1]] = $t
    }
    $ipucu = New-Object System.Windows.Forms.Label; $ipucu.Dock = 'Top'; $ipucu.Height = 38
    $ipucu.Text = "  Each line is one option. The 'Completed' and 'Cancelled' statuses are fixed.`r`n  People (requesters, analysts) come from the registered users."
    $alt = New-Object System.Windows.Forms.Panel; $alt.Dock = 'Bottom'; $alt.Height = 46
    $ok = New-Object System.Windows.Forms.Button; $ok.Text = 'Save'; $ok.Width = 100; $ok.Height = 28
    $iptal = New-Object System.Windows.Forms.Button; $iptal.Text = 'Cancel'; $iptal.Width = 100; $iptal.Height = 28
    $akis = New-Object System.Windows.Forms.FlowLayoutPanel; $akis.Dock = 'Fill'; $akis.FlowDirection = 'RightToLeft'; $akis.Padding = New-Object System.Windows.Forms.Padding(8, 8, 8, 0)
    $akis.Controls.Add($iptal); $akis.Controls.Add($ok)
    $alt.Controls.Add($akis)
    $f.Controls.Add($tc); $f.Controls.Add($ipucu); $f.Controls.Add($alt)
    $script:LD = @{ F = $f; Kutu = $kutular; Anahtarlar = $anahtarlar; Liste = $l }
    $iptal.Add_Click({ $script:LD.F.DialogResult = 'Cancel'; $script:LD.F.Close() })
    $ok.Add_Click({
        Guvenli {
            $yeni = Get-Listeler
            foreach ($a in $script:LD.Anahtarlar) {
                $yeni[$a[1]] = @($script:LD.Kutu[$a[1]].Text -split "`r?`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ })
            }
            Save-Listeler $yeni
            $script:LD.F.DialogResult = 'OK'; $script:LD.F.Close()
        }
    })
    $script:DlgAcik = $true
    try { $sonuc = $f.ShowDialog($ana) } finally { $script:DlgAcik = $false }
    $f.Dispose()
    if ($sonuc -eq 'OK') { Yenile-Hepsi }
}


# ---------------------------------------------------------------- kimlik / kayit
function Show-Girdi {
    # kucuk giris penceresi: metin veya (Secenekler verilirse) secim kutusu. Iptalde $null doner.
    param([string]$Baslik, [string]$Etiket, [string]$Varsayilan = '', $Secenekler = $null)
    $f = New-Object System.Windows.Forms.Form
    $f.Text = $Baslik; $f.StartPosition = 'CenterParent'; $f.Size = Get-Boyut 440 175; $f.FormBorderStyle = 'FixedDialog'
    $f.MinimizeBox = $false; $f.MaximizeBox = $false; $f.ShowInTaskbar = $false; $f.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $f.Padding = New-Object System.Windows.Forms.Padding(12)
    $l = New-Object System.Windows.Forms.Label; $l.Text = $Etiket; $l.Dock = 'Top'; $l.Height = 44
    if ($null -ne $Secenekler) {
        $c = New-Object System.Windows.Forms.ComboBox; $c.DropDownStyle = 'DropDownList'; Stop-Tekerlek $c
        [void]$c.Items.AddRange([object[]]@($Secenekler))
        if ($Varsayilan) { $c.SelectedItem = $Varsayilan }
    } else { $c = New-Object System.Windows.Forms.TextBox; $c.Text = $Varsayilan }
    $c.Dock = 'Top'
    $alt = New-Object System.Windows.Forms.FlowLayoutPanel; $alt.Dock = 'Bottom'; $alt.Height = 44; $alt.FlowDirection = 'RightToLeft'
    $ok = New-Object System.Windows.Forms.Button; $ok.Text = 'OK'; $ok.Width = 90; $ok.Height = 28
    $ip = New-Object System.Windows.Forms.Button; $ip.Text = 'Cancel'; $ip.Width = 90; $ip.Height = 28
    $alt.Controls.Add($ip); $alt.Controls.Add($ok)
    $f.Controls.Add($c); $f.Controls.Add($l); $f.Controls.Add($alt)
    $f.AcceptButton = $ok; $f.CancelButton = $ip
    $script:GD = @{ F = $f; C = $c; Sonuc = $null }
    $ok.Add_Click({ $script:GD.Sonuc = ([string]$script:GD.C.Text).Trim(); $script:GD.F.DialogResult = 'OK'; $script:GD.F.Close() })
    $ip.Add_Click({ $script:GD.F.DialogResult = 'Cancel'; $script:GD.F.Close() })
    [void]$f.ShowDialog($ana)
    $f.Dispose()
    return $script:GD.Sonuc
}

function Show-KullaniciKayit {
    # ilk acilista: ad + rol. Iptal edilirse $null doner (program kapanir).
    param([string]$Hesap)
    $ilk = (@(Get-TumKullanicilar).Count -eq 0)
    $f = New-Object System.Windows.Forms.Form
    $f.Text = 'Analysis Tracker - First login'; $f.StartPosition = 'CenterScreen'; $f.Size = Get-Boyut 520 450; $f.FormBorderStyle = 'FixedDialog'
    $f.MinimizeBox = $false; $f.MaximizeBox = $false; $f.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $l1 = New-Object System.Windows.Forms.Label; $l1.Location = New-Object System.Drawing.Point(16, 14); $l1.Size = Get-Boyut 470 112
    $l1.Text = "Welcome. This Windows account is not registered yet:`r`n$Hesap`r`n`r`nPlease enter your name and role. You will be recognised automatically next time."
    $l2 = New-Object System.Windows.Forms.Label; $l2.Text = 'Your name (full name):'; $l2.Location = New-Object System.Drawing.Point(16, 134); $l2.Size = Get-Boyut 470 22
    $tAd = New-Object System.Windows.Forms.TextBox; $tAd.Location = New-Object System.Drawing.Point(16, 158); $tAd.Size = Get-Boyut 470 26
    $l3 = New-Object System.Windows.Forms.Label; $l3.Text = 'Your role:'; $l3.Location = New-Object System.Drawing.Point(16, 198); $l3.Size = Get-Boyut 470 22
    $cRol = New-Object System.Windows.Forms.ComboBox; $cRol.DropDownStyle = 'DropDownList'; $cRol.Location = New-Object System.Drawing.Point(16, 222); $cRol.Size = Get-Boyut 470 26
    Stop-Tekerlek $cRol; [void]$cRol.Items.AddRange([object[]]$script:Roller)
    $not = New-Object System.Windows.Forms.Label; $not.Location = New-Object System.Drawing.Point(16, 266); $not.Size = Get-Boyut 470 44
    $not.Text = $(if ($ilk) { 'You are the first user, so you become the administrator.' } else { 'Only the administrator can change your name or role later.' })
    $ok = New-Object System.Windows.Forms.Button; $ok.Text = 'Continue'; $ok.Size = Get-Boyut 110 32; $ok.Location = New-Object System.Drawing.Point(262, 336)
    $ip = New-Object System.Windows.Forms.Button; $ip.Text = 'Exit'; $ip.Size = Get-Boyut 100 32; $ip.Location = New-Object System.Drawing.Point(382, 336)
    foreach ($c in @($l1, $l2, $tAd, $l3, $cRol, $not, $ok, $ip)) { $f.Controls.Add($c) }
    $f.AcceptButton = $ok; $f.CancelButton = $ip
    $script:KD = @{ F = $f; Ad = $tAd; Rol = $cRol; Hesap = $Hesap; Sonuc = $null }
    $ip.Add_Click({ $script:KD.F.DialogResult = 'Cancel'; $script:KD.F.Close() })
    $ok.Add_Click({
        try {
            $script:KD.Sonuc = New-Kullanici $script:KD.Hesap $script:KD.Ad.Text ([string]$script:KD.Rol.Text)
            $script:KD.F.DialogResult = 'OK'; $script:KD.F.Close()
        } catch { [void](Mesaj $_.Exception.Message 'Warning') }
    })
    [void]$f.ShowDialog()
    $f.Dispose()
    return $script:KD.Sonuc
}

function Initialize-Kimlik {
    $hesap = Get-HesapAdi
    $k = Get-Kullanici $hesap
    if ($null -eq $k) { $k = Show-KullaniciKayit $hesap }
    if ($null -eq $k) { return $null }
    try { $k.SonGiris = (Get-Date).ToString('s'); $k.Bilgisayar = [Environment]::MachineName; Save-Kullanici $k } catch { }
    return $k
}

# ---------------------------------------------------------------- bildirim penceresi
function Goster-Bildirimler {
    if (-not (Yenile-Ben)) { return }
    $liste = @(Get-BenimBildirimler $script:Ben)
    $f = New-Object System.Windows.Forms.Form
    $f.Text = 'Notifications'; $f.StartPosition = 'CenterParent'; $f.Size = Get-Boyut 900 460; $f.Font = New-Object System.Drawing.Font('Segoe UI', 9); $f.ShowInTaskbar = $false; $f.MinimizeBox = $false
    $g = New-Object System.Windows.Forms.DataGridView
    Set-CiftTampon $g
    $g.Dock = 'Fill'; $g.ReadOnly = $true; $g.AllowUserToAddRows = $false; $g.RowHeadersVisible = $false; $g.BackgroundColor = [System.Drawing.Color]::White
    $g.SelectionMode = 'FullRowSelect'; $g.MultiSelect = $false; $g.EnableHeadersVisualStyles = $false
    $g.ColumnHeadersDefaultCellStyle.BackColor = Get-Renk 31 56 100; $g.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White
    foreach ($c in @(@('New', 'Yeni', 45), @('Time', 'Zaman', 130), @('Record', 'Kayit', 75), @('Message', 'Mesaj', 520), @('From', 'Kimden', 100))) { Add-Kolon $g $c[0] $c[1] $c[2] }
    foreach ($b in $liste) { [void]$g.Rows.Add($(if ($b.Okunmadi) { [string][char]0x25CF } else { '' }), ($b.Zaman -replace 'T', ' ').Substring(0, 16), $b.KayitID, $b.Mesaj, $b.Kimden) }
    $kalin = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
    $g.Add_CellFormatting({ param($s, $e) if ($e.RowIndex -ge 0 -and [string]$s.Rows[$e.RowIndex].Cells[0].Value) { $e.CellStyle.Font = $script:BD.Kalin } })
    $g.Add_MouseDown({ param($s, $e) try { if ($s.HitTest($e.X, $e.Y).Type -eq 'None') { Sec-Temizle $s } } catch { } })
    $alt = New-Object System.Windows.Forms.FlowLayoutPanel; $alt.Dock = 'Bottom'; $alt.Height = 46; $alt.FlowDirection = 'RightToLeft'; $alt.Padding = New-Object System.Windows.Forms.Padding(8, 8, 8, 0)
    $bKapat = New-Object System.Windows.Forms.Button; $bKapat.Text = 'Close (marks all as read)'; $bKapat.Width = 170; $bKapat.Height = 28
    $bAc = New-Object System.Windows.Forms.Button; $bAc.Text = 'Open record'; $bAc.Width = 110; $bAc.Height = 28
    $alt.Controls.Add($bKapat); $alt.Controls.Add($bAc)
    $f.Controls.Add($g); $f.Controls.Add($alt)
    $script:BD = @{ F = $f; G = $g; Ac = ''; Kalin = $kalin }
    $script:BD.Sec = {
        $r = $script:BD.G.CurrentRow
        if ($null -ne $r) { $id = [string]$r.Cells[2].Value; if ($id) { $script:BD.Ac = $id; $script:BD.F.Close() } }
    }
    $bAc.Add_Click({ & $script:BD.Sec })
    $g.Add_CellDoubleClick({ param($s, $e) if ($e.RowIndex -ge 0) { & $script:BD.Sec } })
    $bKapat.Add_Click({ $script:BD.F.Close() })
    $script:DlgAcik = $true
    try { [void]$f.ShowDialog($ana) } finally { $script:DlgAcik = $false }
    $f.Dispose()
    $script:Ben = Set-BildirimOkundu $script:Ben
    Yenile-Bildirim
    $id = $script:BD.Ac
    if ($id -and (Test-Path -LiteralPath (Join-Path $script:KayitDir ($id + '.json')))) { Ac-Kayit $id }
}

# ---------------------------------------------------------------- yonetici ekranlari
function Yonet-Kullanicilar {
    $f = New-Object System.Windows.Forms.Form
    $f.Text = 'Manage users'; $f.StartPosition = 'CenterParent'; $f.Size = Get-Boyut 960 460; $f.Font = New-Object System.Drawing.Font('Segoe UI', 9); $f.ShowInTaskbar = $false; $f.MinimizeBox = $false
    $g = New-Object System.Windows.Forms.DataGridView
    Set-CiftTampon $g
    $g.Dock = 'Fill'; $g.ReadOnly = $true; $g.AllowUserToAddRows = $false; $g.RowHeadersVisible = $false; $g.BackgroundColor = [System.Drawing.Color]::White
    $g.SelectionMode = 'FullRowSelect'; $g.MultiSelect = $false; $g.AutoSizeColumnsMode = 'Fill'; $g.EnableHeadersVisualStyles = $false
    $g.ColumnHeadersDefaultCellStyle.BackColor = Get-Renk 31 56 100; $g.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White
    foreach ($c in 'Name', 'Role', 'Admin', 'Windows account', 'Computer', 'Registered', 'Last login') { [void]$g.Columns.Add($c, $c) }
    $alt = New-Object System.Windows.Forms.FlowLayoutPanel; $alt.Dock = 'Bottom'; $alt.Height = 46; $alt.FlowDirection = 'RightToLeft'; $alt.Padding = New-Object System.Windows.Forms.Padding(8, 8, 8, 0)
    $bs = @{}
    foreach ($b in @(@('Close', 90), @('Remove', 90), @('Admin on/off', 100), @('Rename', 90), @('Change role', 100))) {
        $x = New-Object System.Windows.Forms.Button; $x.Text = $b[0]; $x.Width = $b[1]; $x.Height = 28; $alt.Controls.Add($x); $bs[$b[0]] = $x
    }
    $f.Controls.Add($g); $f.Controls.Add($alt)
    $script:UD = @{ F = $f; G = $g }
    $script:UD.Doldur = {
        $g = $script:UD.G
        $g.Rows.Clear()
        foreach ($k in @(Get-TumKullanicilar)) {
            [void]$g.Rows.Add($k.Ad, $k.Rol, $(if (Test-Admin $k) { 'Yes' } else { '' }), $k.Hesap, $k.Bilgisayar, ($k.Kayit -replace 'T', ' '), ($k.SonGiris -replace 'T', ' '))
        }
        Sec-Temizle $g
    }
    $script:UD.Sec = {
        $r = $script:UD.G.CurrentRow
        if ($null -eq $r) { [void](Mesaj 'Please select a user first.'); return $null }
        return (Get-Kullanici ([string]$r.Cells[3].Value))
    }
    $g.Add_MouseDown({ param($s, $e) try { if ($s.HitTest($e.X, $e.Y).Type -eq 'None') { Sec-Temizle $s } } catch { } })
    $bs['Close'].Add_Click({ $script:UD.F.Close() })
    $bs['Change role'].Add_Click({
        Guvenli {
            $u = & $script:UD.Sec; if ($null -eq $u) { return }
            $r = Show-Girdi 'Change role' "New role for $($u.Ad):" $u.Rol $script:Roller
            if ($r) { $u.Rol = $r; Save-Kullanici $u; & $script:UD.Doldur }
        }
    })
    $bs['Rename'].Add_Click({
        Guvenli {
            $u = & $script:UD.Sec; if ($null -eq $u) { return }
            $yeni = Show-Girdi 'Rename user' "New name for $($u.Ad):" $u.Ad
            if (-not $yeni -or $yeni -eq $u.Ad) { return }
            if ($yeni.Length -lt 2) { [void](Mesaj 'The name is too short.' 'Warning'); return }
            if (Test-AdKullanimda $yeni $u.Hesap) { [void](Mesaj 'This name is already used by another user.' 'Warning'); return }
            $s = Update-KayitlardaAd $u.Ad $yeni @('TalepEden', 'AnalizYapan') $false
            $eskiAd = $u.Ad; $u.Ad = $yeni; Save-Kullanici $u
            [void](Mesaj "Renamed '$eskiAd' to '$yeni'. Updated $($s.Guncellenen) record(s)$(if ($s.Hatali) { "; $($s.Hatali) could not be updated (changed by someone else meanwhile)" })." )
            & $script:UD.Doldur
        }
    })
    $bs['Admin on/off'].Add_Click({
        Guvenli {
            $u = & $script:UD.Sec; if ($null -eq $u) { return }
            $adminler = @(Get-TumKullanicilar | Where-Object { $_.Admin })
            if ($u.Admin -and $adminler.Count -le 1) { [void](Mesaj 'At least one administrator is required.' 'Warning'); return }
            $u.Admin = (-not $u.Admin); Save-Kullanici $u; & $script:UD.Doldur
        }
    })
    $bs['Remove'].Add_Click({
        Guvenli {
            $u = & $script:UD.Sec; if ($null -eq $u) { return }
            if ($u.Hesap -eq $script:Ben.Hesap) { [void](Mesaj 'You cannot remove your own registration.' 'Warning'); return }
            if ((Mesaj "Remove $($u.Ad) ($($u.Hesap))?`r`nThe user will have to register again at the next start. Their records are kept." 'Question' 'YesNo') -ne 'Yes') { return }
            Remove-Kullanici $u.Hesap; & $script:UD.Doldur
        }
    })
    & $script:UD.Doldur
    $script:DlgAcik = $true
    try { [void]$f.ShowDialog($ana) } finally { $script:DlgAcik = $false }
    $f.Dispose()
    Yenile-Hepsi
}

function Geri-Yukle-Silinenler {
    $f = New-Object System.Windows.Forms.Form
    $f.Text = 'Restore deleted records'; $f.StartPosition = 'CenterParent'; $f.Size = Get-Boyut 820 440; $f.Font = New-Object System.Drawing.Font('Segoe UI', 9); $f.ShowInTaskbar = $false; $f.MinimizeBox = $false
    $g = New-Object System.Windows.Forms.DataGridView
    Set-CiftTampon $g
    $g.Dock = 'Fill'; $g.ReadOnly = $true; $g.AllowUserToAddRows = $false; $g.RowHeadersVisible = $false; $g.BackgroundColor = [System.Drawing.Color]::White
    $g.SelectionMode = 'FullRowSelect'; $g.MultiSelect = $false; $g.AutoSizeColumnsMode = 'Fill'; $g.EnableHeadersVisualStyles = $false
    $g.ColumnHeadersDefaultCellStyle.BackColor = Get-Renk 31 56 100; $g.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White
    foreach ($c in 'ID', 'Deleted', 'Part', 'Project', 'Requester', 'File') { [void]$g.Columns.Add($c, $c) }
    $alt = New-Object System.Windows.Forms.FlowLayoutPanel; $alt.Dock = 'Bottom'; $alt.Height = 46; $alt.FlowDirection = 'RightToLeft'; $alt.Padding = New-Object System.Windows.Forms.Padding(8, 8, 8, 0)
    $bK = New-Object System.Windows.Forms.Button; $bK.Text = 'Close'; $bK.Width = 90; $bK.Height = 28
    $bR = New-Object System.Windows.Forms.Button; $bR.Text = 'Restore'; $bR.Width = 100; $bR.Height = 28
    $alt.Controls.Add($bK); $alt.Controls.Add($bR)
    $f.Controls.Add($g); $f.Controls.Add($alt)
    $script:RD = @{ F = $f; G = $g }
    $script:RD.Doldur = {
        $script:RD.G.Rows.Clear()
        foreach ($s in @(Get-SilinenKayitlar)) { [void]$script:RD.G.Rows.Add($s.ID, $s.Silinme.ToString('yyyy-MM-dd HH:mm'), $s.Parca, $s.Proje, $s.Talep, $s.Dosya) }
        Sec-Temizle $script:RD.G
    }
    $g.Add_MouseDown({ param($s, $e) try { if ($s.HitTest($e.X, $e.Y).Type -eq 'None') { Sec-Temizle $s } } catch { } })
    $bK.Add_Click({ $script:RD.F.Close() })
    $bR.Add_Click({
        Guvenli {
            $r = $script:RD.G.CurrentRow
            if ($null -eq $r) { [void](Mesaj 'Please select a record first.'); return }
            Restore-Kayit ([string]$r.Cells[5].Value) ([string]$r.Cells[0].Value)
            & $script:RD.Doldur
        }
    })
    & $script:RD.Doldur
    $script:DlgAcik = $true
    try { [void]$f.ShowDialog($ana) } finally { $script:DlgAcik = $false }
    $f.Dispose()
    Yenile-Hepsi
}

function Devret-Analizci {
    $tum = @(Get-TumKayitlar)
    $adlar = @($tum | Where-Object { -not (Test-Kapali $_) -and $_.AnalizYapan } | ForEach-Object { $_.AnalizYapan } | Sort-Object -Unique)
    if ($adlar.Count -eq 0) { [void](Mesaj 'There are no open analyses with an assigned analyst.'); return }
    $kimden = Show-Girdi 'Reassign analyst' 'Reassign the OPEN analyses of this analyst:' '' $adlar
    if (-not $kimden) { return }
    $kime = @(Get-TumKullanicilar | Where-Object { $_.Rol -eq 'Structural Analysis Engineer' -and $_.Ad -ne $kimden } | ForEach-Object { $_.Ad })
    if ($kime.Count -eq 0) { [void](Mesaj 'There is no other registered structural analysis engineer.' 'Warning'); return }
    $hedef = Show-Girdi 'Reassign analyst' "Reassign the open analyses of $kimden to:" '' $kime
    if (-not $hedef) { return }
    $sayi = @($tum | Where-Object { -not (Test-Kapali $_) -and $_.AnalizYapan -eq $kimden }).Count
    if ((Mesaj "Reassign $sayi open analysis(es) from $kimden to ${hedef}?" 'Question' 'YesNo') -ne 'Yes') { return }
    $s = Update-KayitlardaAd $kimden $hedef @('AnalizYapan') $true
    $h = @(Get-TumKullanicilar | Where-Object { $_.Ad -eq $hedef })
    if ($s.Guncellenen -gt 0) { try { New-BildirimDosyasi '' 'Reassigned' "$($s.Guncellenen) open analysis(es) were reassigned from $kimden to you by $($script:Ben.Ad)" @($h | ForEach-Object { $_.Hesap }) } catch { } }
    [void](Mesaj "Reassigned $($s.Guncellenen) analysis(es)$(if ($s.Hatali) { "; $($s.Hatali) could not be updated (changed by someone else meanwhile) - please repeat" })." )
    Yenile-Hepsi
}

function Yedekle-Veri {
    $s = Backup-Veri
    [void](Mesaj "Backup created ($($s.Dosya) files):`r`n$($s.Yol)`r`n`r`nThe last 10 backups are kept.")
}

function Goster-Profil {
    $k = $script:Ben
    [void](Mesaj ("Name: $($k.Ad)`r`nRole: $($k.Rol)`r`nAdministrator: $(if (Test-Admin $k) { 'Yes' } else { 'No' })`r`nWindows account: $($k.Hesap)`r`nComputer: $($k.Bilgisayar)`r`nRegistered: $(($k.Kayit -replace 'T', ' '))`r`n`r`nOnly the administrator can change your name or role."))
}

function Degistir-VeriKlasoru {
    $fb = New-Object System.Windows.Forms.FolderBrowserDialog
    $fb.Description = 'Select the shared data folder (everyone must use the same folder)'
    $fb.SelectedPath = $script:Kok
    if ($fb.ShowDialog($ana) -ne 'OK') { return }
    Initialize-Veri $fb.SelectedPath
    Set-VeriKlasoru $fb.SelectedPath
    Yenile-Hepsi
}

# ---------------------------------------------------------------- olaylar
$tbYeni.Add_Click({ Guvenli { Yeni-Analiz } })
$tbDuzenle.Add_Click({ Guvenli { Duzenle-Secili } })
$tbSil.Add_Click({ Guvenli { Sil-Secili } })
$tbKopya.Add_Click({ Guvenli { Kopyala-Secili } })
$tbYenile.Add_Click({ Guvenli { Yenile-Hepsi } })
$tbExcel.Add_Click({ Guvenli { Aktar-Excel } })
$miListe.Add_Click({ Guvenli { Duzenle-Listeler } })
$miKlasor.Add_Click({ Guvenli { Degistir-VeriKlasoru } })
$miAc.Add_Click({ Guvenli { [void][Diagnostics.Process]::Start('explorer.exe', ('"' + $script:Kok + '"')) } })
$miOrnek.Add_Click({
    Guvenli {
        if ((Mesaj "Add sample records for testing?`r`n(8 master models with sub-analyses + 4 standalone analyses)`r`n(To remove them, select the records and press Delete.)" 'Question' 'YesNo') -eq 'Yes') { Add-OrnekVeri; Yenile-Hepsi }
    }
})
$miHak.Add_Click({ Guvenli { [void](Mesaj ("Analysis Tracker`r`n`r`nData folder: $($script:Kok)`r`nUser: $($script:Ben.Ad) ($($script:Ben.Rol))`r`nError log: $($script:HataLog)") ) } })
$miProfil.Add_Click({ Guvenli { Goster-Profil } })
$tbBildirim.Add_Click({ Guvenli { Goster-Bildirimler } })
$miKull.Add_Click({ Guvenli { Yonet-Kullanicilar } })
$miSilinen.Add_Click({ Guvenli { Geri-Yukle-Silinenler } })
$miDevir.Add_Click({ Guvenli { Devret-Analizci } })
$miYedek.Add_Click({ Guvenli { Yedekle-Veri } })

# bos alana tiklaninca mavi secim kalksin (tum ekranlar)
foreach ($gg in @($grid, $gDurum, $gKisi, $gYakin)) {
    $gg.Add_MouseDown({ param($s, $e) try { if ($s.HitTest($e.X, $e.Y).Type -eq 'None') { Sec-Temizle $s } } catch { } })
}
$lv.Add_MouseDown({ param($s, $e) try { if ($null -eq $s.GetItemAt($e.X, $e.Y)) { $s.SelectedItems.Clear() } } catch { } })
$bosTikla = { foreach ($gg in @($gDurum, $gKisi, $gYakin, $grid)) { Sec-Temizle $gg }; try { $lv.SelectedItems.Clear() } catch { } }
foreach ($ct in @($tp1, $tp2, $tp3, $fp, $ozetTablo, $fpAgac)) { $ct.Add_Click($bosTikla) }

$grid.Add_CellDoubleClick({ param($s, $e) if ($e.RowIndex -ge 0) { Guvenli { Duzenle-Secili } } })
$grid.Add_KeyDown({ param($s, $e) if ($e.KeyCode -eq 'Return') { $e.Handled = $true; $e.SuppressKeyPress = $true; Guvenli { Duzenle-Secili } } })
$grid.Add_CellFormatting({
    param($s, $e)
    if ($e.RowIndex -lt 0) { return }
    $drv = $s.Rows[$e.RowIndex].DataBoundItem
    if ($null -eq $drv) { return }
    $durum = [string]$drv['Durum']
    if ($script:Kapali -contains $durum) { $e.CellStyle.ForeColor = [System.Drawing.Color]::Gray }
    $kol = $s.Columns[$e.ColumnIndex].Name
    if ($kol -eq 'Gecikme' -and [string]$drv['Gecikme']) { $e.CellStyle.ForeColor = [System.Drawing.Color]::White; $e.CellStyle.BackColor = Get-Renk 192 0 0 }
    if ($kol -eq 'Rol' -and [string]$drv['Rol'] -eq 'Master Model') { if ($null -eq $script:FontKalin) { $script:FontKalin = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold) }; $e.CellStyle.Font = $script:FontKalin; $e.CellStyle.ForeColor = Get-Renk 31 56 140 }
})
$lv.Add_DoubleClick({ if ($lv.SelectedItems.Count -gt 0) { Guvenli { Duzenle-Secili } } })
$lv.Add_KeyDown({ param($s, $e) if ($e.KeyCode -eq 'Return' -and $s.SelectedItems.Count -gt 0) { $e.Handled = $true; Guvenli { Duzenle-Secili } } })
$agacDegisti = { Guvenli { Yenile-Agac } }
$txtAgac.Add_TextChanged($agacDegisti)
$chkAgac.Add_CheckedChanged($agacDegisti)

$filtreDegisti = { if (-not $script:FiltreKilit) { Guvenli { Yenile-Liste } } }
$txtAra.Add_TextChanged($filtreDegisti)
$cmbDurum.Add_SelectedIndexChanged($filtreDegisti)
$cmbKisi.Add_SelectedIndexChanged($filtreDegisti)
$chkAcik.Add_CheckedChanged($filtreDegisti)
$chkGec.Add_CheckedChanged($filtreDegisti)
$btnTemizle.Add_Click({
    $script:FiltreKilit = $true
    $txtAra.Text = ''; $chkAcik.Checked = $false; $chkGec.Checked = $false
    $cmbDurum.SelectedItem = 'All'; $cmbKisi.SelectedItem = 'All'
    # basliga tiklayarak degisen siralama da sifirlanir: en son girilen talep en ustte
    $dt.DefaultView.Sort = 'ID DESC'
    try { $grid.ClearSelection() } catch { }
    $script:FiltreKilit = $false
    Guvenli { Yenile-Liste }
})

$ana.KeyPreview = $true
$ana.Add_KeyDown({
    param($s, $e)
    if ($e.Control -and $e.KeyCode -eq 'N') { $e.Handled = $true; Guvenli { Yeni-Analiz } }
    elseif ($e.KeyCode -eq 'F5') { $e.Handled = $true; Guvenli { Yenile-Hepsi } }
})

$zamanlayici = New-Object System.Windows.Forms.Timer
$zamanlayici.Interval = 20000
$zamanlayici.Add_Tick({ Guvenli { Otomatik-Yenile } })

$ana.Add_Shown({
    Guvenli { Yenile-Hepsi }
    Guvenli {
        if ($script:Tum.Count -eq 0 -and (Test-Admin $script:Ben)) {
            if ((Mesaj "There are no analysis records yet.`r`n`r`nAdd sample data (8 master models with sub-analyses and 4 standalone analyses) so you can try the program?`r`n(You can delete them later, or add them from Admin > Add sample data.)" 'Question' 'YesNo') -eq 'Yes') { Add-OrnekVeri; Yenile-Hepsi }
        }
    }
    Guvenli { Yenile-Bildirim }
    $zamanlayici.Start()
    $ana.Activate()
})

# ---------------------------------------------------------------- başlat
try {
    $veri = if ($VeriKlasoru) { $VeriKlasoru } else { Get-VeriKlasoru }
    try { Initialize-Veri $veri }
    catch {
        Write-HataLog $_.Exception.ToString()
        [void](Mesaj ("Cannot access the data folder:`r`n$veri`r`n`r`n" + $_.Exception.Message + "`r`n`r`nYou can choose another folder in the next window.") 'Warning')
        $fb = New-Object System.Windows.Forms.FolderBrowserDialog
        $fb.Description = 'Select the shared data folder'
        if ($fb.ShowDialog() -ne 'OK') { exit 1 }
        $veri = $fb.SelectedPath
        Initialize-Veri $veri
        Set-VeriKlasoru $veri
    }
    $script:Ben = Initialize-Kimlik
    if ($null -eq $script:Ben) { exit 0 }
    try { Remove-EskiBildirimler } catch { }
    Yetki-Uygula
    try {
        $script:Tray = New-Object System.Windows.Forms.NotifyIcon
        $script:Tray.Icon = [System.Drawing.SystemIcons]::Information
        $script:Tray.Text = 'Analysis Tracker'
        $script:Tray.Visible = $true
        $script:Tray.Add_BalloonTipClicked({ Guvenli { Goster-Bildirimler } })
    } catch { $script:Tray = $null }
    $ana.Add_FormClosed({ try { if ($null -ne $script:Tray) { $script:Tray.Visible = $false; $script:Tray.Dispose() } } catch { } })
    [void][System.Windows.Forms.Application]::Run($ana)
} catch {
    Write-HataLog $_.Exception.ToString()
    [void][System.Windows.Forms.MessageBox]::Show("Unexpected error:`r`n`r`n" + $_.Exception.Message + "`r`n`r`nDetails: " + $script:HataLog, 'Analysis Tracker', 'OK', 'Error')
    exit 1
}
