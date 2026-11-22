# Edit laporan langsung di Microsoft Word lewat COM (Windows).
#
# Tanpa -Docx: menempel ke Word yang sedang terbuka (GetActiveObject) dan memakai dokumen aktif,
# lalu menimpa berkas itu. Dengan -Docx: membuka berkas tersebut (memakai jendela Word yang sudah
# membukanya bila ada, atau Word tersembunyi baru).
# Sebelum aksi yang mengubah dokumen, berkas di disk disalin ke .bak (tidak pernah menimpa .bak lama).
#
# Aksi:
#   info      daftar heading, gambar, tabel, dan caption beserta indeksnya
#   backup    hanya membuat .bak
#   caption   pasang caption Word asli (Range.InsertCaption) pada gambar atau tabel ke-N
#             -Jenis Gambar|Tabel -Indeks N -Bab B -Teks "Keterangan"
#   update    perbarui semua field (nomor caption, daftar isi) lalu simpan
#
# Contoh:
#   powershell -NoProfile -ExecutionPolicy Bypass -File word_live.ps1 -Aksi info
#   powershell -NoProfile -ExecutionPolicy Bypass -File word_live.ps1 -Aksi caption -Jenis Gambar -Indeks 3 -Bab 6 -Teks "Topologi Percobaan 1"
param(
    [ValidateSet('info', 'backup', 'caption', 'update')][string]$Aksi = 'info',
    [string]$Docx,
    [ValidateSet('Gambar', 'Tabel')][string]$Jenis,
    [int]$Indeks,
    [int]$Bab,
    [string]$Teks,
    [switch]$TanpaSimpan
)
$ErrorActionPreference = 'Stop'

# ---------- sambung ke Word ----------
$milikSendiri = $false
$w = $null
try { $w = [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application') } catch { }

if ($Docx) {
    $Docx = (Resolve-Path $Docx).Path
    $doc = $null
    if ($w) { foreach ($d in $w.Documents) { if ($d.FullName -eq $Docx) { $doc = $d } } }
    if (-not $doc) {
        $w = New-Object -ComObject Word.Application
        $w.Visible = $false
        $w.DisplayAlerts = 0
        $milikSendiri = $true
        $doc = $w.Documents.Open($Docx, $false, $false)
    }
}
else {
    if (-not $w -or $w.Documents.Count -eq 0) { throw 'Tidak ada dokumen Word yang terbuka. Buka laporannya di Word atau pakai -Docx.' }
    $doc = $w.ActiveDocument
}

function Buat-Backup {
    if (-not $doc.Path) { throw 'Dokumen belum pernah disimpan; simpan dulu supaya bisa dibuat .bak.' }
    $asal = $doc.FullName
    $bak = "$asal.bak"
    if (Test-Path $bak) { $bak = "$asal.$(Get-Date -Format yyyyMMdd-HHmmss).bak" }
    [IO.File]::Copy($asal, $bak)
    Write-Output "backup: $bak"
}

function Rapikan-Caption($p, $gambar) {
    $r = $p.Range
    $r.Style = -35                       # wdStyleCaption, tidak bergantung bahasa Word
    $r.Font.Name = 'Times New Roman'
    $r.Font.Size = 10
    $r.Font.Italic = 0
    $r.Font.Bold = 0
    $r.Font.Color = 0                    # hitam
    $f = $r.ParagraphFormat
    $f.SpaceBefore = 0
    $f.SpaceAfter = 0
    if ($gambar) {
        $f.Alignment = 1                 # tengah
        $f.LineSpacingRule = 1           # 1,5
        $f.LeftIndent = 0
        $f.FirstLineIndent = 0
    }
    else {
        $f.Alignment = 0                 # kiri
        $f.LineSpacingRule = 0           # 1
        $f.LeftIndent = 36
        $f.FirstLineIndent = -36
        $f.KeepWithNext = -1
    }
}

function Cari-Caption($gambar, $obj) {
    # Paragraf tepat di bawah gambar atau tepat di atas tabel, bila memuat field SEQ jenis ini.
    if ($gambar) {
        $akhir = $obj.Range.Paragraphs.Item(1).Range.End
        if ($akhir -ge $doc.Content.End) { return $null }
        $p = $doc.Range($akhir, $akhir).Paragraphs.Item(1)
    }
    else {
        $awal = $obj.Range.Start
        if ($awal -le 0) { return $null }
        $p = $doc.Range($awal - 1, $awal - 1).Paragraphs.Item(1)
    }
    foreach ($f in $p.Range.Fields) {
        if ($f.Code.Text -match "SEQ\s+$($Jenis)_") { return $p }
    }
    return $null
}

function Pasang-Caption {
    if (-not $Jenis -or -not $Indeks -or -not $Bab -or -not $Teks) { throw 'caption butuh -Jenis, -Indeks, -Bab, dan -Teks.' }
    $gambar = $Jenis -eq 'Gambar'
    $label = "$Jenis $Bab."
    $ada = $false
    foreach ($l in $w.CaptionLabels) { if ($l.Name -eq $label) { $ada = $true } }
    if (-not $ada) {
        $l = $w.CaptionLabels.Add($label)
        $l.NumberStyle = 0
        $l.IncludeChapterNumber = $false
    }

    if ($gambar) {
        if ($Indeks -gt $doc.InlineShapes.Count) { throw "gambar ke-$Indeks tidak ada (hanya $($doc.InlineShapes.Count))" }
        $obj = $doc.InlineShapes.Item($Indeks)
    }
    else {
        if ($Indeks -gt $doc.Tables.Count) { throw "tabel ke-$Indeks tidak ada (hanya $($doc.Tables.Count))" }
        $obj = $doc.Tables.Item($Indeks)
    }

    $ada = Cari-Caption $gambar $obj
    if ($ada) {
        Write-Output "lewati: $Jenis ke-$Indeks sudah punya caption: $($ada.Range.Text.Trim())"
        $script:berubah = $false
        return
    }

    Buat-Backup
    $posisi = if ($gambar) { 1 } else { 0 }   # 1 di bawah gambar, 0 di atas tabel
    $obj.Range.InsertCaption($label, " $Teks", '', $posisi)

    $cap = Cari-Caption $gambar $obj
    if (-not $cap) { throw 'caption tersisip tetapi paragrafnya tidak ditemukan; periksa dokumen di Word.' }
    # Word menyisipkan spasi antara label dan nomor ("Gambar 6. 1"); hapus supaya jadi "Gambar 6.1".
    # Smart cut and paste mengembalikan spasi itu, jadi dimatikan sebentar lalu dipulihkan.
    $smart = $w.Options.SmartCutPaste
    $w.Options.SmartCutPaste = $false
    try {
        foreach ($f in $cap.Range.Fields) {
            if ($f.Code.Text -match 'SEQ') {
                $cs = $f.Code.Start
                $spasi = $doc.Range($cs - 2, $cs - 1)
                if ($spasi.Text -eq ' ') { [void]$spasi.Delete() }
            }
        }
    }
    finally { $w.Options.SmartCutPaste = $smart }
    Rapikan-Caption $cap $gambar
    if ($gambar) {
        $objPar = $obj.Range.Paragraphs.Item(1)
        $pf = $objPar.Format
        $pf.SpaceBefore = 0
        $pf.SpaceAfter = 0
        $pf.LineSpacingRule = 0
        $pf.KeepWithNext = -1
    }
    [void]$cap.Range.Fields.Update()
    Write-Output "caption: $($cap.Range.Text.Trim())"
}

function Perbarui-Field {
    Buat-Backup
    foreach ($story in $doc.StoryRanges) {
        $r = $story
        while ($r -ne $null) { [void]$r.Fields.Update(); $r = $r.NextStoryRange }
    }
    foreach ($toc in $doc.TablesOfContents) { [void]$toc.Update() }
    Write-Output 'field diperbarui'
}

function Tampil-Info {
    Write-Output "dokumen: $($doc.FullName)"
    Write-Output '-- heading'
    foreach ($p in $doc.Paragraphs) {
        if ($p.OutlineLevel -le 4) { Write-Output ("  {0} {1}" -f $p.Range.ListFormat.ListString, $p.Range.Text.Trim()) }
    }
    Write-Output "-- gambar ($($doc.InlineShapes.Count))"
    for ($i = 1; $i -le $doc.InlineShapes.Count; $i++) {
        $n = $doc.InlineShapes.Item($i).Range.Paragraphs.Item(1).Next(1)
        $c = if ($n) { $n.Range.Text.Trim() } else { '' }
        if ($c -notlike 'Gambar *') { $c = '(tanpa caption)' }
        Write-Output "  [$i] $c"
    }
    Write-Output "-- tabel ($($doc.Tables.Count))"
    for ($i = 1; $i -le $doc.Tables.Count; $i++) {
        $t = $doc.Tables.Item($i)
        $sel = $t.Cell(1, 1).Range.Text -replace '[\r\a]', ''
        $prev = $t.Range.Paragraphs.Item(1).Previous(1)
        $c = if ($prev) { $prev.Range.Text.Trim() } else { '' }
        if ($c -notlike 'Tabel *') { $c = '(tanpa caption)' }
        $jenis = if ($t.Rows.Count -eq 1 -and $t.Columns.Count -eq 1) { 'kode/1x1' } else { "$($t.Rows.Count)x$($t.Columns.Count)" }
        Write-Output "  [$i] $jenis sel1='$($sel.Substring(0, [Math]::Min(30, $sel.Length)))' $c"
    }
}

$script:berubah = $Aksi -in 'caption', 'update'
try {
    switch ($Aksi) {
        'info' { Tampil-Info }
        'backup' { Buat-Backup }
        'caption' { Pasang-Caption }
        'update' { Perbarui-Field }
    }
    if ($script:berubah -and -not $TanpaSimpan) {
        $doc.Save()
        Write-Output "disimpan: $($doc.FullName)"
    }
}
finally {
    if ($milikSendiri) {
        $doc.Close(0)
        $w.Quit()
    }
}
