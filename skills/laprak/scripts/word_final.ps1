# Finalisasi laporan lewat Microsoft Word (Windows): perbarui semua field
# (nomor caption SEQ, daftar isi), simpan .docx, dan opsional ekspor PDF untuk dicek.
#
# Pakai:
#   powershell -NoProfile -ExecutionPolicy Bypass -File word_final.ps1 -Docx Laporan.docx [-Pdf cek.pdf]
#
# Setelah skrip ini, pengguna tidak perlu lagi menekan Ctrl+A lalu F9.
param(
    [Parameter(Mandatory = $true)][string]$Docx,
    [string]$Pdf
)
$ErrorActionPreference = 'Stop'
$Docx = (Resolve-Path $Docx).Path
if ($Pdf) { $Pdf = [IO.Path]::GetFullPath($Pdf) }

$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
try {
    $d = $w.Documents.Open($Docx, $false, $false)
    try {
        foreach ($story in $d.StoryRanges) {
            $r = $story
            while ($r -ne $null) {
                [void]$r.Fields.Update()
                $r = $r.NextStoryRange
            }
        }
        foreach ($toc in $d.TablesOfContents) { [void]$toc.Update() }
        $d.Save()
        if ($Pdf) {
            if (Test-Path $Pdf) { Remove-Item $Pdf -Force }
            $d.ExportAsFixedFormat($Pdf, 17)
            Write-Output "pdf: $Pdf"
        }
        $caps = @()
        foreach ($p in $d.Paragraphs) {
            $t = $p.Range.Text.Trim()
            if ($t -match '^(Gambar|Tabel)\s+\d+\.\d+') { $caps += $t }
        }
        Write-Output "caption diperbarui: $($caps.Count)"
    }
    finally { $d.Close(0) }
}
finally {
    $w.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($w)
}
Write-Output "docx: $Docx"
