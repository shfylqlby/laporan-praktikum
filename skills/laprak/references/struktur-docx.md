# Struktur template

> **Catatan format umum.** Angka di bawah adalah hasil pembedahan dokumen acuan. Generator kini
> menimpa sebagian di semua profil sesuai Format umum di `SKILL.md`: TNR 12 eksplisit (termasuk
> nomor heading), before/after 0 di semua paragraf, indent 720 twip (36 pt) untuk paragraf,
> heading, daftar, dan caption tabel, caption gambar spasi 1,5, caption tabel spasi 1,
> tabel data `Table Grid` selebar teks (7938 dxa) tanpa warna (header tebal hitam, tanpa shading)
> dengan isi spasi 1,5, dan label SEQ
> `Gambar_N.` di kedua profil (sama dengan Word InsertCaption). Template eldas tidak punya
> style `Table Grid`; generator menyalinnya dari srwe saat build. Berkas di `assets/` tidak diubah.

## SRWE (hasil pembedahan `Format Laporan Prak SRWE Modul 4.docx`)

Dokumen acuan: 185 paragraf, 6 tabel, 46 gambar, 27 halaman.

## Halaman

A4 21 x 29,7 cm. Margin kiri 4, kanan 3, atas 4, bawah 3 cm. Header dan footer kosong, jarak 1,25 cm.

## docDefaults

`Calibri sz 24`, spacing after 160 line 278 auto. Badan dokumen selalu menimpa dengan Times New Roman eksplisit pada tiap run, jadi jangan mengandalkan default.

## Style

| Style | Isi |
|---|---|
| `Normal` | tanpa format khusus |
| `Heading1` | TNR tebal, rata tengah, before 360 after 80, ind left 432 hanging 432, outlineLvl 0 |
| `Heading2` | TNR tebal, justify, line 360 auto, ind left 576 hanging 576 (paragraf menimpa jadi 709/709) |
| `Heading3` | TNR tebal, justify, line 360 auto, ind left 720 hanging 720 (paragraf menimpa jadi 709/709) |
| `Caption` | TNR 10 pt, keepNext, line 240 auto, after 0, rata kiri |
| `captiongambar` | custom, TNR 10 pt, rata tengah, line 1,5, after 0 |
| `ListParagraph` | ind left 720, contextualSpacing |

Pemakaian pada dokumen acuan: Normal 69, List Paragraph 55, captiongambar 43, Caption 7, Heading 3 6, Heading 2 4, Heading 1 1.

## Penomoran

- `numId 1` → abstract 6, ilvl 0 `BAB %1` upperRoman start 5. Dipakai judul BAB.
- `numId 6` dan `numId 7` → ilvl 1 `5.%2`, ilvl 2 `5.2.%3`. Dipakai Heading 2 dan Heading 3 percobaan.
- `numId 2` → `5.3.%1`. Dipakai Heading 3 pada Hasil dan Analisis.
- `numId 8/12/16` → daftar langkah decimal `%1.`, ind left 720 hanging 360.
- `numId 10` → daftar tujuan, left/hanging 709.
- `numId 14` → daftar kesimpulan.

Generator menambal seluruh `w:lvlText` yang memuat `5.` menjadi nomor BAB yang diminta, dan menyetel `w:start` level Romawi ke nomor itu. Daftar bernomor dibuat baru per blok supaya tiap percobaan mulai dari 1.

## Paragraf

- Judul BAB: `Heading 1` + numPr(1,0), spacing before 0 after 0 line 360 auto, ind left 0 firstLine 0, diawali `<w:br/>` sebelum teks judul sehingga "BAB V" dan judul berada pada dua baris.
- Butir tujuan/langkah/kesimpulan: `List Paragraph`, after 0, line 360 auto, justify, ind left 720 hanging 720.
- Paragraf gambar: `List Paragraph`, keepNext, after 0, line 240 auto, ind left 0, rata tengah, berisi satu inline picture.
- Paragraf analisis: `Normal`, after 0, line 360 auto, justify, ind firstLine 709.

## Caption

```xml
<w:r><w:t>Gambar 5.</w:t></w:r>
<w:fldSimple w:instr=" SEQ Gambar_5. \* ARABIC "><w:r><w:rPr><w:noProof/></w:rPr><w:t>1</w:t></w:r></w:fldSimple>
<w:r><w:t xml:space="preserve"> Topologi Percobaan 1</w:t></w:r>
```

Tabel memakai `SEQ Tabel_5.` dengan style `Caption`.

## Tabel data

`Table Grid`, `tblW 7938 dxa`, `jc center`, border single sz 4 hitam penuh. Lebar kolom acuan `[2.9, 2.1, 3.0, 3.1, 2.9]` cm. Sel: `vAlign center`, paragraf rata tengah, spacing before 40 after 40 twips. Header: `shd fill FF89FF`, run tebal warna `FFFFFF`.

## Tabel kode

`Table Grid`, `tblW 0 auto`, `tblInd 709 dxa`, satu sel. Tiap baris perintah jadi satu paragraf, run Courier New `sz 20` (10 pt), rata kanan kiri.

## Gambar

Lebar bervariasi, maksimum tercatat 11,67 cm, tinggi maksimum 12,74 cm. Rasio asli dipertahankan, jangan dipaksa ke satu lebar tetap.

## Urutan bab

```
BAB V / judul modul
5.1 Tujuan            daftar bernomor
    page break
5.2 Percobaan         5.2.1 .. 5.2.n  (caption tabel, tabel, langkah, kode, gambar, caption)
    page break
5.3 Hasil dan Analisis 5.3.1 .. 5.3.n (paragraf, gambar opsional)
    page break
5.4 Kesimpulan        daftar bernomor
```

Pemisah halaman memakai `pageBreakBefore` pada Heading 2 berikutnya, bukan paragraf kosong berisi `w:br type=page` seperti dokumen acuan. Hasil cetak sama tetapi tanpa risiko halaman kosong.

## Inkonsistensi dokumen acuan yang sengaja dirapikan

- Ada satu caption gambar memakai style `Normal`, bukan `captiongambar` (paragraf 172).
- Daftar langkah memakai tiga numId berbeda (8, 12, 16) dengan `start` berbeda, termasuk satu yang mulai dari 2.
- Sebagian paragraf analisis memakai indent 709 dan sebagian 720.
- Beberapa paragraf `Heading 2` menyimpan `rPr color EE0000` sisa penyuntingan.

Generator menyeragamkan semuanya.

---

# Template kedua: ELDAS2026_MODUL4_Kel26.docx

Dokumen acuan: 160 paragraf, 2 tabel, 20 gambar, 20 halaman. Halaman dan margin sama
persis dengan SRWE (A4, 4/3/4/3 cm, header-footer kosong).

## Beda pokok dari SRWE

| Hal | SRWE | ELDAS |
|---|---|---|
| `docDefaults` | Calibri sz 24 | Times New Roman sz 24, line 360 auto, ind left 357 |
| Judul | `Heading 1` + numPr `BAB %1` upperRoman | `Heading 1` tanpa numPr, teks langsung, dua baris dipisah `w:br` |
| Tingkat heading | tiga (`Heading 1..3`) | empat (`Heading 1..4`), `5.1` sampai `5.1.1.1` |
| Indent heading | left/hanging 709 | left/hanging 851 |
| Caption gambar | style `captiongambar` | style `Caption` |
| Instruksi SEQ | `SEQ Gambar_5. \* ARABIC` | `SEQ Gambar_5._ \* ARABIC` (ada garis bawah tambahan) |
| Gaya tabel | `Table Grid`, `tblW 7938 dxa` | `Normal Table` + `tblBorders` eksplisit, `tblW 0 auto` |
| Header tabel | fill `FF89FF`, teks putih | fill `E8EEF5`, teks hitam tebal |
| Gambar maksimum | 11,67 x 12,74 cm | 14,0 x 11,92 cm |
| Daftar | left/hanging 720 | left/hanging 851 |
| Struktur | Tujuan, Percobaan, Hasil dan Analisis, Kesimpulan | Tujuan, Alat dan Bahan, Langkah Kerja, Hasil Percobaan dan Analisis, Kesimpulan |
| Tabel kode | ada | tidak ada |
| Tabel hasil | hanya di Percobaan | juga di Hasil dan Analisis |

## Yang sama

Caption tetap memakai field `SEQ`. Paragraf prosa tetap TNR 12 pt, rata kanan kiri,
spasi 1,5, indent baris pertama sekitar 709 twips. Paragraf gambar tetap `keepNext`,
rata tengah, indent 0. Pemisah antar bagian utama tetap halaman baru.

## Bagian Alat dan Bahan

Pola berulang: heading tingkat dua berisi nama alat, paragraf gambar, caption,
satu paragraf penjelasan, lalu satu paragraf berisi URL sumber dalam tanda kurung
dengan gaya tautan. Generator menempatkan URL sebagai butir terakhir `paragraf_akhir`.

## Inkonsistensi dokumen acuan yang dirapikan generator

- Sebagian heading diketik manual dengan nomor dan tab (`5.2.2. \tLaptop`) sehingga
  tidak ikut penomoran otomatis. Satu heading bahkan bergaya `Normal`.
- Indent baris pertama paragraf berpindah-pindah antara 630, 709, 714, 720, dan 851.
- `Normal (Web)` muncul di bagian analisis sebagai sisa tempel dari peramban.
- Penomoran langkah memakai `numId` berbeda tiap percobaan dengan `ilvl` yang tidak
  konsisten (0, 3, dan 6).
