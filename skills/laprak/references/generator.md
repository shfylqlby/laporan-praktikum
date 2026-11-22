# Generator laporan (Jalur B)

> Rujukan untuk Jalur B di `../SKILL.md`. Dulu ini SKILL.md skill generator yang terpisah;
> sejak versi 1.2.2 isinya menjadi bagian skill `laprak`. Path `scripts/`, `assets/`, dan
> `references/` di bawah relatif terhadap folder skill `laprak`.

Hasilkan `.docx` yang identik struktur dengan template resmi. Yang berganti antar modul
hanya **nomor BAB, judul, dan isi**. Format tidak diubah.

## Profil

| Profil | Template asal | Ciri |
|---|---|---|
| `srwe` | Format Laporan Prak SRWE Modul 4 | judul berawalan "BAB V", caption gambar `captiongambar`, gambar maks 11,67 cm |
| `eldas` | ELDAS2026_MODUL4_Kel26 | judul tanpa "BAB", seluruh caption style `Caption`, gambar maks 14 cm, heading sampai empat tingkat |

Pilih lewat `"profil"` di spec. Daftar lengkap kunci ada di `assets/profil.json`;
timpa sebagian lewat `"profil_override"` kalau dosen minta variasi kecil.

Profil baru: salin `.docx` acuan ke `assets/`, tambahkan entri di `assets/profil.json`,
jalankan `scripts/selfcheck.py`.

## Format umum (semua profil, semua mata kuliah)

Generator menerapkan ini tanpa perlu diatur di spec. Bila dosen meminta lain, ketentuan dosen
yang dipakai lewat `profil_override`.

| Unsur | Format |
|---|---|
| Halaman | A4, margin kiri 4, atas 4, kanan 3, bawah 3 cm |
| Teks | Times New Roman 12 pt di semua paragraf, heading, nomor heading, daftar, dan isi tabel |
| Spasi | before dan after **0** di semua paragraf; spasi baris 1,5 kecuali yang disebut lain |
| Page break | setiap sub-bab level 2 (x.2, x.3, ...) mulai di halaman baru; level 3 tidak |
| Paragraf | rata kiri-kanan, indent baris pertama 36 pt |
| Heading, daftar, caption tabel | indent menggantung li 36 pt fi −36 pt; nomor mepet kiri, teks mulai di 36 pt |
| Caption gambar | di bawah gambar, TNR 10 hitam, tidak miring, rata tengah, spasi 1,5 |
| Caption tabel | di atas tabel, TNR 10 hitam, tidak miring, spasi 1 |
| Nomor caption | "Gambar 6.1", field `SEQ Gambar_6.` (sama dengan Word InsertCaption berlabel "Gambar 6."), tanpa spasi sebelum nomor |
| Paragraf gambar | spasi 1, rata tengah |
| Tabel data | `Table Grid` selebar teks, tanpa warna (tanpa shading, teks hitam), header tebal, isi rata tengah, spasi 1,5, tanpa indent |
| Kode CLI | tabel 1x1 `Table Grid`, Courier New 10, spasi 1, tidak miring |
| Miring | istilah asing di prosa saja. Heading, caption, tabel, kode, dan nama merek (Cisco, Packet Tracer) tidak miring |
| Gambar belum ada | kotak abu-abu bergaris putus berisi `[ISI: gambar ...]`, caption tetap ada |

## Aturan isi

- Jangan menambah bagian Dasar Teori kecuali diminta eksplisit.
- Langkah kerja tiap percobaan dipisah (Percobaan 1, Percobaan 2, ...), tidak digabung jadi satu daftar.
- Tiap butir Tujuan cukup dua kalimat.
- Kesimpulan murni menjawab Tujuan: tanpa penanda `[ISI: ...]` dan tanpa data kelompok (alamat IP, nilai ukur).
- Bagian yang datanya belum ada (analisis, screenshot) tetap ditulis dengan penanda; struktur template tidak dihapus.

Lint memeriksa empat aturan pertama dan mencetak `PERINGATAN` bila dilanggar.

## Alur

Path `scripts/`, `assets/`, `references/` di bawah relatif terhadap folder skill ini, bukan folder kerja pengguna.

1. Baca modul (markdown/PDF/docx) dan daftar gambar yang diberikan pengguna.
2. Tentukan profil. Kalau ragu, tanya. Jangan menebak.
3. Petakan gambar ke langkah percobaan. Urutan berkas biasanya urutan langkah;
   konfirmasi kalau tidak jelas.
4. Tulis `spec.json` sesuai `references/spec-schema.md`.
5. Tulis prosa Indonesia sesuai `references/bahasa-id.md`. **Wajib dibaca**, bagian
   ini yang paling sering gagal.
6. Bangun:
   ```bash
   python scripts/build_laprak.py spec.json -o "Laporan-Modul-N.docx"
   ```
   Butuh `python-docx` dan `pillow`. Setelah build, `scripts/lint_laprak.py` berjalan otomatis
   dan mencetak `PERINGATAN` untuk penanda `[ISI: ...]` yang tersisa, rujukan Gambar atau
   Tabel yang nomornya tidak ada, gambar atau tabel yang tidak dirujuk, dan pola gaya.
   Tangani tiap temuan satu per satu, jangan diganti otomatis. Gambar yang belum ada diganti
   kotak `[ISI: gambar ...]`; tambahkan `--ketat` bila ingin build berhenti.
7. Finalisasi dan verifikasi visual. Pilih sesuai mesin:
   - **Windows dengan Microsoft Word** (paling baik): perbarui semua field, simpan, dan ekspor PDF.
     Pengguna tidak perlu lagi menekan Ctrl+A lalu F9.
     ```bash
     powershell -NoProfile -ExecutionPolicy Bypass -File scripts/word_final.ps1 -Docx Laporan-Modul-N.docx -Pdf cek.pdf
     ```
     Dari WSL, salin dulu ke folder Windows dan panggil `powershell.exe` dengan path Windows.
   - **Tanpa Word** (Linux, sandbox): `soffice --headless --convert-to pdf Laporan-Modul-N.docx`
     lalu `pdftoppm -png -r 70 Laporan-Modul-N.pdf cek/p`. Nomor caption di render ini bisa
     masih "1"; Word menawarkan pembaruan field saat dokumen dibuka (`updateFields` sudah dipasang).
   Baca PDF atau PNG hasilnya dan periksa halaman berisi tabel, kode, dan gambar.

### Mengedit dokumen yang sudah terbuka di Word (Windows)

Bila pengguna sudah punya laporan atau template dosen yang terbuka di Word dan minta dilengkapi
di tempat, jangan bangun ulang dari spec. Pakai `scripts/word_live.ps1`, yang menempel ke Word
lewat COM (GetActiveObject) dan menimpa berkas yang terbuka:

```bash
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/word_live.ps1 -Aksi info
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/word_live.ps1 -Aksi caption -Jenis Gambar -Indeks 3 -Bab 6 -Teks "Topologi Percobaan 1"
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/word_live.ps1 -Aksi update
```

- `info` mendaftar heading, gambar, dan tabel beserta indeks dan caption-nya.
- `caption` memakai `Range.InsertCaption` asli berlabel "Gambar 6." atau "Tabel 6.", menghapus
  spasi sebelum nomor, lalu menerapkan format caption di atas. Gambar atau tabel yang sudah
  punya caption dilewati.
- `update` memperbarui semua field lalu menyimpan.
- Sebelum mengubah apa pun, skrip menyalin berkas ke `.bak` (tidak pernah menimpa `.bak` lama).
- Tabel data (alamat IP, VLAN) tidak diubah isinya; cek `info` sebelum memasang caption supaya
  indeks tabel tidak tertukar dengan tabel kode.
- Isi teks lain ditulis lewat COM Word yang sama, mengikuti Format umum dan Aturan isi di atas.
- `-Docx berkas.docx` membuka berkas tertentu bila Word belum membukanya.

## Aturan yang tidak boleh dilanggar

- Nomor gambar dan tabel memakai field `SEQ`, jangan diketik manual. Nomor baru muncul
  setelah field diperbarui di Word (`word_final.ps1`, atau Word menawarkannya saat dibuka).
- Data yang belum ada ditulis `[ISI: ...]`, sumber yang belum pasti `[CEK SUMBER: ...]`.
  Jangan mengarang angka hasil percobaan.
- Nomor BAB menyetir seluruh penomoran. Cukup set `"bab": N`; skrip menambal `numbering.xml`.
- Ukuran gambar dibatasi profil, rasio asli dijaga, rata tengah, caption tepat di bawahnya.
- Caption tabel diletakkan **di atas** tabel.
- Heading tidak pernah miring walaupun memuat istilah asing.
- Istilah asing dimiringkan di prosa. `*teks*` memaksa miring.
- Header dan footer dibiarkan kosong, sama seperti template.
- Jangan menulis `.docx` dari nol. Selalu lewat template di `assets/`, karena di situlah
  styles, numbering, `sectPr`, dan font tertanam berada.

## Berkas

- `assets/srwe.docx`, `assets/eldas.docx` — template kosong hasil pelucutan dokumen resmi. Jangan ditimpa.
- `assets/profil.json` — parameter format tiap profil.
- `scripts/build_laprak.py` — generator.
- `scripts/lint_laprak.py` — pemeriksa rujukan, penanda, dan gaya; juga bisa dipakai pada .docx lain: `python scripts/lint_laprak.py Laporan.docx`.
- `scripts/word_final.ps1` — perbarui field, simpan, dan ekspor PDF lewat Word (Windows).
- `scripts/word_live.ps1` — edit dokumen yang terbuka di Word: backup `.bak`, caption InsertCaption, update field (Windows).
- `scripts/selfcheck.py` — uji mandiri seluruh profil.
- `references/spec-schema.md` — bentuk spec.
- `references/struktur-docx.md` — angka format hasil pembedahan kedua template.
- `references/bahasa-id.md` — kaidah menulis agar tidak terbaca sebagai keluaran AI.

## Uji mandiri

```bash
python scripts/selfcheck.py
```
Membangun dokumen contoh tiap profil lalu memeriksa style, penomoran, tabel tanpa warna,
ukuran gambar, field SEQ, pemisah halaman, kemiringan heading, spasi kode, penanda gambar
kosong, nomor persamaan, urutan pustaka, dan hasil lint. Format umum diperiksa di setiap
paragraf: font dan ukuran, before/after 0, indent, format caption, tabel tanpa miring,
dan page break hanya sebelum sub-bab level 2.
