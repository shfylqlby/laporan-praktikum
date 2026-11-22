# Cara pakai laprak

Panduan dari nol: siapkan bahan, minta agent, cek hasil, sampai menyesuaikan format.
Instalasi ada di [README](../README.md).

## Gambaran alur

```
modul praktikum + screenshot
        │  kamu kasih ke agent
        ▼
agent tulis spec.json + prosa Indonesia
        │  build_laprak.py + template resmi
        ▼
Laporan.docx
        │  kamu buka di Word, update field, baca ulang
        ▼
laporan siap kumpul
```

Agent hanya menyusun isi (`spec.json`). Format (font, margin, penomoran, gaya tabel) datang dari template di `assets/`, jadi hasilnya konsisten antar modul.

## 1. Siapkan bahan

1. **Modul praktikum**: markdown, PDF, atau docx. Makin jelas langkah dan tujuannya, makin bagus drafnya.
2. **Folder screenshot**, nama berurutan sesuai langkah. Contoh:

   ```
   shots/
   ├── 01-percobaan-1/
   │   ├── 01.png      # topologi
   │   └── 02.png      # konfigurasi
   └── 02-percobaan-2/
       └── 01.png
   ```

   Urutan nama berkas dianggap urutan langkah. Kalau urutanmu beda, bilang ke agent.
3. **Data hasil percobaan**: angka pengukuran, output ping, tabel kebenaran, dsb. Agent tidak bisa menebak data praktikummu. Kalau tidak diberi, bagian analisis akan umum dan perlu kamu isi sendiri.
4. **Profil format**: pilih sesuai mata kuliah.

| Profil | Pakai untuk | Beda utama |
|---|---|---|
| `srwe` | praktikum jaringan komputer (router, switch, wireless) | judul "BAB V" dua baris, gambar maks 11,67 cm |
| `eldas` | praktikum elektronika dasar dan sistem digital | judul tanpa "BAB", gambar maks 14 cm, heading sampai 4 tingkat |

   Ragu? Lihat laporan contoh dari asisten atau dosen. Kalau ada kata "BAB" di judul, itu `srwe`.

## 2. Minta ke agent

Beri tiga hal: modul, folder gambar, profil. Contoh prompt:

```
bikin laprak profil eldas, bab 4.
modul: ./modul4.pdf
screenshot: ./shots (urut sesuai langkah)
data hasil: ./data-hasil.md
output: Laporan-Modul-4.docx
```

Agent akan:

1. Membaca modul dan daftar gambar.
2. Memilih profil. Kalau ragu, ia bertanya. Jawab saja.
3. Memetakan gambar ke langkah percobaan. Kalau tidak jelas, ia konfirmasi.
4. Menulis `spec.json` dan prosa mengikuti [kaidah bahasa](../skills/laprak/references/bahasa-id.md).
5. Menjalankan `build_laprak.py` dan memberi tahu lokasi `.docx`.

Cara panggil per agent: Claude Code `/laprak`, Codex `$laprak`, agent lain cukup minta langsung. Lihat tabel di README.

## 3. Cek hasil di Word

Langkah ini wajib. Nomor gambar dan tabel memakai field `SEQ` Word, jadi **belum muncul** saat file pertama dibuka (terlihat sebagai "Gambar 5.  Topologi").

1. Buka `.docx` di Microsoft Word.
2. `Ctrl+A`, lalu `F9`. Nomor gambar dan tabel terisi.
3. Cek rujukan silang di prosa. "Seperti pada Gambar 5.2" ditulis manual, jadi cocokkan dengan nomor sebenarnya. Kalau urutan gambar berubah, rujukannya ikut digeser.
4. Baca keseluruhan:
   - Gambar ada di langkah yang benar dan keterangannya cocok.
   - Caption tabel di atas tabel, caption gambar di bawah gambar.
   - Data di tabel dan analisis sesuai praktikum aslimu.
   - Tidak ada bagian generik yang tidak mencerminkan apa yang kamu lakukan.
5. Simpan dengan nama baru.

Kalau ada yang salah, jangan edit `.docx` berulang-ulang. Perbaiki `spec.json` (atau minta agent memperbaikinya) lalu build ulang. Hasil build selalu bisa diulang.

## 4. Build manual tanpa agent

```bash
pip install python-docx pillow
python skills/laprak/scripts/build_laprak.py spec.json -o "Laporan-Modul-4.docx"
```

Path gambar di `spec.json` relatif terhadap folder `spec.json` (atau `base_dir` bila diisi). Contoh siap jalan ada di [`docs/contoh/`](contoh/spec.json):

```bash
python skills/laprak/scripts/build_laprak.py docs/contoh/spec.json -o contoh.docx
```

Kesalahan umum yang dilaporkan skrip:

| Pesan | Penyebab |
|---|---|
| `profil tidak dikenal: xyz. Pilihan: srwe, eldas` | nilai `"profil"` salah ketik |
| `gambar tidak ada: <path>` | `file` gambar salah atau `base_dir` keliru |

## 5. Menyesuaikan isi: struktur `spec.json`

Bentuk lengkap ada di [spec-schema.md](../skills/laprak/references/spec-schema.md). Ringkasnya:

```jsonc
{
  "profil": "srwe",           // srwe | eldas
  "bab": 5,                   // menyetir "BAB V", 5.1, Gambar 5.x
  "judul": "Baris 1\nBaris 2",// \n = pindah baris
  "bagian": [ ... ]
}
```

Tiap `bagian` punya heading (`judul`), lalu isinya dirender berurutan: `pengantar`, `isi`, `daftar`, `penutup`, `sub`. `sub` membuat heading satu tingkat lebih dalam, maksimal tiga tingkat (5.1, 5.1.1, 5.1.1.1).

Blok di dalam `isi` atau butir `daftar` bisa memuat: `teks`, `paragraf`, `kode`, `tabel`, `gambar`, `paragraf_akhir` (dirender dalam urutan itu).

Resep umum:

| Mau | Lakukan |
|---|---|
| Tambah percobaan | tambah objek di `sub` pada bagian "Percobaan" dan "Hasil dan Analisis" |
| Buang bagian tidak perlu | hapus objeknya dari `bagian`. Tidak ada bagian yang dipaksakan |
| Modul tanpa kode | jangan pakai kunci `kode` |
| Beberapa gambar berurutan | `"gambar": [{...}, {...}]` |
| Beberapa tabel | `"tabel": [{...}, {...}]` |
| Langkah bernomor dengan gambar | taruh di `daftar`: `{"teks": "...", "gambar": {...}}` |
| Paksa cetak miring | `*teks*` |
| Atur lebar kolom tabel | `"widths": [2.9, 3.0]` (cm) |

## 6. Menyesuaikan format

### Dosen minta variasi kecil: `profil_override`

Timpa sebagian parameter profil langsung di `spec.json`:

```json
"profil_override": {
  "paragraf": {"first_line": 709, "line": 1.5, "after": 0},
  "kode": {"font": "Consolas", "size": 9, "indent": 709}
}
```

**Penting:** penimpaan bersifat dangkal. Yang diganti adalah satu kunci level atas secara utuh. Jadi kalau kamu menulis `"tabel": {"header_fill": "CCCCCC"}`, kunci `tabel` lain (`style`, `width`, `border`) hilang dan build error `KeyError: 'style'`. Selalu tulis sub-objeknya **lengkap**: salin dari `assets/profil.json` lalu ubah nilai yang perlu.

```json
"profil_override": {
  "tabel": {"style": "Table Grid", "width": 7938, "header_fill": "CCCCCC",
            "header_color": "FFFFFF", "border": false}
}
```

Kunci yang sering disesuaikan:

| Kunci | Arti |
|---|---|
| `gambar.max_w`, `gambar.max_h` | batas lebar dan tinggi gambar (cm) |
| `tabel.header_fill` | warna latar header tabel (hex). Bawaan `null` (tanpa warna); isi hanya bila dosen meminta |
| `tabel.header_color` | warna teks header (hex, `null` = default) |
| `paragraf.first_line` | indentasi baris pertama (twips, 709 = 1,25 cm) |
| `paragraf.line` | spasi baris (1.5 = satu setengah) |
| `kode.font`, `kode.size` | font dan ukuran blok kode |
| `judul_bab` | `true` = judul berawalan "BAB" |
| `break_antar_bagian` | `true` = tiap bagian level atas mulai halaman baru |

Daftar lengkap kunci ada di [`assets/profil.json`](../skills/laprak/assets/profil.json). Arti angkanya ada di [struktur-docx.md](../skills/laprak/references/struktur-docx.md).

### Istilah asing yang dimiringkan otomatis

Skrip memiringkan istilah dari daftar bawaan (*router*, *switch*, *gateway*, *breadboard*, dll) di prosa, bukan di heading.

```jsonc
"istilah": ["gateway", "firmware"],   // ganti seluruh daftar bawaan
"istilah_kecuali": ["port", "output"] // buang istilah tertentu dari daftar
```

Pakai `istilah_kecuali` kalau satu kata tidak mau dimiringkan. Pakai `istilah` bila ingin daftar sendiri sepenuhnya. Untuk satu kasus saja, `*teks*` lebih cepat.

### Beda format total: profil baru

Kalau kampus atau mata kuliahmu pakai format yang benar-benar lain:

1. Ambil dokumen resmi (atau contoh laporan terbaik) dan kosongkan isinya, sisakan style, margin, dan penomoran.
2. Hapus nama dan data pribadi dari properti dokumen: Word, **File > Info > Periksa Dokumen**.
3. Simpan sebagai `skills/laprak/assets/<nama>.docx`.
4. Tambah entri `"<nama>"` di `assets/profil.json`. Salin entri `srwe` atau `eldas` sebagai titik awal, lalu cocokkan `heading`, `caption_*_style`, `tabel`, `gambar`, dan indentasi dengan template barumu. Panduan angka ada di `struktur-docx.md`.
5. Uji:

   ```bash
   python skills/laprak/scripts/selfcheck.py
   ```

   `selfcheck.py` menguji semua profil yang terdaftar di `profil.json`, termasuk yang baru. Cek juga hasilnya secara visual di Word.
6. Pakai dengan `"profil": "<nama>"`.

## 7. Menyesuaikan gaya tulisan

Prosa mengikuti [bahasa-id.md](../skills/laprak/references/bahasa-id.md): tanpa em dash, tanpa bullet di badan laporan, tanpa klaim kepastian berlebihan, kalimat aktif yang menerangkan sebab akibat, tanpa kata ganti orang pertama di kesimpulan.

Kalau dosenmu punya aturan lain (misalnya wajib "kami" atau wajib bullet), edit file itu atau katakan ke agent di prompt. Aturan di prompt berlaku untuk sesi itu saja. Aturan di berkas berlaku seterusnya.

## 8. Troubleshooting

| Gejala | Penyebab | Solusi |
|---|---|---|
| "Gambar 5." tanpa angka | field `SEQ` belum diperbarui | Word: `Ctrl+A`, `F9` |
| Nomor BAB atau heading salah | `"bab"` di spec salah | ubah `"bab"` lalu build ulang |
| Gambar tidak masuk | path salah atau berkas tidak ada | cek `file`, gunakan path relatif dari folder `spec.json` |
| `KeyError` setelah memakai `profil_override` | sub-objek ditimpa tidak lengkap | tulis sub-objek utuh, lihat bagian 6 |
| Istilah dimiringkan tidak diinginkan | ada di daftar bawaan | `istilah_kecuali` atau tulis ulang dengan `*...*` |
| Nama font atau tampilan beda di komputer lain | font Times New Roman tidak ada | pasang font atau buka di Word |
| Agent tidak mengenali skill | belum terpasang atau belum restart | cek folder instalasi di README, restart agent |

## 9. Sebelum dikumpulkan

- [ ] `Ctrl+A`, `F9` sudah dijalankan
- [ ] Rujukan "Gambar x.y" dan "Tabel x.y" cocok dengan nomor sebenarnya
- [ ] Data dan angka sesuai praktikum aslimu
- [ ] Analisis dan kesimpulan kamu baca ulang dan pahami
- [ ] Cover, nama, NIM, dan kelompok (kalau diwajibkan) sudah dilengkapi. Skill tidak membuat cover, jadi cek sendiri
- [ ] Format sudah dicek terhadap panduan terbaru dosen atau asisten

Skill menyusun format dan draf. Isi laporan tetap tanggung jawabmu: pahami yang kamu kumpulkan.
