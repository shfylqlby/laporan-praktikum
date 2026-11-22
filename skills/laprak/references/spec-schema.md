# Skema spec.json

Dokumen dibangun dari daftar **bagian** bersarang. Tidak ada bagian yang dipaksakan,
jadi satu skema melayani praktikum jaringan, elektronika dasar, sistem digital, dan modul lain
dengan struktur serupa.

```jsonc
{
  "profil": "srwe",               // "srwe" atau "eldas", lihat assets/profil.json
  "bab": 5,                       // menyetir "BAB V", "5.1", "Gambar 5.x"
  "judul": "Judul Modul",         // "\n" memecah jadi beberapa baris
  "output": "Laporan.docx",       // opsional, bisa ditimpa lewat -o
  "base_dir": ".",                // akar path gambar relatif; default folder spec
  "istilah": ["gateway"],         // opsional, menimpa daftar bawaan auto-italic
  "profil_override": {},          // opsional, menimpa sebagian kunci profil

  "bagian": [ /* lihat di bawah */ ]
}
```

## Bagian

Tiap bagian menghasilkan satu heading, lalu isinya dalam urutan tetap:
`pengantar` → `isi` → `daftar` → `pustaka` → `penutup` → `sub`.

```jsonc
{
  "judul": "Percobaan",           // opsional; tanpa judul berarti tanpa heading
  "pengantar": ["paragraf pembuka"],
  "isi": [ /* blok */ ],
  "daftar": [ /* string atau blok, dinomori 1, 2, 3 */ ],
  "pustaka": ["Penulis, A. (2020). *Judul Buku*. Penerbit."],  // APA, diurutkan abjad, indent menggantung 1,25 cm
  "penutup": [ /* blok */ ],
  "sub": [ /* bagian bersarang, otomatis turun satu tingkat heading */ ]
}
```

Bagian tingkat teratas dipisah halaman baru secara otomatis.
Kedalaman heading maksimal tiga tingkat: `5.1` → `5.1.1` → `5.1.1.1`.

## Blok

Satu blok boleh menggabungkan beberapa kunci. Urutan render selalu:
`teks` → `paragraf` → `persamaan` → `kode` → `tabel` → `gambar` → `paragraf_akhir`.

```jsonc
{
  "teks": "Satu kalimat langkah.",          // jadi butir bernomor di dalam "daftar"
  "paragraf": ["paragraf biasa"],
  "persamaan": {                            // baris sendiri, rata tengah, nomor (bab.n) di kanan
    "rumus": "Galat = |Vu - Vt| / Vt x 100%",
    "keterangan": ["dengan Vu nilai ukur dan Vt nilai teori."]   // opsional, tanpa indent
  },
  "kode": ["Router>enable", "Router#configure terminal"],
  "tabel": {
    "caption": "Data Hasil Pengukuran",
    "header": ["Sub-Percobaan", "Hasil Ukur"],
    "rows": [["Percobaan 1.1", "0,7 V"]],
    "widths": [2.9, 3.0]                    // cm, opsional
  },
  "gambar": [{"file": "02-percobaan-1/01.png", "caption": "Topologi"}],
  "paragraf_akhir": ["paragraf setelah gambar"]
}
```

`gambar` boleh berupa string, satu objek, atau daftar objek. `tabel` dan `persamaan` boleh
satu objek atau daftar objek. `persamaan` juga boleh berupa string rumus saja.

Gambar yang berkasnya belum ada tidak menghentikan build: skrip memasang kotak
`[ISI: gambar nama.png]` dengan caption tetap ada, lalu mencetak `PERINGATAN`.
Pakai `--ketat` bila ingin build berhenti.

Nomor persamaan dihitung skrip (bab.1, bab.2, ...) dan tidak butuh pembaruan field.
Rujuk di prosa sebagai "Persamaan (5.1)".

## Contoh ringkas: jaringan (profil `srwe`)

```jsonc
"bagian": [
  {"judul": "Tujuan", "daftar": ["..."]},
  {"judul": "Percobaan", "sub": [
    {"judul": "Percobaan 1",
     "isi": [{"tabel": {...}}],
     "daftar": [{"teks": "...", "gambar": {...}}, {"teks": "...", "kode": [...]}]}
  ]},
  {"judul": "Hasil dan Analisis", "sub": [
    {"judul": "Percobaan 1", "isi": [{"paragraf": ["..."], "gambar": [...]}]}
  ]},
  {"judul": "Kesimpulan", "daftar": ["..."]}
]
```

## Contoh ringkas: elektronika dasar (profil `eldas`)

```jsonc
"bagian": [
  {"judul": "Tujuan", "daftar": ["..."]},
  {"judul": "Alat dan Bahan", "sub": [
    {"judul": "Trainer Board",
     "isi": [{"gambar": {...}, "paragraf_akhir": ["penjelasan", "(sumber URL)"]}]}
  ]},
  {"judul": "Langkah Kerja", "sub": [
    {"judul": "Modul Software", "sub": [
      {"judul": "Percobaan 1", "daftar": [{"teks": "...", "gambar": {...}}]}
    ]}
  ]},
  {"judul": "Hasil Percobaan dan Analisis", "sub": [
    {"judul": "Modul Software", "sub": [
      {"judul": "Percobaan 1", "isi": [{"gambar": {...}}, {"tabel": {...}}, {"paragraf": ["..."]}]}
    ]}
  ]},
  {"judul": "Kesimpulan", "daftar": ["..."]}
]
```

## Catatan

- `*miring*` memaksa cetak miring. Sisanya dimiringkan otomatis bila cocok daftar `istilah`.
  Ini hanya berlaku di prosa (paragraf, daftar, pustaka). Heading, caption, sel tabel, dan kode
  tidak pernah miring; tanda `*` di sana dibuang. Nama merek tidak perlu ditandai, karena tidak
  ada di daftar `istilah`; bila ada yang terlanjur cocok, keluarkan lewat `"istilah_kecuali"`.
- Tiap `daftar` memperoleh penomoran sendiri yang mulai dari 1.
- Rujukan silang ditulis manual, misal "seperti pada Gambar 5.4". Setelah build,
  `lint_laprak.py` berjalan otomatis dan mencetak `PERINGATAN` untuk rujukan yang nomornya
  tidak ada, gambar atau tabel yang tidak dirujuk, penanda `[ISI: ...]` atau
  `[CEK SUMBER: ...]` yang tersisa, dan pola gaya (em dash, "di mana", "dilakukan", "kami",
  "membuktikan"). Temuan ditimbang satu per satu, bukan diganti otomatis.
- Modul tanpa kode tinggal tidak memakai kunci `kode`. Tidak ada yang perlu dimatikan.
