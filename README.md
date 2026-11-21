# laporan-praktikum

**Laprak, versi 1.2.2.** Satu skill `laprak` untuk agent coding yang mengisi atau membangun **laporan praktikum** mata kuliah apa pun sesuai template dosen. Skill ini bisa mengisi template yang sedang terbuka di Word, atau membangun `.docx` dari modul dan folder screenshot (agent menyusun `spec.json`, lalu skrip membangun dokumen dari template resmi). Penomoran karya ilmiah, format umum (Times New Roman 12, spasi 0, indent 36 pt), dan humanizer bahasa Indonesia sudah termasuk, jadi cukup panggil satu skill.

Diturunkan dari [RavLyr/laporan-praktikum](https://github.com/RavLyr/laporan-praktikum). Humanizer di `skills/laprak/humanizer-indonesia/` berasal dari [adhiprib/humanizer-indonesia](https://github.com/adhiprib/humanizer-indonesia) (MIT, lihat `LICENSE` di folder itu).

Format skill-nya mengikuti standar [Agent Skills](https://agentskills.io), jadi satu paket yang sama jalan di **Claude Code, Codex, OpenCode, Antigravity**, dan puluhan agent lain.

| Agent | Kompatibilitas | Pasang | Panggil |
|---|---|---|---|
| Claude Code | sesuai dokumentasi | `~/.claude/skills/` | `/laprak` |
| Codex (CLI, IDE, app) | sesuai dokumentasi | `~/.agents/skills/` | `$laprak` |
| OpenCode | sesuai dokumentasi | `~/.config/opencode/skills/` | `@laprak` atau minta langsung |
| Antigravity (IDE, CLI) | sesuai dokumentasi | `~/.gemini/config/skills/` | minta langsung |
| Cursor, Cline, Copilot, Kiro, dll | lewat `npx skills` | otomatis | sesuai agent |

Folder dan cara panggil diambil dari dokumentasi masing-masing agent. Skill ini baru diuji lewat `npx skills` (terdeteksi) dan di OpenCode. Di agent lain mengikuti dokumentasinya, jadi kabari lewat issue bila ada yang tidak jalan. Detail dan catatan versi ada di [Pemasangan](#pemasangan).

## Mulai cepat

```bash
# 1. pasang (pilih agent saat ditanya, tambah -g untuk semua proyek)
npx skills add shfylqlby/laporan-praktikum -g

# 2. pasang dependensi Python
pip install python-docx pillow

# 3. di agent-mu
#    "bikin laprak profil eldas, bab 4. modul: ./modul4.pdf, screenshot: ./shots"
```

Buka `.docx` hasilnya di Word dan setujui tawaran memperbarui field (atau tekan **Ctrl+A lalu F9**) supaya nomor gambar dan tabel muncul, lalu baca ulang. Di Windows, agent bisa melakukannya otomatis lewat `scripts/word_final.ps1`. Panduan lengkap: [docs/CARA-PAKAI.md](docs/CARA-PAKAI.md).

## Daftar isi

- [Fitur](#fitur)
- [Cara kerja](#cara-kerja)
- [Prasyarat](#prasyarat)
- [Pemasangan](#pemasangan)
- [Pemakaian](#pemakaian)
- [Profil format](#profil-format)
- [Menyesuaikan](#menyesuaikan)
- [Struktur repo](#struktur-repo)
- [Pertanyaan umum](#pertanyaan-umum)
- [Batasan](#batasan)

## Fitur

- Heading bernomor otomatis (BAB V, 5.1, 5.1.1) dari satu nilai `"bab"`.
- Caption gambar dan tabel memakai field `SEQ` Word, bukan ketikan manual.
- Tabel data berwarna, tabel kode (Courier New), gambar rata tengah dengan rasio asli dan caption tepat di bawahnya.
- Istilah asing dimiringkan otomatis di prosa, tidak di heading, caption, tabel, atau kode.
- Format umum semua mata kuliah: Times New Roman 12, before/after 0, indent 36 pt, dan setiap sub-bab level 2 di halaman baru.
- Caption identik dengan hasil Word InsertCaption ("Gambar 6.1", field `SEQ Gambar_6.`).
- Edit dokumen yang sedang terbuka di Word lewat COM dengan backup `.bak` otomatis (`word_live.ps1`).
- Selalu dibangun dari template di `assets/`, jadi style, penomoran, dan font identik dengan dokumen resmi.
- Pedoman menulis prosa Indonesia yang tidak terbaca kaku atau seperti keluaran AI ([bahasa-id.md](skills/laprak/references/bahasa-id.md)).
- Uji mandiri untuk semua profil: `python skills/laprak/scripts/selfcheck.py`.

## Cara kerja

```
modul praktikum + screenshot + data hasil
              │  kamu kasih ke agent
              ▼
   agent menulis spec.json dan prosa
              │  scripts/build_laprak.py + template di assets/
              ▼
        Laporan.docx
              │  kamu buka di Word, Ctrl+A, F9, baca ulang
              ▼
      laporan siap dikumpulkan
```

Agent hanya mengisi **isi** (`spec.json`). **Format** datang dari template, jadi hasilnya konsisten antar modul. Yang berganti per modul hanya nomor BAB, judul, dan isi.

Agent memuat `SKILL.md` saat permintaanmu cocok dengan `description`-nya (misalnya "laprak", "laporan praktikum", "teknik komputer", "jaringan", "elektronika", "sistem digital"), atau saat kamu memanggilnya langsung.

## Prasyarat

| Kebutuhan | Keterangan |
|---|---|
| Agent coding | salah satu dari tabel di atas, dengan akses shell untuk menjalankan skrip |
| Python 3.9+ | `python3` di macOS dan Linux, `python` di Windows |
| `python-docx`, `pillow` | `pip install python-docx pillow` |
| Node.js 22.20+ | hanya bila memakai `npx skills` |
| Microsoft Word | untuk memperbarui field `SEQ` dan memeriksa hasil. LibreOffice belum diuji |

Agent menjalankan skrip di mesin yang sama dengan agent-nya. Kalau Python-mu ada di virtualenv, aktifkan sebelum menjalankan agent.

## Pemasangan

Semua agent memakai folder `laprak` yang sama. Beda hanya lokasi tujuan.

### Cara cepat: `npx skills`

```bash
npx skills add shfylqlby/laporan-praktikum              # proyek ini saja
npx skills add shfylqlby/laporan-praktikum -g           # semua proyek
npx skills add shfylqlby/laporan-praktikum -a claude-code -a codex -g   # agent tertentu
```

Perintah ini ([vercel-labs/skills](https://github.com/vercel-labs/skills)) mendeteksi agent yang terpasang dan menaruh skill di folder yang benar. Alat bantu lain: `npx skills list`, `npx skills update`, `npx skills remove laprak`.

### Manual

```bash
git clone https://github.com/shfylqlby/laporan-praktikum.git
cp -r laporan-praktikum/skills/laprak <FOLDER-TUJUAN>/
```

| Agent | Global (semua proyek) | Per proyek |
|---|---|---|
| Claude Code | `~/.claude/skills/` | `.claude/skills/` |
| Codex | `~/.agents/skills/` | `.agents/skills/` |
| OpenCode | `~/.config/opencode/skills/` | `.opencode/skills/` |
| Antigravity | `~/.gemini/config/skills/` | `.agents/skills/` |

Catatan per agent:

- **Claude Code.** Skill juga bisa dipanggil lewat `/laprak`. Kalau tidak muncul, jalankan `/reload-skills` atau mulai ulang sesi.
- **Codex.** Mendeteksi perubahan folder otomatis, mulai ulang bila belum muncul. Ketik `/skills` untuk melihat daftar, lalu `$laprak` untuk memanggil. Folder per proyek dicari dari direktori kerja sampai akar repo.
- **OpenCode.** Langkah lengkap ada di [Pemasangan di OpenCode](#pemasangan-di-opencode). Selain `~/.config/opencode/skills/`, OpenCode juga membaca `~/.claude/skills/`, `~/.agents/skills/`, `.claude/skills/`, dan `.agents/skills/`. Jadi pemasangan untuk Claude Code atau Codex otomatis terbaca juga di OpenCode. Jangan menaruh skill yang sama di dua tempat, karena ID yang bentrok ditimpa sumber yang didaftarkan belakangan.
- **Antigravity.** Folder global di atas mengikuti `npx skills`. Codelab resmi Google menyebut Antigravity CLI bisa memakai `~/.gemini/antigravity-cli/skills/`. Kalau skill tidak terdeteksi di CLI, salin ke folder itu.
- **Satu pemasangan untuk beberapa agent.** `~/.agents/skills/` dibaca Codex dan OpenCode (serta banyak agent lain di daftar `npx skills`). Kalau memakai keduanya, cukup satu salinan di sana. Claude Code memakai folder sendiri.
- **Agent lain.** Pakai `npx skills add`. Daftar agent dan foldernya ada di [docs agents](https://github.com/vercel-labs/skills/blob/main/docs/agents.md).

### Pemasangan di OpenCode

**Cara cepat, tanpa clone** (butuh Node.js dan Python):

```bash
npx skills add https://github.com/shfylqlby/laporan-praktikum/tree/laprak-1.2.2/skills/laprak -g -a opencode -y
pip install python-docx pillow
```

Skill terpasang di `~/.config/opencode/skills/laprak`, lengkap dengan `scripts/`, `assets/`, dan humanizer. URL berisi nama branch, jadi yang terpasang adalah versi 1.2.2 walaupun `main` masih versi lama. Untuk memperbarui, jalankan perintah yang sama lagi.

**Cara manual:**

1. Pasang Python 3.9+ dan dependensinya:
   ```bash
   pip install python-docx pillow
   ```
2. Ambil repo ini:
   ```bash
   git clone https://github.com/shfylqlby/laporan-praktikum.git
   ```
3. Salin folder skill ke folder skill global OpenCode.
   - macOS / Linux:
     ```bash
     mkdir -p ~/.config/opencode/skills
     cp -r laporan-praktikum/skills/laprak ~/.config/opencode/skills/
     ```
   - Windows (PowerShell):
     ```powershell
     New-Item -ItemType Directory -Force "$env:USERPROFILE\.config\opencode\skills"
     Copy-Item -Recurse laporan-praktikum\skills\laprak "$env:USERPROFILE\.config\opencode\skills\"
     ```
   Untuk satu proyek saja, salin ke `.opencode/skills/laprak` di akar proyek.
4. Uji generator:
   ```bash
   python ~/.config/opencode/skills/laprak/scripts/selfcheck.py
   ```
   Harus mencetak `selfcheck OK`.
5. Buka atau mulai ulang OpenCode, lalu tanya "skill apa saja yang tersedia?". `laprak` harus ada di daftar.
6. Panggil dengan menyebut `@laprak` di prompt, misalnya:
   ```
   @laprak isi BAB 6 profil srwe. modul: ./modul6.pdf, screenshot: ./shots, data: ./data.md
   ```
   OpenCode tidak memakai perintah garis miring untuk skill. Menulis "bikin laprak ..." tanpa `@` juga biasanya memicu skill lewat deskripsinya.

Catatan OpenCode:
- ID skill diambil dari nama folder, jadi foldernya harus tetap bernama `laprak`.
- Jangan pasang `laprak` di dua tempat sekaligus (misalnya `~/.config/opencode/skills/` dan `~/.claude/skills/`), karena OpenCode membaca keduanya dan salah satunya menimpa yang lain.
- Bila `opencode.json` memakai aturan izin `skill`, pastikan `laprak` tidak berstatus `deny`.
- Memperbarui: `git pull` di folder repo, lalu salin ulang folder `skills/laprak`.

### Memeriksa pemasangan

Mulai ulang agent bila perlu, lalu:

| Agent | Cara cek |
|---|---|
| Claude Code, Codex | ketik `/skills` |
| OpenCode, Antigravity | tanya "skill apa saja yang tersedia?" |
| Semua | `python <folder-skill>/laprak/scripts/selfcheck.py` harus mencetak `selfcheck OK` |

## Pemakaian

### 1. Siapkan bahan

- Modul praktikum (markdown, PDF, atau docx).
- Folder screenshot, nama berkas berurutan sesuai langkah.
- Data hasil percobaan (angka, output, tabel). Agent tidak bisa menebak data praktikummu.
- Pilihan [profil](#profil-format): `srwe` atau `eldas`.

### 2. Minta ke agent

```
bikin laprak profil eldas, bab 4.
modul: ./modul4.pdf
screenshot: ./shots (urut sesuai langkah)
data hasil: ./data-hasil.md
output: Laporan-Modul-4.docx
```

Panggil lewat `/laprak` (Claude Code), `$laprak` (Codex), `@laprak` (OpenCode), atau minta langsung di agent mana pun. Kalau agent tidak menemukan skrip, sebutkan lokasi skill-nya. Path `scripts/`, `assets/`, dan `references/` relatif terhadap folder `laprak`, bukan folder kerjamu.

### 3. Periksa hasil

1. Buka `.docx` di Word, **Ctrl+A lalu F9**.
2. Cocokkan rujukan silang ("seperti pada Gambar 5.2") dengan nomor sebenarnya.
3. Periksa data, gambar, dan caption, lalu baca analisisnya.
4. Perlu perbaikan? Ubah `spec.json` atau minta agent memperbaikinya, lalu build ulang. Jangan menambal `.docx` berulang kali.

### Tanpa agent

```bash
python skills/laprak/scripts/build_laprak.py spec.json -o "Laporan-Modul-4.docx"
```

Contoh siap jalan: [`docs/contoh/spec.json`](docs/contoh/spec.json).

```bash
python skills/laprak/scripts/build_laprak.py docs/contoh/spec.json -o contoh.docx
```

Contoh `spec.json` minimal:

```json
{
  "profil": "eldas",
  "bab": 4,
  "judul": "Gerbang Logika Dasar",
  "bagian": [
    {"judul": "Tujuan", "daftar": ["Praktikan mampu menjelaskan cara kerja gerbang AND."]},
    {"judul": "Hasil Percobaan dan Analisis", "sub": [
      {"judul": "Percobaan 1", "isi": [
        {"paragraf": ["Rangkaian diuji dengan empat kombinasi input."],
         "tabel": {"caption": "Tabel Kebenaran", "header": ["A", "B", "Y"], "rows": [["0", "0", "0"], ["1", "1", "1"]]},
         "gambar": [{"file": "shots/01.png", "caption": "Rangkaian AND"}]}
      ]}
    ]},
    {"judul": "Kesimpulan", "daftar": ["Keluaran AND bernilai 1 hanya jika kedua input 1."]}
  ]
}
```

Skema lengkap: [spec-schema.md](skills/laprak/references/spec-schema.md).

## Profil format

| Profil | Untuk | Ciri |
|---|---|---|
| `srwe` | praktikum jaringan komputer (router, switch, wireless) | judul berawalan "BAB V" dua baris, caption gambar `captiongambar`, header tabel `FF89FF`, gambar maks 11,67 cm |
| `eldas` | praktikum elektronika dasar dan sistem digital (rangkaian, gerbang logika) | judul tanpa "BAB", seluruh caption style `Caption`, header tabel `E8EEF5`, gambar maks 14 cm, heading sampai 4 tingkat |

Ragu memilih? Lihat contoh laporan dari asisten atau dosen. Ada kata "BAB" di judul dan tabel pink berarti `srwe`.

## Menyesuaikan

| Kebutuhan | Cara | Rujukan |
|---|---|---|
| Dosen minta variasi kecil (font kode, lebar gambar, warna header) | `profil_override` di `spec.json` | [CARA-PAKAI, bagian 6](docs/CARA-PAKAI.md) |
| Istilah asing salah dimiringkan | `istilah` atau `istilah_kecuali` | [CARA-PAKAI, bagian 6](docs/CARA-PAKAI.md) |
| Format kampus lain | tambah profil baru | [Menambah profil](#menambah-profil-baru) |
| Gaya tulisan beda | edit atau tambahkan aturan di prompt | [bahasa-id.md](skills/laprak/references/bahasa-id.md) |

`profil_override` menimpa dangkal: menulis `"tabel": {"header_fill": "CCCCCC"}` saja menghapus kunci `tabel` lain dan build gagal dengan `KeyError: 'style'`. Selalu tulis sub-objeknya lengkap. Contoh lengkap ada di panduan.

### Menambah profil baru

1. Kosongkan dokumen resmi (sisakan style, margin, penomoran) dan simpan sebagai `skills/laprak/assets/<nama>.docx`.
2. Hapus nama dan data pribadi dari properti dokumen (Word: File > Info > Periksa Dokumen).
3. Tambah entri `<nama>` di `skills/laprak/assets/profil.json`. Salin `srwe` atau `eldas` sebagai titik awal.
4. Jalankan `python skills/laprak/scripts/selfcheck.py`, lalu periksa visual di Word.

## Struktur repo

```
README.md
docs/
├── CARA-PAKAI.md             # panduan lengkap: bahan, prompt, cek, penyesuaian
└── contoh/                   # spec.json + gambar contoh siap build
.claude-plugin/plugin.json    # manifest plugin Claude Code (versi 1.2.2)
skills/laprak/
├── SKILL.md                  # satu-satunya skill: jalur, format umum, aturan isi, humanizer, penomoran
├── humanizer-indonesia/      # Humanizer Indonesia 1.7.2 (adhiprib, MIT)
│   ├── PANDUAN.md            # alur kerja humanizer
│   ├── references/           # katalog pola tulisan AI, preset suara
│   └── scripts/metrik.py     # pengukur gaya teks Indonesia
├── assets/
│   ├── srwe.docx             # template kosong profil srwe (jaringan)
│   ├── eldas.docx            # template kosong profil eldas (elektronika, digital)
│   └── profil.json           # parameter format tiap profil
├── scripts/
│   ├── build_laprak.py       # generator spec.json -> .docx
│   ├── lint_laprak.py        # cek penanda, rujukan Gambar/Tabel, dan gaya
│   ├── word_final.ps1        # perbarui field + ekspor PDF lewat Word (Windows)
│   ├── word_live.ps1         # edit dokumen yang terbuka di Word: .bak, InsertCaption (Windows)
│   └── selfcheck.py          # uji mandiri semua profil
└── references/
    ├── generator.md          # panduan generator Jalur B (dulu SKILL.md generator)
    ├── spec-schema.md        # bentuk spec.json
    ├── struktur-docx.md      # angka format hasil pembedahan template
    └── bahasa-id.md          # kaidah menulis prosa Indonesia
```

### Satu skill, dua jalur

Sejak 1.2.2 hanya ada satu skill, `laprak`. Skill `laprak-word` dan `penomoran-akademik` yang dulu terpisah sudah digabung ke dalamnya.

- **Jalur A, edit langsung.** Template dosen terbuka di Word (lewat add-in seperti Claude for Word, atau lewat COM dengan `scripts/word_live.ps1` di Windows) atau dilampirkan sebagai `.docx`. Skill mengisi satu BAB, menjaga style template, dan memasang caption Word asli.
- **Jalur B, bangun dari spec.** Tidak ada dokumen; agent menulis `spec.json` dan `scripts/build_laprak.py` membangun `.docx` dari template di `assets/`.

Di kedua jalur, analisis ditulis ringkas (langkah 4b) lalu disunting humanizer (langkah 5b) sebelum masuk dokumen, dan penomoran akademik (Bagian 9) diterapkan otomatis.

Aturan yang dipegang skill, dan sebaiknya tidak dilanggar:

- Jangan menulis `.docx` dari nol. Selalu lewat template di `assets/`.
- Jangan menimpa berkas di `assets/`.
- Nomor gambar dan tabel selalu field `SEQ`.
- Caption tabel di atas tabel, caption gambar di bawah gambar.

## Pertanyaan umum

**Agent tidak mengenali skill.** Cek folder pemasangan untuk agent-mu, mulai ulang, lalu `/skills` atau tanya daftar skill. Di OpenCode, pastikan tidak ada ID `laprak` ganda di folder yang berbeda.

**Nomor gambar tampil "Gambar 5." tanpa angka.** Field `SEQ` belum diperbarui. Word: Ctrl+A, F9.

**`ModuleNotFoundError: No module named 'docx'`.** Jalankan `pip install python-docx pillow` di Python yang dipakai agent.

**Agent menulis path skrip yang salah.** Sebutkan lokasi folder `laprak`, atau jalankan build manual seperti di atas.

**Bisa dipakai tanpa agent?** Bisa. `build_laprak.py` berdiri sendiri, cukup `spec.json` dan gambar.

**Verifikasi visual otomatis?** Di Windows dengan Word, `scripts/word_final.ps1` memperbarui field, menyimpan, dan mengekspor PDF. Tanpa Word, pakai `soffice --headless --convert-to pdf` lalu periksa hasilnya; nomor caption baru final setelah dibuka di Word.

**Cek ulang laporan yang sudah jadi?** `python skills/laprak/scripts/lint_laprak.py Laporan.docx` melaporkan penanda `[ISI: ...]` yang tersisa, rujukan Gambar atau Tabel yang salah, dan pola gaya.

**Kenapa skill tidak membuatkan cover?** Template hanya mencakup badan laporan. Lengkapi cover, nama, NIM, dan kelompok sendiri.

## Batasan

- Format kampus bisa berubah per semester. Cocokkan dengan panduan terbaru dosen atau asisten.
- Hasil tetap perlu dibaca ulang. Skill menyusun format dan draf prosa, bukan menggantikan data percobaanmu. Pahami yang kamu kumpulkan.
- Dua profil bawaan mengikuti template yang ada. Format lain butuh [profil baru](#menambah-profil-baru).
- Belum diuji langsung di Claude Code, Codex, dan Antigravity. Kompatibilitasnya mengikuti standar Agent Skills dan dokumentasi tiap agent.
- Folder skill tiap agent bisa berubah antar versi. Bila berbeda dari tabel, ikuti dokumentasi resmi agent-nya atau pakai `npx skills add`.
