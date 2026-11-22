---
name: "laprak"
description: "Mengisi atau membangun laporan praktikum mata kuliah apa pun (elektronika, jaringan, pemrograman, fisika, kimia, dan lainnya) mengikuti template dosen, langsung di dokumen Word yang terbuka atau dibangun sebagai .docx dari modul dan data percobaan, dengan aturan penomoran karya ilmiah, humanizer bahasa Indonesia, dan pengecekan otomatis. Trigger: laprak, laporan praktikum, isi laporan, bikin laporan modul, hasil dan analisis, kesimpulan praktikum, penomoran, nomor halaman, kerangka bab, rapikan bahasa laporan, biar tidak kaku, penomoran makalah, proposal, skripsi, caption gambar dan tabel."
---

# Laporan Praktikum (Word) dengan penomoran akademik

Skill ini mengisi template laporan praktikum yang sedang terbuka dan memastikan penomorannya benar, dalam satu kali pakai. Hasilnya harus langsung bisa dikumpulkan: format sesuai template, isi berasal dari data pengguna, dan tidak terbaca seperti jawaban chat.

## Kapan dipakai
- Pengguna meminta mengerjakan, melengkapi, atau merapikan laporan praktikum atau salah satu bagiannya.
- Dokumen terbuka berupa template dosen (judul BAB, heading dengan petunjuk dalam kurung, kolom kosong).
- Pengguna meminta membuat kerangka bab, merapikan, atau mengecek penomoran laporan, makalah, proposal, atau skripsi: nomor halaman (Romawi kecil dan Arab), hierarki bab dan sub-bab, serta penomoran tabel, gambar, dan caption per bab (Bagian 9).
- Pengguna cukup memanggil `/laprak`. Isi laporan, Format umum, humanizer, dan penomoran akademik semuanya berjalan dari skill ini; tidak ada skill pendamping yang perlu dipanggil.

## Dua mode
- **Mode isi BAB** (default untuk laprak): template sudah ada, tugasnya mengisi satu BAB. Bagian 1 sampai 8 berlaku. Dari aturan penomoran, yang berlaku hanya Bagian 9 (caption dan hierarki). Jangan menyentuh section break, nomor halaman, cover, daftar isi, atau daftar gambar kecuali diminta.
- **Mode dokumen lengkap** (kerangka, perapian, atau pengecekan laporan, makalah, proposal, skripsi): seluruh Bagian 9 berlaku, termasuk nomor halaman.

## Lingkungan kerja
- Bila tool Word add-in tersedia (`execute_office_js`, `verify_doc`, `verify_doc_visual`), pakai itu langsung di dokumen yang terbuka.
- Bila tool add-in tidak ada tetapi dokumen terbuka di Microsoft Word di Windows dan agent punya shell (misalnya Claude Code), edit live lewat COM Word (`GetActiveObject('Word.Application')`) dan timpa berkas yang terbuka. Buat backup `.bak` dulu sebelum mengubah apa pun, dan jangan menimpa `.bak` lama. Pakai `scripts/word_live.ps1` dari folder skill ini (`-Aksi info`, `caption`, `update`); skrip itu sudah menangani backup, InsertCaption, dan spasi sebelum nomor. Contoh: `powershell -NoProfile -ExecutionPolicy Bypass -File scripts/word_live.ps1 -Aksi caption -Jenis Gambar -Indeks 3 -Bab 6 -Teks "Topologi"`. Rincian aksi ada di `references/generator.md`, bagian "Mengedit dokumen yang sudah terbuka di Word".
- Bila tidak ada keduanya, dan pengguna melampirkan berkas .docx, kerjakan salinannya dengan skill `docx`, kembalikan sebagai .docx, dan sebutkan satu kalimat bahwa verifikasi visual dilakukan lewat render, bukan Word langsung.

## Dua jalur kerja
Pilih jalur di awal, berdasarkan apa yang ada, lalu sebutkan satu kalimat jalur mana yang dipakai.

| Keadaan | Jalur |
|---|---|
| Template dosen sedang terbuka di Word, atau pengguna melampirkan .docx yang harus diisi atau diperbaiki | **A. Edit langsung** (langkah 1 sampai 8 di bawah) |
| Tidak ada dokumen terbuka; pengguna memberi modul plus folder screenshot dan minta laporan jadi | **B. Bangun dari spec** |
| Dokumen terbuka tetapi pengguna ingin hasil dibangun ulang bersih dari template | tanya satu kali, default A |

Jalur A dan B memakai aturan yang sama: Prinsip utama, gaya bahasa, analisis, sitasi, penomoran (Bagian 9), dan pengecekan di langkah 8. Yang berbeda hanya cara menulis ke dokumen.

### Jalur B: bangun dari spec
Memakai generator yang ada di folder skill ini: `scripts/`, `assets/`, dan `references/`. Semua path di bawah relatif terhadap folder skill ini, bukan folder kerja pengguna. Format dikunci oleh kode dan template `.docx`, sehingga konsisten antar modul. Panduan lengkap generator ada di `references/generator.md`.
1. Pastikan dependensi ada: `pip install python-docx pillow` (tambah `--break-system-packages` di sandbox Linux bila pip menolak). Uji sekali dengan `python scripts/selfcheck.py`; hasilnya harus `selfcheck OK`.
   Bila folder `scripts/` tidak ada (skill dipasang hanya dengan SKILL.md, misalnya diunggah ke claude.ai tanpa folder lain), ambil generatornya ke folder kerja sendiri: `git clone -b laprak-1.2.2 --depth 1 https://github.com/shfylqlby/laporan-praktikum`, lalu pakai `skills/laprak/` di dalamnya.
2. Baca `references/spec-schema.md` sebelum menulis spec. `references/bahasa-id.md` boleh dibaca sebagai pelengkap; bila berbeda, Bagian 5 dan 5b skill ini yang dipakai.
3. Tentukan profil (`eldas` atau `srwe`) dari ciri di `assets/profil.json`. Bila ragu, tanya satu kali. Bila mata kuliah lain dan belum ada profilnya, pakai Jalur A, atau tambahkan profil baru mengikuti `references/generator.md`.
4. Petakan screenshot ke langkah percobaan (urutan berkas biasanya urutan langkah; konfirmasi bila tidak jelas), tulis `spec.json`, lalu bangun dengan `python scripts/build_laprak.py spec.json -o Laporan-Modul-N.docx`.
5. Prosa di spec tetap tunduk pada Prinsip utama, langkah 4b, dan humanizer 5b di sini: data kosong ditulis `[ISI: ...]`, bukan dikarang.
6. Skrip menjalankan `scripts/lint_laprak.py` otomatis dan mencetak `PERINGATAN` (penanda tersisa, rujukan Gambar atau Tabel tidak cocok, gambar tidak dirujuk, pola gaya, aturan isi). Tangani satu per satu: betulkan spec lalu bangun ulang.
7. Render untuk dicek. Di Windows dengan Word, pakai `scripts/word_final.ps1 -Docx ... -Pdf ...` (memperbarui field sekaligus, jadi nomor caption sudah final). Tanpa Word: `soffice --headless --convert-to pdf` lalu `pdftoppm -png -r 70`; nomor caption akhir baru pasti setelah pengguna memperbarui field di Word. Kirim .docx ke pengguna dengan SendUserFile bila tool itu ada, atau sebutkan path lengkapnya.
8. Bila generator tidak bisa dipakai (Python tidak ada, atau clone ditolak), jangan mengarang skrip: pakai Jalur A dengan skill `docx`, dan beri tahu pengguna satu kalimat.

## Prinsip utama
1. **Data pengguna adalah acuan.** Angka, rumus, screenshot, dan foto dari pengguna dipakai apa adanya. Jangan mengarang data hasil percobaan, nilai komponen, kutipan, nomor halaman, atau referensi. Data yang belum ada ditandai `[ISI: ...]` dan masuk ke daftar cek di akhir.
2. **Template adalah hukum.** Ikuti struktur, urutan, dan style yang sudah ada. Jangan mengubah margin, font, atau style kecuali pengguna meminta. Urutan prioritas bila ada pertentangan: ketentuan dosen atau kampus di dokumen, lalu template, lalu Bagian 9 skill ini, lalu Format umum di bawah. Format umum adalah default untuk semua mata kuliah bila template atau dosen tidak menetapkan lain.
3. **Tulisan, bukan chat.** Isi dokumen tidak boleh memuat komentar kepada pengguna, basa-basi, atau kalimat seperti "berikut adalah". Catatan untuk pengguna hanya ada di balasan chat.
4. **Salah di sumber, bilang.** Kalau rumus atau data pengguna tampak keliru, beri tahu singkat di chat dan jangan mengubahnya diam-diam.
5. **Humanizer selalu aktif.** Setiap kali skill ini dipakai, semua prosa yang ditulis (terutama analisis) melewati langkah 4b dan 5b sebelum masuk dokumen, tanpa perlu diminta pengguna. Tidak ada prosa yang ditulis langsung dari draf pertama. Bila waktu atau alat terbatas, yang boleh dilewati hanya pengukuran `metrik.py`, bukan penyuntingannya.

## Format umum
Berlaku untuk semua mata kuliah bila template atau dosen tidak menetapkan lain.
- **Halaman**: A4, margin kiri 4, atas 4, kanan 3, bawah 3 cm.
- **Teks**: Times New Roman 12 pt di paragraf, heading, nomor heading, daftar, dan isi tabel. Normal/Body Text rata kiri-kanan, spasi 1,5.
- **Spasi paragraf**: before dan after **0** di semua paragraf, termasuk heading, caption, dan sel tabel.
- **Page break**: setiap sub-bab level 2 dimulai di halaman baru. Contoh: setelah 12.1 selesai, page break, lalu 12.2; setelah 12.2 selesai, page break, lalu 12.3. Sub-bab level 3 (12.2.1) tidak dipisah halaman.
- **Paragraf biasa** (bukan daftar): FirstLineIndent 36 pt, LeftIndent 0, sehingga baris pertama sejajar dengan teks heading sub-bab.
- **Heading**: Heading 1 rata tengah dan tebal. Heading 2 dan 3 tebal dengan nomor otomatis, li 36 pt fi −36 pt. Nomor bab dan sub-bab tidak diubah.
- **Daftar** (Tujuan, langkah kerja, Kesimpulan): LeftIndent 36 pt, FirstLineIndent −36 pt; nomor mepet kiri, teks mulai di 36 pt.
- **Caption gambar**: di bawah gambar (Below selected item), TNR 10 hitam, tidak miring, rata tengah, spasi 1,5, before/after 0. Paragraf gambar atau placeholder di atasnya spasi 1, before/after 0.
- **Caption tabel**: di atas tabel (Above selected item), TNR 10 hitam, tidak miring, spasi 1, before/after 0, li 36 pt fi −36 pt.
- **Tabel data**: `Table Grid`, TNR 12, **tanpa warna**: tidak ada shading atau warna latar di header maupun isi, semua teks hitam. Header cukup tebal dan rata tengah; isi tidak tebal, rata tengah, spasi 1,5, tanpa indent. Tabel berwarna (header pink, biru, dan sejenisnya) terbaca sebagai AI slop. Isi tabel data (alamat IP, VLAN) tidak diubah saat merapikan tabel lain.
- **Kode CLI**: tabel 1x1 `Table Grid`, Courier New 10, spasi 1, before/after 0, tidak miring.
- **Miring**: kata atau kalimat bahasa Inggris di paragraf isi (*port security*, *native VLAN*, *sticky*, *DHCP snooping*). Tidak miring: nama merek, produk, atau perangkat (Cisco, Packet Tracer, Command Prompt, Rogue Laptop), heading, caption, tabel, dan kode CLI.
- **Placeholder**: gambar atau data yang belum ada ditandai `[ISI: ...]` dengan teks abu-abu. Struktur dan isi template tidak dihapus tanpa diminta.

## Aturan isi
- Jangan menambah bagian Dasar Teori di dokumen apa pun kecuali diminta eksplisit.
- Langkah kerja tiap percobaan atau konfigurasi (Percobaan 1, 2, dan seterusnya) wajib dipisah. Jangan digabung atau diringkas menjadi satu daftar.
- Tujuan: tiap butir cukup dua kalimat.
- Kesimpulan murni menjawab Tujuan, tanpa penanda `[ISI: ...]` dan tanpa data kelompok (alamat IP, nilai ukur, nama kelompok).
- Bagian yang datanya belum ada (analisis, screenshot) dibiarkan dengan penanda saja; Kesimpulan tetap ditulis dari Tujuan.

## Langkah kerja

### 1. Baca dokumen dan inventarisasi
- Baca outline dan seluruh teks dokumen. Catat: nomor BAB, judul modul, heading beserta petunjuk dalam kurung (misalnya "Minimal 5"), butir bernomor kosong, tabel, dan gambar yang sudah ada.
- Baca ketentuan format di dokumen (ukuran halaman, font, spasi, margin, aturan caption, kotak kode). Periksa halaman lewat `section.pageSetup` bila tersedia. Laporkan bila tidak cocok, jangan diubah sendiri.
- Daftar apa yang sudah diberikan pengguna: modul, data, screenshot, foto. Daftar apa yang belum.
- Nilai yang dibaca dari foto atau screenshot (layar osiloskop, multimeter, simulasi) disalin dulu ke tabel kerja singkat di chat. Nilai yang kabur atau ambigu ditanyakan, tidak ditebak. Angka yang sudah dikonfirmasi dipakai apa adanya.

### 2. Tanya dulu bila kurang jelas
Maksimal dua pertanyaan penting, dalam satu pesan, hanya untuk hal yang benar-benar menghalangi penulisan (nilai ukur, foto mana untuk percobaan mana, tugas pendahuluan bila template memintanya). Contoh: "Data VCE tiap percobaan sudah lengkap? Foto mana yang untuk Percobaan 1?" Jangan bertanya hal yang bisa dibaca dari dokumen. Jangan mulai menulis analisis tanpa data hasil. Bila data sudah cukup untuk sebagian bagian, kerjakan bagian itu dulu dan tandai sisanya `[ISI: ...]` daripada menunggu jawaban.

### 3. Tulis per bagian, satu bagian per panggilan `execute_office_js`
Sebelum menulis, tulis rencana bagian di chat (satu baris tiap bagian). Alur tiap bagian prosa: draf, lalu 4b (khusus analisis), lalu 5b humanizer, lalu tulis ke dokumen. Setiap panggilan menulis dengan cara idempoten: baca dulu heading bagian target sampai heading berikutnya. Bila bagian sudah berisi, ganti isinya, jangan menambahkan di bawahnya, supaya tidak ada bagian ganda.

Urutan umum (Eldas): Tujuan, Alat dan Bahan, Langkah Kerja, Hasil Percobaan dan Analisis, Kesimpulan, lalu Daftar Pustaka bila ada sitasi.

| Bagian | Isi | Bentuk |
|---|---|---|
| Tujuan | Sesuai jumlah minimal di petunjuk. Tiap butir paling banyak dua kalimat, diawali kata kerja (mengamati, menentukan, menganalisis). | Daftar bernomor |
| Alat dan Bahan | Satu sub-heading per alat. Gambar, caption, satu paragraf fungsi singkat. Sumber gambar internet ditulis di bawahnya. | Sub-heading, tanpa bullet |
| Langkah Kerja | Dibagi Modul Software dan Modul Hardware bila template memakainya. Langkah ringkas, urut, satu tindakan per langkah, dipisah per percobaan. Gambar secukupnya. | Daftar bernomor, mulai dari 1 tiap percobaan |
| Hasil dan Analisis | Tiap percobaan: gambar atau tabel hasil, lalu paragraf analisis. | Heading tingkat empat per percobaan, tanpa bullet |
| Kesimpulan | Sesuai jumlah minimal. Menjawab tiap Tujuan, tanpa penanda dan tanpa data kelompok. | Daftar bernomor |
| Daftar Pustaka | Hanya sumber yang benar-benar dikutip. APA. | Urut abjad, indent menggantung |

Petunjuk dalam kurung pada heading (misalnya "(Minimal 5)") dihapus setelah bagiannya terisi. Sisanya, teks heading tidak diubah.

### 4. Menulis analisis (bagian yang paling sering gagal)
Susun tiap paragraf analisis dengan urutan: apa yang teramati (dengan angka), pola atau perbandingan, penyebab menurut prinsip rangkaian, lalu kaitan ke tujuan. Contoh arah untuk transistor: bagaimana VCE berubah terhadap VCC pada IB tetap, apakah transistor berada di daerah saturasi atau aktif, dan apa artinya bagi LED menyala atau mati.

- Hitung dulu dengan kode bila ada perhitungan (selisih, rata-rata, persentase galat, gain). Tulis hasil yang sudah dicek, jangan menghitung di kepala. Galat relatif memakai |nilai ukur - nilai teori| / nilai teori x 100%, ditulis sebagai persamaan bernomor bila rumusnya baru muncul. Desimal mengikuti resolusi alat ukur, konsisten di seluruh tabel.
- Setiap paragraf analisis merujuk satu gambar atau tabel yang menjadi dasarnya, dan setiap angka di paragraf harus bisa ditemukan di data pengguna atau hasil hitung tadi.
- Hubungkan penyebab hanya bila bisa dijelaskan dari data atau teori yang disitasi. Bila hasil menyimpang dari teori, tulis penyimpangannya dan sebab yang masuk akal (toleransi komponen, resistansi kontak, kesalahan baca alat), diberi kata hati-hati seperti "kemungkinan".
- Hindari klaim "membuktikan". Laporan praktikum melaporkan pengamatan.

### 4b. Analisis ringkas, tidak bertele-tele (wajib)
Analisis yang bertele-tele dan terbaca AI hampir selalu punya ciri yang sama: kalimat pembuka yang mengulang konteks, data tabel yang disalin ulang ke prosa, penjelasan teori umum yang tidak dipakai, dan kalimat penutup yang merangkum ulang. Buang keempatnya.

Ukuran:
- Satu paragraf analisis untuk satu gambar atau tabel, tiga sampai lima kalimat. Lebih dari enam kalimat berarti ada yang diulang atau ada teori yang tidak perlu.
- Setiap kalimat membawa paling tidak satu dari ini: angka hasil, perbandingan, penyebab, atau kaitan ke tujuan. Kalimat yang tidak membawa satu pun dihapus.
- Kalimat pertama langsung menyebut temuan dengan angka dan rujukannya. Contoh: "Tegangan VCE turun dari 4,8 V menjadi 0,2 V saat IB dinaikkan ke 50 µA (Tabel 6.1)."
- Dari tabel, kutip hanya angka yang dibahas (nilai ekstrem, titik perubahan, selisih). Jangan menyalin seluruh isi tabel ke kalimat.
- Teori hanya sejauh menjelaskan angka yang teramati, satu sampai dua kalimat. Definisi umum (pengertian transistor, fungsi router) tidak ditulis di analisis.

Pembuka dan penutup yang dilarang di analisis (diganti langsung dengan temuan):
- "Berdasarkan hasil percobaan yang telah dilakukan, dapat dilihat bahwa ..."
- "Pada percobaan ini dilakukan pengamatan terhadap ..."
- "Hal ini menunjukkan bahwa ..." berulang di setiap paragraf
- "Dengan demikian, dapat disimpulkan bahwa ..." di akhir paragraf analisis (kesimpulan tempatnya di bagian Kesimpulan)
- "Secara keseluruhan", "Perlu diketahui bahwa", "Penting untuk dicatat", "Dalam hal ini"

Contoh. Sebelum (bertele-tele, terbaca AI):
> Berdasarkan hasil percobaan yang telah dilakukan, dapat dilihat bahwa terdapat perubahan nilai tegangan VCE seiring dengan dilakukannya peningkatan arus basis. Hal ini menunjukkan bahwa transistor memiliki peranan yang sangat penting dalam rangkaian. Pada saat arus basis sebesar 10 µA, nilai VCE adalah 4,8 V, sedangkan pada saat arus basis sebesar 30 µA, nilai VCE adalah 2,1 V, dan pada saat arus basis sebesar 50 µA, nilai VCE adalah 0,2 V. Dengan demikian, dapat disimpulkan bahwa semakin besar arus basis, maka semakin kecil tegangan VCE.

Sesudah (ringkas, angka dan sebab):
> Tegangan VCE turun dari 4,8 V menjadi 0,2 V ketika IB dinaikkan dari 10 µA ke 50 µA (Tabel 6.1). Penurunan itu terjadi karena arus kolektor, yang besarnya kira-kira β kali arus basis, ikut naik sehingga tegangan yang jatuh pada resistor kolektor RC di rangkaian ini juga makin besar. Karena itu, transistor sudah saturasi pada 50 µA. Pada kondisi itu LED menyala penuh, dan VCE terukur hanya berselisih 4,3 persen dari perhitungan. Selisih ini kemungkinan berasal dari toleransi resistor 5 persen.

Diukur dengan `metrik.py --preset akademik`: versi Sebelum gagal di empat metrik (rerata kalimat 21 kata, kalimat 28 kata ke atas 25 persen, koordinasi 2,0, konektor 1,19 per 100 kata) dan memuat dua frasa kosong; versi Sesudah lulus semuanya (rerata 16,2, simpangan baku 7,7). Perhatikan juga ritmenya: kalimat panjang penjelas diapit kalimat pendek yang dibuka penghubung. Paragraf yang semua kalimatnya sama panjang tetap terbaca AI walaupun ringkas.

Periksa setiap paragraf analisis dengan daftar di atas sebelum langkah 5b, lalu jalankan 5b.

### 5. Gaya bahasa
Bahasa Indonesia baku tapi luwes, seperti mahasiswa yang menulis rapi.
- Satu gagasan per kalimat. Variasikan panjang: sisipkan kalimat pendek (lima sampai delapan kata) di antara kalimat panjang. Cek bahwa selisih terpanjang dan terpendek dalam satu paragraf lebih dari lima kata. Kalimat pendek yang bukan penutup paragraf diberi penghubung atau digabung ke kalimat sesudahnya (humanizer, bagian 2.15).
- Hindari nominalisasi berlebih: "mengatur", bukan "melakukan pengaturan"; "menguji", bukan "dilakukan pengujian".
- Hindari "di mana" dan "yang mana" sebagai penghubung, pleonasme ("agar supaya", "saling satu sama lain"), dan konjungsi berpasangan bertubi.
- Hindari klise: "sangat penting untuk dicatat", "dalam era modern ini", "secara keseluruhan dapat disimpulkan bahwa".
- Tanpa em dash dan en dash di badan teks. Pakai koma, titik dua, kurung, atau pecah kalimat.
- Tanpa kata ganti orang pertama ("kami", "penulis").
- Langkah kerja memakai pola kalimat yang seragam dalam satu daftar (semua imperatif atau semua pasif). Hasil dan analisis memakai kalimat aktif.
- Istilah asing belum baku ditulis miring (*collector*, *breadboard*, *waveform*, *ground*). Singkatan teknis (LED, BJT, DC, VCC) tidak dimiringkan. Heading tidak pernah dimiringkan.
- Imbuhan mengikuti KBBI: mengonfigurasi, memverifikasi, menganalisis.
- Istilah teknis yang jarang dipakai dijelaskan sekali, singkat, pada kemunculan pertama.

### 5b. Humanizer bahasa Indonesia (wajib untuk prosa)
Skill ini membawa Humanizer Indonesia 1.7.2 karya Pribadi Mumpuni Adhi (MIT) di folder `humanizer-indonesia/`, relatif terhadap folder skill ini. Bagian 5 di atas adalah ringkasannya; aturan lengkapnya ada di sana.

**Yang disunting**: prosa yang ditulis skill ini, yaitu paragraf analisis, pengantar, penjelasan alat, butir Tujuan, dan butir Kesimpulan. Langkah kerja hanya diperiksa kalke dan nominalisasinya; pola imperatif atau pasif yang seragam dibiarkan, karena pasif di Metode dan Langkah kerja adalah konvensi (katalog bagian F).
**Yang tidak disentuh**: heading, caption, isi tabel, kode, persamaan, angka dan satuan, rujukan Gambar, Tabel, dan Persamaan, sitasi, kutipan langsung, serta teks yang diminta pengguna dipakai apa adanya.

Langkah ini tidak menunggu permintaan pengguna. Setiap pemanggilan laprak menjalankannya. Untuk analisis, langkah 4b dijalankan lebih dulu, lalu urutan di bawah. Aturan humanizer "tulis ulang, jangan pangkas" berlaku untuk teks milik pengguna; untuk prosa yang ditulis skill ini sendiri, padding yang dilarang di 4b justru dibuang.

Urutan kerja, setelah draf satu bagian selesai dan sebelum ditulis ke dokumen:
1. Baca `humanizer-indonesia/PANDUAN.md` (Prinsip dasar, Alur kerja Langkah 1 sampai 8, Kaidah yang sering terlewat, Yang tidak boleh ditandai) dan `humanizer-indonesia/references/id-ai-tells.md`, sekali per sesi.
2. Register `akademik`. Bila modul atau sumber teori berbahasa Inggris, anggap jawabannya "ya" untuk pertanyaan Langkah 1, sehingga Langkah 5b (uji kalke) menjadi pekerjaan utama.
3. Preset suara: pakai `voice-*.md` milik pengguna bila ada di folder `humanizer-indonesia/references/` atau contoh tulisan pengguna di percakapan. Jangan memakai `voice-prib-id.md` kecuali pengguna memintanya; berkas itu preset milik penulis humanizer. Tanpa preset, pakai suara netral akademik dan sebutkan bahwa kalibrasi suara tidak dilakukan.
4. Jalankan Langkah 5, 5b, 5c, dan 5d PANDUAN. Keempatnya wajib, juga ketika metrik sudah lulus.
5. Ukur bila Python tersedia: tulis prosa bagian itu ke berkas `.md` sementara di folder kerja sendiri (bukan folder pengguna), lalu `python humanizer-indonesia/scripts/metrik.py prosa.md --preset akademik`. Baris PERIKSA dan diagnostik menjadi daftar kerja revisi. Tanpa Python (misalnya hanya add-in Word), hitung manual rerata dan simpangan baku panjang kalimat. Metrik hanya dilaporkan bila benar-benar dihitung; selain itu tulis *tidak diukur*.
6. Audit dan revisi akhir sesuai Langkah 6 dan 7 PANDUAN, lalu tulis hasilnya ke dokumen.

Bila aturan humanizer dan laprak bertentangan, aturan laprak yang dipakai: Format umum, Aturan isi (Tujuan dua kalimat, Kesimpulan tanpa data kelompok, tanpa Dasar Teori), istilah asing miring, tanpa kata ganti orang pertama, dan *jika* sebagai konjungsi syarat. Aturan humanizer "tulis ulang, jangan pangkas" dan "pertahankan makna dan angka" selalu berlaku.

Pengguna yang membawa laporan hasil suntingan model lain (Gemini, ChatGPT) dan minta dicek: ikuti bagian "Menerima berkas yang sudah disunting model lain" di PANDUAN, terutama pemeriksaan rujukan silang dan angka yang hilang.

Humanizer dipakai supaya laporan enak dibaca, bukan untuk mengelabui detektor AI; PANDUAN mencatat teks hasil suntingan otomatis tetap terdeteksi sebagai tulisan AI. Hasilnya draf yang tetap harus dibaca dan disunting praktikan. Bila dosen atau kampus mewajibkan pengungkapan pemakaian AI, ingatkan pengguna di penutup.

### 6. Tabel, persamaan, kode, gambar, rujukan
Aturan caption dan nomornya ada di Bagian 9. Di sini yang berlaku untuk isi:
- Setelah menyisipkan caption, baca ulang teksnya. Bila field gagal dibuat, laporkan dan jangan mengetik nomor diam-diam tanpa memberi tahu pengguna.
- Setiap gambar dan tabel dirujuk minimal sekali di prosa ("seperti pada Gambar 6.2", "Tabel 6.1 memuat..."), dengan huruf kapital pada Gambar dan Tabel.
- Persamaan ditulis di baris sendiri, rata tengah, dengan nomor di kanan: (6.1), (6.2). Simbol dijelaskan satu kali setelah persamaan. Rumus dari pengguna dipakai apa adanya.
- Tabel data: tanpa warna latar, header tebal hitam, isi rata tengah, angka dengan desimal konsisten dan satuan di header. Bila tabel di template dosen sudah berwarna, biarkan tabel itu; tabel baru yang dibuat skill tetap tanpa warna. Ikuti gaya tabel yang sudah ada di dokumen (baca `style` dan `headerRowCount` dari tabel sibling).
- Kode program atau perintah: dalam tabel satu sel `Table Grid` (bukan shape), Courier New 10 pt, spasi 1, before 0 after 0, warna hitam, tidak miring.
- Gambar: rata tengah, rasio asli dijaga, lebar tidak melebihi area teks (maksimal sekitar 14 cm untuk profil Eldas). Paragraf gambar tidak dipisah dari caption.

### 7. Sumber dan sitasi
- Teori, angka, dan data selain milik pengguna hanya dari sumber tepercaya: buku teks universitas, jurnal, datasheet resmi produsen, situs pemerintah atau universitas. Jangan Wikipedia, blog, atau situs tanpa penulis jelas.
- Sitasi dalam teks (Penulis, tahun). Daftar pustaka APA di akhir. Setiap sumber di daftar pustaka dikutip di teks, dan sebaliknya.
- Jangan menulis sitasi dari ingatan. Verifikasi lewat pencarian web atau gunakan sumber yang diberikan pengguna. Bila tidak yakin, tulis `[CEK SUMBER: ...]` dan masukkan ke daftar cek.
- Tujuan dan Kesimpulan tidak membutuhkan sitasi. Pembahasan boleh, secukupnya.

### 8. Verifikasi
- Setelah tiap bagian, baca ulang teks dan style yang baru ditulis, lalu bandingkan dengan paragraf template di sekitarnya (font, ukuran, spasi, before dan after, indent). Acuannya template, bukan angka tetap.
- Setelah semua selesai: `verify_doc` untuk struktur dan penomoran, `verify_doc_visual` untuk tampilan, lalu perbarui field (caption, daftar isi bila ada).
- Cari sisa penanda `[ISI:`, `[CEK SUMBER:`, dan petunjuk dalam kurung. Tidak boleh ada yang tertinggal tanpa dilaporkan.
- Cek cepat bahasa: cari "di mana", "yang mana", "melakukan", "dilakukan", "terhadap", em dash. Tinjau tiap temuan satu per satu, jangan diganti otomatis.
- Humanizer: pastikan langkah 5b sudah dijalankan untuk setiap bagian prosa. Bila Python tersedia, jalankan `humanizer-indonesia/scripts/metrik.py` sekali lagi pada seluruh prosa BAB dan catat baris yang masih PERIKSA.
- Cek penomoran sesuai Bagian 9: caption berbentuk bab.urutan, nomor mulai lagi tiap bab, heading memakai style Heading, dan (mode dokumen lengkap) nomor halaman tiap bagian benar.
- **Lint rujukan** (kedua jalur; di Jalur A jalankan lewat `execute_office_js` dengan membaca teks isi dokumen). Kumpulkan semua paragraf prosa dan sel tabel, lalu periksa:
  1. Penanda tersisa, pola `\[(ISI|CEK SUMBER)\s*:[^\]]*\]`.
  2. Rujukan, pola `\b(Gambar|Tabel)\s+(\d+)\.(\d+)\b`: nomor BAB harus sama dengan BAB dokumen, dan nomor urut harus ada. Jumlah gambar dan tabel dihitung dari paragraf caption (style `Caption`) yang diawali "Gambar" atau "Tabel".
  3. Setiap gambar dan tabel dirujuk minimal sekali di prosa.
  4. Pola gaya: em dash atau en dash, "di mana" dan "yang mana", "melakukan" dan "dilakukan", "kami" dan "penulis", "membuktikan".
  5. Aturan isi: tidak ada bagian Dasar Teori yang tidak diminta, butir Tujuan paling banyak dua kalimat, Kesimpulan tanpa penanda dan tanpa data kelompok (alamat IP, nilai ukur).
  6. Format umum: before/after 0 dan TNR 12 di paragraf, heading, dan tabel; caption dan tabel tidak miring; page break hanya sebelum sub-bab level 2.
  Bila ada berkas .docx dan Python, `python scripts/lint_laprak.py Laporan.docx` dari folder skill ini menjalankan butir 1 sampai 5 sekaligus.
  Hasil lint dilaporkan sebagai daftar temuan. Tiap temuan ditimbang, tidak diganti otomatis. Rujukan yang nomornya salah dilaporkan dan dibetulkan hanya bila jelas gambar mana yang dimaksud.

### 9. Penomoran akademik
Aturan baku karya ilmiah Indonesia. Bila template atau aturan kampus berbeda, template dan aturan kampus menang. Bagian ini menggantikan skill `penomoran-akademik` yang dulu terpisah, jadi seluruh aturan penomoran ada di sini.

Cara merespons:
- Setiap kali diminta membuat kerangka bab, merapikan teks, atau mengecek draf laporan, makalah, proposal, atau skripsi, terapkan aturan di bawah secara otomatis, tanpa menunggu diminta. Ini juga berlaku untuk dokumen yang bukan laporan praktikum.
- Bila pengguna menyebut aturan khusus kampus, aturan itu didahulukan.
- Untuk dokumen non-praktikum (makalah, proposal, skripsi), Format umum dan Aturan isi laprak tidak dipaksakan. Yang berlaku: Bagian 9 ini, Prinsip utama, gaya bahasa (Bagian 5), dan humanizer (5b) bila pengguna meminta teksnya dirapikan.

#### 9.1 Caption, tabel, gambar (berlaku di kedua mode)
- Tabel dan gambar bernomor per bab: Tabel 2.1 berarti tabel pertama di Bab II, Gambar 3.2 berarti gambar kedua di Bab III.
- Judul tabel di atas tabel. Judul gambar di bawah gambar.
- Caption wajib fitur caption Word, bukan teks ketikan. Cara baku: `Range.InsertCaption` dengan label manual berisi nomor bab, misalnya "Gambar 6." atau "Tabel 6." (Below untuk gambar, Above untuk tabel). Hasilnya field `SEQ Gambar_6. \* ARABIC`. Word menyisipkan spasi antara label dan nomor; hapus spasi itu sehingga teks menjadi "Gambar 6.1 Keterangan". Lewat COM, matikan sementara `Options.SmartCutPaste` saat menghapus spasi itu (lalu pulihkan), karena kalau tidak Word mengembalikannya.
- Lewat add-in tanpa InsertCaption: sisipkan teks "Gambar 6." lalu field `SEQ Gambar_6. \* ARABIC` tanpa spasi di antaranya. Hasilnya identik dengan InsertCaption.
- Bila template sudah memakai pola field lain (misalnya `STYLEREF 1 \s` + `SEQ ... \s 1`), ikuti template. Setelah menyisipkan, perbarui semua field dan baca ulang teksnya.
- Rujukan di badan teks ditulis "Gambar 6.1" dan "Tabel 6.2", sama persis dengan caption.
- Selesai bila seluruh caption berbentuk bab.urutan dan urutannya mulai lagi tiap bab.

#### 9.2 Hierarki bab dan sub-bab
- Tingkat 1 (Bab): BAB I, BAB II, ... Romawi besar, huruf kapital semua, tebal, rata tengah.
- Tingkat 2 (Sub-bab): A., B., C. tebal.
- Tingkat 3: 1., 2., 3. tebal.
- Tingkat 4: a., b., c. huruf kecil.
- Tingkat 5: 1), 2), 3).
- Tingkat 6: a), b), c).
- Template laprak umumnya memakai format desimal (misalnya 6.3.1.1 untuk Eldas). Bila template atau pengguna memakai format desimal, pakai itu secara konsisten dan hindari bullet untuk struktur bab. Pola A., 1., a. di atas dipakai hanya untuk kerangka dokumen baru yang tidak punya template.
- Heading memakai style Heading 1, 2, 3 (bukan teks tebal biasa) supaya bisa dibuat daftar isi. Penomoran selalu otomatis lewat daftar bertingkat Word, tidak diketik manual.
- Selesai bila `verify_doc` menunjukkan style Heading terpasang dan nomor daftar berurutan tanpa lompat.

#### 9.3 Nomor halaman (mode dokumen lengkap saja)
- **Bagian awal** (Kata Pengantar, Abstrak, Daftar Isi, Daftar Tabel, Daftar Gambar): angka Romawi kecil (i, ii, iii, ...) di bawah tengah. Cover dan Halaman Pengesahan tetap dihitung urutannya, tetapi nomornya tidak ditampilkan.
- **Bagian inti** (Bab I sampai Bab Penutup): angka Arab, mulai dari 1 di halaman pertama Bab I.
  - Halaman awal bab (yang memuat judul BAB): nomor di bawah tengah.
  - Halaman lanjutan bab: nomor di pojok kanan atas.
- **Bagian akhir** (Daftar Pustaka, Lampiran): nomor Arab melanjutkan bagian inti tanpa mengulang dari 1, posisi seperti halaman lanjutan (pojok kanan atas).
- Implementasi di Word: pisahkan bagian dengan Section Break (Next Page), matikan Link to Previous di header dan footer tiap bagian, atur Format Page Number (Romawi kecil atau Arab, Start at), dan pakai Different First Page untuk halaman awal bab. Nomor halaman memakai field, bukan teks ketikan.
- Selesai bila tiap bagian punya format dan posisi nomor sesuai aturan, dan halaman pertama Bab I bernomor 1.
- Bila pengguna bertanya cara mengaturnya di Word atau Google Docs, beri langkah teknis memakai Section Break (Next Page) dan mematikan Link to Previous.

## Hal yang tidak dilakukan
- Tidak menghapus gambar, tabel, atau paragraf yang menjadi jangkar gambar.
- Tidak menghapus bagian "Laporan", "Catatan", atau "Kontak Asisten" di akhir template tanpa izin. Itu ketentuan dosen. Tanyakan apakah perlu dibuang sebelum dikumpulkan.
- Tidak mengubah teks tracked change atau komentar orang lain.
- Tidak menyalin teks dari laporan lain. Ada aturan plagiarisme di ketentuan.
- Tidak menyentuh section break, nomor halaman, cover, dan daftar isi pada mode isi BAB kecuali diminta.

## Variasi profil
Eldas dan SRWE hanyalah dua profil yang sudah dikenal, bukan batas cakupan skill. Skill ini berlaku untuk laporan praktikum mata kuliah apa pun. Untuk mata kuliah lain, template dosen menjadi acuan penuh, dan dua profil di bawah hanya contoh cara membaca ciri template.

- **Eldas**: judul tanpa heading bertingkat berlebih, heading sampai empat tingkat (6.3.1.1), caption seluruhnya style `Caption`, header tabel tebal hitam tanpa warna, tabel hasil boleh muncul di Hasil dan Analisis.
- **SRWE**: judul "BAB V", Tujuan, Percobaan, Hasil dan Analisis, Kesimpulan, ada tabel kode, header tabel tebal hitam tanpa warna.
Bila profil tidak jelas, baca template dulu dan tanyakan satu pertanyaan. Bila ternyata profil baru (mata kuliah lain), kerjakan mengikuti template dan di akhir tawarkan menyimpan cirinya sebagai profil baru di sini.

## Penutup setiap tugas
Balasan chat singkat berisi:
1. Bagian apa yang sudah diisi (dengan tautan heading).
2. Yang perlu dicek atau dilengkapi pengguna: foto atau screenshot yang belum ada, data yang belum diberikan, sumber yang perlu diverifikasi, bagian ketentuan di akhir template yang perlu dihapus, dan nomor caption yang perlu dilihat setelah pembaruan field.
3. Bila ada data atau rumus pengguna yang tampak janggal, satu kalimat penjelasan.
4. Hasil humanizer: metrik akhir (atau *tidak diukur*), preset suara yang dipakai, dan poin konfirmasi substansi dari Langkah 8 PANDUAN. Ingatkan bahwa prosa adalah draf yang perlu dibaca ulang, dan bahwa pemakaian AI perlu diungkapkan bila dosen memintanya.