# Menulis isi laporan dalam bahasa Indonesia

Diturunkan dari skill `humanizer-indonesia` (adhiprib), disempitkan ke konteks laporan praktikum.

## Prinsip

Kekakuan teks Indonesia jarang berasal dari salah tata bahasa. Penyebabnya kepadatan nominalisasi, kohesi yang digantung pada konektor, dan keseragaman panjang kalimat. Perbaiki tiga itu dulu sebelum menyentuh diksi.

## Larangan keras

- Tanpa em dash (—) dan en dash (–) di badan prosa. Pakai koma, titik dua, kurung, atau pecah jadi dua kalimat.
- Tanpa bullet di badan laporan. Yang berupa daftar bernomor hanya Tujuan, langkah percobaan, dan Kesimpulan.
- Tanpa klaim derajat kepastian yang dibuat-buat: "hal ini membuktikan bahwa", "dapat disimpulkan secara pasti". Laporan praktikum melaporkan pengamatan, bukan membuktikan teori.

## Tell yang harus dihindari

**A1 Nominalisasi berlebih.** Verba dijadikan nomina lalu dipasangkan verba ringan.
- Buruk: "melakukan pengaturan terhadap alamat IP", "dilakukan pengujian konektivitas"
- Baik: "mengatur alamat IP", "menguji konektivitas"

**A3 `di mana` / `yang mana` sebagai penghubung.** Kalke dari *where* dan *which*.
- Buruk: "Pool DHCP dibuat di mana rentang alamat ditentukan"
- Baik: "Pool DHCP dibuat dengan rentang alamat tertentu"

**A4 Panjang kalimat seragam.** Kalau semua kalimat 20 sampai 25 kata, teks terbaca sebagai mesin. Sisipkan kalimat pendek lima sampai delapan kata. Contoh: "Ping gagal." lalu kalimat panjang yang menjelaskan sebabnya.

**A7 Konjungsi berpasangan bertubi.** "Tidak hanya ... tetapi juga", "baik ... maupun". Cukup sekali per bagian.

**B1 Pleonasme.** "saling berkomunikasi satu sama lain", "agar supaya", "demi untuk", "sangat penting sekali".

**B4 Derajat kepastian.** Hindari "tentu saja", "sudah pasti", "tidak diragukan lagi".

**B5 Klaim tanpa rujukan.** Jangan menulis "menurut penelitian" kalau tidak ada rujukan di modul.

## Register laporan praktikum

Dokumen acuan memakai campuran yang konsisten:

- **Bagian Percobaan**: kalimat pasif atau imperatif yang melaporkan tindakan. "Berkas topologi Percobaan 1 diunduh untuk mengamati susunan perangkat." atau "Konfigurasikan ISP-Router dengan alamat berikut." Pilih satu pola per percobaan, jangan berganti-ganti di dalam satu daftar.
- **Bagian Hasil dan Analisis**: kalimat aktif yang menerangkan sebab akibat. "Ping ke gateway digunakan untuk memverifikasi bahwa klien tidak hanya memperoleh alamat, tetapi juga dapat menjangkau jalur keluar."
- **Kesimpulan**: kalimat padat satu sampai dua baris per butir, subjek berupa teknologi yang dibahas.

Hindari kata ganti orang pertama. Tidak ada "kami" atau "penulis".

## Istilah asing

Istilah jaringan yang belum punya padanan baku ditulis miring: *gateway*, *access point*, *router*, *switch*, *server*, *port*, *interface*, *pool*, *ping*, *troubleshooting*. Singkatan teknis tidak dimiringkan: DHCP, VLAN, SSID, WLAN, NAT, IP, DNS, CAPWAP.

Bentuk berimbuhan ikut kaidah Indonesia: "mengonfigurasi", bukan "mengkonfigurasi". "memverifikasi", "menganalisis", "mengelola".

## Rujukan silang

Selalu sebut nomor gambar atau tabel di badan kalimat: "seperti pada Gambar 5.4", "Tabel 5.1 memuat alamat tiap perangkat". Huruf kapital pada Gambar dan Tabel.

## Pemeriksaan cepat sebelum kirim

1. Cari `di mana`, `yang mana`, `melakukan`, `dilakukan`, `terhadap`. Tiap temuan ditimbang ulang.
2. Hitung panjang kalimat tiap paragraf. Kalau selisih maksimum dan minimum kurang dari lima kata, pecah satu kalimat.
3. Baca paragraf analisis keras-keras. Kalau tidak ada tempat bernapas, tambahkan kalimat pendek.
4. Pastikan tiap gambar dirujuk minimal sekali di prosa.
