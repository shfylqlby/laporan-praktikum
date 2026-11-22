#!/usr/bin/env python3
"""Periksa laporan praktikum (.docx): penanda tersisa, rujukan Gambar/Tabel, pola gaya,
dan aturan isi (tanpa Dasar Teori, Tujuan 2 kalimat, Kesimpulan tanpa data kelompok).

Pakai: python lint_laprak.py Laporan.docx [--bab N]
Dipanggil otomatis oleh build_laprak.py setelah dokumen dibangun.
Hasilnya daftar temuan untuk ditimbang satu per satu, bukan untuk diganti otomatis.
"""
import argparse
import re
import sys

from docx import Document
from docx.oxml.ns import qn

PENANDA = re.compile(r"\[(ISI|CEK SUMBER)\s*:[^\]]*\]")
RUJUKAN = re.compile(r"\b(Gambar|Tabel)\s+(\d+)\.(\d+)\b")
CAPTION = re.compile(r"^(Gambar|Tabel)\s+(\d+)\.")
GAYA = [
    (re.compile("[—–]"), "em dash atau en dash"),
    (re.compile(r"\b(di mana|yang mana)\b", re.I), "'di mana' atau 'yang mana' sebagai penghubung"),
    (re.compile(r"\b(melakukan|dilakukan)\b", re.I), "nominalisasi 'melakukan/dilakukan'"),
    (re.compile(r"\b(kami|penulis)\b", re.I), "kata ganti orang pertama"),
    (re.compile(r"\bmembuktikan\b", re.I), "klaim 'membuktikan'"),
]


TEORI = re.compile(r"\b(dasar teori|landasan teori|tinjauan pustaka)\b", re.I)
KALIMAT = re.compile(r"(?<=[.!?])\s+(?=[A-Z0-9(\"])")
DATA = re.compile(
    r"\b\d{1,3}(?:\.\d{1,3}){3}\b"                                   # alamat IP
    r"|\b\d+(?:[.,]\d+)?\s*(?:mV|V|mA|uA|µA|A|kΩ|Ω|ohm|ms|Hz|kHz|MHz|dB)\b"  # nilai ukur
    r"|\bkelompok\b", re.I)


def isi_bagian(doc):
    """Kembalikan [(judul heading, [paragraf isi])] berurutan, hanya paragraf badan."""
    hasil, judul, isi = [], None, []
    for p in doc.paragraphs:
        if p.style.name.lower().startswith("heading"):
            if judul is not None:
                hasil.append((judul, isi))
            judul, isi = p.text.strip(), []
        elif judul is not None and p.text.strip():
            isi.append(p.text.strip())
    if judul is not None:
        hasil.append((judul, isi))
    return hasil


def lint_isi(doc):
    """Aturan isi: tanpa Dasar Teori, Tujuan maksimal 2 kalimat per butir, Kesimpulan tanpa data."""
    temuan = []
    for judul, isi in isi_bagian(doc):
        kecil = judul.lower()
        if TEORI.search(judul):
            temuan.append(f"isi: ada bagian '{judul}'; hapus bila dosen tidak memintanya")
        elif kecil.startswith("tujuan"):
            for t in isi:
                n = len([k for k in KALIMAT.split(t) if k.strip()])
                if n > 2:
                    temuan.append(f"isi: butir Tujuan {n} kalimat, cukup 2: {t[:60]}...")
        elif kecil.startswith("kesimpulan"):
            for t in isi:
                if PENANDA.search(t):
                    temuan.append(f"isi: Kesimpulan memuat penanda, harus menjawab Tujuan saja: {t[:60]}")
                for m in DATA.finditer(t):
                    temuan.append(f"isi: Kesimpulan memuat data kelompok ({m.group(0)}): {potong(t, m)}")
    return temuan


def lint_warna(doc):
    """Tabel tanpa warna: header cukup tebal hitam di latar putih. Warna sel terbaca AI slop."""
    temuan = []
    for i, t in enumerate(doc.tables, start=1):
        warna = set()
        for tc in t._tbl.iter(qn("w:tc")):
            for shd in tc.iter(qn("w:shd")):
                fill = (shd.get(qn("w:fill")) or "auto").upper()
                if fill not in ("AUTO", "FFFFFF", "NONE"):
                    warna.add(fill)
        if warna:
            sel = t.rows[0].cells[0].text.strip()[:25] if t.rows else ""
            temuan.append(f"gaya: tabel berwarna (tabel ke-{i}, sel pertama '{sel}', warna {', '.join(sorted(warna))}); hapus shading, header cukup tebal hitam")
    return temuan


def potong(teks, m, lebar=40):
    a, b = max(0, m.start() - lebar), min(len(teks), m.end() + lebar)
    return ("..." if a else "") + teks[a:b].strip() + ("..." if b < len(teks) else "")


def kumpulkan(doc):
    """Kembalikan (prosa, caption) dari paragraf badan dan sel tabel."""
    prosa, caption = [], []

    def lihat(p):
        teks = p.text.strip()
        if not teks:
            return
        if "caption" in p.style.name.lower() and CAPTION.match(teks):
            caption.append(teks)
        else:
            prosa.append(teks)

    for p in doc.paragraphs:
        lihat(p)
    for t in doc.tables:
        for row in t.rows:
            seen = set()
            for cell in row.cells:
                if id(cell._tc) in seen:
                    continue
                seen.add(id(cell._tc))
                for p in cell.paragraphs:
                    lihat(p)
    return prosa, caption


def lint(path, bab=None):
    doc = Document(path)
    prosa, caption = kumpulkan(doc)
    temuan = []

    jumlah = {"Gambar": 0, "Tabel": 0}
    for c in caption:
        m = CAPTION.match(c)
        jumlah[m.group(1)] += 1
        if bab is None:
            bab = int(m.group(2))

    for teks in prosa:
        for m in PENANDA.finditer(teks):
            temuan.append(f"penanda tersisa: {m.group(0)}")

    dirujuk = {"Gambar": set(), "Tabel": set()}
    for teks in prosa:
        for m in RUJUKAN.finditer(teks):
            jenis, b, n = m.group(1), int(m.group(2)), int(m.group(3))
            if bab is not None and b != bab:
                temuan.append(f"rujukan {m.group(0)} memakai BAB {b}, dokumen BAB {bab}: {potong(teks, m)}")
            elif n < 1 or n > jumlah[jenis]:
                temuan.append(f"rujukan {m.group(0)} tidak ada (hanya {jumlah[jenis]} {jenis.lower()}): {potong(teks, m)}")
            else:
                dirujuk[jenis].add(n)

    for jenis in ("Gambar", "Tabel"):
        for n in range(1, jumlah[jenis] + 1):
            if n not in dirujuk[jenis]:
                temuan.append(f"{jenis} {bab}.{n} tidak dirujuk di prosa")

    for teks in prosa:
        for pola, nama in GAYA:
            for m in pola.finditer(teks):
                temuan.append(f"gaya ({nama}): {potong(teks, m)}")

    temuan += lint_isi(doc)
    temuan += lint_warna(doc)
    return temuan


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("docx")
    ap.add_argument("--bab", type=int)
    a = ap.parse_args()
    temuan = lint(a.docx, a.bab)
    for t in temuan:
        print("PERINGATAN: " + t)
    print(f"lint selesai: {len(temuan)} temuan")
    sys.exit(1 if temuan else 0)


if __name__ == "__main__":
    main()
