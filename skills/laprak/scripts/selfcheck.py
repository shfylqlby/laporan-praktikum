#!/usr/bin/env python3
"""Uji mandiri generator laporan untuk tiap profil. Jalankan: python selfcheck.py"""
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from PIL import Image

import build_laprak as B
import lint_laprak as L


def spec(profil, bab):
    return {
        "profil": profil,
        "bab": bab,
        "judul": "Judul Modul Uji",
        "bagian": [
            {"judul": "Tujuan", "daftar": ["Praktikan mampu menjelaskan konsep."]},
            {"judul": "Percobaan", "sub": [{
                "judul": "Percobaan 1",
                "isi": [{"tabel": {"caption": "Alamat Perangkat",
                                   "header": ["Devices", "IP"],
                                   "rows": [["Router", "192.168.1.1"]],
                                   "widths": [2.9, 3.0]}}],
                "daftar": [
                    {"teks": "Berkas topologi diunduh.", "gambar": {"file": "a.png", "caption": "Topologi"}},
                    {"teks": "Switch dikonfigurasi.", "kode": ["Switch>enable", "Switch#configure terminal"]},
                    {"teks": "Hasil ping dicatat.", "gambar": {"file": "belum-ada.png", "caption": "Hasil Ping"}},
                ],
            }]},
            {"judul": "Hasil dan Analisis", "sub": [{
                "judul": "Percobaan 1",
                "isi": [{"paragraf": [f"Seluruh tautan aktif dan port trunk menyala, seperti pada Gambar {bab}.1 dan Gambar {bab}.9."],
                         "persamaan": {"rumus": "Galat = |Vu - Vt| / Vt x 100%",
                                       "keterangan": ["dengan Vu nilai ukur dan Vt nilai teori."]}}],
            }]},
            {"judul": "Kesimpulan", "daftar": ["VLAN memisahkan domain siar."]},
            {"judul": "Daftar Pustaka", "pustaka": [
                "Zeta, A. (2020). *Jaringan Komputer*. Penerbit.",
                "Alpha, B. (2019). *Elektronika Dasar*. Penerbit.",
            ]},
        ],
    }


def cek(profil, bab, tmp):
    s = dict(spec(profil, bab), base_dir=tmp)
    out = os.path.join(tmp, f"{profil}.docx")
    b = B.Builder(s)
    b.build(out)
    P = b.P
    d = Document(out)
    styles = [p.style.name for p in d.paragraphs]

    assert P["judul_style"] in styles, "judul hilang"
    assert styles.count(P["heading"][0]) == 5, f"{profil}: heading tingkat 1 harus 5"
    assert styles.count(P["heading"][1]) == 2, f"{profil}: heading tingkat 2 harus 2"

    numb = d.part.numbering_part.element
    texts = {l.find(qn("w:lvlText")).get(qn("w:val")) for l in numb.iter(qn("w:lvl"))
             if l.find(qn("w:lvlText")) is not None}
    assert f"{bab}.%1" in texts, f"{profil}: lvlText heading tidak ditambal"

    fields = [f.get(qn("w:instr")) for f in d.element.body.iter(qn("w:fldSimple"))]
    sfx = P["seq_suffix"]
    assert any(f"SEQ Gambar_{bab}.{sfx} " in f for f in fields), f"{profil}: SEQ gambar salah"
    assert any(f"SEQ Tabel_{bab}.{sfx} " in f for f in fields), f"{profil}: SEQ tabel salah"

    t0 = d.tables[0]
    # tabel tanpa warna: tidak ada shading di sel mana pun, teks hitam, header hanya tebal
    for row in t0.rows:
        for cell in row.cells:
            assert cell._tc.tcPr.find(qn("w:shd")) is None, f"{profil}: sel tabel tidak boleh berwarna"
            for r in cell.paragraphs[0].runs:
                assert r.font.color.rgb == RGBColor(0, 0, 0), f"{profil}: teks tabel harus hitam"
    assert t0.rows[0].cells[0].paragraphs[0].runs[0].bold, f"{profil}: header tabel tidak tebal"
    assert not t0.rows[1].cells[0].paragraphs[0].runs[0].bold, f"{profil}: isi tabel tidak boleh tebal"
    assert any("tabel berwarna" in x for x in L.lint_warna(_dok_berwarna(tmp))), f"{profil}: lint tidak menangkap tabel berwarna"

    tcode = d.tables[1]
    assert tcode.cell(0, 0).paragraphs[0].runs[0].font.name == P["kode"]["font"], f"{profil}: font kode salah"
    assert len(tcode.cell(0, 0).paragraphs) == 2, f"{profil}: satu paragraf per baris perintah"

    g = P["gambar"]
    n_img = 0
    for e in d.element.body.iter(qn("wp:extent")):
        n_img += 1
        cx, cy = int(e.get("cx")), int(e.get("cy"))
        assert cx <= int(Cm(g["max_w"])) + 1, f"{profil}: gambar terlalu lebar"
        assert cy <= int(Cm(g["max_h"])) + 1, f"{profil}: gambar terlalu tinggi"
        assert abs(cx / cy - 1.8) < 0.02, f"{profil}: rasio gambar berubah"
    assert n_img == 1, f"{profil}: jumlah gambar salah"

    p = next(p for p in d.paragraphs if "dikonfigurasi" in p.text)
    assert any(r.italic and r.text.strip().lower() == "switch" for r in p.runs), \
        f"{profil}: istilah asing tidak dimiringkan"

    brks = sum(1 for p in d.paragraphs if p.paragraph_format.page_break_before)
    assert brks == 4, f"{profil}: pemisah halaman harus 4, dapat {brks}"

    for p in d.paragraphs:
        if p.style.name in P["heading"]:
            assert not any(r.italic for r in p.runs), f"{profil}: heading tidak boleh miring"

    # caption: "Gambar 5." menempel langsung ke nomor SEQ, tanpa spasi
    caps = [p for p in d.paragraphs if p.text.startswith("Gambar")]
    assert caps and caps[0].runs[0].text == f"Gambar {bab}.", f"{profil}: awalan caption salah"

    # kode: spasi 1, before/after 0
    for p in tcode.cell(0, 0).paragraphs:
        pf = p.paragraph_format
        assert pf.line_spacing == 1.0 and pf.space_after == 0 and pf.space_before == 0, \
            f"{profil}: spasi kode harus 1 dengan before/after 0"

    # gambar yang belum ada jadi penanda, build tetap jalan
    assert any("belum-ada.png" in w for w in b.peringatan), f"{profil}: gambar hilang tidak diperingatkan"
    assert any("[ISI: gambar belum-ada.png]" in t.cell(0, 0).text for t in d.tables), \
        f"{profil}: penanda gambar hilang tidak ada"

    # persamaan bernomor di kanan
    eq = next(p for p in d.paragraphs if "Galat" in p.text)
    assert eq.text.endswith(f"({bab}.1)"), f"{profil}: nomor persamaan salah: {eq.text!r}"

    # daftar pustaka urut abjad, indent menggantung
    pus = [p for p in d.paragraphs if "Penerbit." in p.text]
    assert [p.text[:5] for p in pus] == ["Alpha", "Zeta,"], f"{profil}: pustaka tidak urut abjad"
    ind = pus[0]._p.pPr.find(qn("w:ind"))
    assert ind.get(qn("w:hanging")) == "720", f"{profil}: pustaka tanpa indent menggantung"

    # Word diminta memperbarui field saat dibuka
    assert d.settings.element.find(qn("w:updateFields")) is not None, f"{profil}: updateFields tidak ada"

    # lint menangkap penanda, rujukan salah, dan tabel yang tidak dirujuk
    temuan = L.lint(out, bab)
    assert any("penanda tersisa" in t for t in temuan), f"{profil}: lint tidak menangkap penanda"
    assert any(f"Gambar {bab}.9 tidak ada" in t for t in temuan), f"{profil}: lint tidak menangkap rujukan salah"
    assert any(f"Tabel {bab}.1 tidak dirujuk" in t for t in temuan), f"{profil}: lint tidak menangkap tabel tak dirujuk"
    assert not any(f"Gambar {bab}.1 tidak dirujuk" in t for t in temuan), f"{profil}: lint salah menandai gambar yang dirujuk"

    cek_format_umum(d, P, profil, tcode)
    print(f"  {profil} OK")


def _dok_berwarna(tmp):
    """Dokumen kecil berisi tabel ber-header warna, untuk menguji lint warna tabel."""
    from docx.oxml import OxmlElement
    d = Document()
    t = d.add_table(rows=2, cols=2)
    sh = OxmlElement("w:shd")
    sh.set(qn("w:val"), "clear")
    sh.set(qn("w:fill"), "FF89FF")
    t.cell(0, 0)._tc.get_or_add_tcPr().append(sh)
    out = os.path.join(tmp, "berwarna.docx")
    d.save(out)
    return Document(out)


def isi(p):
    return p.text.strip() or p._p.find(".//" + qn("w:drawing")) is not None


def semua_paragraf(d):
    for p in d.paragraphs:
        yield p, None
    for t in d.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p, t


def cek_format_umum(d, P, profil, tcode):
    """Format umum semua mata kuliah: TNR 12, before/after 0, indent 36 pt, page break per sub-bab."""
    data = d.tables[0]._tbl
    for p, t in semua_paragraf(d):
        if not isi(p):
            continue
        di_kode = t is not None and t._tbl is tcode._tbl
        pf = p.paragraph_format
        assert pf.space_before == 0 and pf.space_after == 0, \
            f"{profil}: before/after harus 0: {p.text[:30]!r}"
        caption = p.style.name in (P["caption_gambar_style"], P["caption_tabel_style"]) \
            and p.text.startswith(("Gambar", "Tabel"))
        for r in p.runs:
            if not r.text.strip():
                continue
            if di_kode:
                assert (r.font.name, r.font.size) == ("Courier New", Pt(10)), f"{profil}: kode harus Courier New 10"
            elif caption:
                assert (r.font.name, r.font.size) == (B.TNR, Pt(10)), f"{profil}: caption harus TNR 10"
            else:
                assert (r.font.name, r.font.size) == (B.TNR, Pt(12)), \
                    f"{profil}: harus TNR 12: {r.text[:30]!r} = {r.font.name} {r.font.size}"
            if caption or t is not None or p.style.name in P["heading"]:
                assert not r.italic, f"{profil}: caption, tabel, dan heading tidak boleh miring: {r.text!r}"
        if caption:
            assert all(r.font.color.rgb == RGBColor(0, 0, 0) for r in p.runs if r.text.strip()), \
                f"{profil}: caption harus hitam"
            if p.text.startswith("Gambar"):
                assert pf.line_spacing == 1.5 and p.alignment == 1, f"{profil}: caption gambar spasi 1,5 rata tengah"
            else:
                assert pf.line_spacing == 1.0, f"{profil}: caption tabel spasi 1"
                assert (pf.left_indent, pf.first_line_indent) == (Pt(36), -Pt(36)), \
                    f"{profil}: caption tabel li 36 fi -36"
        elif t is not None and t._tbl is data:
            assert pf.line_spacing == 1.5 and not pf.left_indent and not pf.first_line_indent, \
                f"{profil}: isi tabel spasi 1,5 tanpa indent"
        elif p.style.name in P["heading"]:
            assert (pf.left_indent, pf.first_line_indent) == (Pt(36), -Pt(36)), f"{profil}: heading li 36 fi -36"
        elif t is None and p.style.name != P["judul_style"] and p._p.pPr.find(qn("w:numPr")) is not None:
            assert (pf.left_indent, pf.first_line_indent) == (Pt(36), -Pt(36)), f"{profil}: daftar li 36 fi -36"

    assert d.tables[0].style.name == "Table Grid", f"{profil}: tabel data harus Table Grid"

    analisis = next(p for p in d.paragraphs if p.text.startswith("Seluruh tautan"))
    assert analisis.paragraph_format.first_line_indent == Pt(36), f"{profil}: indent baris pertama 36 pt"
    assert any(r.italic and r.text == "trunk" for r in analisis.runs), f"{profil}: istilah baru tidak dimiringkan"

    # nomor heading memakai TNR 12
    for lvl in d.part.numbering_part.element.iter(qn("w:lvl")):
        txt = lvl.find(qn("w:lvlText"))
        if txt is not None and txt.get(qn("w:val")).endswith(".%1") and txt.get(qn("w:val"))[0].isdigit():
            f = lvl.find(qn("w:rPr") + "/" + qn("w:rFonts"))
            if f is not None:
                assert f.get(qn("w:ascii")) == B.TNR, f"{profil}: nomor heading bukan TNR"
                break
    else:
        raise AssertionError(f"{profil}: nomor heading tanpa font TNR")

    # page break hanya sebelum sub-bab level 2 (x.2, x.3, ...), tidak di tempat lain
    h2 = [p for p in d.paragraphs if p.style.name == P["heading"][0]]
    for i, p in enumerate(h2):
        assert bool(p.paragraph_format.page_break_before) == (i > 0), \
            f"{profil}: page break salah di {p.text!r}"
    lain = [p.text for p in d.paragraphs
            if p.paragraph_format.page_break_before and p.style.name != P["heading"][0]]
    assert not lain, f"{profil}: page break di luar sub-bab level 2"


def cek_lint_isi(tmp):
    s = {"profil": "srwe", "bab": 3, "judul": "Uji Isi", "bagian": [
        {"judul": "Tujuan", "daftar": [
            "Memahami VLAN. Mengonfigurasi trunk.",
            "Satu. Dua. Tiga kalimat.",
        ]},
        {"judul": "Dasar Teori", "pengantar": ["Teori singkat."]},
        {"judul": "Kesimpulan", "daftar": [
            "VLAN memisahkan domain siar.",
            "PC1 memperoleh alamat 192.168.1.10 dari kelompok kami.",
            "[ISI: kesimpulan percobaan 2]",
        ]},
    ]}
    out = os.path.join(tmp, "isi.docx")
    B.Builder(s).build(out)
    temuan = [t for t in L.lint(out, 3) if t.startswith("isi:")]
    assert any("Dasar Teori" in t for t in temuan), "lint isi: Dasar Teori tidak terdeteksi"
    assert sum("butir Tujuan" in t for t in temuan) == 1, f"lint isi: Tujuan 3 kalimat: {temuan}"
    assert any("192.168.1.10" in t for t in temuan), "lint isi: data di Kesimpulan tidak terdeteksi"
    assert any("Kesimpulan memuat penanda" in t for t in temuan), "lint isi: penanda di Kesimpulan"
    assert not any("domain siar" in t for t in temuan), "lint isi: kesimpulan bersih ikut ditandai"
    print("  lint isi OK")


def main():
    profiles = B.load_profiles()
    with tempfile.TemporaryDirectory() as tmp:
        Image.new("RGB", (1800, 1000), (200, 220, 240)).save(os.path.join(tmp, "a.png"))
        for i, profil in enumerate(profiles):
            cek(profil, 5 + i, tmp)
        cek_lint_isi(tmp)
    print("selfcheck OK")


if __name__ == "__main__":
    main()
