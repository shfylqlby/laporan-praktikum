#!/usr/bin/env python3
"""Bangun laporan praktikum (.docx) dari spesifikasi JSON memakai template resmi.

Pakai: python build_laprak.py spec.json [-o keluaran.docx]
Skema spec: references/spec-schema.md
Profil format: assets/profil.json
"""
import argparse
import copy
import json
import os
import re
import sys

from docx import Document
from docx.enum.table import WD_ROW_HEIGHT_RULE, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(HERE, os.pardir, "assets")
TNR = "Times New Roman"
TEKS = Pt(12)
KECIL = Pt(10)

ISTILAH = [
    "access point", "default gateway", "gateway", "interface", "port",
    "router", "switch", "server", "client", "host", "hostname", "pool",
    "wireless", "username", "password", "passphrase", "browser",
    "firmware", "lightweight", "management", "guest", "staff",
    "end-to-end", "troubleshooting", "ping", "login", "setup", "subnet",
    "trainer board", "breadboard", "jumper", "software", "hardware",
    "forward bias", "reverse bias", "bias", "output", "input", "probe",
    "screenshot", "sheet", "datasheet", "power supply", "ground",
    "clock", "flip-flop", "truth table", "timer", "counter", "gate",
    "port security", "port-security", "native vlan", "native", "sticky",
    "dhcp snooping", "snooping", "trunk", "access port", "violation",
]


def el(tag, **attrs):
    e = OxmlElement(tag)
    for k, v in attrs.items():
        e.set(qn("w:" + k), str(v))
    return e


def load_profiles():
    with open(os.path.join(ASSETS, "profil.json"), encoding="utf-8") as f:
        return json.load(f)


# ---------- penomoran ----------

def patch_numbering(doc, bab):
    """Ganti nomor BAB pada tiap lvlText dan start level teratas."""
    numb = doc.part.numbering_part.element
    for lvl in numb.iter(qn("w:lvl")):
        fmt = lvl.find(qn("w:numFmt"))
        txt = lvl.find(qn("w:lvlText"))
        start = lvl.find(qn("w:start"))
        if txt is not None:
            v = txt.get(qn("w:val"))
            txt.set(qn("w:val"), re.sub(r"(?<![0-9%])5(?=[.%])", str(bab), v))
        if start is not None and lvl.get(qn("w:ilvl")) == "0":
            fv = fmt.get(qn("w:val")) if fmt is not None else ""
            if fv in ("upperRoman", "decimal") and start.get(qn("w:val")) == "5":
                start.set(qn("w:val"), str(bab))


def new_num(doc, levels, left=720, hanging=720):
    """Buat daftar bernomor baru agar penomoran mulai dari 1."""
    numb = doc.part.numbering_part.element
    aid = max(int(a.get(qn("w:abstractNumId"))) for a in numb.findall(qn("w:abstractNum"))) + 1
    nid = max(int(n.get(qn("w:numId"))) for n in numb.findall(qn("w:num"))) + 1

    a = el("w:abstractNum", abstractNumId=aid)
    a.append(el("w:multiLevelType", val="hybridMultilevel"))
    for i, (fmt, text) in enumerate(levels):
        lvl = el("w:lvl", ilvl=i)
        lvl.append(el("w:start", val=1))
        lvl.append(el("w:numFmt", val=fmt))
        lvl.append(el("w:lvlText", val=text))
        lvl.append(el("w:lvlJc", val="left"))
        pPr = el("w:pPr")
        pPr.append(el("w:ind", left=left + 720 * i, hanging=hanging if i == 0 else 360))
        lvl.append(pPr)
        # nomor selalu Times New Roman 12, tidak ikut font style heading template
        rPr = el("w:rPr")
        rPr.append(el("w:rFonts", ascii=TNR, hAnsi=TNR, cs=TNR, eastAsia=TNR))
        rPr.append(el("w:sz", val=24))
        rPr.append(el("w:szCs", val=24))
        lvl.append(rPr)
        a.append(lvl)
    numb.findall(qn("w:abstractNum"))[-1].addnext(a)

    n = el("w:num", numId=nid)
    n.append(el("w:abstractNumId", val=aid))
    numb.append(n)
    return nid


def set_num(p, num_id, ilvl=0):
    pPr = p._p.get_or_add_pPr()
    numPr = el("w:numPr")
    numPr.append(el("w:ilvl", val=ilvl))
    numPr.append(el("w:numId", val=num_id))
    pPr.insert(0, numPr)


def set_ind(p, **kw):
    pPr = p._p.get_or_add_pPr()
    old = pPr.find(qn("w:ind"))
    if old is not None:
        pPr.remove(old)
    pPr.append(el("w:ind", **kw))


# ---------- teks ----------

def split_italic(text, istilah):
    parts = []
    for chunk in re.split(r"(\*[^*]+\*)", text):
        if not chunk:
            continue
        if chunk.startswith("*") and chunk.endswith("*") and len(chunk) > 2:
            parts.append((chunk[1:-1], True))
        else:
            parts.extend(auto_italic(chunk, istilah))
    return parts


def auto_italic(text, istilah):
    if not istilah:
        return [(text, False)]
    # tanda hubung dihitung bagian kata agar "port-security" tidak terpenggal
    pat = re.compile(r"(?<![\w-])(" + "|".join(re.escape(t) for t in sorted(istilah, key=len, reverse=True)) + r")(?![\w-])", re.I)
    out, pos = [], 0
    for m in pat.finditer(text):
        if m.start() > pos:
            out.append((text[pos:m.start()], False))
        out.append((m.group(0), True))
        pos = m.end()
    if pos < len(text):
        out.append((text[pos:], False))
    return out or [(text, False)]


def add_runs(p, text, istilah, size=TEKS, bold=False, color=None, font=TNR, miring=True):
    """Tulis teks sebagai run. miring=False: tidak ada yang miring (heading, caption, tabel)."""
    for chunk, ital in split_italic(text, istilah):
        r = p.add_run(chunk)
        r.font.name = font
        r._element.rPr.rFonts.set(qn("w:eastAsia"), font)
        r.italic = bool(ital) if miring else False
        if bold:
            r.bold = True
        if size:
            r.font.size = size
        if color:
            r.font.color.rgb = RGBColor.from_string(color)
    return p


def rapat(p, line):
    """Spasi before dan after 0, spasi baris sesuai argumen."""
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = line


def mark_rpr(p, size_half_pt, bold=False, color="000000"):
    """Format tanda paragraf: dipakai Word untuk nomor daftar dan heading."""
    pPr = p._p.get_or_add_pPr()
    old = pPr.find(qn("w:rPr"))
    if old is not None:
        pPr.remove(old)
    rPr = el("w:rPr")
    rPr.append(el("w:rFonts", ascii=TNR, hAnsi=TNR, cs=TNR, eastAsia=TNR))
    if bold:
        rPr.append(el("w:b"))
    rPr.append(el("w:i", val="0"))
    rPr.append(el("w:color", val=color))
    rPr.append(el("w:sz", val=size_half_pt))
    rPr.append(el("w:szCs", val=size_half_pt))
    pPr.append(rPr)


def pastikan_table_grid(doc):
    """Template tanpa style Table Grid (eldas) mendapat salinan dari template srwe."""
    if any(s.style_id == "TableGrid" for s in doc.styles):
        return
    src = Document(os.path.join(ASSETS, "srwe.docx"))
    for s in src.styles.element.findall(qn("w:style")):
        if s.get(qn("w:styleId")) == "TableGrid":
            doc.styles.element.append(copy.deepcopy(s))
            return


# ---------- pembangun ----------

class Builder:
    def __init__(self, spec, ketat=False):
        self.spec = spec
        self.ketat = ketat
        self.peringatan = []
        self.n_persamaan = 0
        self.bab = int(spec["bab"])
        self.istilah = [t for t in spec.get("istilah", ISTILAH)
                        if t not in spec.get("istilah_kecuali", [])]
        self.base = spec.get("base_dir", ".")
        profiles = load_profiles()
        name = spec.get("profil", "srwe")
        if name not in profiles:
            sys.exit(f"profil tidak dikenal: {name}. Pilihan: {', '.join(profiles)}")
        self.P = profiles[name]
        self.P.update(spec.get("profil_override", {}))
        self.doc = Document(os.path.join(ASSETS, self.P["template"]))
        pastikan_table_grid(self.doc)
        patch_numbering(self.doc, self.bab)
        self.pending_break = False
        self.hnum = new_num(
            self.doc,
            [("decimal", f"{self.bab}.%1"), ("decimal", f"{self.bab}.%1.%2"),
             ("decimal", f"{self.bab}.%1.%2.%3"), ("decimal", "%4.")],
            left=self.P["heading_ind"]["left"],
            hanging=self.P["heading_ind"]["hanging"],
        )

    # -- primitif --

    def img_path(self, f):
        return f if os.path.isabs(f) else os.path.join(self.base, f)

    def heading(self, text, level):
        """level 0 = x.1, level 1 = x.1.1, level 2 = x.1.1.1"""
        p = self.doc.add_paragraph(style=self.P["heading"][level])
        add_runs(p, text, self.istilah, bold=True, color="000000", miring=False)
        set_num(p, self.hnum, level)
        set_ind(p, **self.P["heading_ind"])
        rapat(p, self.P["paragraf"]["line"])
        mark_rpr(p, 24, bold=True)
        if self.pending_break:
            p.paragraph_format.page_break_before = True
            self.pending_break = False
        return p

    def para(self, text, first_line=None):
        cfg = self.P["paragraf"]
        p = self.doc.add_paragraph()
        rapat(p, cfg["line"])
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        fl = cfg["first_line"] if first_line is None else first_line
        if fl or first_line == 0:
            set_ind(p, left=0, firstLine=fl)
        add_runs(p, text, self.istilah)
        return p

    def item(self, text, num_id):
        cfg = self.P["daftar"]
        p = self.doc.add_paragraph(style="List Paragraph")
        set_num(p, num_id)
        rapat(p, self.P["paragraf"]["line"])
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        set_ind(p, **cfg)
        mark_rpr(p, 24)
        add_runs(p, text, self.istilah)
        return p

    def jeda(self):
        """Paragraf kosong pemisah setelah tabel, tanpa before/after."""
        p = self.doc.add_paragraph()
        rapat(p, self.P["paragraf"]["line"])
        set_ind(p, left=0, firstLine=0)
        return p

    def caption(self, kind, text):
        gambar = kind == "Gambar"
        style = self.P["caption_gambar_style"] if gambar else self.P["caption_tabel_style"]
        p = self.doc.add_paragraph(style=style)
        if gambar:
            # di bawah gambar: rata tengah, spasi 1,5
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            rapat(p, 1.5)
            set_ind(p, left=0, firstLine=0)
        else:
            # di atas tabel: spasi 1, indent menggantung seperti daftar
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            rapat(p, 1.0)
            set_ind(p, **self.P["caption_tabel_ind"])
            p.paragraph_format.keep_with_next = True
        mark_rpr(p, 20)
        add_runs(p, f"{kind} {self.bab}.", [], size=KECIL, color="000000", miring=False)
        # sama dengan hasil Word InsertCaption berlabel "Gambar N.": SEQ Gambar_N.
        sfx = self.P["seq_suffix"]
        fld = el("w:fldSimple", instr=f" SEQ {kind}_{self.bab}.{sfx} \\* ARABIC ")
        inner = OxmlElement("w:r")
        rPr = el("w:rPr")
        rPr.append(el("w:rFonts", ascii=TNR, hAnsi=TNR, cs=TNR, eastAsia=TNR))
        rPr.append(el("w:i", val="0"))
        rPr.append(el("w:noProof"))
        rPr.append(el("w:color", val="000000"))
        rPr.append(el("w:sz", val=20))
        inner.append(rPr)
        t = OxmlElement("w:t")
        t.text = "1"
        inner.append(t)
        fld.append(inner)
        p._p.append(fld)
        add_runs(p, " " + text, [], size=KECIL, color="000000", miring=False)
        return p

    def figure(self, path, cap):
        path = self.img_path(path)
        if not os.path.exists(path):
            if self.ketat:
                sys.exit(f"gambar tidak ada: {path}")
            return self.figure_kosong(path, cap)
        g = self.P["gambar"]
        with Image.open(path) as im:
            w_px, h_px = im.size
        w = min(int(Cm(g["max_w"])), int(Cm(w_px / 96 * 2.54)))
        h = int(w * h_px / w_px)
        if h > int(Cm(g["max_h"])):
            w = int(w * int(Cm(g["max_h"])) / h)
        p = self.doc.add_paragraph(style=g["style"])
        rapat(p, 1.0)
        p.paragraph_format.keep_with_next = True
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_ind(p, left=0, firstLine=0)
        p.add_run().add_picture(path, width=w)
        if cap:
            self.caption("Gambar", cap)
        return p

    def figure_kosong(self, path, cap):
        """Gambar belum ada: kotak penanda [ISI: ...] supaya build tetap jalan."""
        self.peringatan.append(f"gambar tidak ada, diganti penanda: {path}")
        t = self.doc.add_table(rows=1, cols=1)
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        b = el("w:tblBorders")
        for side in ("top", "left", "bottom", "right"):
            b.append(el("w:" + side, val="dashed", sz=8, space=0, color="808080"))
        t._tbl.tblPr.append(b)
        cell = t.cell(0, 0)
        cell.width = Cm(self.P["gambar"]["max_w"] * 0.6)
        cell._tc.get_or_add_tcPr().append(el("w:vAlign", val="center"))
        t.rows[0].height = Cm(3)
        t.rows[0].height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        rapat(p, 1.0)
        p.paragraph_format.keep_with_next = True
        set_ind(p, left=0, firstLine=0)
        add_runs(p, f"[ISI: gambar {os.path.basename(path)}]", [], color="808080")
        if cap:
            self.caption("Gambar", cap)
        return p

    def code(self, lines):
        c = self.P["kode"]
        t = self.doc.add_table(rows=1, cols=1)
        t.style = self.P["tabel"]["style"]
        self.borders(t)
        t._tbl.tblPr.append(el("w:tblInd", w=c["indent"], type="dxa"))
        cell = t.cell(0, 0)
        for r in list(cell.paragraphs[0].runs):
            r._element.getparent().remove(r._element)
        for i, line in enumerate(lines):
            p = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            rapat(p, 1.0)
            set_ind(p, left=0, firstLine=0)
            r = p.add_run(line)
            r.font.name = c["font"]
            r._element.rPr.rFonts.set(qn("w:eastAsia"), c["font"])
            r.font.size = Pt(c["size"])
            r.italic = False
        self.jeda()
        return t

    def lebar_teks(self):
        s = self.doc.sections[-1]
        return int(s.page_width - s.left_margin - s.right_margin)

    def equation(self, eq):
        """Persamaan di baris sendiri, rata tengah, nomor (bab.n) di kanan."""
        if isinstance(eq, str):
            eq = {"rumus": eq}
        self.n_persamaan += 1
        lebar = self.lebar_teks()
        p = self.doc.add_paragraph()
        rapat(p, self.P["paragraf"]["line"])
        p.paragraph_format.keep_together = True
        set_ind(p, left=0, firstLine=0)
        tabs = el("w:tabs")
        tabs.append(el("w:tab", val="center", pos=lebar // 635 // 2))
        tabs.append(el("w:tab", val="right", pos=lebar // 635))
        p._p.get_or_add_pPr().append(tabs)
        p.add_run("\t")
        add_runs(p, eq["rumus"], [])
        add_runs(p, f"\t({self.bab}.{self.n_persamaan})", [], miring=False)
        for par in eq.get("keterangan", []):
            self.para(par, first_line=0)
        return p

    def references(self, items):
        """Daftar pustaka: urut abjad, indent menggantung 36 pt."""
        def kunci(s):
            return re.sub(r"[*\W_]+", "", s).lower()
        for item in sorted(items, key=kunci):
            p = self.doc.add_paragraph()
            rapat(p, self.P["paragraf"]["line"])
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            set_ind(p, **self.P["daftar"])
            add_runs(p, item, [])

    def table(self, spec):
        cfg = self.P["tabel"]
        self.caption("Tabel", spec["caption"])
        header, rows = spec["header"], spec["rows"]
        t = self.doc.add_table(rows=1 + len(rows), cols=len(header))
        t.style = cfg["style"]
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        tblPr = t._tbl.tblPr
        tblW = tblPr.find(qn("w:tblW"))
        if tblW is None:
            tblW = el("w:tblW")
            tblPr.append(tblW)
        tblW.set(qn("w:w"), str(cfg["width"]))
        tblW.set(qn("w:type"), "dxa" if cfg["width"] else "auto")
        self.borders(t)
        for j, text in enumerate(header):
            cell = t.rows[0].cells[j]
            if cfg.get("header_fill"):
                cell._tc.get_or_add_tcPr().append(
                    el("w:shd", val="clear", color="auto", fill=cfg["header_fill"]))
            # tanpa warna: header cukup tebal hitam di latar putih (warna terbaca AI slop)
            self.fill(cell, str(text), bold=True, color=cfg.get("header_color") or "000000")
        for i, row in enumerate(rows, start=1):
            for j, text in enumerate(row):
                self.fill(t.rows[i].cells[j], str(text), color="000000")
        for j, w in enumerate(spec.get("widths") or []):
            for row in t.rows:
                row.cells[j].width = Cm(w)
        self.jeda()
        return t

    def borders(self, t):
        if not self.P["tabel"].get("border"):
            return
        b = el("w:tblBorders")
        for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
            b.append(el("w:" + side, val="single", sz=4, space=0, color="000000"))
        t._tbl.tblPr.append(b)

    def fill(self, cell, text, bold=False, color=None):
        cell._tc.get_or_add_tcPr().append(el("w:vAlign", val="center"))
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        rapat(p, self.P["paragraf"]["line"])
        set_ind(p, left=0, firstLine=0)
        add_runs(p, text, [], bold=bold, color=color, miring=False)

    def brk(self):
        if self.P.get("break_antar_bagian"):
            self.pending_break = True

    # -- blok generik --

    def blok(self, b, num_id=None):
        """Render satu blok isi. Kunci yang dikenali boleh digabung."""
        if b.get("teks"):
            if num_id:
                self.item(b["teks"], num_id)
            else:
                self.para(b["teks"])
        for par in b.get("paragraf", []):
            self.para(par)
        for eq in listify(b.get("persamaan")):
            self.equation(eq)
        if b.get("kode"):
            self.code(b["kode"])
        for tb in listify(b.get("tabel")):
            self.table(tb)
        for g in gambar_list(b):
            self.figure(g["file"], g.get("caption", ""))
        for par in b.get("paragraf_akhir", []):
            self.para(par)

    def bagian(self, sec, level=0):
        """Bagian berisi sub-bagian bersarang, daftar bernomor, atau blok bebas."""
        if sec.get("judul"):
            self.heading(sec["judul"], level)
        for par in sec.get("pengantar", []):
            self.para(par)
        for b in sec.get("isi", []):
            self.blok(b)
        if sec.get("daftar"):
            nid = new_num(self.doc, [("decimal", "%1."), ("lowerLetter", "%2."),
                                     ("lowerRoman", "%3."), ("decimal", "%4.")],
                          **self.P["daftar"])
            for it in sec["daftar"]:
                if isinstance(it, str):
                    self.item(it, nid)
                else:
                    self.blok(it, nid)
        if sec.get("pustaka"):
            self.references(sec["pustaka"])
        for b in sec.get("penutup", []):
            self.blok(b)
        for sub in sec.get("sub", []):
            self.bagian(sub, level + 1)

    def build(self, out):
        s = self.spec
        self.judul()
        for i, sec in enumerate(s["bagian"]):
            if i:
                self.brk()
            self.bagian(sec, 0)
        self.minta_update_field()
        self.doc.save(out)
        return out

    def minta_update_field(self):
        """Word menawarkan memperbarui field (nomor caption) saat dokumen dibuka."""
        st = self.doc.settings.element
        if st.find(qn("w:updateFields")) is None:
            st.append(el("w:updateFields", val="true"))

    def judul(self):
        p = self.doc.add_paragraph(style=self.P["judul_style"])
        pf = p.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.line_spacing = 1.5
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_ind(p, left=0, firstLine=0)
        baris = self.spec["judul"].split("\n")
        if self.P["judul_bab"]:
            set_num(p, 1, 0)
            baris = [""] + baris
        mark_rpr(p, 24, bold=True)
        for i, line in enumerate(baris):
            r = p.add_run()
            r.font.name = TNR
            r.font.size = TEKS
            r.bold = True
            if i:
                r._element.append(OxmlElement("w:br"))
            if line:
                t = OxmlElement("w:t")
                t.text = line
                t.set(qn("xml:space"), "preserve")
                r._element.append(t)


def listify(x):
    if not x:
        return []
    return x if isinstance(x, list) else [x]


def gambar_list(entry):
    g = entry.get("gambar")
    if not g:
        return []
    if isinstance(g, str):
        return [{"file": g, "caption": entry.get("caption", "")}]
    if isinstance(g, dict):
        return [g]
    return g


def build(spec, out, ketat=False):
    return Builder(spec, ketat=ketat).build(out)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("spec")
    ap.add_argument("-o", "--output")
    ap.add_argument("--ketat", action="store_true",
                    help="berhenti bila gambar tidak ada, alih-alih memasang penanda [ISI: ...]")
    ap.add_argument("--tanpa-lint", action="store_true", help="lewati pemeriksaan setelah build")
    a = ap.parse_args()
    # utf-8-sig: spec yang ditulis PowerShell 5 / Notepad diawali BOM
    with open(a.spec, encoding="utf-8-sig") as f:
        spec = json.load(f)
    spec.setdefault("base_dir", os.path.dirname(os.path.abspath(a.spec)))
    out = a.output or spec.get("output") or "Laporan.docx"
    b = Builder(spec, ketat=a.ketat)
    b.build(out)
    temuan = list(b.peringatan)
    if not a.tanpa_lint:
        sys.path.insert(0, HERE)
        from lint_laprak import lint
        temuan += lint(out, b.bab)
    for t in temuan:
        print("PERINGATAN: " + t)
    print(f"{len(temuan)} peringatan")
    print(out)


if __name__ == "__main__":
    main()
