import json
import os
import re
import zipfile
from xml.sax.saxutils import escape

BASE = r"C:\Users\Abu Fazimaroyatih\.copilot\session-state\f8af1ad5-8bda-4f16-82f6-d9e68747d7f3\files\roads.json"
OUT = "Analisa_Jam_Alat_dan_Kebutuhan_Material.xlsx"
roads = [r for r in json.load(open(BASE, encoding="utf-8")) if r["rekomendasi"].strip()]


def ccol(n):
    out = ""
    while n:
        n, rem = divmod(n - 1, 26)
        out = chr(65 + rem) + out
    return out


def esc(v):
    return escape(str(v), {'"': "&quot;"})


class S:
    def __init__(self, name):
        self.name = name
        self.rows = []
        self.widths = {}

    def add(self, *values):
        self.rows.append(list(values))


def cell(r, c, v, style=0):
    if v is None:
        return ""
    ref = f"{ccol(c)}{r}"
    if isinstance(v, str) and v.startswith("="):
        return f'<c r="{ref}" s="{style}"><f>{esc(v[1:])}</f></c>'
    if isinstance(v, (int, float)):
        return f'<c r="{ref}" s="{style}" t="n"><v>{v}</v></c>'
    return f'<c r="{ref}" s="{style}" t="inlineStr"><is><t>{esc(v)}</t></is></c>'


def sheet_xml(s):
    rows = []
    for r, vals in enumerate(s.rows, 1):
        rows.append(f'<row r="{r}">{"".join(cell(r,c,v,1 if r == 1 else 0) for c,v in enumerate(vals,1) if v is not None)}</row>')
    cols = "".join(f'<col min="{c}" max="{c}" width="{w}" customWidth="1"/>' for c,w in s.widths.items())
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<dimension ref="A1:{ccol(max(len(x) for x in s.rows))}{len(s.rows)}"/><sheetViews><sheetView workbookViewId="0"/></sheetViews>
<sheetFormatPr defaultRowHeight="15"/><cols>{cols}</cols><sheetData>{''.join(rows)}</sheetData></worksheet>'''


def pct(text, code):
    m = re.search(rf"\b{code}\s*(\d+)", text or "", re.I)
    return (int(m.group(1)) / 100) if m else 0


def build():
    guide = S("Petunjuk")
    guide.add("ANALISA JAM ALAT DAN KEBUTUHAN MATERIAL")
    guide.add("Basis: rekomendasi ruas pada workbook sumber dan SE-DJBK No. 68/2024 Lampiran IV SDA.")
    guide.add("Formula produktivitas Bulldozer mengikuti contoh PDF/gambar: q=L x H^2; Ts=(L2x60/VF)+(L2x60/VR)+T3; Q=q x Fb x Fm x Fa x 60/(Ts x 60); koefisien=1/Q.")
    guide.add("Kode pekerjaan: GRD Grading, RIP Ripping, SPR Spreading, CMP Compacting, DRN Drainase, SUB Perbaikan tanah dasar, BSC Basecourse baru, PTH Patching lubang, SEM Semenisasi.")
    guide.add("Kode alat: AL-DZR Dozer/Bulldozer, AL-GRD Motor Grader, AL-RLR Roller/Compactor, AL-EXC Excavator, AL-WTR Water Tanker.")
    guide.add("Sel input pada sheet Asumsi dapat diganti. Semua jam, biaya, dan material pada sheet analisa menggunakan formula antar-sheet.")
    guide.add("Ruas kondisi baik tanpa rekomendasi tidak dihitung.")
    guide.widths = {1: 125}

    asum = S("Asumsi")
    asum.add("PARAMETER GEOMETRIK DAN MATERIAL")
    asum.add("Parameter", "Nilai", "Satuan", "Keterangan")
    asum.add("Lebar jalan JU", 5, "m", "Input lapangan")
    asum.add("Lebar jalan JP", 4, "m", "Input lapangan")
    asum.add("Tebal olah ripping/compact", 0.15, "m", "Input spesifikasi")
    asum.add("Tebal basecourse baru", 0.20, "m", "Input spesifikasi")
    asum.add("Lebar efektif drainase", 0.50, "m", "Asumsi penampang")
    asum.add("Kedalaman efektif drainase", 0.40, "m", "Asumsi penampang")
    asum.add("Air pemadatan", 0.05, "m3/m3", "Input kadar air")
    asum.add()
    asum.add("PARAMETER FORMULA BULLDOZER - SE-DJBK 68/2024 LAMPIRAN IV")
    asum.add("Parameter", "Nilai", "Satuan", "Keterangan")
    asum.add("Lebar pisau L", 3.175, "m", "Contoh PDF/gambar")
    asum.add("Tinggi pisau H", 1.30, "m", "Contoh PDF/gambar")
    asum.add("Faktor pisau Fb", 1.00, "-", "Tabel faktor")
    asum.add("Faktor kemiringan Fm", 1.00, "-", "Gambar faktor")
    asum.add("Faktor efisiensi Fa", 0.83, "-", "Tabel efisiensi")
    asum.add("Kecepatan maju VF", 3.40, "km/jam", "Contoh PDF/gambar")
    asum.add("Kecepatan mundur VR", 6.97, "km/jam", "Contoh PDF/gambar")
    asum.add("Jarak penggusuran L2", 100, "m", "Contoh PDF/gambar")
    asum.add("Waktu tetap T3", 0.10, "menit", "Contoh PDF/gambar")
    asum.add("Kapasitas pisau q", "=B13*B14^2", "m3", "q=L x H^2")
    asum.add("Waktu gusur T1", "=B20*60/B18", "menit", "T1=(L2x60)/VF")
    asum.add("Waktu kembali T2", "=B20*60/B19", "menit", "T2=(L2x60)/VR")
    asum.add("Waktu siklus Ts", "=B23+B24+B21", "menit", "Ts=T1+T2+T3")
    asum.add("Kapasitas produksi Q", "=B22*B15*B16*B17*60/B25", "m3/jam", "Q=q x Fb x Fm x Fa x 60/Ts")
    asum.add("Koefisien Dozer", "=1/B26", "jam/m3", "1/Q")
    asum.add()
    asum.add("PARAMETER ALAT DAN BIAYA OPERASI")
    asum.add("Kode alat", "Nama alat", "Harga pokok", "Nilai sisa", "Umur (th)", "Jam/tahun", "Bunga", "Asuransi", "Daya kW", "BBM l/jam", "Harga BBM", "Pelumas/l", "Bengkel", "Perbaikan", "Operator", "Upah/jam", "Produktivitas m2/jam", "Produktivitas m3/jam", "Biaya alat/jam")
    equip = [
        ("AL-GRD", "Motor Grader", 1800000000, .1, 8, 1600, .1, .02, 130, 15, 15000, 18000, .025, .075, 1, 35000, 3000, 90),
        ("AL-DZR", "Bulldozer 155 HP", 2500000000, .1, 8, 1600, .1, .02, 150, 20, 15000, 18000, .025, .075, 1, 35000, 2500, 110),
        ("AL-RLR", "Vibro/Tire Roller", 1400000000, .1, 8, 1600, .1, .02, 100, 12, 15000, 18000, .025, .075, 1, 35000, 2200, 75),
        ("AL-EXC", "Excavator", 2200000000, .1, 8, 1600, .1, .02, 120, 16, 15000, 18000, .025, .075, 1, 35000, 1800, 60),
        ("AL-WTR", "Water Tanker", 900000000, .1, 8, 1600, .1, .02, 100, 10, 15000, 18000, .025, .075, 1, 35000, 2000, 50),
    ]
    for rr, x in enumerate(equip, 31):
        asum.add(*x, f"=((C{rr}-C{rr}*D{rr})*(G{rr}*(1+G{rr})^E{rr}/((1+G{rr})^E{rr}-1))+C{rr}*H{rr})/F{rr}+J{rr}*K{rr}+0.003*I{rr}*L{rr}+M{rr}*C{rr}/F{rr}+N{rr}*C{rr}/F{rr}+O{rr}*P{rr}")
    asum.add()
    asum.add("PARAMETER PRODUKTIVITAS PEKERJAAN - SE-DJBK 68/2024")
    asum.add("Parameter", "Nilai", "Satuan", "Keterangan")
    asum.add("Lebar kerja grader", 3.70, "m", "Lebar efektif lintasan")
    asum.add("Kecepatan kerja grader", 4.00, "km/jam", "Kecepatan kerja efektif")
    asum.add("Efisiensi grader", 0.75, "-", "Faktor efisiensi kerja")
    asum.add("Lintasan roller", 6, "lintasan", "Lintasan pemadatan")
    asum.add("Lebar efektif roller", 2.10, "m", "Lebar drum efektif")
    asum.add("Kecepatan roller", 3.00, "km/jam", "Kecepatan pemadatan")
    asum.add("Efisiensi roller", 0.80, "-", "Faktor efisiensi kerja")
    asum.add("Kapasitas bucket excavator", 0.80, "m3", "Kapasitas bucket")
    asum.add("Faktor isi bucket", 0.85, "-", "Faktor isi material")
    asum.add("Efisiensi excavator", 0.83, "-", "Faktor efisiensi kerja")
    asum.add("Waktu siklus excavator", 0.50, "menit", "Waktu gali-muat-buang")
    asum.add("Koefisien semen semenisasi", 0.35, "zak/m2", "Input spesifikasi pekerjaan")
    asum.add("Koefisien pasir semenisasi", 0.06, "m3/m2", "Input spesifikasi pekerjaan")
    asum.add("Koefisien air semenisasi", 0.015, "m3/m2", "Input spesifikasi pekerjaan")
    asum.widths = {1: 25, 2: 22, 3: 16, 4: 12, 5: 12, 6: 12, 7: 10, 8: 12, 9: 12, 10: 14, 11: 14, 12: 14, 13: 12, 14: 12, 15: 12, 16: 14, 17: 20, 18: 20, 19: 18}

    work_specs = [
        ("GRD", "Grading/perataan", "m2", "AL-GRD", "AL-GRD", "A.3.02", "Perataan/pembentukan badan jalan", "Luas / Q grader"),
        ("RIP", "Ripping/penggemburan", "m3", "AL-DZR", "AL-DZR", "A.3.02", "Penggemburan lapisan existing", "Volume / Q dozer"),
        ("SPR", "Spreading/penyebaran", "m3", "AL-DZR + AL-GRD", "AL-DZR,AL-GRD", "A.3.02", "Penyebaran dan pembentukan material", "Volume / Q alat"),
        ("CMP", "Compacting/pemadatan", "m3", "AL-RLR + AL-WTR", "AL-RLR,AL-WTR", "A.3.03", "Pemadatan dengan pengaturan kadar air", "Volume / Q roller; air / kapasitas tanker"),
        ("DRN", "Drainase/galian", "m3", "AL-EXC", "AL-EXC", "A.3.01.1a", "Galian tanah dan galian batu", "Volume galian / Q excavator"),
        ("SUB", "Perbaikan tanah dasar", "m3", "AL-EXC + AL-DZR + AL-GRD + AL-RLR + AL-WTR", "AL-EXC,AL-DZR,AL-GRD,AL-RLR,AL-WTR", "A.3.01.1a", "Striping Top Soil, tebas-tebang pohon, dan produktivitas excavator", "Volume / Q masing-masing alat"),
        ("BSC", "Basecourse baru", "m3", "AL-DZR + AL-GRD + AL-RLR + AL-WTR", "AL-DZR,AL-GRD,AL-RLR,AL-WTR", "A.3.04", "Penghamparan, pembentukan, pembasahan, dan pemadatan basecourse", "Volume / Q masing-masing alat"),
        ("PTH", "Patching lubang", "m3", "AL-EXC + AL-GRD + AL-RLR + AL-WTR", "AL-EXC,AL-GRD,AL-RLR,AL-WTR", "A.3.05", "Perbaikan lokal lubang dan pemadatan", "Volume / Q masing-masing alat"),
        ("SEM", "Semenisasi", "m2", "MAN", "MAN", "A.3.06", "Pekerjaan manual semenisasi", "Material dan tenaga manual"),
    ]
    work = S("Kode_Pekerjaan")
    work.add("Kode", "Jenis pekerjaan", "Satuan", "Alat/tenaga", "Kode alat yang dihitung", "Kode analisa SE-DJBK", "Uraian acuan", "Dasar produktivitas")
    for row in work_specs:
        work.add(*row)
    work.widths = {1: 12, 2: 28, 3: 10, 4: 48, 5: 38, 6: 24, 7: 58, 8: 30}

    alat = S("Kode_Alat")
    alat.add("Kode alat", "Nama alat", "Keterangan")
    for row in [("AL-DZR", "Bulldozer 155 HP", "Ripping, spreading, basecourse"), ("AL-GRD", "Motor Grader", "Grading"), ("AL-RLR", "Vibro/Tire Roller", "Compacting"), ("AL-EXC", "Excavator", "Drainase dan tanah dasar"), ("AL-WTR", "Water Tanker", "Air pemadatan")]:
        alat.add(*row)
    alat.widths = {1: 15, 2: 24, 3: 42}

    inp = S("Input_Jalan")
    inp.add("No", "Afdeling", "Kondisi", "Titik Kode", "Jenis Jalan", "Blok", "Jenis Kerusakan", "Rekomendasi", "Panjang m", "Lebar m", "Luas m2", "WS", "LM", "CR", "UN", "PH")
    for i, x in enumerate(roads, 2):
        inp.add(x["no"], x["afd"], x["kondisi"], x["kode"], x["jenis"], x["blok"], x["kerusakan"], x["rekomendasi"], x["panjang"], f'=IF(E{i}="JU",Asumsi!$B$3,Asumsi!$B$4)', f"=I{i}*J{i}", pct(x["kerusakan"], "WS"), pct(x["kerusakan"], "LM"), pct(x["kerusakan"], "CR"), pct(x["kerusakan"], "UN"), pct(x["kerusakan"], "PH"))
    inp.widths = {1: 8, 2: 12, 3: 16, 4: 12, 5: 12, 6: 10, 7: 24, 8: 70, 9: 14, 10: 12, 11: 14, 12: 8, 13: 8, 14: 8, 15: 8, 16: 8}

    prod = S("Analisa_Produktivitas")
    prod.add("Kode", "Jenis pekerjaan", "Alat", "Formula/parameter", "Produktivitas", "Satuan", "Koefisien jam", "Satuan koef.", "Kode AHSP", "Halaman acuan", "Uraian analisa")
    prod.add("GRD", "Grading/perataan", "AL-GRD", "Q=L x V x 1.000 x Fa", "=Asumsi!B39*Asumsi!B40*Asumsi!B41", "m2/jam", "=1/E2", "jam/m2", "AHSP-SDA-GRD-PROV", "Lampiran IV; halaman belum terverifikasi", "Perataan/pembentukan badan jalan")
    prod.add("RIP", "Ripping/penggemburan", "AL-DZR", "Q Bulldozer SE-DJBK", "=Asumsi!B26", "m3/jam", "=1/E3", "jam/m3", "AHSP-SDA-RIP-PROV", "Lampiran IV; halaman belum terverifikasi", "Penggemburan lapisan existing")
    prod.add("SPR", "Spreading/penyebaran", "AL-DZR", "Q Bulldozer SE-DJBK", "=Asumsi!B26", "m3/jam", "=1/E4", "jam/m3", "AHSP-SDA-SPR-PROV", "Lampiran IV; halaman belum terverifikasi", "Penyebaran dan pembentukan material")
    prod.add("CMP", "Compacting/pemadatan", "AL-RLR", "Q=L x V x 1.000 x Fa/n", "=Asumsi!B43*Asumsi!B44*Asumsi!B45*1000/Asumsi!B42", "m3/jam", "=1/E5", "jam/m3", "AHSP-SDA-CMP-PROV", "Lampiran IV; halaman belum terverifikasi", "Pemadatan dengan pengaturan kadar air")
    prod.add("DRN", "Drainase/galian", "AL-EXC", "Q=Kb x Fb x Fa x 60/Ts", "=Asumsi!B46*Asumsi!B47*Asumsi!B48*60/Asumsi!B49", "m3/jam", "=1/E6", "jam/m3", "A.3.01.1a", "Sekitar hal. 324; verifikasi terhadap edisi PDF", "Galian tanah dan galian batu")
    prod.add("SUB", "Perbaikan tanah dasar", "AL-EXC", "Q excavator SE-DJBK", "=Asumsi!B46*Asumsi!B47*Asumsi!B48*60/Asumsi!B49", "m3/jam", "=1/E7", "jam/m3", "A.3.01.1a", "Lampiran IV; cocokkan halaman PDF", "Striping Top Soil, tebas-tebang pohon, dan produktivitas excavator")
    prod.add("BSC", "Basecourse baru", "AL-DZR", "Q Bulldozer SE-DJBK", "=Asumsi!B26", "m3/jam", "=1/E8", "jam/m3", "AHSP-SDA-BSC-PROV", "Lampiran IV; halaman belum terverifikasi", "Penghamparan material basecourse")
    prod.add("PTH", "Patching lubang", "AL-EXC", "Q excavator SE-DJBK", "=Asumsi!B46*Asumsi!B47*Asumsi!B48*60/Asumsi!B49", "m3/jam", "=1/E9", "jam/m3", "AHSP-SDA-PTH-PROV", "Lampiran IV; halaman belum terverifikasi", "Perbaikan lokal lubang")
    prod.add("SEM", "Semenisasi", "MAN", "Pekerjaan manual; alat berat tidak dihitung", 0, "m2/jam", 0, "jam/m2", "AHSP-SDA-SEM-PROV", "Lampiran IV; halaman belum terverifikasi", "Pekerjaan manual semenisasi")
    prod.widths = {1: 14, 2: 28, 3: 14, 4: 24, 5: 16, 6: 12, 7: 18, 8: 14, 9: 24, 10: 34, 11: 58}

    ref = S("Referensi_AHSP")
    ref.add("Kode analisa", "Uraian pekerjaan", "Kode AHSP", "Alat/tenaga", "Dasar perhitungan jam alat", "Halaman PDF", "Status")
    for row in [
        ("GRD", "Grading/perataan", "AHSP-SDA-GRD-PROV", "AL-GRD", "Luas / produktivitas grader", "Belum terverifikasi", "Kode internal sementara"),
        ("RIP", "Ripping/penggemburan", "AHSP-SDA-RIP-PROV", "AL-DZR", "Volume / Q dozer", "Belum terverifikasi", "Kode internal sementara"),
        ("SPR", "Spreading/penyebaran", "AHSP-SDA-SPR-PROV", "AL-DZR", "Volume / Q dozer", "Belum terverifikasi", "Kode internal sementara"),
        ("CMP", "Compacting/pemadatan", "AHSP-SDA-CMP-PROV", "AL-RLR + AL-WTR", "Volume / Q roller; air dihitung terpisah", "Belum terverifikasi", "Kode internal sementara"),
        ("DRN", "Drainase/galian", "A.3.01.1a", "AL-EXC", "Volume galian / Q excavator", "Sekitar 324", "Perlu dicocokkan dengan PDF sumber"),
        ("SUB", "Perbaikan tanah dasar / Striping Top Soil", "A.3.01.1a", "AL-EXC + AL-DZR + AL-GRD + AL-RLR + AL-WTR", "Volume / Q excavator, dozer, grader, roller, tanker", "Cocokkan PDF", "Kode analisa mengacu contoh pengguna"),
        ("BSC", "Basecourse baru", "AHSP-SDA-BSC-PROV", "AL-DZR + AL-RLR", "Volume basecourse / produktivitas pekerjaan", "Belum terverifikasi", "Kode internal sementara"),
        ("PTH", "Patching lubang", "AHSP-SDA-PTH-PROV", "AL-EXC", "Volume patching / Q excavator", "Belum terverifikasi", "Kode internal sementara"),
        ("SEM", "Semenisasi", "AHSP-SDA-SEM-PROV", "Tenaga manual", "Tidak ada jam alat berat pada model ini", "Belum terverifikasi", "Kode internal sementara"),
    ]:
        ref.add(*row)
    ref.add()
    ref.add("Catatan penting")
    ref.add("Nomor halaman resmi selain rujukan contoh galian belum diisi karena file PDF SE-DJBK No. 68/2024 Lampiran IV tidak tersedia di workspace saat workbook dibuat.")
    ref.add("Kode AHSP-SDA-*-PROV adalah kode internal analisa, bukan nomor item resmi. Ganti setelah PDF sumber dan edisi halaman dikonfirmasi.")
    ref.widths = {1: 16, 2: 28, 3: 24, 4: 20, 5: 42, 6: 22, 7: 34}

    detail = S("Analisa_Alat_Pekerjaan")
    detail.add("No", "Afdeling", "Blok", "Kode pekerjaan", "Uraian pekerjaan", "Kode AHSP", "Kode alat", "Nama alat", "Volume dihitung", "Produktivitas", "Jam alat", "Biaya alat Rp", "Rumus/trigger")
    tool_names = {"AL-DZR": "Bulldozer 155 HP", "AL-GRD": "Motor Grader", "AL-RLR": "Vibro/Tire Roller", "AL-EXC": "Excavator", "AL-WTR": "Water Tanker", "MAN": "Tenaga manual"}
    prod_rows = {spec[0]: i + 2 for i, spec in enumerate(work_specs)}
    tool_cost_rows = {"AL-GRD": 31, "AL-DZR": 32, "AL-RLR": 33, "AL-EXC": 34, "AL-WTR": 35}
    trigger_map = {
        "GRD": "GRADING",
        "RIP": "RIPPING",
        "SPR": "SPREADING",
        "CMP": "COMPACTING",
        "DRN": "DRAINASE",
        "SUB": "PERBAIKAN TANAH DASAR",
        "BSC": "BASECOURSE BARU",
        "PTH": "PATCHING LUBANG",
        "SEM": "SEMENISASI",
    }
    for input_row in range(2, len(roads) + 2):
        for spec in work_specs:
            code, description, unit, _primary, tools, ahsp, ref_desc, _basis = spec
            for tool_code in tools.split(","):
                tool_code = tool_code.strip()
                if tool_code == "MAN":
                    continue
                trigger = trigger_map[code]
                if code == "DRN":
                    volume = f"=Input_Jalan!I{input_row}*Asumsi!$B$7*Asumsi!$B$8*Input_Jalan!L{input_row}"
                elif code == "GRD":
                    volume = f"=Input_Jalan!K{input_row}"
                else:
                    volume = f"=Input_Jalan!K{input_row}*Asumsi!$B$5*(Input_Jalan!M{input_row}+Input_Jalan!N{input_row}+Input_Jalan!O{input_row})"
                if tool_code == "AL-WTR":
                    productivity = "=Asumsi!$R$35"
                    hours = f'=IF(ISNUMBER(SEARCH("{trigger}",Input_Jalan!H{input_row})),{volume}*Asumsi!$B$9/{productivity},0)'
                elif code == "GRD" or tool_code == "AL-GRD":
                    productivity = "=Analisa_Produktivitas!$E$2"
                    hours = f'=IF(ISNUMBER(SEARCH("{trigger}",Input_Jalan!H{input_row})),{volume}/{productivity},0)'
                elif tool_code == "AL-DZR":
                    productivity = f"=Analisa_Produktivitas!$E${prod_rows['RIP'] if code == 'RIP' else prod_rows['SPR'] if code == 'SPR' else prod_rows['BSC']}"
                    hours = f'=IF(ISNUMBER(SEARCH("{trigger}",Input_Jalan!H{input_row})),{volume}/{productivity},0)'
                elif tool_code == "AL-RLR":
                    productivity = "=Analisa_Produktivitas!$E$5"
                    hours = f'=IF(ISNUMBER(SEARCH("{trigger}",Input_Jalan!H{input_row})),{volume}/{productivity},0)'
                else:
                    productivity = f"=Analisa_Produktivitas!$E${prod_rows[code]}"
                volume_expr = volume[1:] if volume.startswith("=") else volume
                productivity_expr = productivity[1:] if productivity.startswith("=") else productivity
                if tool_code == "AL-WTR":
                    hours = f'=IF(ISNUMBER(SEARCH("{trigger}",Input_Jalan!H{input_row})),{volume_expr}*Asumsi!$B$9/{productivity_expr},0)'
                else:
                    hours = f'=IF(ISNUMBER(SEARCH("{trigger}",Input_Jalan!H{input_row})),{volume_expr}/{productivity_expr},0)'
                cost = f"={hours[1:]}*Asumsi!$S${tool_cost_rows[tool_code]}"
                detail.add(
                    f"=Input_Jalan!A{input_row}", f"=Input_Jalan!B{input_row}", f"=Input_Jalan!F{input_row}",
                    code, description, ahsp, tool_code, tool_names[tool_code], volume, productivity, hours, cost,
                    f"{trigger}; volume / produktivitas {tool_code}",
                )
    detail.widths = {1: 8, 2: 12, 3: 10, 4: 16, 5: 42, 6: 18, 7: 14, 8: 22, 9: 18, 10: 18, 11: 14, 12: 18, 13: 38}
    detail_end = len(detail.rows)

    jam = S("Analisa_Jam_Alat")
    jam.add("No", "Afdeling", "Blok", "Panjang m", "Luas m2", "Vol olah m3", "Vol drainase m3", "Jam GRD", "Jam DZR", "Jam RLR", "Jam EXC", "Jam WTR", "Total jam", "Biaya alat Rp", "Kode pekerjaan terdeteksi", "Kode alat")
    for r in range(2, len(roads) + 2):
        vol = f"=Input_Jalan!K{r}*Asumsi!$B$5*(Input_Jalan!M{r}+Input_Jalan!N{r}+Input_Jalan!O{r})"
        drain = f"=Input_Jalan!I{r}*Asumsi!$B$7*Asumsi!$B$8*Input_Jalan!L{r}"
        grd = f'=SUMIFS(Analisa_Alat_Pekerjaan!$K$2:$K${detail_end},Analisa_Alat_Pekerjaan!$A$2:$A${detail_end},Input_Jalan!A{r},Analisa_Alat_Pekerjaan!$G$2:$G${detail_end},"AL-GRD")'
        dzr = f'=SUMIFS(Analisa_Alat_Pekerjaan!$K$2:$K${detail_end},Analisa_Alat_Pekerjaan!$A$2:$A${detail_end},Input_Jalan!A{r},Analisa_Alat_Pekerjaan!$G$2:$G${detail_end},"AL-DZR")'
        rlr = f'=SUMIFS(Analisa_Alat_Pekerjaan!$K$2:$K${detail_end},Analisa_Alat_Pekerjaan!$A$2:$A${detail_end},Input_Jalan!A{r},Analisa_Alat_Pekerjaan!$G$2:$G${detail_end},"AL-RLR")'
        exc = f'=SUMIFS(Analisa_Alat_Pekerjaan!$K$2:$K${detail_end},Analisa_Alat_Pekerjaan!$A$2:$A${detail_end},Input_Jalan!A{r},Analisa_Alat_Pekerjaan!$G$2:$G${detail_end},"AL-EXC")'
        wtr = f'=SUMIFS(Analisa_Alat_Pekerjaan!$K$2:$K${detail_end},Analisa_Alat_Pekerjaan!$A$2:$A${detail_end},Input_Jalan!A{r},Analisa_Alat_Pekerjaan!$G$2:$G${detail_end},"AL-WTR")'
        flags = f'=IF(ISNUMBER(SEARCH("GRADING",Input_Jalan!H{r})),"GRD ","")&IF(ISNUMBER(SEARCH("RIPPING",Input_Jalan!H{r})),"RIP ","")&IF(ISNUMBER(SEARCH("SPREADING",Input_Jalan!H{r})),"SPR ","")&IF(ISNUMBER(SEARCH("COMPACTING",Input_Jalan!H{r})),"CMP ","")&IF(ISNUMBER(SEARCH("PEMADATAN",Input_Jalan!H{r})),"CMP ","")&IF(ISNUMBER(SEARCH("DRAINASE",Input_Jalan!H{r})),"DRN ","")&IF(ISNUMBER(SEARCH("TANAH DASAR",Input_Jalan!H{r})),"SUB ","")&IF(ISNUMBER(SEARCH("BASECOURSE",Input_Jalan!H{r})),"BSC ","")&IF(ISNUMBER(SEARCH("PATCHING",Input_Jalan!H{r})),"PTH ","")&IF(ISNUMBER(SEARCH("SEMENISASI",Input_Jalan!H{r})),"SEM ","")'
        alat_flags = f'=IF(I{r}>0,"AL-DZR ","")&IF(H{r}>0,"AL-GRD ","")&IF(J{r}>0,"AL-RLR ","")&IF(K{r}>0,"AL-EXC ","")&IF(L{r}>0,"AL-WTR","")'
        jam.add(f"=Input_Jalan!A{r}", f"=Input_Jalan!B{r}", f"=Input_Jalan!F{r}", f"=Input_Jalan!I{r}", f"=Input_Jalan!K{r}", vol, drain, grd, dzr, rlr, exc, wtr, f"=SUM(H{r}:L{r})", f"=H{r}*Asumsi!$S$31+I{r}*Asumsi!$S$32+J{r}*Asumsi!$S$33+K{r}*Asumsi!$S$34+L{r}*Asumsi!$S$35", flags, alat_flags)
    jam.widths = {1: 8, 2: 12, 3: 10, 4: 12, 5: 14, 6: 16, 7: 18, 8: 12, 9: 12, 10: 12, 11: 12, 12: 12, 13: 12, 14: 18, 15: 26, 16: 24}

    mat = S("Kebutuhan_Material")
    mat.add("No", "Afdeling", "Blok", "Panjang m", "Luas m2", "Existing diolah m3", "Basecourse baru m3", "Air pemadatan m3", "Drainase/galian m3", "Semen zak", "Pasir m3", "Air semenisasi m3", "Kode material", "Pekerjaan")
    for r in range(2, len(roads) + 2):
        vol = f"=Input_Jalan!K{r}*Asumsi!$B$5*(Input_Jalan!M{r}+Input_Jalan!N{r}+Input_Jalan!O{r})"
        base = f"=Input_Jalan!K{r}*Asumsi!$B$6*Input_Jalan!O{r}"
        water = f"={vol[1:] if vol.startswith('=') else vol}*Asumsi!$B$9"
        drain = f"=Input_Jalan!I{r}*Asumsi!$B$7*Asumsi!$B$8*Input_Jalan!L{r}"
        cement = f'=IF(ISNUMBER(SEARCH("SEMENISASI",Input_Jalan!H{r})),Input_Jalan!K{r}*Asumsi!$B$50,0)'
        sand = f'=IF(ISNUMBER(SEARCH("SEMENISASI",Input_Jalan!H{r})),Input_Jalan!K{r}*Asumsi!$B$51,0)'
        mat.add(f"=Input_Jalan!A{r}", f"=Input_Jalan!B{r}", f"=Input_Jalan!F{r}", f"=Input_Jalan!I{r}", f"=Input_Jalan!K{r}", vol, base, water, drain, cement, sand, f'=IF(ISNUMBER(SEARCH("SEMENISASI",Input_Jalan!H{r})),Input_Jalan!K{r}*Asumsi!$B$52,0)', "MAT-EXC / MAT-BSC / MAT-WTR / MAT-SEM", f"=Analisa_Jam_Alat!O{r}")
    mat.widths = {1: 8, 2: 12, 3: 10, 4: 12, 5: 14, 6: 22, 7: 22, 8: 20, 9: 22, 10: 14, 11: 14, 12: 20, 13: 30, 14: 28}

    rec = S("Rekap")
    rec.add("REKAPITULASI")
    rec.add("Indikator", "Nilai", "Satuan")
    n = len(roads) + 1
    rec.add("Jumlah ruas", f"=COUNTA(Input_Jalan!A2:A{n})", "ruas")
    rec.add("Total panjang", f"=SUM(Input_Jalan!I2:I{n})", "m")
    rec.add("Total jam alat", f"=SUM(Analisa_Jam_Alat!M2:M{n})", "jam")
    rec.add("Total biaya alat", f"=SUM(Analisa_Jam_Alat!N2:N{n})", "Rp")
    rec.add()
    rec.add("Alat", "Kode", "Jam", "Biaya Rp")
    cost_row = {"AL-GRD": 31, "AL-DZR": 32, "AL-RLR": 33, "AL-EXC": 34, "AL-WTR": 35}
    for name, code, col in [("Motor Grader", "AL-GRD", "H"), ("Bulldozer", "AL-DZR", "I"), ("Roller", "AL-RLR", "J"), ("Excavator", "AL-EXC", "K"), ("Water Tanker", "AL-WTR", "L")]:
        rec.add(name, code, f"=SUM(Analisa_Jam_Alat!{col}2:{col}{n})", f'=SUMPRODUCT(Analisa_Jam_Alat!{col}2:{col}{n},Asumsi!$S${cost_row[code]})')
    rec.add()
    rec.add("Material", "Volume", "Satuan")
    rec.add("Existing diolah", f"=SUM(Kebutuhan_Material!F2:F{n})", "m3")
    rec.add("Basecourse baru", f"=SUM(Kebutuhan_Material!G2:G{n})", "m3")
    rec.add("Air pemadatan", f"=SUM(Kebutuhan_Material!H2:H{n})", "m3")
    rec.add("Drainase/galian", f"=SUM(Kebutuhan_Material!I2:I{n})", "m3")
    rec.add("Semen semenisasi", f"=SUM(Kebutuhan_Material!J2:J{n})", "zak")
    rec.add("Pasir semenisasi", f"=SUM(Kebutuhan_Material!K2:K{n})", "m3")
    rec.add("Air semenisasi", f"=SUM(Kebutuhan_Material!L2:L{n})", "m3")
    rec.widths = {1: 30, 2: 18, 3: 18, 4: 20}

    sheets = [guide, asum, work, alat, inp, prod, detail, jam, mat, rec, ref]
    content = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>''' + "".join(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' for i in range(1, len(sheets)+1)) + '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'
    wb = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>''' + "".join(f'<sheet name="{esc(s.name)}" sheetId="{i}" r:id="rId{i}"/>' for i,s in enumerate(sheets,1)) + '</sheets><calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/></workbook>'
    rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + "".join(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>' for i in range(1,len(sheets)+1)) + '<Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
    styles = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="solid"><fgColor rgb="D9EAF7"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="1" fillId="1" borderId="0"/></cellXfs></styleSheet>'''
    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content)
        z.writestr("xl/workbook.xml", wb)
        z.writestr("xl/_rels/workbook.xml.rels", rels)
        z.writestr("xl/styles.xml", styles)
        z.writestr("_rels/.rels", '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        for i, s in enumerate(sheets, 1):
            z.writestr(f"xl/worksheets/sheet{i}.xml", sheet_xml(s))
    print(f"Created {OUT} with {len(roads)} analyzed roads and {len(sheets)} linked sheets")


if __name__ == "__main__":
    build()
