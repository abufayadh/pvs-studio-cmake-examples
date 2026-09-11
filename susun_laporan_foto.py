"""Build a block-ordered road damage photo report from Foto 1KCO.xlsx."""

from __future__ import annotations

import hashlib
import os
import re
import tempfile
from collections import Counter, defaultdict
from io import BytesIO
from pathlib import Path

import openpyxl
from PIL import Image, ImageOps

# Field photos are supplied as embedded workbook images; allow PIL to inspect
# their metadata before the report generator applies its output-size bound.
Image.MAX_IMAGE_PIXELS = None
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    Image as PdfImage,
    KeepTogether,
    PageBreak,
    PageTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)


WORKBOOK = Path(
    r"C:\Users\Abu Fazimaroyatih\.copilot\workspaces\14c79664-cd30-40da-9046-fb00ac9f385c\attachments\fca5324f-1245-4d39-b5d8-557377355db7-Foto 1KCO.xlsx"
)
OUTPUT = Path("Laporan_Analisa_Foto_1KCO.pdf")

NAVY = colors.HexColor("#12344D")
BLUE = colors.HexColor("#1F6F8B")
TEAL = colors.HexColor("#2C9C95")
ORANGE = colors.HexColor("#F28E2B")
LIGHT = colors.HexColor("#F2F6F8")
MID = colors.HexColor("#D8E4EA")
DARK = colors.HexColor("#243746")


def clean(value) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def block_key(value: str) -> tuple[str, int]:
    matches = re.findall(r"([A-E])\s*(\d+)(?!\d)", value.upper())
    if not matches:
        return ("Z", 9999)
    letter, number = matches[-1]
    return (letter, int(number))


def class_from_point(name: str) -> str:
    value = name.upper()
    if "SPRB" in value or "RUSAK BERAT" in value:
        return "Rusak berat"
    if "SPRS" in value or "RUSAK SEDANG" in value:
        return "Rusak sedang"
    return "Kondisi khusus"


def make_image(raw: bytes, width: float, height: float) -> PdfImage:
    with Image.open(BytesIO(raw)) as source:
        image = ImageOps.exif_transpose(source).convert("RGB")
        max_w, max_h = max(1, int(width)), max(1, int(height))
        image.thumbnail((max_w, max_h), Image.Resampling.LANCZOS)
        canvas = Image.new("RGB", (max_w, max_h), "white")
        x = (max_w - image.width) // 2
        y = (max_h - image.height) // 2
        canvas.paste(image, (x, y))
        output = BytesIO()
        canvas.save(output, format="JPEG", quality=88)
        output.seek(0)
        with Image.open(output) as checked:
            if checked.width * checked.height > 20_000_000:
                raise ValueError(f"Unexpected rendered image size: {checked.size}")
        output.seek(0)
    result = PdfImage(output, width=width, height=height)
    result.hAlign = "CENTER"
    return result


def load_data():
    workbook = openpyxl.load_workbook(WORKBOOK, data_only=True, read_only=False)
    pairs = workbook["Titik + Foto (2)"]
    detail = workbook["Titik + Foto"]

    detail_by_block = defaultdict(list)
    for row in detail.iter_rows(min_row=2, values_only=True):
        afd, point, photo, description, classification = (
            clean(row[1]),
            clean(row[2]),
            clean(row[6]),
            clean(row[8]),
            clean(row[9]),
        )
        if not afd:
            continue
        key = block_key(point)
        if key[0] != "Z":
            detail_by_block[(afd, key)].append(
                {
                    "classification": classification or "Tidak diklasifikasikan",
                    "description": description,
                    "point": point,
                    "photo": photo,
                }
            )

    anchor_by_row = defaultdict(list)
    for image in pairs._images:
        anchor = image.anchor._from
        side = "Start Point" if anchor.col == 4 else "End Point" if anchor.col == 9 else "Foto"
        anchor_by_row[anchor.row + 1].append((side, image._data()))

    records = []
    for row_number in range(9, pairs.max_row + 1):
        afd = clean(pairs.cell(row_number, 2).value)
        if not afd or not isinstance(pairs.cell(row_number, 1).value, int):
            continue
        start = clean(pairs.cell(row_number, 3).value)
        end = clean(pairs.cell(row_number, 9).value)
        block_text = f"{start} {end}"
        key = block_key(block_text)
        if key[0] == "Z":
            continue
        details = detail_by_block.get((afd, key), [])
        classifications = Counter(
            d["classification"] for d in details if d["classification"]
        )
        if classifications:
            classification = classifications.most_common(1)[0][0]
        else:
            classification = (
                "Rusak berat" if "BERAT" in block_text.upper() else "Rusak sedang"
            )
        description = next(
            (d["description"] for d in details if d["description"]), ""
        )
        evidence = "; ".join(
            sorted({d["classification"] for d in details if d["classification"]})
        )
        images = {"Start Point": None, "End Point": None}
        for side, raw in anchor_by_row.get(row_number, []):
            if images.get(side) is None:
                images[side] = raw
        records.append(
            {
                "afd": afd,
                "block": f"{key[0]}{key[1]}",
                "key": key,
                "start": start or "Data Start Point tidak tersedia",
                "end": end or "Data End Point tidak tersedia",
                "classification": classification,
                "evidence": evidence or classification,
                "description": description
                or "Kerusakan ditentukan dari keterangan titik dan klasifikasi pada workbook.",
                "images": images,
            }
        )
    records.sort(key=lambda item: (item["afd"], item["key"]))
    return records


def report_description(record) -> str:
    classification = record["classification"]
    lower = classification.lower()
    actions = {
        "wet surface": "Prioritaskan pembersihan saluran dan pengaliran genangan; lakukan pemadatan setelah kadar air terkendali.",
        "loose material": "Lakukan grading, penambahan material granular yang sesuai, pengaturan kadar air, dan pemadatan.",
        "pot hole": "Gali material gagal, isi kembali dengan material bergradasi, ratakan, lalu padatkan.",
        "undulating": "Lakukan perataan profil, perbaiki lapisan dasar/subgrade, dan pastikan kemiringan melintang berfungsi.",
        "corrugation": "Lakukan ripping/grading pada kedalaman yang diperlukan, kemudian padatkan secara merata.",
        "upheaval": "Evaluasi tanah dasar dan drainase; ratakan atau gali material yang mengembang sebelum pengisian kembali.",
    }
    for key, action in actions.items():
        if key in lower:
            return action
    if "berat" in lower:
        return "Kerusakan berat memerlukan verifikasi kedalaman, perbaikan drainase, penggantian material gagal, dan pemadatan bertahap."
    return "Lakukan verifikasi lapangan, perbaikan drainase, perataan, dan pemadatan sesuai hasil pengukuran."


def footer(canvas, doc):
    canvas.saveState()
    canvas.setFillColor(NAVY)
    canvas.rect(0, 0, landscape(A4)[0], 0.45 * cm, fill=1, stroke=0)
    canvas.setFillColor(colors.white)
    canvas.setFont("Helvetica", 7)
    canvas.drawString(1.2 * cm, 0.16 * cm, "Laporan Analisa Foto Kondisi Jalan - 1KCO")
    canvas.drawRightString(
        landscape(A4)[0] - 1.2 * cm, 0.16 * cm, f"Halaman {doc.page}"
    )
    canvas.restoreState()


def build_report(records):
    page_w, page_h = landscape(A4)
    styles = getSampleStyleSheet()
    styles.add(
        ParagraphStyle(
            name="PosterTitle",
            parent=styles["Title"],
            fontName="Helvetica-Bold",
            fontSize=24,
            leading=27,
            textColor=colors.white,
            alignment=1,
            spaceAfter=5,
        )
    )
    styles.add(
        ParagraphStyle(
            name="PosterSub",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=10,
            leading=13,
            textColor=colors.white,
            alignment=1,
        )
    )
    styles.add(
        ParagraphStyle(
            name="Section",
            parent=styles["Heading2"],
            fontName="Helvetica-Bold",
            fontSize=13,
            leading=15,
            textColor=NAVY,
            spaceBefore=5,
            spaceAfter=5,
        )
    )
    styles.add(
        ParagraphStyle(
            name="Small",
            parent=styles["Normal"],
            fontSize=8.4,
            leading=10.5,
            textColor=DARK,
        )
    )
    styles.add(
        ParagraphStyle(
            name="SmallWhite",
            parent=styles["Small"],
            textColor=colors.white,
        )
    )
    styles.add(
        ParagraphStyle(
            name="PhotoLabel",
            parent=styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=10,
            leading=12,
            textColor=NAVY,
            alignment=1,
        )
    )

    doc = BaseDocTemplate(
        str(OUTPUT),
        pagesize=landscape(A4),
        leftMargin=1.1 * cm,
        rightMargin=1.1 * cm,
        topMargin=1.0 * cm,
        bottomMargin=0.85 * cm,
        title="Laporan Analisa Foto 1KCO",
        author="OpenAI",
    )
    frame = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id="normal")
    doc.addPageTemplates([PageTemplate(id="main", frames=frame, onPage=footer)])

    story = []
    class_counts = Counter(r["classification"] for r in records)
    afd_counts = Counter(r["afd"] for r in records)

    header = Table(
        [
            [
                Paragraph("ANALISA FOTO KONDISI JALAN", styles["PosterTitle"]),
            ],
            [
                Paragraph(
                    "1KCO - PEMETAAN BLOK DAN PASANGAN START POINT / END POINT",
                    styles["PosterSub"],
                )
            ],
            [
                Paragraph(
                    "Dokumentasi diurutkan berdasarkan AFD dan kode blok, dimulai dari A1. "
                    "Foto yang tidak tersedia ditandai secara eksplisit.",
                    styles["PosterSub"],
                )
            ],
        ],
        colWidths=[doc.width],
        rowHeights=[1.05 * cm, 0.55 * cm, 0.85 * cm],
    )
    header.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), NAVY),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("BOX", (0, 0), (-1, -1), 0.5, NAVY),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("RIGHTPADDING", (0, 0), (-1, -1), 8),
            ]
        )
    )
    story.extend([header, Spacer(1, 0.35 * cm)])

    summary = [
        [
            Paragraph("<b>TOTAL BLOK</b><br/><font size=20>%d</font>" % len(records), styles["Small"]),
            Paragraph("<b>AFD TERDATA</b><br/><font size=20>%d</font>" % len(afd_counts), styles["Small"]),
            Paragraph("<b>FOTO START/END</b><br/><font size=20>%d</font>" % sum(sum(v is not None for v in r["images"].values()) for r in records), styles["Small"]),
            Paragraph("<b>KLASIFIKASI</b><br/><font size=20>%d</font>" % len(class_counts), styles["Small"]),
        ]
    ]
    summary_table = Table(summary, colWidths=[doc.width / 4] * 4, rowHeights=[1.3 * cm])
    summary_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), LIGHT),
                ("GRID", (0, 0), (-1, -1), 0.5, MID),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ]
        )
    )
    story.append(summary_table)
    story.append(Spacer(1, 0.35 * cm))

    summary_rows = [["Klasifikasi kerusakan", "Jumlah blok", "AFD"]]
    for label, count in class_counts.most_common():
        summary_rows.append(
            [
                label,
                str(count),
                ", ".join(sorted({r["afd"] for r in records if r["classification"] == label})),
            ]
        )
    summary_table = Table(summary_rows, colWidths=[doc.width * 0.54, doc.width * 0.16, doc.width * 0.30])
    summary_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), BLUE),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 8.5),
                ("GRID", (0, 0), (-1, -1), 0.35, MID),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, LIGHT]),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ("RIGHTPADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )
    story.extend(
        [
            Paragraph("RINGKASAN HASIL ANALISA", styles["Section"]),
            summary_table,
            Spacer(1, 0.35 * cm),
            Paragraph(
                "<b>METODE PEMBACAAN:</b> klasifikasi mengikuti kolom Klasifikasi Kondisi Jalan "
                "pada workbook dan dikaitkan dengan blok melalui nama titik. Keterangan perbaikan "
                "merupakan rekomendasi awal untuk verifikasi lapangan; ukuran, kedalaman, dan volume "
                "tidak diasumsikan dari foto.",
                styles["Small"],
            ),
            PageBreak(),
        ]
    )

    current_afd = None
    for index, record in enumerate(records, 1):
        if current_afd != record["afd"]:
            current_afd = record["afd"]
            band = Table(
                [[Paragraph(f"AFD {current_afd} - URUTAN BLOK", styles["PosterTitle"])]],
                colWidths=[doc.width],
                rowHeights=[0.8 * cm],
            )
            band.setStyle(
                TableStyle(
                    [
                        ("BACKGROUND", (0, 0), (-1, -1), TEAL),
                        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ]
                )
            )
            story.append(band)
            story.append(Spacer(1, 0.15 * cm))

        title = Table(
            [
                [
                    Paragraph(f"BLOK {record['block']}", styles["Section"]),
                    Paragraph(
                        f"<b>{record['classification']}</b><br/>Dokumentasi ke-{index} dari {len(records)}",
                        styles["Small"],
                    ),
                ]
            ],
            colWidths=[doc.width * 0.62, doc.width * 0.38],
        )
        title.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, -1), LIGHT),
                    ("BOX", (0, 0), (-1, -1), 0.6, MID),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 7),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 7),
                ]
            )
        )
        story.append(title)
        story.append(Spacer(1, 0.12 * cm))

        photo_w = (doc.width - 0.3 * cm) / 2
        photo_h = 8.7 * cm
        photo_cells = []
        for side in ("Start Point", "End Point"):
            raw = record["images"][side]
            image_content = (
                make_image(raw, photo_w - 0.25 * cm, photo_h)
                if raw
                else Paragraph("FOTO TIDAK TERSEDIA", styles["Small"])
            )
            photo_cells.append(
                [
                    Paragraph(side.upper(), styles["PhotoLabel"]),
                    image_content,
                    Paragraph(
                        (record["start"] if side == "Start Point" else record["end"]).replace(
                            "&", "&amp;"
                        ),
                        styles["Small"],
                    ),
                ]
            )
        photos = Table(
            [[photo_cells[0], photo_cells[1]]],
            colWidths=[photo_w, photo_w],
            rowHeights=[photo_h + 1.35 * cm],
        )
        photos.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, -1), colors.white),
                    ("BOX", (0, 0), (-1, -1), 0.6, MID),
                    ("INNERGRID", (0, 0), (-1, -1), 0.6, MID),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 5),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                    ("TOPPADDING", (0, 0), (-1, -1), 4),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ]
            )
        )
        story.append(photos)
        story.append(Spacer(1, 0.12 * cm))

        analysis = Table(
            [
                [
                    Paragraph("<b>ANALISA JENIS KERUSAKAN</b>", styles["SmallWhite"]),
                    Paragraph("<b>REKOMENDASI PENANGANAN AWAL</b>", styles["SmallWhite"]),
                ],
                [
                    Paragraph(
                        f"<b>Klasifikasi:</b> {record['classification']}<br/>"
                        f"<b>Indikasi terkait blok:</b> {record['evidence']}<br/>"
                        f"<b>Keterangan:</b> {record['description']}",
                        styles["Small"],
                    ),
                    Paragraph(report_description(record), styles["Small"]),
                ],
            ],
            colWidths=[doc.width * 0.5, doc.width * 0.5],
        )
        analysis.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), NAVY),
                    ("BACKGROUND", (0, 1), (-1, 1), LIGHT),
                    ("GRID", (0, 0), (-1, -1), 0.45, MID),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 6),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                    ("TOPPADDING", (0, 0), (-1, -1), 5),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                ]
            )
        )
        story.append(analysis)
        if index != len(records):
            story.append(PageBreak())

    doc.build(story)


if __name__ == "__main__":
    data = load_data()
    build_report(data)
    print(f"Created {OUTPUT} with {len(data)} block records.")
