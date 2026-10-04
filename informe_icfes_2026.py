from copy import copy
import json
import re
from pathlib import Path
from numbers import Real
from xml.sax.saxutils import escape

from openpyxl import load_workbook
from openpyxl.worksheet.table import TableColumn
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


BASE_DIR = Path(__file__).resolve().parent
INPUT_CANDIDATES = (
    BASE_DIR / "Informe_ICFES_2026.xlsx",
    BASE_DIR / "Informe_ICFES_2006.xlsx",
)
SHEET_NAME = "BD_Consolidado"
SCORE_HEADER = "Ponderado ICFES"
POSITION_HEADER = "Posici\u00f3n (podium)"
INSTITUTION = "Instituci\u00f3n Educativa Manuel Rodr\u00edguez Torices de San Diego"
SUBJECT_WEIGHTS = {
    "Matematicas": 3,
    "Ciencias": 3,
    "Lectura": 3,
    "Sociales": 3,
    "Ingles": 1,
}
NORMALIZATION_FACTOR = 1 / 13
JSON_OUTPUT = BASE_DIR / "informe_icfes_2026_datos.json"
PDF_OUTPUT = BASE_DIR / "informe_icfes_2026_ponderado_y_posicion.pdf"


def find_source_file():
    for path in INPUT_CANDIDATES:
        if path.is_file():
            return path
    names = ", ".join(path.name for path in INPUT_CANDIDATES)
    raise FileNotFoundError(f"No se encontro el archivo fuente. Se esperaba: {names}")


def calculate_weighted_score(ws, row):
    subject_scores = {
        "Matematicas": ws.cell(row, 6).value,
        "Ciencias": ws.cell(row, 8).value,
        "Lectura": ws.cell(row, 5).value,
        "Sociales": ws.cell(row, 7).value,
        "Ingles": ws.cell(row, 9).value,
    }
    if any(not isinstance(value, Real) for value in subject_scores.values()):
        raise ValueError(f"Puntajes incompletos o no numericos en la fila {row}.")

    weighted_sum = sum(
        subject_scores[subject] * SUBJECT_WEIGHTS[subject]
        for subject in SUBJECT_WEIGHTS
    )
    return round(weighted_sum * NORMALIZATION_FACTOR, 2)


def load_historical_averages():
    workbook = load_workbook(
        BASE_DIR / "ANTERIORES ICFES.xlsx",
        read_only=True,
        data_only=True,
    )
    worksheet = workbook.active
    area_keys = ("reading", "math", "social", "science", "english")
    historical = {}
    for values in worksheet.iter_rows(min_row=3, values_only=True):
        label = values[0]
        match = re.search(r"\b(20\d{2})\b", str(label or ""))
        scores = values[1:6]
        if not match or len(scores) != len(area_keys):
            continue
        if any(not isinstance(score, Real) for score in scores):
            raise ValueError(f"Promedios historicos incompletos para {match.group(1)}.")
        historical[match.group(1)] = {
            key: round(float(score), 2)
            for key, score in zip(area_keys, scores)
        }
    if not historical:
        raise ValueError("No se encontraron promedios por area en ANTERIORES ICFES.xlsx.")
    return dict(sorted(historical.items(), key=lambda item: int(item[0]), reverse=True))


def build_report_data(ws, rows, weighted_scores, rank_by_score):
    students = []
    for row, weighted in zip(rows, weighted_scores):
        students.append(
            {
                "name": ws.cell(row, 1).value,
                "global": ws.cell(row, 4).value,
                "reading": ws.cell(row, 5).value,
                "math": ws.cell(row, 6).value,
                "social": ws.cell(row, 7).value,
                "science": ws.cell(row, 8).value,
                "english": ws.cell(row, 9).value,
                "weighted": weighted,
                "rank": rank_by_score[weighted],
            }
        )
    students.sort(key=lambda student: (student["rank"], student["name"]))

    def average(key):
        return round(sum(student[key] for student in students) / len(students), 2)

    area_definitions = (
        ("reading", "Lectura cr\u00edtica"),
        ("math", "Matem\u00e1ticas"),
        ("social", "Sociales y ciudadanas"),
        ("science", "Ciencias naturales"),
        ("english", "Ingl\u00e9s"),
    )
    return {
        "institution": INSTITUTION,
        "municipality": "San Diego",
        "department": "Cesar",
        "country": "Colombia",
        "year": 2026,
        "exam": "Saber 11",
        "summary": {
            "students": len(students),
            "globalAverage": average("global"),
            "weightedAverage": average("weighted"),
            "highestGlobal": max(student["global"] for student in students),
            "lowestGlobal": min(student["global"] for student in students),
        },
        "areas": [
            {"key": key, "label": label, "average": average(key)}
            for key, label in area_definitions
        ],
        "historicalAverages": load_historical_averages(),
        "students": students,
    }


def write_pdf(report_data, output_path):
    forest = colors.HexColor("#173d35")
    leaf = colors.HexColor("#397b63")
    coral = colors.HexColor("#e66c54")
    line = colors.HexColor("#d9e2dd")
    pale = colors.HexColor("#f1f5f1")
    muted = colors.HexColor("#66746e")
    page_width, _ = landscape(A4)
    styles = getSampleStyleSheet()
    styles.add(
        ParagraphStyle(
            name="ReportTitle",
            parent=styles["Title"],
            fontName="Times-Bold",
            fontSize=22,
            leading=26,
            textColor=forest,
            alignment=0,
            spaceAfter=2,
        )
    )
    styles.add(
        ParagraphStyle(
            name="ReportMeta",
            parent=styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=8,
            leading=11,
            textColor=leaf,
            spaceAfter=8,
        )
    )
    styles.add(
        ParagraphStyle(
            name="ReportSection",
            parent=styles["Heading2"],
            fontName="Times-Bold",
            fontSize=14,
            leading=17,
            textColor=forest,
            spaceBefore=9,
            spaceAfter=6,
        )
    )
    styles.add(
        ParagraphStyle(
            name="ReportCell",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=7.2,
            leading=8.5,
            textColor=colors.HexColor("#182622"),
        )
    )
    styles.add(
        ParagraphStyle(
            name="ReportCellBold",
            parent=styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=7.2,
            leading=8.5,
            textColor=forest,
        )
    )
    styles.add(
        ParagraphStyle(
            name="ReportHead",
            parent=styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=6.5,
            leading=7.5,
            textColor=colors.white,
        )
    )
    styles.add(
        ParagraphStyle(
            name="ReportHeadRight",
            parent=styles["ReportHead"],
            alignment=2,
        )
    )
    styles.add(
        ParagraphStyle(
            name="ReportFootnote",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=7,
            leading=9,
            textColor=muted,
        )
    )

    doc = SimpleDocTemplate(
        str(output_path),
        pagesize=landscape(A4),
        leftMargin=12 * mm,
        rightMargin=12 * mm,
        topMargin=12 * mm,
        bottomMargin=14 * mm,
        title=f"Informe Saber 11 {report_data['year']} - {report_data['institution']}",
        author=report_data["institution"],
    )
    story = [
        Paragraph(escape(report_data["institution"]), styles["ReportTitle"]),
        Paragraph(
            f"SAN DIEGO, CESAR, COLOMBIA &nbsp; / &nbsp; SABER 11 {report_data['year']} &nbsp; / &nbsp; INFORME CONSOLIDADO",
            styles["ReportMeta"],
        ),
    ]

    summary = report_data["summary"]
    metric_labels = (
        ("ESTUDIANTES", str(summary["students"])),
        ("PROMEDIO GLOBAL", f"{summary['globalAverage']:.2f}"),
        ("PROMEDIO PONDERADO", f"{summary['weightedAverage']:.2f}"),
        ("GLOBAL M\u00c1XIMO", str(summary["highestGlobal"])),
        ("GLOBAL M\u00cdNIMO", str(summary["lowestGlobal"])),
    )
    metric_cells = [
        Paragraph(
            f"<font name='Helvetica-Bold' size='7' color='#66746e'>{label}</font><br/>"
            f"<font name='Times-Bold' size='17' color='#173d35'>{value}</font>",
            styles["Normal"],
        )
        for label, value in metric_labels
    ]
    metric_table = Table(
        [metric_cells], colWidths=[(page_width - 24 * mm) / 5] * 5
    )
    metric_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), pale),
                ("BOX", (0, 0), (-1, -1), 0.5, line),
                ("INNERGRID", (0, 0), (-1, -1), 0.5, line),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )
    story.extend([metric_table, Paragraph("Promedio por \u00e1rea", styles["ReportSection"])])

    area_rows = [
        [Paragraph("PRUEBA", styles["ReportHead"]), Paragraph("PROMEDIO / 100", styles["ReportHeadRight"]) ]
    ]
    for area in report_data["areas"]:
        area_rows.append(
            [
                Paragraph(escape(area["label"]), styles["ReportCell"]),
                Paragraph(f"{area['average']:.2f}", styles["ReportCellBold"]),
            ]
        )
    area_table = Table(area_rows, colWidths=[100 * mm, 35 * mm], repeatRows=1)
    area_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), leaf),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, pale]),
                ("LINEBELOW", (0, 0), (-1, -1), 0.35, line),
                ("ALIGN", (1, 0), (1, -1), "RIGHT"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    story.extend(
        [
            area_table,
            Paragraph("Ranking individual", styles["ReportSection"]),
            Paragraph(
                "Ponderado = (Matem\u00e1ticas \u00d7 3 + Ciencias naturales \u00d7 3 + "
                "Lectura cr\u00edtica \u00d7 3 + Sociales \u00d7 3 + Ingl\u00e9s \u00d7 1) \u00f7 13. "
                "Se consideran dos decimales; los empates comparten rank y la secuencia contin\u00faa sin saltos.",
                styles["ReportFootnote"],
            ),
            Spacer(1, 5),
        ]
    )

    headers = (
        "#",
        "RANK",
        "ESTUDIANTE",
        "PONDERADO",
        "GLOBAL",
        "LECTURA",
        "MATEM\u00c1TICAS",
        "SOCIALES",
        "CIENCIAS",
        "INGL\u00c9S",
    )
    student_rows = [
        [
            Paragraph(label, styles["ReportHeadRight"] if index > 1 else styles["ReportHead"])
            for index, label in enumerate(headers)
        ]
    ]
    for row_number, student in enumerate(report_data["students"], start=1):
        student_rows.append(
            [
                Paragraph(str(row_number), styles["ReportCell"]),
                Paragraph(str(student["rank"]), styles["ReportCellBold"]),
                Paragraph(escape(student["name"]), styles["ReportCell"]),
                Paragraph(f"{student['weighted']:.2f}", styles["ReportCellBold"]),
                *[
                    Paragraph(str(student[key]), styles["ReportCell"])
                    for key in ("global", "reading", "math", "social", "science", "english")
                ],
            ]
        )
    result_table = Table(
        student_rows,
        colWidths=[8 * mm, 12 * mm, 70 * mm, 24 * mm, 21 * mm, 21 * mm, 25 * mm, 23 * mm, 23 * mm, 20 * mm],
        repeatRows=1,
    )
    result_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), forest),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, pale]),
                ("LINEBELOW", (0, 0), (-1, -1), 0.3, line),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (3, 1), (-1, -1), "RIGHT"),
                ("TEXTCOLOR", (3, 1), (3, -1), coral),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]
        )
    )
    story.append(result_table)

    def draw_footer(canvas, document):
        canvas.saveState()
        canvas.setStrokeColor(line)
        canvas.line(12 * mm, 9 * mm, page_width - 12 * mm, 9 * mm)
        canvas.setFont("Helvetica", 7)
        canvas.setFillColor(muted)
        canvas.drawString(
            12 * mm,
            5 * mm,
            f"{report_data['institution']} - Saber 11 {report_data['year']}",
        )
        canvas.drawRightString(page_width - 12 * mm, 5 * mm, f"P\u00e1gina {document.page}")
        canvas.restoreState()

    doc.build(story, onFirstPage=draw_footer, onLaterPages=draw_footer)


def main():
    source_path = find_source_file()
    output_path = source_path.with_name(
        f"{source_path.stem}_ponderado_y_posicion.xlsx"
    )

    workbook = load_workbook(source_path)
    if SHEET_NAME not in workbook.sheetnames:
        raise ValueError(f"No se encontro la hoja '{SHEET_NAME}'.")

    ws = workbook[SHEET_NAME]
    expected_headers = [
        "Nombre",
        "ID",
        "Registro",
        "Global",
        "Lectura",
        "Matematicas",
        "Sociales",
        "Ciencias",
        "Ingles",
    ]
    actual_headers = [ws.cell(1, column).value for column in range(1, 10)]
    if actual_headers != expected_headers:
        raise ValueError(
            "Los encabezados de BD_Consolidado no coinciden con el formato esperado."
        )
    if ws["K1"].value is not None or ws["L1"].value is not None:
        raise ValueError("Las columnas K y L ya contienen datos en el archivo fuente.")

    data_rows = range(2, ws.max_row + 1)
    weighted_scores = [calculate_weighted_score(ws, row) for row in data_rows]
    distinct_scores = sorted(set(weighted_scores), reverse=True)
    rank_by_score = {
        score: position for position, score in enumerate(distinct_scores, start=1)
    }

    header_style = copy(ws["J1"]._style)
    data_style = copy(ws["J2"]._style)
    headers = (SCORE_HEADER, POSITION_HEADER)
    for column, title in ((11, headers[0]), (12, headers[1])):
        cell = ws.cell(row=1, column=column, value=title)
        cell._style = copy(header_style)
        cell.alignment = copy(ws["J1"].alignment)

    for row, score in zip(data_rows, weighted_scores):
        position = rank_by_score[score]
        for column, value in ((11, score), (12, position)):
            cell = ws.cell(row=row, column=column, value=value)
            cell._style = copy(data_style)
            cell.alignment = copy(ws["J2"].alignment)
            if column == 11:
                cell.number_format = "0.00"

    ws.column_dimensions["K"].width = 18
    ws.column_dimensions["L"].width = 20

    if ws.tables:
        table = next(iter(ws.tables.values()))
        table.ref = f"A1:L{ws.max_row}"
        if table.autoFilter:
            table.autoFilter.ref = table.ref
        table.tableColumns.append(TableColumn(id=11, name=SCORE_HEADER))
        table.tableColumns.append(TableColumn(id=12, name=POSITION_HEADER))

    try:
        workbook.save(output_path)
    except PermissionError:
        output_path = source_path.with_name(
            f"{source_path.stem}_ponderado_y_posicion_actualizado.xlsx"
        )
        workbook.save(output_path)

    report_data = build_report_data(ws, data_rows, weighted_scores, rank_by_score)
    JSON_OUTPUT.write_text(
        json.dumps(report_data, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    write_pdf(report_data, PDF_OUTPUT)

    print(f"Archivo creado: {output_path}")
    print(f"Datos de interfaz creados: {JSON_OUTPUT}")
    print(f"PDF de informe creado: {PDF_OUTPUT}")
    print(f"Estudiantes procesados: {len(weighted_scores)}")
    print("Puesto 1 corresponde al puntaje mas alto; los empates comparten puesto.")


if __name__ == "__main__":
    main()