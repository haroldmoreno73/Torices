import base64
import json
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "publicacion_web"
HTML_SOURCE = BASE_DIR / "informe_icfes_2026.html"
DATA_SOURCE = BASE_DIR / "informe_icfes_2026_datos.json"
PDF_SOURCE = BASE_DIR / "informe_icfes_2026_ponderado_y_posicion.pdf"
CREST_SOURCE = BASE_DIR / "escudo_Torices.jpg"
PROMO_IMAGE_SOURCE = BASE_DIR / "resultado_saber11_orgullo_torices.png"
HTML_OUTPUT = OUTPUT_DIR / "index.html"


def main():
    report_data = json.loads(DATA_SOURCE.read_text(encoding="utf-8"))
    inline_json = json.dumps(report_data, ensure_ascii=False).replace("<", "\\u003c")
    inline_pdf = base64.b64encode(PDF_SOURCE.read_bytes()).decode("ascii")
    inline_crest = base64.b64encode(CREST_SOURCE.read_bytes()).decode("ascii")
    html = HTML_SOURCE.read_text(encoding="utf-8")
    crest_src = 'src="./escudo_Torices.jpg"'
    if html.count(crest_src) != 1:
        raise ValueError("No se encontro la imagen unica del escudo institucional.")
    html = html.replace(crest_src, f'src="data:image/jpeg;base64,{inline_crest}"', 1)
    script_marker = "  <script>\n    const numberFormat"
    if html.count(script_marker) != 1:
        raise ValueError("No se encontro el bloque principal del informe HTML.")
    html = html.replace(
        script_marker,
        f'  <script id="embedded-report-data" type="application/json">{inline_json}</script>\n'
        f'  <script>\n    const numberFormat',
        1,
    )
    pdf_link = 'href="./informe_icfes_2026_ponderado_y_posicion.pdf"'
    if html.count(pdf_link) != 1:
        raise ValueError("No se encontro el enlace unico al informe PDF.")
    html = html.replace(
        pdf_link,
        f'href="data:application/pdf;base64,{inline_pdf}"',
        1,
    )

    OUTPUT_DIR.mkdir(exist_ok=True)
    for private_asset in (
        OUTPUT_DIR / "informe_icfes_2026.html",
        OUTPUT_DIR / "informe_icfes_2026_datos.json",
        OUTPUT_DIR / "informe_icfes_2026_ponderado_y_posicion.pdf",
    ):
        private_asset.unlink(missing_ok=True)
    HTML_OUTPUT.write_text(html, encoding="utf-8")
    if PROMO_IMAGE_SOURCE.exists():
        (OUTPUT_DIR / PROMO_IMAGE_SOURCE.name).write_bytes(PROMO_IMAGE_SOURCE.read_bytes())
    print(f"Interfaz publica creada: {HTML_OUTPUT}")
    print(f"Estudiantes incluidos: {report_data['summary']['students']}")
    print("Archivos publicados: 1 HTML independiente")


if __name__ == "__main__":
    main()