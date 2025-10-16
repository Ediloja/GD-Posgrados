from pathlib import Path
from weasyprint import HTML, CSS

BASE = Path(__file__).parent.resolve()

sis_code = input("\nIngrese el código SIS para generar el PDF:\n\n>>> ")

# Rutas de los archivos
html_path = BASE / "finales" / f'{sis_code}.html'
css_path  = BASE / "assets" / "style.css"

output = BASE / "finales" / f'{sis_code}.pdf'

print("\nRenderizando HTML a PDF...")

HTML(filename=str(html_path), base_url=str(html_path.parent)).write_pdf(target=str(output), stylesheets=[CSS(filename=str(css_path))])

print(f'\nPDF generado ({sis_code}.html)\n')