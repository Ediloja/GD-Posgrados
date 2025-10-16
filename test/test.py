from pathlib import Path
from weasyprint import HTML, CSS

BASE = Path(__file__).parent.resolve()
# Rutas de los archivos
html_path = BASE / 'index.html'
css_path  = BASE / 'style.css'

output = BASE / 'Archivo.pdf'

print("\nRenderizando HTML a PDF...")

HTML(filename=str(html_path), base_url=str(html_path.parent)).write_pdf(target=str(output), stylesheets=[CSS(filename=str(css_path))])

print(f'\nPDF generado\n')