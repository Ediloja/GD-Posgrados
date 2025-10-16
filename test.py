from pathlib import Path
from weasyprint import HTML, CSS

BASE = Path(__file__).parent.resolve()

# Rutas de los archivos
html_path = BASE / "test.html"
css_path  = BASE / "test.css"

output = BASE / "finales" / "Test.pdf"

HTML(filename=str(html_path), base_url=str(html_path.parent)).write_pdf(target=str(output), stylesheets=[CSS(filename=str(css_path))])