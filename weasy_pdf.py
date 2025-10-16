from pathlib import Path
from weasyprint import HTML, CSS

BASE = Path(__file__).parent.resolve()

# Rutas de los archivos
html_path = BASE / "curso.html"
css_path  = BASE / "assets" / "style.css"

output = BASE / "finales" / "guia didactica_GD.pdf"

HTML(filename=str(html_path), base_url=str(html_path.parent)).write_pdf(target=str(output), stylesheets=[CSS(filename=str(css_path))])