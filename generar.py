############################################################################
# Importar librerías
############################################################################

import re
import os
import pandas as pd
from pathlib import Path
from canvasapi import Canvas
from bs4 import BeautifulSoup
from weasyprint import HTML, CSS

############################################################################
# Variables globales
############################################################################

API_URL = "https://utpl.instructure.com/"
API_KEY = "2795~Ac8vr3ThZR9PJ44trcerXxJQeuz46JfkuVRQcGXtQBc8W3MGmVKt6LymATYeFTMn"

# Columnas del excel
CODIGO = 'Código Banner'
CODIGO_SIS = 'Código Canvas'
CURSO = 'Módulo (malla curricular)'
AUTOR = 'Apellidos y nombres del Docente'
FACULTAD = 'Facultad'
MAESTRIA = 'Nombre del Programa'
CICLO = 'Ciclo'
BIMESTRE = 'Bimestre'
URL = 'URL metacurso nuevo (lo gestiona la Dirección de Posgrados)'
ISBN = 'ISBN DIGITAL'
FECHA_ISBN = 'FECHA ISBN'

# Ruta base
BASE = Path(__file__).parent.resolve()

############################################################################
# Funciones
############################################################################

def read_file(ruta_archivo):
    """
    Lee un archivo Excel y retorna un DataFrame de pandas.
    :param ruta_archivo: Ruta del archivo Excel (.xlsx, .xls)
    :return: DataFrame con los datos
    """
    df = pd.read_excel(ruta_archivo)
    return df


def get_file_html(ruta_archivo):
    """
    Lee el contenido de un archivo HTML local y lo retorna como string.
    :param ruta_archivo: Ruta del archivo HTML.
    :return: Contenido HTML como string.
    """
    with open(ruta_archivo, "r", encoding="utf-8") as f:
        html = f.read()
    return html
    
def write_file(ruta_archivo, contenido):
    """
    Escribe contenido en un archivo HTML local.
    :param ruta_archivo: Ruta del archivo HTML.
    :param contenido: Contenido a escribir en el archivo.
    """
    with open(ruta_archivo, "w", encoding="utf-8") as f:
        f.write(contenido)

def delete_elements(div):
    """
    Busca la primera etiqueta <h2> y luego la elimina
    Buscar un div con clase button-actividad y lo elimina
    :param div: Objeto soup.
    :return: HTML modificado con el iframe reemplazado.
    """
    h2 = div.find("h2")
                        
    if h2:
        h2.decompose()

    # Eliminar los botones de actividad
    for node in div.select("div.button-actividad"):
        node.decompose()

    # Elimiar espacios en blanco de párrafos sin imágenes y sin iframes
    paragraphs = div.find_all('p')
    
    for p in paragraphs:
        img = p.find("img")
        iframe = p.find("iframe")

        if not img and not iframe:
            # Si el párrafo no tiene texto (después de strip)
            if not p.get_text(strip=True):
                p.decompose()  # Elimina el elemento del DOM

    return div

def replace_iframe(html):
    """
    Reemplaza todos los elementos <iframe> en el HTML por enlaces <a>.
    Si el iframe tiene un atributo title, se usa como texto del enlace.
    Si no tiene title, se usa un texto por defecto "Ver recurso".
    Si el iframe no tiene src, se elimina del HTML.
    :param html: Contenido HTML original.
    :return: HTML modificado con el iframe reemplazado.
    """

    DEFAULT_TEXT = "Ver recurso"

    soup = BeautifulSoup(html, "html.parser")

    for ifr in soup.find_all("iframe"):
        src = ifr.get("src")
        title = ifr.get("title", "").strip()

        if not src:
            ifr.decompose()
            continue

        # Crear el <a> nuevo
        a = soup.new_tag("a", href=src)
        a["target"] = "_blank"
        a["title"] = title
        a.string = title if title else DEFAULT_TEXT

        ifr.replace_with(a)
    
    return str(soup)

def replace_tags(str_html, type=1):

    str_html = str_html.replace('<pre>', '<p>')
    str_html = str_html.replace('</pre>', '</p>')

    if type == 1:
        str_html = str_html.replace('<h5>', '<h6>')
        str_html = str_html.replace('</h5>', '</h6>')
        str_html = str_html.replace('<h4>', '<h5>')
        str_html = str_html.replace('</h4>', '</h5>')
        str_html = str_html.replace('<h3>', '<h4>')
        str_html = str_html.replace('</h3>', '</h4>')
        str_html = str_html.replace('<h2>', '<h3>')
        str_html = str_html.replace('</h2>', '</h3>')
    elif type == 2:
        str_html = str_html.replace('<h5>', '<p>')
        str_html = str_html.replace('</h5>', '</p>')
        str_html = str_html.replace('<h4>', '<h6>')
        str_html = str_html.replace('</h4>', '</h6>')
        str_html = str_html.replace('<h3>', '<h5>')
        str_html = str_html.replace('</h3>', '</h5>')
        str_html = str_html.replace('<h2>', '<h4>')
        str_html = str_html.replace('</h2>', '</h4>')
    elif type == 3:
        str_html = str_html.replace('<h5>', '<p>')
        str_html = str_html.replace('</h5>', '</p>')
        str_html = str_html.replace('<h4>', '<p>')
        str_html = str_html.replace('</h4>', '<p>')
        str_html = str_html.replace('<h3>', '<h6>')
        str_html = str_html.replace('</h3>', '</h6>')
        str_html = str_html.replace('<h2>', '<h5>')
        str_html = str_html.replace('</h2>', '</h5>')

    return str_html


def generate_table_of_contents(html):
    """
    Genera un índice (tabla de contenidos) basado en los encabezados <h1>, <h2>, <h3> del HTML.
    :param html: Contenido HTML original.
    :return: HTML del índice generado.
    """
    soup = BeautifulSoup(html, "html.parser")
    toc = '<div id="indice_contenidos" class="indice"><h1>Índice de contenidos</h1><ul>'

    h1_count = 0
    h2_count = 0
    h3_count = 0
    h4_count = 0
    
    for header in soup.find_all(['h1', 'h2', 'h3', 'h4']):
        if header.name == 'h1':
            h1_count += 1
            header_id = f'semana_{h1_count}'
            toc += f'<li><a class="nivel-1" href="#{header_id}">{header.text}</a></li>'
        elif header.name == 'h2':
            h2_count += 1
            header_id = f'unidad_{h2_count}'
            toc += f'<li><a class="nivel-2" href="#{header_id}">{header.text}</a></li>'
        elif header.name == 'h3':
            h3_count += 1
            header_id = f'tema_{h3_count}'
            toc += f'<li><a class="nivel-3" href="#{header_id}">{header.text}</a></li>'
        elif header.name == 'h4':
            h4_count += 1
            header_id = f'subtema_{h4_count}'
            toc += f'<li><a class="nivel-4" href="#{header_id}">{header.text}</a></li>'

        # Asignar el ID al encabezado en el HTML original
        header['id'] = header_id

    toc += '</ul></div>'

    # Índice de figuras

    toc += '<div class="indice"><h1>Índice de figuras</h1><ul>'

    figuras = soup.find_all(id=re.compile(r'^figura_\d+$'))

    for figura in figuras:
        # Obtener el nombre de la figura
        p = figura.find('p')
        title = str(p).replace('<br/>', '. ').replace('<br>', '. ')
        title = title.replace('<p>', '').replace('</p>', '')
        title = BeautifulSoup(title, 'html.parser')

        # Ibtener el ID de la figura
        header_id = figura.get('id')
        toc += f'<li><a class="nivel" href="#{header_id}">{title}</a></li>'

    toc += '</ul></div>'

    # Índice de tablas

    toc += '<div class="indice"><h1>Índice de tablas</h1><ul>'

    tablas = soup.find_all(id=re.compile(r'^tabla_\d+$'))

    for contenedor_tabla in tablas:
        table = contenedor_tabla.find('table')
        caption = table.find('caption')
        caption = str(caption).replace('<br/>', '. ').replace('<br>', '. ')
        title = BeautifulSoup(caption, 'html.parser')

        # Obtener id de la tabla
        header_id = contenedor_tabla.get('id')
        toc += f'<li><a class="nivel" href="#{header_id}">{title}</a></li>'

    toc += '</ul></div>'

    
    return toc + str(soup)

def clear_styles(div):
    """Versión simple que solo limpia estilos inline"""
    
    estilos_permitidos = ['padding-left', 'margin-left', 'margin-right', 
                          'border', 'list-style-type', 'max-width', 'color']
    
    
    for elemento in div.find_all(style=True):
        estilo_original = elemento['style']
        declaraciones = [d.strip() for d in estilo_original.split(';') if d.strip()]
        
        declaraciones_filtradas = []
        for declaracion in declaraciones:
            if ':' in declaracion:
                propiedad = declaracion.split(':')[0].strip().lower()
                if propiedad in estilos_permitidos:
                    declaraciones_filtradas.append(declaracion)
        
        if declaraciones_filtradas:
            elemento['style'] = '; '.join(declaraciones_filtradas)
        else:
            del elemento['style']
    
    return div

def update_link_quiz(html_string):
    """
    Encuentra etiquetas <a> con data-api-returntype="Quiz" y las 
    reemplaza por <span> con clase "enlace-actualizado" y texto adicional.
    
    Args:
        html_string (str): String con el contenido HTML
        
    Returns:
        str: HTML con enlaces actualizados
    """

    soup=BeautifulSoup(html_string, 'html.parser')
    
    # Buscar todas las etiquetas <a> con el atributo específico
    enlaces_quiz = soup.find_all('a', attrs={'data-api-returntype': 'Quiz'})
    
    for enlace in enlaces_quiz:
        # Crear nueva etiqueta <a>
        nuevo_enlace = soup.new_tag('a')
        nuevo_enlace['class'] = 'enlace-actualizado'
        nuevo_enlace['href'] = 'https://utpl.instructure.com/'
        nuevo_enlace['title'] = 'Entorno Virtual de Aprendizaje'
        
        # Obtener el contenido actual del enlace
        contenido_actual = enlace.decode_contents()  # Obtiene el HTML interno
        
        # Agregar contenido original + texto nuevo
        nuevo_enlace.append(BeautifulSoup(contenido_actual, 'html.parser'))
        nuevo_enlace.append(' (revise en plataforma)')
        
        # Reemplazar el enlace por el span
        enlace.replace_with(nuevo_enlace)
    
    return soup

def etiquetar_indice(div):
    # Figuras
    imgs = div.find_all("div", class_="contenedor-imagen")
    
    for contenedor_img in imgs:
        p = contenedor_img.find('p')
        
        if p:
            texto = p.get_text(strip=True)
            patron = r'Figura\s*[:\.]?\s*(\d+)'
            coincidencia = re.search(patron, texto, re.IGNORECASE)
            
            if coincidencia:
                numero_figura = coincidencia.group(1)
                contenedor_img['id'] = f'figura_{numero_figura}'
    
    # Tablas
    tables = div.find_all("div", class_="contenedor-tabla")

    for contenedor_tabla in tables:
        table = contenedor_tabla.find('table')
        caption = table.find('caption')

        if caption:
            texto = caption.get_text(strip=True)
            patron = r'Tabla\s*[:\.]?\s*(\d+)'
            coincidencia = re.search(patron, texto, re.IGNORECASE)

            if coincidencia:
                numero_tabla = coincidencia.group(1)
                contenedor_tabla['id'] = f'tabla_{numero_tabla}'

    return div

############################################################################
# Programa principal
############################################################################

def main():
    # Rutas de los archivos
    css_path  = BASE / "assets" / "style.css"
    template  = BASE / "templates" / "portada.html"
    contraportada  = BASE / "templates" / "contraportada.html"
    
    canvas = Canvas(API_URL, API_KEY)
    
    # Lectura del excel
    df = read_file("datos.xlsx")

    for i, row in df.iterrows():
        print(f"\nIniciando generación de guía -> {row[CODIGO]} - {row[CURSO]}\n")

        html_path = BASE / "finales" / f'{row[CODIGO_SIS]}.html' # Ruta del archivo generado

        print('Actualizando datos de la oferta\n')
        html_portada = get_file_html(template)

        # Reemplazar datos por documento excel (oferta académica)
        html_portada = html_portada.replace('{{maestria}}', str(row[MAESTRIA]))
        html_portada = html_portada.replace('{{modulo}}', str(row[CURSO]))
        html_portada = html_portada.replace('{{autor}}', str(row[AUTOR]))
        html_portada = html_portada.replace('{{facultad}}', str(row[FACULTAD]))
        html_portada = html_portada.replace('{{ciclo}}', str(row[CICLO]))
        html_portada = html_portada.replace('{{bimestre}}', str(row[BIMESTRE]))
        html_portada = html_portada.replace('{{codigo}}', str(row[CODIGO]))
        
        url = row[URL]
        regex = re.search(r'(\d+)$', url)
        course_id = int(regex.group(1))
        
        course = canvas.get_course(course_id)
        modules = course.get_modules()

        html_course = ""
        menu_semanas = '<header class="menu-semanas"><div class="titulo-semana">Semanas</div>'

        i = 0 # Contador de semanas
        j = 0 # Contador de unidades

        for m in modules:
            if 'semana' in m.name.lower():
                i += 1
                html_course += f'<h1>{m.name}</h1>'
                
                boton_semana = f'<a title="Semana {i}" href="#semana_{i}">{i}</a>'
                menu_semanas += boton_semana

                print(f'Obteniendo contenido de Canvas <{m.name}>')

                items = m.get_module_items()
                page_url = items[0].page_url
                page = course.get_page(page_url)
                soup = BeautifulSoup(page.body, "html.parser")

                h3 = soup.find_all("h3")
                for h in h3:
                    if 'unidad' in h.text.lower():
                        html_course += f'<h2>{h.text}</h2>'
                        j += 1

                links = soup.select('p.tema a')

                if links:
                    for link in links:
                        e = link.get('data-api-endpoint')

                        if e:
                            replace = f'https://utpl.instructure.com/api/v1/courses/{course_id}/pages/'
                            page_id = e.replace(replace, '')
                            page = course.get_page(page_id)

                            print(f'\t<{page.title}>')

                            soup = BeautifulSoup(page.body, "html.parser")

                            # Manipulación del contenido de cada página
                            div = soup.find("div", class_="f-contenido-pagina")

                            # Etiquetar figuras y tablas para obtener el índice
                            div = etiquetar_indice(div)

                            # Eliminar título de la página y botón de actividad
                            div = delete_elements(div)

                            # Limpiar estilos
                            div = clear_styles(div)

                            # Actualizar enlaces de autoevaluaciones
                            div = update_link_quiz(str(div))

                            # Reemplazar iframes por enlaces
                            div = replace_iframe(str(div))

                            # Identificar el nivel de encabezado del enlace
                            style = link.get('style', '') 

                            if style == '':
                                html_course += f'<h3>{link.text}</h3>'
                                div = replace_tags(str(div), 1)
                            elif style == 'padding-left: 40px;':
                                html_course += f'<h4>{link.text}</h4>'
                                div = replace_tags(str(div), 2)
                            elif style == 'padding-left: 60px;':
                                html_course += f'<h5>{link.text}</h5>'
                                div = replace_tags(str(div), 3)

                            html_course += div
                else:
                    html_course += '<p>Revisión de contenidos</p>'


        menu_semanas += '</header>'

        # Arreglar el menú de acceso de las semanas
        # botones_semanas = f'<div class="menu-semanas salto-pagina"><div class="titulo-semana">Semanas</div>{menu_semanas}</div>'

        # Agregar contenedor de todo el HTML del curso
        html_course = f'<div id="contenidos" class="contenedor">{menu_semanas}{html_course}</div>'
        # html_course = f'<div id="contenidos" class="contenedor">{html_course}</div>'

        # Generar índice + contenido
        html_indice = generate_table_of_contents(html_course)

        # Contraportada
        html_contraportada = get_file_html(contraportada)

        # Unificar partes de la guía
        html_final = html_portada + html_indice + html_contraportada

        # Generar PDF 
        write_file(html_path, html_final)

        output = BASE / "finales" / f"{row[CODIGO_SIS]}.pdf"

        print("\nRenderizando HTML a PDF...")

        HTML(filename=str(html_path), base_url=str(html_path.parent)).write_pdf(target=str(output), stylesheets=[CSS(filename=str(css_path))])

        print(f'\nPDF generado ({row[CODIGO_SIS]}.pdf)\n')

if __name__ == "__main__":
    main()