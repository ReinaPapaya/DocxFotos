import os
import io
import zipfile
import json
import re
from flask import Flask, request, render_template, send_file, jsonify
from werkzeug.utils import secure_filename
from PIL import Image
from docx import Document
from docx.shared import Inches, Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB

# ---- Utilidades internas ----

def set_cell_background(cell, color_hex):
    """Pinta fondo de celda en tabla (para imágenes con fondo)"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex.replace('#', ''))
    tcPr.append(shd)

def add_caption(paragraph, text, numbering=False, num_prefix="Figura "):
    """Añade caption con numeración automática (estilo 'Caption' o normal)"""
    run = paragraph.add_run(text)
    run.bold = True
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    try:
        paragraph.style = 'Caption'
    except:
        pass

def insert_image_with_options(doc, img_stream, config, fig_num, caption_text):
    """
    Inserta una imagen en el documento según configuraciones.
    Retorna el párrafo o tabla donde se insertó.
    """
    pil_img = Image.open(img_stream)
    bg_color = config.get('fondo', 'transparente')
    if bg_color != 'transparente' and pil_img.mode in ('RGBA', 'LA', 'P'):
        bg = Image.new('RGB', pil_img.size, bg_color)
        if pil_img.mode == 'P':
            pil_img = pil_img.convert('RGBA')
        bg.paste(pil_img, (0, 0), pil_img if pil_img.mode == 'RGBA' else None)
        pil_img = bg
    elif pil_img.mode != 'RGB':
        pil_img = pil_img.convert('RGB')
    
    img_bytes = io.BytesIO()
    pil_img.save(img_bytes, format='PNG')
    img_bytes.seek(0)

    width_val = config.get('ancho', 80)
    width_unit = config.get('unidad_ancho', '%')
    if width_unit == '%':
        page_width_cm = 16.0
        width_inches = Inches((width_val / 100.0) * page_width_cm / 2.54)
    else:
        width_inches = Cm(float(width_val))

    alignment_map = {
        'izquierda': WD_ALIGN_PARAGRAPH.LEFT,
        'centrado': WD_ALIGN_PARAGRAPH.CENTER,
        'derecha': WD_ALIGN_PARAGRAPH.RIGHT
    }
    align = alignment_map.get(config.get('alineacion', 'centrado'), WD_ALIGN_PARAGRAPH.CENTER)
    
    pie_pos = config.get('pie_posicion', 'abajo')
    if pie_pos == 'lateral':
        table = doc.add_table(rows=1, cols=2)
        table.autofit = False
        table.allow_autofit = False
        cell_img = table.cell(0, 0)
        cell_img.width = int(width_inches) * 2
        cell_cap = table.cell(0, 1)
        paragraph_img = cell_img.paragraphs[0] if cell_img.paragraphs else cell_img.add_paragraph()
        run = paragraph_img.add_run()
        run.add_picture(img_bytes, width=width_inches)
        paragraph_img.alignment = align
        p_cap = cell_cap.paragraphs[0] if cell_cap.paragraphs else cell_cap.add_paragraph()
        if config.get('numeracion', False):
            caption_full = f"Figura {fig_num}: {caption_text}"
        else:
            caption_full = caption_text
        add_caption(p_cap, caption_full, numbering=False)
        if bg_color != 'transparente':
            set_cell_background(cell_img, bg_color)
        return table
    else:
        p_img = doc.add_paragraph()
        p_img.alignment = align
        run = p_img.add_run()
        run.add_picture(img_bytes, width=width_inches)
        p_cap = doc.add_paragraph()
        if config.get('numeracion', False):
            caption_full = f"Figura {fig_num}: {caption_text}"
        else:
            caption_full = caption_text
        add_caption(p_cap, caption_full, numbering=False)
        spacing = config.get('espaciado_pt', 12)
        p_cap.paragraph_format.space_before = Pt(spacing)
        p_cap.paragraph_format.space_after = Pt(spacing)
        return p_img

# ---- Endpoints ----

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generar', methods=['POST'])
def generar():
    semilla = request.files.get('semilla')
    if not semilla:
        return jsonify({'error': 'Semilla requerida'}), 400
    imagenes = request.files.getlist('imagenes')
    config_json = request.form.get('config')
    config = json.loads(config_json) if config_json else {}

    doc = Document(semilla)
    
    orden_str = request.form.get('orden', '')
    orden = json.loads(orden_str) if orden_str else []
    incluidos_str = request.form.get('incluidos', '')
    incluidos = json.loads(incluidos_str) if incluidos_str else []
    
    # Construir mapa clave -> archivo usando zip(orden, imagenes)
    img_map = {}
    for key, file in zip(orden, imagenes):
        if file and file.filename:
            img_map[key] = file
    if not img_map:
        for i, file in enumerate(imagenes):
            if file and file.filename:
                img_map[f'img_{i}'] = file
    
    ordered_keys = orden if orden else list(img_map.keys())
    if incluidos:
        ordered_keys = [k for k in ordered_keys if k in incluidos and k in img_map]
    
    inserted_elements = []
    fig_counter = 1
    # Cambio: valor predeterminado a '###IMAGENES###'
    bookmark = config.get('bookmark', '###IMAGENES###')
    
    one_per_page = config.get('modo_ajuste') == 'una_por_pagina'
    
    for key in ordered_keys:
        file = img_map[key]
        caption_type = config.get('caption_tipo', 'nombre_archivo')
        if caption_type == 'nombre_archivo':
            caption_text = file.filename.rsplit('.', 1)[0]
        elif caption_type == 'secuencial':
            caption_text = f"Imagen {fig_counter}"
        else:  # 'personalizado'
            captions_custom = json.loads(request.form.get('captions_custom', '{}'))
            caption_text = captions_custom.get(key, file.filename)
        
        elem = insert_image_with_options(
            doc, file.stream, config,
            fig_counter if config.get('numeracion', False) else None,
            caption_text
        )
        inserted_elements.append(elem)
        
        if one_per_page and ordered_keys.index(key) < len(ordered_keys)-1:
            doc.add_page_break()
        
        fig_counter += 1
    
    # --- Reemplazo del marcador (versión robusta) ---
    elem_lxml = [e._element for e in inserted_elements if hasattr(e, '_element')]
    bookmark_found = False
    bookmark_clean = bookmark.strip()  # elimina espacios alrededor

    for p in doc.paragraphs:
        p_text = p.text.strip()  # elimina espacios y saltos
        if bookmark_clean in p_text or p_text == bookmark_clean:
            parent = p._element.getparent()
            for el in elem_lxml:
                parent.insert(parent.index(p._element), el)
            parent.remove(p._element)
            bookmark_found = True
            break

    if not bookmark_found:
        # Si no se encuentra, añadir al final
        for el in elem_lxml:
            doc.element.body.append(el)
    
    # Guardar en memoria
    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    
    nombre_salida = request.form.get('nombre_salida', 'documento_generado.docx')
    if not nombre_salida.endswith('.docx'):
        nombre_salida += '.docx'
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        as_attachment=True,
        download_name=nombre_salida
    )

if __name__ == '__main__':
    app.run(debug=False, host='0.0.0.0', port=5000)
