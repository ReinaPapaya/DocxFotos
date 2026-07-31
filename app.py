from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

doc = Document()

# Título
titulo = doc.add_heading('MANUAL DE USO – GENERADOR DOCX CON IMÁGENES', 0)
titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER

# Instrucciones
doc.add_paragraph('1. Ve a https://docxfotos.onrender.com')
doc.add_paragraph('2. Carga este mismo documento como SEMILLA (o tu propia plantilla).')
doc.add_paragraph('3. Carga las imágenes que quieras insertar.')
doc.add_paragraph('4. Configura opciones (ancho, alineación, pies, etc.).')
doc.add_paragraph('5. Haz clic en "Generar DOCX".')
doc.add_paragraph('')
doc.add_paragraph('El marcador #INSERT_IMAGES# se reemplazará por las imágenes.')
doc.add_paragraph('Puedes cambiarlo en el campo "Marcador" de la herramienta.')
doc.add_paragraph('')
# Separador
doc.add_paragraph('----------------------------------------------------------------')
# Marcador (debe estar en un párrafo aparte)
doc.add_paragraph('#INSERT_IMAGES#')
doc.add_paragraph('----------------------------------------------------------------')
doc.add_paragraph('')
doc.add_paragraph('Texto final después de las imágenes.')

doc.save('semilla_con_manual.docx')
print('✅ semilla_con_manual.docx creado')
