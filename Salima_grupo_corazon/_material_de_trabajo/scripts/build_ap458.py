# -*- coding: utf-8 -*-
"""Construye en Word los apartados 4, 5 y 8 (parte de Salima) a partir de texto_ap458.py.
Uso: python3 build_ap458.py [salida.docx]"""
import sys, os, re, json, glob, importlib.util
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

HERE = os.path.dirname(os.path.abspath(__file__))
BIBLIO = os.path.dirname(HERE)
OUT = sys.argv[1] if len(sys.argv) > 1 else '/home/user/fj/Salima_grupo_corazon/Apartados_4_5_8_Beneficiarios_Objetivos_Variables_Salima.docx'

spec = importlib.util.spec_from_file_location('t458', os.path.join(HERE, 'texto_ap458.py'))
T = importlib.util.module_from_spec(spec); spec.loader.exec_module(T)

# ---------- referencias: las citadas en estos apartados que ya están en el apartado 2 ----------
fichas = {}
for f in sorted(glob.glob(os.path.join(BIBLIO, 'fichas', 'fichas_*.json'))):
    for x in json.load(open(f, encoding='utf-8')): fichas.setdefault(x['id'], x)
INITIALS = re.compile(r"\b[A-ZÁÉÍÓÚÑÈÖÜ]\.(?:[\s-]?[A-ZÁÉÍÓÚÑÈÖÜ]\.)*")
def etiqueta(ref):
    m = re.search(r'\((\d{4})', ref); aut = ref[:m.start()].strip(); anio = m.group(1)
    if not INITIALS.search(aut): ap = [aut.rstrip('. ')]
    else:
        s = re.sub(r'\s+y\s+', ', ', INITIALS.sub('', aut)); ap = [a.strip(' .') for a in s.split(',') if a.strip(' .')]
    lab = ap[0] if len(ap) == 1 else (f'{ap[0]} y {ap[1]}' if len(ap) == 2 else f'{ap[0]} et al.')
    return lab, anio
texto = '\n'.join(c for k, c in T.BLOQUES if k in ('p', 'obj')) + '\n' + '\n'.join(
    ' '.join(cell for row in c['rows'] for cell in row) for k, c in T.BLOQUES if k == 'tabla')
citadas_ap2 = []
for fid, x in fichas.items():
    lab, anio = etiqueta(x['referencia_apa7'])
    if any(anio in texto[m.end():m.end() + 45] for m in re.finditer(re.escape(lab), texto)):
        citadas_ap2.append(x['referencia_apa7'])
NORMATIVA_AP2 = [
    "Consejo Escolar del Estado. (2023). *Informe 2023 sobre el estado del sistema educativo. Ceuta y Melilla. Curso 2021-2022*. Ministerio de Educación, Formación Profesional y Deportes. https://www.educacionfpydeportes.gob.es/dam/jcr:cc1dc202-ac37-414e-862e-11e335df7a2e/i23cee09-ceuta-melilla.pdf",
    "Ley Orgánica 2/2006, de 3 de mayo, de Educación. *Boletín Oficial del Estado*, 106, de 4 de mayo de 2006. Texto consolidado. https://www.boe.es/eli/es/lo/2006/05/03/2/con",
    "Orden EFP/608/2022, de 29 de junio, por la que se establece el currículo y se regula la ordenación de la Educación Infantil en el ámbito de gestión del Ministerio de Educación y Formación Profesional. *Boletín Oficial del Estado*, 158, de 2 de julio de 2022. https://www.boe.es/eli/es/o/2022/06/29/efp608",
    "Real Decreto 95/2022, de 1 de febrero, por el que se establece la ordenación y las enseñanzas mínimas de la Educación Infantil. *Boletín Oficial del Estado*, 28, de 2 de febrero de 2022. https://www.boe.es/eli/es/rd/2022/02/01/95",
]
def sin_acentos(s):
    import unicodedata
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
def clave(ref):
    if re.match(r'(Ley|Real Decreto|Orden) ', ref): return sin_acentos(ref.lower())
    return sin_acentos(ref.lower()).replace('-', ' ')
todas = sorted(citadas_ap2 + NORMATIVA_AP2 + T.REFS_NUEVAS, key=clave)

# ---------- documento ----------
doc = Document()
for s in doc.sections:
    s.top_margin = s.bottom_margin = Cm(2.5); s.left_margin = s.right_margin = Cm(2.5)
st = doc.styles['Normal']; st.font.name = 'Arial'; st.font.size = Pt(11)
st.element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')
st.paragraph_format.space_after = Pt(8); st.paragraph_format.line_spacing = 1.5
st.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
for name, size in [('Heading 1', 13), ('Heading 2', 12)]:
    h = doc.styles[name]; h.font.name = 'Arial'; h.font.size = Pt(size); h.font.bold = True; h.font.italic = False
    h.font.color.rgb = RGBColor(0, 0, 0); h.element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')
    rf = h.element.rPr.rFonts
    for att in ('w:asciiTheme', 'w:hAnsiTheme', 'w:eastAsiaTheme', 'w:cstheme'):
        if rf.get(qn(att)) is not None: del rf.attrib[qn(att)]
    h.paragraph_format.space_before = Pt(16 if name == 'Heading 1' else 12); h.paragraph_format.space_after = Pt(6)
    h.paragraph_format.line_spacing = 1.15; h.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
    h.paragraph_format.keep_with_next = True

def runs(p, text, size=None, bold=False, italic=False):
    for i, part in enumerate(text.split('*')):
        if not part: continue
        r = p.add_run(part); r.italic = italic or (i % 2 == 1); r.bold = bold
        if size: r.font.size = Pt(size)

def P(text, align=None, size=None, after=None, bold=False, italic=False, indent=None):
    p = doc.add_paragraph(); runs(p, text, size, bold, italic)
    if align == 'center': p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if after is not None: p.paragraph_format.space_after = Pt(after)
    if indent: p.paragraph_format.left_indent = Cm(indent)
    return p

def shade(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr(); shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto'); shd.set(qn('w:fill'), hex_color); tcPr.append(shd)

def set_widths(t, widths):
    t.autofit = False
    grid = t._tbl.find(qn('w:tblGrid'))
    for i, gc in enumerate(grid.findall(qn('w:gridCol'))): gc.set(qn('w:w'), str(int(widths[i] * 567)))
    for row in t.rows:
        for i, w in enumerate(widths): row.cells[i].width = Cm(w)

def cabecera(t):
    trPr = t.rows[0]._tr.get_or_add_trPr(); el = OxmlElement('w:tblHeader'); el.set(qn('w:val'), 'true'); trPr.append(el)

def TABLA(tb):
    cap = P(tb['titulo'], size=10, italic=True, after=4); cap.alignment = WD_ALIGN_PARAGRAPH.LEFT
    cap.paragraph_format.keep_with_next = True; cap.paragraph_format.line_spacing = 1.15
    t = doc.add_table(rows=1, cols=len(tb['cols'])); t.style = 'Table Grid'; t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(tb['cols']):
        c = t.rows[0].cells[i]; c.text = ''; pr = c.paragraphs[0]; r = pr.add_run(h); r.bold = True; r.font.size = Pt(9)
        pr.alignment = WD_ALIGN_PARAGRAPH.LEFT; pr.paragraph_format.line_spacing = 1.0; pr.paragraph_format.space_after = Pt(2); shade(c, 'E7E6E6')
    cabecera(t)
    for rowv in tb['rows']:
        cells = t.add_row().cells
        for i, v in enumerate(rowv):
            cells[i].text = ''; pr = cells[i].paragraphs[0]; runs(pr, v, size=9)
            pr.alignment = WD_ALIGN_PARAGRAPH.LEFT; pr.paragraph_format.line_spacing = 1.0; pr.paragraph_format.space_after = Pt(2)
    set_widths(t, tb['widths'])
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

def REF(text):
    p = doc.add_paragraph(); pf = p.paragraph_format
    pf.left_indent = Cm(1.25); pf.first_line_indent = Cm(-1.25); pf.space_after = Pt(6)
    pf.line_spacing = 1.15; p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    runs(p, text, size=10.5)

# numeración de páginas
fp = doc.sections[0].footer.paragraphs[0]; fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = fp.add_run()
for tag, txt in [('begin', None), (None, 'PAGE'), ('end', None)]:
    if tag:
        e = OxmlElement('w:fldChar'); e.set(qn('w:fldCharType'), tag); r._r.append(e)
    else:
        e = OxmlElement('w:instrText'); e.set(qn('xml:space'), 'preserve'); e.text = txt; r._r.append(e)
r.font.size = Pt(9)

P('*Un corazón, muchas formas de sentir*', bold=True, size=14, align='center', after=2)
P('Programa de educación emocional intercultural para un aula de 5 años y sus familias. Melilla, curso 2026-2027', align='center', size=11, after=2)
P('Conceptos básicos sobre la diversidad cultural · Máster en Diversidad Cultural (Universidad de Granada, Campus de Melilla)', align='center', size=10, after=2)
P('Apartados 4 (Beneficiarios), 5 (Objetivos) y 8 (Variables contextuales) · Redacta: Salima López Rami', align='center', size=10, italic=True, after=14)

for kind, content in T.BLOQUES:
    if kind == 'h1': doc.add_heading(content, level=1)
    elif kind == 'h2': doc.add_heading(content, level=2)
    elif kind == 'p': P(content)
    elif kind == 'obj':
        p = P(content, indent=0.75, after=6); p.paragraph_format.line_spacing = 1.3
    elif kind == 'tabla': TABLA(content)

doc.add_heading('Referencias citadas en estos apartados', 1)
P('Se incluyen todas las fuentes citadas en los apartados 4, 5 y 8. Las que ya figuran en el apartado 2 se repiten aquí para facilitar la unión de la lista final del trabajo.', size=10, italic=True, after=8)
for ref in todas: REF(ref)

doc.save(OUT)
print('guardado:', OUT, '| referencias:', len(todas))
