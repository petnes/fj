# -*- coding: utf-8 -*-
"""Documento único con todas las partes de Salima: 1 (introducción), 2, 4, 5, 8, anexo VI y referencias.
Uso: python3 build_partes_salima.py [salida.docx]"""
import sys, os, re, json, glob, unicodedata, importlib.util
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

HERE = os.path.dirname(os.path.abspath(__file__))
BIBLIO = os.path.dirname(HERE)
OUT = sys.argv[1] if len(sys.argv) > 1 else '/home/user/fj/Salima_grupo_corazon/Partes_de_Salima_trabajo_grupo_COMPLETO.docx'

def mod(name):
    spec = importlib.util.spec_from_file_location(name, os.path.join(HERE, name + '.py'))
    m = importlib.util.module_from_spec(spec); spec.loader.exec_module(m); return m
A2 = mod('texto_ap2_v2'); A458 = mod('texto_ap458'); IA = mod('texto_intro_anexo')

# Dos obras del Ministerio de 2026: la nota de Think Equal (2026a) y la resolución del calendario (2026b).
def ab(s):
    s = s.replace('Formación Profesional y Deportes, 2026)', 'Formación Profesional y Deportes, 2026a)') if 'Think Equal' in s else s
    return s
S26_txt = [ab(p) for p in A2.S26[1]]
CIERRE_txt = list(A2.CIERRE[1])
bloques458 = []
for k, c in A458.BLOQUES:
    if k in ('p', 'obj'):
        c = c.replace('(Ministerio de Educación, Formación Profesional y Deportes, 2026)', '(Ministerio de Educación, Formación Profesional y Deportes, 2026b)')
    bloques458.append((k, c))

# ---------- referencias ----------
fichas = {}
for f in sorted(glob.glob(os.path.join(BIBLIO, 'fichas', 'fichas_*.json'))):
    for x in json.load(open(f, encoding='utf-8')): fichas.setdefault(x['id'], x)
INITIALS = re.compile(r"\b[A-ZÁÉÍÓÚÑÈÖÜ]\.(?:[\s-]?[A-ZÁÉÍÓÚÑÈÖÜ]\.)*")
def sin_acentos(s): return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
def autores_y_anio(ref):
    m = re.search(r'\((\d{4}[ab]?|s\. f\.)', ref); return ref[:m.start()].strip(), m.group(1)[:4]
def apellidos(aut):
    if not INITIALS.search(aut): return [aut.rstrip('. ')]
    s = re.sub(r'\s+y\s+', ', ', INITIALS.sub('', aut)); return [a.strip(' .') for a in s.split(',') if a.strip(' .')]
def etiqueta(ref):
    aut, anio = autores_y_anio(ref); ap = apellidos(aut)
    lab = ap[0] if len(ap) == 1 else (f'{ap[0]} y {ap[1]}' if len(ap) == 2 else f'{ap[0]} et al.')
    return lab, anio
def clave(ref):
    if re.match(r'(Ley|Real Decreto|Orden) ', ref): return ([sin_acentos(ref.lower())], '0', '')
    aut, anio = autores_y_anio(ref)
    ap = [re.sub(r'[\s\-]', '', sin_acentos(a.lower())) for a in apellidos(aut)]
    m = INITIALS.search(aut); ini = m.group(0).replace(' ', '') if m else ''
    titulo = ref[ref.find(')') + 1:].strip(' .').lower()
    return ([ap[0], ini] + ap[1:], anio, titulo)

todo_texto = '\n'.join(IA.INTRO1[1] + A2.INTRO + [p for s in [A2.S21, A2.S22, A2.S23, A2.S24, A2.S25] for p in s[1]] + S26_txt + CIERRE_txt
                       + [c for k, c in bloques458 if k in ('p', 'obj')]
                       + [cell for k, c in bloques458 if k == 'tabla' for row in c['rows'] for cell in row]
                       + [IA.ANEXO_VI['nota']] + [cell for row in IA.ANEXO_VI['rows'] for cell in row])
refs = []
for fid, x in fichas.items():
    ref = x['referencia_apa7']; lab, anio = etiqueta(ref)
    if any(anio in todo_texto[m.end():m.end() + 45] for m in re.finditer(re.escape(lab), todo_texto)):
        if fid == 'Ministerio_2026_ThinkEqual': ref = ref.replace('(2026, 4 de febrero)', '(2026a, 4 de febrero)')
        refs.append(ref)
NORMATIVA = [
    "Consejo Escolar del Estado. (2023). *Informe 2023 sobre el estado del sistema educativo. Ceuta y Melilla. Curso 2021-2022*. Ministerio de Educación, Formación Profesional y Deportes. https://www.educacionfpydeportes.gob.es/dam/jcr:cc1dc202-ac37-414e-862e-11e335df7a2e/i23cee09-ceuta-melilla.pdf",
    "Ley Orgánica 2/2006, de 3 de mayo, de Educación. *Boletín Oficial del Estado*, 106, de 4 de mayo de 2006. Texto consolidado. https://www.boe.es/eli/es/lo/2006/05/03/2/con",
    "Ley Orgánica 3/2020, de 29 de diciembre, por la que se modifica la Ley Orgánica 2/2006, de 3 de mayo, de Educación. *Boletín Oficial del Estado*, 340, de 30 de diciembre de 2020. https://www.boe.es/eli/es/lo/2020/12/29/3",
    "Orden EFP/608/2022, de 29 de junio, por la que se establece el currículo y se regula la ordenación de la Educación Infantil en el ámbito de gestión del Ministerio de Educación y Formación Profesional. *Boletín Oficial del Estado*, 158, de 2 de julio de 2022. https://www.boe.es/eli/es/o/2022/06/29/efp608",
    "Real Decreto 95/2022, de 1 de febrero, por el que se establece la ordenación y las enseñanzas mínimas de la Educación Infantil. *Boletín Oficial del Estado*, 28, de 2 de febrero de 2022. https://www.boe.es/eli/es/rd/2022/02/01/95",
]
nuevas = [r.replace('Deportes. (2026). *Resolución', 'Deportes. (2026b). *Resolución') for r in A458.REFS_NUEVAS]
refs = sorted(set(refs + NORMATIVA + nuevas), key=clave)

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
def H(text, lvl=1): return doc.add_heading(text, level=lvl)
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
def TABLA(tb, size=9):
    if tb.get('titulo'):
        cap = P(tb['titulo'], size=10, italic=True, after=4); cap.alignment = WD_ALIGN_PARAGRAPH.LEFT
        cap.paragraph_format.keep_with_next = True; cap.paragraph_format.line_spacing = 1.15
    t = doc.add_table(rows=1, cols=len(tb['cols'])); t.style = 'Table Grid'; t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(tb['cols']):
        c = t.rows[0].cells[i]; c.text = ''; pr = c.paragraphs[0]; r = pr.add_run(h); r.bold = True; r.font.size = Pt(size)
        pr.alignment = WD_ALIGN_PARAGRAPH.LEFT; pr.paragraph_format.line_spacing = 1.0; pr.paragraph_format.space_after = Pt(2); shade(c, 'E7E6E6')
    cabecera(t)
    for rowv in tb['rows']:
        cells = t.add_row().cells
        for i, v in enumerate(rowv):
            cells[i].text = ''; pr = cells[i].paragraphs[0]; runs(pr, v, size=size)
            pr.alignment = WD_ALIGN_PARAGRAPH.LEFT; pr.paragraph_format.line_spacing = 1.0; pr.paragraph_format.space_after = Pt(2)
    set_widths(t, tb['widths'])
    doc.add_paragraph().paragraph_format.space_after = Pt(4)
def REF(text):
    p = doc.add_paragraph(); pf = p.paragraph_format
    pf.left_indent = Cm(1.25); pf.first_line_indent = Cm(-1.25); pf.space_after = Pt(6)
    pf.line_spacing = 1.15; p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    runs(p, text, size=10.5)

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
P('Partes redactadas por Salima López Rami: apartados 1, 2, 4, 5 y 8, anexo VI y referencias', align='center', size=10, italic=True, after=14)

H(IA.INTRO1[0], 1)
for p in IA.INTRO1[1]: P(p)

H('2. Marco teórico y antecedentes', 1)
for p in A2.INTRO: P(p)
for titulo, parrafos in [A2.S21, A2.S22, A2.S23, A2.S24, A2.S25]:
    H(titulo, 2)
    for p in parrafos: P(p)
H(A2.S26[0], 2)
for p in S26_txt: P(p)
H(A2.CIERRE[0], 2)
for p in CIERRE_txt: P(p)

for kind, content in bloques458:
    if kind == 'h1': H(content, 1)
    elif kind == 'h2': H(content, 2)
    elif kind == 'p': P(content)
    elif kind == 'obj':
        p = P(content, indent=0.75, after=6); p.paragraph_format.line_spacing = 1.3
    elif kind == 'tabla': TABLA(content)

H(IA.ANEXO_VI['titulo'], 1)
P(IA.ANEXO_VI['nota'], size=10, after=8)
TABLA({k: v for k, v in IA.ANEXO_VI.items() if k != 'titulo' and k != 'nota'}, size=8.5)

H('Referencias (apartados 1, 2, 4, 5 y 8 y anexo VI)', 1)
for ref in refs: REF(ref)

doc.save(OUT)
print('guardado:', OUT, '| referencias:', len(refs))
