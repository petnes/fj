# -*- coding: utf-8 -*-
"""Construye el apartado 2 (v2) en Word a partir de texto_ap2_v2.py y de las fichas.
Uso: python3 build_ap2_v2.py [salida.docx]"""
import sys, os, re, json, glob, unicodedata, importlib.util
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

HERE = os.path.dirname(os.path.abspath(__file__))
BIBLIO = os.path.dirname(HERE)
OUT = sys.argv[1] if len(sys.argv) > 1 else '/home/user/fj/Salima_grupo_corazon/Apartado_2_Marco_teorico_y_antecedentes_Salima_v2.docx'

spec = importlib.util.spec_from_file_location('texto', os.path.join(HERE, 'texto_ap2_v2.py'))
T = importlib.util.module_from_spec(spec); spec.loader.exec_module(T)
SECCIONES = [T.S21, T.S22, T.S23, T.S24, T.S25, T.S26]

# ---------- referencias que no salen de las fichas ----------
MANUALES = [
    "Consejo Escolar del Estado. (2023). *Informe 2023 sobre el estado del sistema educativo. Ceuta y Melilla. Curso 2021-2022*. Ministerio de Educación, Formación Profesional y Deportes. https://www.educacionfpydeportes.gob.es/dam/jcr:cc1dc202-ac37-414e-862e-11e335df7a2e/i23cee09-ceuta-melilla.pdf",
    "Ley Orgánica 2/2006, de 3 de mayo, de Educación. *Boletín Oficial del Estado*, 106, de 4 de mayo de 2006. Texto consolidado. https://www.boe.es/eli/es/lo/2006/05/03/2/con",
    "Ley Orgánica 3/2020, de 29 de diciembre, por la que se modifica la Ley Orgánica 2/2006, de 3 de mayo, de Educación. *Boletín Oficial del Estado*, 340, de 30 de diciembre de 2020. https://www.boe.es/eli/es/lo/2020/12/29/3",
    "Orden EFP/608/2022, de 29 de junio, por la que se establece el currículo y se regula la ordenación de la Educación Infantil en el ámbito de gestión del Ministerio de Educación y Formación Profesional. *Boletín Oficial del Estado*, 158, de 2 de julio de 2022. https://www.boe.es/eli/es/o/2022/06/29/efp608",
    "Real Decreto 95/2022, de 1 de febrero, por el que se establece la ordenación y las enseñanzas mínimas de la Educación Infantil. *Boletín Oficial del Estado*, 28, de 2 de febrero de 2022. https://www.boe.es/eli/es/rd/2022/02/01/95",
]
# patrones de cita en el texto que corresponden a las referencias manuales
MANUAL_CITAS = ['Consejo Escolar del Estado', 'Ley Orgánica 2/2006', 'Ley Orgánica 3/2020',
                'Orden EFP/608/2022', 'Real Decreto 95/2022']

INITIALS = re.compile(r"\b[A-ZÁÉÍÓÚÑÈÖÜ]\.(?:[\s-]?[A-ZÁÉÍÓÚÑÈÖÜ]\.)*")

def sin_acentos(s):
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')

def autores_y_anio(ref):
    m = re.search(r'\((\d{4})', ref)
    return ref[:m.start()].strip(), m.group(1)

def apellidos(autores):
    if not INITIALS.search(autores):  # autor institucional
        return [autores.rstrip('. ')]
    s = INITIALS.sub('', autores)
    s = re.sub(r'\s+y\s+', ', ', s)
    return [a.strip(' .') for a in s.split(',') if a.strip(' .')]

def etiqueta(ref):
    aut, anio = autores_y_anio(ref)
    ap = apellidos(aut)
    if len(ap) == 1: lab = ap[0]
    elif len(ap) == 2: lab = f'{ap[0]} y {ap[1]}'
    else: lab = f'{ap[0]} et al.'
    return lab, anio

def clave_orden(ref):
    if re.match(r'(Ley|Real Decreto|Orden) ', ref):
        return ([sin_acentos(ref.lower())], '0')
    aut, anio = autores_y_anio(ref)
    # orden letra a letra (sin espacios ni guiones); con el mismo primer apellido, deciden las iniciales
    ap = [re.sub(r'[\s\-]', '', sin_acentos(a.lower())) for a in apellidos(aut)]
    m = INITIALS.search(aut)
    ini = m.group(0).replace(' ', '') if m else ''
    return ([ap[0], ini] + ap[1:], anio)

# ---------- cargar fichas ----------
fichas = {}
for f in sorted(glob.glob(os.path.join(BIBLIO, 'fichas', 'fichas_*.json'))):
    for x in json.load(open(f, encoding='utf-8')):
        fichas.setdefault(x['id'], x)

texto = '\n'.join(T.INTRO + [p for s in SECCIONES for p in s[1]] + T.CIERRE[1])

citadas, no_citadas = [], []
for fid, x in fichas.items():
    lab, anio = etiqueta(x['referencia_apa7'])
    hit = False
    for m in re.finditer(re.escape(lab), texto):
        if anio in texto[m.end():m.end() + 45]:
            hit = True; break
    (citadas if hit else no_citadas).append(fid)

# comprobación inversa: citas del texto sin referencia
pat_cita = re.compile(r"((?:[A-ZÁÉÍÓÚÑ][\w'’\-áéíóúñàèòçü]+)(?: (?:de |del |van |San )?[A-ZÁÉÍÓÚÑ][\w'’\-áéíóúñàèòçü]+)*(?: et al\.| y (?:de |De |Van |San )?[A-ZÁÉÍÓÚÑ][\w'’\-áéíóúñ]+(?: [A-ZÁÉÍÓÚÑ][\w'’\-áéíóúñ]+)?)?),? \(?(\d{4})")
etiquetas = {etiqueta(fichas[i]['referencia_apa7']) for i in citadas}
huerfanas = set()
for m in pat_cita.finditer(texto):
    lab, anio = m.group(1), m.group(2)
    if any(lab.endswith(e[0]) and anio == e[1] for e in etiquetas):
        continue
    if any(mc.split()[0] in lab for mc in MANUAL_CITAS):
        continue
    huerfanas.add((lab, anio))

refs = [fichas[i]['referencia_apa7'] for i in citadas] + MANUALES
refs.sort(key=clave_orden)

# ---------- documento ----------
doc = Document()
for s in doc.sections:
    s.top_margin = s.bottom_margin = Cm(2.5); s.left_margin = s.right_margin = Cm(2.5)
st = doc.styles['Normal']; st.font.name = 'Arial'; st.font.size = Pt(11)
st.element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')
st.paragraph_format.space_after = Pt(8); st.paragraph_format.line_spacing = 1.5
st.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
for name, size in [('Heading 1', 13), ('Heading 2', 12)]:
    h = doc.styles[name]; h.font.name = 'Arial'; h.font.size = Pt(size); h.font.bold = True
    h.font.italic = False
    h.font.color.rgb = RGBColor(0, 0, 0); h.element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')
    h.paragraph_format.space_before = Pt(16 if name == 'Heading 1' else 12); h.paragraph_format.space_after = Pt(6)
    h.paragraph_format.line_spacing = 1.15; h.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
    h.paragraph_format.keep_with_next = True
    rf = h.element.rPr.rFonts
    for att in ('w:asciiTheme', 'w:hAnsiTheme', 'w:eastAsiaTheme', 'w:cstheme'):
        if rf.get(qn(att)) is not None: del rf.attrib[qn(att)]

def runs(p, text, size=None, bold=False, italic=False):
    for i, part in enumerate(text.split('*')):
        if not part: continue
        r = p.add_run(part); r.italic = italic or (i % 2 == 1); r.bold = bold
        if size: r.font.size = Pt(size)

def P(text, align=None, size=None, after=None, bold=False, italic=False):
    p = doc.add_paragraph(); runs(p, text, size, bold, italic)
    if align == 'center': p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if after is not None: p.paragraph_format.space_after = Pt(after)
    return p

def H(text, lvl=1): return doc.add_heading(text, level=lvl)

def REF(text):
    p = doc.add_paragraph(); pf = p.paragraph_format
    pf.left_indent = Cm(1.25); pf.first_line_indent = Cm(-1.25); pf.space_after = Pt(6)
    pf.line_spacing = 1.15; p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    runs(p, text, size=10.5)

# numeración de páginas en el pie
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
P('Apartado 2: Marco teórico y antecedentes · Redacta: Salima López Rami', align='center', size=10, italic=True, after=14)

H('2. Marco teórico y antecedentes', 1)
for p in T.INTRO: P(p)
for titulo, parrafos in SECCIONES:
    H(titulo, 2)
    for p in parrafos: P(p)
H(T.CIERRE[0], 2)
for p in T.CIERRE[1]: P(p)

H('Referencias del apartado 2', 1)
for r_ in refs: REF(r_)

doc.save(OUT)
json.dump({'citadas': sorted(citadas), 'no_citadas': sorted(no_citadas)}, open(os.path.join(BIBLIO, 'citadas.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
print('guardado:', OUT)
print('referencias:', len(refs), '| fichas citadas:', len(citadas), '| fichas no citadas:', len(no_citadas))
print('no citadas:', ', '.join(sorted(no_citadas)))
print('citas sin referencia (revisar):', sorted(huerfanas))
