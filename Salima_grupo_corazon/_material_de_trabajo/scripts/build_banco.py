import json, glob, os, re, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from limpieza import limpiar
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

BASE = '/tmp/claude-0/-home-user-fj/17446f98-31ac-5557-ad69-3d259c778366/scratchpad/biblio'
OUT = sys.argv[1] if len(sys.argv) > 1 else '/home/user/fj/Salima_grupo_corazon/Banco_articulos_marco_teorico_Salima.docx'
EXTRA = json.load(open(f'{BASE}/extra_sections.json', encoding='utf-8')) if os.path.exists(f'{BASE}/extra_sections.json') else {}

fichas = []
for f in sorted(glob.glob(f'{BASE}/fichas/fichas_*.json')):
    try:
        data = json.load(open(f, encoding='utf-8'))
        fichas.extend(data)
    except Exception as e:
        print('ERROR leyendo', f, e)
excluded = set(EXTRA.get('excluir', []))
fichas = [x for x in fichas if x.get('id') not in excluded]
# dedupe by id keeping first
seen=set(); uniq=[]
for x in fichas:
    k = x.get('id')
    if k in seen: continue
    seen.add(k); uniq.append(x)
fichas = uniq
for x in fichas:
    for k, v in list(x.items()):
        if isinstance(v, str): x[k] = limpiar(v)
    pl = x.get('pdf_local', '')
    x['pdf_rel'] = pl.split('Salima_grupo_corazon/')[-1] if pl else ''
CIT = set(json.load(open(f'{BASE}/citadas.json', encoding='utf-8'))['citadas']) if os.path.exists(f'{BASE}/citadas.json') else set()
def es_completo(x):
    return x.get('acceso', '').strip().lower().startswith('texto completo')

BLOQUES = [
 ('2.1', 'Educación intercultural'),
 ('2.2', 'Desarrollo emocional en Educación Infantil'),
 ('2.3', 'Socialización emocional: familia, escuela y entorno'),
 ('2.4', 'Diversidad cultural, lingüística y expresión emocional'),
 ('2.5', 'Colaboración entre familia y escuela'),
 ('2.6', 'Antecedentes y experiencias de intervención'),
]
override_bloque = EXTRA.get('reasignar_bloque', {})
for x in fichas:
    if x.get('id') in override_bloque:
        x['bloque'] = override_bloque[x['id']]

doc = Document()
for s in doc.sections:
    s.top_margin = s.bottom_margin = Cm(2.0); s.left_margin = s.right_margin = Cm(2.0)
st = doc.styles['Normal']; st.font.name = 'Arial'; st.font.size = Pt(10)
st.element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')
st.paragraph_format.space_after = Pt(5); st.paragraph_format.line_spacing = 1.15
for name, size in [('Heading 1', 14), ('Heading 2', 12), ('Heading 3', 10.5)]:
    h = doc.styles[name]; h.font.name = 'Arial'; h.font.size = Pt(size); h.font.bold = True
    h.font.color.rgb = RGBColor(0x1F, 0x3A, 0x5F); h.element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')
    h.paragraph_format.space_before = Pt(12 if name != 'Heading 3' else 10); h.paragraph_format.space_after = Pt(4)

def cabecera(t):
    trPr = t.rows[0]._tr.get_or_add_trPr(); el = OxmlElement('w:tblHeader'); el.set(qn('w:val'), 'true'); trPr.append(el)

def shade(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr(); shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto'); shd.set(qn('w:fill'), hex_color); tcPr.append(shd)

def set_widths(t, widths):
    t.autofit = False
    grid = t._tbl.find(qn('w:tblGrid'))
    for i, gc in enumerate(grid.findall(qn('w:gridCol'))): gc.set(qn('w:w'), str(int(widths[i] * 567)))
    for row in t.rows:
        for i, w in enumerate(widths): row.cells[i].width = Cm(w)

def add_rich(par, text, size=None, bold=False):
    parts = (text or '').split('*')
    for i, part in enumerate(parts):
        if not part: continue
        r = par.add_run(part); r.italic = (i % 2 == 1); r.bold = bold
        if size: r.font.size = Pt(size)

def P(text, size=None, italic=False, bold=False, align=None, after=None):
    p = doc.add_paragraph(); r = p.add_run(text); r.italic = italic; r.bold = bold
    if size: r.font.size = Pt(size)
    if align == 'center': p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if after is not None: p.paragraph_format.space_after = Pt(after)
    return p

def short_ref(x):
    ref = x.get('referencia_apa7', '')
    m = re.match(r'(.+?)\((\d{4}[a-z]?|s\. f\.)\)', ref)
    if not m: return x.get('id', '')
    autores = m.group(1).strip().rstrip('.,')
    year = m.group(2)
    apellidos = re.findall(r"([A-ZÁÉÍÓÚÑÜ][\wÁÉÍÓÚÑÜáéíóúñüàèçïò'\-]+(?: [A-ZÁÉÍÓÚÑÜ][\wáéíóúñüàèçïò'\-]+)*), [A-Z]", autores)
    if not apellidos:
        return f'{autores} ({year})'
    if len(apellidos) == 1: a = apellidos[0]
    elif len(apellidos) == 2: a = f'{apellidos[0]} y {apellidos[1]}'
    else: a = f'{apellidos[0]} et al.'
    return f'{a} ({year})'

# ---------- Portada ----------
P('Banco de artículos y fichas de lectura', bold=True, size=16, align='center', after=2)
P('Apartado 2. Marco teórico y antecedentes', bold=True, size=12, align='center', after=2)
P('«Un corazón, muchas formas de sentir». Programa intercultural de educación emocional para un aula de 5 años de Melilla y sus familias', size=10.5, align='center', after=2)
P('Documento de trabajo de Salima López Rami para el grupo · Máster en Diversidad Cultural (UGR, Campus de Melilla)', italic=True, size=9.5, align='center', after=12)

for sec in EXTRA.get('intro', []):
    if sec.get('h'): doc.add_heading(sec['h'], level=1)
    for ptxt in sec.get('p', []):
        p = doc.add_paragraph(); add_rich(p, ptxt)

# ---------- Tabla resumen ----------
doc.add_heading('Resumen de las fuentes seleccionadas', level=1)
P('Cada fuente tiene su ficha completa más abajo. La columna «Acceso» indica si se ha leído el texto completo o solo el resumen, y la siguiente, si la fuente se cita en la versión 2 del apartado. Los PDF de acceso abierto están en la carpeta «articulos», ordenados por subapartado.', size=9.5)
for code, name in BLOQUES:
    items = [x for x in fichas if str(x.get('bloque', '')).startswith(code)]
    if not items: continue
    doc.add_heading(f'{code}. {name}', level=3)
    t = doc.add_table(rows=1, cols=5); t.style = 'Table Grid'; t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(['Fuente', 'Tipo', 'Acceso', 'Citada', 'Para qué sirve en el trabajo']):
        c = t.rows[0].cells[i]; c.text = ''; r = c.paragraphs[0].add_run(h); r.bold = True; r.font.size = Pt(8.5); shade(c, 'D9E2F3')
    cabecera(t)
    for x in items:
        cells = t.add_row().cells
        vals = [short_ref(x), x.get('tipo', ''), ('Texto completo' if es_completo(x) else 'Solo resumen'), ('Sí' if x.get('id') in CIT else 'No'), x.get('uso_en_el_trabajo', '')]
        for i, v in enumerate(vals):
            cells[i].text = ''; r = cells[i].paragraphs[0].add_run(v); r.font.size = Pt(8.5)
    set_widths(t, [3.4, 2.6, 1.8, 1.5, 7.7])
    doc.add_paragraph().paragraph_format.space_after = Pt(2)

# ---------- Fichas ----------
doc.add_page_break()
doc.add_heading('Fichas de lectura', level=1)
P('Formato de la «Ficha de lectura por investigación» de la guía del grupo: referencia, objetivo, participantes y contexto, método, resultados útiles, limitaciones y uso en el trabajo. Se añade una cita textual breve cuando se ha leído el texto completo.', size=9.5)
CAMPOS = [('referencia_apa7', 'Referencia (APA 7)'), ('enlace', 'Texto completo / DOI'), ('pdf_rel', 'PDF en la carpeta'), ('acceso', 'Acceso'), ('objetivo', 'Objetivo'),
          ('participantes_contexto', 'Participantes y contexto'), ('metodo', 'Método'), ('resultados_utiles', 'Resultados útiles'),
          ('limitaciones', 'Limitaciones'), ('cita_textual', 'Cita textual'), ('uso_en_el_trabajo', 'Uso en el trabajo'), ('verificacion', 'Verificación')]
for code, name in BLOQUES:
    items = [x for x in fichas if str(x.get('bloque', '')).startswith(code)]
    if not items: continue
    doc.add_heading(f'{code}. {name}', level=2)
    for x in items:
        doc.add_heading(short_ref(x), level=3)
        t = doc.add_table(rows=0, cols=2); t.style = 'Table Grid'
        for key, label in CAMPOS:
            val = (x.get(key) or '').strip()
            if not val: continue
            row = t.add_row().cells
            row[0].text = ''; r = row[0].paragraphs[0].add_run(label); r.bold = True; r.font.size = Pt(8.5); shade(row[0], 'EEF2F8')
            row[1].text = ''; add_rich(row[1].paragraphs[0], val, size=8.5)
        set_widths(t, [3.4, 13.6])
        doc.add_paragraph().paragraph_format.space_after = Pt(2)

for sec in EXTRA.get('final', []):
    doc.add_heading(sec['h'], level=1)
    for ptxt in sec.get('p', []):
        p = doc.add_paragraph(); add_rich(p, ptxt)
    for tb in sec.get('tablas', []):
        t = doc.add_table(rows=1, cols=len(tb['cols'])); t.style = 'Table Grid'
        for i, h in enumerate(tb['cols']):
            c = t.rows[0].cells[i]; c.text = ''; r = c.paragraphs[0].add_run(h); r.bold = True; r.font.size = Pt(8.5); shade(c, 'D9E2F3')
        cabecera(t)
        for rowv in tb['rows']:
            cells = t.add_row().cells
            for i, v in enumerate(rowv):
                cells[i].text = ''; add_rich(cells[i].paragraphs[0], v, size=8.5)
        set_widths(t, tb['widths'])
        doc.add_paragraph().paragraph_format.space_after = Pt(2)

doc.save(OUT); print('saved', OUT, '| fichas:', len(fichas))
