# -*- coding: utf-8 -*-
"""Genera (1) el informe de la entrevista en versión preparada y (2) el guion para imprimir.
Uso: python3 build_entrevista.py [carpeta_salida]"""
import sys, os, importlib.util
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

HERE = os.path.dirname(os.path.abspath(__file__))
OUTDIR = sys.argv[1] if len(sys.argv) > 1 else '/home/user/fj/Salima_entrevista'
spec = importlib.util.spec_from_file_location('te', os.path.join(HERE, 'texto_entrevista.py'))
T = importlib.util.module_from_spec(spec); spec.loader.exec_module(T)

def nuevo_doc():
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
    fp = doc.sections[0].footer.paragraphs[0]; fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = fp.add_run()
    for tag, txt in [('begin', None), (None, 'PAGE'), ('end', None)]:
        if tag:
            e = OxmlElement('w:fldChar'); e.set(qn('w:fldCharType'), tag); r._r.append(e)
        else:
            e = OxmlElement('w:instrText'); e.set(qn('xml:space'), 'preserve'); e.text = txt; r._r.append(e)
    r.font.size = Pt(9)
    return doc

class W:
    def __init__(self, doc): self.doc = doc
    def runs(self, p, text, size=None, bold=False, italic=False):
        for i, part in enumerate(text.split('*')):
            if not part: continue
            r = p.add_run(part); r.italic = italic or (i % 2 == 1); r.bold = bold
            if size: r.font.size = Pt(size)
    def P(self, text, align=None, size=None, after=None, bold=False, italic=False, indent=None, spacing=None):
        p = self.doc.add_paragraph(); self.runs(p, text, size, bold, italic)
        if align == 'center': p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if align == 'left': p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        if after is not None: p.paragraph_format.space_after = Pt(after)
        if indent: p.paragraph_format.left_indent = Cm(indent)
        if spacing: p.paragraph_format.line_spacing = spacing
        return p
    def H(self, text, lvl=1): return self.doc.add_heading(text, level=lvl)
    def shade(self, cell, hex_color):
        tcPr = cell._tc.get_or_add_tcPr(); shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto'); shd.set(qn('w:fill'), hex_color); tcPr.append(shd)
    def TABLA(self, cols, rows, widths, size=9.5):
        t = self.doc.add_table(rows=1, cols=len(cols)); t.style = 'Table Grid'; t.alignment = WD_TABLE_ALIGNMENT.CENTER
        for i, h in enumerate(cols):
            c = t.rows[0].cells[i]; c.text = ''; pr = c.paragraphs[0]; r = pr.add_run(h); r.bold = True; r.font.size = Pt(size)
            pr.alignment = WD_ALIGN_PARAGRAPH.LEFT; pr.paragraph_format.line_spacing = 1.0; pr.paragraph_format.space_after = Pt(2); self.shade(c, 'E7E6E6')
        trPr = t.rows[0]._tr.get_or_add_trPr(); el = OxmlElement('w:tblHeader'); el.set(qn('w:val'), 'true'); trPr.append(el)
        for rowv in rows:
            cells = t.add_row().cells
            for i, v in enumerate(rowv):
                cells[i].text = ''; pr = cells[i].paragraphs[0]; self.runs(pr, v, size=size)
                pr.alignment = WD_ALIGN_PARAGRAPH.LEFT; pr.paragraph_format.line_spacing = 1.0; pr.paragraph_format.space_after = Pt(2)
        t.autofit = False
        grid = t._tbl.find(qn('w:tblGrid'))
        for i, gc in enumerate(grid.findall(qn('w:gridCol'))): gc.set(qn('w:w'), str(int(widths[i] * 567)))
        for row in t.rows:
            for i, w in enumerate(widths): row.cells[i].width = Cm(w)
        self.doc.add_paragraph().paragraph_format.space_after = Pt(4)
    def REF(self, text):
        p = self.doc.add_paragraph(); pf = p.paragraph_format
        pf.left_indent = Cm(1.25); pf.first_line_indent = Cm(-1.25); pf.space_after = Pt(6)
        pf.line_spacing = 1.15; p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        self.runs(p, text, size=10.5)

# ---------------- 1. Informe preparado ----------------
doc = nuevo_doc(); w = W(doc)
w.P('INFORME DE LA ENTREVISTA', bold=True, size=13, align='center', after=4)
w.P('Título: ' + T.TITULO, bold=True, size=12, align='center', after=4)
w.P('Salima López Rami · Máster Universitario en Diversidad Cultural, Universidad de Granada, Campus de Melilla', align='center', size=10, after=2)
w.P('La Investigación en Contextos Multiculturales e Iniciación a la Investigación · [fecha de entrega]', align='center', size=10, after=4)
w.P('Versión preparada: los corchetes marcan lo que se completa cuando se haya hecho la entrevista. El informe final y los anexos se redactan a partir de la transcripción.', italic=True, size=9.5, align='center', after=14)

w.H('Justificación metodológica', 1)
for p in T.JUSTIFICACION: w.P(p)

w.H('Objetivos', 1)
w.P('Objetivo general. ' + T.OBJ_GENERAL)
w.P('Objetivos específicos:', after=4)
for i, o in enumerate(T.OBJ_ESPECIFICOS, 1): w.P(f'{i}. {o}', indent=0.75, after=4, spacing=1.3)

w.H('Participantes', 1)
for p in T.PARTICIPANTES: w.P(p)

w.H('Instrumento de investigación', 1)
for p in T.INSTRUMENTO: w.P(p)
w.TABLA(T.GUION['cols'], T.GUION['rows'], T.GUION['widths'])

w.H('Análisis de datos', 1)
for p in T.ANALISIS: w.P(p)
w.P('Tabla 1. Clasificación de las etiquetas', italic=True, size=10, after=4, align='left')
ded = '\n'.join(c for c, _ in T.CODIGOS_DEDUCTIVOS)
w.TABLA(['Etiquetas deductivas', 'Etiquetas inductivas'], [[ded, T.CODIGOS_INDUCTIVOS_PREVISTOS]], [8.2, 8.2])
w.H('Definición de códigos deductivos', 2)
w.TABLA(['Código', 'Definición'], T.CODIGOS_DEDUCTIVOS, [5.4, 11.0])
w.H('Definición de códigos inductivos', 2)
w.P('[Se definen tras la codificación, con el mismo formato: nombre del código y descripción de una o dos líneas.]', italic=True)
w.H('Familias de códigos (epígrafes del informe final)', 2)
w.TABLA(['Familia', 'Códigos que agrupa'], T.FAMILIAS, [10.0, 6.4])

w.H('Criterios de calidad empleados para asegurar la calidad de la investigación', 1)
for p in T.CALIDAD: w.P(p)

w.H('Informe final', 1)
w.P(T.INFORME_FINAL_NOTA, italic=True)

w.H('Referencias bibliográficas', 1)
for r in T.REFERENCIAS: w.REF(r)

w.H('Anexos', 1)
w.P(T.ANEXOS_NOTA, italic=True)
w.H('Consentimiento informado (texto leído antes de grabar)', 2)
w.P(T.CONSENTIMIENTO, italic=True, size=10.5)

os.makedirs(OUTDIR, exist_ok=True)
out1 = os.path.join(OUTDIR, 'Informe_entrevista_Salima_version_preparada.docx'); doc.save(out1); print('guardado:', out1)

# ---------------- 2. Guion para imprimir ----------------
doc = nuevo_doc(); w = W(doc)
doc.styles['Normal'].paragraph_format.line_spacing = 1.2
w.P('Guion de la entrevista', bold=True, size=14, align='center', after=2)
w.P('Cómo responde un centro educativo de Melilla a la diversidad de su alumnado', size=11, align='center', after=2)
w.P('Entrevista semiestructurada, individual, 30-45 minutos, grabada en audio', italic=True, size=10, align='center', after=12)
w.H('Antes de empezar: consentimiento', 2)
w.P(T.CONSENTIMIENTO, italic=True, size=10.5)
w.H('Preguntas', 2)
w.P('Las preguntas son una guía, no un cuestionario. Se pueden cambiar de orden, saltar las que ya se hayan respondido y añadir «¿me pones un ejemplo?» o «¿qué sentiste?» en cualquier momento.', size=10.5, after=6)
w.TABLA(T.GUION['cols'], T.GUION['rows'], T.GUION['widths'], size=10)
w.H('Consejos prácticos', 2)
for c in T.CONSEJOS: w.P(c, size=10.5, after=4, spacing=1.2)
out2 = os.path.join(OUTDIR, 'Guion_entrevista_Salima.docx'); doc.save(out2); print('guardado:', out2)
