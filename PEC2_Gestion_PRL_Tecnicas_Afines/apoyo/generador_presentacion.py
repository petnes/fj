# -*- coding: utf-8 -*-
"""PEC 02 (Gestión de la PRL y Técnicas Afines, MUPRL UNED): presentación corregida con citas APA 7 y guion en las notas."""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree
import copy, re

NAVY = RGBColor(0x0F, 0x25, 0x45); BLUE = RGBColor(0x2E, 0x75, 0xB6); LIGHT = RGBColor(0xEE, 0xF2, 0xF7)
WHITE = RGBColor(0xFF, 0xFF, 0xFF); TEXT = RGBColor(0x22, 0x22, 0x22); GRAY = RGBColor(0x66, 0x66, 0x66)
ORANGE = RGBColor(0xC5, 0x5A, 0x11); SKY = RGBColor(0xB4, 0xC8, 0xDC)
FONT = 'Calibri'
W, H = Inches(13.333), Inches(7.5)

prs = Presentation(); prs.slide_width = W; prs.slide_height = H
BLANK = prs.slide_layouts[6]

# ---------- utilidades ----------
def rect(slide, x, y, w, h, fill, line=None, shape=MSO_SHAPE.RECTANGLE):
    s = slide.shapes.add_shape(shape, x, y, w, h)
    s.fill.solid(); s.fill.fore_color.rgb = fill
    if line is None: s.line.fill.background()
    else: s.line.color.rgb = line; s.line.width = Pt(0.75)
    s.shadow.inherit = False
    return s

def set_bullet(par, char='•', indent=Inches(0.22)):
    pPr = par._p.get_or_add_pPr()
    pPr.set('marL', str(int(indent))); pPr.set('indent', str(-int(indent)))
    for tag in ('a:buNone', 'a:buChar', 'a:buAutoNum'):
        for e in pPr.findall(qn(tag)): pPr.remove(e)
    bu = etree.SubElement(pPr, qn('a:buChar')); bu.set('char', char)

def add_runs(par, text, size, color=TEXT, bold=False):
    """Texto con marcas: **negrita**, [[placeholder]] en naranja."""
    for tok in re.split(r'(\*\*.*?\*\*|\[\[.*?\]\])', text):
        if not tok: continue
        r = par.add_run()
        if tok.startswith('**'):
            r.text = tok[2:-2]; r.font.bold = True; r.font.color.rgb = color
        elif tok.startswith('[['):
            r.text = '[' + tok[2:-2] + ']'; r.font.bold = True; r.font.color.rgb = ORANGE
        else:
            r.text = tok; r.font.bold = bold; r.font.color.rgb = color
        r.font.size = Pt(size); r.font.name = FONT

def textbox(slide, x, y, w, h, items, size=14, color=TEXT, bullets=True, space_after=6, anchor=MSO_ANCHOR.TOP, align=PP_ALIGN.LEFT, line_spacing=1.05):
    tb = slide.shapes.add_textbox(x, y, w, h); tf = tb.text_frame; tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.08); tf.margin_top = tf.margin_bottom = Inches(0.05); tf.vertical_anchor = anchor
    for i, it in enumerate(items):
        if isinstance(it, tuple): txt, opts = it
        else: txt, opts = it, {}
        par = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        par.alignment = align; par.space_after = Pt(opts.get('after', space_after)); par.line_spacing = line_spacing
        add_runs(par, txt, opts.get('size', size), opts.get('color', color), opts.get('bold', False))
        if opts.get('bullet', bullets): set_bullet(par)
    return tb

def header(slide, title, num):
    rect(slide, 0, 0, W, Inches(0.95), NAVY)
    textbox(slide, Inches(0.45), Inches(0.12), Inches(10.5), Inches(0.75), [(title, {'size': 24, 'bold': True, 'color': WHITE, 'bullet': False})], anchor=MSO_ANCHOR.MIDDLE)
    textbox(slide, Inches(10.2), Inches(0.2), Inches(2.9), Inches(0.6), [('PEC 02 · MUPRL UNED', {'size': 12, 'color': SKY, 'bullet': False})], align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
    rect(slide, 0, Inches(7.1), W, Inches(0.4), LIGHT)
    textbox(slide, Inches(0.45), Inches(7.12), Inches(9), Inches(0.35), [('Gestión de la PRL y Técnicas Afines · Milagros Moncada Roa', {'size': 10, 'color': GRAY, 'bullet': False})], anchor=MSO_ANCHOR.MIDDLE)
    textbox(slide, Inches(11.2), Inches(7.12), Inches(1.9), Inches(0.35), [(f'Diapositiva {num}', {'size': 10, 'color': GRAY, 'bullet': False})], align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

def card(slide, x, y, w, h, title, items, size=13):
    rect(slide, x, y, w, h, WHITE, line=SKY, shape=MSO_SHAPE.ROUNDED_RECTANGLE).adjustments[0] = 0.04
    textbox(slide, x + Inches(0.15), y + Inches(0.12), w - Inches(0.3), Inches(0.5), [(title, {'size': 15, 'bold': True, 'color': NAVY, 'bullet': False})])
    textbox(slide, x + Inches(0.15), y + Inches(0.62), w - Inches(0.3), h - Inches(0.75), items, size=size)

def table(slide, x, y, w, h, widths, header_row, rows, size=11, header_size=12):
    shp = slide.shapes.add_table(len(rows) + 1, len(header_row), x, y, w, h); t = shp.table
    tot = sum(widths)
    for i, cw in enumerate(widths): t.columns[i].width = int(w * cw / tot)
    def fill_cell(c, text, sz, bold=False, color=TEXT, bg=None):
        c.text = ''; tf = c.text_frame; tf.word_wrap = True
        c.margin_left = c.margin_right = Inches(0.06); c.margin_top = c.margin_bottom = Inches(0.03)
        par = tf.paragraphs[0]; add_runs(par, text, sz, color, bold); par.line_spacing = 1.0
        if bg is not None: c.fill.solid(); c.fill.fore_color.rgb = bg
    for j, htxt in enumerate(header_row): fill_cell(t.cell(0, j), htxt, header_size, True, WHITE, NAVY)
    for i, row in enumerate(rows, start=1):
        for j, txt in enumerate(row): fill_cell(t.cell(i, j), txt, size, False, TEXT, WHITE if i % 2 else LIGHT)
    # quitar estilo de tabla por defecto (bandas)
    tblPr = t._tbl.tblPr; tblPr.set('bandRow', '0'); tblPr.set('firstRow', '1')
    t.rows[0].height = Inches(0.5)
    for i in range(1, len(rows) + 1): t.rows[i].height = int((h - Inches(0.5)) / len(rows))
    return t

def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text

NOTAS = []
def new_slide(title=None, num=None):
    s = prs.slides.add_slide(BLANK)
    if title: header(s, title, num)
    return s

# ================= 1. PORTADA =================
s = new_slide()
rect(s, 0, 0, W, H, NAVY)
rect(s, Inches(0.7), Inches(1.35), Inches(0.12), Inches(2.3), BLUE)
textbox(s, Inches(1.0), Inches(1.2), Inches(11.5), Inches(1.6), [('Supervisión inteligente y cumplimiento del Reglamento de IA en la construcción', {'size': 32, 'bold': True, 'color': WHITE, 'bullet': False})], anchor=MSO_ANCHOR.MIDDLE)
textbox(s, Inches(1.0), Inches(2.85), Inches(11.5), Inches(0.9), [('Análisis del marco regulador europeo (Reglamento (UE) 2024/1689) y de la prevención de riesgos laborales en obras', {'size': 18, 'color': SKY, 'bullet': False})])
textbox(s, Inches(1.0), Inches(4.3), Inches(11.5), Inches(2.4), [
    ('PEC 02. Tema: implementación de la IA en labores de prevención en el sector de la construcción', {'size': 15, 'color': WHITE, 'bullet': False, 'after': 14}),
    ('Máster Universitario en Prevención de Riesgos Laborales (MUPRL), UNED', {'size': 14, 'color': WHITE, 'bullet': False, 'after': 4}),
    ('Asignatura: Gestión de la PRL y Técnicas Afines', {'size': 14, 'color': WHITE, 'bullet': False, 'after': 4}),
    ('Autora: Milagros Moncada Roa', {'size': 14, 'color': WHITE, 'bullet': False, 'after': 4}),
    ('Curso académico 2026-2027', {'size': 14, 'color': WHITE, 'bullet': False}),
])
notes(s, "Hola. Soy Milagros Moncada Roa y esta es mi PEC 2 de Gestión de la Prevención de Riesgos Laborales y Técnicas Afines. He elegido el tema de la implementación de la inteligencia artificial en las labores de prevención en el sector de la construcción, y lo he enfocado como una pregunta de cumplimiento: qué puede aportar la IA a la vigilancia preventiva en obra y bajo qué condiciones legales puede usarse, ahora que el Reglamento europeo de Inteligencia Artificial ya está en aplicación.")

# ================= 2. ÍNDICE =================
s = new_slide('1. Índice', 2)
items = ['Introducción y justificación', 'Antecedentes I: marco de prevención de riesgos laborales en obras', 'Antecedentes II: marco europeo de la inteligencia artificial y protección de datos', 'Metodología', 'Resultados I: aplicaciones de IA con evidencia publicada', 'Resultados II: encaje normativo de cada aplicación', 'Discusión: la supervisión humana como condición', 'Conclusiones y recomendaciones', 'Bibliografía (APA 7.ª edición)', 'Anexo: declaración de uso de inteligencia artificial']
textbox(s, Inches(0.8), Inches(1.3), Inches(6.2), Inches(5.6), [(f'{i+1}. {t}', {'bullet': False, 'after': 10}) for i, t in enumerate(items[:5])], size=17)
textbox(s, Inches(7.0), Inches(1.3), Inches(6.0), Inches(5.6), [(f'{i+6}. {t}', {'bullet': False, 'after': 10}) for i, t in enumerate(items[5:])], size=17)
notes(s, "La presentación sigue la estructura recomendada para la PEC: introducción y justificación, dos bloques de antecedentes, uno sobre el marco de prevención en obras y otro sobre el Reglamento de Inteligencia Artificial y la protección de datos, la metodología, dos bloques de resultados, la discusión, las conclusiones, la bibliografía en APA séptima edición y el anexo con la declaración de uso de inteligencia artificial que exige el enunciado.")

# ================= 3. INTRODUCCIÓN =================
s = new_slide('2. Introducción y justificación', 3)
card(s, Inches(0.5), Inches(1.25), Inches(6.1), Inches(5.65), 'Por qué la construcción', [
    'Índice de incidencia de **5.530** accidentes con baja por cada 100.000 personas afiliadas, frente a **2.542** en el conjunto de los sectores (INSST, 2026).',
    'Índice de incidencia mortal de **11,8** frente a 2,8 en el total; **173 accidentes mortales en jornada** en los doce meses de octubre de 2025 a septiembre de 2026 (INSST, 2026).',
    'Caídas a distinto nivel, atropellos por maquinaria y trabajo sin equipos de protección individual concentran la gravedad (Real Decreto 1627/1997, 1997).',
    'La obra cambia cada día: la inspección periódica llega tarde a los riesgos que aparecen entre dos visitas.',
], size=14)
card(s, Inches(6.75), Inches(1.25), Inches(6.1), Inches(5.65), 'Pregunta y objetivos', [
    ('**Pregunta:** ¿puede la IA reforzar la vigilancia preventiva en obra sin vulnerar los derechos de la plantilla ni el Reglamento de Inteligencia Artificial?', {'bullet': False, 'after': 10}),
    'Objetivo 1: identificar aplicaciones de IA con evidencia publicada para los riesgos principales de la obra.',
    'Objetivo 2: situar cada aplicación en el marco español de PRL y en el Reglamento (UE) 2024/1689.',
    'Objetivo 3: proponer condiciones de implantación compatibles con la supervisión humana y la protección de datos.',
], size=14)
notes(s, "Empiezo por la justificación con datos oficiales del INSST. Entre octubre de 2025 y septiembre de 2026 la construcción tuvo un índice de incidencia de cinco mil quinientos treinta accidentes con baja por cada cien mil personas afiliadas, más del doble de la media, y un índice mortal de once coma ocho frente a dos coma ocho: ciento setenta y tres muertes en jornada en un año. Caídas, atropellos y trabajo sin protección explican esa gravedad, y como la obra cambia cada día, la inspección periódica llega tarde. Mi pregunta es si la inteligencia artificial puede vigilar de forma continua sin vulnerar los derechos de la plantilla ni el Reglamento europeo, con tres objetivos: identificar aplicaciones con evidencia, situarlas en el marco legal y proponer condiciones de implantación.")

# ================= 4. ANTECEDENTES I =================
s = new_slide('3. Antecedentes I: marco de prevención en obras', 4)
card(s, Inches(0.5), Inches(1.25), Inches(6.1), Inches(5.65), 'Ley 31/1995 y Real Decreto 1627/1997', [
    'Evaluación de riesgos y planificación preventiva (arts. 15 y 16); el art. 15.1.e) obliga a tener en cuenta la evolución de la técnica (Ley 31/1995, 1995).',
    'Presencia de recursos preventivos en obra: disposición adicional decimocuarta de la LPRL y disposición adicional única del RD 1627/1997 (Real Decreto 1627/1997, 1997).',
    'Estudio y plan de seguridad y salud, coordinador en fase de ejecución y libro de incidencias como instrumentos de control de la obra (Real Decreto 1627/1997, 1997).',
    'Consulta y participación: Comité de Seguridad y Salud y delegados de prevención (arts. 33 a 39 de la Ley 31/1995).',
], size=13)
card(s, Inches(6.75), Inches(1.25), Inches(6.1), Inches(5.65), 'Control empresarial y datos de la plantilla', [
    'El empresario puede vigilar el cumplimiento de las obligaciones laborales con respeto a la dignidad (art. 20.3 del Estatuto de los Trabajadores).',
    'La representación legal debe ser informada de los parámetros, reglas e instrucciones de los algoritmos o sistemas de IA que afecten a las condiciones de trabajo (art. 64.4.d) (Real Decreto Legislativo 2/2015, 2015).',
    'Videovigilancia (art. 89) y geolocalización (art. 90) en el trabajo: información previa, expresa y clara a la plantilla y a sus representantes (Ley Orgánica 3/2018, 2018).',
], size=13)
notes(s, "El marco español ya contiene lo esencial. La Ley de Prevención exige evaluar y planificar, y en su artículo quince manda tener en cuenta la evolución de la técnica. En obras, la Ley y el Real Decreto mil seiscientos veintisiete regulan la presencia de los recursos preventivos, el plan de seguridad y salud, el coordinador y el libro de incidencias. El Estatuto de los Trabajadores permite el control empresarial respetando la dignidad y obliga a informar a la representación legal de los algoritmos que afectan a las condiciones de trabajo. Y la Ley Orgánica de Protección de Datos regula la videovigilancia y la geolocalización en el trabajo con información previa a la plantilla.")

# ================= 5. ANTECEDENTES II =================
s = new_slide('3. Antecedentes II: Reglamento de IA y protección de datos', 5)
card(s, Inches(0.5), Inches(1.25), Inches(6.1), Inches(5.65), 'Reglamento (UE) 2024/1689', [
    'Enfoque por niveles de riesgo. **Prohibido** usar IA para inferir las emociones de una persona en el lugar de trabajo, salvo por motivos médicos o de seguridad (art. 5.1.f).',
    '**Alto riesgo:** sistemas destinados a supervisar y evaluar el rendimiento y el comportamiento de las personas en la relación laboral (anexo III, punto 4.b).',
    'Supervisión humana efectiva (art. 14). El empleador, como responsable del despliegue, usa el sistema según las instrucciones, encomienda la supervisión a personas competentes e informa previamente a trabajadores y representantes (art. 26, apartados 1, 2 y 7).',
    'Alfabetización en IA del personal que opera los sistemas (art. 4).',
], size=13.5)
card(s, Inches(6.75), Inches(1.25), Inches(6.1), Inches(5.65), 'Calendario y protección de datos', [
    'Prohibiciones y alfabetización aplicables desde el 2 de febrero de 2025; régimen general desde el 2 de agosto de 2026 (art. 113) (Reglamento (UE) 2024/1689, 2024).',
    'Obligaciones de los sistemas de alto riesgo del anexo III aplazadas al **2 de diciembre de 2027** por el Ómnibus digital sobre IA (Reglamento (UE) 2026/1744, 2026).',
    'RGPD: evaluación de impacto cuando hay observación sistemática de personas (art. 35) y límites a las decisiones basadas únicamente en tratamientos automatizados (art. 22) (Reglamento (UE) 2016/679, 2016).',
], size=13.5)
notes(s, "El Reglamento europeo de Inteligencia Artificial funciona por niveles de riesgo. Hay una prohibición que afecta a nuestro caso: inferir las emociones de una persona en el lugar de trabajo, salvo por motivos médicos o de seguridad. Y una categoría de alto riesgo que afecta a la vigilancia en obra: los sistemas destinados a supervisar y evaluar el rendimiento y el comportamiento de las personas en la relación laboral. Para esos sistemas el empleador es el responsable del despliegue: debe usarlos según las instrucciones, encomendar la supervisión humana a personas competentes e informar a los trabajadores y a sus representantes antes de ponerlos en servicio. Las prohibiciones se aplican desde febrero de dos mil veinticinco, pero las obligaciones de alto riesgo del anexo tres se han aplazado al dos de diciembre de dos mil veintisiete por el Ómnibus digital de julio de este año. Y el Reglamento de Protección de Datos exige una evaluación de impacto cuando hay observación sistemática de personas.")

# ================= 6. METODOLOGÍA =================
s = new_slide('4. Metodología', 6)
card(s, Inches(0.5), Inches(1.25), Inches(6.1), Inches(5.65), 'Diseño y fuentes', [
    'Revisión documental de carácter cualitativo, sin trabajo de campo.',
    'Normativa en texto consolidado (BOE y EUR-Lex) y estadística oficial de siniestralidad (INSST, 2026).',
    'Informes de la Agencia Europea para la Seguridad y la Salud en el Trabajo sobre IA en la gestión de trabajadores (2022a, 2022b).',
    'Artículos revisados por pares en Automation in Construction y Safety Science (Fang et al., 2018; Nath et al., 2020; Teizer et al., 2010; Zhang et al., 2015).',
], size=13)
card(s, Inches(6.75), Inches(1.25), Inches(6.1), Inches(5.65), 'Criterios, análisis y limitaciones', [
    'Criterios de selección: aplicación probada en obra real o simulada, evidencia publicada y relación directa con caídas, atropellos o uso de EPI.',
    'Análisis mediante una matriz que cruza cada tecnología con el riesgo que trata, la evidencia disponible y los requisitos jurídicos aplicables.',
    'Limitaciones: literatura mayoritariamente anglosajona y anterior al Reglamento de IA; sin datos de implantación en pymes españolas; el calendario legal sigue en evolución.',
], size=13)
notes(s, "La metodología es una revisión documental cualitativa con tres tipos de fuentes: la normativa en texto consolidado del BOE y de EUR-Lex, la estadística oficial del INSST y la literatura técnica, es decir, dos informes de la Agencia Europea para la Seguridad y la Salud en el Trabajo y cuatro artículos revisados por pares. Se seleccionaron aplicaciones probadas en obra, con evidencia publicada y relacionadas con caídas, atropellos o uso de EPI, y se analizaron en una matriz que cruza tecnología, riesgo, evidencia y requisitos legales. Las limitaciones: literatura anglosajona y anterior al Reglamento, y sin datos de pymes españolas.")

# ================= 7. RESULTADOS I =================
s = new_slide('5. Resultados I: aplicaciones de IA con evidencia publicada', 7)
table(s, Inches(0.45), Inches(1.2), Inches(12.45), Inches(4.9), [2.4, 2.0, 4.6, 3.2],
      ['Tecnología', 'Riesgo que trata', 'Evidencia publicada', 'Limitaciones observadas'], [
    ['Visión artificial sobre las cámaras de la obra: detección de casco, chaleco y arnés', 'Trabajo sin EPI; caídas y golpes', 'Detección en tiempo real con redes neuronales entrenadas con imágenes de obra (Nath et al., 2020). Detección de la ausencia de casco en vídeo de vigilancia a larga distancia (Fang et al., 2018).', 'Oclusiones, distancia, luz y polvo; falsos positivos; trata imágenes de personas identificables.'],
    ['Sensores de proximidad entre maquinaria y personas (radiofrecuencia)', 'Atropellos y aplastamientos por maquinaria móvil', 'Alerta en tiempo real al operador y al trabajador a pie, validada en pruebas de campo (Teizer et al., 2010).', 'Alcance variable; fatiga por exceso de alarmas; exige que todas las personas lleven el dispositivo.'],
    ['Comprobación automática de reglas de seguridad sobre el modelo BIM', 'Caídas a distinto nivel, desde la fase de planificación', 'Identificación automática de huecos y bordes sin protección y propuesta de medidas, probada en dos casos de estudio (Zhang et al., 2015).', 'Depende de la calidad del modelo y de su actualización durante la ejecución.'],
], size=11.5)
textbox(s, Inches(0.5), Inches(6.2), Inches(12.3), Inches(0.8), [('**Resultado 1.** La evidencia más sólida se concentra en tres usos: uso de EPI, proximidad a maquinaria y riesgos de caída anticipados en el modelo. Ninguno sustituye al recurso preventivo; todos le dan información antes.', {'bullet': False})], size=13)
notes(s, "Primer resultado: tres aplicaciones con evidencia. La visión artificial sobre las cámaras de la obra para detectar casco, chaleco y arnés: Nath y colaboradores demostraron en dos mil veinte la detección en tiempo real con redes neuronales, y Fang y colaboradores, en dos mil dieciocho, la detección de la ausencia de casco en vídeo de vigilancia a larga distancia; sus límites son las oclusiones, la luz, el polvo, los falsos positivos y, sobre todo, que tratan imágenes de personas identificables. Los sensores de proximidad entre maquinaria y personas, validados en campo por Teizer y colaboradores en dos mil diez. Y la comprobación automática de reglas de seguridad sobre el modelo BIM, que Zhang y colaboradores aplicaron en dos mil quince a los riesgos de caída desde la planificación. Ninguna sustituye al recurso preventivo: le dan información antes.")

# ================= 8. RESULTADOS II =================
s = new_slide('6. Resultados II: encaje normativo de cada aplicación', 8)
table(s, Inches(0.45), Inches(1.2), Inches(12.45), Inches(4.9), [2.2, 1.9, 3.6, 4.5],
      ['Aplicación', 'Datos personales', 'Clasificación en el Reglamento de IA', 'Obligaciones principales'], [
    ['Visión artificial sobre personas (uso de EPI)', 'Sí: imágenes de trabajadores identificables', 'Alto riesgo cuando sirve para supervisar o evaluar el comportamiento (anexo III, 4.b). Nunca inferencia de emociones (art. 5.1.f).', 'Supervisión humana competente e información previa a plantilla y representantes (art. 26); evaluación de impacto (art. 35 RGPD); art. 89 LOPDGDD; art. 64.4.d ET; finalidad preventiva, no disciplinaria.'],
    ['Proximidad entre maquinaria y personas', 'Mínimos: dispositivos sin identidad', 'Fuera del anexo III si no evalúa a la persona; posible componente de seguridad de la máquina (anexo I).', 'Instrucciones del fabricante; integración en el plan de seguridad y salud; formación (art. 19 LPRL); mantenimiento y registro de alarmas.'],
    ['Reglas automáticas sobre el modelo BIM', 'No', 'Riesgo mínimo: no trata personas ni decide sobre ellas.', 'Validación por el coordinador de seguridad y salud; trazabilidad de las medidas en el plan de seguridad y salud.'],
], size=11.5)
textbox(s, Inches(0.5), Inches(6.2), Inches(12.3), Inches(0.8), [('**Resultado 2.** El encaje jurídico depende menos de la tecnología que de la finalidad y de los datos que trata: la misma cámara es riesgo mínimo si cuenta cascos y alto riesgo si evalúa a personas.', {'bullet': False})], size=13)
notes(s, "Segundo resultado, la matriz de encaje normativo, que es la aportación central. La visión artificial sobre personas trata imágenes identificables y, si sirve para supervisar o evaluar el comportamiento, es de alto riesgo; nunca puede inferir emociones. Eso obliga a supervisión humana competente, información previa a la plantilla y a sus representantes, evaluación de impacto de protección de datos y a cumplir la Ley Orgánica y el Estatuto. Y la finalidad debe ser preventiva, no disciplinaria. Los sensores de proximidad tratan datos mínimos y no evalúan a la persona: quedan fuera del anexo tres, aunque pueden ser un componente de seguridad de la máquina. Las reglas sobre el modelo BIM no tratan personas: riesgo mínimo. La misma cámara es riesgo mínimo si cuenta cascos y alto riesgo si evalúa a personas.")

# ================= 9. DISCUSIÓN =================
s = new_slide('7. Discusión: la supervisión humana como condición', 9)
textbox(s, Inches(0.6), Inches(1.3), Inches(12.1), Inches(5.6), [
    '**El sistema propone y la persona decide.** El recurso preventivo y el coordinador conservan la decisión y la acción correctora; el algoritmo no sanciona ni sustituye la vigilancia (arts. 14 y 26.2 del Reglamento (UE) 2024/1689).',
    '**Riesgos documentados de la gestión algorítmica:** intensificación del trabajo, pérdida de autonomía, sensación de vigilancia y sesgos en los datos (Agencia Europea para la Seguridad y la Salud en el Trabajo, 2022a).',
    '**Medidas de prevención:** participación de la plantilla desde el diseño, transparencia sobre qué se mide y para qué, minimización de datos y revisión periódica de los falsos positivos (Agencia Europea para la Seguridad y la Salud en el Trabajo, 2022b).',
    '**Canal adecuado:** consulta en el Comité de Seguridad y Salud (arts. 33 y 38 de la Ley 31/1995) e información a la representación legal (art. 64.4.d ET) antes de la puesta en servicio.',
    '**Viabilidad:** coste y mantenimiento en pymes, condiciones de obra que degradan la detección y necesidad de personal con alfabetización en IA (art. 4). La tecnología es la parte fácil.',
], size=14, space_after=12)
notes(s, "En la discusión destaco la supervisión humana como condición. El sistema propone y la persona decide: el recurso preventivo y el coordinador conservan la decisión, y el algoritmo no sanciona. La Agencia Europea ha documentado los riesgos de la gestión algorítmica, intensificación del trabajo, pérdida de autonomía, sensación de vigilancia y sesgos, y propone medidas que encajan con nuestro sistema: participación desde el diseño, transparencia sobre qué se mide y para qué, minimización de datos y revisión de los falsos positivos. El canal ya existe: el Comité de Seguridad y Salud y la información a la representación legal. Queda la viabilidad en pymes, las condiciones de obra y la formación del personal. La tecnología es la parte fácil.")

# ================= 10. CONCLUSIONES =================
s = new_slide('8. Conclusiones y recomendaciones', 10)
textbox(s, Inches(0.6), Inches(1.3), Inches(12.1), Inches(5.6), [
    '**1.** La IA aporta vigilancia continua en tres frentes con evidencia (uso de EPI, proximidad a maquinaria y caídas anticipadas en BIM) en un sector cuyos índices de incidencia duplican con creces la media (INSST, 2026).',
    '**2.** El marco español ya exige lo esencial: evaluación, planificación, recursos preventivos y control respetuoso con la dignidad. El Reglamento de IA añade la clasificación por riesgo, la supervisión humana y la información a la plantilla.',
    '**3.** Lo prohibido es claro (inferir emociones en el trabajo) y lo de alto riesgo también (supervisar y evaluar el comportamiento), con obligaciones plenas desde el 2 de diciembre de 2027.',
    '**4. Recomendación:** implantar con finalidad exclusivamente preventiva, con el recurso preventivo al mando del sistema, con evaluación de impacto y consulta previa, y medir el resultado en indicadores de siniestralidad y no en sanciones.',
], size=14.5, space_after=14, bullets=False)
notes(s, "Cuatro conclusiones. Primera: la inteligencia artificial aporta vigilancia continua en tres frentes con evidencia, EPI, proximidad a maquinaria y caídas anticipadas en BIM, en un sector que duplica con creces la siniestralidad media. Segunda: el marco español ya exige lo esencial y el Reglamento de IA añade la clasificación por riesgo, la supervisión humana y la información a la plantilla. Tercera: lo prohibido es claro y lo de alto riesgo también, con obligaciones plenas desde el dos de diciembre de dos mil veintisiete. Y la recomendación: implantar con finalidad exclusivamente preventiva, con el recurso preventivo al mando, con evaluación de impacto y consulta previa, y medir el resultado en siniestralidad, no en sanciones.")

# ================= 11-12. BIBLIOGRAFÍA =================
REFS1 = [
    'Ley 31/1995, de 8 de noviembre, de Prevención de Riesgos Laborales. Boletín Oficial del Estado, 269, de 10 de noviembre de 1995. https://www.boe.es/buscar/act.php?id=BOE-A-1995-24292',
    'Ley Orgánica 3/2018, de 5 de diciembre, de Protección de Datos Personales y garantía de los derechos digitales. Boletín Oficial del Estado, 294, de 6 de diciembre de 2018. https://www.boe.es/buscar/act.php?id=BOE-A-2018-16673',
    'Real Decreto 1627/1997, de 24 de octubre, por el que se establecen disposiciones mínimas de seguridad y de salud en las obras de construcción. Boletín Oficial del Estado, 256, de 25 de octubre de 1997. https://www.boe.es/buscar/act.php?id=BOE-A-1997-22614',
    'Real Decreto Legislativo 2/2015, de 23 de octubre, por el que se aprueba el texto refundido de la Ley del Estatuto de los Trabajadores. Boletín Oficial del Estado, 255, de 24 de octubre de 2015. https://www.boe.es/buscar/act.php?id=BOE-A-2015-11430',
    'Reglamento (UE) 2016/679 del Parlamento Europeo y del Consejo, de 27 de abril de 2016, relativo a la protección de las personas físicas en lo que respecta al tratamiento de datos personales y a la libre circulación de estos datos (Reglamento general de protección de datos). Diario Oficial de la Unión Europea, L 119, de 4 de mayo de 2016. https://eur-lex.europa.eu/eli/reg/2016/679/oj',
    'Reglamento (UE) 2024/1689 del Parlamento Europeo y del Consejo, de 13 de junio de 2024, por el que se establecen normas armonizadas en materia de inteligencia artificial (Reglamento de Inteligencia Artificial). Diario Oficial de la Unión Europea, L, de 12 de julio de 2024. https://eur-lex.europa.eu/eli/reg/2024/1689/oj',
    'Reglamento (UE) 2026/1744 del Parlamento Europeo y del Consejo, de 8 de julio de 2026, por el que se modifican los Reglamentos (UE) 2024/1689, (UE) 2018/1139 y (UE) 2023/1230 en lo que respecta a la simplificación de la aplicación de normas armonizadas en materia de inteligencia artificial (Ómnibus digital sobre IA). Diario Oficial de la Unión Europea, L, de 24 de julio de 2026. https://eur-lex.europa.eu/eli/reg/2026/1744/oj',
]
REFS2 = [
    'Agencia Europea para la Seguridad y la Salud en el Trabajo. (2022a). Inteligencia artificial para la gestión de trabajadores: implicaciones para la salud y la seguridad en el trabajo. Recuperado el 24 de noviembre de 2026, de https://osha.europa.eu/es/publications/artificial-intelligence-worker-management-implications-occupational-safety-and-health',
    'Agencia Europea para la Seguridad y la Salud en el Trabajo. (2022b). Inteligencia artificial para la gestión de trabajadores: medidas de prevención. Recuperado el 24 de noviembre de 2026, de https://osha.europa.eu/es/publications/artificial-intelligence-worker-management-prevention-measures',
    'Fang, Q., Li, H., Luo, X., Ding, L., Luo, H., Rose, T. M., y An, W. (2018). Detecting non-hardhat-use by a deep learning method from far-field surveillance videos. Automation in Construction, 85, 1-9. https://doi.org/10.1016/j.autcon.2017.09.018',
    'Instituto Nacional de Seguridad y Salud en el Trabajo. (2026). Últimos datos de siniestralidad: periodo octubre 2025 a septiembre 2026. Recuperado el 24 de noviembre de 2026, de https://www.insst.es/el-observatorio/estadisticas/ultimos-datos-de-siniestralidad',
    'Nath, N. D., Behzadan, A. H., y Paal, S. G. (2020). Deep learning for site safety: Real-time detection of personal protective equipment. Automation in Construction, 112, 103085. https://doi.org/10.1016/j.autcon.2020.103085',
    'Teizer, J., Allread, B. S., Fullerton, C. E., y Hinze, J. (2010). Autonomous pro-active real-time construction worker and equipment operator proximity safety alert system. Automation in Construction, 19, 630-640. https://doi.org/10.1016/j.autcon.2010.02.009',
    'Zhang, S., Sulankivi, K., Kiviniemi, M., Romo, I., Eastman, C. M., y Teizer, J. (2015). BIM-based fall hazard identification and prevention in construction safety planning. Safety Science, 72, 31-45. https://doi.org/10.1016/j.ssci.2014.08.001',
]
for num, title, refs in ((11, '9. Bibliografía (APA 7.ª edición), 1 de 2: normativa', REFS1), (12, '9. Bibliografía (APA 7.ª edición), 2 de 2: estadística, informes y artículos', REFS2)):
    s = new_slide(title, num)
    tb = textbox(s, Inches(0.6), Inches(1.15), Inches(12.1), Inches(5.9), [(r, {'bullet': False, 'after': 7}) for r in refs], size=11.5, line_spacing=1.0)
    for par in tb.text_frame.paragraphs:
        pPr = par._p.get_or_add_pPr(); pPr.set('marL', str(int(Inches(0.4)))); pPr.set('indent', str(-int(Inches(0.4))))
    notes(s, "Bibliografía en APA séptima edición, con la normativa española y europea en texto consolidado y, en la segunda diapositiva, la estadística del INSST, los dos informes de la Agencia Europea y los cuatro artículos científicos citados en los resultados." if num == 11 else "Todas las fuentes de internet llevan la fecha de consulta, y los artículos su identificador DOI.")

# ================= 13. ANEXO IA =================
s = new_slide('10. Anexo: declaración de uso de inteligencia artificial', 13)
textbox(s, Inches(0.6), Inches(1.15), Inches(12.1), Inches(5.9), [
    ('**Declaración.** Para la elaboración de esta PEC **sí** se ha utilizado inteligencia artificial generativa, en la primera versión de la presentación y en su revisión.', {'bullet': False, 'after': 8}),
    ('**Herramientas empleadas (cita APA):**', {'bullet': False, 'after': 2}),
    ('OpenAI. (2026). ChatGPT [Modelo de lenguaje de gran tamaño]. https://chatgpt.com/', {'bullet': True, 'after': 2}),
    ('Google. (2026). Gemini [Modelo de lenguaje de gran tamaño]. https://gemini.google.com/', {'bullet': True, 'after': 2}),
    ('Anthropic. (2026). Claude [Modelo de lenguaje de gran tamaño]. https://claude.ai/', {'bullet': True, 'after': 8}),
    ('**Consultas e instrucciones:** en la primera versión, a ChatGPT y Gemini, la estructuración de la presentación académica, la adaptación del contenido al Reglamento (UE) 2024/1689 y el guion expositivo. En la revisión, a Claude, se le facilitaron el enunciado de la PEC 02 y la primera versión, con la petición de corregirla para que cumpliera el enunciado.', {'bullet': False, 'after': 8}),
    ('**Resultados obtenidos:** la primera versión aportó la estructura y el enfoque, sin citas ni datos y con erratas. La revisión incorporó los datos de siniestralidad del INSST, los artículos concretos de la normativa y el calendario actualizado del Reglamento, los cuatro artículos científicos y las dos tablas de resultados, la bibliografía en APA 7 y el guion de narración, y corrigió las erratas.', {'bullet': False, 'after': 8}),
    ('**Reflexión y aportación personal:** la elección del tema y del enfoque, centrado en el cumplimiento del Reglamento de IA en obra, son míos, igual que la primera versión, hecha con ayuda de ChatGPT y Gemini. La IA ha sido útil para completar fuentes, datos y citas, pero sus resultados hay que comprobarlos: la primera versión tenía errores que solo aparecieron al contrastarla con las fuentes oficiales. He revisado la versión final y la narración del vídeo es mía.', {'bullet': False}),
], size=12.5, line_spacing=1.03)
notes(s, "Termino con la declaración de uso de inteligencia artificial que pide el enunciado. Sí la he utilizado: ChatGPT y Gemini en la primera versión, para la estructura, la adaptación al Reglamento y el guion, y Claude en la revisión, que añadió los datos del INSST, las citas, la normativa concreta, los artículos científicos y el guion definitivo. El tema y el enfoque son míos, y la conclusión que saco es que la inteligencia artificial ayuda, pero hay que contrastarla con las fuentes oficiales, porque la primera versión tenía errores. Muchas gracias.")

out = 'Moncada_Roa_Milagros_PEC02_IA_Construccion.pptx'
prs.save(out)
words = sum(len(sl.notes_slide.notes_text_frame.text.split()) for sl in prs.slides)
print('guardado', out, '| diapositivas:', len(prs.slides), '| palabras del guion:', words, '| minutos a 140 ppm:', round(words / 140, 1))
