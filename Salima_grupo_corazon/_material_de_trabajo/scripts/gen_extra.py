# -*- coding: utf-8 -*-
"""Genera extra_sections.json (introducción y anexos) para build_banco.py."""
import os, re, json, glob, sys, importlib.util
HERE = os.path.dirname(os.path.abspath(__file__))
BIBLIO = os.path.dirname(HERE)
sys.path.insert(0, HERE)
from limpieza import limpiar

spec = importlib.util.spec_from_file_location('b', os.path.join(HERE, 'build_ap2_v2.py'))
# solo necesitamos las funciones de etiqueta: las copiamos sin ejecutar la construcción del documento
src = open(os.path.join(HERE, 'build_ap2_v2.py'), encoding='utf-8').read()
ns = {'__file__': os.path.join(HERE, 'build_ap2_v2.py'), '__name__': 'build_ap2_v2_parcial'}
exec(src.split('# ---------- cargar fichas ----------')[0].replace("OUT = sys.argv[1] if len(sys.argv) > 1 else", "OUT = '' if True else"), ns)
etiqueta = ns['etiqueta']; T = ns['T']

F = {}
for f in sorted(glob.glob(os.path.join(BIBLIO, 'fichas', 'fichas_*.json'))):
    for x in json.load(open(f, encoding='utf-8')): F.setdefault(x['id'], x)
cit = json.load(open(os.path.join(BIBLIO, 'citadas.json'), encoding='utf-8'))

SECC = [('Intro', T.INTRO), ('2.1', T.S21[1]), ('2.2', T.S22[1]), ('2.3', T.S23[1]), ('2.4', T.S24[1]),
        ('2.5', T.S25[1]), ('2.6', T.S26[1]), ('Cierre', T.CIERRE[1])]

def donde(fid):
    lab, anio = etiqueta(F[fid]['referencia_apa7'])
    out = []
    for nom, ps in SECC:
        txt = '\n'.join(ps)
        if any(anio in txt[m.end():m.end() + 45] for m in re.finditer(re.escape(lab), txt)):
            out.append(nom)
    return ', '.join(out)

def corta(fid):
    lab, anio = etiqueta(F[fid]['referencia_apa7'])
    return f'{lab} ({anio})'

def primera_frase(s):
    s = limpiar(s)
    m = re.search(r'^(.+?[.;])\s', s + ' ')
    return m.group(1) if m else s

n_total = len(F)
n_comp = sum(1 for x in F.values() if x['acceso'].strip().lower().startswith('texto completo'))
n_res = n_total - n_comp
n_cit = len(cit['citadas'])
cit_res = [i for i in cit['citadas'] if not F[i]['acceso'].strip().lower().startswith('texto completo')]

intro = [
 {'h': 'Para qué sirve este documento', 'p': [
  'Este banco reúne las fuentes que sostienen el apartado 2 del trabajo (marco teórico y antecedentes). Cada fuente tiene una ficha con el formato que propone la guía del grupo y, cuando el artículo es de acceso abierto, su PDF está en la carpeta «articulos», ordenado por subapartado. Sirve para redactar, para comprobar una cita antes de entregar y para que el resto del grupo pueda reutilizar fuentes en sus apartados, por ejemplo en la metodología o en la evaluación.']},
 {'h': 'Cómo se ha hecho la búsqueda', 'p': [
  'La búsqueda se hizo en noviembre de 2026 siguiendo los seis bloques de la guía. Se consultaron ERIC, PubMed y Europe PMC, Crossref, Semantic Scholar y el buscador académico SciSpace. También se revisaron las webs de revistas españolas (*Revista de Educación*, *Educación XX1*, *Revista Complutense de Educación*, *REIS* o *Revista de Psicodidáctica*, entre otras) y repositorios universitarios como Digibug (UGR), y la normativa se consultó directamente en el BOE.',
  'Se combinaron términos en castellano y en inglés: «educación intercultural» AND «educación infantil», «competencia emocional» AND preescolar, «socialización emocional», «participación familiar» AND inmigración, *emotion socialization* AND *culture*, *social-emotional learning* AND *preschool* AND *meta-analysis* o *family-school partnership* AND *immigrant*, con búsquedas específicas sobre Melilla (convivencia escolar, amazige o tamazight, bilingüismo).',
  'Se dio prioridad a artículos revisados por pares, metaanálisis y revisiones sistemáticas, a estudios con niños de tres a seis años, a investigaciones hechas en España y a lo publicado sobre Melilla. Se mantuvieron algunos clásicos anteriores a 2010 cuando siguen siendo referencia. Los datos de cada referencia (autoría, año, revista, volumen, páginas y DOI) se comprobaron en Crossref, en la web de la revista o en el BOE, y en varias fichas se anotan correcciones respecto a cómo circulan esas referencias.']},
 {'h': 'Qué se ha leído y qué no', 'p': [
  f'El banco tiene {n_total} fichas. En {n_comp} se ha leído el texto completo y las citas textuales están copiadas del original con su página; en {n_res} solo se pudo leer el resumen, porque el artículo es de pago o la web de la editorial no permitía la descarga. Las fichas de resumen lo indican en el campo «Acceso» y no se han usado para datos que no aparezcan en el resumen. De todas ellas, {n_cit} se citan en la versión 2 del apartado y las demás quedan como reserva.',
  f'Conviene tener en cuenta dos limitaciones. Dialnet y Google Scholar no se pudieron consultar, así que merece la pena hacer una búsqueda rápida en Dialnet con los mismos términos por si falta algún estudio español reciente, sobre todo de Melilla. Y antes de la entrega final sería bueno descargar desde la biblioteca de la UGR las {len(cit_res)} fuentes citadas que solo se han leído en resumen (están en la penúltima tabla). Algunas son fáciles de conseguir: el metaanálisis de Silberman et al. (2026) es de acceso abierto en la web de la revista y la revisión de Aboud et al. (2012), que ha quedado en reserva, también lo es en ScienceDirect.']},
 {'h': 'Cómo está organizado', 'p': [
  'Primero aparece una tabla resumen por subapartado con la fuente, el tipo de trabajo, el acceso, si se cita en la versión 2 del apartado y para qué sirve. Después vienen las fichas completas en el mismo orden. Al final se recogen los artículos de la normativa con su texto literal, las fuentes citadas que conviene leer completas y las que han quedado en reserva. Algunas notas de las fichas («Uso en el trabajo», «Verificación») se escribieron al revisar la primera versión del apartado; la versión 2 ya incorpora esas correcciones.']},
]

normativa_rows = [
 ['Ley Orgánica 2/2006 (texto consolidado), art. 1, letras a bis y b', '«a bis) La calidad de la educación para todo el alumnado, sin que exista discriminación alguna por razón de nacimiento, sexo, origen racial, étnico o geográfico, discapacidad, edad, enfermedad, religión o creencias, orientación sexual o identidad sexual o cualquier otra condición o circunstancia personal o social. b) La equidad, que garantice la igualdad de oportunidades para el pleno desarrollo de la personalidad a través de la educación, la inclusión educativa, la igualdad de derechos y oportunidades [...]»', '2.1'],
 ['LOE, art. 2.1.g', '«La formación en el respeto y reconocimiento de la pluralidad lingüística y cultural de España y de la interculturalidad como un elemento enriquecedor de la sociedad.»', '2.1'],
 ['LOE, art. 12.4', '«Con objeto de respetar la responsabilidad fundamental de las madres y padres o tutores legales en esta etapa, los centros de educación infantil cooperarán estrechamente con ellos.»', '2.5 y cierre'],
 ['LOE, art. 12.5', '«La programación, la gestión y el desarrollo de la educación infantil atenderán, en todo caso, a la compensación de los efectos que las desigualdades de origen cultural, social y económico tienen en el aprendizaje y evolución infantil [...]»', 'Reserva (introducción o justificación)'],
 ['LOE, art. 13, letras a, d y e', '«a) Conocer su propio cuerpo y el de los otros, sus posibilidades de acción y aprender a respetar las diferencias. [...] d) Desarrollar sus capacidades afectivas. e) Relacionarse con los demás en igualdad y adquirir progresivamente pautas elementales de convivencia y relación social, así como ejercitarse en el uso de la empatía y la resolución pacífica de conflictos, evitando cualquier tipo de violencia.»', 'Reserva (objetivos del programa)'],
 ['Real Decreto 95/2022, art. 7, letras d y e', '«d) Desarrollar sus capacidades emocionales y afectivas. e) Relacionarse con los demás en igualdad y adquirir progresivamente pautas elementales de convivencia y relación social, así como ejercitarse en el uso de la empatía y la resolución pacífica de conflictos, evitando cualquier tipo de violencia.»', 'Reserva (objetivos del programa)'],
 ['RD 95/2022, anexo I (competencia plurilingüe)', '«En esta etapa, se inicia el contacto con lenguas y culturas distintas de la familiar, fomentando en niños y niñas las actitudes de respeto y aprecio por la diversidad lingüística y cultural, así como el interés por el enriquecimiento de su repertorio lingüístico. Se promueven de este modo el diálogo y la convivencia democrática.»', '2.1'],
 ['RD 95/2022, anexo II (introducción a las áreas)', '«Un aula inclusiva da cabida a diferentes posibilidades y limitaciones, a distintos modos de lenguaje y comunicación, a diversas formas de movilidad y a múltiples maneras de acercarse y comprender el entorno.»', '2.1'],
 ['RD 95/2022, anexo II, Crecimiento en Armonía, competencia específica 2', '«Reconocer, manifestar y regular progresivamente sus emociones, expresando necesidades y sentimientos para lograr bienestar emocional y seguridad afectiva.»', '2.2'],
 ['RD 95/2022, anexo II, Crecimiento en Armonía, criterio de evaluación 2.2', '«Relacionarse con las otras personas aceptando y mostrando afecto de manera libre, segura, respetuosa y alejada de todo tipo de estereotipos.»', '2.1'],
 ['Orden EFP/608/2022, art. 14.1', '«Cada grupo de niños y niñas tendrá un tutor o tutora, que será su persona de referencia y asumirá la responsabilidad tanto de la relación con las familias o tutores o tutoras legales, como de la coordinación con el resto de profesionales que intervengan en su grupo.»', 'Reserva (metodología: papel de la tutora)'],
 ['Orden EFP/608/2022, art. 16.1', '«Cuando el niño o la niña se incorpore al centro, el tutor o tutora realizará una evaluación inicial cuyos resultados recogerá en un informe. Asimismo, añadirá en dicho documento la información aportada por las familias o tutores o tutoras legales [...]»', 'Reserva (diagnóstico de septiembre)'],
 ['Orden EFP/608/2022, art. 19.2', '«Los centros cooperarán estrechamente con los padres, madres, tutores o tutoras legales, generando cauces de participación y colaboración, y harán explícitas las actuaciones previstas para favorecer su participación en el proceso educativo de los niños y niñas.»', '2.5 y cierre'],
 ['Orden EFP/608/2022, anexo IV (Atención a las diferencias individuales)', '«Es importante, por tanto, evitar atribuir etiquetas y calificativos a la conducta, rasgos de personalidad o capacidades personales [...]. Cada niño o niña llega a la escuela cargado de experiencias, costumbres y tradiciones que enriquecen el aula. El equipo educativo debe aprovechar la oportunidad que ofrece esta pluralidad cultural para trabajar los valores de la convivencia democrática. Esta diversidad implica a su vez una atención respetuosa y personalizada de los distintos perfiles lingüísticos, así como cualquier otra práctica que favorezca la comprensión y el respeto intercultural.»', '2.1'],
]

def como_conseguir(fid):
    x = F[fid]; acc = x['acceso'].lower(); enl = x.get('enlace', '')
    if 'acceso abierto' in acc and 'puede descargarse' in acc:
        return f'Acceso abierto en la web de la revista: {enl}'
    return f'Biblioteca de la UGR (con el usuario de la universidad): {enl}'

pend_rows = [[corta(i), donde(i), como_conseguir(i)] for i in sorted(cit_res, key=lambda i: corta(i))]
res_rows = [[corta(i), 'Texto completo' if F[i]['acceso'].strip().lower().startswith('texto completo') else 'Solo resumen', primera_frase(F[i]['uso_en_el_trabajo'])]
            for i in sorted(cit['no_citadas'], key=lambda i: corta(i))]

final = [
 {'h': 'Normativa consultada (texto literal)', 'p': ['Todos los textos se han copiado del BOE (texto consolidado de la Ley Orgánica 2/2006 tras la Ley Orgánica 3/2020, Real Decreto 95/2022 y Orden EFP/608/2022). La última columna indica dónde se usa en el apartado 2 o si queda en reserva para otros apartados del trabajo.'],
  'tablas': [{'cols': ['Norma y artículo', 'Texto literal', 'Dónde se usa'], 'rows': normativa_rows, 'widths': [3.6, 10.2, 3.2]}]},
 {'h': 'Fuentes citadas que conviene leer completas antes de entregar', 'p': ['Son las fuentes de la versión 2 que solo se han podido leer en resumen. Lo que se dice de ellas en el texto no va más allá del resumen, pero conviene tenerlas a mano por si el profesorado pregunta por algún detalle.'],
  'tablas': [{'cols': ['Fuente', 'Se cita en', 'Cómo conseguirla'], 'rows': pend_rows, 'widths': [4.4, 2.2, 10.4]}]},
 {'h': 'Fuentes en reserva (no citadas en la versión 2)', 'p': ['Se leyeron o localizaron durante la búsqueda, pero no se citan en el texto actual, bien porque solo se pudo leer el resumen, bien porque otra fuente cubría la misma idea. Pueden servir al grupo para otros apartados.'],
  'tablas': [{'cols': ['Fuente', 'Acceso', 'Posible uso'], 'rows': res_rows, 'widths': [4.4, 2.2, 10.4]}]},
]

json.dump({'intro': intro, 'final': final}, open(os.path.join(BIBLIO, 'extra_sections.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
print('extra_sections.json escrito |', len(pend_rows), 'pendientes |', len(res_rows), 'en reserva')
for r in pend_rows: print('  P:', r[0], '|', r[1], '|', r[2][:80])
for r in res_rows: print('  R:', r[0], '|', r[1], '|', r[2][:100])
