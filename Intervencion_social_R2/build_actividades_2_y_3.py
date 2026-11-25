import re
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
BASE = '/home/user/fj/Intervencion_social_R2/'
a2 = open(BASE + 'R2_actividad2_FINAL.md', encoding='utf-8').read()
a3 = open(BASE + 'R2_actividad3_FINAL.md', encoding='utf-8').read()
a2_body, refs = a2.split('## Referencias')
blocks = [('body', b) for b in a2_body.split('\n\n')] + [('pagebreak', '')] + [('body', b) for b in a3.split('\n\n')] \
         + [('body', '# Referencias')] + [('ref', b) for b in refs.split('\n\n')]
def curly(s):
    out, dbl = [], False
    for i, ch in enumerate(s):
        if ch == '"': out.append('”' if dbl else '“'); dbl = not dbl
        elif ch == "'":
            prev = s[i-1] if i else ' '
            out.append('‘' if (prev.isspace() or prev in '“(') else '’')
        else: out.append(ch)
    return ''.join(out)
def add_runs(par, text):
    for part in re.split(r"(\*[^*]+\*)", curly(text)):
        if not part: continue
        it = part.startswith('*') and part.endswith('*')
        r = par.add_run(part[1:-1] if it else part); r.italic = it; r.font.size = Pt(11); r.font.name = 'Calibri'
doc = Document()
for s in doc.sections:
    s.left_margin = s.right_margin = Cm(2.5); s.top_margin = s.bottom_margin = Cm(2.5)
doc.styles['Normal'].font.name = 'Calibri'; doc.styles['Normal'].font.size = Pt(11)
new_page = False
for kind, raw in blocks:
    b = raw.strip()
    if kind == 'pagebreak':
        new_page = True; continue
    if not b: continue
    if b.startswith('#'):
        level = len(b) - len(b.lstrip('#')); text = b.lstrip('#').strip()
        h = doc.add_heading(level=1 if level == 1 else 2)
        r = h.add_run(text); r.font.name = 'Calibri'; r.bold = True
        r.font.color.rgb = RGBColor(0, 0, 0); r.font.size = Pt(14 if level == 1 else 12)
        h.paragraph_format.space_before = Pt(12 if level == 1 else 10); h.paragraph_format.space_after = Pt(6)
        if new_page: h.paragraph_format.page_break_before = True; new_page = False
        continue
    par = doc.add_paragraph(); pf = par.paragraph_format
    pf.line_spacing = 1.15; pf.space_after = Pt(6)
    if kind == 'ref':
        pf.left_indent = Cm(1.25); pf.first_line_indent = Cm(-1.25); par.alignment = WD_ALIGN_PARAGRAPH.LEFT
    else:
        par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    add_runs(par, b.replace('\n', ' '))
doc.save(BASE + 'R2_actividades_2_y_3.docx')
print("docx ok")
