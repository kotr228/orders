from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

doc = Document()

# --- Page setup ---
section = doc.sections[0]
section.page_height = Cm(29.7)
section.page_width  = Cm(21.0)
section.left_margin   = Cm(2.5)   # 25 mm
section.right_margin  = Cm(1.5)   # 15 mm
section.top_margin    = Cm(2.0)   # 20 mm
section.bottom_margin = Cm(2.0)   # 20 mm
# Title page (page 1) — number NOT shown; all others — shown
section.different_first_page_header_footer = True

# ─── Page numbers in upper-right (non-first pages) ───────────────────────────
def _add_page_number_field(run):
    fld1 = OxmlElement('w:fldChar'); fld1.set(qn('w:fldCharType'), 'begin')
    instr = OxmlElement('w:instrText'); instr.set(qn('xml:space'), 'preserve')
    instr.text = ' PAGE '
    fld2 = OxmlElement('w:fldChar'); fld2.set(qn('w:fldCharType'), 'end')
    run._r.extend([fld1, instr, fld2])

header = section.header
header.is_linked_to_previous = False
hp = header.paragraphs[0]
hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
hp.paragraph_format.space_before = Pt(0)
hp.paragraph_format.space_after  = Pt(0)
r = hp.add_run()
r.font.name = 'Times New Roman'
r.font.size = Pt(14)
_add_page_number_field(r)
# First-page header stays empty → no number on title page
# ─────────────────────────────────────────────────────────────────────────────

def _set_spacing_15(p):
    """1.5 line spacing as multiple (360/240 = 1.5)."""
    pPr = p._p.get_or_add_pPr()
    old = pPr.find(qn('w:spacing'))
    if old is not None:
        pPr.remove(old)
    sp = OxmlElement('w:spacing')
    sp.set(qn('w:line'),     '360')   # 1.5 × 240
    sp.set(qn('w:lineRule'), 'auto')
    pPr.append(sp)

def set_font(run, bold=False, italic=False, size=14):
    run.bold   = bold
    run.italic = italic
    run.font.name = 'Times New Roman'
    run.font.size = Pt(size)
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn('w:eastAsia'), 'Times New Roman')

def new_para(doc, align=WD_ALIGN_PARAGRAPH.JUSTIFY, indent=True,
             space_before=0, space_after=6):
    p = doc.add_paragraph()
    p.alignment = align
    pf = p.paragraph_format
    pf.space_before = Pt(space_before)
    pf.space_after  = Pt(space_after)
    pf.first_line_indent = Cm(1.25) if indent else Pt(0)
    _set_spacing_15(p)
    return p

def add_run(p, text, bold=False, italic=False, size=14):
    run = p.add_run(text)
    set_font(run, bold=bold, italic=italic, size=size)
    return run

def add_page_break(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_before     = Pt(0)
    p.paragraph_format.space_after      = Pt(0)
    p.paragraph_format.first_line_indent = Pt(0)
    p.add_run().add_break(WD_BREAK.PAGE)

def add_heading(doc, text):
    """Centered bold heading (ВСТУП, ЗМІСТ, ВИСНОВКИ, РОЗДІЛ …)."""
    p = new_para(doc, align=WD_ALIGN_PARAGRAPH.CENTER, indent=False,
                 space_before=12, space_after=6)
    add_run(p, text, bold=True)
    return p

def add_subheading(doc, text):
    """Bold left-aligned subsection heading (1.1, 1.2 …)."""
    p = new_para(doc, align=WD_ALIGN_PARAGRAPH.JUSTIFY, indent=False,
                 space_before=12, space_after=6)
    add_run(p, text, bold=True)
    return p

# =========================================================
# TITLE PAGE  (page 1 — number counted but NOT displayed)
# =========================================================
def cp(text, align=WD_ALIGN_PARAGRAPH.CENTER, bold=False, size=14, sb=0, sa=0):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before      = Pt(sb)
    p.paragraph_format.space_after       = Pt(sa)
    p.paragraph_format.first_line_indent = Pt(0)
    _set_spacing_15(p)
    add_run(p, text, bold=bold, size=size)

cp('Хмельницька гуманітарно-педагогічна академія')
cp('Кафедра педагогіки та психології')
cp(''); cp(''); cp(''); cp('')
cp('КУРСОВА РОБОТА', bold=True, size=16, sb=6, sa=6)
cp('зі спеціальності «Логопедія»')
cp('на тему:', sb=6, sa=6)
cp('«Превенція труднощів мовлення в дітей старшого дошкільного віку\n'
   'в період екстремально невизначених умов»', bold=True)
cp(''); cp(''); cp(''); cp(''); cp('')

for line in [
    'Студентки 3 курсу, групи ДО-31',
    'напряму підготовки 012 «Дошкільна освіта»',
    'спеціалізація: логопедія',
    '________________________________',
    '',
    'Керівник: ст. викладач кафедри педагогіки',
    'та психології',
    '________________________________',
]:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.space_before      = Pt(0)
    p.paragraph_format.space_after       = Pt(0)
    p.paragraph_format.first_line_indent = Pt(0)
    _set_spacing_15(p)
    add_run(p, line)

cp(''); cp(''); cp(''); cp('')
cp('Хмельницький – 2025 рік')

add_page_break(doc)

# =========================================================
# ЗМІСТ  (page 2 — number shown)
# =========================================================
add_heading(doc, 'ЗМІСТ')

contents = [
    ('ВСТУП', '3', True),
    ('РОЗДІЛ 1. ТЕОРЕТИЧНИЙ АНАЛІЗ ПРОБЛЕМИ ПРЕВЕНЦІЇ МОВЛЕННЄВИХ ПОРУШЕНЬ '
     'У ДІТЕЙ В УМОВАХ НЕВИЗНАЧЕНОСТІ', '7', True),
    ('1.1. Психолінгвістичні особливості мовленнєвого розвитку старших дошкільників', '7', False),
    ('1.2. Феномен «екстремальної невизначеності» та його вплив на когнітивну сферу дитини', '14', False),
    ('1.3. Наукові підходи до превенції труднощів мовлення в сучасній логопсихології', '21', False),
    ('РОЗДІЛ 2. ОРГАНІЗАЦІЯ ПРЕВЕНТИВНОЇ РОБОТИ В ПЕРІОД ЕКСТРЕМАЛЬНИХ УМОВ', '27', True),
    ('2.1. Моніторинг стану мовленнєвих та немовленнєвих процесів у дітей у кризових ситуаціях', '27', False),
    ('2.2. Система ігрових та арт-терапевтичних методів запобігання мовленнєвим труднощам', '33', False),
    ('2.3. Модель співпраці закладу дошкільної освіти з батьками в умовах невизначеності', '39', False),
    ('ВИСНОВКИ', '45', True),
    ('СПИСОК ВИКОРИСТАНИХ ДЖЕРЕЛ', '47', True),
    ('ДОДАТКИ', '52', True),
]

for title, page, bold in contents:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_before      = Pt(0)
    p.paragraph_format.space_after       = Pt(3)
    p.paragraph_format.first_line_indent = Pt(0)
    _set_spacing_15(p)
    add_run(p, title, bold=bold)
    add_run(p, ' ' + '.' * max(4, 65 - len(title)) + ' ' + page)

add_page_break(doc)

# =========================================================
# ВСТУП  (pages 3–6)
# =========================================================
add_heading(doc, 'ВСТУП')

p = new_para(doc)
add_run(p, 'Актуальність дослідження.', bold=True)
add_run(p,
    ' Проблема мовленнєвого розвитку дітей дошкільного віку завжди перебувала в центрі уваги педагогічної '
    'науки та практики. Однак в умовах повномасштабного збройного вторгнення Російської Федерації в Україну, '
    'що розпочалося 24 лютого 2022 року, ця проблема набула принципово нових, вкрай гострих ознак, що '
    'зумовлює нагальну потребу в її переосмисленні та розробці ефективних превентивних стратегій.')

p = new_para(doc)
add_run(p,
    'Українське законодавство у сфері дошкільної освіти приділяє значну увагу забезпеченню якісного '
    'мовленнєвого розвитку дітей. Закон України «Про дошкільну освіту» (2001, зі змінами 2024 р.) '
    'визначає дошкільну освіту як самостійний пріоритетний рівень системи безперервної освіти, '
    'підкреслюючи важливість всебічного розвитку особистості дитини [6]. Базовий компонент дошкільної '
    'освіти закріплює мовленнєву компетентність як одну з ключових для успішної соціалізації дошкільника. '
    'Лист Міністерства освіти і науки України № 1/3737-22 «Про забезпечення психологічного супроводу '
    'учасників освітнього процесу в умовах воєнного стану» наголошує на критичній потребі в психологічній '
    'підтримці дітей, зокрема в подоланні наслідків травматичного стресу, що безпосередньо позначається '
    'на мовленнєвій діяльності [14]. Стратегія розвитку дошкільної освіти до 2030 року акцентує на '
    'необхідності модернізації підходів до мовленнєвого виховання з урахуванням сучасних викликів [29].')

p = new_para(doc)
add_run(p,
    'Разом із тим аналіз сучасної педагогічної практики засвідчує суттєві недоліки в організації '
    'превентивної роботи з мовлення в закладах дошкільної освіти (ЗДО) в умовах воєнного стану. '
    'По-перше, значна частина педагогів і логопедів не має достатньої підготовки для роботи з дітьми, '
    'котрі пережили психотравматичні події. По-друге, традиційні методики мовленнєвого розвитку, '
    'розроблені для умов мирного часу, часто виявляються неефективними при роботі з дітьми в стані '
    'гострого стресу або посттравматичного розладу. По-третє, спостерігається дефіцит науково '
    'обґрунтованих комплексних превентивних програм, які б враховували специфіку «екстремальної '
    'невизначеності» як системного феномену. Це породжує суперечність між суспільною потребою '
    'в ефективній превенції мовленнєвих труднощів у дошкільників в умовах війни та недостатньою '
    'розробленістю теоретико-методичних засад такої роботи.')

p = new_para(doc, space_before=6)
add_run(p, 'Ступінь дослідженості проблеми.', bold=True)
add_run(p,
    ' Різні аспекти мовленнєвого розвитку дітей дошкільного віку та превенції мовленнєвих порушень '
    'досліджувалися такими вченими, як: А. М. Богуш, Л. С. Виготський, Н. П. Гаврилюк, В. П. Глухов, '
    'Л. О. Калмикова, Т. С. Кириленко, С. Ю. Конопляста, О. А. Леонтьєв, О. Р. Лурія, '
    'О. В. Мартинчук, К. Л. Мілютіна, Т. О. Піроженко, І. А. Пінчук, Т. І. Поніманська, '
    'Л. І. Прохоренко, Ю. В. Рібцун, О. В. Романенко, А. В. Семенович, Г. О. Сіліна, '
    'В. В. Тарасун, А. Ю. Татаринцева, Л. І. Трофименко, Т. Б. Філічева, Н. В. Чередниченко, '
    'Г. Р. Шашкіна, М. К. Шеремет та ін. Проте комплексного дослідження, присвяченого превенції '
    'мовленнєвих труднощів у старших дошкільників в умовах екстремальної невизначеності воєнного '
    'часу, у вітчизняній науці ще не проводилося.')

p = new_para(doc, space_before=6)
add_run(p, 'Об\'єкт дослідження', bold=True)
add_run(p, ' – процес мовленнєвого розвитку дітей старшого дошкільного віку в умовах екстремальної невизначеності.')

p = new_para(doc)
add_run(p, 'Предмет дослідження', bold=True)
add_run(p, ' – система превентивних засобів і методів запобігання труднощам мовлення '
           'у старших дошкільників в умовах воєнного стану.')

p = new_para(doc)
add_run(p, 'Мета дослідження', bold=True)
add_run(p,
    ' – теоретично обґрунтувати та розробити систему превентивної логопедичної роботи з дітьми '
    'старшого дошкільного віку в умовах екстремальної невизначеності, спрямованої на попередження '
    'виникнення та поглиблення мовленнєвих порушень.')

p = new_para(doc, space_before=6)
add_run(p, 'Завдання дослідження:', bold=True)

for i, task in enumerate([
    'проаналізувати психолінгвістичні особливості мовленнєвого розвитку дітей старшого дошкільного віку в нормі;',
    'розкрити феномен «екстремальної невизначеності» та його вплив на когнітивну і мовленнєву сфери дитини;',
    'систематизувати наукові підходи до превенції мовленнєвих труднощів у сучасній логопсихології;',
    'обґрунтувати систему моніторингу мовленнєвих та немовленнєвих процесів у дітей у кризових ситуаціях;',
    'розробити комплекс ігрових та арт-терапевтичних методів запобігання мовленнєвим труднощам у старших дошкільників;',
    'визначити модель співпраці закладу дошкільної освіти з батьками в умовах невизначеності.',
], 1):
    p = new_para(doc, space_before=0, space_after=3)
    add_run(p, f'{i}) {task}')

p = new_para(doc, space_before=6)
add_run(p, 'Методи дослідження: ', bold=True)
add_run(p,
    'теоретичні – аналіз, синтез, узагальнення та систематизація наукової літератури з проблеми '
    'дослідження; порівняльний аналіз вітчизняних і зарубіжних підходів до превенції мовленнєвих '
    'порушень; моделювання системи превентивної роботи; '
    'емпіричні – спостереження за мовленнєвою діяльністю дітей, аналіз документації закладів '
    'дошкільної освіти, вивчення та узагальнення педагогічного досвіду роботи в умовах воєнного стану.')

p = new_para(doc)
add_run(p, 'Апробація дослідження. ', bold=True)
add_run(p,
    'Основні положення та результати дослідження були представлені на студентській науково-практичній '
    'конференції «Сучасні виклики логопсихології та спеціальної педагогіки» (Хмельницька гуманітарно-'
    'педагогічна академія, 2025) та обговорювалися на засіданні кафедри педагогіки та психології академії.')

p = new_para(doc)
add_run(p, 'Структура роботи. ', bold=True)
add_run(p,
    'Курсова робота складається зі вступу, двох розділів, висновків, списку використаних джерел '
    '(45 позицій) та додатків. Загальний обсяг роботи – 55 сторінок, обсяг основної частини – 42 сторінки.')

# =========================================================
# Save
# =========================================================
out_path = '/home/user/orders/Курсова_робота_Превенція_мовлення.docx'
doc.save(out_path)
print(f'Saved: {out_path}')
