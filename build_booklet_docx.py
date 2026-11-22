import os
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

from PIL import Image, ImageOps

BASE_DIR = "/Users/ivanermak/Documents/deepseek-harness/default-workspace/aleksandra-baraban-leader"
ASSETS = os.path.join(BASE_DIR, "assets")
OUT_DOCX = os.path.join(BASE_DIR, "Буклет_Лидеры_Года_Александра_Барабан.docx")

def make_uniform_thumb(rel_path, out_name, size=(600, 420)):
    src = os.path.join(ASSETS, rel_path)
    dst = os.path.join(ASSETS, "photos", out_name)
    with Image.open(src) as im:
        fitted = ImageOps.fit(im.convert("RGB"), size, method=Image.Resampling.LANCZOS, centering=(0.5, 0.25))
        fitted.save(dst, "JPEG", quality=90)
    return dst

def set_cell_bg(cell, hex_color):
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shading)

def set_cell_margins(cell, top=80, bottom=80, left=120, right=120):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def add_run(p, text, size=9.5, bold=False, italic=False, color=(96, 77, 83)):
    r = p.add_run(text)
    r.font.name = "Arial"
    r.font.size = Pt(size)
    r.bold = bold
    r.italic = italic
    r.font.color.rgb = RGBColor(*color)
    return r

doc = Document()

for section in doc.sections:
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(0.75)
    section.bottom_margin = Cm(0.75)
    section.left_margin = Cm(1.1)
    section.right_margin = Cm(1.1)

# ================= PAGE 1: FRONT DOSSIER =================
header_tbl = doc.add_table(rows=1, cols=2)
header_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
header_tbl.autofit = False
header_tbl.columns[0].width = Cm(12.6)
header_tbl.columns[1].width = Cm(6.0)

c0, c1 = header_tbl.rows[0].cells
set_cell_bg(c0, "604D53")
set_cell_bg(c1, "DB7F8E")
set_cell_margins(c0, 100, 100, 140, 140)
set_cell_margins(c1, 100, 100, 140, 140)

p0 = c0.paragraphs[0]
add_run(p0, "УО «ГРОДНЕНСКИЙ ГОСУДАРСТВЕННЫЙ УНИВЕРСИТЕТ ИМЕНИ ЯНКИ КУПАЛЫ»\n", 8, True, color=(213, 197, 200))
add_run(p0, "Факультет экономики и управления • Сектор инновационных проектов и стартапов", 8, False, color=(157, 163, 164))

p1 = c1.paragraphs[0]
p1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
add_run(p1, "КОНКУРС «ЛИДЕРЫ ГОДА»\n", 9, True, color=(255, 255, 255))
add_run(p1, "Портфолио-досье участника", 7.5, False, color=(213, 197, 200))

# Spacing
sp = doc.add_paragraph()
sp.paragraph_format.space_before = Pt(3)
sp.paragraph_format.space_after = Pt(3)

# Hero Table (Photo + Main Info)
hero_tbl = doc.add_table(rows=1, cols=2)
hero_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
hero_tbl.autofit = False
hero_tbl.columns[0].width = Cm(5.8)
hero_tbl.columns[1].width = Cm(12.8)

hc0, hc1 = hero_tbl.rows[0].cells
set_cell_bg(hc0, "FFF5F5")
set_cell_bg(hc1, "FFF5F5")
set_cell_margins(hc0, 80, 80, 80, 80)
set_cell_margins(hc1, 80, 80, 140, 120)

hp0 = hc0.paragraphs[0]
hp0.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_img = hp0.add_run()
run_img.add_picture(os.path.join(ASSETS, "photos/student_year_roses_portrait.jpg"), width=Cm(5.3))
add_run(hp0, "\nПризер «Студент года» и Первая Вице-Мисс «Гродненчанка»", 7, True, color=(219, 127, 142))

hp1 = hc1.paragraphs[0]
hp1.paragraph_format.space_after = Pt(3)
add_run(hp1, "БАРАБАН АЛЕКСАНДРА ВИТАЛЬЕВНА\n", 15.5, True, color=(96, 77, 83))
add_run(hp1, "«Время действовать: интеллект, слово, инновации и национальная культура»", 9.5, True, True, color=(157, 163, 164))

facts = [
    ("Образование:", "3 курс, специальность «Электронная экономика» ФЭУ ГрГУ им. Янки Купалы (средний балл 8.0+)"),
    ("Работа в университете:", "Делопроизводитель и SMM-специалист сектора инновационных проектов и стартапов ГрГУ им. Я. Купалы"),
    ("Экосистема стартапов:", "Спикер проекта стартап-инкубатора «СтартИнн» ГрГУ (БРСМ), команда STUFF СММ «Space CorpAcc × МТС × ПВТ», призер «ИнНаСтарт» (IT-решение «LibList»), спикер Хакатона 2026 (Case #1)"),
    ("Республика и наука:", "Диплом Совета Республики НС РБ (от Н. И. Кочановой), Диплом I степени респ. финала «АРТ-вакацыі» Минобразования РБ (скрипка, ансамбль «Виола»), лауреат Международного форума в Бресте (БрГТУ)"),
    ("Лидерство и риторика:", "Первая Вице-Мисс «Гродненчанка-2025», 3 место и «Выдающийся интеллект» («Студент года-2025»), дважды лауреат ораторского конкурса «Цицероний-2025», II место в финале Дебатов ГрГУ"),
]

for label, val in facts:
    fp = hc1.add_paragraph()
    fp.paragraph_format.space_before = Pt(1.5)
    fp.paragraph_format.space_after = Pt(1.5)
    add_run(fp, f"▪ {label} ", 8, True, color=(219, 127, 142))
    add_run(fp, val, 8, False, color=(96, 77, 83))

# Section Banner: 4 Pillars
sec1 = doc.add_paragraph()
sec1.paragraph_format.space_before = Pt(6)
sec1.paragraph_format.space_after = Pt(4)
add_run(sec1, "МАТРИЦА ЛИДЕРСТВА: 4 КЛЮЧЕВЫХ ВЕКТОРА ДОСТИЖЕНИЙ (14+ ДИПЛОМОВ)", 10, True, color=(219, 127, 142))

pillars_tbl = doc.add_table(rows=2, cols=2)
pillars_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
pillars_tbl.autofit = False
for row in pillars_tbl.rows:
    row.cells[0].width = Cm(9.3)
    row.cells[1].width = Cm(9.3)

pillars_data = [
    ("1. ЛИДЕРСТВО, ОРАТОРСТВО И ДЕБАТЫ", "FFDBDA", [
        "Первая Вице-Мисс городского конкурса «Гродненчанка – 2025» (БСЖ).",
        "Диплом за III место в финале конкурса «Студент года – 2025» ГрГУ.",
        "Диплом II степени + спецноминация «За оригинальность подачи материала» в конкурсе ораторского мастерства «Цицероний – 2025».",
        "Диплом II степени финала Дебатного турнира ГрГУ им. Янки Купалы."
    ]),
    ("2. ИНТЕЛЛЕКТ, СТАРТАПЫ, ПВТ И НАУКА", "FFF0F0", [
        "Победа в номинации «Выдающийся интеллект» («Студент года – 2025»).",
        "Защита проекта стартап-инкубатора «СтартИнн» на Республиканском конкурсе молодёжных инициатив; STUFF СММ «Space CorpAcc × МТС × ПВТ».",
        "Диплом III степени «ИнНаСтарт» («IT-решение» — книжная соцсеть «LibList») + Case #1 на Хакатоне 2026 + диплом Международного форума в Бресте.",
        "I место в турнире «Законы семьи», I место с командой «4G» ФЭУ в викторине «История земли моей», II место в «Смузи Квиз» («Шутка Фейербаха»)."
    ]),
    ("3. МУЗЫКА, ПОЭЗИЯ И КУЛЬТУРНЫЙ КОД", "FFE8E8", [
        "Диплом Совета Республики НС РБ (вручен лично Н. И. Кочановой) на республиканском конкурсе «Вдохновение для души» (А. Кулешов «Дарагі! Не хаваю тугі»).",
        "Два Диплома I степени (областной этап и республиканский финал «АРТ-вакацыі – 2026» Минобразования РБ) в составе ансамбля скрипачей «Виола».",
        "Сертификат 7 форматов гостиной «Беларусь – перекресток национальных культур». Игра на скрипке и акустической гитаре."
    ]),
    ("4. ГРАЖДАНСКАЯ ПОЗИЦИЯ И ВОЛОНТЁРСТВО", "F5E6E8", [
        "Участница визита делегации в Военную академию Республики Беларусь в рамках меморандума о сотрудничестве.",
        "Лицо республиканской акции молодых избирателей «Мой голос первый!».",
        "Профориентационный спикер ФЭУ ГрГУ в школах; интервью ТРК «Беларусь 4».",
        "Добровольчество по зову сердца: регулярные поездки в приют для бездомных животных (выгул и социализация собак)."
    ])
]

for idx, (title, bg, items) in enumerate(pillars_data):
    r_idx, c_idx = divmod(idx, 2)
    cell = pillars_tbl.rows[r_idx].cells[c_idx]
    set_cell_bg(cell, bg)
    set_cell_margins(cell, 70, 70, 100, 100)
    cp = cell.paragraphs[0]
    cp.paragraph_format.space_after = Pt(2)
    add_run(cp, title, 8.5, True, color=(96, 77, 83))
    for item in items:
        ip = cell.add_paragraph()
        ip.paragraph_format.space_before = Pt(1)
        ip.paragraph_format.space_after = Pt(1)
        add_run(ip, "• " + item, 7.6, False, color=(96, 77, 83))

# Photo Strip Page 1 (4 photos)
p_strip_lbl = doc.add_paragraph()
p_strip_lbl.paragraph_format.space_before = Pt(5)
p_strip_lbl.paragraph_format.space_after = Pt(2)
add_run(p_strip_lbl, "ФОТОХРОНИКА КЛЮЧЕВЫХ СОБЫТИЙ:", 8.5, True, color=(219, 127, 142))

photo_tbl1 = doc.add_table(rows=1, cols=4)
photo_tbl1.alignment = WD_TABLE_ALIGNMENT.CENTER
photo_tbl1.autofit = False
p1_imgs = [
    ("photos/vdohnovenie_kochanova_award_duo.jpg", "Награждение в Совете Республики (с Н. И. Кочановой)"),
    ("photos/startinn_pitch_title.jpg", "Питч стартап-инкубатора «СтартИнн» (БРСМ)"),
    ("photos/grodnenchanka_violin_stage.jpg", "Первая Вице-Мисс «Гродненчанка» (скрипка)"),
    ("photos/military_academy_trio_armor.jpg", "Визит в Военную академию РБ")
]
for i, (rel_path, cap) in enumerate(p1_imgs):
    cell = photo_tbl1.rows[0].cells[i]
    cell.width = Cm(4.65)
    set_cell_bg(cell, "FFF5F5")
    set_cell_margins(cell, 40, 40, 40, 40)
    cp = cell.paragraphs[0]
    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    thumb_path = make_uniform_thumb(rel_path, f"_booklet_thumb_{i+1}.jpg", size=(640, 440))
    cp.add_run().add_picture(thumb_path, width=Cm(4.35))
    add_run(cp, "\n" + cap, 6.5, True, color=(96, 77, 83))

doc.add_page_break()

# ================= PAGE 2: CHRONOLOGICAL TABLE & QR BRIDGE =================
p2_head = doc.add_paragraph()
p2_head.paragraph_format.space_after = Pt(4)
add_run(p2_head, "ПОЛНЫЙ ХРОНОЛОГИЧЕСКИЙ РЕЕСТР ДИПЛОМОВ, ПРОЕКТОВ И ИНИЦИАТИВ (2024–2026)", 10.5, True, color=(96, 77, 83))

reg_tbl = doc.add_table(rows=1, cols=4)
reg_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
reg_tbl.autofit = False

col_widths = [Cm(2.3), Cm(3.0), Cm(6.3), Cm(7.0)]
headers = ["Дата", "Уровень", "Мероприятие / Конкурс / Форум", "Достижение / Статус / Роль"]
for i, h_text in enumerate(headers):
    cell = reg_tbl.rows[0].cells[i]
    cell.width = col_widths[i]
    set_cell_bg(cell, "604D53")
    set_cell_margins(cell, 55, 55, 70, 70)
    hp = cell.paragraphs[0]
    add_run(hp, h_text, 7.5, True, color=(213, 197, 200))

rows_data = [
    ("2026", "Республиканский", "Республиканский фестиваль «АРТ-вакацыі» (Минобразования РБ)", "Диплом I степени («Инструментальный жанр») в составе ансамбля скрипачей «Виола» ГрГУ (подпись Министра А. И. Иванца)."),
    ("2025–2026", "Республиканский", "Республиканский конкурс «Вдохновение для души» (Совет Республики НС РБ)", "Диплом Совета Республики НС РБ (вручен лично Н. И. Кочановой) за номер по поэзии А. Кулешова «Дарагі! Не хаваю тугі»."),
    ("2026", "Республиканский", "Респ. конкурс молодёжных инициатив (БРСМ) & StartUp Space × ПВТ × МТС", "Защита проекта стартап-инкубатора «СтартИнн» ГрГУ; работа в команде STUFF СММ «Space CorpAcc with МТС & Hi-Tech Park»."),
    ("29.05.2026", "Стартап-конкурс", "XVI Открытый конкурс бизнес-идей инновационных стартапов «ИнНаСтарт»", "Диплом III степени в номинации «IT-решение» за проект «LibList» — книжная социальная сеть."),
    ("13.05.2026", "Университетский", "Интеллектуальная игра «Законы семьи» («Неделя семьи» ГрГУ)", "Диплом за I место (командная победа: А. Барабан, М. Гурин, А. Белявская)."),
    ("2026", "Областной этап", "Областной этап фестиваля «АРТ-вакацыі – 2026» (г. Гродно)", "Диплом I степени («Инструментальное творчество») в составе ансамбля скрипачей «Виола» ГрГУ."),
    ("2026", "Кейс-чемпионат", "Университетский Хакатон – 2026 (с ОАО «Белагропромбанк»)", "Разработка и публичная защита Case #1 (by Александра Барабан, Анна Белявская)."),
    ("18–19.12.2025", "Международный", "XIX Международный научный форум «Студенческая научная зима в Бресте»", "Диплом в номинации «Гражданские инициативы» (конкурс «Мой вклад в устойчивое будущее», БрГТУ)."),
    ("29.10.2025", "Университетский", "Университетский конкурс «Студент года – 2025» («Время действовать»)", "Два диплома: 1) Диплом за III место; 2) Победа в номинации «Выдающийся интеллект»."),
    ("2025", "Городской", "Городской конкурс «Гродненчанка – 2025» (ОО «Белорусский союз женщин»)", "Диплом победителя в номинации «ПЕРВАЯ ВИЦЕ-МИСС» (номера со скрипкой и гитарой, эфир «Беларусь 4»)."),
    ("18.09.2025", "Университетский", "Историческая викторина «История земли моей» ко Дню народного единства", "I место (победитель викторины) в составе команды «4G» Факультета экономики и управления."),
    ("2025–2026", "Университетский", "Финал Дебатного турнира ГрГУ имени Янки Купалы", "Диплом II степени финала университетского турнира парламентских дебатов."),
    ("25.04.2025", "Университетский", "Конкурс ораторского мастерства «Цицероний – 2025» (ГрГУ)", "Два диплома: 1) Диплом II степени; 2) Диплом «За оригинальность подачи материала»."),
    ("21–29.11.2024", "Город / ВУЗ", "Творческая гостиная «Беларусь – перекресток национальных культур» & «Смузи Квиз»", "Сертификат участника 7 форматов межкультурной гостиной ГрГУ + Диплом за II место в «Смузи Квиз» («Шутка Фейербаха»)."),
]

for idx, r_vals in enumerate(rows_data):
    row = reg_tbl.add_row()
    bg = "FFFFFF" if idx % 2 == 0 else "FFF5F5"
    for c_i, val in enumerate(r_vals):
        cell = row.cells[c_i]
        cell.width = col_widths[c_i]
        set_cell_bg(cell, bg)
        set_cell_margins(cell, 20, 20, 55, 55)
        cp = cell.paragraphs[0]
        cp.paragraph_format.space_before = Pt(0)
        cp.paragraph_format.space_after = Pt(0)
        is_bold = (c_i in (0, 1))
        color = (30, 58, 138) if c_i == 1 else (15, 23, 42)
        add_run(cp, val, 6.85, is_bold, color=color)

# Miniature scans strip on Page 2
dip_lbl = doc.add_paragraph()
dip_lbl.paragraph_format.space_before = Pt(3)
dip_lbl.paragraph_format.space_after = Pt(1)
add_run(dip_lbl, "КЛЮЧЕВЫЕ ДИПЛОМЫ И МЕДИААКТИВНОСТИ:", 8.0, True, color=(219, 127, 142))

dip_tbl = doc.add_table(rows=1, cols=4)
dip_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
dip_tbl.autofit = False
p2_imgs = [
    ("diplomas/diploma_vdohnovenie_kochanova.jpg", "Диплом Совета Республики НС РБ"),
    ("diplomas/diploma_viola_republic_1st.jpg", "I степень Респ. финал «АРТ-вакацыі»"),
    ("diplomas/diploma_grodnenchanka_vicemiss.jpg", "Первая Вице-Мисс «Гродненчанка»"),
    ("diplomas/diploma_innastart_liblist_3rd.jpg", "III степень «ИнНаСтарт» («LibList»)")
]
for i, (rel_path, cap) in enumerate(p2_imgs):
    cell = dip_tbl.rows[0].cells[i]
    cell.width = Cm(4.65)
    set_cell_bg(cell, "FFF5F5")
    set_cell_margins(cell, 20, 20, 30, 30)
    cp = cell.paragraphs[0]
    cp.paragraph_format.space_before = Pt(0)
    cp.paragraph_format.space_after = Pt(0)
    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cp.add_run().add_picture(os.path.join(ASSETS, rel_path), height=Cm(2.45))
    add_run(cp, "\n" + cap, 6.0, True, color=(96, 77, 83))

# QR Footer Bridge
sp2 = doc.add_paragraph()
sp2.paragraph_format.space_before = Pt(2)
sp2.paragraph_format.space_after = Pt(1)

qr_tbl = doc.add_table(rows=1, cols=3)
qr_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
qr_tbl.autofit = False
qr_tbl.columns[0].width = Cm(11.4)
qr_tbl.columns[1].width = Cm(3.6)
qr_tbl.columns[2].width = Cm(3.6)

qc0, qc1, qc2 = qr_tbl.rows[0].cells
set_cell_bg(qc0, "604D53")
set_cell_bg(qc1, "DB7F8E")
set_cell_bg(qc2, "DB7F8E")
for qc in (qc0, qc1, qc2):
    set_cell_margins(qc, 45, 45, 80, 80)

qp0 = qc0.paragraphs[0]
add_run(qp0, "ЦИФРОВОЕ ПОРТФОЛИО И КОНТАКТЫ\n", 9, True, color=(213, 197, 200))
add_run(qp0, "Интерактивный веб-лендинг содержит все 14 полноразмерных сканов дипломов, 50+ фотографий и 8 видеороликов (выступления на «Студенте года», «Цицеронии», «Гродненчанке» и волонтёрство в приюте для животных).\n", 7.2, False, color=(157, 163, 164))
add_run(qp0, "Telegram: @alexwzzw   |   YouTube: https://youtu.be/bfQRvv3K-Ck", 7.6, True, color=(219, 127, 142))

qp1 = qc1.paragraphs[0]
qp1.alignment = WD_ALIGN_PARAGRAPH.CENTER
qp1.add_run().add_picture(os.path.join(ASSETS, "qr_telegram.png"), width=Cm(1.75))
add_run(qp1, "\nTelegram @alexwzzw", 6.2, True, color=(255, 255, 255))

qp2 = qc2.paragraphs[0]
qp2.alignment = WD_ALIGN_PARAGRAPH.CENTER
qp2.add_run().add_picture(os.path.join(ASSETS, "qr_youtube.png"), width=Cm(1.75))
add_run(qp2, "\nYouTube Выступление", 6.2, True, color=(213, 197, 200))

doc.save(OUT_DOCX)
print("Saved DOCX:", OUT_DOCX)
