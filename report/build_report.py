"""
Сборка отчёта по ЛР1 (вариант 7, Lab1_Dataset.xlsx) в формате .docx по результатам последнего запуска main.py.
Оформление — по «Правилам оформления и требованиям к содержанию КП(Р) и ВКР» СМК ДГТУ.

Запуск (из каталога lab1):  python report/build_report.py
"""
import ast
import glob
import json
import os
import re
from importlib.metadata import version

import pandas as pd
from scipy import stats

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(BASE, "output")
# ФИО студента (титульный лист) и адрес репозитория
STUDENT = "Казаков А. В."  # заменить на ФИО студента
GROUP = "ИПБП41"
VARIANT = 7
DOCX = os.path.join(BASE, f"Отчет_ЛР1_вариант{VARIANT}.docx")
REPO_URL = ""  # при наличии — адрес репозитория, например https://github.com/<логин>/MachineLearningDSTU

R = json.load(open(sorted(glob.glob(os.path.join(OUT, "results_*.json")))[-1], encoding="utf-8"))
TS = R["run_ts"]
IMG = R["image"]


def fig_path(name):
    return os.path.join(OUT, f"{name}_{TS}.png")


# ---------------------------------------------------------------- форматирование чисел
def num(x, d=2):
    """Число для основного текста: десятичная запятая, «минус» вместо знака."""
    s = f"{abs(x):.{d}f}".replace(".", ",")
    return ("минус " + s) if x < 0 and float(s.replace(",", ".")) != 0 else s


def tn(x, d=3):
    """Число для таблиц: десятичная запятая, знак «−»."""
    s = f"{x:.{d}f}".replace(".", ",").replace("-", "−")
    return "0," + "0" * d if s == "−0," + "0" * d else s


def sci(x, d=1):
    """Экспоненциальная запись: 1,1·10⁻¹⁶."""
    if x == 0:
        return "0"
    m, e = f"{x:.{d}e}".split("e")
    sup = str.maketrans("-0123456789", "⁻⁰¹²³⁴⁵⁶⁷⁸⁹")
    return f"{m.replace('.', ',').replace('-', '−')}·10{str(int(e)).translate(sup)}"


# ---------------------------------------------------------------- документ и стили
doc = Document()
sec = doc.sections[0]
sec.page_height, sec.page_width = Cm(29.7), Cm(21.0)
sec.left_margin, sec.right_margin = Cm(3.0), Cm(1.5)
sec.top_margin, sec.bottom_margin = Cm(2.0), Cm(2.0)
sec.different_first_page_header_footer = True


def set_rfonts(el, name):
    rpr = el.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(a), name)
    # шрифты темы (asciiTheme и т. п.) имеют приоритет над явными — удаляем их
    for a in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
        if rfonts.get(qn(a)) is not None:
            del rfonts.attrib[qn(a)]


normal = doc.styles["Normal"]
normal.font.name, normal.font.size = "Times New Roman", Pt(14)
set_rfonts(normal.element, "Times New Roman")
normal.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
normal.paragraph_format.space_after = Pt(0)
normal.paragraph_format.space_before = Pt(0)

for lvl in (1, 2):
    st = doc.styles[f"Heading {lvl}"]
    st.font.name, st.font.size, st.font.bold = "Times New Roman", Pt(14), True
    st.font.italic = False
    st.font.color.rgb = None
    set_rfonts(st.element, "Times New Roman")
    pf = st.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    pf.space_before, pf.space_after = Pt(0), Pt(21)  # одна пустая строка после заголовка
    pf.keep_with_next = True
    pf.first_line_indent = Cm(1.25)
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
# цвет заголовков в шаблоне задан темой — убираем
for st in (doc.styles["Heading 1"], doc.styles["Heading 2"]):
    rpr = st.element.get_or_add_rPr()
    for c in rpr.findall(qn("w:color")):
        rpr.remove(c)

CNT = {"fig": 0, "tab": 0, "eq": 0}


def _runs(p, text, size=None, bold=False):
    """
    Добавляет текст с разметкой: *курсив* (имена функций и библиотек по п. 9.5 Правил),
    _{индекс} — подстрочный, ^{степень} — надстрочный.
    """
    text = re.sub(r"(\d) (%|ч\b|дБ|МиБ|ГиБ)", r"\1 \2", text)  # число не отрывается от единицы
    # после «=», «≈», «<», «>» отрицательное число пишется со знаком «−», а не словом «минус»
    text = re.sub(r"([=≈<>≤≥]\s*)минус\s+", r"\1−", text)
    for part in re.split(r"(\*.+?\*|_\{.+?\}|\^\{.+?\})", text):
        if not part:
            continue
        kind = part[0] if part[0] in "*_^" and len(part) > 2 else ""
        txt = part[1:-1] if kind == "*" else part[2:-1] if kind else part
        r = p.add_run(txt)
        r.italic = kind == "*"
        r.font.subscript = kind == "_"
        r.font.superscript = kind == "^"
        r.bold = bold
        if size:
            r.font.size = Pt(size)
    return p


def para(text, indent=True, align=WD_ALIGN_PARAGRAPH.JUSTIFY, size=None, bold=False, keep=False):
    p = doc.add_paragraph()
    p.alignment = align
    if indent:
        p.paragraph_format.first_line_indent = Cm(1.25)
    p.paragraph_format.keep_with_next = keep
    return _runs(p, text, size, bold)


def bullets(items):
    for it in items:
        para("– " + it)


def page_break():
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


def h1(text, structural=False):
    """Заголовок раздела; структурные элементы — прописными по центру, без номера."""
    p = doc.add_paragraph(style="Heading 1")
    p.paragraph_format.page_break_before = True
    if structural:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Cm(0)
        text = text.upper()
    p.add_run(text)
    return p


def h2(text):
    p = doc.add_paragraph(style="Heading 2")
    p.paragraph_format.space_before = Pt(21)
    p.add_run(text)
    return p


def next_fig():
    CNT["fig"] += 1
    return CNT["fig"]


def figure(path, caption, width=16.0):
    """Рисунок по центру, подпись «Рисунок N — …» 12 pt под ним, пустая строка сверху и снизу."""
    n = CNT["fig"]
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(21)
    p.paragraph_format.keep_with_next = True
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.add_run().add_picture(path, width=Cm(width))
    c = doc.add_paragraph()
    c.alignment = WD_ALIGN_PARAGRAPH.CENTER
    c.paragraph_format.space_after = Pt(21)
    r = c.add_run(f"Рисунок {n} — {caption}")
    r.font.size = Pt(12)


def next_tab():
    CNT["tab"] += 1
    return CNT["tab"]


def _cell_borders(table):
    tbl = table._tbl
    tblPr = tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:color"), "000000")
        borders.append(el)
    tblPr.append(borders)


def table(caption, header, rows, widths=None, size=12):
    """Таблица: «Таблица N — …» над таблицей слева без отступа, шрифт 12 pt."""
    n = CNT["tab"]
    c = doc.add_paragraph()
    c.paragraph_format.space_before = Pt(10)
    c.paragraph_format.keep_with_next = True
    c.add_run(f"Таблица {n} — {caption}")
    t = doc.add_table(rows=1 + len(rows), cols=len(header))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    _cell_borders(t)
    for i, row in enumerate([header] + rows):
        tr = t.rows[i]
        trPr = tr._tr.get_or_add_trPr()
        h = OxmlElement("w:trHeight")
        h.set(qn("w:val"), str(int(Cm(0.8).twips)))
        trPr.append(h)
        if i == 0:  # повтор головки при переносе
            th = OxmlElement("w:tblHeader")
            trPr.append(th)
        for j, val in enumerate(row):
            cell = tr.cells[j]
            cell.text = ""
            p = cell.paragraphs[0]
            p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            p.paragraph_format.keep_with_next = i < len(rows) and len(rows) <= 12
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if (i == 0 or j > 0) else WD_ALIGN_PARAGRAPH.LEFT
            r = p.add_run(str(val))
            r.font.size = Pt(size)
            r.bold = i == 0
            if widths:
                cell.width = Cm(widths[j])
    sp = doc.add_paragraph()
    sp.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    return t


def formula(text):
    """Формула по центру, номер в круглых скобках у правого края; свободная строка сверху и снизу."""
    CNT["eq"] += 1
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(14)
    ts = p.paragraph_format.tab_stops
    ts.add_tab_stop(Cm(8.25), WD_TAB_ALIGNMENT.CENTER)
    ts.add_tab_stop(Cm(16.5), WD_TAB_ALIGNMENT.RIGHT)
    p.add_run("\t")
    _runs(p, text)
    p.add_run(f"\t({CNT['eq']})")
    return CNT["eq"]


def where(lines):
    """Расшифровка: первая строка с «где» без отступа и без двоеточия, далее с отступом."""
    for i, ln in enumerate(lines):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        if i > 0:
            p.paragraph_format.first_line_indent = Cm(1.25)
        _runs(p, ("где " if i == 0 else "") + ln)


def add_field(paragraph, instr):
    r = paragraph.add_run()
    for tag, txt in (("begin", None), (None, instr), ("separate", None), (None, "1"), ("end", None)):
        if tag:
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), tag)
            r._r.append(el)
        elif txt == instr:
            el = OxmlElement("w:instrText")
            el.set(qn("xml:space"), "preserve")
            el.text = instr
            r._r.append(el)
        else:
            el = OxmlElement("w:t")
            el.text = txt
            r._r.append(el)


# ---------------------------------------------------------------- номер страницы
fp = sec.footer.paragraphs[0]
fp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
add_field(fp, "PAGE")

# ================================================================ ТИТУЛЬНЫЙ ЛИСТ
def tline(text, size=12, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER, before=0, single=True):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(before)
    if single:
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    r = p.add_run(text)
    r.font.size, r.bold = Pt(size), bold
    return p


tline("МИНИСТЕРСТВО НАУКИ И ВЫСШЕГО ОБРАЗОВАНИЯ РОССИЙСКОЙ ФЕДЕРАЦИИ", 14)
tline("ФЕДЕРАЛЬНОЕ ГОСУДАРСТВЕННОЕ БЮДЖЕТНОЕ", 14, True, before=6)
tline("ОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ ВЫСШЕГО ОБРАЗОВАНИЯ", 14, True)
tline("«ДОНСКОЙ ГОСУДАРСТВЕННЫЙ ТЕХНИЧЕСКИЙ УНИВЕРСИТЕТ»", 14, True)
tline("(ДГТУ)", 14, True)
tline("Кафедра «ИПМ Ростсельмаш»", 12, before=30)
tline("ОТЧЁТ", 14, True, before=60)
tline("по лабораторной работе № 1", 12, before=6)
tline("по дисциплине «Машинное обучение»", 12)
tline("на тему:", 12, before=12)
tline("«БИБЛИОТЕКИ PYTHON ДЛЯ МАШИННОГО ОБУЧЕНИЯ.", 14)
tline("РАБОТА С ДАННЫМИ»", 14)
tline("Направление подготовки 09.03.02 Информационные системы и технологии", 12, before=18)

RIGHT = Cm(9.5)
for i, (txt, gap) in enumerate([("Выполнил:", 50), (f"студент {STUDENT}", 0), (f"Группа {GROUP}", 0),
                                (f"Вариант {VARIANT}", 0), ("Проверил:", 24), ("Дольников М.Г.", 0)]):
    p = tline(txt, 12, align=WD_ALIGN_PARAGRAPH.LEFT, before=gap, single=False)
    p.paragraph_format.left_indent = RIGHT
tline("Ростов-на-Дону", 12, before=60)
tline("2026", 12)

# ================================================================ СОДЕРЖАНИЕ
page_break()
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(21)
r = p.add_run("СОДЕРЖАНИЕ")
r.bold = True
add_field(doc.add_paragraph(), r'TOC \o "1-2" \h \z \u')


# ---------------------------------------------------------------- данные варианта для текста
VC = ["x", "y"]
VERSIONS = ", ".join(f"{name} {'.'.join(version(pkg).split('.')[:2])}" for name, pkg in [
    ("NumPy", "numpy"), ("pandas", "pandas"), ("SciPy", "scipy"), ("scikit-learn", "scikit-learn"),
    ("Matplotlib", "matplotlib"), ("TensorFlow", "tensorflow"), ("Keras", "keras"), ("PyTorch", "torch")])
REPO_TEXT = (f"в репозитории {REPO_URL} (каталог lab1)" if REPO_URL else "в каталоге lab1 репозитория работы")


def vec(v, d=3):
    return "; ".join(num(x, d) for x in v)


def pv(p):
    return "< 0,001" if p < 1e-3 else tn(p, 3)


st_ = {s["col"]: s for s in R["stats"]}
crg = R["col_range"]
q = R["quality"]
rs_ = R["resid"]
N = R["X_shape"][0]
G1, G2 = R["gaps_excel"]
GN1, GN2 = R["gaps_n"]
sheet = R["sheet"]
bad_vars = [s for s in sheet if s["non_numeric"]]
miss_vars = [s for s in sheet if s["missing"]]

# ================================================================ ВВЕДЕНИЕ
h1("Введение", structural=True)
para("Подготовка данных занимает значительную часть любой задачи машинного обучения: модели "
     "обучаются только на корректных числовых массивах, а пропуски, ошибки ввода и несовпадение "
     "типов напрямую ухудшают результат. Поэтому умение читать, проверять, очищать, "
     "масштабировать и статистически описывать данные средствами стандартных библиотек Python "
     "является базовым навыком для дальнейших лабораторных работ.")
para("Цель работы — изучить основные библиотеки Python для машинного обучения (*NumPy*, "
     "*pandas*, *Matplotlib*, *SciPy*, *scikit-learn*, *TensorFlow*, *Keras*, *PyTorch*) и "
     "приобрести практические навыки работы с табличными данными: чтения и записи, "
     "предобработки, статистического анализа, тензорных вычислений и обработки изображений [1].")
para("Для достижения цели были поставлены следующие задачи:")
bullets([
    "освоить чтение и запись табличных данных в форматах .xlsx, .csv, .txt и .mat (задания 1–3);",
    "извлечь данные варианта, выявить и обработать проблемы в них (задания 4–8);",
    "реализовать собственные функции масштабирования и разбиения данных и сравнить их со "
    "*scikit-learn* (задания 9–13, 34–35);",
    "выполнить статистический и спектральный анализ данных средствами *NumPy*, *pandas* и "
    "*SciPy* (задания 14–33, 36–37);",
    "освоить работу с тензорами *TensorFlow*, *PyTorch* и *Keras* (задания 38–45);",
    "выполнить матричные операции и построить пайплайн с методом главных компонент "
    "(задания 46–49);",
    "выполнить обработку цветного изображения средствами *SciPy* (задание 50).",
])
para(f"Исходные данные — файл Lab1_Dataset.xlsx: лист из {R['raw_shape'][0]} строк и "
     f"{R['raw_shape'][1]} столбцов, в котором по два столбца отведено каждому из "
     f"{R['n_variants']} вариантов. Первая строка листа содержит надпись «Вариант N», вторая — "
     "заголовки столбцов, далее идут значения; длина столбцов у вариантов разная. Данным "
     f"варианта 7 соответствуют столбцы {R['variant_cols'][0]} и {R['variant_cols'][1]} листа (M и "
     "N) с заголовками x и y. Их описание приведено в таблице 1.")
next_tab()
table("Состав данных варианта 7",
      ["Столбец", "Строки листа", "Содержание"],
      [["x", f"3–{N + 2}", f"{N} значений, из них {q['missing']['x']['count']} пустые ячейки (строки "
                           f"{G1} и {G2})"],
       ["y", f"3–{N + 2}", f"{N} значений без пропусков"]],
      widths=[2.5, 3.0, 11.0])
para("Работа выполнена на Python 3.13 с библиотеками " + VERSIONS + " [2–9]. "
     f"Исходный код и все файлы результатов размещены {REPO_TEXT}. Программа разбита на "
     "модули по заданиям методических указаний (каталог src), точка входа — main.py. Все "
     "рисунки и таблицы, создаваемые программой, сохраняются в каталог output с датой и временем "
     "запуска в имени файла, полный консольный вывод записывается в файл run_log. Генераторы "
     "случайных чисел зафиксированы (SEED = 42), поэтому повторный запуск даёт те же значения.")

# ================================================================ 1
h1("1 Считывание, проверка и запись данных")
h2("1.1 Чтение и запись файлов")
para("Для чтения реализована функция *read_table* (модуль io_utils.py), которая по расширению "
     "файла выбирает способ чтения: *pandas.read_excel* для .xlsx, *pandas.read_csv* для .csv, "
     "*read_csv* с разделителем «пробелы или табуляция» для .txt и *scipy.io.loadmat* для .mat. "
     "Наличие заголовков определяется по первой строке: если все её значения числовые, "
     "заголовков в файле нет, и столбцам автоматически присваиваются имена col_1, col_2 и т. д. "
     "Лист Lab1_Dataset.xlsx не имеет единой строки заголовков, поэтому он читается целиком "
     f"без заголовков: получена таблица {R['raw_shape'][0]} × {R['raw_shape'][1]} со столбцами "
     f"col_1…col_{R['raw_shape'][1]}.")
para("Функция *write_table* (задание 2) сохраняет таблицу сразу в четыре формата с "
     "заголовками, добавляя к имени файла дату и время запуска (например, "
     f"task2_variant7_clean_{TS}.csv); в файл .mat записываются массив значений data и список "
     "имён столбцов columns. Функция *save_figure* (задание 3) аналогично сохраняет рисунки в "
     "формате PNG. Для проверки очищенные данные варианта (после задания 8) были записаны во все "
     "форматы и прочитаны обратно, а также записаны без заголовков. Результат приведён в "
     f"таблице {CNT['tab'] + 1}.")
next_tab()
rt, nh = R["roundtrip"], R["noheader"]
table("Проверка записи и обратного чтения файлов",
      ["Формат", "Файл с заголовками: прочитанные столбцы", "Файл без заголовков: прочитанные столбцы",
       "Совпадение значений"],
      [[k, f"{', '.join(rt[k]['columns'])} ({rt[k]['shape'][0]} × {rt[k]['shape'][1]})",
        f"{', '.join(nh[k]['columns'])} ({nh[k]['shape'][0]} × {nh[k]['shape'][1]})",
        "да" if rt[k]["equal"] and nh[k]["equal"] else "нет"] for k in rt],
      widths=[2.0, 5.5, 5.5, 3.5])
para("Во всех четырёх форматах заголовки x и y сохранились и были прочитаны обратно, а для "
     "файлов без заголовков столбцы получили автоматические имена col_1 и col_2; значения во "
     "всех случаях совпали с исходными, то есть запись и чтение выполняются без потерь.")

h2("1.2 Извлечение данных варианта и проверка качества")
para("Функция *extract_variant* (задание 6) находит в первой строке листа ячейку «Вариант 7», "
     "берёт два столбца, начиная с неё, отбрасывает строку «Вариант N» (заголовок таблицы) и "
     "данные остальных вариантов, а строку с именами x и y использует как заголовки столбцов. "
     "Конец данных определяется автоматически — по последней непустой строке, поэтому "
     f"извлечены все {N} значений (строки листа 3–{N + 2}). Функция возвращает копию: проверено, "
     "что таблица листа не изменилась, а контрольная сумма MD5 файла до и после чтения совпадает, "
     "то есть исходный файл не изменяется. Функция *show_table* (задание 5) выводит на экран "
     "содержимое таблицы полностью — и всего листа, и данных варианта; вывод сохранён в файле "
     "run_log.")
para("Функция *check_data_quality* (задание 4) для каждого столбца находит пустые ячейки, "
     "значения, которые нельзя преобразовать в число, бесконечные значения и нечисловой тип "
     "столбца, и при нахождении выдаёт предупреждение средствами модуля *warnings*. Номера строк "
     "приводятся в нумерации листа Excel. Для варианта 7 получены предупреждения:")
bullets([w.replace("[", "").replace("]", "") + (";" if i < len(R["warnings"]) - 1 else ".")
         for i, w in enumerate(R["warnings"])])
para("Несовпадение типов возникает потому, что в столбцах листа, кроме чисел, находятся "
     "надписи «Вариант 7», «x» и «y», и *pandas* хранит такие столбцы с типом object; после "
     "отделения заголовков значения остаются объектами Python, хотя все они числовые. "
     "Нечисловых и бесконечных значений в данных варианта 7 нет. Чтобы проверить, что функция "
     "находит и некорректные значения, она была применена ко всем 24 вариантам листа "
     f"(таблица {CNT['tab'] + 1}, файл task4_sheet_scan_*.xlsx).")
next_tab()
table("Проблемы в данных всех вариантов листа",
      ["Вариант", "Строк", "Пустых ячеек", "Нечисловых значений", "Нечисловые значения"],
      [[s["variant"], s["rows"], s["missing"], s["non_numeric"],
        ", ".join(f"«{v}»" for v in s["values"]) if s["values"] else "–"] for s in sheet],
      widths=[2.0, 2.0, 3.0, 3.7, 5.3], size=10)
para(f"Пустые ячейки найдены в {len(miss_vars)} вариантах из 24, нечисловые значения «ERROR» и "
     f"«-» — в {len(bad_vars)} вариантах (например, в варианте {R['bad_demo']['variant']} в строке "
     f"{R['bad_demo']['rows_excel'][0]} столбца {R['bad_demo']['col']}). У варианта 7 только "
     f"пустые ячейки — {q['missing']['x']['count']} в столбце x.")

h2("1.3 Явные типы данных")
para("Функция *cast_types* (задание 7) задаёт типы столбцов по словарю. До приведения оба "
     f"столбца имели тип {R['dtypes_before']['x']}, после приведения — "
     f"{R['dtypes_after']['x']} (64-разрядное число с плавающей точкой, тип *pandas* float64). "
     "Тип float64 выбран потому, что значения непрерывные, а пустые ячейки этот тип хранит как "
     "NaN.")

# ================================================================ 2
h1("2 Предобработка данных")
h2("2.1 Обработка проблем в данных")
para("Функция *handle_missing* (задание 8) реализует средствами *pandas* все действия из "
     "методических указаний: удаление строк с некорректными данными (drop_bad), приведение "
     "типов (cast, функция *pd.to_numeric*), удаление строк с пропусками (drop), заполнение "
     "средним (mean), предыдущим (ffill) или последующим (bfill) значением, нулями (zero) и "
     "указанным значением (value) — во всех пропусках или в одной указанной ячейке. Функция "
     "*interactive_menu* (команда «python main.py --interactive») показывает найденные проблемы и "
     "для каждой предлагает только подходящие действия: для некорректных данных и несовпадения "
     "типов — удаление строк и приведение типов, для пустых ячеек — удаление и способы "
     "заполнения. На варианте "
     f"{R['bad_demo']['variant']} действие drop_bad сократило таблицу с {R['bad_demo']['n']} до "
     f"{R['bad_demo']['drop_bad_rows']} строк, а cast перевело столбец в тип "
     f"{R['bad_demo']['cast_dtype']}, заменив значение «{R['bad_demo']['values'][0]}» пустой ячейкой. "
     "Результат действий для пропусков столбца x варианта 7 приведён в таблице "
     f"{CNT['tab'] + 1}.")
next_tab()
hd = R["handle_demo"]
table("Действия обработки пропусков в столбце x",
      ["Действие", "Строк", "Пропусков", "Значения на месте пропусков", "Отклонение от √n"],
      [[d["action"], d["rows"], d["missing"], "; ".join(tn(v, 4) for v in d["values"]) or "–",
        tn(d["err"], 4) if d["err"] is not None else "–"] for d in hd], widths=[2.4, 1.8, 2.4, 5.8, 3.6])
para(f"Проверено, что все {N - 2} известных значений x в точности равны √n, где n — номер "
     "измерения: x — гладкий монотонно возрастающий ряд (коэффициент автокорреляции соседних "
     f"значений {num(R['x_autocorr'], 5)}). Поэтому способ обработки выбран по характеру ряда:")
bullets([
    "несовпадение типов устранено приведением обоих столбцов к числу (cast) и к типу float64;",
    f"пропуски x заполнены предыдущим значением (ffill) — {num(R['gap_filled'][0], 4)} при "
    f"следующем известном значении {num(R['gap_neighbors'][1], 4)}; отклонение от точных "
    f"значений √{GN1} = {num(R['gap_true'][0], 4)} и √{GN2} = {num(R['gap_true'][1], 4)} не "
    f"превышает {num(hd[2]['err'], 3)}. Заполнение средним ({num(R['x_mean_before'], 2)}) дало бы "
    f"посреди ряда провал на {num(hd[1]['err'], 2)}, нулём или константой — выбросы около 11–12 "
    f"единиц, а удаление строк сократило бы выборку до {hd[0]['rows']} значений и нарушило "
    "равномерный шаг по номеру измерения.",
])
para(f"После обработки пропусков не осталось, объём выборки сохранился — {N} строк, оба столбца "
     "имеют тип float64.")

h2("2.2 Перевод в NumPy и масштабирование")
sc = R["scale"]
para(f"Функция *to_numpy* (задание 9) переводит таблицу в массив X размером {N} × 2 типа "
     "float64 методом *DataFrame.to_numpy*. Для задания 10 реализованы классы "
     "*MinMaxScalerCustom* и *StandardScalerCustom* (листинг А.1 приложения А) с методами fit, "
     "transform и inverse_transform. MinMax-нормализация выполняется по формуле")
formula("*x*_{new} = (*x* − min(*x*)) / (max(*x*) − min(*x*)),")
where(["*x* — исходное значение признака;", "min(*x*), max(*x*) — минимальное и максимальное "
       "значения признака."])
para("MinMax-нормализация в произвольных границах [a, b] выполняется по формуле")
formula("*x*_{new} = *a* + (*x* − min(*x*))·(*b* − *a*) / (max(*x*) − min(*x*)),")
where(["*a*, *b* — границы нового диапазона (при a = 0, b = 1 формула совпадает с формулой (1))."])
para("Стандартизация выполняется по формуле")
formula("*z* = (*x* − *μ*) / *σ*,")
where(["*μ* — выборочное среднее признака;",
       "*σ* — выборочное среднеквадратическое отклонение (с делителем n, как в *scikit-learn*)."])
p_ = R["scale_params"]
para("Значения, используемые при масштабировании, сохраняются функцией *save_scaler_params* "
     f"в файл {R['scale_params_file'].replace(TS, '*')}; они приведены в таблице "
     f"{CNT['tab'] + 1}.")
next_tab()
table("Сохранённые параметры масштабирования",
      ["Параметр", "x", "y"],
      [["min (MinMax)", tn(p_["minmax_0_1"]["min"][0], 4), tn(p_["minmax_0_1"]["min"][1], 4)],
       ["max (MinMax)", tn(p_["minmax_0_1"]["max"][0], 4), tn(p_["minmax_0_1"]["max"][1], 4)],
       ["μ (z-оценка)", tn(p_["zscore"]["mean"][0], 4), tn(p_["zscore"]["mean"][1], 4)],
       ["σ (z-оценка)", tn(p_["zscore"]["std"][0], 4), tn(p_["zscore"]["std"][1], 4)]],
      widths=[5, 4, 4])
para("После MinMax-нормализации оба столбца лежат точно в [0, 1], в границах a = минус 1, "
     "b = 1 — точно в [минус 1, 1]; после стандартизации средние равны нулю с точностью до "
     f"{sci(max(abs(v) for v in sc['zscore']['mean']) or 1e-16, 0)}, среднеквадратические "
     "отклонения — единице. Функции обратного преобразования (задание 11) восстанавливают "
     "исходный масштаб по параметрам, прочитанным из сохранённого файла: максимальная ошибка "
     f"восстановления {sci(sc['minmax_0_1']['inv_err'])} для MinMax, "
     f"{sci(sc['minmax_m1_1']['inv_err'])} для границ [минус 1, 1] и {sci(sc['zscore']['inv_err'])} "
     "для стандартизации, то есть на уровне погрешности вычислений с плавающей точкой.")

h2("2.3 Разбиение данных и фиксация seed")
sp = R["split"]
para("Функция *split_train_val_test* (задание 12, листинг А.1) делит массив на три части в "
     "заданном соотношении, записанном в процентах, долях или частях: соотношение нормируется "
     "на сумму, размеры первых двух частей округляются, третья часть получает остаток. Части "
     "возвращаются в словаре с метками _train, _val и _test. Проверка на примерах из "
     f"методических указаний и на данных варианта приведена в таблице {CNT['tab'] + 1}.")
next_tab()
v_ = lambda d: [d[k] for k in d]
table("Результаты разбиения массивов",
      ["Массив", "Соотношение", "_train", "_val", "_test"],
      [["150 элементов", "1 : 1 : 1"] + v_(sp["a150"]),
       ["100 элементов", "75 % : 10 % : 15 %"] + v_(sp["a100"]),
       [f"X ({N} строк)", "70 % : 15 % : 15 %"] + v_(sp["pct"]),
       [f"X ({N} строк), с перемешиванием", "60 : 20 : 20"] + v_(sp["shuffle"])],
      widths=[5.2, 4.3, 2.3, 2.3, 2.3])
para("Разбиения в процентах (70; 15; 15) и в долях (0,70; 0,15; 0,15) дали одинаковые массивы. По "
     "умолчанию порядок строк сохраняется, так как данные упорядочены по номеру измерения; с "
     "параметром shuffle = True строки предварительно перемешиваются функцией "
     "*np.random.permutation*.")
para("Функция *set_seeds* (задание 13, модуль config.py) фиксирует генераторы *random*, "
     "*NumPy*, *TensorFlow* и *PyTorch* и переменную PYTHONHASHSEED; в функциях *scikit-learn* "
     "задаётся random_state = 42. Проверка: два вызова np.random.rand(3) после set_seeds(42) "
     "вернули одинаковые массивы, повторный запуск программы даёт те же результаты.")

# ================================================================ 3
h1("3 Статистический анализ данных")
h2("3.1 Визуализация исходных данных")
n = next_fig()
para(f"График исходных данных (задание 14) приведён на рисунке {n}: слева — ряды x и y в "
     "зависимости от номера измерения n, справа — диаграмма рассеяния y(x) с прямой y = x.",
     keep=True)
figure(fig_path("task14_series"), "Ряды x и y и диаграмма рассеяния y(x)", 16.0)
para(f"Ряд x — гладкая кривая √n от {num(crg['x'][0], 3)} до {num(crg['x'][1], 3)}: рост "
     "замедляется с увеличением n. Красными кружками отмечены значения, заполненные в "
     f"задании 8 (n = {GN1} и {GN2}); на графике они образуют едва заметную «ступеньку». Ряд y "
     f"повторяет форму x с шумом: y изменяется от {num(crg['y'][0], 3)} до {num(crg['y'][1], 3)}, "
     f"и во всех {rs_['y_gt_x']} точках y > x. Отклонение y − √n лежит в пределах от "
     f"{num(rs_['min'], 3)} до {num(rs_['max'], 3)}, его среднее {num(rs_['mean'], 3)}, "
     f"среднеквадратическое отклонение {num(rs_['std'], 3)}, а коэффициент автокорреляции соседних "
     f"значений {num(rs_['ac1'], 3)}, то есть шум не зависит от n и от предыдущих значений. На "
     "диаграмме рассеяния точки лежат в узкой полосе выше прямой y = x, что означает почти "
     "линейную связь y ≈ x + 0,5 с постоянным разбросом.")
n = next_fig()
para(f"Гистограммы частот и нормализованные гистограммы (задание 15) с 15 интервалами и "
     f"ядерной оценкой плотности *scipy.stats.gaussian_kde* приведены на рисунке {n}. При "
     "нормализации (density = True) высота столбца равна частоте, делённой на объём выборки и "
     "ширину интервала, поэтому площадь гистограммы равна единице и её можно сравнивать с "
     "плотностью распределения.", keep=True)
figure(fig_path("task15_histograms"), "Гистограммы частот и оценки плотности распределения x и y")
hb = R["hist"]
para("Распределение x — возрастающее, почти треугольное: число значений в интервале растёт с "
     f"3 в первом до 20 в последнем, плотность — с {num(hb['x']['first'], 3)} до "
     f"{num(hb['x']['last'], 3)}. Это прямое следствие закона x = √n: при равномерном шаге по n "
     "на интервал значений [x; x + Δx] приходится около 2x·Δx номеров, то есть плотность "
     "пропорциональна x. Распределение y имеет ту же форму, размытую шумом: максимум (по 20 "
     "значений) приходится на интервалы около 11–13, а последний интервал ниже соседних, так как "
     "значения y > 13 бывают только при наибольших n. Ядерная оценка на правом краю резко "
     "опускается: это особенность метода, который сглаживает обрыв плотности на границе данных.")
para("Функция *sorted_columns* (задание 16) выводит на экран каждый столбец, отсортированный по "
     "возрастанию функцией *np.sort*; результат сохранён в файл task16_sorted_columns_*.xlsx. "
     "Столбец x уже упорядочен (сортировка его не изменила), а в столбце y из-за шума "
     f"{R['y_sorted_changed']} значений из {N} изменили позицию. Наименьшие и наибольшие значения: "
     f"x — {num(crg['x'][0], 3)} и {num(crg['x'][1], 3)}, y — {num(crg['y'][0], 3)} и "
     f"{num(crg['y'][1], 3)}.")
ec = R["ecdf"]
n = next_fig()
para(f"Эмпирические функции распределения (задание 17), построенные функцией *scipy.stats.ecdf*, "
     f"показаны на рисунке {n}.", keep=True)
figure(fig_path("task17_ecdf"), "Эмпирические функции распределения x и y")
para("Функция распределения x — гладкая выпуклая вниз кривая, близкая к параболе: для x = √n "
     "доля значений, не превышающих t, равна t²/166. Поэтому в нижнюю половину диапазона "
     f"(до середины {num(ec['x']['mid'], 2)}) попадает только {num(100 * ec['x']['F_mid'], 1)} % "
     f"значений, для y — {num(100 * ec['y']['F_mid'], 1)} % (середина {num(ec['y']['mid'], 2)}). "
     "Небольшой скачок функции x около 11,3 соответствует трём одинаковым значениям, полученным "
     "при заполнении пропусков. Функция y повторяет форму функции x, но ступенчатая из-за шума.")

h2("3.2 Описательные статистики и доверительные интервалы")
para("Функция *column_stats* (задание 18) вычисляет средствами *NumPy* выборочное среднее "
     "(*np.mean*), выборочную дисперсию с делителем n − 1 (*np.var*, ddof = 1), медиану "
     "(*np.median*) и моду (*scipy.stats.mode*); функция *pandas_stats* вычисляет те же величины "
     "методами *DataFrame.mean*, *var*, *median* и *mode*. Результаты *NumPy* и *pandas* совпали. "
     "Значения непрерывные, поэтому дополнительно вычислена мода интервального ряда — середина "
     "интервала гистограммы с наибольшей частотой. Доверительный интервал для математического "
     "ожидания при неизвестной дисперсии (задание 19) вычисляется по формуле")
formula("*x̄* ± *t*_{1−α/2; n−1} · *s* / √*n*,")
where(["*x̄* — выборочное среднее;", "*t*_{1−α/2; n−1} — квантиль распределения Стьюдента;",
       "*s* — выборочное среднеквадратическое отклонение;",
       f"*n* = {N} — объём выборки; α = 0,05."])
para("Доверительный интервал для дисперсии вычисляется по формуле")
formula("[(*n* − 1)*s*² / χ²_{1−α/2; n−1};  (*n* − 1)*s*² / χ²_{α/2; n−1}],")
where(["χ²_{1−α/2; n−1}, χ²_{α/2; n−1} — квантили распределения хи-квадрат."])
para(f"Результаты приведены в таблице {CNT['tab'] + 1} и сохранены в файл task18_19_stats_*.xlsx.")
next_tab()
table("Описательные статистики и доверительные интервалы (α = 0,05)",
      ["Столбец", "Среднее", "Дисперсия", "Медиана", "Мода (частота)", "Мода интервала",
       "ДИ среднего", "ДИ дисперсии"],
      [[s["col"], tn(s["mean"]), tn(s["var"]), tn(s["median"]), f"{tn(s['mode'])} ({s['mode_count']})",
        tn(s["mode_interval_mid"]), f"{tn(s['ci_mean'][0])} … {tn(s['ci_mean'][1])}",
        f"{tn(s['ci_var'][0])} … {tn(s['ci_var'][1])}"] for s in R["stats"]],
      widths=[1.6, 1.7, 1.9, 1.7, 2.4, 1.8, 2.7, 2.9], size=10)
sx, sy = st_["x"], st_["y"]
n = next_fig()
para(f"Среднее, медиана, мода интервального ряда и интервал «среднее ± СКО» отображены на "
     f"гистограммах на рисунке {n}.", keep=True)
figure(fig_path("task18_stats_plot"), "Среднее, медиана и мода на гистограммах x и y")
para(f"У обоих столбцов медиана больше среднего (для x {num(sx['median'], 3)} и "
     f"{num(sx['mean'], 3)}, для y {num(sy['median'], 3)} и {num(sy['mean'], 3)}), а мода "
     f"интервального ряда лежит у правого края ({num(sx['mode_interval_mid'], 2)} и "
     f"{num(sy['mode_interval_mid'], 2)}): распределения скошены влево, длинный «хвост» малых "
     "значений тянет среднее вниз. Для y два соседних интервала имеют одинаковую наибольшую "
     "частоту 20, на рисунке отмечен первый из них. Дисперсии почти равны "
     f"({num(sx['var'], 3)} и {num(sy['var'], 3)}): шум y мал по сравнению с размахом тренда. "
     "Точная мода у непрерывных данных малоинформативна: в y все значения различны, и функция "
     f"возвращает наименьшее ({num(sy['mode'], 3)}), а мода x {num(sx['mode'], 4)} встречается "
     f"{sx['mode_count']} раза — это значение √128, скопированное в пропуски n = {GN1} и {GN2}. "
     f"Среднее y больше среднего x на {num(sy['mean'] - sx['mean'], 3)}, что совпадает со средним "
     "отклонением y − √n. Доверительные интервалы средних x и y пересекаются, а интервалы "
     "дисперсий практически совпадают.")

h2("3.3 Корреляционный анализ")
cr, cv_ = R["corr"], R["cov"]
sg = R["sig"][0]
n = next_fig()
para("Ковариационная и корреляционная матрицы (задание 20) вычислены функциями *np.cov* и "
     "*np.corrcoef*; результаты совпали с методами *DataFrame.cov* и *DataFrame.corr* и "
     f"сохранены в файл task20_21_cov_corr_*.xlsx. Обе матрицы показаны на рисунке {n}.", keep=True)
figure(fig_path("task20_cov_corr"), "Ковариационная и корреляционная матрицы столбцов x и y", 15.0)
para(f"Ковариация x и y равна {num(cv_['y']['x'], 4)} — почти столько же, сколько дисперсии "
     f"столбцов ({num(cv_['x']['x'], 4)} и {num(cv_['y']['y'], 4)}), коэффициент корреляции "
     f"Пирсона r = {num(cr['y']['x'], 4)}. Связь практически функциональная: y = x + шум, и почти "
     f"вся дисперсия y объясняется x (r² = {num(cr['y']['x'] ** 2, 4)}).")
para("Значимость коэффициента корреляции (задание 21) проверена функцией "
     "*scipy.stats.pearsonr*: нулевая гипотеза — истинный коэффициент равен нулю, статистика")
formula("*t* = *r* · √(*n* − 2) / √(1 − *r*²)")
para(f"имеет распределение Стьюдента с n − 2 = {N - 2} степенями свободы. Дополнительно вычислен "
     f"ранговый коэффициент Спирмена (*scipy.stats.spearmanr*). Результаты приведены в таблице "
     f"{CNT['tab'] + 1}.", indent=False)
next_tab()
table("Проверка значимости коэффициента корреляции x–y",
      ["Показатель", "Значение"],
      [["r (Пирсон)", tn(sg["r"], 4)], ["t", tn(sg["t"], 2)],
       ["Критическое значение t при α = 0,05", tn(sg["t_crit"], 3)],
       ["p-значение", "< 10⁻¹⁰⁰"], ["95-процентный ДИ для r", f"{tn(sg['ci_low'], 4)} … {tn(sg['ci_high'], 4)}"],
       ["ρ (Спирмен)", tn(sg["spearman"], 4)]], widths=[9, 5])
para(f"Статистика t = {num(sg['t'], 1)} многократно превышает критическое значение "
     f"{num(sg['t_crit'], 3)}, p-значение имеет порядок 10⁻¹⁷⁴, поэтому гипотеза о незначимости "
     "отвергается: корреляция x и y статистически значима, доверительный интервал коэффициента "
     f"узкий ({num(sg['ci_low'], 4)} … {num(sg['ci_high'], 4)}). При этом отклонение y − √n с x "
     f"не коррелирует (r = {num(R['corr_e_x'][0], 3)}, p = {tn(R['corr_e_x'][1], 3)}), то есть шум "
     "аддитивный и не зависит от уровня сигнала.")
xc = R["xcorr"]
a_, d_ = xc["x–y"], xc["Δx–Δy"]
n = next_fig()
para("Взаимная корреляционная функция (задание 22) вычислена функцией *scipy.signal.correlate* "
     "с нормировкой на произведение среднеквадратических отклонений и объём выборки; сдвиги "
     "получены функцией *scipy.signal.correlation_lags*. Функция построена для исходных столбцов "
     "x–y и, чтобы исключить влияние общего тренда, для их первых разностей Δx–Δy при сдвигах от "
     f"минус 40 до 40 элементов (рисунок {n}).", keep=True)
figure(fig_path("task22_xcorr"), "Взаимная корреляционная функция x–y и Δx–Δy")
para(f"Для исходных рядов максимум {num(a_['value'], 3)} достигается при нулевом сдвиге, затем "
     f"функция убывает почти линейно: {num(a_['at10'], 3)} при сдвиге 10 и около "
     f"{num(a_['at40'], 2)} при сдвиге 40. Медленное убывание вызвано общим монотонным трендом: "
     "сдвинутый на несколько элементов ряд почти не отличается от исходного. Для первых "
     f"разностей тренд исключён, и функция мала: максимум {num(d_['value'], 3)} при нулевом "
     "сдвиге, при остальных сдвигах модуль не превышает "
     f"{num(d_['max_abs_nonzero'], 3)}. Это значит, что быстрые изменения y определяются в "
     "основном шумом, а x и y согласованы между собой через тренд без запаздывания.")

h2("3.4 Производная, свёртка, произведения и нормы")
g = R["grad"]
n = next_fig()
para("Численная производная (задание 23) вычислена функцией *np.diff* (разность соседних "
     f"значений, {g['diff_len']} элементов для каждого столбца), градиент — функцией "
     "*np.gradient* (центральные разности внутри ряда и односторонние на концах) по оси строк "
     f"массива X, результат имеет размер {N} × 2. Для x = √n производная известна точно: "
     f"dx/dn = 1/(2√n). Результаты показаны на рисунке {n}.", keep=True)
figure(fig_path("task23_gradient"), "Численная производная и градиент столбцов x и y")
para(f"Градиент x совпадает с теоретической кривой 1/(2√n): вне концов ряда и окрестности "
     f"заполненных пропусков наибольшее отклонение равно {num(g['x_err_inner'], 4)} (при n = 2, "
     "где кривизна наибольшая); уже при n = 21 оно меньше 10⁻⁴, а при n = 101 — меньше 10⁻⁶. Градиент убывает от "
     f"{num(g['x_first'][0], 3)} при n = 1 до {num(g['x_last'], 3)} при n = {N}. На месте "
     f"заполненных пропусков ряд на два шага постоянен, поэтому градиент там падает до "
     f"{num(g['x_at_gaps'][1], 1)}, а на следующем шаге скачок выдаёт выброс до "
     f"{num(g['x_at_gaps'][3], 3)}. Производная y определяется шумом: градиент колеблется от "
     f"{num(g['y_min'], 3)} до {num(g['y_max'], 3)} со среднеквадратическим отклонением "
     f"{num(g['y_std'], 3)}, и в {g['y_neg']} из {g['diff_len']} разностей y убывает, хотя тренд "
     f"возрастает. Средний градиент y ({num(g['y_mean'], 4)}) близок к среднему градиенту x "
     f"({num(g['x_mean'], 4)}).")
cvv = R["conv"]
n = next_fig()
para(f"Свёртка столбцов (задание 24) x ∗ y вычислена функциями *np.convolve* и "
     f"*scipy.signal.convolve* в режиме full: длина результата {cvv['len']} = {N} + {N} − 1, "
     f"расхождение двух реализаций равно {sci(cvv['diff']) if cvv['diff'] else 'нулю'}. Кроме того, "
     f"ряд y свёрнут с прямоугольным ядром из {cvv['width']} одинаковых коэффициентов "
     f"1/{cvv['width']} в режиме valid ({cvv['len_valid']} значений), что соответствует "
     f"сглаживанию. Результаты показаны на рисунке {n}.", keep=True)
figure(fig_path("task24_convolution"), "Свёртка столбцов x ∗ y и сглаживание y прямоугольным ядром")
para("Элемент свёртки с номером k — сумма произведений x_{i}·y_{k−i} по всем допустимым i. "
     "Результат возрастает, пока перекрытие рядов увеличивается, достигает максимума "
     f"{num(cvv['max'], 0)} при k = {cvv['k_max']} и убывает до {num(cvv['c_last'], 2)} "
     "на конце; кривая несимметрична, так как оба ряда возрастают и наибольшие произведения дают "
     f"концы рядов. Сглаживание ядром ширины {cvv['width']} уменьшает отклонение y от тренда "
     f"√n + {num(rs_['mean'], 3)} с {num(cvv['noise_before'], 3)} до {num(cvv['noise_after'], 3)}, "
     f"то есть примерно в √{cvv['width']} ≈ {num(cvv['width'] ** 0.5, 2)} раза, как и следует для "
     "некоррелированного шума; форма тренда при этом сохраняется.")
pr = R["products"]
nm = R["norms"]
para(f"Скалярное произведение столбцов x и y (задание 25, *np.dot*) равно {num(pr['dot'], 2)}; "
     f"косинус угла между векторами {num(pr['cos'], 4)}, угол {num(pr['angle_deg'], 2)}°, то есть "
     "векторы почти сонаправлены. Векторное произведение (*np.cross*) определено для "
     "трёхмерных векторов, поэтому оно вычислено для первых трёх элементов столбцов: "
     f"({vec(pr['a'])}) × ({vec(pr['b'])}) = ({vec(pr['cross'])}). Скалярные произведения "
     "результата с сомножителями равны нулю с точностью до "
     f"{sci(max(abs(v) for v in pr['cross_orth']), 0)}, то есть полученный вектор ортогонален "
     "обоим. Нормы L1 (сумма модулей) и L2 (евклидова длина) столбцов (задание 26, "
     f"*np.linalg.norm*) приведены в таблице {CNT['tab'] + 1}.")
next_tab()
table("Нормы L1 и L2 столбцов", ["Столбец", "L1", "L2"],
      [[c, tn(v["L1"], 3), tn(v["L2"], 3)] for c, v in nm.items()], widths=[3, 5, 5])
para(f"Нормы y больше норм x на {num(100 * (nm['y']['L1'] / nm['x']['L1'] - 1), 1)} % (L1) и "
     f"{num(100 * (nm['y']['L2'] / nm['x']['L2'] - 1), 1)} % (L2): все значения положительны, и "
     f"y в среднем больше x на {num(rs_['mean'], 2)}.")

h2("3.5 Проверка гипотез о виде распределения")
tt = {t["Столбец"]: t for t in R["tests"]}
P_DAG_Y = tt["y"]["p (Д'Агостино)"]
para("По заданию 27 для первого столбца (x) проверена гипотеза о равномерном распределении, "
     "для второго (y) — о нормальном; для сравнения обе гипотезы проверены для обоих столбцов. "
     "Равномерность проверялась критерием Колмогорова–Смирнова (*scipy.stats.kstest*) для "
     "равномерного распределения на отрезке [min, max], нормальность — критериями "
     "Колмогорова–Смирнова, Шапиро–Уилка (*scipy.stats.shapiro*) и Д'Агостино "
     "(*scipy.stats.normaltest*); параметры распределений оценивались по выборке. Результаты "
     f"приведены в таблице {CNT['tab'] + 1}; при p > 0,05 оснований отвергнуть гипотезу нет.")
next_tab()
table("Проверка гипотез о виде распределения (p-значения)",
      ["Столбец", "Равномерное (КС)", "Нормальное (КС)", "Нормальное (Шапиро–Уилк)",
       "Нормальное (Д'Агостино)"],
      [[t["Столбец"], pv(t["p (КС, равномерное)"]), pv(t["p (КС, нормальное)"]), pv(t["p (Шапиро–Уилк)"]),
        pv(t["p (Д'Агостино)"])] for t in R["tests"]], widths=[2.0, 3.3, 3.3, 3.9, 3.9])
para("Гипотеза о равномерном распределении x отвергается (p < 0,001): как показано в "
     "подразделе 3.1, плотность x растёт пропорционально значению, и такое распределение "
     "называют треугольным. Гипотеза о нормальном распределении y отвергается критериями "
     f"Шапиро–Уилка (p < 0,001) и Д'Агостино (p = {tn(P_DAG_Y, 3)}): "
     "распределение асимметрично и скошено влево. Критерий Колмогорова–Смирнова с оценёнными по "
     f"выборке параметрами нормальность не отвергает (p = {tn(tt['y']['p (КС, нормальное)'], 3)}), "
     "но при оценке параметров по той же выборке он заметно теряет мощность, поэтому решение "
     "принимается по более мощным специализированным критериям. Вывод: x и y не являются ни "
     "равномерными, ни нормальными, что согласуется с видом гистограмм на рисунке 2. Для "
     f"отклонения y − √n гипотеза о нормальности также отвергается (p < 0,001), а равномерность "
     f"на границе значимости (p = {tn(R['resid_tests']['p_unif'], 3)}): шум ограничен интервалом от "
     "0 до 1 и по форме ближе к равномерному.")

h2("3.6 Спектральный анализ")
spc = R["spectral"]
n = next_fig()
para("Для каждого столбца построены спектрограмма (*scipy.signal.spectrogram*, окно 32 отсчёта "
     "с перекрытием 24, задание 28), периодограмма (*scipy.signal.periodogram*, задание 29) и "
     "амплитудно-частотная характеристика через быстрое преобразование Фурье (*np.fft.rfft*, "
     "задание 30). Частота выражена в долях частоты дискретизации (1/отсчёт), среднее перед "
     "анализом вычитается. Поскольку x и y почти целиком состоят из тренда, тот же анализ "
     f"выполнен для отклонения y − √n. Результаты показаны на рисунке {n}.", keep=True)
figure(fig_path("task28_30_spectral"), "Спектрограммы, периодограммы и амплитудно-частотные "
       "характеристики x, y и y − √n", 16.5)
para(f"Спектры x и y почти одинаковы: на самую низкую частоту 1/{N} ≈ {num(R['f1'], 4)} "
     f"(период, равный длине ряда) приходится {num(100 * spc['x']['share_f1'], 1)} % и "
     f"{num(100 * spc['y']['share_f1'], 1)} % мощности, на три низшие частоты — "
     f"{num(100 * spc['x']['share_f3'], 1)} % и {num(100 * spc['y']['share_f3'], 1)} %, а на "
     f"нижнюю половину частот — более {num(100 * min(spc['x']['low_half'], spc['y']['low_half']), 0)} %. "
     "Такой монотонно спадающий спектр характерен для трендового ряда без периодичности; "
     "спектрограммы показывают ту же картину во всех окнах — яркая полоса у нулевой частоты. "
     "Различие видно только на высоких частотах: у гладкого x спектр продолжает плавно убывать, "
     f"а у y выходит на постоянный уровень шума (средняя амплитуда верхней половины частот "
     f"{num(spc['y']['amp_mean_high'], 4)} против {num(spc['x']['amp_mean_high'], 4)} у x).")
para("Спектр отклонения y − √n не имеет выраженной структуры: периодограмма колеблется около "
     "среднего уровня во всём диапазоне частот, а на нижнюю половину частот приходится "
     f"{num(100 * spc['e']['low_half'], 1)} % мощности, как и должно быть для белого шума. "
     f"Наибольшее значение (на частоте {num(spc['e']['f_max'], 3)}) превышает средний уровень в "
     f"{num(spc['e']['max_to_mean'], 1)} раза; для {spc['e']['n_freq']} независимых частот шума "
     "такой одиночный выброс ожидаем и не означает периодичности, что подтверждает и "
     f"коэффициент автокорреляции {num(rs_['ac1'], 3)}. Спектрограмма отклонения однородна и не "
     "содержит устойчивых горизонтальных полос.")

h2("3.7 Интерполяция и бинарные маски")
ip = R["interp"]
gi = R["gap_interp"]
n = next_fig()
para("Кубическая интерполяция (*scipy.interpolate.interp1d*, kind = 'cubic', задание 31) и "
     "интерполяция сплайном третьей степени (*scipy.interpolate.UnivariateSpline*, k = 3, s = 0, "
     f"задание 32) выполнены для первых {ip['x']['n']} значений каждого столбца с шагом 0,1 "
     f"({ip['x']['n_new']} точек). Для y дополнительно построен сглаживающий сплайн с параметром "
     f"s = n·σ², где σ — среднеквадратическое отклонение шума. Результат показан на рисунке {n}.",
     keep=True)
figure(fig_path("task31_32_interpolation"), "Кубическая интерполяция и интерполяция сплайнами "
       "фрагментов x и y")
para("Кубическая интерполяция и интерполяционный сплайн проходят через все исходные точки и "
     f"совпадают между собой (расхождение {sci(max(ip['x']['diff'], ip['y']['diff']))}), так как "
     "оба метода строят один и тот же кубический сплайн. Для гладкого x промежуточные значения "
     f"отличаются от точных √n не более чем на {num(ip['x']['err_sqrt'], 4)}. Для зашумлённого y "
     "интерполяция повторяет шум и колеблется между узлами, а сглаживающий сплайн проходит "
     "между точками и отклоняется от тренда √n + 0,54 со среднеквадратическим отклонением "
     f"{num(ip['y']['sm_vs_trend'], 3)} при шуме {num(rs_['std'], 3)}. Кубическая интерполяция по "
     f"всем известным точкам x восстанавливает пропуски значениями {vec(gi['cubic'], 4)} с ошибкой "
     f"{sci(gi['err_cubic'])} относительно √{GN1} и √{GN2}; заполнение предыдущим значением, "
     f"выбранное в задании 8 из действий методических указаний, даёт ошибку {num(gi['err_ffill'], 3)}.")
mk, mz = R["masks"], R["masks_z"]
para("Бинарные маски (задание 33) — логические массивы, полученные сравнением столбца с "
     "условием: x > 0, x < 0, x = 0 и −1 ≤ x ≤ 1. Все значения x и y положительны и больше "
     "единицы (кроме x = 1 при n = 1), поэтому маски исходных столбцов почти тривиальны. Чтобы "
     "маски были содержательными, те же условия применены и к стандартизованным значениям "
     f"z = (x − μ)/σ. Число элементов, удовлетворяющих условиям, приведено в таблице "
     f"{CNT['tab'] + 1}.")
next_tab()
keys = ["x > 0", "x < 0", "x = 0", "-1 <= x <= 1"]
table("Число элементов, выделенных бинарными масками",
      ["Столбец", "> 0", "< 0", "= 0", "От −1 до 1"],
      [["x"] + [mk["x"][k] for k in keys], ["y"] + [mk["y"][k] for k in keys],
       ["z-оценка x"] + [mz["x"][k] for k in keys], ["z-оценка y"] + [mz["y"][k] for k in keys]],
      widths=[3.6, 2.7, 2.7, 2.7, 3.3])
n = next_fig()
para(f"Маски стандартизованных столбцов показаны на рисунке {n}.", keep=True)
figure(fig_path("task33_masks"), "Бинарные маски стандартизованных столбцов x и y")
zb = R["masks_z_bounds"]
para(f"Для z-оценки x маски образуют сплошные участки: значения ниже среднего (z < 0) занимают "
     f"первые {mz['x']['x < 0']} измерений, выше среднего — остальные {mz['x']['x > 0']}, так как "
     "ряд монотонен. В интервал от минус 1 до 1 (в пределах одного СКО от среднего) попали "
     f"{mz['x']['-1 <= x <= 1']} значений — это x от {num(zb['x'][0], 2)} до {num(zb['x'][1], 2)}. У y "
     f"переход через среднее происходит не в одной точке: из-за шума около n = {R['z_first_pos']['y']}–80 "
     "знак z несколько раз меняется, и маски «> 0» и «< 0» чередуются; на границах интервала "
     "[минус 1, 1] маска также имеет разрывы. Точных нулей нет ни в одном столбце: значения "
     "непрерывные.")
para(f"По разделу 3 можно сделать следующие выводы. Столбец x — детерминированный ряд √n без "
     f"шума с треугольным распределением, столбец y — тот же ряд со сдвигом {num(rs_['mean'], 2)} и "
     f"аддитивным некоррелированным шумом со среднеквадратическим отклонением "
     f"{num(rs_['std'], 2)}, ограниченным интервалом от 0 до 1. Связь x и y почти функциональная "
     f"(r = {num(cr['y']['x'], 4)}, p < 10⁻¹⁰⁰) и не имеет запаздывания; периодических "
     "составляющих в данных нет, спектр отклонения y − √n соответствует белому шуму, а "
     "сглаживание с окном 11 уменьшает шум примерно в 4 раза.")

# ================================================================ 4
h1("4 Масштабирование и скользящее окно")
h2("4.1 Сравнение с масштабированием scikit-learn")
scm = R["scaling_cmp"]
mx_diff = max(max(v["minmax"]["max_diff"], v["standard"]["max_diff"]) for v in scm.values())
mx_inv = max(v[k][e] for v in scm.values() for k in ("minmax", "standard")
             for e in ("inv_err_custom", "inv_err_sklearn"))
mx_invd = max(v[k]["inv_diff"] for v in scm.values() for k in ("minmax", "standard"))
n = next_fig()
para("В задании 34 каждый столбец масштабирован собственными классами и классами "
     "*MinMaxScaler* и *StandardScaler* из *scikit-learn*; совпадение проверено функцией "
     f"*np.allclose* с допуском 10⁻¹⁰. Результаты наложены на рисунке {n} и записаны функцией "
     "*write_table* в файлы task34_35_scaling_* во всех четырёх форматах (исходные, "
     "масштабированные и восстановленные значения обеих реализаций).", keep=True)
figure(fig_path("task34_scaling"), "Сравнение собственного масштабирования и scikit-learn", 16.5)
z0 = R["scale"]["zscore"]
para("Кривые собственной реализации и *scikit-learn* совпадают полностью: максимальная разность "
     f"не превышает {sci(mx_diff)} для MinMax и равна нулю для стандартизации, так как "
     "используется та же формула со среднеквадратическим отклонением с делителем n. "
     "Масштабирование не меняет форму рядов, а только переносит их в новый диапазон: после "
     "стандартизации x лежит от "
     f"{num(z0['min'][0], 2)} до {num(z0['max'][0], 2)}, y — от {num(z0['min'][1], 2)} до "
     f"{num(z0['max'][1], 2)}; несимметричность диапазона отражает левую асимметрию распределений. "
     "Инвертирование масштабирования (задание 35) собственными функциями и методом "
     f"*inverse_transform* восстанавливает исходные данные с ошибкой не более {sci(mx_inv)}, "
     f"результаты двух реализаций различаются не более чем на {sci(mx_invd)}.")
h2("4.2 Скользящее окно и скользящее среднее")
wn = R["windows"]
nl, nt = wn["noise_left"], wn["noise_theory"]
w1, w2, w3 = wn["widths"]
para("Функция *sliding_window* (задание 36, листинг А.2) принимает массив и ширину окна и "
     "формирует матрицу окон индексацией: строка i содержит элементы с i-го по (i + w − 1)-й. "
     f"Для y при ширине 10 получена матрица {wn['shape'][0]} × {wn['shape'][1]} "
     f"({N} − 10 + 1 окон), совпадающая с результатом "
     "*numpy.lib.stride_tricks.sliding_window_view*. Функция *moving_average* (задание 37) "
     "вычисляет скользящее среднее как среднее по строкам матрицы окон; результат совпал с "
     "методом *pandas.Series.rolling*.")
n = next_fig()
para(f"Скользящие средние y с окнами {w1}, {w2} и {w3} показаны на рисунке {n}; каждое значение "
     "отнесено к центру своего окна.", keep=True)
figure(fig_path("task37_moving_average"), f"Скользящие средние ряда y с окнами {w1}, {w2} и {w3}")
para("Периода в ряду нет, поэтому окна выбраны по уровню шума: усреднение w независимых "
     "значений уменьшает среднеквадратическое отклонение шума в √w раз. Отклонение от тренда "
     f"√n + {num(rs_['mean'], 2)} до сглаживания равно {num(wn['noise_before'], 3)}; после "
     f"сглаживания окном {w1} — {num(nl[str(w1)], 3)} (оценка {num(nt[str(w1)], 3)}), окном {w2} — "
     f"{num(nl[str(w2)], 3)} (оценка {num(nt[str(w2)], 3)}), окном {w3} — {num(nl[str(w3)], 3)} "
     f"(оценка {num(nt[str(w3)], 3)}). Фактическое уменьшение несколько больше оценки: соседние "
     "значения скользящего среднего сильно коррелированы, поэтому оценка СКО по одному ряду "
     f"из {N} значений имеет заметный случайный разброс. Широкое окно теряет по {(w3 - 1) // 2} "
     "точек на каждом конце ряда, а в начале ряда, где √n изгибается сильнее всего, сглаживает "
     f"и сам тренд (кривая окна {w3} в начале проходит ниже данных). Окно "
     f"{w2} уменьшает шум примерно вчетверо и сохраняет форму тренда, поэтому оно использовано "
     "и для свёртки в задании 24.")

# ================================================================ 5
h1("5 Тензорные вычисления")
tsr = R["tensors"]
para(f"Массив X ({N} × 2) преобразован в тензор *TensorFlow* (*tf.constant*, задание 38) и тензор "
     "*PyTorch* (*torch.tensor*, задание 39) типа float32; значения обоих тензоров совпадают. Для "
     "операции A·X·W + B (задание 40) созданы случайные тензоры: целочисленный A размером "
     f"5 × {N} из равномерного распределения на {{0, …, 9}}, W размером 2 × 4 из стандартного "
     "нормального распределения и B размером 5 × 4 из непрерывного равномерного распределения на "
     f"[0, 1). Размеры согласованы: (5 × {N})·({N} × 2)·(2 × 4) + (5 × 4) даёт матрицу 5 × 4.")
para("Операция выполнена в *TensorFlow* через *tf.matmul* и *tf.cast* (задание 41) и в *PyTorch* "
     f"через оператор @ (задание 42). Первые строки результатов приведены в таблице {CNT['tab'] + 1}.")
next_tab()
table("Первая строка результата A·X·W + B", ["Библиотека", "Столбец 1", "Столбец 2", "Столбец 3",
                                             "Столбец 4"],
      [["TensorFlow"] + [tn(v, 1) for v in tsr["axwb_tf"][0]],
       ["PyTorch"] + [tn(v, 1) for v in tsr["axwb_pt"][0]]], widths=[3, 3.3, 3.3, 3.3, 3.3], size=11)
AXWB_MAX = max(abs(v) for row in tsr["axwb_tf"] + tsr["axwb_pt"] for v in row)
para("Результаты различаются, потому что в каждой библиотеке генерируются свои случайные "
     "A, W и B. Для проверки тензоры A, W и B из *TensorFlow* были переданы в *PyTorch*: "
     "максимальная относительная разность результатов составила "
     f"{sci(tsr['axwb_same_inputs_rel_diff'])}, что соответствует точности типа float32. Большие "
     f"по модулю значения (до {num(AXWB_MAX, 0)}) объясняются суммированием {N} произведений с "
     "целыми коэффициентами до 9 при значениях x и y до 14.")
tr = tsr["TPQ_range"]
para("Для заданий 43–45 средствами *NumPy* созданы матрицы T, P и Q размером 3 × 10, "
     "отличающиеся распределением элементов: T — равномерное на [минус π, π] (значения от "
     f"{num(tr['T'][0], 2)} до {num(tr['T'][1], 2)}), P — нормальное со средним 0 и СКО 0,5 (от "
     f"{num(tr['P'][0], 2)} до {num(tr['P'][1], 2)}), Q — равномерное на [0, 4] (от "
     f"{num(tr['Q'][0], 3)} до {num(tr['Q'][1], 2)}; Q неотрицательна, что нужно для квадратного "
     "корня). Вычислено выражение")
formula("*V* = |sin *T* − *e*^{P} · √*Q*|,")
where(["операции выполняются поэлементно."])
para("Матрицы преобразованы в тензоры *keras.ops.convert_to_tensor* и *torch.tensor*. В *Keras* "
     "использованы функции *keras.ops* (*ops.sin*, *ops.exp*, *ops.sqrt*, *ops.abs*, задание 44), "
     "в *PyTorch* — аналогичные функции *torch* (задание 45). Значения V лежат в диапазоне от "
     f"{num(tsr['V_range'][0], 3)} до {num(tsr['V_range'][1], 3)}; максимальная разность "
     f"результатов *Keras* и *PyTorch* равна {sci(tsr['V_diff'])}, то есть результаты совпадают с "
     "точностью вычислений в двойной точности.")

# ================================================================ 6
h1("6 Матричные операции и пайплайн")
spr, inv = R["sparse"], R["inverse"]
para("В задании 46 функцией *scipy.sparse.random* создана случайная разреженная матрица "
     "100 000 × 100 000 с плотностью 10⁻⁵ в формате CSR: "
     f"{spr['nnz']} ненулевых элементов (не более {spr['max_row_nnz']} в одной строке), которые "
     f"вместе с индексами занимают {num(spr['mem_mib'], 2)} МиБ. Такая же плотная матрица типа "
     f"float64 заняла бы {num(spr['dense_gib'], 1)} ГиБ, что показывает необходимость разреженных "
     "форматов для больших матриц с малым числом ненулевых элементов.")
para("В задании 47 для случайной матрицы M размером 100 × 100 с элементами из [0, 1) вычислены "
     f"ранг ({inv['rank']}), определитель ({sci(inv['det'], 3)}) и число обусловленности "
     f"({num(inv['cond'], 1)}). Ранг полный, поэтому обратная матрица существует и вычислена "
     "функцией *np.linalg.inv*. Максимальное отклонение произведений M·M⁻¹ и M⁻¹·M от единичной "
     f"матрицы равно {sci(inv['max_err'])} и {sci(inv['max_err_left'])}.")
Fi = R["F"]
para(f"В задании 48 исходный массив X достроен до трёхмерного массива F размером "
     f"{' × '.join(map(str, Fi['shape']))}: по третьей оси расположены 8 слоёв — исходные данные, "
     "MinMax-нормализация в [0, 1] и в [минус 1, 1], стандартизация и четыре бинарные маски "
     "z-оценок из задания 33 (z > 0, z < 0, z = 0, −1 ≤ z ≤ 1) в виде нулей и единиц. Массив "
     "создаётся один раз, а наборы признаков извлекаются срезами по оси слоёв — представлениями "
     "(views), которые ссылаются на память F без копирования: для всех пяти наборов функция "
     "*np.shares_memory* вернула True, атрибут base указывает на F; развёрнутая матрица "
     "F.reshape(n, −1) также является представлением.")
pca = R["pca"]
best = pca["best"]
para("В задании 49 для каждого набора признаков (исходные, масштабированные, маски, исходные "
     "вместе с масштабированными и все слои) и каждого варианта масштабирования (без "
     "масштабирования для сравнения, *MinMaxScaler*, *StandardScaler*) построен пайплайн "
     "*Pipeline*(масштабирование → *PCA*). Постоянные признаки (маска z = 0) исключались. "
     "Наилучшей считалась проекция на две главные компоненты, сохраняющая наибольшую долю "
     "дисперсии. Для каждой комбинации вычислен средний логарифм правдоподобия модели *PCA* "
     f"(*PCA.score*) в зависимости от числа компонент. Результаты приведены в таблице "
     f"{CNT['tab'] + 1} и сохранены в файл task49_pca_combinations_*.xlsx.")
next_tab()
table("Результаты пайплайнов масштабирование → PCA",
      ["Набор признаков", "Масштабирование", "Признаков", "Ранг", "Доля дисперсии ГК 1",
       "Доля дисперсии ГК 1–2"],
      [[r["set"], r["scaler"], r["n_features"], r["rank_centered"], tn(r["evr1"], 4), tn(r["evr2"], 4)]
       for r in pca["rows"]], widths=[4.4, 3.6, 2.0, 1.4, 2.6, 2.6], size=10)
n = next_fig()
para(f"График логарифма правдоподобия и проекция наилучшей комбинации приведены на рисунке {n}.",
     keep=True)
figure(fig_path("task49_pca"), "Логарифм правдоподобия PCA и проекция наилучшей комбинации", 16.5)
rr = {(r["set"], r["scaler"]): r for r in pca["rows"]}
para("Исходные и масштабированные слои — аффинные преобразования одних и тех же двух "
     "столбцов, поэтому любой набор только из них имеет ранг 2 после центрирования: две "
     "компоненты сохраняют 100 % дисперсии, а первая — "
     f"{num(100 * rr[('исходные', 'StandardScaler')]['evr1'], 1)} %, так как x и y почти "
     "коллинеарны. Такие наборы проецируются на плоскость без потерь тривиально, поэтому в "
     "выборе наилучшей комбинации участвовали наборы с масками, ранг которых равен "
     f"{rr[('маски', 'MinMaxScaler')]['rank_centered']} и {rr[('все слои', 'MinMaxScaler')]['rank_centered']}. "
     f"Наилучшая проекция получена для набора «{best['set']}» с *{best['scaler']}*: две компоненты "
     f"сохраняют {num(100 * best['evr2'], 2)} % дисперсии против "
     f"{num(100 * rr[('все слои', 'StandardScaler')]['evr2'], 2)} % и "
     f"{num(100 * rr[('все слои', 'MinMaxScaler')]['evr2'], 2)} % для всех слоёв со "
     "стандартизацией и MinMax. Маски принимают значения 0 и 1, поэтому *MinMaxScaler* их не "
     "меняет (кривые без масштабирования и с MinMax на графике совпадают), а *StandardScaler* "
     "выравнивает дисперсии и немного снижает долю первых компонент. Без масштабирования у "
     f"всех слоёв две компоненты сохраняют {num(100 * rr[('все слои', 'без масштабирования')]['evr2'], 2)} %, "
     "но лишь потому, что дисперсия исходных значений (около 9) на порядок больше дисперсии "
     "остальных слоёв и подавляет их. На проекции маски образуют несколько плотных групп точек — "
     "каждая соответствует одной комбинации значений масок, а цвет показывает, что группы "
     "последовательно сменяют друг друга с ростом n. Логарифм правдоподобия растёт с числом "
     "компонент у всех комбинаций; значения для разных вариантов масштабирования несопоставимы "
     "напрямую, так как правдоподобие зависит от единиц измерения преобразованных данных.")

# ================================================================ 7
h1("7 Обработка изображения")
para("Для задания 50 программно (функция *generate_image*) построено цветное изображение "
     f"размером {IMG['shape'][1]} × {IMG['shape'][0]} пикселей (размер кратен 8): воздушные шары "
     "над полем — голубое небо с градиентом, белые облака, пять полосатых шаров разных цветов с "
     "бликами и корзинами, холмы, поле с полосами посевов, жёлтое поле подсолнечника, дорога и "
     "деревья; файл data/image.png. Сюжет содержит области с сильно различающимися цветом и "
     "яркостью. Все операции выполнены функциями модуля image_proc.py (листинг А.3).")
cm = IMG["channel_means"]
n = next_fig()
para("Разделение на цветовые каналы — выделение трёх двумерных массивов img[..., 0], "
     f"img[..., 1], img[..., 2] (рисунок {n}); чем темнее цвет на изображении канала, тем больше "
     "значение.", keep=True)
figure(fig_path("task50_1_channels"), "Исходное изображение и его красный, зелёный и синий каналы")
para(f"Средние значения каналов: красный {num(cm[0], 0)}, зелёный {num(cm[1], 0)}, синий "
     f"{num(cm[2], 0)}. Зелёный канал преобладает за счёт поля и холмов. Небо и облака дают "
     "большие значения синего канала, поле подсолнечника — красного и зелёного, но не синего. "
     "Полосы шаров по-разному видны в каналах: красно-жёлтый шар в красном канале почти "
     "однороден (обе краски содержат много красного), а в зелёном и синем видны полосы; "
     "фиолетовые полосы выделяются в красном и синем каналах.")
rs = IMG["rotated_shape"]
n = next_fig()
para("Поворот выполнен функцией *scipy.ndimage.rotate* на 40° против часовой стрелки с "
     "кубической интерполяцией (order = 3) и расширением холста (reshape = True), чтобы углы не "
     f"обрезались (рисунок {n}).", keep=True)
figure(fig_path("task50_2_rotated"), "Исходное изображение и изображение, повёрнутое на 40°")
para(f"Размер изображения после поворота — {rs[1]} × {rs[0]} пикселей: повёрнутый прямоугольник "
     "занимает больше места, освободившиеся области заполнены чёрным цветом (cval = 0). Линия "
     "горизонта поднимается слева направо под углом 40°, шары и дорога наклонены на тот же угол.")
n = next_fig()
para("Преобразование в чёрно-белое изображение (оттенки серого) выполнено по формуле яркости "
     f"Y = 0,299R + 0,587G + 0,114B (рекомендация ITU-R BT.601); результат на рисунке {n}.",
     keep=True)
figure(fig_path("task50_3_gray"), "Исходное и чёрно-белое изображения")
para(f"Средняя яркость {num(IMG['gray_mean'], 1)} из 255 — изображение светлое. Облака и нижняя "
     "часть неба остались светлыми, поле подсолнечника — светло-серым, а полосы посевов, "
     "деревья и тёмные полосы шаров стали тёмно-серыми, так как зелёный и красный цвета "
     "средней насыщенности имеют небольшую яркость.")
hs = IMG["hist_stats"]
n = next_fig()
para("Гистограмма чёрно-белого изображения построена функцией *np.histogram* по 256 интервалам "
     "шириной 1: левая граница каждого столбца совпадает с уровнем яркости, поэтому подписи оси "
     f"соответствуют столбцам без смещения (рисунок {n}).", keep=True)
figure(fig_path("task50_4_gray_hist"), "Гистограмма чёрно-белого изображения")
para(f"Яркость пикселей изменяется от {hs['min']} до {hs['max']}; темнее 60 лишь "
     f"{num(100 * hs['lt60'], 1)} % пикселей, ярче 200 — {num(100 * hs['gt200'], 1)} %. Гистограмма "
     "многовершинная, каждая вершина соответствует группе объектов: небольшая вершина около 52 — "
     "кроны деревьев; широкая область 75–140 — полосы поля, верхняя часть неба и шары, с "
     "острой вершиной около 124 — дальние холмы и светлые полосы поля; плато 160–200 — нижняя "
     "часть неба, поле подсолнечника и тени облаков; область 210–245 — края облаков; самая "
     f"высокая вершина около {IMG['hist_peak']} — белые облака.")
spl = IMG["split"]
sh = spl["shares"]
n = next_fig()
para("Разбиение на три изображения выполнено по порогам, найденным по накопленной "
     "гистограмме так, чтобы в каждую часть попала примерно треть пикселей: "
     f"t1 = {spl['t1']}, t2 = {spl['t2']}; пиксели вне диапазона части заменены чёрными "
     f"(рисунок {n}).", keep=True)
figure(fig_path("task50_5_split3"), "Изображения, полученные разбиением по порогам гистограммы")
para(f"Доли пикселей в частях: тёмные {num(100 * sh[0], 1)} %, средние {num(100 * sh[1], 1)} %, "
     f"светлые {num(100 * sh[2], 1)} %. Порог t1 = {spl['t1']} проходит по вершине гистограммы "
     "около 124, поэтому полосы поля, деревья, верх неба и тёмные полосы шаров попали в тёмную "
     "часть, а холмы оказались на границе и разделились зернисто между тёмной и средней частями. "
     "В среднюю часть попали нижняя часть неба, поле подсолнечника, тени облаков и дорога, в "
     "светлую — облака и светлые полосы шаров. Зернистость на границах частей вызвана шумом, "
     "добавленным при построении изображения: пиксели с яркостью около порога случайно "
     "попадают в соседнюю часть.")
H, W = IMG["shape"][0], IMG["shape"][1]
n = next_fig()
para("Чёрная круговая рамка построена алгоритмически: по координатной сетке *np.ogrid* "
     "вычисляется маска пикселей, расстояние от которых до центра больше радиуса "
     f"R = min(H, W)/2 − 10 = {min(H, W) // 2 - 10} пикселя, и эти пиксели обнуляются "
     f"(рисунок {n}). За пределы круга попали углы изображения — левый и правый шары и большая "
     "часть поля подсолнечника.", keep=True)
figure(fig_path("task50_6_frame"), "Изображение в чёрной круговой рамке")
n = next_fig()
para("Случайный гауссовский шум с нулевым средним и σ = 25 добавлен функцией "
     f"*np.random.normal*, результат ограничен диапазоном 0…255 (рисунок {n}).", keep=True)
figure(fig_path("task50_7_noisy"), "Исходное изображение и изображение с гауссовским шумом")
para(f"Средний квадрат отклонения от исходного изображения составил {num(IMG['noise_mse'], 1)}, "
     "что меньше σ² = 625: у ярких пикселей облаков, близких к 255, положительные значения шума "
     "обрезаются. Шум хорошо заметен на однородных участках неба и поля, а контуры шаров "
     "остаются различимыми.")
shp = IMG["sharpness"]
n = next_fig()
para("Размытие выполнено функцией *scipy.ndimage.gaussian_filter* с σ = 2 по осям "
     f"изображения и σ = 0 по оси каналов, чтобы цвета не смешивались (рисунок {n}). Ниже "
     "полных изображений показан увеличенный фрагмент с большим шаром (выделен жёлтой рамкой).",
     keep=True)
figure(fig_path("task50_8_blurred"), "Исходное и размытое фильтром Гаусса изображения", 15.0)
para("Границы полос шара, стропы и корзина стали размытыми, ступенчатый край шара сгладился. "
     "Количественно резкость оценивалась средним модулем разности яркостей соседних пикселей: "
     f"после размытия она уменьшилась с {num(shp['orig'], 2)} до {num(shp['blurred'], 2)}, то есть в "
     f"{num(shp['orig'] / shp['blurred'], 1)} раза.")
n = next_fig()
para("Повышение резкости (sharpening) выполнено методом нерезкого маскирования: к изображению "
     "добавляется разность исходного и размытого изображений с коэффициентом k = 1,5")
formula("*I*′ = *I* + *k* · (*I* − *G*_{σ} ∗ *I*),")
where(["*I* — исходное изображение;", "*G*_{σ} ∗ *I* — изображение, размытое фильтром Гаусса с σ = 2;",
       "*k* — коэффициент усиления."])
para(f"Результат показан на рисунке {n}.", keep=True)
figure(fig_path("task50_9_sharpened"), "Исходное изображение и изображение после повышения резкости",
       15.0)
para("Полосы шаров стали контрастнее, вокруг шаров и деревьев появились светлые ореолы, "
     "характерные для этого метода; одновременно усилился мелкий шум на однородных участках "
     f"неба. Средний модуль разности соседних пикселей вырос с {num(shp['orig'], 2)} до "
     f"{num(shp['sharpened'], 2)}, то есть в {num(shp['sharpened'] / shp['orig'], 1)} раза.")
pi = IMG["pca"]
para("Метод главных компонент применён к изображению по блокам: каждый блок 8 × 8 пикселей "
     f"одного канала рассматривается как объект с 64 признаками, всего {pi['n_objects']} "
     "объектов. Такое представление рекомендовано методическими указаниями для сжатия "
     "пространственных фрагментов: в отличие от трёх цветовых признаков пикселя, 64 признака "
     "блока позволяют исследовать зависимость качества от числа компонент в широком диапазоне.")
n = next_fig()
para(f"Изображение, восстановленное по 8 из 64 главных компонент, показано на рисунке {n}.", keep=True)
figure(fig_path("task50_10_pca_restored"), "Изображение, восстановленное по 8 главным компонентам")
para(f"Восемь компонент объясняют {num(100 * pi['evr_k'], 1)} % дисперсии блоков; средний квадрат "
     f"ошибки восстановления {num(pi['mse'], 1)}, пиковое отношение сигнал/шум {num(pi['psnr'], 1)} дБ. "
     "Сцена полностью узнаваема: шары, облака, поле и деревья на месте, но узкие полосы на "
     "маленьких шарах и стропы размыты, а на резких границах заметны блоки 8 × 8. Хранение 8 "
     "коэффициентов вместо 64 сжимает данные в 8 раз.")
n = next_fig()
para(f"Зависимость ошибки восстановления от числа компонент показана на рисунке {n}.", keep=True)
figure(fig_path("task50_10_pca_mse"), "Ошибка восстановления от числа главных компонент", 13.0)
mc = pi["mse_curve"]
para(f"Ошибка быстро убывает: {num(mc[0], 0)} при одной компоненте, {num(mc[3], 0)} при четырёх, "
     f"{num(mc[7], 0)} при восьми. Для сохранения 95 % дисперсии достаточно {pi['k95']} компонент, "
     f"99 % — {pi['k99']}. Большая часть изображения состоит из плавных переходов (небо, облака, "
     "поля), поэтому почти вся энергия блоков сосредоточена в первых компонентах. При 64 "
     "компонентах ошибка равна нулю с точностью вычислений, так как преобразование становится "
     "обратимым.")
n = next_fig()
para(f"Первые 16 главных компонент в виде блоков 8 × 8 показаны на рисунке {n}.", keep=True)
figure(fig_path("task50_10_pca_components"), "Первые 16 главных компонент блоков изображения", 11.0)
para(f"Первая компонента ({num(100 * pi['evr1'], 1)} % дисперсии) почти постоянна: все 64 элемента "
     f"одного знака и лежат в пределах от {num(pi['comp1_range'][0], 3)} до "
     f"{num(pi['comp1_range'][1], 3)}, то есть она описывает среднюю яркость блока. Вторая и третья "
     "компоненты — плавные перепады яркости слева направо и сверху вниз (вертикальные полосы "
     "шаров и поля, горизонт и градиент неба), четвёртая — светлая вертикальная полоса в центре "
     "блока. Следующие компоненты описывают всё более мелкие детали с большим числом смен "
     "знака, и вклад каждой начиная с 5-й не превышает 0,3 %.")

# ================================================================ ЗАКЛЮЧЕНИЕ
h1("Заключение", structural=True)
para("В ходе лабораторной работы выполнены все 50 заданий методических указаний на данных "
     "варианта 7 файла Lab1_Dataset.xlsx. Программа разбита на модули, все результаты "
     "сохраняются в файлы с датой и временем в имени, запуск воспроизводим благодаря "
     "фиксированному seed.")
para("Реализованы чтение и запись таблиц в форматах .xlsx, .csv, .txt и .mat с сохранением "
     "заголовков и автоматическим именованием столбцов при их отсутствии; все форматы "
     f"читаются обратно без потерь. Из листа {R['raw_shape'][0]} × {R['raw_shape'][1]} извлечены "
     f"{N} строк варианта 7 без изменения исходного файла. Проверка обнаружила несовпадение типов "
     f"(тип object) и {q['missing']['x']['count']} пустые ячейки x в строках {G1} и {G2}; на всём "
     f"листе нечисловые значения найдены в {len(bad_vars)} вариантах. Столбцы приведены к float64, "
     "пропуски x заполнены предыдущим значением с ошибкой не более "
     f"{num(hd[2]['err'], 3)} относительно закона √n вместо провала на {num(hd[1]['err'], 2)} при "
     "заполнении средним.")
para("Собственные функции MinMax-нормализации, нормализации в границах [a, b] и стандартизации "
     f"дают тот же результат, что и *scikit-learn* (разность не более {sci(mx_diff)}), и по "
     f"сохранённым в файл параметрам восстанавливают данные с ошибкой не более {sci(mx_inv)}. "
     "Функция разбиения даёт 50/50/50 элементов для соотношения 1 : 1 : 1 и 75/10/15 для "
     "75 % : 10 % : 15 %.")
para(f"Статистический анализ показал, что x — детерминированный ряд √n (все {N - 2} известных "
     "значений совпадают с законом), а y — тот же ряд со сдвигом "
     f"{num(rs_['mean'], 3)} и некоррелированным шумом с СКО {num(rs_['std'], 3)}. Средние x и y "
     f"равны {num(sx['mean'], 3)} и {num(sy['mean'], 3)}, дисперсии — {num(sx['var'], 3)} и "
     f"{num(sy['var'], 3)}, распределения скошены влево (медианы {num(sx['median'], 3)} и "
     f"{num(sy['median'], 3)} больше средних); гипотезы о равномерности x и нормальности y "
     "отвергнуты (p < 0,001). Коэффициент корреляции "
     f"r = {num(cr['y']['x'], 4)} значим (t = {num(sg['t'], 1)}), взаимная корреляция максимальна при "
     "нулевом сдвиге, спектры x и y сосредоточены на низших частотах "
     f"({num(100 * spc['x']['share_f3'], 1)} % мощности на трёх частотах), а спектр отклонения "
     "y − √n соответствует белому шуму. Градиент x совпадает с теоретической производной "
     f"1/(2√n) с точностью {num(g['x_err_inner'], 3)}, сглаживание окном {w2} уменьшает шум y с "
     f"{num(wn['noise_before'], 3)} до {num(nl[str(w2)], 3)}, кубическая интерполяция "
     f"восстанавливает пропуски x с ошибкой {sci(gi['err_cubic'])}.")
para("Операция A·X·W + B в *TensorFlow* и *PyTorch* при одинаковых входных тензорах даёт "
     f"одинаковый результат (относительная разность {sci(tsr['axwb_same_inputs_rel_diff'])}), "
     "выражение |sin T − e^{P}·√Q| в *Keras* и *PyTorch* совпадает с точностью "
     f"{sci(tsr['V_diff'])}. Разреженная матрица 100 000 × 100 000 занимает "
     f"{num(spr['mem_mib'], 2)} МиБ вместо {num(spr['dense_gib'], 1)} ГиБ, обратная матрица "
     f"100 × 100 найдена с погрешностью {sci(inv['max_err'])}. Массив F размером "
     f"{' × '.join(map(str, Fi['shape']))} построен с использованием представлений; наилучшая "
     f"проекция PCA получена для набора «{best['set']}» с *{best['scaler']}* "
     f"({num(100 * best['evr2'], 2)} % дисперсии на двух компонентах).")
para("Обработка изображения выполнена средствами *SciPy*: разделение каналов, поворот на 40°, "
     "перевод в оттенки серого, гистограмма и разбиение на три части по порогам "
     f"{spl['t1']} и {spl['t2']}, круговая рамка, шум, размытие (резкость уменьшилась в "
     f"{num(shp['orig'] / shp['blurred'], 1)} раза) и повышение резкости. PCA по блокам 8 × 8 "
     f"показал, что 8 компонент из 64 сохраняют {num(100 * pi['evr_k'], 1)} % дисперсии при "
     f"пиковом отношении сигнал/шум {num(pi['psnr'], 1)} дБ.")
para("Результатом работы является программа на Python, выполняющая чтение, проверку, очистку, "
     "статистический анализ, масштабирование и визуализацию табличных данных, тензорные "
     f"вычисления и обработку изображения; код и файлы результатов размещены {REPO_TEXT}.")

# ================================================================ ИСТОЧНИКИ
h1("Список использованных источников", structural=True)
DATE = "11.11.2026"
src = [
    "Библиотеки Python для машинного обучения. Работа с данными : методические указания по "
    "выполнению лабораторной работы № 1 по дисциплине «Машинное обучение» / ДГТУ, кафедра "
    "«ИПМ Ростсельмаш». – Ростов-на-Дону : ДГТУ, 2026.",
    "NumPy documentation [Электронный ресурс]. – URL: https://numpy.org/doc/stable/ "
    f"(дата обращения: {DATE}).",
    "pandas documentation [Электронный ресурс]. – URL: https://pandas.pydata.org/docs/ "
    f"(дата обращения: {DATE}).",
    "SciPy documentation [Электронный ресурс]. – URL: https://docs.scipy.org/doc/scipy/ "
    f"(дата обращения: {DATE}).",
    "scikit-learn: machine learning in Python. User guide [Электронный ресурс]. – URL: "
    f"https://scikit-learn.org/stable/user_guide.html (дата обращения: {DATE}).",
    "Matplotlib documentation [Электронный ресурс]. – URL: https://matplotlib.org/stable/ "
    f"(дата обращения: {DATE}).",
    "TensorFlow API documentation [Электронный ресурс]. – URL: "
    f"https://www.tensorflow.org/api_docs/python/tf (дата обращения: {DATE}).",
    "Keras 3 API documentation [Электронный ресурс]. – URL: https://keras.io/api/ "
    f"(дата обращения: {DATE}).",
    "PyTorch documentation [Электронный ресурс]. – URL: https://pytorch.org/docs/stable/ "
    f"(дата обращения: {DATE}).",
]
for i, s in enumerate(src, 1):
    para(f"{i}. {s}", align=WD_ALIGN_PARAGRAPH.LEFT)

# ================================================================ ПРИЛОЖЕНИЕ А
p = doc.add_paragraph(style="Heading 1")
p.paragraph_format.page_break_before = True
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.first_line_indent = Cm(0)
p.paragraph_format.space_after = Pt(0)
p.add_run("ПРИЛОЖЕНИЕ А")
t = para("Листинги основных функций программы", indent=False, align=WD_ALIGN_PARAGRAPH.CENTER, bold=True)
t.paragraph_format.space_after = Pt(21)
para(f"Полный исходный код размещён {REPO_TEXT}. Ниже приведены собственные функции "
     "масштабирования и разбиения (листинг А.1), скользящего окна (листинг А.2), основные "
     "операции обработки изображения (листинг А.3) и функции извлечения и проверки данных "
     "варианта (листинг А.4).", align=WD_ALIGN_PARAGRAPH.JUSTIFY)


def source_of(path, names=None):
    code = open(os.path.join(BASE, path), encoding="utf-8").read()
    if not names:
        return code.strip()
    tree = ast.parse(code)
    parts = [ast.get_source_segment(code, n) for n in tree.body
             if isinstance(n, (ast.FunctionDef, ast.ClassDef)) and n.name in names]
    return "\n\n\n".join(parts)


def listing(num_, title, code):
    c = doc.add_paragraph()
    c.paragraph_format.space_before = Pt(12)
    c.paragraph_format.keep_with_next = True
    c.add_run(f"Листинг А.{num_} – {title}")
    for line in code.splitlines():
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        r = p.add_run(line if line else " ")
        r.font.name, r.font.size = "Courier New", Pt(12)  # п. 9.4 Правил
        set_rfonts(r._r, "Courier New")


listing(1, "Собственные классы масштабирования и разбиение данных (src/preprocess.py)",
        source_of("src/preprocess.py", ["MinMaxScalerCustom", "StandardScalerCustom",
                                        "split_train_val_test"]))
listing(2, "Скользящее окно и скользящее среднее (src/windows.py)",
        source_of("src/windows.py", ["sliding_window", "moving_average"]))
listing(3, "Основные операции обработки изображения (src/image_proc.py)",
        source_of("src/image_proc.py", ["rotate_image", "to_grayscale", "gray_histogram",
                                        "split_by_histogram", "add_circular_frame",
                                        "sharpen_image"]))
listing(4, "Извлечение и проверка данных варианта (src/data_loader.py)",
        source_of("src/data_loader.py", ["check_data_quality", "variant_columns", "_variant_block",
                                         "extract_variant"]))

doc.save(DOCX)
print("Сохранено:", DOCX, "| рисунков:", CNT["fig"], "| таблиц:", CNT["tab"], "| формул:", CNT["eq"])
