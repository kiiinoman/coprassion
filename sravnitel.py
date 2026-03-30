"""
compare_docx.py — Сравнивает два Word-документа и выделяет различия красным маркером.

Использование:
    python compare_docx.py file1.docx file2.docx [output.docx]

Результат:
    Новый документ, в котором:
      - Текст из file1, которого нет в file2, выделен красным (удалён)
      - Текст из file2, которого нет в file1, выделен красным жирным (добавлен)
      - Одинаковый текст — без изменений
"""

import sys
import difflib
import zipfile
import shutil
import os
import re
from copy import deepcopy
from lxml import etree

# ───────────────────────────── Константы XML ──────────────────────────────────
W  = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WNS = f"{{{W}}}"

def wtag(name: str) -> str:
    return f"{WNS}{name}"


# ───────────────────────────── Утилиты ────────────────────────────────────────

def extract_paragraphs(docx_path: str) -> list[str]:
    """Извлекает текст всех абзацев документа."""
    with zipfile.ZipFile(docx_path) as z:
        xml = z.read("word/document.xml")
    root = etree.fromstring(xml)
    paragraphs = []
    for p in root.iter(wtag("p")):
        text = "".join(t.text or "" for t in p.iter(wtag("t")))
        paragraphs.append(text)
    return paragraphs


def parse_document(docx_path: str):
    """Возвращает (root, namespace_map) разобранного document.xml."""
    with zipfile.ZipFile(docx_path) as z:
        xml = z.read("word/document.xml")
    return etree.fromstring(xml)


def make_red_run(text: str, bold: bool = False) -> etree.Element:
    """Создаёт <w:r> с красным цветом и, опционально, жирным шрифтом."""
    r = etree.Element(wtag("r"))
    rPr = etree.SubElement(r, wtag("rPr"))

    color = etree.SubElement(rPr, wtag("color"))
    color.set(wtag("val"), "FF0000")

    highlight = etree.SubElement(rPr, wtag("highlight"))
    highlight.set(wtag("val"), "none")

    if bold:
        etree.SubElement(rPr, wtag("b"))
        etree.SubElement(rPr, wtag("bCs"))

    # Подчёркивание для наглядности
    u = etree.SubElement(rPr, wtag("u"))
    u.set(wtag("val"), "single")

    t = etree.SubElement(r, wtag("t"))
    t.text = text
    if text != text.strip() or text.startswith(" ") or text.endswith(" "):
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return r


def make_normal_run(text: str) -> etree.Element:
    """Создаёт обычный <w:r>."""
    r = etree.Element(wtag("r"))
    t = etree.SubElement(r, wtag("t"))
    t.text = text
    if text != text.strip() or text.startswith(" ") or text.endswith(" "):
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return r


def make_paragraph(runs: list[etree.Element], style_elem=None) -> etree.Element:
    """Создаёт <w:p> из списка <w:r>."""
    p = etree.Element(wtag("p"))
    if style_elem is not None:
        p.append(deepcopy(style_elem))
    for r in runs:
        p.append(r)
    return p


def get_pPr(p_elem) -> etree.Element | None:
    """Возвращает элемент <w:pPr> абзаца, если есть."""
    return p_elem.find(wtag("pPr"))


# ───────────────────────────── Основная логика ────────────────────────────────

def build_diff_document(docx1: str, docx2: str, output: str):
    """Строит результирующий документ с выделением различий."""

    lines1 = extract_paragraphs(docx1)
    lines2 = extract_paragraphs(docx2)

    # Получаем исходные <w:p> из первого документа (для сохранения стилей)
    root1 = parse_document(docx1)
    paras1 = list(root1.iter(wtag("p")))

    # Получаем исходные <w:p> из второго документа
    root2 = parse_document(docx2)
    paras2 = list(root2.iter(wtag("p")))

    # Словарь текст → <w:pPr> для переноса стилей абзацев
    def build_style_map(paras):
        m = {}
        for p in paras:
            txt = "".join(t.text or "" for t in p.iter(wtag("t")))
            pPr = p.find(wtag("pPr"))
            if txt not in m:
                m[txt] = pPr
        return m

    style_map1 = build_style_map(paras1)
    style_map2 = build_style_map(paras2)

    # Diff на уровне абзацев
    matcher = difflib.SequenceMatcher(None, lines1, lines2, autojunk=False)
    opcodes = matcher.get_opcodes()

    new_paras: list[etree.Element] = []

    for tag, i1, i2, j1, j2 in opcodes:
        if tag == "equal":
            # Одинаковые абзацы — берём из файла 2 без изменений
            for idx in range(j1, j2):
                txt = lines2[idx]
                pPr = style_map2.get(txt)
                p = make_paragraph([make_normal_run(txt)], pPr)
                new_paras.append(p)

        elif tag == "replace":
            # Изменённые абзацы — детальный diff по словам
            old_block = "\n".join(lines1[i1:i2])
            new_block = "\n".join(lines2[j1:j2])

            old_words = re.split(r"(\s+)", old_block)
            new_words = re.split(r"(\s+)", new_block)

            word_matcher = difflib.SequenceMatcher(None, old_words, new_words, autojunk=False)
            runs: list[etree.Element] = []

            for wtag2, wi1, wi2, wj1, wj2 in word_matcher.get_opcodes():
                if wtag2 == "equal":
                    text = "".join(old_words[wi1:wi2])
                    if text:
                        runs.append(make_normal_run(text))
                elif wtag2 == "delete":
                    text = "".join(old_words[wi1:wi2])
                    if text:
                        # Удалённый текст — красный зачёркнутый
                        r = make_red_run(text, bold=False)
                        strike = etree.SubElement(r.find(wtag("rPr")), wtag("strike"))
                        runs.append(r)
                elif wtag2 == "insert":
                    text = "".join(new_words[wj1:wj2])
                    if text:
                        # Добавленный текст — красный жирный
                        runs.append(make_red_run(text, bold=True))
                elif wtag2 == "replace":
                    old_text = "".join(old_words[wi1:wi2])
                    new_text = "".join(new_words[wj1:wj2])
                    if old_text:
                        r = make_red_run(old_text, bold=False)
                        strike = etree.SubElement(r.find(wtag("rPr")), wtag("strike"))
                        runs.append(r)
                    if new_text:
                        runs.append(make_red_run(new_text, bold=True))

            # Разбиваем результат обратно по абзацам
            pPr = style_map2.get(lines2[j1]) if j1 < len(lines2) else None
            p = make_paragraph(runs, pPr)
            new_paras.append(p)

        elif tag == "delete":
            # Абзацы только в file1 — красный зачёркнутый
            for idx in range(i1, i2):
                txt = lines1[idx]
                pPr = style_map1.get(txt)
                r = make_red_run(txt, bold=False)
                strike_el = etree.SubElement(r.find(wtag("rPr")), wtag("strike"))
                p = make_paragraph([r], pPr)
                new_paras.append(p)

        elif tag == "insert":
            # Абзацы только в file2 — красный жирный
            for idx in range(j1, j2):
                txt = lines2[idx]
                pPr = style_map2.get(txt)
                r = make_red_run(txt, bold=True)
                p = make_paragraph([r], pPr)
                new_paras.append(p)

    # ── Собираем итоговый документ на основе file2 ──────────────────────────
    shutil.copy2(docx2, output)

    with zipfile.ZipFile(output, "r") as z:
        xml_bytes = z.read("word/document.xml")

    root = etree.fromstring(xml_bytes)

    # Находим <w:body> и заменяем все <w:p> нашими
    body = root.find(f".//{wtag('body')}")
    if body is None:
        raise RuntimeError("Не найден элемент <w:body> в document.xml")

    # Удаляем все старые абзацы, сохраняем sectPr (настройки страницы)
    sect_pr = body.find(wtag("sectPr"))
    for child in list(body):
        body.remove(child)

    for p in new_paras:
        body.append(p)

    if sect_pr is not None:
        body.append(sect_pr)

    new_xml = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

    # Перезаписываем document.xml внутри ZIP
    tmp_zip = output + ".tmp"
    with zipfile.ZipFile(output, "r") as zin, zipfile.ZipFile(tmp_zip, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename == "word/document.xml":
                zout.writestr(item, new_xml)
            else:
                zout.writestr(item, zin.read(item.filename))

    os.replace(tmp_zip, output)
    print(f"✅ Готово! Результат сохранён в: {output}")


# ───────────────────────────── Точка входа ────────────────────────────────────

def main():
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)

    file1 = sys.argv[1]
    file2 = sys.argv[2]
    output = sys.argv[3] if len(sys.argv) > 3 else "diff_result.docx"

    for f in (file1, file2):
        if not os.path.exists(f):
            print(f"❌ Файл не найден: {f}")
            sys.exit(1)

    print(f"📄 Файл 1: {file1}")
    print(f"📄 Файл 2: {file2}")
    print(f"🔍 Сравниваю...")

    build_diff_document(file1, file2, output)

    print()
    print("Легенда:")
    print("  🔴 Зачёркнутый красный  — текст удалён (был в файле 1, нет в файле 2)")
    print("  🔴 Жирный красный       — текст добавлен (нет в файле 1, есть в файле 2)")
    print("  ⚪ Обычный              — текст одинаков в обоих файлах")


if __name__ == "__main__":
    main()
