#!/usr/bin/env python3
"""
compare_docx.py — Сравнивает два Word-документа, не меняя структуру документа.

Результат:
  • file1_marked.docx — копия файла 1, красным выделено удалённое/изменённое
  • file2_marked.docx — копия файла 2, зелёным выделено добавленное/изменённое

Использование:
    python compare_docx.py file1.docx file2.docx

Зависимости:
    pip install lxml
"""

import sys
import os
import re
import zipfile
import shutil
import difflib
from copy import deepcopy
from lxml import etree

W   = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML = "http://www.w3.org/XML/1998/namespace"
WNS = f"{{{W}}}"

def wtag(name):
    return f"{WNS}{name}"

# ──────────────────────────── ZIP helpers ─────────────────────────────────────

def read_zip(path):
    """Читает все файлы из ZIP, возвращает dict {filename: bytes}."""
    data = {}
    with zipfile.ZipFile(path) as z:
        for info in z.infolist():
            data[info.filename] = z.read(info.filename)
    return data

def write_zip(src_docx, dst_docx, replacements):
    """
    Копирует src_docx в dst_docx, заменяя файлы из словаря replacements
    {filename: new_bytes}.
    """
    shutil.copy2(src_docx, dst_docx)
    tmp = dst_docx + ".tmp"
    with zipfile.ZipFile(dst_docx, "r") as zin, \
         zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            if info.filename in replacements:
                zout.writestr(info, replacements[info.filename])
            else:
                zout.writestr(info, zin.read(info.filename))
    os.replace(tmp, dst_docx)

# ──────────────────────────── XML helpers ─────────────────────────────────────

def para_text(p_elem):
    return "".join(t.text or "" for t in p_elem.iter(wtag("t")))

def collect_paragraphs(root):
    """
    Возвращает ВСЕ <w:p> в документе включая те, что внутри таблиц.
    Порядок — document order.
    """
    return list(root.iter(wtag("p")))

def set_color_on_rPr(rPr_elem, hex_color):
    old = rPr_elem.find(wtag("color"))
    if old is not None:
        rPr_elem.remove(old)
    el = etree.Element(wtag("color"))
    el.set(wtag("val"), hex_color)
    rPr_elem.insert(0, el)

def colorize_run(r_elem, hex_color):
    """Красит один run не трогая ничего кроме <w:color>."""
    rPr = r_elem.find(wtag("rPr"))
    if rPr is None:
        rPr = etree.Element(wtag("rPr"))
        r_elem.insert(0, rPr)
    set_color_on_rPr(rPr, hex_color)

def colorize_paragraph_full(p_elem, hex_color):
    """Красит все run-ы абзаца целиком."""
    for r in p_elem.findall(f".//{wtag('r')}"):
        colorize_run(r, hex_color)

def colorize_paragraph_partial(p_elem, old_text, new_text, hex_color, use_old):
    """
    Красит только run-ы, попадающие в изменённые диапазоны (word-level diff).
    Структура XML не меняется — только добавляется/заменяется <w:color>.
    """
    base_words  = re.split(r"(\s+)", old_text)
    other_words = re.split(r"(\s+)", new_text)

    sm = difflib.SequenceMatcher(None, base_words, other_words, autojunk=False)

    # Строим список изменённых символьных диапазонов в базовом тексте
    changed_ranges = []
    pos = 0
    for op, i1, i2, j1, j2 in sm.get_opcodes():
        if use_old:
            seg = "".join(base_words[i1:i2])
        else:
            seg = "".join(other_words[j1:j2])
        seg_len = len(seg)
        if op != "equal" and seg_len > 0:
            changed_ranges.append((pos, pos + seg_len))
        pos += seg_len

    if not changed_ranges:
        return

    # Проходим по run-ам, отслеживая позицию в тексте
    char_pos = 0
    for r in p_elem.findall(f".//{wtag('r')}"):
        run_text = "".join(t.text or "" for t in r.findall(wtag("t")))
        run_len  = len(run_text)
        if run_len == 0:
            continue
        run_start = char_pos
        run_end   = char_pos + run_len
        # Красим если run хоть частично пересекается с изменённым диапазоном
        if any(s < run_end and e > run_start for s, e in changed_ranges):
            colorize_run(r, hex_color)
        char_pos += run_len

# ──────────────────────────── Основная логика ─────────────────────────────────

def _output_name(path, suffix):
    base, ext = os.path.splitext(path)
    return f"{base}_{suffix}{ext}"

def process(docx1, docx2):
    out1 = _output_name(docx1, "marked")
    out2 = _output_name(docx2, "marked")

    zip1 = read_zip(docx1)
    zip2 = read_zip(docx2)

    # Парсим document.xml каждого файла
    root1 = etree.fromstring(zip1["word/document.xml"])
    root2 = etree.fromstring(zip2["word/document.xml"])

    # Собираем все абзацы (включая внутри таблиц) — НЕ вырезаем их из дерева
    paras1 = collect_paragraphs(root1)
    paras2 = collect_paragraphs(root2)
    texts1 = [para_text(p) for p in paras1]
    texts2 = [para_text(p) for p in paras2]

    RED   = "FF0000"
    GREEN = "00B050"

    sm = difflib.SequenceMatcher(None, texts1, texts2, autojunk=False)

    for op, i1, i2, j1, j2 in sm.get_opcodes():
        if op == "equal":
            continue  # ничего не делаем — оригинальные run-ы остаются как есть

        elif op == "replace":
            old_block = paras1[i1:i2]
            new_block = paras2[j1:j2]
            count = max(len(old_block), len(new_block))
            for idx in range(count):
                has_old = idx < len(old_block)
                has_new = idx < len(new_block)
                if has_old and has_new:
                    t1 = texts1[i1 + idx]
                    t2 = texts2[j1 + idx]
                    colorize_paragraph_partial(old_block[idx], t1, t2, RED,   use_old=True)
                    colorize_paragraph_partial(new_block[idx], t1, t2, GREEN, use_old=False)
                elif has_old:
                    colorize_paragraph_full(old_block[idx], RED)
                else:
                    colorize_paragraph_full(new_block[idx], GREEN)

        elif op == "delete":
            for p in paras1[i1:i2]:
                colorize_paragraph_full(p, RED)

        elif op == "insert":
            for p in paras2[j1:j2]:
                colorize_paragraph_full(p, GREEN)

    # Сериализуем изменённые деревья обратно в байты
    new_xml1 = etree.tostring(root1, xml_declaration=True, encoding="UTF-8", standalone=True)
    new_xml2 = etree.tostring(root2, xml_declaration=True, encoding="UTF-8", standalone=True)

    write_zip(docx1, out1, {"word/document.xml": new_xml1})
    write_zip(docx2, out2, {"word/document.xml": new_xml2})

    print(f"OK {out1}  <- красным: что удалили/изменили")
    print(f"OK {out2}  <- зелёным: что добавили/изменили")

# ──────────────────────────── Точка входа ─────────────────────────────────────

def main():
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)
    f1, f2 = sys.argv[1], sys.argv[2]
    for f in (f1, f2):
        if not os.path.exists(f):
            print(f"Файл не найден: {f}")
            sys.exit(1)
    print(f"Файл 1: {f1}")
    print(f"Файл 2: {f2}")
    print("Сравниваю...\n")
    process(f1, f2)
    print()
    print("Легенда:")
    print("  Красный  (файл 1) -- текст, которого больше нет в файле 2")
    print("  Зелёный  (файл 2) -- текст, которого не было в файле 1")

if __name__ == "__main__":
    main()
