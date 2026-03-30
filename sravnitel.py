#!/usr/bin/env python3
"""
compare_docx.py — Сравнивает два Word-документа, сохраняя оригинальное форматирование.

Результат:
  • file1_marked.docx — копия файла 1, где красным выделено всё, что было удалено/изменено
  • file2_marked.docx — копия файла 2, где зелёным выделено всё, что было добавлено/изменено

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

# ─────────────────────────── XML-пространства имён ────────────────────────────
W   = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML = "http://www.w3.org/XML/1998/namespace"
WNS = f"{{{W}}}"

def wtag(name: str) -> str:
    return f"{WNS}{name}"

# ─────────────────────────────── Утилиты ──────────────────────────────────────

def read_docx_xml(path: str) -> bytes:
    with zipfile.ZipFile(path) as z:
        return z.read("word/document.xml")

def write_docx_xml(src_docx: str, dst_docx: str, new_xml: bytes):
    """Копирует src_docx в dst_docx, заменяя word/document.xml."""
    shutil.copy2(src_docx, dst_docx)
    tmp = dst_docx + ".tmp"
    with zipfile.ZipFile(dst_docx, "r") as zin, \
         zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            if info.filename == "word/document.xml":
                zout.writestr(info, new_xml)
            else:
                zout.writestr(info, zin.read(info.filename))
    os.replace(tmp, dst_docx)

def para_text(p_elem) -> str:
    return "".join(t.text or "" for t in p_elem.iter(wtag("t")))

def get_body_paragraphs(root) -> list:
    body = root.find(f".//{wtag('body')}")
    return [c for c in body if c.tag == wtag("p")]

# ─────────────────── Работа с форматированием ────────────────────────────────

def set_color_on_rPr(rPr_elem, hex_color: str):
    existing = rPr_elem.find(wtag("color"))
    if existing is not None:
        rPr_elem.remove(existing)
    color_el = etree.Element(wtag("color"))
    color_el.set(wtag("val"), hex_color)
    rPr_elem.insert(0, color_el)

def make_colored_run(text: str, hex_color: str, source_rPr=None) -> etree.Element:
    r = etree.Element(wtag("r"))
    rPr = deepcopy(source_rPr) if source_rPr is not None else etree.Element(wtag("rPr"))
    set_color_on_rPr(rPr, hex_color)
    r.append(rPr)
    t = etree.SubElement(r, wtag("t"))
    t.text = text
    if text and (text[0] == " " or text[-1] == " "):
        t.set(f"{{{XML}}}space", "preserve")
    return r

def make_plain_run(text: str, source_rPr=None) -> etree.Element:
    r = etree.Element(wtag("r"))
    if source_rPr is not None:
        r.append(deepcopy(source_rPr))
    t = etree.SubElement(r, wtag("t"))
    t.text = text
    if text and (text[0] == " " or text[-1] == " "):
        t.set(f"{{{XML}}}space", "preserve")
    return r

def paragraph_char_tape(p_elem) -> list:
    """Список (char, rPr) для всех символов абзаца."""
    tape = []
    for r in p_elem.iter(wtag("r")):
        rPr = r.find(wtag("rPr"))
        rPr_copy = deepcopy(rPr) if rPr is not None else None
        for t in r.findall(wtag("t")):
            for ch in (t.text or ""):
                tape.append((ch, rPr_copy))
    return tape

def make_fully_colored_paragraph(p_elem, color: str) -> etree.Element:
    new_p = etree.Element(wtag("p"))
    pPr = p_elem.find(wtag("pPr"))
    if pPr is not None:
        new_p.append(deepcopy(pPr))
    for r in p_elem.iter(wtag("r")):
        new_r = deepcopy(r)
        rPr = new_r.find(wtag("rPr"))
        if rPr is None:
            rPr = etree.Element(wtag("rPr"))
            new_r.insert(0, rPr)
        set_color_on_rPr(rPr, color)
        new_p.append(new_r)
    return new_p

def copy_paragraph_unchanged(p_elem) -> etree.Element:
    return deepcopy(p_elem)

# ─────────────── Перестройка абзаца с точечной подсветкой ────────────────────

def rebuild_paragraph_with_highlights(p_elem, old_text: str, new_text: str,
                                       color: str, use_old: bool) -> etree.Element:
    """
    Строит новый <w:p>, сохраняя форматирование оригинальных run-ов.
    Изменённые слова перекрашивает в color.
    use_old=True  → базовый текст old_text (для файла 1, красим удалённое)
    use_old=False → базовый текст new_text (для файла 2, красим добавленное)
    """
    new_p = etree.Element(wtag("p"))
    pPr = p_elem.find(wtag("pPr"))
    if pPr is not None:
        new_p.append(deepcopy(pPr))

    tape = paragraph_char_tape(p_elem)

    base_words  = re.split(r"(\s+)", old_text)
    other_words = re.split(r"(\s+)", new_text)

    sm = difflib.SequenceMatcher(None, base_words, other_words, autojunk=False)

    # Строим список (segment_text, is_changed) для нужной стороны
    segments = []  # list of (str, bool)
    for op, i1, i2, j1, j2 in sm.get_opcodes():
        if use_old:
            seg = "".join(base_words[i1:i2])
            segments.append((seg, op != "equal"))
        else:
            seg = "".join(other_words[j1:j2])
            segments.append((seg, op != "equal"))

    # Сопоставляем символы сегментов с лентой форматирования
    tape_idx = 0
    for seg_text, is_changed in segments:
        if not seg_text:
            continue
        # rPr первого символа сегмента
        seg_rPr = tape[tape_idx][1] if tape_idx < len(tape) else None
        tape_idx += len(seg_text)

        if is_changed:
            new_p.append(make_colored_run(seg_text, color, seg_rPr))
        else:
            new_p.append(make_plain_run(seg_text, seg_rPr))

    return new_p

# ──────────────────────────── Основная функция ────────────────────────────────

def _replace_body_paragraphs(root, new_paras: list):
    body = root.find(f".//{wtag('body')}")
    sect_pr = body.find(wtag("sectPr"))
    for p in [c for c in list(body) if c.tag == wtag("p")]:
        body.remove(p)
    if sect_pr is not None:
        idx = list(body).index(sect_pr)
        for i, p in enumerate(new_paras):
            body.insert(idx + i, p)
    else:
        for p in new_paras:
            body.append(p)

def _output_name(path: str, suffix: str) -> str:
    base, ext = os.path.splitext(path)
    return f"{base}_{suffix}{ext}"

def process(docx1: str, docx2: str):
    out1 = _output_name(docx1, "marked")
    out2 = _output_name(docx2, "marked")

    root1 = etree.fromstring(read_docx_xml(docx1))
    root2 = etree.fromstring(read_docx_xml(docx2))

    paras1 = get_body_paragraphs(root1)
    paras2 = get_body_paragraphs(root2)
    texts1 = [para_text(p) for p in paras1]
    texts2 = [para_text(p) for p in paras2]

    sm = difflib.SequenceMatcher(None, texts1, texts2, autojunk=False)

    new_paras1: list = []
    new_paras2: list = []

    RED   = "FF0000"
    GREEN = "00B050"

    for op, i1, i2, j1, j2 in sm.get_opcodes():
        if op == "equal":
            for idx in range(i2 - i1):
                new_paras1.append(copy_paragraph_unchanged(paras1[i1 + idx]))
                new_paras2.append(copy_paragraph_unchanged(paras2[j1 + idx]))

        elif op == "replace":
            old_block = paras1[i1:i2]
            new_block = paras2[j1:j2]
            count = max(len(old_block), len(new_block))
            for idx in range(count):
                has_old = idx < len(old_block)
                has_new = idx < len(new_block)
                if has_old and has_new:
                    p1, p2 = old_block[idx], new_block[idx]
                    t1, t2 = para_text(p1), para_text(p2)
                    new_paras1.append(rebuild_paragraph_with_highlights(p1, t1, t2, RED,   use_old=True))
                    new_paras2.append(rebuild_paragraph_with_highlights(p2, t1, t2, GREEN, use_old=False))
                elif has_old:
                    new_paras1.append(make_fully_colored_paragraph(old_block[idx], RED))
                else:
                    new_paras2.append(make_fully_colored_paragraph(new_block[idx], GREEN))

        elif op == "delete":
            for p in paras1[i1:i2]:
                new_paras1.append(make_fully_colored_paragraph(p, RED))

        elif op == "insert":
            for p in paras2[j1:j2]:
                new_paras2.append(make_fully_colored_paragraph(p, GREEN))

    _replace_body_paragraphs(root1, new_paras1)
    _replace_body_paragraphs(root2, new_paras2)

    write_docx_xml(docx1, out1,
                   etree.tostring(root1, xml_declaration=True, encoding="UTF-8", standalone=True))
    write_docx_xml(docx2, out2,
                   etree.tostring(root2, xml_declaration=True, encoding="UTF-8", standalone=True))

    print(f"✅ {out1}  ← красным: что удалили/изменили")
    print(f"✅ {out2}  ← зелёным: что добавили/изменили")

# ────────────────────────────── Точка входа ───────────────────────────────────

def main():
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)
    f1, f2 = sys.argv[1], sys.argv[2]
    for f in (f1, f2):
        if not os.path.exists(f):
            print(f"❌ Файл не найден: {f}")
            sys.exit(1)
    print(f"📄 Файл 1: {f1}")
    print(f"📄 Файл 2: {f2}")
    print("🔍 Сравниваю...\n")
    process(f1, f2)
    print()
    print("Легенда:")
    print("  🔴 Красный  (файл 1) — текст, которого больше нет в файле 2")
    print("  🟢 Зелёный  (файл 2) — текст, которого не было в файле 1")

if __name__ == "__main__":
    main()
