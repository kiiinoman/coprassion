#!/usr/bin/env python3
"""
compare_docx.py — Сравнивает два Word-документа с сохранением форматирования.

Результат:
  • <file1>_diff.docx — оригинал файла 1, удалённые (относительно файла 2)
                        фрагменты выделены КРАСНЫМ цветом
  • <file2>_diff.docx — оригинал файла 2, добавленные (которых не было в файле 1)
                        фрагменты выделены ЗЕЛЁНЫМ цветом

Использование:
    python compare_docx.py file1.docx file2.docx

Зависимости:
    pip install lxml
"""

import sys
import os
import re
import difflib
import zipfile
import shutil
from copy import deepcopy
from lxml import etree

# ─────────────────────────── XML-пространства имён ────────────────────────────
W   = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML = "http://www.w3.org/XML/1998/namespace"
WNS = f"{{{W}}}"
XSPACE = f"{{{XML}}}space"

RED   = "FF0000"
GREEN = "00AA00"


def wtag(name: str) -> str:
    return f"{WNS}{name}"


# ─────────────────────────── Работа с ZIP/XML ─────────────────────────────────

def read_xml(docx_path: str) -> bytes:
    with zipfile.ZipFile(docx_path) as z:
        return z.read("word/document.xml")


def write_xml(docx_src: str, docx_dst: str, new_xml: bytes):
    """Копирует docx_src в docx_dst, заменяя document.xml."""
    shutil.copy2(docx_src, docx_dst)
    tmp = docx_dst + ".tmp"
    with zipfile.ZipFile(docx_dst, "r") as zin, \
         zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename == "word/document.xml":
                zout.writestr(item, new_xml)
            else:
                zout.writestr(item, zin.read(item.filename))
    os.replace(tmp, docx_dst)


def serialize(root) -> bytes:
    return etree.tostring(root, xml_declaration=True,
                          encoding="UTF-8", standalone=True)


# ─────────────────────────── Утилиты ──────────────────────────────────────────

def para_text(para) -> str:
    return "".join(t.text or "" for t in para.iter(wtag("t")))


def clone_rPr(run) -> etree.Element:
    """Возвращает копию <w:rPr> рана, или новый пустой <w:rPr>."""
    rpr = run.find(wtag("rPr"))
    return deepcopy(rpr) if rpr is not None else etree.Element(wtag("rPr"))


def apply_color(rPr: etree.Element, hex_color: str):
    """Устанавливает или заменяет <w:color> в rPr."""
    existing = rPr.find(wtag("color"))
    if existing is not None:
        rPr.remove(existing)
    color_el = etree.Element(wtag("color"))
    color_el.set(wtag("val"), hex_color)
    rPr.insert(0, color_el)


def make_run(text: str, orig_run: etree.Element,
             color: str | None = None) -> etree.Element:
    """
    Создаёт <w:r> с текстом text и форматированием из orig_run.
    Если color задан — добавляет цвет.
    """
    r = etree.Element(wtag("r"))
    rPr = clone_rPr(orig_run)
    if color:
        apply_color(rPr, color)
    r.append(rPr)

    t = etree.SubElement(r, wtag("t"))
    t.text = text
    if text != text.strip() or text.startswith(" ") or text.endswith(" "):
        t.set(XSPACE, "preserve")
    return r


# ─────────────────────────── Атомы абзаца ─────────────────────────────────────
# Атом = (text_token, source_run)

def para_atoms(para) -> list[tuple[str, etree.Element]]:
    """
    Разбивает все <w:r> абзаца на токены (слово или пробел),
    сохраняя ссылку на исходный run для копирования форматирования.
    """
    atoms = []
    for r in para.iter(wtag("r")):
        for t in r.findall(wtag("t")):
            text = t.text or ""
            tokens = re.split(r"(\s+)", text)
            for tok in tokens:
                if tok:
                    atoms.append((tok, r))
    return atoms


# ─────────────────────────── Перестройка абзаца ───────────────────────────────

def rebuild_para(orig_para: etree.Element,
                 diff_atoms: list[tuple[str, etree.Element, bool]],
                 color: str) -> etree.Element:
    """
    Возвращает новый <w:p> на основе orig_para:
      - сохраняет <w:pPr> и все не-<w:r> элементы
      - заменяет <w:r> новыми согласно diff_atoms
        (highlighted=True → окрашиваем в color)
    """
    new_p = deepcopy(orig_para)

    # Убираем все <w:r> (и bookmarks/hyperlinks чтобы не дублировать)
    for child in list(new_p):
        if child.tag in (wtag("r"), wtag("hyperlink"),
                         wtag("bookmarkStart"), wtag("bookmarkEnd"),
                         wtag("proofErr")):
            new_p.remove(child)

    for text, orig_run, highlighted in diff_atoms:
        c = color if highlighted else None
        new_p.append(make_run(text, orig_run, c))

    return new_p


# ─────────────────────────── Основная логика ──────────────────────────────────

def process(docx1: str, docx2: str, out1: str, out2: str):
    xml1 = read_xml(docx1)
    xml2 = read_xml(docx2)

    root1 = etree.fromstring(xml1)
    root2 = etree.fromstring(xml2)

    body1 = root1.find(f".//{wtag('body')}")
    body2 = root2.find(f".//{wtag('body')}")

    paras1 = [p for p in body1 if p.tag == wtag("p")]
    paras2 = [p for p in body2 if p.tag == wtag("p")]

    texts1 = [para_text(p) for p in paras1]
    texts2 = [para_text(p) for p in paras2]

    # ── Diff абзацев ──────────────────────────────────────────────────────
    seq = difflib.SequenceMatcher(None, texts1, texts2, autojunk=False)

    # Индексы абзацев, которые нужно перекрасить
    # para_idx → list of (text, run, highlighted)
    paint1: dict[int, list] = {}   # file1: красим красным
    paint2: dict[int, list] = {}   # file2: красим зелёным

    for tag, i1, i2, j1, j2 in seq.get_opcodes():

        if tag == "equal":
            pass  # Ничего не делаем

        elif tag == "delete":
            # Абзацы только в file1 → все токены красные
            for idx in range(i1, i2):
                atoms = para_atoms(paras1[idx])
                paint1[idx] = [(t, r, True) for t, r in atoms]

        elif tag == "insert":
            # Абзацы только в file2 → все токены зелёные
            for idx in range(j1, j2):
                atoms = para_atoms(paras2[idx])
                paint2[idx] = [(t, r, True) for t, r in atoms]

        elif tag == "replace":
            # Для каждой пары абзацев делаем пословный diff
            block1 = paras1[i1:i2]
            block2 = paras2[j1:j2]

            # Собираем токены всего блока, разделяя абзацы маркером None
            def collect(paras_list):
                result = []
                for p in paras_list:
                    result.extend(para_atoms(p))
                    result.append(("\n", None))
                return result

            atoms1_block = collect(block1)
            atoms2_block = collect(block2)

            toks1 = [a[0] for a in atoms1_block]
            toks2 = [a[0] for a in atoms2_block]

            wseq = difflib.SequenceMatcher(None, toks1, toks2, autojunk=False)

            # Результирующие токены: (text, run, highlighted)
            res1: list[tuple[str, object, bool]] = []
            res2: list[tuple[str, object, bool]] = []

            for op, wi1, wi2, wj1, wj2 in wseq.get_opcodes():
                if op == "equal":
                    for k in range(wi1, wi2):
                        res1.append((toks1[k], atoms1_block[k][1], False))
                    for k in range(wj1, wj2):
                        res2.append((toks2[k], atoms2_block[k][1], False))
                elif op == "delete":
                    for k in range(wi1, wi2):
                        res1.append((toks1[k], atoms1_block[k][1], True))
                elif op == "insert":
                    for k in range(wj1, wj2):
                        res2.append((toks2[k], atoms2_block[k][1], True))
                elif op == "replace":
                    for k in range(wi1, wi2):
                        res1.append((toks1[k], atoms1_block[k][1], True))
                    for k in range(wj1, wj2):
                        res2.append((toks2[k], atoms2_block[k][1], True))

            # Разбиваем обратно по абзацам (по маркеру "\n")
            def split_chunks(res, orig_paras):
                chunks, cur = [], []
                for item in res:
                    if item[0] == "\n":
                        chunks.append(cur)
                        cur = []
                    else:
                        cur.append(item)
                if cur:
                    chunks.append(cur)
                while len(chunks) < len(orig_paras):
                    chunks.append([])
                return chunks[:len(orig_paras)]

            chunks1 = split_chunks(res1, block1)
            chunks2 = split_chunks(res2, block2)

            for local_idx, chunk in enumerate(chunks1):
                if chunk:
                    paint1[i1 + local_idx] = chunk
            for local_idx, chunk in enumerate(chunks2):
                if chunk:
                    paint2[j1 + local_idx] = chunk

    # ── Применяем изменения к абзацам ─────────────────────────────────────

    def apply_paint(body, paras, paint_map, color):
        """Заменяет нужные абзацы в body перекрашенными версиями."""
        for idx, new_para_data in paint_map.items():
            orig = paras[idx]
            new_p = rebuild_para(orig, new_para_data, color)
            # Заменяем в body
            parent = orig.getparent()
            pos = list(parent).index(orig)
            parent.remove(orig)
            parent.insert(pos, new_p)

    apply_paint(body1, paras1, paint1, RED)
    apply_paint(body2, paras2, paint2, GREEN)

    write_xml(docx1, out1, serialize(root1))
    write_xml(docx2, out2, serialize(root2))


# ─────────────────────────── Точка входа ──────────────────────────────────────

def main():
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)

    f1, f2 = sys.argv[1], sys.argv[2]

    for f in (f1, f2):
        if not os.path.exists(f):
            print(f"❌ Файл не найден: {f}")
            sys.exit(1)

    base1 = os.path.splitext(os.path.basename(f1))[0]
    base2 = os.path.splitext(os.path.basename(f2))[0]
    out1  = f"{base1}_diff.docx"
    out2  = f"{base2}_diff.docx"

    print(f"📄 Файл 1: {f1}")
    print(f"📄 Файл 2: {f2}")
    print("🔍 Сравниваю...\n")

    process(f1, f2, out1, out2)

    print(f"✅ {out1}  — удалённое (было в файле 1) выделено КРАСНЫМ")
    print(f"✅ {out2}  — добавленное (только в файле 2) выделено ЗЕЛЁНЫМ")
    print()
    print("Легенда:")
    print("  🔴 Красный  — текст был в файле 1, отсутствует в файле 2")
    print("  🟢 Зелёный  — текст появился в файле 2, отсутствовал в файле 1")
    print("  ⚪ Обычный  — текст одинаков в обоих файлах")


if __name__ == "__main__":
    main()
