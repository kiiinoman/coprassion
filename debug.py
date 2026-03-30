#!/usr/bin/env python3
"""
debug_compare.py — показывает что именно видит скрипт при сравнении двух docx.

Использование:
    python debug_compare.py file1.docx file2.docx
"""

import sys
import os
import re
import zipfile
import difflib
from lxml import etree

W   = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WNS = f"{{{W}}}"

def wtag(name):
    return f"{WNS}{name}"

def read_docx_xml(path):
    with zipfile.ZipFile(path) as z:
        return z.read("word/document.xml")

def para_text(p_elem):
    return "".join(t.text or "" for t in p_elem.iter(wtag("t")))

def collect_paragraphs(root):
    return list(root.iter(wtag("p")))

def repr_text(s):
    """Показывает невидимые символы явно."""
    return repr(s)

def main():
    if len(sys.argv) < 3:
        print("Использование: python debug_compare.py file1.docx file2.docx")
        sys.exit(1)

    f1, f2 = sys.argv[1], sys.argv[2]

    root1 = etree.fromstring(read_docx_xml(f1))
    root2 = etree.fromstring(read_docx_xml(f2))

    paras1 = collect_paragraphs(root1)
    paras2 = collect_paragraphs(root2)
    texts1 = [para_text(p) for p in paras1]
    texts2 = [para_text(p) for p in paras2]

    print(f"Абзацев в файле 1: {len(texts1)}")
    print(f"Абзацев в файле 2: {len(texts2)}")
    print()

    sm = difflib.SequenceMatcher(None, texts1, texts2, autojunk=False)

    equal_count   = 0
    replace_count = 0
    delete_count  = 0
    insert_count  = 0

    for op, i1, i2, j1, j2 in sm.get_opcodes():
        if op == "equal":
            equal_count += i2 - i1
        elif op == "replace":
            replace_count += max(i2 - i1, j2 - j1)
            # Показываем детали replace-блоков
            for idx in range(max(i2 - i1, j2 - j1)):
                has_old = idx < (i2 - i1)
                has_new = idx < (j2 - j1)
                if has_old and has_new:
                    t1 = texts1[i1 + idx]
                    t2 = texts2[j1 + idx]
                    if t1 == t2:
                        print(f"[REPLACE но ТЕКСТЫ ОДИНАКОВЫ!]")
                        print(f"  Файл1: {repr_text(t1)}")
                        print(f"  Файл2: {repr_text(t2)}")
                        print()
                    else:
                        print(f"[REPLACE]")
                        print(f"  Файл1: {repr_text(t1)}")
                        print(f"  Файл2: {repr_text(t2)}")
                        # Word-level diff
                        w1 = re.split(r"(\s+)", t1)
                        w2 = re.split(r"(\s+)", t2)
                        wsm = difflib.SequenceMatcher(None, w1, w2, autojunk=False)
                        changes = [(op2, "".join(w1[a1:a2]), "".join(w2[b1:b2]))
                                   for op2, a1, a2, b1, b2 in wsm.get_opcodes()
                                   if op2 != "equal"]
                        if changes:
                            print(f"  Изменения по словам:")
                            for cop, cold, cnew in changes:
                                print(f"    {cop}: '{cold}' -> '{cnew}'")
                        print()
        elif op == "delete":
            delete_count += i2 - i1
            for idx in range(i1, i2):
                print(f"[DELETE] {repr_text(texts1[idx])}")
            print()
        elif op == "insert":
            insert_count += j2 - j1
            for idx in range(j1, j2):
                print(f"[INSERT] {repr_text(texts2[idx])}")
            print()

    print("=" * 60)
    print(f"Итого абзацев: equal={equal_count}, replace={replace_count}, "
          f"delete={delete_count}, insert={insert_count}")

if __name__ == "__main__":
    main()
