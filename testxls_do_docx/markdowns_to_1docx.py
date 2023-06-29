#!/usr/bin/python3.8
# -*- coding: utf-8 -*-

# dzis: 2022-06-02
# uzyto Python 3.8.5
# pisano w Emacs >= 27
# na GNU/Linux Mint

import codecs
import os
import re
import subprocess

together = ""


def getIntFromStr(file_name):
    return int(re.search(r'[0-9]+', file_name)[0])


md_files = [f for f in os.listdir(
    "./") if (f.startswith("arkusz") and f.endswith("md"))]
md_files.sort(key=getIntFromStr)

for file in md_files:
    curFile = codecs.open(file, "r", "utf-8")
    together += (
        file.replace(".md", "").upper() + "\n\n" + curFile.read() + "\n\n"
    )
    curFile.close()

# zapisywanie pliku glownego
f = codecs.open("all_together.md", "w", "utf-8")
f.write(together)
f.close()

# zapisywanie/eksport do *.docx przez pandoc-a
subprocess.run(
    [
        "pandoc",
        "-o",
        "all_together.docx",
        "-f",
        "markdown",
        "-t",
        "docx",
        "all_together.md",
        "--wrap=preserve",
        "--reference-doc=reference.docx",
    ]
)
print("writing " + "all_together.md" + " to docx")
