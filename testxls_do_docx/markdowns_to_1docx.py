#!/usr/bin/python3.8
# -*- coding: utf-8 -*-

# dzis: 2022-06-02
# uzyto Python 3.8.5
# pisano w Emacs >= 27
# na GNU/Linux Mint

import codecs
import subprocess
import os

together = ""

for file in os.listdir("./"):
    if file.startswith("arkusz") and file.endswith("md"):
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
