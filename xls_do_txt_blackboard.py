#!/usr/bin/python3.8
# -*- coding: utf-8 -*-

################################################################################
#              program pisany na kolanie na prosbe (ladna) E.Z
#                           nie optymalizowany
#       program wczytuje plik_xls z pytaniami i tworzy pliki *.txt
#        (po 1 dla kazdego arkusza) do wczytania przez Blackboard
#                          wywolanie z bash-a:
#          python3.8.5 xls_do_blackboard.py plik_xlsx ile_arkuszy
###############################################################################

# dzis: 2020-12-01
# uzyto Python 3.8.5
# pisano w Emacs >= 26
# na GNU/Linux Mint
import pandas as pd
import numpy as np
import sys
import codecs


if len(sys.argv) != 3:
    print("Podano nieprawidlowa liczbe argumentow wejsciowych")
    print("Nie wykonano programu. Prosze poprawic input")
else:

    # z dokumentacji wynika, ze ma byc tab-delimited text file
    # bez header row, bez blank linesow miedzy wierszami
    # 1 pole w wierszu okresla typ pytania, pola oddzielone TAB
    # przyklad (spacji spacje wstawiono dla lepszej czytelnosci):
    # MC TAB question_text TAB answ1_text TAB correct|incorrect TAB answ2_text TAB correct| incorrect

    quest_type = "MC"
    field_sep = "\t"
    cor_answ = "correct"
    wrong_answ = "incorrect"

    def tworz_pytanie(start, stop):
        """
        wyodrebnia pytanie z obiektu pd.DataFrame o nazwie cur_arkusz
        Input:
        ---
        start - Int (inclusive) - indeks gdzie zaczyna sie dane pytanie
        stop - Int (exclusive) - indeks gdzie konczy sie dane pytanie

        Output:
        ---
        String - pytanie sformatowane pod blackboarda do wczytania
        """
        pytanie = (
            quest_type
            + field_sep
            + cur_arkusz.loc[start, "treść pytania"]
            + field_sep
        )
        for i in range(start, stop):
            pytanie += cur_arkusz.loc[i, "odpowiedzi"] + field_sep
            if cur_arkusz.loc[i, "prawidłowa*"] == 1:
                pytanie += "correct" + field_sep
            else:
                pytanie += "incorrect" + field_sep
        return pytanie + "\n"

    nazwa_pliku = sys.argv[1]
    l_arkuszy = int(sys.argv[2])
    nazwy_kol = ["nr_pyt", "treść pytania", "odpowiedzi", "prawidłowa*"]

    for arkusz_id in range(l_arkuszy):
        # wczytuje puste pola jako NaN
        cur_arkusz = pd.read_excel(
            io=nazwa_pliku,
            sheet_name=arkusz_id,
            usecols=nazwy_kol,
            na_values="",
        )
        cur_arkusz = cur_arkusz.replace(np.nan, "", regex=True)
        cur_arkusz.head()  # do us

        # wybieramy niepuste pytania, tj. te wiersze gdzie
        # kolumna "odpowiedzi" nie jest pusta
        niepuste = []
        for i in range(cur_arkusz.shape[0]):
            if cur_arkusz.loc[i, "odpowiedzi"] != "":
                niepuste.append(i)
        cur_arkusz = cur_arkusz.iloc[niepuste, :]  # wybranie niepustych

        # wybieramy indeksy poczatkow pytan
        pocz_pytan = []
        for i in range(cur_arkusz.shape[0]):
            if cur_arkusz.loc[i, "nr_pyt"] != "":
                pocz_pytan.append(i)
        pocz_pytan.append(cur_arkusz.shape[0])  # dodaje id ost wiersz-a

        tekst_wynikowy = ""
        for i in range(len(pocz_pytan) - 1):
            tekst_wynikowy += tworz_pytanie(
                start=pocz_pytan[i], stop=pocz_pytan[i + 1]
            )

        # zapisywanie pliku
        # kodowanie polskich znakow, aby blackboard to odczytal
        f = codecs.open("arkusz" + str(arkusz_id) + ".txt", "w", "utf-16")
        f.write(tekst_wynikowy)
        f.close()

    print("utworzono plik(i) TXT do blackboard-a")
