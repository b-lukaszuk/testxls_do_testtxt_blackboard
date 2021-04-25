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

if len(sys.argv) != 3:
    print("Podano nieprawidlowa liczbe argumentow wejsciowych")
    print("Nie wykonano programu. Prosze poprawic input")
else:
    # z dokumentacji wynika, ze ma byc tab-delimited text file
    # bez header row, bez blank linesow miedzy wierszami
    # 1 pole w wierszu okresla typ pytania, pola oddzielone TAB
    # przyklad (spacje wstawiono dla lepszej czytelnosci):
    # MC TAB questText TAB answ1Text TAB correct|incorrect TAB answ2Text TAB correct|incorrect

    quest_type = "MC"
    field_sep = "\t"
    cor_answ = "correct"
    wrong_answ = "incorrect"

    def tworz_pytanie(start, stop):
        """
        wyodrebnia pytanie z obiektu pd.DataFrame o nazwie cur_arkusz
        i zwraca je jako string w formacie rozpoznawanym przez blackboard

        Input:
        ---
        start - Int (inclusive) - indeks (wiersz) gdzie zaczyna sie dane pytanie
        stop - Int (exclusive) - indeks (wiersz) gdzie konczy sie dane pytanie

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
            # str(), bo w odpowiedzi moga byc same cyferki (inty, floaty)
            # wstaw jesli pole z odpowiedzia nie jest puste
            if str(cur_arkusz.loc[i, "odpowiedzi"]).strip() != "":
                pytanie += (
                    str(cur_arkusz.loc[i, "odpowiedzi"]).strip() + field_sep
                )
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

        # wybieramy niepuste pytania, tj. te wiersze gdzie
        # kolumna "odpowiedzi" lub komorka "treść pytania" nie jest pusta
        niepuste = []
        for i in range(cur_arkusz.shape[0]):
            # str() bo w odp moga byc same cyferki (inty, floaty)
            # str.strip(), bo po edycji mogly zostac np. 3 spacje
            # niepuste jesli jest komorka z pytaniem lub z odpowiedzia
            if (
                str(cur_arkusz.loc[i, "odpowiedzi"]).strip() != ""
                or str(cur_arkusz.loc[i, "treść pytania"]).strip() != ""
            ):
                niepuste.append(i)
        cur_arkusz = cur_arkusz.iloc[niepuste, :]  # wybranie niepustych
        # reset indeksu, cyferki po kolei wym. przez pd.DataFrame.loc[] ponizej
        cur_arkusz = cur_arkusz.reset_index()

        # wybieramy indeksy poczatkow pytan
        pocz_pytan = []
        for i in range(cur_arkusz.shape[0]):
            # str.strip(), bo po edycji mogly zostac np. 3 spacje
            if str(cur_arkusz.loc[i, "treść pytania"]).strip() != "":
                pocz_pytan.append(i)
        pocz_pytan.append(cur_arkusz.shape[0])  # dodaje id ost wiersz-a

        tekst_wynikowy = ""
        # -1, bo pod tym indeksem (row) jest koncowka ost pytania (ost odp)
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
