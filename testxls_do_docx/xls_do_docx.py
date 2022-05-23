#!/usr/bin/python3.8
# -*- coding: utf-8 -*-

################################################################################
#              program pisany na kolanie na prosbe (ladna) E.Z
#                           nie optymalizowany
#       program wczytuje plik_xls z pytaniami i tworzy pliki *.md
#        (po 1 dla kazdego arkusza) a potem *.docx (przez pandoca)
#                          wywolanie z bash-a:
#          python3.8.5 xls_do_blackboard.py plik_xlsx ile_arkuszy
###############################################################################

# dzis: 2022-05-23
# uzyto Python 3.8.5
# pisano w Emacs >= 27
# na GNU/Linux Mint

if len(sys.argv) != 3:
    print("Podano nieprawidlowa liczbe argumentow wejsciowych")
    print("Nie wykonano programu. Prosze poprawic input")
else:
    quest_sep = "\n"
    answer_sep = "\n"

    def removeTabsNewLines(tekst):
        what = ["\n", "\t"]
        with_what = " "
        for symbol in what:
            tekst = tekst.replace(symbol, with_what)
            return tekst

    def getBoldText(text):
        return "**" + text + "**"  # markdown bold

    def getFormattedQuestion(text):
        return getBoldText(text) + quest_sep

    def getFormattedAnswer(letter, text, is_correct):
        result = getBoldText(letter + ")") + " " + text
        # underlining, that would be exporterd md -> docx by pandoc
        if is_correct:
            result = '<span class="underline">' + result + "</span>"
        return result

    def getQuestionWithAnswers(start, stop):
        """
        wyodrebnia pytanie z obiektu DataFrame o nazwie cur_arkusz
        i zwraca je jako string w formacie markdown

        Input:
        ---
        start - Int (inclusive) - indeks (wiersz) gdzie zaczyna sie pytanie
        stop - Int (exclusive) - indeks (wiersz) gdzie konczy sie dane pytanie

        Output:
        ---
        String - pytanie sformatowane pod blackboarda do wczytania
        """

        # tu najpierw idzie pytanie, a potem doklejona beda odpowiedzi
        pytanie = removeTabsNewLines(cur_arkusz.loc[start, "treść pytania"])
        pytanie = getFormattedQuestion(pytanie)

        letterId = 0

        for i in range(start, stop):
            # str(), bo w odpowiedzi moga byc same cyferki (inty, floaty)
            # wstaw jesli pole z odpowiedzia nie jest puste
            if str(cur_arkusz.loc[i, "odpowiedzi"]).strip() != "":
                pytanie += (
                    "\n"
                    + getFormattedAnswer(
                        ascii_uppercase[letterId],
                        str(cur_arkusz.loc[i, "odpowiedzi"]).strip(),
                        cur_arkusz.loc[i, "prawidłowa*"] == 1,
                    )
                    + answer_sep
                )
                letterId += 1
        # "\n\\\n\\\n" oddziel. od siebie pytan md->docx (eksport przez pandoc)
        return pytanie + "\n\\\n\\\n"

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
        # reset indeksu, cyferki po kolei wym. przez DataFrame.loc[] ponizej
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
            tekst_wynikowy += getQuestionWithAnswers(
                start=pocz_pytan[i], stop=pocz_pytan[i + 1]
            )
        # usuniecie niepotrzebnych znakow ("\n") po ostatnim pytaniu
        tekst_wynikowy = tekst_wynikowy[:-4]

        # zapisywanie pliku
        # kodowanie polskich znakow, aby blackboard to odczytal
        f = codecs.open("arkusz" + str(arkusz_id) + ".md", "w", "utf-8")
        f.write(tekst_wynikowy)
        f.close()

        # zapisywanie/eksport do *.docx przez pandoc-a
        subprocess.run(
            [
                "pandoc",
                "-o",
                "arkusz" + str(arkusz_id) + ".docx",
                "-f",
                "markdown",
                "-t",
                "docx",
                "arkusz" + str(arkusz_id) + ".md",
                "--wrap=preserve",
            ]
        )

    print("utworzono plik(i) *.md i *.docx")
