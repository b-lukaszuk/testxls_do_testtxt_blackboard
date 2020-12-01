# test_xls do test(y).txt w formacie blackboarda

Program pisany na kolanie na prosbe (ladna) E.Z,
nie optymalizowany,

program wczytuje plik_xls z pytaniami i tworzy pliki *.txt
(po 1 dla kazdego arkusza) do wczytania przez Blackboard

Zawartosc:
- test_format_uproszczony.xlsx - schemat wstawiania pytan testowych
- xls_do_txt_blackboard.py - plik programu w jezyku Python (ver. 3.8.5)


Wywolanie z bash-a:
```bash
python3.8 xls_do_blackboard.py plik_xlsx ile_arkuszy
```

Output:
- arkusz0.txt
- arkusz1.txt
- ...

Program (*.py) powinien dzialac poprawnie, ale nie udzielam zadnych gwarancji

Jelsli na podstawie informacji w tym pliku, przegladania pliku *.xls oraz
czytania kodu Python-a nie potrafisz go uruchomic, to nic nie ruszaj i
popros o pomoc autora :)

Po utworzeniu testu w blackboardzie, wczytujemy pliki z pytaniami
po kliknieciu "Import" ("Przekaż pytania")

## Do użytku własnego, nie powinno być używane przez nikogo innego.
## For personal use only, should not be used by anyone else.
