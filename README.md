Program "Synonimy" to skrypt napisany w języku Python, który pozwala na znajdowanie synonimów i zmianę wyrazów w pliku .docx. Wykorzystuje on moduły takie jak BeautifulSoup, requests, docx, collections, prettytable oraz re.

Klasa "Synonyms" posiada trzy metody: "change_all", "change_my_choice" oraz "show_statistics". Metoda "change_all" zmienia wszystkie słowa w pliku, natomiast "change_my_choice" pozwala na wybór konkretnych słów, które mają zostać zmienione. Metoda "show_statistics" prezentuje statystyki dotyczące częstotliwości występowania poszczególnych słów w dokumencie .docx.

Program pobiera synonimy ze strony "synonimy.pl" i wykorzystuje je do zmiany wyrazów w dokumencie. Słowa, dla których nie zostaną znalezione synonimy, są oznaczane gwiazdką, a słowa, dla których istnieje więcej niż jeden synonim, są oznaczane myślnikiem. Wynik programu można zapisać w pliku .docx.

Program jest przyjazny dla użytkownika dzięki wykorzystaniu kolorów do wyróżnienia poszczególnych słów w statystykach oraz do wyróżnienia zmienionych słów w wynikowym pliku .docx.
