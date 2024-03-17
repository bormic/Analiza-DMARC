# Analiza-DMARC
Skrypt służy do analizy raportów DMARC zawartych w plikach archiwum, takich jak pliki ZIP i GZ. Użytkownik może wybrać te pliki, a następnie skrypt ekstrahuje dane z plików XML wewnątrz archiwów, przetwarza je i tworzy wykresy prezentujące różne aspekty analizy raportów DMARC.

Skrypt Pythona jest narzędziem analizy raportów DMARC zawartych w plikach archiwum (ZIP, GZ). Początkowo użytkownik jest zachęcany do wyboru plików archiwum za pomocą interfejsu użytkownika opartego na Tkinter. Następnie skrypt ekstrahuje dane XML z wybranych plików archiwum do folderu roboczego. Pliki archiwum są przenoszone do folderu zapasowego w celu zachowania oryginałów.

Po ekstrakcji danych z plików XML skrypt przetwarza je, tworząc ramkę danych pandas. W tym procesie dane są odpowiednio formatowane, a niektóre wartości są zamieniane na bardziej czytelne etykiety za pomocą mapowań. Na przykład kolumna "dkim" może zawierać wartości "pass", "fail" lub "none", które są zamieniane na "Poprawna weryfikacja DKIM", "Błąd weryfikacji DKIM" i "Brak autentykacji DKIM". Analogiczne mapowania są stosowane do innych kolumn, takich jak "adkim", "aspf", "comment" itp.

Po przetworzeniu danych skrypt generuje różne wykresy na podstawie przetworzonych danych, które prezentują różne aspekty raportów DMARC. Wykresy te obejmują wykresy słupkowe, wykresy kołowe i wykresy słupkowe skumulowane, które prezentują informacje o wynikach autentykacji DKIM, SPF, politykach DMARC, komentarzach, zakresie SPF itp.

Wygenerowane wykresy są zapisywane jako pliki JPEG w odpowiednim folderze, który jest tworzony, jeśli nie istnieje. Dodatkowo, istnieje możliwość wczytania istniejącego pliku Excela z danymi i generowania dodatkowych wykresów na podstawie tych danych, co daje użytkownikowi większą elastyczność i kontrolę nad analizą danych raportów DMARC.

Repozytorium zawiera:

1. wersję konsolową: dmarc_raport.py
2. wersję okienkową: dmarc_raport_v1.py
3. wersję skompilowaną do exe: AnalizaDMARC.py
4. skompilowany program exe: AnalizaDMARC.exe
