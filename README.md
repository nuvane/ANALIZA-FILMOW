<h1>Analiza danych Netflixa</h1>

Projekt oparty na Pythonie i bibliotekach <b>pandas</b>, <b>openpyxl</b>, <b>matplotlib</b>, <b>seaborn</b> oraz <b>plotly</b>.<br>
Automatyczny generator raportu z danych Netflixa zawierający statystyki, wykresy oraz interaktywne wizualizacje.<br>
Skrypt tworzy raport Excel (<b>dane/analiza_netflix.xlsx</b>) z przetworzonymi danymi, wykresami i hiperlinkami do wersji interaktywnych.<br>

<hr>

<h2>Funkcje</h2>

- Analiza danych z pliku <code>netflix.csv</code>
- Czyszczenie danych i konwersja dat (bez godzin)
- Obliczanie statystyk:
  - Typy produkcji (film/serial)
  - Najczęstsze kraje
  - Gatunki (zliczane ze złożonych tagów)
  - Reżyserzy
  - Rozkład według lat
- Eksport do Excela z:
  - Arkuszem danych źródłowych
  - Osobnymi arkuszami z topowymi kategoriami i wykresami
  - Statycznymi wykresami (PNG)
  - Interaktywnymi wykresami (HTML) z linkami w Excelu
  - Arkuszem <b>Podsumowanie</b> z metadanymi i hiperlinkami

<hr>

<h3>Arkusz: Gatunki</h3>
<img src="images/netflix_gatunki.png" alt="Wykres gatunków" width="700"/>

<hr>

<h3>Wykres interaktywny: Produkcje wg lat</h3>
<img src="images/netflix_lata_interaktywny.png" alt="Wykres roczny" width="700"/>

<hr>

<h3>Arkusz: Podsumowanie</h3>
<img src="images/netflix_podsumowanie.png" alt="Podsumowanie Excel" width="700"/>
