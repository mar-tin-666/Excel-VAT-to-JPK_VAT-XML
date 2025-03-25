# Excel to JPK VAT (XML) Converter

Skrypt w Pythonie do przetwarzania plików Excel z rejestrem zakupów VAT i generowania osobnych plików JPK VAT XML dla każdej zakładki (arkusza).

## Wymagania

- Python 3.13 (na wcześniejszych nie testowane)
- Pakiety: `pandas`, `openpyxl`

Instalacja zależności:

```bash
pip install -r requirements.txt
```

## Użycie

1. Uruchom skrypt:

```bash
python main.py
```

2. Wprowadź/wklej ścieżkę do pliku Excel (zawierającego dane VAT):
```
D:\ścieżka\do\pliku.xlsx
```

3. Dla każdej zakładki, która zawiera dane `Type = Purchase`, zostanie wygenerowany osobny plik `.xml`.

Przykład dla pliku Excel `Plik_z_danymi.xlsx`:
```
Plik_z_danymi__Zakladka1.xml
Plik_z_danymi__Faktury_marzec.xml
```

## Format pliku wejściowego

Plik Excel musi zawierać następujące kolumny (lub zbliżone nazwy):

- **Type** – typ transakcji (np. `Purchase`)
- **VAT Registration No** – numer VAT dostawcy
- **External Document No** – numer zewnętrzny dokumentu
- **Document Date** – data wystawienia
- **Document Receipt/Sales Date** – data wpływu
- **VAT Base** – podstawa opodatkowania
- **VAT Amount** – kwota VAT

> Kolumny mogą zawierać dodatkowe przecinki, spacje itp. – skrypt je rozpozna automatycznie (a przynajmniej powinien).

## Bezpieczeństwo danych

Dane z Excela nie są nigdzie przesyłane – całość przetwarzana jest lokalnie.
