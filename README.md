# Eplan.EplAddin.ERP

# EPLAN Add-In: Analyze BOM with Prices

This EPLAN 2025 add-in analyzes the entire Bill of Materials (BOM) in a project, compares it with supplier data from a CSV file, and helps you choose the most cost-effective and timely supplier for each part.

## Features

- Full BOM scan (articles assigned to functions)
- Supplier comparison based on:
  - Price (with discounts)
  - Delivery time
  - Availability (including stock)
- Budget and deadline constraints
- Auto-selection of optimal supplier
- Manual supplier override in GUI
- Excel export of final BOM with chosen supplier
- Advanced GUI showing:
  - All suppliers for each part
  - Delivery deadlines
  - Total cost calculations (min, max, actual)
  - Support for Catalog Number, EAN, ERP Code, Category
- Priority handling for conflicting constraints (deadline vs budget)

## Requirements

- EPLAN Electric P8 2025
- .NET Framework 4.8
- Visual Studio 2022
- ClosedXML v0.97.0
- CSV input: `parts_stock_full.csv`

## How it works

1. Run the add-in from EPLAN.
2. It automatically:
   - Loads the BOM
   - Loads the CSV data
   - Sets default deadline (1 year ahead) and budget (unlimited)
   - Calculates best supplier based on current conditions
3. You can:
   - Adjust budget and deadline
   - Choose priority (cost or delivery)
   - Recalculate and export final Excel sheet

---

# Dodatek EPLAN: Analiza BOM z cenami

Dodatek do EPLAN 2025 analizuje cały zestaw materiałowy (BOM) projektu, porównuje go z danymi hurtowni z pliku CSV i pomaga wybrać najkorzystniejszego dostawcę – cenowo i czasowo.

## Funkcje

- Pełna analiza BOM (artykuły przypisane do funkcji)
- Porównanie dostawców pod kątem:
  - Ceny (z rabatami)
  - Terminu dostawy
  - Dostępności (w tym magazyn)
- Ograniczenia budżetu i terminu
- Automatyczny wybór najlepszego dostawcy
- Możliwość ręcznej zmiany w GUI
- Eksport do Excela z wybranym dostawcą
- Zaawansowane GUI pokazujące:
  - Wszystkich dostawców dla każdej pozycji
  - Przedziały terminów dostaw
  - Całkowite koszty (minimalne, maksymalne, rzeczywiste)
  - Obsługa: Numer katalogowy, EAN, kod ERP, kategoria
- Obsługa konfliktów (budżet vs deadline)

## Wymagania

- EPLAN Electric P8 2025
- .NET Framework 4.8
- Visual Studio 2022
- ClosedXML v0.97.0
- Plik CSV wejściowy: `parts_stock_full.csv`

## Jak działa

1. Uruchom dodatek z poziomu EPLAN.
2. Program automatycznie:
   - Wczytuje BOM
   - Ładuje dane z CSV
   - Ustawia domyślnie: deadline = +1 rok, budżet = brak ograniczeń
   - Oblicza najlepszego dostawcę wg warunków
3. Możesz:
   - Zmienić budżet i termin
   - Wybrać priorytet (koszt vs termin)
   - Przeliczyć i wyeksportować do Excela