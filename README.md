# XlsxToJira

Konsolowe narzędzie .NET 9, które wycina zakres wierszy z arkusza `.xlsx` i wypisuje go jako tabelę gotową do wklejenia w Jira.

- domyślnie **Jira wiki markup** (`||nagłówek||`, `|wiersz|`) – stary edytor, Jira Server/DC
- z flagą `--md` **Markdown** – nowy edytor Jira Cloud rozpoznaje go przy wklejaniu

Pierwszy wiersz podanego zakresu jest traktowany jako nagłówek.

## Użycie

```
XlsxToJira [plik.xlsx] [od] [do] [--md] [--sheet <nazwa|nr>]
```

```bash
dotnet run -- "G:\dane\tabela.xlsx" 2 15
dotnet run -- "G:\dane\tabela.xlsx" 2 15 --md --sheet Arkusz2
```

### Tryb interaktywny

Jeśli program uruchomisz bez parametrów (np. dwuklikiem na `.exe`) albo z niepełnymi lub błędnymi (brak pliku, `do` < `od`), dopyta o brakujące rzeczy:

```
Ścieżka do pliku .xlsx: G:\dane\tabela.xlsx
Arkusze:
  1. Dane
  2. Drugi
Arkusz (nr) [1]:
Arkusz „Dane”: ostatni użyty wiersz = 15
Wiersz nagłówka (od) [1]: 2
Ostatni wiersz (do) [15]:
Format [Enter = Jira wiki markup, m = Markdown]:
```

Enter przyjmuje wartość domyślną w nawiasie. Pytania idą na stderr, więc `XlsxToJira.exe | clip` nadal wrzuca do schowka samą tabelę.

Czysty tekst prosto do schowka:

```bash
dotnet build -c Release
bin\Release\net9.0\XlsxToJira.exe "G:\dane\tabela.xlsx" 2 15 | clip
```

## Szczegóły

- `--sheet` przyjmuje nazwę lub numer arkusza (domyślnie pierwszy)
- plik może być w tym czasie otwarty w Excelu
- komórki są wypisywane tak, jak Excel je formatuje (daty, separator dziesiętny)
- `|` w treści jest escapowany, komórki wielolinijkowe sklejane (`<br>` w Markdown, `\\` w wiki markup)
- biblioteka: [ClosedXML](https://github.com/ClosedXML/ClosedXML)
