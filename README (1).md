# Excel VBA Expert 2.0.0 - pełna lokalna baza Microsoft VBA

23 października 2026. Rozbudowa istniejącego skilla, nie nowy projekt.

## Dwa pakiety, jedna instalacja
`skill.zip` zawiera skilla, zachowane 19 modułów eksperckich, siedem workflow,
stare narzędzia i przykłady oraz nowe narzędzia do pełnej dokumentacji offline.
`excel-vba-knowledge.zip` zawiera **rzeczywiste pliki Microsoft VBA-Docs**: wszystkie
38 187 dokumentów Markdown z dostarczonego repozytorium, pozostałe pliki,
oryginalne licencje, indeks pełnotekstowy i indeksy przenośne JSONL.
Nie trzeba ponownie pobierać repozytorium ani budować indeksu.

Struktura po rozpakowaniu OBU archiwów do tego samego katalogu:
```text
C:\Users\skowal\.copilot\skills\
  excel-vba-expert\
    SKILL.md
    scripts\offline_docs.py
    metadata\knowledge-config.json
  excel-vba-expert-knowledge\
    manifest.json
    index.sqlite
    documents.jsonl
    repository\
      api\Excel.Range.Find.md
      Language\...
      LICENSE
      LICENSE-CODE
      ThirdPartyNotices
```
Folder wiedzy jest rodzeństwem skilla, nie drugim skillem. Nie zawiera SKILL.md.
Oba ZIP-y mają już własny katalog główny. Nie twórz dodatkowego poziomu.
Pliki zajmują około 177 MB danych logicznych; system plików może wykorzystać
więcej miejsca ze względu na tysiące małych plików. Zostaw co najmniej 400 MB.

## Instalacja / aktualizacja na Windows
Zachowaj pobrane ZIP-y w jednym katalogu. Poniższy skrypt przenosi poprzednie
katalogi do kopii poza katalogiem wykrywania Skills; nie kasuje poprzedniej wersji.
Uruchom PowerShell w katalogu z oboma ZIP-ami:

```powershell
$ErrorActionPreference = "Stop"
foreach ($zip in @("skill.zip", "excel-vba-knowledge.zip")) {
    if (-not (Test-Path -LiteralPath $zip -PathType Leaf)) {
        throw "Missing archive: $zip"
    }
}
$base = Join-Path $HOME ".copilot\skills"
$backup = Join-Path $HOME (".copilot\skill-backups\vba-" + (Get-Date -Format "yyyyMMdd-HHmmss"))
New-Item -ItemType Directory -Force -Path $base | Out-Null
foreach ($name in @("excel-vba-expert", "excel-vba-expert-knowledge")) {
    $existing = Join-Path $base $name
    if (Test-Path -LiteralPath $existing) {
        New-Item -ItemType Directory -Force -Path $backup | Out-Null
        Move-Item -LiteralPath $existing -Destination (Join-Path $backup $name)
    }
}
Expand-Archive -LiteralPath .\skill.zip -DestinationPath $base
Expand-Archive -LiteralPath .\excel-vba-knowledge.zip -DestinationPath $base
$skill = Join-Path $base "excel-vba-expert"
py -3 "$skill\scripts\offline_docs.py" verify
if ($LASTEXITCODE -ne 0) { throw "Knowledge verification failed; keep the backup." }
```
Nie usuwaj kopii poprzedniej wersji przed sprawdzeniem nowej. Po instalacji otwórz
nową rozmowę agenta i wskaż `excel-vba-expert`. Wykrywanie skilli zależy od
klienta/rozszerzenia. Nie wykonano zdalnej instalacji ani testu Twojego VS Code/Qwen.

## Wymagania
Python 3.10+ z modułem sqlite3 i obsługą FTS5. Nie potrzeba pip, Git, serwera
wektorowego, modelu embeddingowego, kluczy API, Excela ani internetu do wyszukiwania
źródeł. Zwykłe pliki Markdown można czytać także bez Pythona.
Sprawdzenie SQLite/FTS5 bez internetu:
```powershell
py -3 -c "import sqlite3; c=sqlite3.connect(':memory:'); c.execute('CREATE VIRTUAL TABLE t USING fts5(x)'); print('FTS5 OK')"
```

## Sprawdzenie bez internetu
Polecenia mogą być wykonywane z dowolnego katalogu. Baza jest odnajdywana względem
skilla, a nie bieżącego projektu:
```powershell
$skill = Join-Path $HOME ".copilot\skills\excel-vba-expert"
py -3 "$skill\scripts\offline_docs.py" search "Range.Find" --exact --json
py -3 "$skill\scripts\offline_docs.py" read "Range.Find" --expand-includes
py -3 "$skill\scripts\offline_docs.py" read "UserForm.Show"
py -3 "$skill\scripts\offline_docs.py" read "xlWhole"
py -3 "$skill\scripts\offline_docs.py" links "Range.Find"
```
`read` domyślnie zwraca **cały dokument**, nie streszczenie. Każdy odczyt sprawdza
sumę SHA-256 pliku. Dostępne są `--section Remarks`, `--start-line`, `--end-line`
i `--json`. Przy niejednoznacznej nazwie podaj dokładną ścieżkę z wyszukiwania.
Przykład: `read "api/Excel.Range(object).md"` odróżnia Range Excela od Graph.
`search "Document" --scope all` przeszukuje też pozostałe produkty Office.

## Polecenie kontrolne dla agenta
> Użyj skilla excel-vba-expert. Bez internetu wyszukaj lokalną dokumentację
> Range.Find. Odczytaj pełny dokument, wskaż lokalną ścieżkę i linie z sygnaturą,
> parametrami i uwagami. Następnie zaproponuj bezpieczny kod na podstawie źródła
> oraz modułu eksperckiego. Nie uruchamiaj Excela i nie zmieniaj plików projektu.

Analogicznie sprawdź Workbooks.Open, Worksheet.Change, UserForm.Show i
ChartObjects.Add. Brak sekcji w oryginale ma być jawnie zgłoszony, nie uzupełniany
z domysłu. [Dokładne wyniki tych pięciu prób](OFFLINE-VERIFICATION.md).

## Aktualizacja z nowego lokalnego repozytorium
Nie nadpisuj opracowanych reguł ani działającej bazy. Import do NOWEGO katalogu:
```powershell
py -3 "$skill\scripts\import_knowledge.py" --source ".\VBA-Docs-main.zip" --dest "D:\VBA-KB-next"
py -3 "$skill\scripts\offline_docs.py" --kb "D:\VBA-KB-next" verify
py -3 "$skill\scripts\offline_docs.py" --kb "D:\VBA-KB-next" read "Range.Find"
```
`--source` może wskazywać także katalog repozytorium. Import nie pobiera niczego,
nie uruchamia zawartości repozytorium i nie nadpisuje istniejącego celu. W katalogu
Git pomija tylko metadane `.git`; źródła i licencje zachowuje bez zmian.
Po weryfikacji zamień katalog rodzeństwa z zachowaniem kopii albo ustaw w procesie
agenta `EXCEL_VBA_KB` na nowy katalog. Opcja `--kb` ma pierwszeństwo.

## Zachowane narzędzia i testy
`search_docs.py` nadal przeszukuje 70 opracowanych rekordów; `index_repository.py`
pozostaje starszym inwentaryzatorem. Nie zastępuj nimi pełnego odczytu przez
`offline_docs.py`. Zachowano inspekcję VBA, inspekcję skoroszytów, walidator,
jawnie uruchamiany test COM oraz opcjonalne sprawdzanie publicznych URL.
Sprawdzanie URL pozostaje wyłączone bez `--allow-network` i nie jest potrzebne
przy pracy offline. Oryginalne osiem plików narzędzi/helpera jest niezmienione.

Z katalogu skilla:
```powershell
py -3 scripts\validate_skill.py
py -3 -m unittest discover -s tests -v
py -3 tests\verify_snapshot.py --archive "<original-VBA-Docs-main.zip>" --out "<new-report.json>"
```
Ostatnie polecenie jest opcjonalne: porównuje każdy plik z archiwum wejściowym.
Testy są wykonywane na Linux; test dowiązań może wymagać dodatkowego uprawnienia
Windows. Weryfikator testów z `--enforce-os` wymaga Linux/libseccomp, ale samo
wyszukiwanie i odczyt nie wymagają tej biblioteki.

## Zakres i ograniczenia
Pełne pokrycie oznacza **wszystkie dokumenty w dostarczonym archiwum**, a nie
wszystkie możliwe API ani niezależnie potwierdzoną aktualność GitHub. Nie wykonano
kompilacji VBA, uruchomienia Excela/COM/UserForms ani rzeczywistego testu agenta.
Oryginalne braki i błędy dokumentacji pozostają jawne. Teksty zewnętrznych stron
Win32, VS Code i pywin32 nie są częścią repozytorium VBA-Docs.

Skill nie odcina sieci całemu klientowi AI i nie zmienia modelu chmurowego w lokalny.
Można uruchomić lokalne wyszukiwanie bez sieci; możliwość pracy samego modelu
bez internetu zależy od klienta i konfiguracji.

[Źródła](SOURCES.md) | [Jakość](QUALITY-REPORT.md) |
[Odbiór](ACCEPTANCE.md) | [Ograniczenia](LIMITATIONS.md) |
[Licencje](LICENSES.md) | [Historia](CHANGELOG.md)

## Local update candidate: namespace API preflight (2.0.1)

This repackaging adds `scripts/api_preflight.py`, its evidence guide and
regression tests. Check `FIVE-LIBRARY-AUDIT.md` for five-library coverage.
This is a local candidate, **not** an update to the GitHub v2.0.0 release.
The separate knowledge ZIP and its Microsoft documentation have not changed.

```powershell
$preflight = Join-Path $HOME '.copilot\skills\excel-vba-expert\scripts\api_preflight.py'
python $preflight --namespace msforms --symbol ListBox.ItemHeight --json
python $preflight --namespace vbide --symbol VBComponent --json
```
