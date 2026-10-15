// Użycie: XlsxToJira [plik.xlsx] [od] [do] [--md] [--sheet <nazwa|nr>]
// Wiersz <od> traktowany jest jako nagłówek tabeli.
// Brakujące lub błędne parametry program dopyta interaktywnie.
using System.Text;
using ClosedXML.Excel;

Console.OutputEncoding = Encoding.UTF8;
Console.InputEncoding = Encoding.UTF8;

const string usage = "Użycie: XlsxToJira [plik.xlsx] [od] [do] [--md] [--sheet <nazwa|nr>]";

// --- parametry z linii poleceń -------------------------------------------
var flags = args.Where(a => a.StartsWith("--")).ToList();
var positional = new List<string>();
for (var i = 0; i < args.Length; i++)
{
    if (args[i] == "--sheet") { i++; continue; }      // wartość --sheet nie jest pozycyjna
    if (!args[i].StartsWith("--")) positional.Add(args[i]);
}

var markdown = flags.Contains("--md");
var sheetIdx = Array.IndexOf(args, "--sheet");
var sheetArg = sheetIdx >= 0 && sheetIdx + 1 < args.Length ? args[sheetIdx + 1] : null;

var path = positional.Count > 0 ? positional[0].Trim('"') : null;
int? from = positional.Count > 1 && int.TryParse(positional[1], out var f) && f > 0 ? f : null;
int? to = positional.Count > 2 && int.TryParse(positional[2], out var t) && t > 0 ? t : null;

var interactive = path is null || !File.Exists(path) || from is null || to is null || to < from;
if (interactive) Console.Error.WriteLine(usage + "\n");

// --- dopytanie o brakujące parametry -------------------------------------
while (path is null || !File.Exists(path))
{
    if (path is not null) Console.Error.WriteLine($"Nie znaleziono pliku: {path}");
    path = Ask("Ścieżka do pliku .xlsx")?.Trim('"');
    if (path is null) return 1;
}

using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite); // działa też gdy plik otwarty w Excelu
using var wb = new XLWorkbook(stream);

IXLWorksheet ws;
if (sheetArg is not null)
{
    ws = int.TryParse(sheetArg, out var n) ? wb.Worksheet(n) : wb.Worksheet(sheetArg);
}
else if (interactive && wb.Worksheets.Count > 1)
{
    Console.Error.WriteLine("Arkusze:");
    foreach (var s in wb.Worksheets) Console.Error.WriteLine($"  {s.Position}. {s.Name}");
    ws = wb.Worksheet(AskInt("Arkusz (nr)", 1, wb.Worksheets.Count, 1));
}
else
{
    ws = wb.Worksheet(1);
}

if (interactive)
{
    var last = ws.LastRowUsed()?.RowNumber() ?? 1;
    Console.Error.WriteLine($"Arkusz „{ws.Name}”: ostatni użyty wiersz = {last}");
    if (from is null || to is null || to < from)
    {
        from = AskInt("Wiersz nagłówka (od)", 1, last, from ?? 1);
        to = AskInt("Ostatni wiersz (do)", from.Value, last, Math.Max(from.Value, last));
    }
    if (!markdown)
        markdown = Ask("Format [Enter = Jira wiki markup, m = Markdown]")?.Trim().ToLowerInvariant() is "m" or "md";
    Console.Error.WriteLine();
}

// --- generowanie tabeli ---------------------------------------------------
var rows = Enumerable.Range(from!.Value, to!.Value - from.Value + 1).Select(ws.Row).ToList();
var lastCol = rows.Max(r => r.LastCellUsed()?.Address.ColumnNumber ?? 0);
var firstCol = rows.Where(r => r.FirstCellUsed() != null).Select(r => r.FirstCellUsed()!.Address.ColumnNumber).DefaultIfEmpty(1).Min();

string Cell(IXLRow r, int c)
{
    var s = r.Cell(c).GetFormattedString().Replace("\r", "").Replace("\n", markdown ? "<br>" : " \\\\ ").Trim();
    s = s.Replace("|", "\\|");
    return s.Length == 0 ? " " : s;
}

var cols = Enumerable.Range(firstCol, lastCol - firstCol + 1).ToList();
for (var i = 0; i < rows.Count; i++)
{
    var cells = cols.Select(c => Cell(rows[i], c));
    if (markdown)
    {
        Console.WriteLine("| " + string.Join(" | ", cells) + " |");
        if (i == 0) Console.WriteLine("|" + string.Join("|", cols.Select(_ => "---")) + "|");
    }
    else
    {
        Console.WriteLine(i == 0 ? "||" + string.Join("||", cells) + "||" : "|" + string.Join("|", cells) + "|");
    }
}

if (interactive && !Console.IsInputRedirected)
{
    Console.Error.Write("\nEnter, aby zakończyć...");
    Console.ReadLine();
}
return 0;

// --- pomocnicze -----------------------------------------------------------
static string? Ask(string prompt)
{
    Console.Error.Write($"{prompt}: ");
    return Console.ReadLine();
}

static int AskInt(string prompt, int min, int max, int @default)
{
    while (true)
    {
        var s = Ask($"{prompt} [{@default}]");
        if (s is null) Environment.Exit(1);
        if (string.IsNullOrWhiteSpace(s)) return @default;
        if (int.TryParse(s.Trim(), out var v) && v >= min && v <= max) return v;
        Console.Error.WriteLine($"Podaj liczbę z zakresu {min}–{max}.");
    }
}
