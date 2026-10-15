// Użycie: XlsxToJira <plik.xlsx> <od> <do> [--md] [--sheet <nazwa|nr>]
// Wiersz <od> traktowany jest jako nagłówek tabeli.
using ClosedXML.Excel;

if (args.Length < 3 || !int.TryParse(args[1], out var from) || !int.TryParse(args[2], out var to) || from > to)
{
    Console.Error.WriteLine("Użycie: XlsxToJira <plik.xlsx> <od> <do> [--md] [--sheet <nazwa|nr>]");
    return 1;
}

var markdown = args.Contains("--md");
var sheetIdx = Array.IndexOf(args, "--sheet");
var sheetArg = sheetIdx >= 0 && sheetIdx + 1 < args.Length ? args[sheetIdx + 1] : null;

using var stream = new FileStream(args[0], FileMode.Open, FileAccess.Read, FileShare.ReadWrite); // działa też gdy plik otwarty w Excelu
using var wb = new XLWorkbook(stream);
var ws = sheetArg is null ? wb.Worksheet(1)
    : int.TryParse(sheetArg, out var n) ? wb.Worksheet(n) : wb.Worksheet(sheetArg);

var rows = Enumerable.Range(from, to - from + 1).Select(ws.Row).ToList();
var lastCol = rows.Max(r => r.LastCellUsed()?.Address.ColumnNumber ?? 0);
var firstCol = rows.Where(r => r.FirstCellUsed() != null).Select(r => r.FirstCellUsed()!.Address.ColumnNumber).DefaultIfEmpty(1).Min();

string Cell(IXLRow r, int c)
{
    var s = r.Cell(c).GetFormattedString().Replace("\r", "").Replace("\n", markdown ? "<br>" : " \\\\ ").Trim();
    s = s.Replace("|", markdown ? "\\|" : "\\|");
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
return 0;
