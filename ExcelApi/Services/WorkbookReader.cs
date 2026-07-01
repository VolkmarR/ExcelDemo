using ClosedXML.Excel;
using ExcelApi.Models;

namespace ExcelApi.Services;

/// <summary>
/// Reads an .xlsx file with ClosedXML (MIT) into the shared WorkbookModel:
/// values (number-format applied), cached formula results, styles, merges and dimensions.
/// Read-only — the file is never modified.
/// </summary>
public static class WorkbookReader
{
    public static WorkbookModel Read(string path)
    {
        // FileShare.ReadWrite so a copy open in Excel does not block the preview.
        using var fs = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var wb = new XLWorkbook(fs);

        var sheets = new List<SheetModel>();
        foreach (var ws in wb.Worksheets)
            sheets.Add(ReadSheet(ws));
        return new WorkbookModel(sheets);
    }

    private static SheetModel ReadSheet(IXLWorksheet ws)
    {
        var last = ws.LastCellUsed();
        int rowCount = last?.Address.RowNumber ?? 0;
        int colCount = last?.Address.ColumnNumber ?? 0;

        // Baseline font: anything matching the sheet default is omitted to keep the payload small.
        string defFont = ws.Style.Font.FontName;
        double defSize = ws.Style.Font.FontSize;

        var cells = new List<CellModel>();
        foreach (var cell in ws.CellsUsed())
        {
            int r = cell.Address.RowNumber - 1;
            int c = cell.Address.ColumnNumber - 1;
            try
            {
                string text = cell.GetFormattedString();
                object? raw = RawValue(cell);
                string? formula = cell.HasFormula ? "=" + cell.FormulaA1 : null;
                var style = StyleOf(cell, defFont, defSize);
                cells.Add(new CellModel(r, c, text, raw, formula, style));
            }
            catch
            {
                // A single unreadable cell must never break the whole preview.
                cells.Add(new CellModel(r, c, "", null, null, null));
            }
        }

        var colWidths = new List<double>(colCount);
        for (int c = 1; c <= colCount; c++)
            colWidths.Add(Math.Round(ws.Column(c).Width * 7 + 5)); // Excel char units -> px

        var rowHeights = new List<double>(rowCount);
        for (int r = 1; r <= rowCount; r++)
            rowHeights.Add(Math.Round(ws.Row(r).Height * 4.0 / 3.0)); // points -> px

        var merges = new List<MergeModel>();
        foreach (var mr in ws.MergedRanges)
        {
            var a = mr.RangeAddress.FirstAddress;
            var b = mr.RangeAddress.LastAddress;
            merges.Add(new MergeModel(a.RowNumber - 1, a.ColumnNumber - 1, b.RowNumber - 1, b.ColumnNumber - 1));
        }

        // Freeze panes: read reliably by the ExcelJS path; best-effort null on the backend.
        FreezeModel? freeze = null;

        return new SheetModel(ws.Name, rowCount, colCount, cells, colWidths, rowHeights, merges, freeze);
    }

    private static object? RawValue(IXLCell cell) => cell.DataType switch
    {
        XLDataType.Number => cell.Value.GetNumber(),
        XLDataType.Boolean => cell.Value.GetBoolean(),
        XLDataType.DateTime => cell.Value.GetDateTime().ToString("o"),
        XLDataType.Text => cell.Value.GetText(),
        _ => null,
    };

    private static CellStyleModel? StyleOf(IXLCell cell, string defFont, double defSize)
    {
        var s = cell.Style;
        var f = s.Font;

        bool? bold = f.Bold ? true : null;
        bool? italic = f.Italic ? true : null;
        bool? underline = f.Underline != XLFontUnderlineValues.None ? true : null;

        string? color = ToHex(f.FontColor);
        if (color == "#000000") color = null; // black is the default; omit

        string? fontFamily = !string.IsNullOrEmpty(f.FontName) && f.FontName != defFont ? f.FontName : null;
        double? fontSize = Math.Abs(f.FontSize - defSize) > 0.01 ? f.FontSize : null;

        string? bg = null;
        if (s.Fill.PatternType != XLFillPatternValues.None)
        {
            bg = ToHex(s.Fill.BackgroundColor);
            if (bg == "#FFFFFF") bg = null; // white == no visible fill
        }

        string? hAlign = s.Alignment.Horizontal switch
        {
            XLAlignmentHorizontalValues.Left => "left",
            XLAlignmentHorizontalValues.Center => "center",
            XLAlignmentHorizontalValues.CenterContinuous => "center",
            XLAlignmentHorizontalValues.Right => "right",
            XLAlignmentHorizontalValues.Justify => "left",
            _ => null, // General -> let the grid auto-align (numbers right, text left)
        };
        string? vAlign = s.Alignment.Vertical switch
        {
            XLAlignmentVerticalValues.Top => "top",
            XLAlignmentVerticalValues.Center => "middle",
            _ => null, // Bottom is the Excel default
        };
        bool? wrap = s.Alignment.WrapText ? true : null;

        var border = BordersOf(s.Border);

        if (bold is null && italic is null && underline is null && color is null && bg is null
            && fontFamily is null && fontSize is null && hAlign is null && vAlign is null
            && wrap is null && border is null)
            return null;

        return new CellStyleModel(bold, italic, underline, color, bg, fontFamily, fontSize, hAlign, vAlign, wrap, border);
    }

    private static BordersModel? BordersOf(IXLBorder b)
    {
        var top = Side(b.TopBorder, b.TopBorderColor);
        var right = Side(b.RightBorder, b.RightBorderColor);
        var bottom = Side(b.BottomBorder, b.BottomBorderColor);
        var left = Side(b.LeftBorder, b.LeftBorderColor);
        if (top is null && right is null && bottom is null && left is null) return null;
        return new BordersModel(top, right, bottom, left);
    }

    private static BorderModel? Side(XLBorderStyleValues style, XLColor color)
    {
        if (style == XLBorderStyleValues.None) return null;
        return new BorderModel(style.ToString().ToLowerInvariant(), ToHex(color));
    }

    private static string? ToHex(XLColor color)
    {
        if (color is null) return null;
        try
        {
            var c = color.Color; // throws for pure theme colors -> caught below
            return $"#{c.R:X2}{c.G:X2}{c.B:X2}";
        }
        catch
        {
            return null;
        }
    }
}
