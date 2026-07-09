using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using ExcelApi.Models;

namespace ExcelApi.Services;

/// <summary>
/// Reads an .xlsx file with ClosedXML (MIT) into the shared WorkbookModel:
/// values (number-format applied), cached formula results, styles, merges and dimensions.
/// Read-only — the file is never modified.
/// </summary>
public sealed class ClosedXmlWorkbookReader : IWorkbookReader
{
    public WorkbookModel Read(string path)
    {
        using var fs = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        return Read(fs);
    }

    public WorkbookModel Read(Stream stream)
    {
        if (stream.CanSeek)
            stream.Position = 0;

        using var wb = new XLWorkbook(stream);

        var sheets = new List<SheetModel>();

        foreach (var ws in wb.Worksheets)
            sheets.Add(ReadSheet(ws));

        return new WorkbookModel(sheets, ThemeOf(wb));
    }

    // Resolved theme palette (accents + fore/back), so the frontend can render exact
    // table-style colors. Falls back to null (frontend uses its own default) on any error.
    private static ThemePalette? ThemeOf(XLWorkbook wb)
    {
        try
        {
            var t = wb.Theme;
            if (t is null) return null;
            return new ThemePalette(
                ToHex(t.Accent1) ?? "#4472C4",
                ToHex(t.Accent2) ?? "#ED7D31",
                ToHex(t.Accent3) ?? "#A5A5A5",
                ToHex(t.Accent4) ?? "#FFC000",
                ToHex(t.Accent5) ?? "#5B9BD5",
                ToHex(t.Accent6) ?? "#70AD47",
                ToHex(t.Text1) ?? "#000000",
                ToHex(t.Background1) ?? "#FFFFFF",
                ToHex(t.Text2) ?? "#44546A",
                ToHex(t.Background2) ?? "#E7E6E6");
        }
        catch
        {
            return null;
        }
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

        var merges = new List<MergeModel>();
        foreach (var mr in ws.MergedRanges)
        {
            var a = mr.RangeAddress.FirstAddress;
            var b = mr.RangeAddress.LastAddress;
            merges.Add(new MergeModel(a.RowNumber - 1, a.ColumnNumber - 1, b.RowNumber - 1, b.ColumnNumber - 1));
        }

        var pictures = ReadPictures(ws, out int picMaxCol, out int picMaxRow);

        // Extend the used range so overlays that reach past the last content cell
        // (wide merges, floating pictures) have grid geometry to anchor against.
        foreach (var m in merges)
        {
            colCount = Math.Max(colCount, m.C1 + 1);
            rowCount = Math.Max(rowCount, m.R1 + 1);
        }

        colCount = Math.Max(colCount, picMaxCol);
        rowCount = Math.Max(rowCount, picMaxRow);

        var colWidths = new List<double>(colCount);
        for (int c = 1; c <= colCount; c++)
            colWidths.Add(Math.Round(ws.Column(c).Width * 7 + 5)); // Excel char units -> px

        var rowHeights = new List<double>(rowCount);
        for (int r = 1; r <= rowCount; r++)
            rowHeights.Add(Math.Round(ws.Row(r).Height * 4.0 / 3.0)); // points -> px

        // Freeze panes: ClosedXML exposes the frozen row/column counts on the sheet view.
        // When panes are frozen these are > 0 (a plain split leaves them 0 for this file).
        var sv = ws.SheetView;
        FreezeModel? freeze = (sv.SplitRow > 0 || sv.SplitColumn > 0)
            ? new FreezeModel(sv.SplitRow, sv.SplitColumn)
            : null;

        var tables = ReadTables(ws);
        var conditionalFormats = ReadConditionalFormats(ws);

        return new SheetModel(ws.Name, rowCount, colCount, cells, colWidths, rowHeights, merges, freeze, pictures,
            tables, conditionalFormats);
    }

    // Excel Tables (ListObjects): name, range, built-in style name and the display flags.
    // Exact colors are resolved on the frontend from StyleName + the theme palette.
    private static List<TableModel> ReadTables(IXLWorksheet ws)
    {
        var tables = new List<TableModel>();
        foreach (var t in ws.Tables)
        {
            try
            {
                var a = t.RangeAddress.FirstAddress;
                var b = t.RangeAddress.LastAddress;
                tables.Add(new TableModel(
                    t.Name,
                    new RangeModel(a.RowNumber - 1, a.ColumnNumber - 1, b.RowNumber - 1, b.ColumnNumber - 1),
                    t.Theme?.Name,
                    t.ShowHeaderRow,
                    t.ShowTotalsRow,
                    t.ShowRowStripes,
                    t.ShowColumnStripes,
                    t.EmphasizeFirstColumn,
                    t.EmphasizeLastColumn));
            }
            catch
            {
                // A single unreadable table must never break the preview.
            }
        }

        return tables;
    }

    // Conditional formatting: only color scales are emitted for now. The stop colors are
    // read here; the actual min/max (or percentile) is computed from the data on the frontend.
    private static List<ConditionalFormatModel> ReadConditionalFormats(IXLWorksheet ws)
    {
        var result = new List<ConditionalFormatModel>();
        foreach (var cf in ws.ConditionalFormats)
        {
            try
            {
                if (cf.ConditionalFormatType != XLConditionalFormatType.ColorScale) continue;

                // Stop colors ordered by key (1 = min, [2 = mid], last = max).
                var byKey = new SortedDictionary<int, string?>();
                foreach (var kv in cf.Colors) byKey[kv.Key] = ToHex(kv.Value);
                var colors = byKey.Values.ToList();
                if (colors.Count < 2) continue;

                var stops = new List<ColorScaleStopModel>();
                for (int i = 0; i < colors.Count; i++)
                {
                    string kind = i == 0 ? "min" : i == colors.Count - 1 ? "max" : "percentile";
                    double? val = kind == "percentile" ? 50 : null;
                    stops.Add(new ColorScaleStopModel(kind, val, colors[i] ?? "#FFFFFF"));
                }

                var ranges = cf.Ranges != null && cf.Ranges.Any() ? cf.Ranges.AsEnumerable() : new[] { cf.Range };
                foreach (var rng in ranges)
                {
                    if (rng is null) continue;
                    var a = rng.RangeAddress.FirstAddress;
                    var b = rng.RangeAddress.LastAddress;
                    result.Add(new ConditionalFormatModel(
                        "colorScale",
                        new RangeModel(a.RowNumber - 1, a.ColumnNumber - 1, b.RowNumber - 1, b.ColumnNumber - 1),
                        stops));
                }
            }
            catch
            {
                // A single unreadable rule must never break the preview.
            }
        }

        return result;
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

        return new CellStyleModel(bold, italic, underline, color, bg, fontFamily, fontSize, hAlign, vAlign, wrap,
            border);
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

    // Floating pictures via ClosedXML's native API (no separate parser). Emits each
    // image as a base64 data URL plus its pixel anchor; also reports how far the
    // pictures reach so the caller can grow the sheet's used range to fit them.
    private static List<PictureModel> ReadPictures(IXLWorksheet ws, out int maxCol, out int maxRow)
    {
        maxCol = 0;
        maxRow = 0;
        var pictures = new List<PictureModel>();
        foreach (var pic in ws.Pictures)
        {
            try
            {
                byte[] bytes = pic.ImageStream.ToArray();
                string src = $"data:{MimeOf(pic.Format)};base64,{Convert.ToBase64String(bytes)}";
                int fromCol = pic.TopLeftCell.Address.ColumnNumber - 1;
                int fromRow = pic.TopLeftCell.Address.RowNumber - 1;
                pictures.Add(new PictureModel(src, fromCol, fromRow, pic.Left, pic.Top, pic.Width, pic.Height));

                // Grow the used range to fit the picture. BottomRightCell throws for
                // non-resizing anchors, so walk widths/heights from the pixel extent instead.
                maxCol = Math.Max(maxCol,
                    FarIndex(fromCol, pic.Left + pic.Width, i => Math.Round(ws.Column(i + 1).Width * 7 + 5)));
                maxRow = Math.Max(maxRow,
                    FarIndex(fromRow, pic.Top + pic.Height, i => Math.Round(ws.Row(i + 1).Height * 4.0 / 3.0)));
            }
            catch
            {
                // A single unreadable picture must never break the preview.
            }
        }

        return pictures;
    }

    // Cell count (from a 0-based start) needed to span `extentPx`, walking cell sizes.
    private static int FarIndex(int start, double extentPx, Func<int, double> sizePx)
    {
        double acc = 0;
        int i = start;
        while (acc < extentPx && i < start + 1000)
        {
            acc += sizePx(i);
            i++;
        }

        return i + 1;
    }

    private static string MimeOf(XLPictureFormat fmt) => fmt switch
    {
        XLPictureFormat.Jpeg => "image/jpeg",
        XLPictureFormat.Png => "image/png",
        XLPictureFormat.Gif => "image/gif",
        XLPictureFormat.Bmp => "image/bmp",
        XLPictureFormat.Tiff => "image/tiff",
        _ => "application/octet-stream",
    };
}