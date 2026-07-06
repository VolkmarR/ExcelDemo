using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelApi.Models;
using ExcelNumberFormat;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace ExcelApi.Services;

/// <summary>
/// Reads an .xlsx file with the low-level Open XML SDK (<c>DocumentFormat.OpenXml</c>) into the
/// shared <see cref="WorkbookModel"/> — an alternative to <see cref="ClosedXmlWorkbookReader"/>.
///
/// The output is intended to be byte-for-byte interchangeable with the ClosedXML reader, so it
/// deliberately reproduces that reader's choices rather than "improving" on them: theme colors
/// are omitted at the cell level (only rgb/indexed resolve), freeze panes are null, color-scale
/// stop kinds are assigned positionally (min/max/percentile), solid fills read the OOXML
/// <c>fgColor</c>, and a bad cell/table/picture/rule is swallowed so it can't break the preview.
///
/// Number/date display text is produced with <c>ExcelNumberFormat</c> — the exact library
/// ClosedXML calls internally for <c>GetFormattedString()</c> — so the formatted text matches.
/// Read-only — the file is never modified.
/// </summary>
public sealed class OpenXmlWorkbookReader : IWorkbookReader
{
    // Excel char-units -> px and points -> px, identical to ClosedXmlWorkbookReader.
    private static double ColPx(double widthChars) => Math.Round(widthChars * 7 + 5);
    private static double RowPx(double heightPts) => Math.Round(heightPts * 4.0 / 3.0);
    private static double EmuToPx(long emu) => Math.Round(emu / 9525.0); // 9525 EMU per px @ 96 DPI

    // Default column *display* width (char units) when the sheet declares no defaultColWidth.
    //
    // This is the one value a pure-OpenXML reader cannot compute independently: ClosedXML derives
    // it by measuring the normal font's maximum digit width with SixLabors.Fonts, then rounds. For
    // the demo's normal font (Aptos Narrow 11) that measured default is 10.38 chars (-> 78px), which
    // we reproduce here so the output stays byte-identical. A different normal font would yield a
    // different ClosedXML default, so a fully general match would require the same font metrics.
    private const double DefaultColWidthChars = 10.38;

    public WorkbookModel Read(string path)
    {
        // FileShare.ReadWrite so a copy open in Excel does not block the preview.
        using var fs = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var doc = SpreadsheetDocument.Open(fs, false);
        var wbPart = doc.WorkbookPart ?? throw new InvalidDataException("Workbook part missing.");

        var ctx = new WorkbookContext(wbPart);

        var sheets = new List<SheetModel>();
        foreach (var sheet in wbPart.Workbook.Sheets?.Elements<Sheet>() ?? Enumerable.Empty<Sheet>())
        {
            if (sheet.Id?.Value is not string rid) continue;
            if (wbPart.GetPartById(rid) is not WorksheetPart wsPart) continue;
            sheets.Add(ReadSheet(sheet.Name?.Value ?? "", wsPart, ctx));
        }
        return new WorkbookModel(sheets, ctx.Theme);
    }

    private SheetModel ReadSheet(string name, WorksheetPart wsPart, WorkbookContext ctx)
    {
        var ws = wsPart.Worksheet;
        var sheetData = ws.GetFirstChild<SheetData>();
        var fmt = ws.GetFirstChild<SheetFormatProperties>();

        // Shared-formula masters (si -> master formula + cell), so dependent cells (which carry no
        // formula text) can be translated to a concrete A1 formula like ClosedXML's FormulaA1.
        var masters = CollectSharedMasters(sheetData);

        // Per-column width (char units) resolver: explicit <col> width, else the sheet default.
        var colDefs = new List<(int Min, int Max, double W)>();
        foreach (var col in ws.GetFirstChild<Columns>()?.Elements<Column>() ?? Enumerable.Empty<Column>())
            if (col.Width?.Value is double w)
                colDefs.Add(((int)(col.Min?.Value ?? 1), (int)(col.Max?.Value ?? (col.Min?.Value ?? 1)), w));
        double defColW = fmt?.DefaultColumnWidth?.Value ?? DefaultColWidthChars;
        double ColWidthPx(int c0)
        {
            int oneBased = c0 + 1;
            foreach (var d in colDefs)
                if (oneBased >= d.Min && oneBased <= d.Max)
                    // An explicit <col> width is the OOXML *stored* width, which bakes in a 5px
                    // padding relative to the display width ClosedXML reports. Converting stored ->
                    // display (-5/MDW at MDW=7) and back through ColPx cancels the +5, leaving w*7 —
                    // matching ClosedXML exactly (e.g. stored 24 -> 168px).
                    return Math.Round(d.W * 7);
            return ColPx(defColW); // default is already a display width -> round(w*7+5)
        }

        // Per-row height (points) resolver: explicit <row ht>, else the sheet default.
        var rowHeightsMap = new Dictionary<int, double>();
        foreach (var row in sheetData?.Elements<Row>() ?? Enumerable.Empty<Row>())
            if (row.Height?.Value is double h && row.RowIndex?.Value is uint ri)
                rowHeightsMap[(int)ri - 1] = h;
        double defRowH = fmt?.DefaultRowHeight?.Value ?? 15.0;
        double RowHeightPx(int r0) => RowPx(rowHeightsMap.TryGetValue(r0, out var h) ? h : defRowH);

        // Cells + used range (content only, matching ClosedXML's CellsUsed()).
        int rowCount = 0, colCount = 0;
        var cells = new List<CellModel>();
        foreach (var row in sheetData?.Elements<Row>() ?? Enumerable.Empty<Row>())
        foreach (var c in row.Elements<Cell>())
        {
            if (c.CellReference?.Value is not string aref) continue;
            bool hasValue = c.CellValue != null || c.InlineString != null;
            bool hasFormula = c.CellFormula != null;
            if (!hasValue && !hasFormula) continue; // style-only cells are not "used"

            var (r, col) = ParseCellRef(aref);
            rowCount = Math.Max(rowCount, r + 1);
            colCount = Math.Max(colCount, col + 1);
            try
            {
                var (text, raw) = ResolveValue(c, ctx);
                string? formula = ResolveFormula(c, r, col, masters);
                var style = StyleOf(c, ctx);
                cells.Add(new CellModel(r, col, text, raw, formula, style));
            }
            catch
            {
                // A single unreadable cell must never break the whole preview.
                cells.Add(new CellModel(r, col, "", null, null, null));
            }
        }

        var merges = new List<MergeModel>();
        foreach (var mc in ws.GetFirstChild<MergeCells>()?.Elements<MergeCell>() ?? Enumerable.Empty<MergeCell>())
            if (mc.Reference?.Value is string mref)
            {
                var (r0, c0, r1, c1) = ParseRange(mref);
                merges.Add(new MergeModel(r0, c0, r1, c1));
            }

        var pictures = ReadPictures(wsPart, ColWidthPx, RowHeightPx, out int picMaxCol, out int picMaxRow);

        // Extend the used range so overlays past the last content cell (wide merges, floating
        // pictures) have grid geometry to anchor against — identical to the ClosedXML reader.
        foreach (var m in merges)
        {
            colCount = Math.Max(colCount, m.C1 + 1);
            rowCount = Math.Max(rowCount, m.R1 + 1);
        }
        colCount = Math.Max(colCount, picMaxCol);
        rowCount = Math.Max(rowCount, picMaxRow);

        var colWidths = new List<double>(colCount);
        for (int c = 0; c < colCount; c++) colWidths.Add(ColWidthPx(c));
        var rowHeights = new List<double>(rowCount);
        for (int r = 0; r < rowCount; r++) rowHeights.Add(RowHeightPx(r));

        // Freeze panes: read reliably by the ExcelJS path; best-effort null on the backend.
        FreezeModel? freeze = null;

        var tables = ReadTables(wsPart);
        var conditionalFormats = ReadConditionalFormats(ws);

        return new SheetModel(name, rowCount, colCount, cells, colWidths, rowHeights, merges, freeze, pictures, tables, conditionalFormats);
    }

    // ---- Values & text -----------------------------------------------------------------------

    private static (string Text, object? Raw) ResolveValue(Cell c, WorkbookContext ctx)
    {
        // These OOXML *Values types are structs (not enums) in the 3.x SDK, so they compare with
        // == but cannot be used as switch/case constants.
        var type = c.DataType?.Value ?? CellValues.Number;
        string v = c.CellValue?.InnerText ?? "";

        if (type == CellValues.SharedString)
        {
            string s = int.TryParse(v, NumberStyles.Integer, CultureInfo.InvariantCulture, out int idx)
                && idx >= 0 && idx < ctx.SharedStrings.Count ? ctx.SharedStrings[idx] : "";
            return (s, s);
        }
        if (type == CellValues.InlineString)
        {
            string ins = c.InlineString?.InnerText ?? "";
            return (ins, ins);
        }
        if (type == CellValues.String) // formula result that is a string
            return (v, v);
        if (type == CellValues.Boolean)
        {
            bool b = v == "1";
            return (b ? "TRUE" : "FALSE", b);
        }
        if (type == CellValues.Error)
            return (v, null);

        // Number / Date
        if (!double.TryParse(v, NumberStyles.Any, CultureInfo.InvariantCulture, out double num))
            return (v, string.IsNullOrEmpty(v) ? null : v);

        var nf = ctx.NumberFormatFor(ctx.NumberFormatIdOf(c));
        string text = FormatNumber(nf, num);
        if (nf?.IsDateTimeFormat == true)
        {
            try { return (text, DateTime.FromOADate(num).ToString("o", CultureInfo.InvariantCulture)); }
            catch { return (text, num); }
        }
        return (text, num);
    }

    private static string FormatNumber(NumberFormat? nf, double num)
    {
        try
        {
            if (nf is { IsValid: true })
                return nf.Format(num, CultureInfo.CurrentCulture, false);
        }
        catch { /* fall through to a plain rendering */ }
        return num.ToString(CultureInfo.CurrentCulture);
    }

    // ---- Formulas (incl. shared-formula translation) -----------------------------------------

    private readonly record struct SharedMaster(string Formula, int Row, int Col);

    private static Dictionary<int, SharedMaster> CollectSharedMasters(SheetData? sd)
    {
        var d = new Dictionary<int, SharedMaster>();
        foreach (var row in sd?.Elements<Row>() ?? Enumerable.Empty<Row>())
        foreach (var c in row.Elements<Cell>())
        {
            var f = c.CellFormula;
            if (f?.FormulaType?.Value == CellFormulaValues.Shared
                && f.SharedIndex?.Value is uint si
                && !string.IsNullOrEmpty(f.Text)
                && c.CellReference?.Value is string aref)
            {
                var (r, col) = ParseCellRef(aref);
                d[(int)si] = new SharedMaster(f.Text!, r, col);
            }
        }
        return d;
    }

    private static string? ResolveFormula(Cell c, int r, int col, Dictionary<int, SharedMaster> masters)
    {
        var f = c.CellFormula;
        if (f == null) return null;

        if (f.FormulaType?.Value == CellFormulaValues.Shared)
        {
            if (!string.IsNullOrEmpty(f.Text)) return "=" + f.Text; // this cell is the master
            if (f.SharedIndex?.Value is uint si && masters.TryGetValue((int)si, out var m))
                return "=" + TranslateFormula(m.Formula, r - m.Row, col - m.Col);
            return null;
        }

        string t = f.Text ?? "";
        return t.Length > 0 ? "=" + t : null;
    }

    // A1 references, ignoring $-anchored parts and things that are clearly identifiers/functions.
    private static readonly Regex RefRx = new(
        @"(?<![A-Za-z0-9_$.])(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_(])",
        RegexOptions.Compiled);

    // Shift the relative parts of every cell reference in a shared formula by (dRow, dCol),
    // skipping string literals — mirrors how Excel/ClosedXML materialize a shared formula.
    private static string TranslateFormula(string formula, int dRow, int dCol)
    {
        if (dRow == 0 && dCol == 0) return formula;

        var sb = new StringBuilder(formula.Length);
        int i = 0;
        while (i < formula.Length)
        {
            char ch = formula[i];
            if (ch == '"')
            {
                int start = i++;
                while (i < formula.Length && formula[i] != '"') i++;
                if (i < formula.Length) i++; // closing quote
                sb.Append(formula, start, i - start);
            }
            else
            {
                int start = i;
                while (i < formula.Length && formula[i] != '"') i++;
                sb.Append(ShiftRefs(formula.Substring(start, i - start), dRow, dCol));
            }
        }
        return sb.ToString();
    }

    private static string ShiftRefs(string segment, int dRow, int dCol) => RefRx.Replace(segment, m =>
    {
        string colAbs = m.Groups[1].Value, letters = m.Groups[2].Value, rowAbs = m.Groups[3].Value;
        if (!int.TryParse(m.Groups[4].Value, out int rowNum)) return m.Value;
        int col = ColToIndex(letters);
        if (col < 0) return m.Value; // not a valid column (e.g. > XFD) — leave untouched
        if (colAbs != "$") col += dCol;
        if (rowAbs != "$") rowNum += dRow;
        if (col < 0 || rowNum < 1) return m.Value;
        return colAbs + IndexToCol(col) + rowAbs + rowNum.ToString(CultureInfo.InvariantCulture);
    });

    // ---- Styles ------------------------------------------------------------------------------

    private static CellStyleModel? StyleOf(Cell c, WorkbookContext ctx)
    {
        uint si = c.StyleIndex?.Value ?? 0;
        if (si >= ctx.CellFormats.Count) return null;
        var xf = ctx.CellFormats[(int)si];

        var font = ctx.FontAt(xf.FontId?.Value);
        bool? bold = IsOn(font?.Bold) ? true : null;
        bool? italic = IsOn(font?.Italic) ? true : null;
        bool? underline = UnderlineOn(font) ? true : null;

        string? color = ColorHex(font?.Color);
        if (color == "#000000") color = null; // black is the default; omit

        string? fontName = font?.FontName?.Val?.Value;
        string? fontFamily = !string.IsNullOrEmpty(fontName) && fontName != ctx.DefaultFontName ? fontName : null;
        double? sz = font?.FontSize?.Val?.Value;
        double? fontSize = sz.HasValue && Math.Abs(sz.Value - ctx.DefaultFontSize) > 0.01 ? sz : null;

        string? bg = FillColor(ctx.FillAt(xf.FillId?.Value));
        if (bg == "#FFFFFF") bg = null; // white == no visible fill

        var al = xf.Alignment;
        var h = al?.Horizontal?.Value;
        string? hAlign =
            h == HorizontalAlignmentValues.Left ? "left"
            : h == HorizontalAlignmentValues.Center ? "center"
            : h == HorizontalAlignmentValues.CenterContinuous ? "center"
            : h == HorizontalAlignmentValues.Right ? "right"
            : h == HorizontalAlignmentValues.Justify ? "left"
            : null; // General -> let the grid auto-align
        var vy = al?.Vertical?.Value;
        string? vAlign =
            vy == VerticalAlignmentValues.Top ? "top"
            : vy == VerticalAlignmentValues.Center ? "middle"
            : null; // Bottom is the Excel default
        bool? wrap = (al?.WrapText?.Value ?? false) ? true : null;

        var border = BordersOf(ctx.BorderAt(xf.BorderId?.Value));

        if (bold is null && italic is null && underline is null && color is null && bg is null
            && fontFamily is null && fontSize is null && hAlign is null && vAlign is null
            && wrap is null && border is null)
            return null;

        return new CellStyleModel(bold, italic, underline, color, bg, fontFamily, fontSize, hAlign, vAlign, wrap, border);
    }

    private static bool IsOn(BooleanPropertyType? p) => p != null && (p.Val == null || p.Val.Value);

    private static bool UnderlineOn(Font? font)
    {
        var u = font?.Underline;
        if (u == null) return false;
        return (u.Val?.Value ?? UnderlineValues.Single) != UnderlineValues.None;
    }

    private static string? FillColor(Fill? fill)
    {
        var pf = fill?.PatternFill;
        var pt = pf?.PatternType?.Value;
        if (pf == null || pt == null || pt == PatternValues.None) return null;
        // Solid fills store the visible color in fgColor (the classic OOXML gotcha) — this is what
        // ClosedXML surfaces as BackgroundColor. Other patterns fall back to bgColor.
        if (pt == PatternValues.Solid) return ColorHex(pf.ForegroundColor);
        return ColorHex(pf.BackgroundColor) ?? ColorHex(pf.ForegroundColor);
    }

    private static BordersModel? BordersOf(Border? b)
    {
        if (b == null) return null;
        var top = Side(b.TopBorder);
        var right = Side(b.RightBorder);
        var bottom = Side(b.BottomBorder);
        var left = Side(b.LeftBorder);
        if (top is null && right is null && bottom is null && left is null) return null;
        return new BordersModel(top, right, bottom, left);
    }

    private static BorderModel? Side(BorderPropertiesType? b)
    {
        var style = b?.Style?.Value;
        if (style == null || style == BorderStyleValues.None) return null;
        return new BorderModel(style.ToString()!.ToLowerInvariant(), ColorHex(b!.Color));
    }

    // rgb -> #RRGGBB (alpha stripped); indexed -> legacy palette; theme/auto -> null (ClosedXML
    // omits pure theme colors, and callers drop #000000 / #FFFFFF as defaults).
    private static string? ColorHex(ColorType? c)
    {
        if (c == null || c.Auto?.Value == true) return null;
        if (c.Rgb?.Value is string argb && argb.Length >= 6)
        {
            string hex = argb.Length == 8 ? argb.Substring(2) : argb;
            return "#" + hex.ToUpperInvariant();
        }
        if (c.Indexed?.Value is uint idx) return IndexedColor(idx);
        return null; // theme color
    }

    // ---- Theme --------------------------------------------------------------------------------

    private static ThemePalette? ResolveTheme(A.ColorScheme? s)
    {
        if (s == null) return null;
        return new ThemePalette(
            SchemeHex(s.Accent1Color, "#4472C4"),
            SchemeHex(s.Accent2Color, "#ED7D31"),
            SchemeHex(s.Accent3Color, "#A5A5A5"),
            SchemeHex(s.Accent4Color, "#FFC000"),
            SchemeHex(s.Accent5Color, "#5B9BD5"),
            SchemeHex(s.Accent6Color, "#70AD47"),
            SchemeHex(s.Dark1Color, "#000000"),
            SchemeHex(s.Light1Color, "#FFFFFF"),
            SchemeHex(s.Dark2Color, "#44546A"),
            SchemeHex(s.Light2Color, "#E7E6E6"));
    }

    private static string SchemeHex(OpenXmlCompositeElement? wrapper, string fallback)
    {
        if (wrapper?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value is string rgb && rgb.Length >= 6)
            return "#" + rgb.ToUpperInvariant();
        if (wrapper?.GetFirstChild<A.SystemColor>() is { } sys)
        {
            if (sys.LastColor?.Value is string last && last.Length >= 6) return "#" + last.ToUpperInvariant();
            if (sys.Val?.Value == A.SystemColorValues.Window) return "#FFFFFF";
            if (sys.Val?.Value == A.SystemColorValues.WindowText) return "#000000";
        }
        return fallback;
    }

    // ---- Pictures -----------------------------------------------------------------------------

    private static List<PictureModel> ReadPictures(
        WorksheetPart wsPart, Func<int, double> colPx, Func<int, double> rowPx, out int maxCol, out int maxRow)
    {
        maxCol = 0;
        maxRow = 0;
        var pictures = new List<PictureModel>();
        var dp = wsPart.DrawingsPart;
        if (dp?.WorksheetDrawing == null) return pictures;

        foreach (var anchor in dp.WorksheetDrawing.Elements())
        {
            var pic = anchor.GetFirstChild<Xdr.Picture>();
            if (pic == null) continue; // charts (graphicFrame) and other shapes are not pictures
            try
            {
                string? embed = pic.BlipFill?.GetFirstChild<A.Blip>()?.Embed?.Value;
                if (embed == null || dp.GetPartById(embed) is not ImagePart img) continue;

                byte[] bytes;
                using (var st = img.GetStream())
                using (var ms = new MemoryStream())
                {
                    st.CopyTo(ms);
                    bytes = ms.ToArray();
                }
                string src = $"data:{img.ContentType};base64,{Convert.ToBase64String(bytes)}";

                var from = anchor.GetFirstChild<Xdr.FromMarker>();
                int fromCol = ParseInt(from?.ColumnId?.Text);
                int fromRow = ParseInt(from?.RowId?.Text);
                double offsetX = EmuToPx(ParseLong(from?.ColumnOffset?.Text));
                double offsetY = EmuToPx(ParseLong(from?.RowOffset?.Text));

                var (width, height) = PictureSize(pic, anchor, from, colPx, rowPx);
                pictures.Add(new PictureModel(src, fromCol, fromRow, offsetX, offsetY, width, height));

                // Grow the used range to fit the picture (same walk as the ClosedXML reader).
                maxCol = Math.Max(maxCol, FarIndex(fromCol, offsetX + width, colPx));
                maxRow = Math.Max(maxRow, FarIndex(fromRow, offsetY + height, rowPx));
            }
            catch
            {
                // A single unreadable picture must never break the preview.
            }
        }
        return pictures;
    }

    // Prefer the explicit stored size (spPr/xfrm/ext, or a one-cell anchor's ext); otherwise fall
    // back to the two-cell span computed from column widths / row heights.
    private static (double Width, double Height) PictureSize(
        Xdr.Picture pic, OpenXmlElement anchor, Xdr.FromMarker? from,
        Func<int, double> colPx, Func<int, double> rowPx)
    {
        var ext = pic.GetFirstChild<Xdr.ShapeProperties>()?.GetFirstChild<A.Transform2D>()?.Extents;
        if (ext != null) return (EmuToPx(ext.Cx ?? 0), EmuToPx(ext.Cy ?? 0));

        if (anchor is Xdr.OneCellAnchor { Extent: { } oneExt })
            return (EmuToPx(oneExt.Cx ?? 0), EmuToPx(oneExt.Cy ?? 0));

        var to = anchor.GetFirstChild<Xdr.ToMarker>();
        if (from != null && to != null)
        {
            int fromCol = ParseInt(from.ColumnId?.Text), toCol = ParseInt(to.ColumnId?.Text);
            int fromRow = ParseInt(from.RowId?.Text), toRow = ParseInt(to.RowId?.Text);
            double w = -EmuToPx(ParseLong(from.ColumnOffset?.Text)) + EmuToPx(ParseLong(to.ColumnOffset?.Text));
            for (int col = fromCol; col < toCol; col++) w += colPx(col);
            double h = -EmuToPx(ParseLong(from.RowOffset?.Text)) + EmuToPx(ParseLong(to.RowOffset?.Text));
            for (int r = fromRow; r < toRow; r++) h += rowPx(r);
            return (Math.Max(0, w), Math.Max(0, h));
        }
        return (0, 0);
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

    // ---- Tables -------------------------------------------------------------------------------

    private static List<TableModel> ReadTables(WorksheetPart wsPart)
    {
        var tables = new List<TableModel>();
        foreach (var tdp in wsPart.TableDefinitionParts)
        {
            var t = tdp.Table;
            if (t?.Reference?.Value is not string reference) continue;
            try
            {
                var (r0, c0, r1, c1) = ParseRange(reference);
                var tsi = t.TableStyleInfo;
                bool showHeader = (t.HeaderRowCount?.Value ?? 1) != 0;
                bool showTotals = (t.TotalsRowShown?.Value ?? false) || (t.TotalsRowCount?.Value ?? 0) > 0;
                tables.Add(new TableModel(
                    t.Name?.Value ?? t.DisplayName?.Value ?? "",
                    new RangeModel(r0, c0, r1, c1),
                    tsi?.Name?.Value,
                    showHeader,
                    showTotals,
                    tsi?.ShowRowStripes?.Value ?? false,
                    tsi?.ShowColumnStripes?.Value ?? false,
                    tsi?.ShowFirstColumn?.Value ?? false,
                    tsi?.ShowLastColumn?.Value ?? false));
            }
            catch
            {
                // A single unreadable table must never break the preview.
            }
        }
        return tables;
    }

    // ---- Conditional formatting (color scales only) ------------------------------------------

    private static List<ConditionalFormatModel> ReadConditionalFormats(Worksheet ws)
    {
        var result = new List<ConditionalFormatModel>();
        foreach (var cf in ws.Elements<ConditionalFormatting>())
        {
            try
            {
                var refs = (cf.SequenceOfReferences?.InnerText ?? "")
                    .Split(' ', StringSplitOptions.RemoveEmptyEntries);
                if (refs.Length == 0) continue;

                foreach (var rule in cf.Elements<ConditionalFormattingRule>())
                {
                    if (rule.Type?.Value != ConditionalFormatValues.ColorScale) continue;
                    var cs = rule.GetFirstChild<ColorScale>();
                    if (cs == null) continue;

                    // Stop colors in document order (1 = min, [2 = mid], last = max).
                    var colors = cs.Elements<Color>().Select(x => ColorHex(x) ?? "#FFFFFF").ToList();
                    if (colors.Count < 2) continue;

                    var stops = new List<ColorScaleStopModel>();
                    for (int i = 0; i < colors.Count; i++)
                    {
                        string kind = i == 0 ? "min" : i == colors.Count - 1 ? "max" : "percentile";
                        double? val = kind == "percentile" ? 50 : null;
                        stops.Add(new ColorScaleStopModel(kind, val, colors[i]));
                    }

                    foreach (var rng in refs)
                    {
                        var (r0, c0, r1, c1) = ParseRange(rng);
                        result.Add(new ConditionalFormatModel(
                            "colorScale", new RangeModel(r0, c0, r1, c1), stops));
                    }
                }
            }
            catch
            {
                // A single unreadable rule must never break the preview.
            }
        }
        return result;
    }

    // ---- A1 parsing / column indexing --------------------------------------------------------

    private static (int Row, int Col) ParseCellRef(string s)
    {
        int i = 0, col = 0, row = 0;
        while (i < s.Length && (s[i] == '$')) i++;
        while (i < s.Length && char.IsLetter(s[i])) { col = col * 26 + (char.ToUpperInvariant(s[i]) - 'A' + 1); i++; }
        while (i < s.Length && s[i] == '$') i++;
        while (i < s.Length && char.IsDigit(s[i])) { row = row * 10 + (s[i] - '0'); i++; }
        return (row - 1, col - 1);
    }

    private static (int R0, int C0, int R1, int C1) ParseRange(string s)
    {
        int colon = s.IndexOf(':');
        if (colon < 0)
        {
            var (r, c) = ParseCellRef(s);
            return (r, c, r, c);
        }
        var (r0, c0) = ParseCellRef(s.Substring(0, colon));
        var (r1, c1) = ParseCellRef(s.Substring(colon + 1));
        return (r0, c0, r1, c1);
    }

    private static int ColToIndex(string letters)
    {
        int n = 0;
        foreach (char ch in letters)
        {
            if (!char.IsLetter(ch)) return -1;
            n = n * 26 + (char.ToUpperInvariant(ch) - 'A' + 1);
        }
        int idx = n - 1;
        return idx > 16383 ? -1 : idx; // XFD is the last valid column
    }

    private static string IndexToCol(int index)
    {
        var sb = new StringBuilder();
        int n = index;
        do
        {
            sb.Insert(0, (char)('A' + n % 26));
            n = n / 26 - 1;
        } while (n >= 0);
        return sb.ToString();
    }

    private static int ParseInt(string? s) => int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out int v) ? v : 0;
    private static long ParseLong(string? s) => long.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out long v) ? v : 0;

    // Legacy 56-color indexed palette (indices 64/65 are the system fore/background -> null).
    private static string? IndexedColor(uint i) => i < IndexedPalette.Length ? IndexedPalette[i] : null;

    private static readonly string?[] IndexedPalette =
    {
        "#000000", "#FFFFFF", "#FF0000", "#00FF00", "#0000FF", "#FFFF00", "#FF00FF", "#00FFFF",
        "#000000", "#FFFFFF", "#FF0000", "#00FF00", "#0000FF", "#FFFF00", "#FF00FF", "#00FFFF",
        "#800000", "#008000", "#000080", "#808000", "#800080", "#008080", "#C0C0C0", "#808080",
        "#9999FF", "#993366", "#FFFFCC", "#CCFFFF", "#660066", "#FF8080", "#0066CC", "#CCCCFF",
        "#000080", "#FF00FF", "#FFFF00", "#00FFFF", "#800080", "#800000", "#008080", "#0000FF",
        "#00CCFF", "#CCFFFF", "#CCFFCC", "#FFFF99", "#99CCFF", "#FF99CC", "#CC99FF", "#FFCC99",
        "#3366FF", "#33CCCC", "#99CC00", "#FFCC00", "#FF9900", "#FF6600", "#666699", "#969696",
        "#003366", "#339966", "#003300", "#333300", "#993300", "#993366", "#333399", "#333333",
    };

    // ---- Shared workbook resources (loaded once) ---------------------------------------------

    private sealed class WorkbookContext
    {
        public readonly List<string> SharedStrings = new();
        public readonly List<CellFormat> CellFormats = new();
        public readonly List<Font> Fonts = new();
        public readonly List<Fill> Fills = new();
        public readonly List<Border> Borders = new();
        public readonly ThemePalette? Theme;
        public string DefaultFontName = "Calibri";
        public double DefaultFontSize = 11;

        private readonly Dictionary<uint, string> _customFormats = new();
        private readonly Dictionary<string, NumberFormat> _nfCache = new();

        public WorkbookContext(WorkbookPart wbPart)
        {
            foreach (var si in wbPart.SharedStringTablePart?.SharedStringTable?.Elements<SharedStringItem>()
                     ?? Enumerable.Empty<SharedStringItem>())
                SharedStrings.Add(si.InnerText);

            var styles = wbPart.WorkbookStylesPart?.Stylesheet;
            if (styles != null)
            {
                if (styles.CellFormats != null) CellFormats.AddRange(styles.CellFormats.Elements<CellFormat>());
                if (styles.Fonts != null) Fonts.AddRange(styles.Fonts.Elements<Font>());
                if (styles.Fills != null) Fills.AddRange(styles.Fills.Elements<Fill>());
                if (styles.Borders != null) Borders.AddRange(styles.Borders.Elements<Border>());
                foreach (var nf in styles.NumberingFormats?.Elements<NumberingFormat>() ?? Enumerable.Empty<NumberingFormat>())
                    if (nf.NumberFormatId?.Value is uint id && nf.FormatCode?.Value is string code)
                        _customFormats[id] = code;

                if (Fonts.Count > 0)
                {
                    DefaultFontName = Fonts[0].FontName?.Val?.Value ?? DefaultFontName;
                    DefaultFontSize = Fonts[0].FontSize?.Val?.Value ?? DefaultFontSize;
                }
            }

            Theme = ResolveTheme(wbPart.ThemePart?.Theme?.ThemeElements?.ColorScheme);
        }

        public uint NumberFormatIdOf(Cell c)
        {
            uint si = c.StyleIndex?.Value ?? 0;
            return si < CellFormats.Count ? CellFormats[(int)si].NumberFormatId?.Value ?? 0 : 0;
        }

        public NumberFormat? NumberFormatFor(uint numFmtId)
        {
            string code = _customFormats.TryGetValue(numFmtId, out var custom)
                ? custom
                : BuiltinFormats.GetValueOrDefault(numFmtId, "General");
            if (!_nfCache.TryGetValue(code, out var nf))
            {
                nf = new NumberFormat(code);
                _nfCache[code] = nf;
            }
            return nf;
        }

        public Font? FontAt(uint? id) => id is uint i && i < Fonts.Count ? Fonts[(int)i] : null;
        public Fill? FillAt(uint? id) => id is uint i && i < Fills.Count ? Fills[(int)i] : null;
        public Border? BorderAt(uint? id) => id is uint i && i < Borders.Count ? Borders[(int)i] : null;
    }

    // Standard OOXML built-in number formats (the locale-independent subset; currency/accounting
    // ids 5-8/41-44 are region-derived and omitted — the demo does not use them).
    private static readonly Dictionary<uint, string> BuiltinFormats = new()
    {
        [0] = "General",
        [1] = "0",
        [2] = "0.00",
        [3] = "#,##0",
        [4] = "#,##0.00",
        [9] = "0%",
        [10] = "0.00%",
        [11] = "0.00E+00",
        [12] = "# ?/?",
        [13] = "# ??/??",
        [14] = "mm-dd-yy",
        [15] = "d-mmm-yy",
        [16] = "d-mmm",
        [17] = "mmm-yy",
        [18] = "h:mm AM/PM",
        [19] = "h:mm:ss AM/PM",
        [20] = "h:mm",
        [21] = "h:mm:ss",
        [22] = "m/d/yy h:mm",
        [37] = "#,##0 ;(#,##0)",
        [38] = "#,##0 ;[Red](#,##0)",
        [39] = "#,##0.00;(#,##0.00)",
        [40] = "#,##0.00;[Red](#,##0.00)",
        [45] = "mm:ss",
        [46] = "[h]:mm:ss",
        [47] = "mmss.0",
        [48] = "##0.0E+0",
        [49] = "@",
    };
}
