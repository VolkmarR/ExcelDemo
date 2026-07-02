namespace ExcelApi.Models;

// Shape mirrors excel-web/src/model.ts so the frontend consumes it directly.
// Minimal APIs serialize with JsonSerializerDefaults.Web => camelCase on the wire.

public record WorkbookModel(List<SheetModel> Sheets, ThemePalette? Theme);

public record SheetModel(
    string Name,
    int RowCount,
    int ColCount,
    List<CellModel> Cells,
    List<double> ColWidths,
    List<double> RowHeights,
    List<MergeModel> Merges,
    FreezeModel? Freeze,
    List<PictureModel>? Pictures,
    List<TableModel>? Tables,
    List<ConditionalFormatModel>? ConditionalFormats);

// r/c are 0-based. Text is the display string (number format applied).
// Raw carries the typed value (number/bool/string) so the grid can align numbers right.
public record CellModel(
    int R,
    int C,
    string Text,
    object? Raw,
    string? Formula,
    CellStyleModel? Style);

public record MergeModel(int R0, int C0, int R1, int C1);

// A floating picture anchored on the sheet. Src is a data URL (base64) so the
// browser renders it directly; position/offset/size are in pixels at 96 DPI.
public record PictureModel(
    string Src,
    int FromCol,
    int FromRow,
    double OffsetX,
    double OffsetY,
    double Width,
    double Height);

public record FreezeModel(int Rows, int Cols);

public record CellStyleModel(
    bool? Bold,
    bool? Italic,
    bool? Underline,
    string? Color,
    string? Bg,
    string? FontFamily,
    double? FontSize,
    string? HAlign,
    string? VAlign,
    bool? Wrap,
    BordersModel? Border);

public record BordersModel(BorderModel? Top, BorderModel? Right, BorderModel? Bottom, BorderModel? Left);

// Style is an Excel border-style token (e.g. "thin", "medium", "double"); the frontend maps it to CSS.
public record BorderModel(string Style, string? Color);

// A rectangular range, 0-based inclusive (mirrors MergeModel's shape).
public record RangeModel(int R0, int C0, int R1, int C1);

// An Excel Table (ListObject). The frontend resolves exact colors from StyleName + the theme.
public record TableModel(
    string Name,
    RangeModel Range,
    string? StyleName,
    bool ShowHeaderRow,
    bool ShowTotalsRow,
    bool ShowRowStripes,
    bool ShowColumnStripes,
    bool ShowFirstColumn,
    bool ShowLastColumn);

// One stop of a color scale. Kind is "min"/"max"/"num"/"percent"/"percentile"/"formula";
// Value is the literal for num/percent/percentile. Color is #RRGGBB.
public record ColorScaleStopModel(string Kind, double? Value, string Color);

// A conditional-format rule. Only Type == "colorScale" is emitted for now (discriminated on Type).
public record ConditionalFormatModel(string Type, RangeModel Range, List<ColorScaleStopModel> Stops);

// Resolved workbook theme palette (hex). Lets the frontend render exact table-style colors.
public record ThemePalette(
    string Accent1,
    string Accent2,
    string Accent3,
    string Accent4,
    string Accent5,
    string Accent6,
    string Dk1,
    string Lt1,
    string Dk2,
    string Lt2);
