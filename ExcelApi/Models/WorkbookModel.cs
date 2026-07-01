namespace ExcelApi.Models;

// Shape mirrors excel-web/src/model.ts so the frontend consumes it directly.
// Minimal APIs serialize with JsonSerializerDefaults.Web => camelCase on the wire.

public record WorkbookModel(List<SheetModel> Sheets);

public record SheetModel(
    string Name,
    int RowCount,
    int ColCount,
    List<CellModel> Cells,
    List<double> ColWidths,
    List<double> RowHeights,
    List<MergeModel> Merges,
    FreezeModel? Freeze);

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
