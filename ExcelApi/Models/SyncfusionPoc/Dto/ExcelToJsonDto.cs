namespace ExcelApi.Models.SyncfusionPoc.Dto;

public record BorderDto(
    string Color,
    string LineStyle
);

public sealed record CellStyleDto(
    string Bg,
    string Fg,
    bool Bold,
    bool Italic,
    bool Underline,
    double FontSize,
    string FontName,
    string HorizontalAlignment,
    string VerticalAlignment,
    BorderDto? TopBorder,
    BorderDto? BottomBorder,
    BorderDto? LeftBorder,
    BorderDto? RightBorder
);

public sealed record CellDto(
    string? V,
    int S);

public record ChartDto(
    string Src,
    double X,
    double Y,
    double Width,
    double Height
);

public record ImageDto(
    double Height,
    double Width,
    double Left,
    double Top,
    string Image
);

public record MergeCellDto(
    int StartRow,
    int StartColumn,
    int EndRow,
    int EndColumn
);

public sealed record WorksheetDto(
    List<List<CellDto>> Grid,
    Dictionary<int, CellStyleDto> Styles,
    List<ChartDto> Charts,
    List<ImageDto> Images,
    int FrozenRows,
    int FrozenColumns,
    List<MergeCellDto> MergedCells
);