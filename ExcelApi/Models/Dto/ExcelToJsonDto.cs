namespace ExcelApi.Models.Dto;

public record BorderDto(
    string Color,
    string LineStyle
);

public record CellStyleDto(
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

public record CellDto(
    string Text,
    string Bg,
    string Fg,
    CellStyleDto Style
);

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

public record WorksheetDto(
    List<List<CellDto>> Grid,
    List<ChartDto> Charts,
    List<ImageDto> Images,
    int FrozenRows,
    int FrozenColumns,
    List<MergeCellDto> MergedCells
);