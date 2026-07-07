namespace ExcelApi.Models.Dto;

public record CellDto(
    string Text,
    string Bg,
    string Fg
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