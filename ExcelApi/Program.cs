using System.Text.Json;
using System.Text.Json.Serialization;
using ExcelApi.Models;
using ExcelApi.Models.SyncfusionPoc.Dto;
using ExcelApi.Models.SyncfusionPoc.Helper;
using ExcelApi.Services;
using Microsoft.AspNetCore.Http.Features;
using Microsoft.AspNetCore.Mvc;
using Syncfusion.Drawing;
using Syncfusion.EJ2.Spreadsheet;
using Syncfusion.XlsIO;
using Syncfusion.XlsIORenderer;
using Syncfusion.EJ2.Spreadsheet;
using ExcelEngine = Syncfusion.XlsIO.ExcelEngine;
using FormFile = Microsoft.AspNetCore.Http.FormFile;
using MemoryStream = System.IO.MemoryStream;
using Results = Microsoft.AspNetCore.Http.Results;
using Stream = System.IO.Stream;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
// Learn more about configuring OpenAPI at https://aka.ms/aspnet/openapi

builder.Services.Configure<FormOptions>(options =>
{
    options.MultipartBodyLengthLimit = int.MaxValue;
    options.ValueLengthLimit = int.MaxValue;
});

builder.WebHost.ConfigureKestrel(options => { options.Limits.MaxRequestBodySize = int.MaxValue; });

builder.Services.AddCors(options =>
{
    options.AddPolicy("React",
        policy => { policy.WithOrigins("http://localhost:5173").AllowAnyHeader().AllowAnyMethod(); });
});

builder.Services.AddOpenApi();

var syncfusionKey = builder.Configuration["SyncfusionKey"];
// The active backend workbook parser. Both implementations produce the same WorkbookModel, so
// swapping this single line switches /api/workbook between ClosedXML and the low-level Open XML SDK.
builder.Services.AddSingleton<IWorkbookReader, ClosedXmlWorkbookReader>();
// builder.Services.AddSingleton<IWorkbookReader, OpenXmlWorkbookReader>();

var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    // Serve the OpenAPI (Swagger) document at /openapi/v1.json …
    app.MapOpenApi();

    // …and render an interactive Swagger UI at /swagger that reads that document.
    app.UseSwaggerUI(options =>
    {
        options.SwaggerEndpoint("/openapi/v1.json", "ExcelApi v1");
        options.DocumentTitle = "ExcelApi — Swagger UI";
    });
}

Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense(syncfusionKey);

app.UseCors("React");

app.UseHttpsRedirection();

// Resolve the workbook path from configuration (defaults to the repo-root DemoData.xlsx,
// one level above the project content root). The file never leaves this backend.
string ResolveExcelPath() =>
    Path.GetFullPath(Path.Combine(
        app.Environment.ContentRootPath,
        app.Configuration["ExcelFile"] ?? "../DemoData.xlsx"));

// Approach A: parse on the backend with the configured IWorkbookReader and return the shared
// WorkbookModel as JSON (ClosedXML or the Open XML SDK — see the DI registration above).
app.MapGet("/api/workbook", (IWorkbookReader reader) =>
    {
        var path = ResolveExcelPath();
        if (!File.Exists(path))
            return Results.NotFound(new { error = $"Excel file not found at {path}" });
        return Results.Ok(reader.Read(path));
    })
    .WithName("GetWorkbook")
    .WithTags("Workbook")
    .WithSummary("Parsed workbook as JSON")
    .WithDescription(
        "Parses DemoData.xlsx on the server with the configured IWorkbookReader and returns the shared WorkbookModel (sheets, cells, styles, merges, pictures).")
    .Produces<WorkbookModel>(StatusCodes.Status200OK)
    .Produces(StatusCodes.Status404NotFound);

// Approaches B & C: serve the raw .xlsx bytes for client-side parsing (same-origin, never external).
app.MapGet("/api/workbook/file", () =>
    {
        var path = ResolveExcelPath();
        if (!File.Exists(path))
            return Results.NotFound();
        var bytes = File.ReadAllBytes(path);
        return Results.File(
            bytes,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            Path.GetFileName(path));
    })
    .WithName("GetWorkbookFile")
    .WithTags("Workbook")
    .WithSummary("Raw .xlsx file")
    .WithDescription(
        "Returns the raw DemoData.xlsx bytes (same-origin) for client-side parsing by the ExcelJS and Univer tabs.")
    .Produces(StatusCodes.Status200OK, contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    .Produces(StatusCodes.Status404NotFound);


#region Syncfusion POC

app.MapPost("/api/excel/read/{sheetName}", async (
        string sheetName,
        IFormFile file) =>
    {
        if (file.Length == 0)
            return Results.BadRequest("No file uploaded.");

        using var excelEngine = new ExcelEngine();

        var application = excelEngine.Excel;
        application.DefaultVersion = ExcelVersion.Excel2016;

        application.XlsIORenderer = new XlsIORenderer();
        application.XlsIORenderer.ChartRenderingOptions.ImageFormat = ExportImageFormat.Png;

        application.XlsIORenderer.ChartRenderingOptions.ScalingMode = ScalingMode.Best;

        await using var stream = file.OpenReadStream();

        var workbook = application.Workbooks.Open(stream);

        var worksheet = workbook.Worksheets
            .FirstOrDefault(w =>
                string.Equals(
                    w.Name,
                    sheetName,
                    StringComparison.OrdinalIgnoreCase));

        if (worksheet is null)
        {
            return Results.NotFound(
                $"Worksheet '{sheetName}' not found.");
        }

        worksheet.EnableSheetCalculations();
        worksheet.UsedRangeIncludesFormatting = true;

        var lastRow = worksheet.UsedRange.LastRow;
        var lastColumn = worksheet.UsedRange.LastColumn;

        // ======================
        // Grid
        // ======================

        var styleMap = new Dictionary<CellStyleDto, int>();
        var styles = new Dictionary<int, CellStyleDto>();

        var nextStyleId = 1;

        var sheetRows = new List<List<CellDto>>(lastRow);

        for (var row = 1; row <= lastRow; row++)
        {
            var rowData = new List<CellDto>(lastColumn);

            for (var col = 1; col <= lastColumn; col++)
            {
                var cell = worksheet[row, col];

                var bg = ColorTranslator.ToHtml(
                    cell.CellStyle.Color);

                var fg = ColorTranslator.ToHtml(
                    cell.CellStyle.Font.RGBColor);

                if (bg == "#000000" &&
                    cell.CellStyle.FillPattern == ExcelPattern.None)
                {
                    bg = "transparent";
                }

                if (fg == "#000000")
                    fg = "inherit";

                var style = cell.CellStyle;

                var styleDto = new CellStyleDto(
                    Bg: bg,
                    Fg: fg,
                    Bold: style.Font.Bold,
                    Italic: style.Font.Italic,
                    Underline: style.Font.Underline != ExcelUnderline.None,
                    FontSize: style.Font.Size,
                    FontName: style.Font.FontName,
                    HorizontalAlignment: style.HorizontalAlignment.ToString(),
                    VerticalAlignment: style.VerticalAlignment.ToString(),
                    TopBorder: ExcelHelper.GetBorder(style.Borders[ExcelBordersIndex.EdgeTop]),
                    BottomBorder: ExcelHelper.GetBorder(style.Borders[ExcelBordersIndex.EdgeBottom]),
                    LeftBorder: ExcelHelper.GetBorder(style.Borders[ExcelBordersIndex.EdgeLeft]),
                    RightBorder: ExcelHelper.GetBorder(style.Borders[ExcelBordersIndex.EdgeRight])
                );

                if (!styleMap.TryGetValue(styleDto, out var styleId))
                {
                    styleId = nextStyleId++;

                    styleMap.Add(styleDto, styleId);
                    styles.Add(styleId, styleDto);
                }

                rowData.Add(
                    new CellDto(
                        V: cell.DisplayText,
                        S: styleId
                    ));
            }

            sheetRows.Add(rowData);
        }

        // ======================
        // Charts
        // ======================

        var sheetCharts = new List<ChartDto>();

        foreach (IChart chart in worksheet.Charts)
        {
            using var chartStream = new MemoryStream();

            chart.SaveAsImage(chartStream);

            sheetCharts.Add(new ChartDto(
                Src: $"data:image/png;base64,{Convert.ToBase64String(chartStream.ToArray())}",
                X: chart.XPos,
                Y: chart.YPos,
                Width: chart.Width,
                Height: chart.Height));
        }

        // ======================
        // Images
        // ======================

        var sheetImages = new List<ImageDto>();

        foreach (IPictureShape picture in worksheet.Pictures)
        {
            sheetImages.Add(
                new ImageDto(
                    Height: picture.HeightDouble,
                    Width: picture.WidthDouble,
                    Left: picture.LeftDouble,
                    Top: picture.TopDouble,
                    Image: Convert.ToBase64String(
                        picture.Picture.ImageData)
                ));
        }

        // ======================
        // Merged Cells
        // ======================

        var mergedCells = new List<MergeCellDto>();

        if (worksheet.MergedCells is not null)
        {
            foreach (var area in worksheet.MergedCells)
            {
                if (area is null)
                    continue;

                mergedCells.Add(
                    new MergeCellDto(
                        StartRow: area.Row,
                        StartColumn: area.Column,
                        EndRow: area.LastRow,
                        EndColumn: area.LastColumn));
            }
        }

        var response = new WorksheetDto(
            Grid: sheetRows,
            Styles: styles,
            Charts: sheetCharts,
            Images: sheetImages,
            FrozenRows: worksheet.HorizontalSplit,
            FrozenColumns: worksheet.VerticalSplit,
            MergedCells: mergedCells
        );

        var jsonOptions = new JsonSerializerOptions
        {
            WriteIndented = false,
            DefaultIgnoreCondition =
                JsonIgnoreCondition.WhenWritingDefault
        };

        var jsonBytes = JsonSerializer.SerializeToUtf8Bytes(
            response,
            jsonOptions);

        // 2. Return it as a file streaming response
        return Results.File(
            fileContents: jsonBytes,
            contentType: "application/json",
            fileDownloadName: "excel-extracted-data.json"
        );
    })
    .Accepts<IFormFile>("multipart/form-data")
    .DisableAntiforgery()
    .WithName("ReadWorksheet");

app.MapPost("/api/excel/sheets", async (IFormFile file) =>
    {
        if (file.Length == 0)
            return Results.BadRequest("No file uploaded.");

        using var excelEngine = new ExcelEngine();

        var application = excelEngine.Excel;
        application.DefaultVersion = ExcelVersion.Excel2016;

        await using var stream = file.OpenReadStream();

        var workbook = application.Workbooks.Open(stream);

        var sheetNames = workbook.Worksheets
            .Select(w => w.Name)
            .ToList();

        return Results.Json(sheetNames);
    })
    .Accepts<IFormFile>("multipart/form-data")
    .DisableAntiforgery()
    .WithName("GetWorkbookSheets");

#endregion


#region Syncfusion POC with React Viewer

// Open action

// app.MapPost("/api/spreadsheet/open", async (HttpRequest request) =>
// {
//     var form = await request.ReadFormAsync();
//
//     if (form.Files.Count == 0)
//     {
//         return Results.BadRequest("No file uploaded.");
//     }
//
//     var openRequest = new OpenRequest
//     {
//         File = form.Files[0]
//     };
//
//     var result = Workbook.Open(openRequest);
//
//
//     return Results.Content(result, "application/json");
// });


app.MapPost("/api/spreadsheet/open", async (HttpRequest request) =>
{
    var form = await request.ReadFormAsync();

    if (form.Files.Count == 0)
        return Results.BadRequest();

    using var excelEngine = new ExcelEngine();
    var appExcel = excelEngine.Excel;

    using Stream stream = form.Files[0].OpenReadStream();

    var workbook = appExcel.Workbooks.Open(stream);
    foreach (var sheet in workbook.Worksheets)
    {
        if (sheet is null)
            continue;

        sheet.EnableSheetCalculations();

        foreach (var cell in sheet.UsedRange.Cells)
        {
            if (cell is not null && !string.IsNullOrEmpty(cell.Formula))
            {
                var value = cell.CalculatedValue;

                // Remove formula
                cell.Clear(ExcelClearOptions.ClearContent);

                // Write calculated value as text/value
                cell.Value = value;
            }
        }

        //In case of table we can filter the table data
        foreach (var table in sheet.ListObjects)
        {
            sheet.AutoFilters.FilterRange = table.Location;
        }
    }

    // Save workbook to memory
    using var memoryStream = new MemoryStream();
    workbook.SaveAs(memoryStream);
    memoryStream.Position = 0;

    var openRequest = new OpenRequest
    {
        File = new FormFile(
            memoryStream,
            0,
            memoryStream.Length,
            "file",
            form.Files[0].FileName)
    };

    var result = Workbook.Open(openRequest);

    return Results.Content(result, "application/json");
});

#endregion


app.Run();