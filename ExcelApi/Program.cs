using System.Text.Json;
using ExcelApi.Models;
using ExcelApi.Models.Dto;
using ExcelApi.Models.Helper;
using ExcelApi.Services;
using Syncfusion.Drawing;
using Syncfusion.XlsIO;
using Syncfusion.XlsIORenderer;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
// Learn more about configuring OpenAPI at https://aka.ms/aspnet/openapi
builder.Services.AddOpenApi();

var syncfusionKey = builder.Configuration["SyncfusionKey"];
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


app.UseHttpsRedirection();

// Resolve the workbook path from configuration (defaults to the repo-root DemoData.xlsx,
// one level above the project content root). The file never leaves this backend.
string ResolveExcelPath() =>
    Path.GetFullPath(Path.Combine(
        app.Environment.ContentRootPath,
        app.Configuration["ExcelFile"] ?? "../DemoData.xlsx"));

// Approach A: parse on the backend with ClosedXML and return the shared WorkbookModel as JSON.
app.MapGet("/api/workbook", () =>
    {
        var path = ResolveExcelPath();
        if (!File.Exists(path))
            return Results.NotFound(new { error = $"Excel file not found at {path}" });
        return Results.Ok(WorkbookReader.Read(path));
    })
    .WithName("GetWorkbook")
    .WithTags("Workbook")
    .WithSummary("Parsed workbook as JSON")
    .WithDescription(
        "Parses DemoData.xlsx on the server with ClosedXML and returns the shared WorkbookModel (sheets, cells, styles, merges, pictures).")
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


app.MapPost("/api/excel/read", async (IFormFile file) =>
    {
        if (file.Length == 0)
            return Results.BadRequest("No file uploaded.");

        using var excelEngine = new ExcelEngine();

        var application = excelEngine.Excel;
        application.DefaultVersion = ExcelVersion.Excel2016;

        application.XlsIORenderer = new XlsIORenderer();
        application.XlsIORenderer.ChartRenderingOptions.ImageFormat =
            ExportImageFormat.Png;

        application.XlsIORenderer.ChartRenderingOptions.ScalingMode =
            ScalingMode.Best;

        await using var stream = file.OpenReadStream();

        var workbook =
            application.Workbooks.Open(stream);


        var responsePayload =
            new Dictionary<string, WorksheetDto>(
                workbook.Worksheets.Count);

        foreach (var worksheet in workbook.Worksheets)
        {
            worksheet.EnableSheetCalculations();
            worksheet.UsedRangeIncludesFormatting = true;

            var lastRow = worksheet.UsedRange.LastRow;
            var lastColumn = worksheet.UsedRange.LastColumn;

            #region Grid

            var sheetRows =
                new List<List<CellDto>>(lastRow);

            for (var row = 1; row <= lastRow; row++)
            {
                var rowData =
                    new List<CellDto>(lastColumn);

                for (var col = 1; col <= lastColumn; col++)
                {
                    var cell = worksheet[row, col];

                    var bg =
                        ColorTranslator.ToHtml(
                            cell.CellStyle.Color);

                    var fg =
                        ColorTranslator.ToHtml(
                            cell.CellStyle.Font.RGBColor);

                    if (bg == "#000000" &&
                        cell.CellStyle.FillPattern ==
                        ExcelPattern.None)
                    {
                        bg = "transparent";
                    }

                    if (fg == "#000000")
                        fg = "inherit";

                    #region Style

                    var style = cell.CellStyle;
                    var cellStyle = new CellStyleDto(
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

                    #endregion

                    rowData.Add(new CellDto(
                        cell.DisplayText ?? string.Empty,
                        bg,
                        fg,
                        cellStyle));
                }

                sheetRows.Add(rowData);
            }

            #endregion

            #region Charts (Sequential - safest)

            var sheetCharts =
                new List<ChartDto>(
                    worksheet.Charts.Count);

            foreach (IChart chart in worksheet.Charts)
            {
                using var chartStream = new MemoryStream();

                chart.SaveAsImage(chartStream);

                sheetCharts.Add(new ChartDto(
                    Src:
                    $"data:image/png;base64,{Convert.ToBase64String(chartStream.ToArray())}",
                    X: chart.XPos,
                    Y: chart.YPos,
                    Width: chart.Width,
                    Height: chart.Height));
            }

            #endregion

            #region Images (Safe parallel CPU conversion)

            var pictures =
                worksheet.Pictures
                    .Cast<IPictureShape>()
                    .ToList();

            var imageTasks = pictures.Select(picture =>
                Task.Run(() =>
                {
                    var imageBytes =
                        picture.Picture.ImageData;

                    return new ImageDto(
                        Height: picture.HeightDouble,
                        Width: picture.WidthDouble,
                        Left: picture.LeftDouble,
                        Top: picture.TopDouble,
                        Image: Convert.ToBase64String(imageBytes)
                    );
                }));

            var sheetImages =
                (await Task.WhenAll(imageTasks))
                .ToList();

            #endregion

            #region Merged Cells

            var mergedCells =
                new List<MergeCellDto>();

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

            #endregion


            #region Conditional Formatting
            
            //Evaluate if it is needed.
            //XLSIO can only get rules, So we need to evaluate the rules.
            // https://help.syncfusion.com/document-processing/excel/excel-library/net/working-with-conditional-formatting#reading-an-existing-conditional-format

            #endregion

            responsePayload[worksheet.Name] =
                new WorksheetDto(
                    Grid: sheetRows,
                    Charts: sheetCharts,
                    Images: sheetImages,
                    FrozenRows: worksheet.HorizontalSplit,
                    FrozenColumns: worksheet.VerticalSplit,
                    MergedCells: mergedCells);
        }

        workbook.Close();

        // 1. Serialize the payload object into an in-memory JSON byte array
        var jsonBytes = JsonSerializer.SerializeToUtf8Bytes(responsePayload, new JsonSerializerOptions
        {
            WriteIndented = true // Makes the downloaded file human-readable and clean
        });

        // 2. Return it as a file streaming response
        return Results.File(
            fileContents: jsonBytes,
            contentType: "application/json",
            fileDownloadName: "excel-extracted-data.json"
        );
    })
    .Accepts<IFormFile>("multipart/form-data")
    .DisableAntiforgery()
    .WithName("ReadExcelWorkbook");

app.Run();