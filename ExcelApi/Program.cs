using ExcelApi.Models;
using ExcelApi.Services;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
// Learn more about configuring OpenAPI at https://aka.ms/aspnet/openapi
builder.Services.AddOpenApi();

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
    .WithDescription("Parses DemoData.xlsx on the server with ClosedXML and returns the shared WorkbookModel (sheets, cells, styles, merges, pictures).")
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
    .WithDescription("Returns the raw DemoData.xlsx bytes (same-origin) for client-side parsing by the ExcelJS and Univer tabs.")
    .Produces(StatusCodes.Status200OK, contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    .Produces(StatusCodes.Status404NotFound);

app.Run();
