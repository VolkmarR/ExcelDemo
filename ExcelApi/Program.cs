using ExcelApi.Services;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
// Learn more about configuring OpenAPI at https://aka.ms/aspnet/openapi
builder.Services.AddOpenApi();

var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.MapOpenApi();
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
    .WithName("GetWorkbook");

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
    .WithName("GetWorkbookFile");

app.Run();
