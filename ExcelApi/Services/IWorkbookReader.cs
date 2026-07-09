using ExcelApi.Models;

namespace ExcelApi.Services;

/// <summary>
/// Reads an .xlsx file into the shared <see cref="WorkbookModel"/>. Implementations must
/// produce interchangeable output so the active parser can be swapped in <c>Program.cs</c>
/// without the frontend noticing — see <see cref="ClosedXmlWorkbookReader"/> (ClosedXML) and
/// <see cref="OpenXmlWorkbookReader"/> (the low-level Open XML SDK).
/// </summary>
public interface IWorkbookReader
{
    WorkbookModel Read(string path);
    WorkbookModel Read(Stream stream);
}
