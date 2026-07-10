using ExcelApi.Models.SyncfusionPoc.Dto;
using Syncfusion.Drawing;
using Syncfusion.XlsIO;

namespace ExcelApi.Models.SyncfusionPoc.Helper;

public static class ExcelHelper
{
    public static BorderDto? GetBorder(IBorder? border)
    {
        if (border == null)
            return null;

        if (border.LineStyle == ExcelLineStyle.None)
            return null;

        return new BorderDto(
            Color: ColorTranslator.ToHtml(border.ColorRGB),
            LineStyle: border.LineStyle.ToString()
        );
    }
}