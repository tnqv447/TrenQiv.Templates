using NPOI.SS.UserModel;

namespace TreynQuiv.Templates.Extensions.NPOI;

public static class IWorkbookExtensions
{
    /// <summary>
    /// Initialize a built-in <see cref="ICellStyleBuilder"/>.
    /// </summary>
    public static ICellStyleBuilder CreateCellStyleBuilder(this IWorkbook wb)
    {
        return new CellStyleBuilder(wb);
    }

    /// <summary>
    /// Initialize a built-in <see cref="IFontBuilder"/>.
    /// </summary>
    public static IFontBuilder CreateFontBuilder(this IWorkbook wb)
    {
        return new FontBuilder(wb);
    }
}
