using System.Runtime.CompilerServices;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace TrenQiv.Templates.Extensions.NPOI;

public static class ISheetExtensions
{
    /// <summary>
    /// Initialize a built-in <see cref="ICellStyleBuilder"/>.
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static ICellStyleBuilder CreateCellStyleBuilder(this ISheet sheet)
    {
        return sheet.Workbook.CreateCellStyleBuilder();
    }

    /// <summary>
    /// Initialize a built-in <see cref="IFontBuilder"/>.
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static IFontBuilder CreateFontBuilder(this ISheet sheet)
    {
        return sheet.Workbook.CreateFontBuilder();
    }

    /// <summary>
    /// Get or create a <see cref="IRow"/> at <paramref name="rowIndex"/>.
    /// </summary>
    public static IRow EnsureRow(this ISheet sheet, int rowIndex, int? rowHeight = null)
    {
        var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
        if (rowHeight.HasValue)
        {
            row.HeightInPoints = rowHeight.Value;
        }

        return row;
    }
}
