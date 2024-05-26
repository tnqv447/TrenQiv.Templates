using System.Runtime.CompilerServices;
using NPOI.SS.UserModel;

namespace TrenQiv.Templates.Extensions.NPOI;

public static class IRowExtensions
{
    /// <summary>
    /// Initialize a built-in <see cref="ICellStyleBuilder"/>.
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static ICellStyleBuilder CreateCellStyleBuilder(this IRow row)
    {
        return row.Sheet.CreateCellStyleBuilder();
    }

    /// <summary>
    /// Initialize a built-in <see cref="IFontBuilder"/>.
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static IFontBuilder CreateFontBuilder(this IRow row)
    {
        return row.Sheet.CreateFontBuilder();
    }

    /// <summary>
    /// Get or create a cell with the enforced <see cref="CellType"/>.
    /// </summary>
    public static ICell EnsureCell(this IRow row, int columnIndex, CellType cellType = CellType.Blank)
    {
        var cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);
        cell.SetCellType(cellType);
        return cell;
    }

    /// <summary>
    /// Create and set cells starts from <paramref name="startColIndex"/> in sequence to <paramref name="values"/> with <paramref name="cellStyle"/> if specified.
    /// </summary>
    public static void SetRowValues(this IRow row, int startColIndex = 0, ICellStyle? cellStyle = null, params dynamic[] values)
    {
        var valuesNum = values.Length;
        for (var i = 0; i < valuesNum; i++)
        {
            var cell = row.EnsureCell(startColIndex + i);
            if (cellStyle is not null)
            {
                cell.CellStyle = cellStyle;
            }

            cell.SetCellValue(values[i]);
        }
    }
}
