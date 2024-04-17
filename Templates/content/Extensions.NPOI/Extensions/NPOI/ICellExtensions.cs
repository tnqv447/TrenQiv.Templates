using System.Runtime.CompilerServices;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace TreynQuiv.Templates.Extensions.NPOI;

public static class ICellExtensions
{
    /// <summary>
    /// Fluently set cell value and it's <see cref="CellType"/> based on type of <paramref name="value"/>.
    /// </summary>
    /// <returns>This <see cref="ICell"/>.</returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static ICell SetValue(this ICell cell, dynamic value)
    {
        cell.SetCellValue(value);
        return cell;
    }

    /// <summary>
    /// Fluently set cell formula.
    /// </summary>
    /// <returns>This <see cref="ICell"/>.</returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static ICell SetFormula(this ICell cell, string formula)
    {
        cell.SetCellFormula(formula);
        return cell;
    }

    /// <summary>
    /// Fluently set cell's <see cref="ICellStyle"/>.
    /// <para>If <paramref name="cloneStyle"/> is <see langword="true"/>, set using new instance of <paramref name="cellStyle"/> instead of the same reference.</para>
    /// </summary>
    /// <returns>This <see cref="ICell"/>.</returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static ICell SetStyle(this ICell cell, ICellStyle cellStyle, bool cloneStyle = false)
    {
        cell.CellStyle = cloneStyle ? cell.Row.CreateCellStyleBuilder().CloneFrom(cellStyle).Build() : cellStyle;
        return cell;
    }
}
