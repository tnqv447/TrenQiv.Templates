using NPOI.SS.UserModel;

namespace TrenQiv.Templates.Extensions.NPOI;

public static class IWorkbookExtensions
{
    /// <summary>
    /// Initialize a built-in <see cref="IFontBuilder"/>.
    /// </summary>
    public static IFontBuilder CreateFontBuilder(this IWorkbook wb)
    {
        return new FontBuilder(wb);
    }

    /// <summary>
    /// Initialize a built-in <see cref="ICellStyleBuilder"/>.
    /// </summary>
    public static ICellStyleBuilder CreateCellStyleBuilder(this IWorkbook wb)
    {
        return new CellStyleBuilder(wb);
    }

    /// <summary>
    /// Initialize a built-in <see cref="ICellStyleBuilder"/> with <see cref="CellStyleOptions"/>.
    /// </summary>
    public static ICellStyleBuilder CreateCellStyleBuilder(this IWorkbook wb, CellStyleOptions options)
    {
        return new CellStyleBuilder(wb, options);
    }

    /// <summary>
    /// Create a <see cref="ICellStyle"/> with <see cref="CellStyleOptions"/>.
    /// </summary>
    public static ICellStyle CreateCellStyle(this IWorkbook wb, CellStyleOptions options)
    {
        return new CellStyleBuilder(wb, options).Build();
    }

    /// <summary>
    /// Create a <see cref="IFont"/> with default parameters.
    /// </summary>
    /// <remarks>Let <paramref name="color"/> be <c>null</c> for default black color.</remarks>
    public static IFont CreateFont(this IWorkbook wb,
                                   string name = "Arial",
                                   short fontHeight = 12,
                                   bool bold = false,
                                   bool strikeout = false,
                                   bool italic = false,
                                   short? color = null,
                                   FontSuperScript fontSuperScript = FontSuperScript.None,
                                   FontUnderlineType fontUnderlineType = FontUnderlineType.None)
    {
        return wb.FindFont(bold, color ?? IndexedColors.Black.Index, fontHeight, name, italic, strikeout, fontSuperScript, fontUnderlineType);
    }
}
