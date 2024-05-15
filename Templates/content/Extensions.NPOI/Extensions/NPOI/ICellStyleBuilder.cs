using NPOI.SS.UserModel;

namespace TrenQiv.Templates.Extensions.NPOI;

/// <summary>
/// Base interface for configuring common properties of <see cref="ICellStyle"/>.
/// </summary>
public interface ICellStyleBuilder
{
    /// <summary>
    /// Compiles and returns a <see cref="ICellStyle"/>.
    /// </summary>
    ICellStyle Build();

    /// <summary>
    /// Clones from another <see cref="ICellStyle"/>.
    /// </summary>
    ICellStyleBuilder CloneFrom(ICellStyle otherStyle);

    /// <summary>
    /// Clones from another <see cref="ICellStyleBuilder"/>.
    /// </summary>
    ICellStyleBuilder CloneFrom(ICellStyleBuilder otherStyleBuilder);

    /// <summary>
    /// Apply <see cref="CellStyleOptions"/> to this <see cref="ICellStyleBuilder"/>.
    /// </summary>
    /// <param name="options"></param>
    /// <returns></returns>
    ICellStyleBuilder ApplyOptions(CellStyleOptions options);

    /// <summary>
    /// Enables wrapping text. If <paramref name="enable"/> is <see langword="false"/>, disable instead.
    /// </summary>
    ICellStyleBuilder WrapText(bool enable = true);

    /// <summary>
    /// Set cell's font.
    /// </summary>
    ICellStyleBuilder Font(IFont font);

    /// <summary>
    /// Set cell's foreground color.
    /// </summary>
    ICellStyleBuilder ForegroundColor(IColor color);

    /// <summary>
    /// Set cell's foreground color.
    /// </summary>
    ICellStyleBuilder ForegroundColor(IndexedColors color);

    /// <summary>
    /// Set cell's foreground color by index.
    /// </summary>
    ICellStyleBuilder ForegroundColor(short color);

    /// <summary>
    /// Set cell's background color.
    /// </summary>
    ICellStyleBuilder BackgroundColor(IColor color);

    /// <summary>
    /// Set cell's background color.
    /// </summary>
    ICellStyleBuilder BackgroundColor(IndexedColors color);

    /// <summary>
    /// Set cell's background color by index.
    /// </summary>
    ICellStyleBuilder BackgroundColor(short color);

    /// <summary>
    /// Set cell's fill pattern.
    /// </summary>
    ICellStyleBuilder ApplyFillPattern(FillPattern fillPattern = FillPattern.NoFill);

    /// <summary>
    /// Set cell's border style by <see cref="BorderOptions"/>.
    /// </summary>
    ICellStyleBuilder Border(BorderOptions options);

    /// <summary>
    /// Set cell's border style for all four edges bt <see cref="BorderEdgeOptions"/> .
    /// </summary>
    ICellStyleBuilder Border(BorderEdgeOptions options);

    /// <summary>
    /// Set cell's top border style.
    /// </summary>
    ICellStyleBuilder BorderTop(BorderEdgeOptions options);

    /// <summary>
    /// Set cell's bottom border style.
    /// </summary>
    ICellStyleBuilder BorderBottom(BorderEdgeOptions options);

    /// <summary>
    /// Set cell's left border style.
    /// </summary>
    ICellStyleBuilder BorderLeft(BorderEdgeOptions options);

    /// <summary>
    /// Set cell's right border style.
    /// </summary>
    ICellStyleBuilder BorderRight(BorderEdgeOptions options);

    /// <summary>
    /// Set cell's diagonal border style.
    /// </summary>
    ICellStyleBuilder BorderDiagonal(BorderDiagonalOptions options);

    /// <summary>
    /// Set cell's data format string.
    /// </summary>
    /// <remarks>
    /// <para>Useful for display <see cref="DateTime"/> as string, delimited number,...</para>
    /// <para>Look into <see cref="IDataFormat"/> for <paramref name="formatString"/> syntax.</para>
    /// </remarks>
    /// <param name="formatString">Look into <see cref="IDataFormat"/> for syntax.</param>
    ICellStyleBuilder DataFormat(string formatString);

    /// <summary>
    /// Set cell's horizontal alignment.
    /// </summary>
    ICellStyleBuilder HorizontalAlignment(HorizontalAlignment alignment);

    /// <summary>
    /// Set cell's vertical alignment.
    /// </summary>
    ICellStyleBuilder VerticalAlignment(VerticalAlignment alignment);
}
