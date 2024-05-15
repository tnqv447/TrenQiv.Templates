using NPOI.SS.UserModel;

namespace TrenQiv.Templates.Extensions.NPOI;

public record class CellStyleOptions
{
    public required IFont Font { get; set; }
    public bool WrapText { get; set; }

    /// <summary>
    /// Use <see cref="IndexedColor.Black"/> by <c>MS-Excel</c> default if not set.
    /// </summary>
    public short ForegroundColor { get; set; } = IndexedColors.Black.Index;

    /// <summary>
    /// Use <see cref="IndexedColor.White"/> by <c>MS-Excel</c> default if not set.
    /// </summary>
    public short BackgroundColor { get; set; } = IndexedColors.White.Index;

    /// <summary>
    /// Use <see cref="FillPattern.NoFill"/> by <c>MS-Excel</c> default if not set.
    /// </summary>
    public FillPattern FillPattern { get; set; } = FillPattern.NoFill;

    /// <summary>
    /// Use <see cref="HorizontalAlignment.General"/> by <c>MS-Excel</c> default if not set.
    /// </summary>
    public HorizontalAlignment HorizontalAlignment { get; set; } = HorizontalAlignment.General;

    /// <summary>
    /// Use <see cref="VerticalAlignment.None"/> by <c>MS-Excel</c> default if not set.
    /// </summary>
    public VerticalAlignment VerticalAlignment { get; set; } = VerticalAlignment.None;
    public BorderOptions? Border { get; set; }
}

public record struct BorderOptions
{
    public BorderOptions(BorderStyle style, short color)
    {
        Top = new(style, color);
        Bottom = new(style, color);
        Left = new(style, color);
        Right = new(style, color);
        Diagonal = new(BorderDiagonal.None, style, color);
    }

    public BorderOptions(BorderStyle style, IColor color) : this(style, color.Indexed) { }
    public BorderOptions(BorderStyle style, IndexedColors color) : this(style, color.Index) { }

    public BorderEdgeOptions Top { get; set; }
    public BorderEdgeOptions Right { get; set; }
    public BorderEdgeOptions Bottom { get; set; }
    public BorderEdgeOptions Left { get; set; }
    public BorderDiagonalOptions Diagonal { get; set; }
}

public record struct BorderEdgeOptions(BorderStyle Style, short Color)
{
    public BorderEdgeOptions(BorderStyle style, IColor color) : this(style, color.Indexed) { }
    public BorderEdgeOptions(BorderStyle style, IndexedColors color) : this(style, color.Index) { }
}

public record struct BorderDiagonalOptions(BorderDiagonal Diagonal, BorderStyle Style, short Color)
{
    public BorderDiagonalOptions(BorderDiagonal diagonal, BorderStyle style, IColor color) : this(diagonal, style, color.Indexed) { }
    public BorderDiagonalOptions(BorderDiagonal diagonal, BorderStyle style, IndexedColors color) : this(diagonal, style, color.Index) { }
}
