using System.Runtime.CompilerServices;
using NPOI.SS.UserModel;
using NPOI.SS;

namespace TrenQiv.Templates.Extensions.NPOI;

/// <summary>
/// Built-in implementation of <see cref="ICellStyleBuilder"/>.
/// </summary>
internal class CellStyleBuilder : ICellStyleBuilder
{
    private readonly IWorkbook _wb;
    private readonly ICellStyle _style;

    public CellStyleBuilder(IWorkbook wb)
    {
        _wb = wb;
        _style = _wb.CreateCellStyle();
    }

    public CellStyleBuilder(IWorkbook wb, CellStyleOptions options)
    {
        _wb = wb;
        _style = _wb.CreateCellStyle();
        ApplyOptions(options);
    }

    public ICellStyle Build()
    {
        var exportStyle = _wb.CreateCellStyle();
        exportStyle.CloneStyleFrom(_style);
        return exportStyle;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public ICellStyleBuilder CloneFrom(ICellStyle otherStyle)
    {
        _style.CloneStyleFrom(otherStyle);
        return this;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public ICellStyleBuilder CloneFrom(ICellStyleBuilder otherStyleBuilder)
    {
        _style.CloneStyleFrom(otherStyleBuilder.Build());
        return this;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public ICellStyleBuilder ApplyOptions(CellStyleOptions options)
    {
        Font(options.Font);
        WrapText(options.WrapText);
        ForegroundColor(options.ForegroundColor);
        BackgroundColor(options.BackgroundColor);
        ApplyFillPattern(options.FillPattern);
        if (options.Border is BorderOptions borderOptions)
        {
            Border(borderOptions);
        }

        return this;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public ICellStyleBuilder Font(IFont font)
    {
        _style.SetFont(font);
        return this;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public ICellStyleBuilder WrapText(bool enable = true)
    {
        _style.WrapText = enable;
        return this;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public ICellStyleBuilder ForegroundColor(IColor color)
    {
        return ForegroundColor(color.Indexed);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public ICellStyleBuilder ForegroundColor(IndexedColors color)
    {
        return ForegroundColor(color.Index);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public ICellStyleBuilder ForegroundColor(short color)
    {
        _style.FillForegroundColor = color;
        return this;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public ICellStyleBuilder BackgroundColor(IColor color)
    {
        return BackgroundColor(color.Indexed);

    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public ICellStyleBuilder BackgroundColor(IndexedColors color)
    {
        return BackgroundColor(color.Index);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public ICellStyleBuilder BackgroundColor(short color)
    {
        _style.FillBackgroundColor = color;
        return this;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public ICellStyleBuilder ApplyFillPattern(FillPattern fillPattern = FillPattern.NoFill)
    {
        _style.FillPattern = fillPattern;
        return this;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public ICellStyleBuilder DataFormat(string formatString)
    {
        _style.DataFormat = _wb.CreateDataFormat().GetFormat(formatString);
        return this;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public ICellStyleBuilder HorizontalAlignment(HorizontalAlignment alignment)
    {
        _style.Alignment = alignment;
        return this;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public ICellStyleBuilder VerticalAlignment(VerticalAlignment alignment)
    {
        _style.VerticalAlignment = alignment;
        return this;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public ICellStyleBuilder Border(BorderOptions options)
    {
        BorderTop(options.Top);
        BorderBottom(options.Bottom);
        BorderLeft(options.Left);
        BorderRight(options.Right);
        BorderDiagonal(options.Diagonal);
        return this;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public ICellStyleBuilder Border(BorderEdgeOptions options)
    {
        BorderTop(options);
        BorderBottom(options);
        BorderLeft(options);
        BorderRight(options);
        return this;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public ICellStyleBuilder BorderTop(BorderEdgeOptions options)
    {
        _style.BorderTop = options.Style;
        _style.TopBorderColor = options.Color;
        return this;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public ICellStyleBuilder BorderBottom(BorderEdgeOptions options)
    {
        _style.BorderBottom = options.Style;
        _style.BottomBorderColor = options.Color;
        return this;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public ICellStyleBuilder BorderLeft(BorderEdgeOptions options)
    {
        _style.BorderLeft = options.Style;
        _style.LeftBorderColor = options.Color;
        return this;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public ICellStyleBuilder BorderRight(BorderEdgeOptions options)
    {
        _style.BorderRight = options.Style;
        _style.RightBorderColor = options.Color;
        return this;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public ICellStyleBuilder BorderDiagonal(BorderDiagonalOptions options)
    {
        _style.BorderDiagonal = options.Diagonal;
        _style.BorderDiagonalLineStyle = options.Style;
        _style.BorderDiagonalColor = options.Color;
        return this;
    }
}
