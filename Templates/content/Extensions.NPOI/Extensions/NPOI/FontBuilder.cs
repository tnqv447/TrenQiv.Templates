using System.Runtime.CompilerServices;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace TrenQiv.Templates.Extensions.NPOI;

/// <summary>
/// Built-in implementation of <see cref="IFontBuilder"/>.
/// </summary>
public class FontBuilder : IFontBuilder
{
    protected readonly IWorkbook _wb;
    protected readonly IFont _font;
    public FontBuilder(IWorkbook wb, string fontName = "Arial", int fontHeightInPoints = 12)
    {
        _wb = wb;
        _font = _wb.CreateFont();
        _font.FontName = fontName;
        _font.FontHeightInPoints = fontHeightInPoints;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public IFont Build()
    {
        var exportFont = _wb.CreateFont();
        exportFont.CloneStyleFrom(_font);
        return exportFont;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public IFontBuilder CloneFrom(IFont font)
    {
        _font.CloneStyleFrom(font);
        return this;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public IFontBuilder CloneFrom(IFontBuilder otherFontBuilder)
    {
        _font.CloneStyleFrom(otherFontBuilder.Build());
        return this;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public IFontBuilder FontName(string fontName)
    {
        _font.FontName = fontName;
        return this;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public IFontBuilder Bold(bool enable = true)
    {
        _font.IsBold = enable;
        return this;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public IFontBuilder Italic(bool enable = true)
    {
        _font.IsItalic = enable;
        return this;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public IFontBuilder Strikeout(bool enable = true)
    {
        _font.IsStrikeout = enable;
        return this;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public IFontBuilder Color(IColor color)
    {
        return Color(color.Indexed);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public IFontBuilder Color(IndexedColors color)
    {
        return Color(color.Index);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public IFontBuilder Color(short color)
    {
        _font.Color = color;
        return this;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public IFontBuilder HeightInPoints(short value)
    {
        _font.FontHeightInPoints = value;
        return this;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public IFontBuilder Underline(FontUnderlineType underlineType)
    {
        _font.Underline = underlineType;
        return this;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public IFontBuilder TypeOffset(FontSuperScript fontSuperScript)
    {
        _font.TypeOffset = fontSuperScript;
        return this;
    }
}
