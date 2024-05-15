using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace TrenQiv.Templates.Extensions;

public static class StringExtensions
{
    /// <summary>
    /// Inline extension method for <c>string.IsNullOrEmpty(value)</c>.
    /// </summary>
    /// <param name="value"></param>
    /// <returns></returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool IsNullOrEmpty([NotNullWhen(false)] this string? value)
    {
        return string.IsNullOrEmpty(value);
    }

    /// <summary>
    /// Inline extension method for <c>!string.IsNullOrEmpty(value)</c>.
    /// </summary>
    /// <param name="value"></param>
    /// <returns></returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool IsNotNullOrEmpty([NotNullWhen(true)] this string? value)
    {
        return !string.IsNullOrEmpty(value);
    }

    /// <summary>
    /// Short-hand for trimming and to upper cases. Use <see cref="CultureInfo.InvariantCulture"/> by defaults.
    /// </summary>
    /// <param name="text"></param>
    /// <returns></returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static string TrimUpper(this string text, CultureInfo? culture = null)
    {
        return text.Trim().ToUpper(culture: culture ?? CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// Short-hand for trimming and to lower cases. Use <see cref="CultureInfo.InvariantCulture"/> by defaults.
    /// </summary>
    /// <param name="text"></param>
    /// <returns></returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static string TrimLower(this string text, CultureInfo? culture = null)
    {
        return text.Trim().ToLower(culture: culture ?? CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// Format the string to <c>camelCase</c>.
    /// </summary>
    /// <remarks>By default, this method will remove the spaces, <see cref="Environment.NewLine"/> and these characters: [<c>_</c>, <c>-</c>, <c>|</c>, <c>#</c>] </remarks>
    /// <param name="text"></param>
    /// <returns></returns>
    public static string ToCamelCase(this string text, params string[] additionalSeparators)
    {
        ArgumentNullException.ThrowIfNull(text);
        Span<string> separators = [" ", "_", "-", "|", "#", Environment.NewLine, .. additionalSeparators];
        var words = text.Split(separators.ToArray(), StringSplitOptions.RemoveEmptyEntries);
        return string.Join(string.Empty, words);
    }

    /// <summary>
    /// Format the string to <c>PascalCase</c>
    /// </summary>
    /// <param name="text"></param>
    /// <returns></returns>
    public static string ToPascalCase(this string text)
    {
        var camelCased = text.ToCamelCase();
        if (camelCased.Length != 0)
        {
            return char.ToUpperInvariant(camelCased[0]) + camelCased[1..];
        }

        return camelCased;
    }

    /// <summary>
    /// Format the string to <c>snake_case</c>.
    /// </summary>
    /// <param name="text"></param>
    /// <returns></returns>
    public static string ToSnakeCase(this string text)
    {
        ArgumentNullException.ThrowIfNull(text);
        if (text.Length < 2)
        {
            return text.ToLowerInvariant();
        }

        var sb = new StringBuilder();
        sb.Append(char.ToLowerInvariant(text[0]));

        for (int i = 1; i < text.Length; ++i)
        {
            char c = text[i];
            if (char.IsUpper(c))
            {
                sb.Append('_');
                sb.Append(char.ToLowerInvariant(c));
            }
            else
            {
                sb.Append(c);
            }
        }

        return sb.ToString();
    }

    /// <summary>
    /// Format the string to <c>kebab-case</c>.
    /// </summary>
    /// <param name="text"></param>
    /// <returns></returns>
    public static string ToKebabCase(this string text)
    {
        ArgumentNullException.ThrowIfNull(text);
        if (text.Length < 2)
        {
            return text.ToLowerInvariant();
        }

        var sb = new StringBuilder();
        sb.Append(char.ToLowerInvariant(text[0]));

        for (int i = 1; i < text.Length; ++i)
        {
            char c = text[i];
            if (char.IsUpper(c))
            {
                sb.Append('-');
                sb.Append(char.ToLowerInvariant(c));
            }
            else
            {
                sb.Append(c);
            }
        }

        return sb.ToString();
    }
}
