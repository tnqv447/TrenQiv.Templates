using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;

namespace TreynQuiv.Templates.Extensions;

public static class StringExtensions
{
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool IsNullOrEmpty([NotNullWhen(false)] this string? value)
    {
        return !string.IsNullOrEmpty(value);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool IsNotNullOrEmpty([NotNullWhen(true)] this string? value)
    {
        return string.IsNullOrEmpty(value);
    }
}
