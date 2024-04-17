using System.Drawing.Imaging;
using System.Runtime.CompilerServices;
using System.Text.Encodings.Web;
using System.Text.Json;

namespace TreynQuiv.Templates.Extensions;

public static class JsonExtensions
{
    /// <summary>
    /// Get the JSON representation of this object with options
    /// </summary>
    /// <param name="obj"></param>
    /// <param name="options"></param>
    /// <returns>A JSON string</returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static string ToJsonString(this object obj, JsonSerializerOptions? options)
    {
        dynamic @this = Convert.ChangeType(obj, obj.GetType());
        return JsonSerializer.Serialize(@this, options);
    }

    /// <summary>
    /// Get the JSON representation of this object.
    /// <para><paramref name="writeIndented"/> settings overrides</para>
    /// </summary>
    /// <param name="obj"></param>
    /// <param name="writeIndented"></param>
    /// <returns>A JSON string</returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static string ToJsonString(this object obj, bool writeIndented = false)
    {
        dynamic @this = Convert.ChangeType(obj, obj.GetType());
        return JsonSerializer.Serialize(@this, new JsonSerializerOptions()
        {
            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
            WriteIndented = writeIndented
        });
    }

    /// <summary>
    /// Deserialize this string to <typeparamref name="T"/> object.
    /// </summary>
    /// <param name="jsonString"></param>
    /// <param name="options"></param>
    /// <returns>Return <typeparamref name="T"/> if successful, otherwise <see langword="null"/></returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static T? DeserializeJsonString<T>(this string jsonString, JsonSerializerOptions? options)
    {
        return JsonSerializer.Deserialize<T>(jsonString, options);
    }
}
