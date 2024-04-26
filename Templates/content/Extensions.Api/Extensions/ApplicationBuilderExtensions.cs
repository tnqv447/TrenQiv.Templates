using System.Text;
using Microsoft.AspNetCore.Builder;

namespace TreynQuiv.Templates.Extensions;

public static class ApplicationBuilderExtensions
{
    /// <summary>
    /// A middleware to set X-Content-Type-Options header to <see cref="HttpResponse"/>. Default: <c>NOSNIF</c>.
    /// </summary>
    public static IApplicationBuilder UseXContentTypeOptions(this IApplicationBuilder builder, string headerValue = "NOSNIFF")
    {
        headerValue = SanitizeHeaderValue(headerValue);
        return builder.Use(async (context, next) =>
        {
            context.Response.Headers.XContentTypeOptions = headerValue;
            await next();
        });
    }

    /// <summary>
    /// A middleware to set X-Frame-Options header to <see cref="HttpResponse"/>. Default: <c>SAMEORIGIN</c>.
    /// </summary>
    public static IApplicationBuilder UseXFrameOptions(this IApplicationBuilder builder, string headerValue = "SAMEORIGIN")
    {
        headerValue = SanitizeHeaderValue(headerValue);
        return builder.Use(async (context, next) =>
        {
            context.Response.Headers.XFrameOptions = headerValue;
            await next();
        });
    }

    /// <summary>
    /// A middleware to set Content-Security-Policy header to <see cref="HttpResponse"/>. Default: <c>default-src 'self'; script-src 'self';</c>.
    /// </summary>
    public static IApplicationBuilder UseContentSecurityPolicy(this IApplicationBuilder builder, string headerValue = "default-src 'self'; script-src 'self';")
    {
        headerValue = SanitizeHeaderValue(headerValue);
        return builder.Use(async (context, next) =>
        {
            context.Response.Headers.ContentSecurityPolicy = headerValue;
            await next();
        });
    }

    /// <summary>
    /// Remove invalid chars from HTTP header value and encode to ASCII.
    /// </summary>
    /// <param name="headerValue"></param>
    /// <returns>An ASCII string safe for HTTP header value</returns>
    private static string SanitizeHeaderValue(string headerValue)
    {
        headerValue = headerValue.ReplaceLineEndings().Replace(Environment.NewLine, " ");
        headerValue = Encoding.ASCII.GetString(Encoding.UTF8.GetBytes(headerValue));
        return headerValue;
    }
}
