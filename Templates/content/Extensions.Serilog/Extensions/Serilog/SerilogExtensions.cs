using System.Runtime.CompilerServices;
using Serilog;
using Serilog.Core;
using Serilog.Events;

namespace TreynQuiv.Templates.Extensions.Serilog;

public static class SerilogSettings
{
    public const string LogTemplate =
    "[{Timestamp:dd/MM/yyyy HH:mm:ss}][{Level:u3}] <{ThreadId}> [{SourceContextName}{MemberName}]{Executioner}{SubContext} {Message:lj}{NewLine}{Exception}";
    public static void Initialize()
    {
        Log.Logger = new LoggerConfiguration()
            .MinimumLevel.Debug()
            .Enrich.FromLogContext()
            .Enrich.With(new ThreadIdEnricher())
            .Enrich.WithComputed("SourceContextName", "Substring(SourceContext, LastIndexOf(SourceContext, '.') + 1)")
            .WriteTo.Console(LogEventLevel.Debug, LogTemplate)
            .WriteTo.Map(
                ev => $"{DateTime.Now:dd-MM-yyyy}",
                (date, wt) => wt.File($"./logs/{date}.log",
                                      LogEventLevel.Information,
                                      LogTemplate,
                                      fileSizeLimitBytes: 10000000,
                                      rollOnFileSizeLimit: true))
            .CreateLogger()
            .ForContext("Executioner", $" (SYSTEM)");
    }

    public static ILogger Here(this ILogger logger, [CallerMemberName] string memberName = "")
    {
        return logger.ForContext("MemberName", $"{(string.IsNullOrEmpty(memberName) ? "" : ("." + memberName))}");
    }

    public static ILogger ExecuteBy(this ILogger logger, string executioner)
    {
        return logger.ForContext("Executioner", $" ({executioner})");
    }

    public static ILogger SubContext<T>(this ILogger logger)
    {
        return logger.ForContext("SubContext", $" {typeof(T).Name} |");
    }

    public static ILogger SubContextHere<T>(this ILogger logger, [CallerMemberName] string memberName = "")
    {
        return logger.ForContext("SubContext", $" {typeof(T).Name}{(string.IsNullOrEmpty(memberName) ? "" : ("." + memberName))} |");
    }

    /// <summary>
    /// Default log error with message: <c>An error occurred:</c>.
    /// </summary>
    public static void Error(this ILogger logger, Exception ex)
    {
        logger.Error(ex, "An error occurred:");
    }
}

internal class ThreadIdEnricher : ILogEventEnricher
{
    public void Enrich(LogEvent logEvent, ILogEventPropertyFactory propertyFactory)
    {
        logEvent.AddPropertyIfAbsent(propertyFactory.CreateProperty(
            "ThreadId", Environment.CurrentManagedThreadId.ToString("D3")));
    }
}
