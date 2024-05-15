using System.Diagnostics;
using System.Net;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Http.Extensions;
using Microsoft.AspNetCore.Mvc.Filters;

namespace TrenQiv.Templates.Attributes;

public sealed class HttpAutoLogAttribute : ActionFilterAttribute
{
    private readonly Serilog.ILogger _logger = Serilog.Log.ForContext("SourceContext", "HttpAutoLog");
    private const string RequestLogTemplate = "REQ> {0}";
    private const string ResponseLogTemplate = "RES> {0} -- {1} -- {2} ms";
    private const string LogTimerKey = "log_timer";
    private static Stopwatch GetTimer(HttpContext context)
    {
        if (context.Items.TryGetValue(LogTimerKey, out var stopwatch))
        {
            return (Stopwatch)stopwatch!;
        }

        var newStopwatch = new Stopwatch();
        context.Items[LogTimerKey] = newStopwatch;
        return newStopwatch;
    }

    public override void OnActionExecuting(ActionExecutingContext context)
    {
        var request = context.HttpContext.Request;
        var requestDisplay = request.Method + " " + request.GetDisplayUrl();
        _logger.Information(string.Format(RequestLogTemplate, requestDisplay));

        var timer = GetTimer(context.HttpContext);
        timer.Start();
        base.OnActionExecuting(context);
    }

    public override void OnResultExecuted(ResultExecutedContext context)
    {
        var timer = GetTimer(context.HttpContext);
        timer.Stop();

        var request = context.HttpContext.Request;
        var requestDisplay = request.Method + " " + request.GetDisplayUrl();
        if (context.HttpContext.RequestAborted.IsCancellationRequested)
        {
            _logger.Information(string.Format(ResponseLogTemplate, requestDisplay, "ABORTED", timer.ElapsedMilliseconds));
        }
        else
        {
            var response = context.HttpContext.Response;
            var responseStatusCode = (HttpStatusCode)response.StatusCode;
            var responseDisplay = $"{(int)responseStatusCode} {responseStatusCode}";
            _logger.Information(string.Format(ResponseLogTemplate, requestDisplay, responseDisplay, timer.ElapsedMilliseconds));
        }

        base.OnResultExecuted(context);
    }
}
