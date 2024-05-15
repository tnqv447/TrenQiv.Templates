using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using IComponent = Microsoft.AspNetCore.Components.IComponent;

namespace TreynQuiv.Templates.Helpers;

internal class BlazorRenderer : IAsyncDisposable
{
    private readonly ServiceProvider _serviceProvider;
    private readonly ILoggerFactory _loggerFactory;
    private readonly HtmlRenderer _htmlRenderer;
    public BlazorRenderer()
    {
        // Build all the dependencies for the HtmlRenderer
        var services = new ServiceCollection();
        services.AddLogging();
        _serviceProvider = services.BuildServiceProvider();
        _loggerFactory = _serviceProvider.GetRequiredService<ILoggerFactory>();
        _htmlRenderer = new HtmlRenderer(_serviceProvider, _loggerFactory);
    }

    // Dispose the services and DI container we created
    public async ValueTask DisposeAsync()
    {
        await _htmlRenderer.DisposeAsync();
        _loggerFactory.Dispose();
        await _serviceProvider.DisposeAsync();
    }

    // The other public methods are identical
    public Task<string> RenderComponentAsync<T>() where T : IComponent
    {
        return RenderComponentAsync<T>(ParameterView.Empty);
    }

    public Task<string> RenderComponentAsync<T>(Dictionary<string, object?> dictionary) where T : IComponent
    {
        return RenderComponentAsync<T>(ParameterView.FromDictionary(dictionary));
    }

    private Task<string> RenderComponentAsync<T>(ParameterView parameters) where T : IComponent
    {
        return _htmlRenderer.Dispatcher.InvokeAsync(async () =>
        {
            var output = await _htmlRenderer.RenderComponentAsync<T>(parameters);
            return output.ToHtmlString();
        });
    }
}
