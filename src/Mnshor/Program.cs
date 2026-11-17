using Microsoft.AspNetCore.SignalR;
using Microsoft.AspNetCore.StaticFiles;
using Mnshor.Components;
using Mnshor.Engines;
using Mnshor.Services;

var builder = WebApplication.CreateBuilder(new WebApplicationOptions
{
    Args = args,
    ContentRootPath = AppContext.BaseDirectory
});

builder.Services.AddWindowsService(options =>
{
    options.ServiceName = "MnshorService";
});

// Add services to the container.
builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents();

// Configure SignalR buffer sizes for handling large documents in-memory
builder.Services.Configure<HubOptions>(options =>
{
    options.MaximumReceiveMessageSize = 500 * 1024 * 1024; // 500 MB
});

// High-Performance Services Registration
builder.Services.AddLocalization();
builder.Services.AddScoped<LocalizationState>();
builder.Services.AddSingleton<IPdfEngine, FastPdfEngine>();
builder.Services.AddSingleton<IWordConverter, LocalWordConverter>();
builder.Services.AddSingleton<ISpreadsheetMarkdownConverter, LocalSpreadsheetMarkdownConverter>();
builder.Services.AddSingleton<OsFontService>();

PdfSharp.Fonts.GlobalFontSettings.FontResolver ??= new SystemFontResolver();

var app = builder.Build();

var localizationOptions = new RequestLocalizationOptions()
    .SetDefaultCulture(LocalizationState.DefaultCulture)
    .AddSupportedCultures(LocalizationState.SupportedCultures)
    .AddSupportedUICultures(LocalizationState.SupportedCultures);

app.UseRequestLocalization(localizationOptions);

var configuredUrls = builder.Configuration["Urls"] ?? builder.Configuration["ASPNETCORE_URLS"] ?? string.Empty;
var hasHttpsEndpoint = configuredUrls.Contains("https://", StringComparison.OrdinalIgnoreCase);

// Configure HTTP pipeline
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error", createScopeForErrors: true);
    if (hasHttpsEndpoint)
    {
        app.UseHsts();
    }
}

if (hasHttpsEndpoint)
{
    app.UseHttpsRedirection();
}
var contentTypes = new FileExtensionContentTypeProvider();
contentTypes.Mappings[".mjs"] = "text/javascript";
contentTypes.Mappings[".wasm"] = "application/wasm";
app.UseStaticFiles(new StaticFileOptions
{
    ContentTypeProvider = contentTypes
});
app.UseAntiforgery();

app.MapRazorComponents<App>()
    .AddInteractiveServerRenderMode();

app.Run();
