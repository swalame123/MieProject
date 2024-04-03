using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Authentication.Google;
using IndiaEventsWebApi.Models;
using DinkToPdf.Contracts;
using DinkToPdf;
using Serilog;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.IdentityModel.Tokens;
using System.Text;
using IndiaEventsWebApi.Helper;
using Microsoft.Extensions.DependencyInjection;
using Aspose.Pdf.Plugins;
using iTextSharp.text.pdf.security;
using Smartsheet.Api.Models;
using Microsoft.AspNetCore.Server.Kestrel.Core;

var builder = WebApplication.CreateBuilder(args);
var configuration = builder.Configuration;
int maxRequestLimit = 1073741824;

//builder.Services.AddSession(options => { options.IdleTimeout = TimeSpan.FromMinutes(60); });
builder.Services.AddAuthentication(options =>
{
    options.DefaultScheme = JwtBearerDefaults.AuthenticationScheme;
    options.DefaultChallengeScheme = JwtBearerDefaults.AuthenticationScheme;
})
.AddJwtBearer(options =>
{
    options.SaveToken = true;
    options.RequireHttpsMetadata = false;
    options.TokenValidationParameters = new TokenValidationParameters()
    {
        ValidateIssuer = true,
        ValidateAudience = true,
        ValidAudience = "ABM",
        ValidIssuer = "http://localhost:5098",
        IssuerSigningKey = new SymmetricSecurityKey(Encoding.UTF8.GetBytes("veryveryveryveryverysecret......................"))
    };
   
})

.AddGoogle(options =>
{
    options.ClientId = "644106526561-5899nb8044t0k47h4bdu6lk2aebs4g1s.apps.googleusercontent.com";
    options.ClientSecret = "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";

});

builder.Services.Configure<IISServerOptions>(options =>
{
    options.MaxRequestBodySize = 1073741824; // 1 GB in bytes
    options.MaxRequestBodyBufferSize = 1073741824;
});

builder.Services.Configure<KestrelServerOptions>(options =>
{
    options.Limits.MaxRequestBodySize = 1073741824; // 1 GB in bytes

    options.Limits.RequestHeadersTimeout = TimeSpan.FromMinutes(20); //  20 minutes
});
builder.Services.AddCors(option =>
{
    option.AddPolicy("MyPolicy", builder =>
    {
        builder.AllowAnyOrigin()
        .AllowAnyMethod()
        .AllowAnyHeader();
    });
});
IHostBuilder CreateHostBuilder(string[] args) =>
    Host.CreateDefaultBuilder(args)
    .ConfigureAppConfiguration((hostingContext, config) =>
    {
        config.AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);
    })
.ConfigureServices((hostContext, services) =>
{
    IConfiguration configuration = hostContext.Configuration;
    services.AddSingleton(typeof(IConverter), new SynchronizedConverter(new PdfTools()));
    services.AddOptions();
    services.Configure<SmartsheetSettings>(configuration.GetSection("SmartsheetSettings"));


});

Log.Logger = new LoggerConfiguration()
    .WriteTo.Console()
    .CreateLogger();

builder.Host.UseSerilog((hostingContext, loggerConfiguration) =>
{
    loggerConfiguration.ReadFrom.Configuration(hostingContext.Configuration)
        .Enrich.FromLogContext()
        .WriteTo.Console();
}, preserveStaticLogger: true);



builder.Services.AddControllers();

builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();
builder.Services.AddSingleton(typeof(IConverter), new SynchronizedConverter(new PdfTools()));


var app = builder.Build();


//app.UseCors(options =>
//options.WithOrigins("http://localhost:4200")
//.AllowAnyMethod()
//.AllowAnyHeader()
//);
app.UseSwagger();
app.UseSwaggerUI();
app.UseHttpsRedirection();
app.UseCors("MyPolicy");
app.UseSerilogRequestLogging();
app.UseAuthentication();
app.UseAuthorization();

app.MapControllers();

app.Run();
