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

var builder = WebApplication.CreateBuilder(args);
var configuration = builder.Configuration;



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

    //options.Events = new JwtBearerEvents
    //{
    //    OnAuthenticationFailed = context =>
    //    {
    //        context.Response.ContentType = "application/json";
    //        context.Response.StatusCode = StatusCodes.Status401Unauthorized;

    //        //return context.Response.WriteAsync(new
    //        //{
    //        //    Message = "Invalid token",
    //        //    Error = context.Exception.Message
    //        //}.ToString());
            
    //    }
    //};
})
.AddGoogle(options =>
{
    options.ClientId = "200698853522-5b3nkgrgal38n7eqjqrrt6biinbt46ca.apps.googleusercontent.com";
    options.ClientSecret = "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";
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
