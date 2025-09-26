using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using ContractGeneratorBlazor.Data;
using ContractGeneratorBlazor.Models;
using QuestPDF.Infrastructure;

var builder = WebApplication.CreateBuilder(args);
QuestPDF.Settings.License = LicenseType.Community;

// Bind config
builder.Services.Configure<ContractConfig>(
    builder.Configuration.GetSection("ContractSettings"));

// Add services to the container.
builder.Services.AddRazorPages();
builder.Services.AddServerSideBlazor();
builder.Services.AddSingleton<WeatherForecastService>();
builder.Services.AddScoped<ContractGeneratorBlazor.Services.IDocumentGenerator, ContractGeneratorBlazor.Services.DocumentGenerator>();
builder.Services.AddScoped<ContractGeneratorBlazor.Services.IContractService, ContractGeneratorBlazor.Services.ContractService>();

var app = builder.Build();

// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error");
    // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
    app.UseHsts();
}

app.UseHttpsRedirection();

app.UseStaticFiles();

app.UseRouting();

app.MapBlazorHub();
app.MapFallbackToPage("/_Host");

app.Run();
