using BackEnd.Services;
using Microsoft.AspNetCore.Server.IIS;
using Microsoft.AspNetCore.Http.Features;
using Microsoft.OpenApi.Models;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
// Learn more about configuring OpenAPI at https://aka.ms/aspnet/openapi
builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen(c =>
{
    c.SwaggerDoc("v1", new OpenApiInfo { 
        Title = "Excel Data Import API", 
        Version = "v1",
        Description = "API for importing Excel data into SQL Server" 
    });
});

// Register services
builder.Services.AddScoped<ExcelService>();
builder.Services.AddScoped<DataValidationService>();

// Configure CORS
builder.Services.AddCors(options =>
{
    options.AddDefaultPolicy(policy =>
    {
        policy.AllowAnyOrigin()
              .AllowAnyHeader()
              .AllowAnyMethod();
    });
});

// Configure file upload size limit
builder.Services.Configure<IISServerOptions>(options =>
{
    options.MaxRequestBodySize = 30 * 1024 * 1024; // 30 MB
});

builder.Services.Configure<FormOptions>(options =>
{
    options.MultipartBodyLengthLimit = 30 * 1024 * 1024; // 30 MB
});

var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI(c => 
    {
        c.SwaggerEndpoint("/swagger/v1/swagger.json", "Excel Data Import API v1");
    });
}

app.UseHttpsRedirection();

app.UseCors();

app.UseAuthorization();

app.MapControllers();

app.Run();
