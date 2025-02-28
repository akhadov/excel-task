using ExcelTask.Api.Data;
using ExcelTask.Api.Services;
using ExcelTask.Api.Services.Interfaces;

var builder = WebApplication.CreateBuilder(args);


string connectionString = builder.Configuration.GetConnectionString("Database")
                          ?? throw new InvalidOperationException("Database connection string is missing.");

builder.Services.AddSingleton<IDbConnectionFactory>(new DbConnectionFactory(connectionString));

builder.Services.AddScoped<IExcelProcessingService, ExcelProcessingService>();

builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

var app = builder.Build();

if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseHttpsRedirection();

app.UseAuthorization();

app.MapControllers();

app.Run();
