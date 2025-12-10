using ExcelToQuery;
using ExcelToQuery.Services;
using Microsoft.Data.SqlClient;
using Microsoft.OpenApi.Models;
using System.Reflection;

var builder = WebApplication.CreateBuilder(args);

var testConnectionString = "Server=.;Database=gtb_wallet;User ID=sa;Password=sasa@123;TrustServerCertificate=True;";

try
{
    using var testConn = new SqlConnection(testConnectionString);
    await testConn.OpenAsync();
    Console.WriteLine("✅ SQL Server connection SUCCESS!");

    // Test query
    using var cmd = new SqlCommand("SELECT @@VERSION", testConn);
    var version = await cmd.ExecuteScalarAsync();
    Console.WriteLine($"SQL Server Version: {version}");
}
catch (Exception ex)
{
    Console.WriteLine($"❌ SQL Server connection FAILED: {ex.Message}");
    // Don't continue if SQL Server isn't accessible
    throw;
}

// Add services to the container.
builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen(options =>
{
    options.SwaggerDoc("v1", new OpenApiInfo
    {
        Title = "Excel to Query API",
        Version = "v1",
        Description = "API for uploading Excel files to SQL Server databases"
    });

    // Add file upload support
    options.OperationFilter<SwaggerFileOperationFilter>();

    // Enable XML comments
    var xmlFile = $"{Assembly.GetExecutingAssembly().GetName().Name}.xml";
    var xmlPath = Path.Combine(AppContext.BaseDirectory, xmlFile);
    if (File.Exists(xmlPath))
    {
        options.IncludeXmlComments(xmlPath);
    }
});

// ✅ Register your service
builder.Services.AddScoped<IExcelImportService, ExcelImportService>();

var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseHttpsRedirection();
app.UseAuthorization();
app.MapControllers();
app.Run();