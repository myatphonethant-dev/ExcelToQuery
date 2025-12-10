using ExcelToQuery.Services;
using Microsoft.OpenApi.Models;
using Swashbuckle.AspNetCore.SwaggerGen;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container
builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen(c =>
{
    c.SwaggerDoc("v1", new OpenApiInfo { Title = "Data Import API", Version = "v1" });
    c.OperationFilter<SwaggerFileUploadOperation>();
});

// Add CORS
builder.Services.AddCors(options =>
{
    options.AddPolicy("AllowAll", policy =>
    {
        policy.AllowAnyOrigin()
              .AllowAnyMethod()
              .AllowAnyHeader();
    });
});

// Add your services
builder.Services.AddScoped<IExcelImportService, ExcelImportService>();

var app = builder.Build();

// Configure the HTTP request pipeline
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseCors("AllowAll");
app.UseHttpsRedirection();
app.UseAuthorization();
app.MapControllers();

app.Run();

// Swagger file upload operation filter
public class SwaggerFileUploadOperation : IOperationFilter
{
    public void Apply(OpenApiOperation operation, OperationFilterContext context)
    {
        var fileParams = context.MethodInfo.GetParameters()
            .Where(p => p.ParameterType == typeof(IFormFile));

        if (fileParams.Any())
        {
            operation.Parameters.Clear();
            operation.Parameters.Add(new OpenApiParameter
            {
                Name = "file",
                In = ParameterLocation.Header,
                Description = "Upload file",
                Required = true,
                Schema = new OpenApiSchema
                {
                    Type = "string",
                    Format = "binary"
                }
            });
        }
    }
}