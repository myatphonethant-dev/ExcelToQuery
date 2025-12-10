using Microsoft.OpenApi.Models;
using Swashbuckle.AspNetCore.SwaggerGen;

namespace ExcelToQuery;

public class SwaggerFileUploadFilter : IOperationFilter
{
    public void Apply(OpenApiOperation operation, OperationFilterContext context)
    {
        var method = context.MethodInfo;

        if (method.Name == "ImportExcel")
        {
            operation.Parameters.Clear();

            operation.RequestBody = new OpenApiRequestBody
            {
                Content = new Dictionary<string, OpenApiMediaType>
                {
                    ["multipart/form-data"] = new OpenApiMediaType
                    {
                        Schema = new OpenApiSchema
                        {
                            Type = "object",
                            Properties = new Dictionary<string, OpenApiSchema>
                            {
                                ["file"] = new OpenApiSchema
                                {
                                    Type = "string",
                                    Format = "binary",
                                    Description = "Excel file to upload"
                                },
                                ["tableName"] = new OpenApiSchema
                                {
                                    Type = "string",
                                    Description = "Target table name"
                                },
                                ["targetDatabase"] = new OpenApiSchema
                                {
                                    Type = "string",
                                    Description = "Target database name"
                                }
                            },
                            Required = new HashSet<string> { "file", "tableName", "targetDatabase" }
                        }
                    }
                }
            };
        }
    }
}