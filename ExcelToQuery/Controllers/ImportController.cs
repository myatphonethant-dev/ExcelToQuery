using ExcelToQuery.Services;
using Microsoft.AspNetCore.Mvc;

namespace ExcelToQuery.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class ImportController : ControllerBase
    {
        private readonly IExcelImportService _excelImportService;

        public ImportController(IExcelImportService excelImportService)
        {
            _excelImportService = excelImportService;
        }

        [HttpPost("upload")]
        [Consumes("multipart/form-data")] // ✅ Add this attribute
        public async Task<IActionResult> ImportExcel(
            [FromForm] IFormFile file,
            [FromForm] string tableName,
            [FromForm] string targetDatabase)
        {
            try
            {
                var recordsImported = await _excelImportService.ImportExcel(
                    file,
                    tableName,
                    targetDatabase);

                return Ok(new
                {
                    success = true,
                    message = "File imported successfully",
                    recordsImported = recordsImported,
                    tableName = tableName,
                    database = targetDatabase
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new
                {
                    success = false,
                    message = $"Import failed: {ex.Message}"
                });
            }
        }
    }
}