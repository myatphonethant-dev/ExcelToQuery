using ExcelToQuery.Models;
using Microsoft.AspNetCore.Mvc;

namespace ExcelToQuery.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class ImportController : ControllerBase
    {
        [HttpPost("upload")]
        public async Task<ActionResult<FileUploadResponseModel>> UploadFile(
            [FromForm] FileUploadRequest request)
        {
            try
            {
                // Validate file
                if (request.File == null || request.File.Length == 0)
                    return BadRequest(new FileUploadResponseModel
                    {
                        Success = false,
                        Message = "No file uploaded"
                    });

                // Auto-detect database from filename if not specified
                if (string.IsNullOrEmpty(request.TargetDatabase))
                {
                    request.TargetDatabase = DetectDatabaseFromFilename(request.File.FileName);
                }

                // Auto-detect table name from filename if not specified
                if (string.IsNullOrEmpty(request.TableName))
                {
                    request.TableName = DetectTableNameFromFilename(request.File.FileName);
                }

                // Process the file...
                var result = new FileUploadResponseModel
                {
                    Success = true,
                    Message = "File imported successfully",
                    TargetDatabase = request.TargetDatabase,
                    TargetTable = request.TableName,
                    ImportedFileName = request.File.FileName
                };

                return Ok(result);
            }
            catch (Exception ex)
            {
                return StatusCode(500, new FileUploadResponseModel
                {
                    Success = false,
                    Message = $"Import failed: {ex.Message}",
                    Errors = new List<string> { ex.Message }
                });
            }
        }

        private string DetectDatabaseFromFilename(string filename)
        {
            var fileName = Path.GetFileNameWithoutExtension(filename);

            if (fileName.Contains("Variya", StringComparison.OrdinalIgnoreCase) ||
                fileName.Contains("WalletTranLog", StringComparison.OrdinalIgnoreCase) ||
                fileName.Contains("ThirdPartyIntegration", StringComparison.OrdinalIgnoreCase))
            {
                return "gtb_wallet";
            }
            else if (fileName.Contains("BillPaymentLog", StringComparison.OrdinalIgnoreCase))
            {
                return "gtb_wallet_log";
            }

            return "gtb_wallet"; // default
        }

        private string DetectTableNameFromFilename(string filename)
        {
            var fileName = Path.GetFileNameWithoutExtension(filename);

            if (fileName.Contains("Variya", StringComparison.OrdinalIgnoreCase))
                return "Tbl_VariyaLog";
            else if (fileName.Contains("WalletTranLog", StringComparison.OrdinalIgnoreCase))
                return "Tbl_WalletTranLog";
            else if (fileName.Contains("BillPaymentLog", StringComparison.OrdinalIgnoreCase))
                return "Tbl_BillPaymentLog";
            else if (fileName.Contains("ThirdPartyIntegration", StringComparison.OrdinalIgnoreCase))
                return "Tbl_ThirdPartyIntegration";

            return fileName;
        }
    }
}