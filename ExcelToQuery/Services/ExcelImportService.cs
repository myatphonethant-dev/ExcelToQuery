using Microsoft.Data.SqlClient;
using OfficeOpenXml;
using System.Data;
using System.Text.RegularExpressions;

namespace ExcelToQuery.Services
{
    public interface IExcelImportService
    {
        Task<int> ImportExcel(IFormFile file, string tableName, string targetDatabase);
    }

    public class ExcelImportService : IExcelImportService
    {
        private readonly ILogger<ExcelImportService> _logger;

        public ExcelImportService(IConfiguration configuration, ILogger<ExcelImportService> logger)
        {
            _configuration = configuration ?? throw new ArgumentNullException(nameof(configuration));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public async Task<int> ImportExcel(IFormFile file, string tableName, string targetDatabase)
        {
            var startTime = DateTime.Now;

            try
            {
                // Validate inputs
                if (file == null || file.Length == 0)
                    throw new ArgumentException("No file uploaded");

                if (string.IsNullOrWhiteSpace(tableName))
                    throw new ArgumentException("Table name is required");

                if (string.IsNullOrWhiteSpace(targetDatabase))
                    throw new ArgumentException("Target database is required");

                _logger.LogInformation($"=== Starting Import ===");
                _logger.LogInformation($"File: {file.FileName}, Size: {file.Length} bytes");
                _logger.LogInformation($"Table: {tableName}, Database: {targetDatabase}");

                // Read file into memory
                using var stream = new MemoryStream();
                await file.CopyToAsync(stream);
                stream.Position = 0;

                // Parse Excel file
                using var package = new ExcelPackage(stream);
                var worksheet = package.Workbook.Worksheets[0];

                if (worksheet.Dimension == null)
                    throw new Exception("Excel file is empty or invalid");

                _logger.LogInformation($"Excel worksheet: {worksheet.Name}, Dimensions: {worksheet.Dimension.Address}");

                // Extract data from Excel (always assume has headers)
                var (data, headers) = ExtractDataFromWorksheet(worksheet, true);

                if (!data.Any())
                    throw new Exception("No data found in Excel file");

                _logger.LogInformation($"Extracted {data.Count} rows, {headers.Count} columns");

                // Import to database
                var recordsImported = await ImportToDatabase(
                    targetDatabase,
                    tableName,
                    data);

                var elapsed = (DateTime.Now - startTime).TotalSeconds;
                _logger.LogInformation($"=== Import Completed ===");
                _logger.LogInformation($"Records imported: {recordsImported}");
                _logger.LogInformation($"Time taken: {elapsed:F2} seconds");

                return recordsImported;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"!!! Import Failed !!!");
                throw;
            }
        }

        private (List<Dictionary<string, object>> data, List<string> headers)
            ExtractDataFromWorksheet(ExcelWorksheet worksheet, bool hasHeaders)
        {
            var data = new List<Dictionary<string, object>>();
            var headers = new List<string>();

            if (worksheet.Dimension == null)
                return (data, headers);

            var rowCount = worksheet.Dimension.Rows;
            var colCount = worksheet.Dimension.Columns;

            // Read headers
            for (int col = 1; col <= colCount; col++)
            {
                if (hasHeaders)
                {
                    var header = worksheet.Cells[1, col].Text?.Trim();
                    if (string.IsNullOrEmpty(header))
                        header = $"Column{col}";
                    headers.Add(SanitizeColumnName(header));
                }
                else
                {
                    headers.Add($"Column{col}");
                }
            }

            // Read data rows
            int startRow = hasHeaders ? 2 : 1;
            for (int row = startRow; row <= rowCount; row++)
            {
                var rowData = new Dictionary<string, object>();
                bool hasValues = false;

                for (int col = 1; col <= colCount; col++)
                {
                    var header = headers[col - 1];
                    var cell = worksheet.Cells[row, col];
                    var value = GetCellValue(cell);

                    if (value != DBNull.Value)
                        hasValues = true;

                    rowData[header] = value;
                }

                if (hasValues)
                    data.Add(rowData);
            }

            return (data, headers);
        }

        private object GetCellValue(ExcelRange cell)
        {
            if (cell.Value == null)
                return DBNull.Value;

            return cell.Value switch
            {
                DateTime dt => dt,
                TimeSpan ts => ts,
                double d when Math.Abs(d % 1) <= double.Epsilon * 100 => Convert.ToInt32(d),
                double d => d,
                decimal dec => dec,
                bool b => b,
                string s when string.IsNullOrWhiteSpace(s) => DBNull.Value,
                string s => s.Trim(),
                _ => cell.Value.ToString()
            };
        }

        private string SanitizeColumnName(string columnName)
        {
            if (string.IsNullOrWhiteSpace(columnName))
                return "Column";

            var sanitized = Regex.Replace(columnName, @"[^\w]", "_");

            if (!Regex.IsMatch(sanitized, @"^[a-zA-Z_]"))
                sanitized = "_" + sanitized;

            sanitized = Regex.Replace(sanitized, @"_+", "_");
            return sanitized;
        }

        private async Task<int> ImportToDatabase(
            string database,
            string tableName,
            List<Dictionary<string, object>> data)
        {
            if (!data.Any())
                throw new Exception("No data to import");

            _logger.LogInformation($"Importing to {database}.{tableName}");

            // Build connection string manually to avoid config issues
            var connectionString = BuildConnectionString(database);

            _logger.LogInformation($"Using connection string: {MaskPassword(connectionString)}");

            using var connection = new SqlConnection(connectionString);

            try
            {
                _logger.LogInformation("Opening connection...");
                await connection.OpenAsync();
                _logger.LogInformation($"Connection opened. State: {connection.State}");

                // Check if table exists
                if (!await TableExists(connection, tableName))
                {
                    throw new Exception($"Table '{tableName}' does not exist in database '{database}'");
                }

                // Insert data
                return await SimpleBulkInsert(connection, tableName, data);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Database import error");
                throw new Exception($"Failed to import to database: {ex.Message}", ex);
            }
        }

        private string BuildConnectionString(string database)
        {
            // Simple hardcoded connection string - adjust as needed
            var dbName = database.Equals("gtb_wallet_log", StringComparison.OrdinalIgnoreCase)
                ? "gtb_wallet_log"
                : "gtb_wallet";

            return $"Server=.;Database={dbName};User ID=sa;Password=sasa@123;TrustServerCertificate=True;Connection Timeout=30;";
        }

        private async Task<bool> TableExists(SqlConnection connection, string tableName)
        {
            try
            {
                var query = "SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = @TableName";
                using var cmd = new SqlCommand(query, connection);
                cmd.Parameters.AddWithValue("@TableName", tableName);

                var result = await cmd.ExecuteScalarAsync();
                return Convert.ToInt32(result) > 0;
            }
            catch
            {
                return false;
            }
        }

        private async Task<int> SimpleBulkInsert(
            SqlConnection connection,
            string tableName,
            List<Dictionary<string, object>> data)
        {
            if (!data.Any()) return 0;

            var columns = data[0].Keys.ToList();
            var columnNames = string.Join(", ", columns.Select(c => $"[{c}]"));
            var paramNames = string.Join(", ", columns.Select(c => $"@{c}"));

            var query = $"INSERT INTO [{tableName}] ({columnNames}) VALUES ({paramNames})";

            _logger.LogInformation($"Insert query: {query}");

            using var transaction = await connection.BeginTransactionAsync();
            var recordsImported = 0;

            try
            {
                for (int i = 0; i < data.Count; i++)
                {
                    using var cmd = new SqlCommand(query, connection, (SqlTransaction)transaction);

                    // Add parameters
                    foreach (var column in columns)
                    {
                        cmd.Parameters.AddWithValue($"@{column}", data[i][column] ?? DBNull.Value);
                    }

                    await cmd.ExecuteNonQueryAsync();
                    recordsImported++;

                    if ((i + 1) % 100 == 0)
                    {
                        _logger.LogInformation($"Inserted {i + 1} rows...");
                    }
                }

                await transaction.CommitAsync();
                _logger.LogInformation($"Committed transaction. Total: {recordsImported} rows");

                return recordsImported;
            }
            catch (Exception ex)
            {
                await transaction.RollbackAsync();
                _logger.LogError(ex, $"Insert failed at row {recordsImported + 1}");
                throw;
            }
        }

        // Helper method to mask password in logs
        private string MaskPassword(string connectionString)
        {
            if (string.IsNullOrEmpty(connectionString))
                return connectionString;

            var regex = new Regex(@"Password=([^;]+)", RegexOptions.IgnoreCase);
            return regex.Replace(connectionString, "Password=***");
        }
    }
}