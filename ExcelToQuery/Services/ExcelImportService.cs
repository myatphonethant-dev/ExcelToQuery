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
        private readonly IConfiguration _configuration;
        private readonly ILogger<ExcelImportService> _logger;

        public ExcelImportService(IConfiguration configuration, ILogger<ExcelImportService> logger)
        {
            _configuration = configuration;
            _logger = logger;
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

            var connectionString = GetConnectionString(database);
            _logger.LogInformation($"Using connection string: {MaskPassword(connectionString)}");

            using var connection = new SqlConnection(connectionString);

            try
            {
                await connection.OpenAsync();
                _logger.LogInformation($"Connected to database: {database}");

                // Check if table exists before inserting
                await CheckIfTableExists(connection, tableName);

                // Get column names from first row
                var columns = data[0].Keys.ToList();

                // Insert data (NO TRUNCATE)
                var recordsImported = await BulkInsertData(
                    connection,
                    tableName,
                    columns,
                    data);

                return recordsImported;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error importing to database {database}, table {tableName}");
                throw new Exception($"Database import failed: {ex.Message}", ex);
            }
        }

        private async Task<int> BulkInsertData(
            SqlConnection connection,
            string tableName,
            List<string> columns,
            List<Dictionary<string, object>> data)
        {
            var recordsImported = 0;

            // Build parameterized INSERT statement
            var columnNames = string.Join(", ", columns.Select(c => $"[{c}]"));
            var paramNames = string.Join(", ", columns.Select(c => $"@{c}"));

            var insertQuery = $@"
        INSERT INTO [{tableName}] ({columnNames})
        VALUES ({paramNames})";

            // Create transaction
            using var transaction = await connection.BeginTransactionAsync() as SqlTransaction;

            try
            {
                using var command = new SqlCommand(insertQuery, connection, transaction);

                // Add parameters
                foreach (var column in columns)
                {
                    command.Parameters.Add($"@{column}", SqlDbType.NVarChar);
                }

                // Insert rows
                for (int i = 0; i < data.Count; i++)
                {
                    try
                    {
                        var row = data[i];

                        // Set parameter values
                        foreach (var column in columns)
                        {
                            var param = command.Parameters[$"@{column}"];
                            var value = row.ContainsKey(column) ? row[column] : DBNull.Value;

                            if (value == DBNull.Value || value == null)
                            {
                                param.Value = DBNull.Value;
                            }
                            else
                            {
                                param.Value = value;
                                param.SqlDbType = GetSqlDbType(value);
                            }
                        }

                        await command.ExecuteNonQueryAsync();
                        recordsImported++;

                        if ((i + 1) % 100 == 0)
                        {
                            _logger.LogInformation($"Inserted {i + 1} rows to {tableName}");
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning($"Error inserting row {i + 1}: {ex.Message}");
                        // Continue with next row
                    }
                }

                await transaction.CommitAsync();
                _logger.LogInformation($"Transaction committed. Total inserted: {recordsImported}");
                return recordsImported;
            }
            catch (Exception ex)
            {
                await transaction.RollbackAsync();
                _logger.LogError(ex, $"Transaction rolled back due to error");
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

        private async Task CheckIfTableExists(SqlConnection connection, string tableName)
        {
            var checkQuery = $@"
                SELECT COUNT(*) 
                FROM INFORMATION_SCHEMA.TABLES 
                WHERE TABLE_NAME = @TableName";

            using var command = new SqlCommand(checkQuery, connection);
            command.Parameters.AddWithValue("@TableName", tableName);

            var tableExists = (int)await command.ExecuteScalarAsync() > 0;

            if (!tableExists)
            {
                throw new Exception($"Table '{tableName}' does not exist in the database. Please create it first.");
            }
        }

        private SqlDbType GetSqlDbType(object value)
        {
            return value switch
            {
                int => SqlDbType.Int,
                long => SqlDbType.BigInt,
                decimal => SqlDbType.Decimal,
                float => SqlDbType.Float,
                double => SqlDbType.Float,
                DateTime => SqlDbType.DateTime,
                bool => SqlDbType.Bit,
                Guid => SqlDbType.UniqueIdentifier,
                _ => SqlDbType.NVarChar
            };
        }

        private string GetConnectionStringV1(string database)
        {
            // Try to get connection string from configuration
            var connectionString = _configuration.GetConnectionString(database);

            if (!string.IsNullOrEmpty(connectionString))
            {
                return connectionString;
            }

            // Fallback to default connection
            connectionString = _configuration.GetConnectionString("DefaultConnection");

            if (!string.IsNullOrEmpty(connectionString))
            {
                return connectionString;
            }

            // If still not found, throw detailed error
            var allConnections = _configuration.GetSection("ConnectionStrings").GetChildren();
            var availableConnections = string.Join(", ", allConnections.Select(c => c.Key));

            throw new Exception($"Connection string not found for '{database}'. Available connections: {availableConnections}");
        }

        private string GetConnectionString(string database)
        {
            // Temporary hardcoded connection string for testing
            string connectionString;

            if (database.Equals("gtb_wallet_log", StringComparison.OrdinalIgnoreCase))
            {
                connectionString = "Server=.;Database=gtb_wallet_log;User ID=sa;Password=sasa@123;TrustServerCertificate=true;";
            }
            else
            {
                connectionString = "Server=.;Database=gtb_wallet;User ID=sa;Password=sasa@123;TrustServerCertificate=true;";
            }

            _logger.LogInformation($"Using connection string for {database}: {MaskPassword(connectionString)}");
            return connectionString;
        }
    }
}