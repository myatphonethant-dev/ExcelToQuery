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

                _logger.LogInformation($"Starting import: File={file.FileName}, Table={tableName}, Database={targetDatabase}");

                // Read file into memory
                using var stream = new MemoryStream();
                await file.CopyToAsync(stream);
                stream.Position = 0;

                // Parse Excel file
                using var package = new ExcelPackage(stream);
                var worksheet = package.Workbook.Worksheets[0];

                if (worksheet.Dimension == null)
                    throw new Exception("Excel file is empty or invalid");

                // Extract data from Excel (always assume has headers)
                var (data, headers) = ExtractDataFromWorksheet(worksheet, true);

                if (!data.Any())
                    throw new Exception("No data found in Excel file");

                _logger.LogInformation($"Found {data.Count} rows, {headers.Count} columns");

                // Import to database (ONLY INSERT, NO TRUNCATE)
                var recordsImported = await ImportToDatabase(
                    targetDatabase,
                    tableName,
                    data);

                _logger.LogInformation($"Import completed: {recordsImported}/{data.Count} records inserted to {tableName}");

                return recordsImported;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error importing Excel file: {ex.Message}");
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
            using var connection = new SqlConnection(connectionString);

            try
            {
                await connection.OpenAsync();

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

            // Create transaction for batch insert
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

                            // Set appropriate value and type
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

                        // Log progress every 100 rows
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
                return recordsImported;
            }
            catch (Exception)
            {
                await transaction.RollbackAsync();
                throw;
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

        private string GetConnectionString(string database)
        {
            var connectionString = _configuration.GetConnectionString(database);

            if (string.IsNullOrEmpty(connectionString))
            {
                // Fallback to default connection string
                connectionString = _configuration.GetConnectionString("DefaultConnection");

                if (string.IsNullOrEmpty(connectionString))
                {
                    throw new Exception($"Connection string not found for database: {database}");
                }
            }

            return connectionString;
        }
    }
}