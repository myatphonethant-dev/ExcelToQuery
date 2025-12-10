using ExcelToQuery.Models;
using Microsoft.Data.SqlClient;
using OfficeOpenXml;
using System.Data;
using System.Text.RegularExpressions;

namespace ExcelToQuery.Services
{
    public interface IExcelImportService
    {
        Task<int> ImportExcel(IFormFile file, string tableName, string targetDatabase);
        Task<TableSchemaModel> GetTableSchema(string database, string schema, string tableName);
    }

    public class ExcelImportService : IExcelImportService
    {
        private readonly IConfiguration _configuration;
        private readonly ILogger<ExcelImportService> _logger;
        private readonly Dictionary<string, string> _databaseMapping;
        private readonly Dictionary<string, string> _tableMapping;

        public ExcelImportService(IConfiguration configuration, ILogger<ExcelImportService> logger)
        {
            _configuration = configuration;
            _logger = logger;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Initialize mappings from configuration or hardcoded
            _databaseMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "Variya", "gtb_wallet" },
                { "WalletTranLog", "gtb_wallet" },
                { "ThirdPartyIntegration", "gtb_wallet" },
                { "BillPaymentLog", "gtb_wallet_log" },
                { "Tbl_VariyaLog", "gtb_wallet" },
                { "Tbl_WalletTranLog", "gtb_wallet" },
                { "Tbl_ThirdPartyIntegration", "gtb_wallet" },
                { "Tbl_BillPaymentLog", "gtb_wallet_log" }
            };

            _tableMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "Variya", "Tbl_VariyaLog" },
                { "WalletTranLog", "Tbl_WalletTranLog" },
                { "BillPaymentLog", "Tbl_BillPaymentLog" },
                { "ThirdPartyIntegration", "Tbl_ThirdPartyIntegration" }
            };
        }

        #region File Import Methods

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
                var worksheet = package.Workbook.Worksheets[0]; // Always use first sheet

                if (worksheet.Dimension == null)
                    throw new Exception("Excel file is empty or invalid");

                // Extract data from Excel (always assume has headers)
                var (data, headers) = ExtractDataFromWorksheet(worksheet, true);
                var totalRecords = data.Count;

                if (!data.Any())
                    throw new Exception("No data found in Excel file");

                // Import to database (always truncate table)
                var recordsImported = await ImportToDatabase(
                    targetDatabase,
                    tableName,
                    data,
                    truncateTable: true);

                _logger.LogInformation($"Import completed: {recordsImported}/{totalRecords} records imported to {tableName}");

                return recordsImported;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error importing Excel file: {ex.Message}");
                throw;
            }
        }

        #endregion

        #region Helper Methods

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

            // Handle different data types
            return cell.Value switch
            {
                DateTime dt => dt,
                TimeSpan ts => ts,
                double d when cell.Style.Numberformat.Format.Contains("%") => d / 100.0, // Percentage
                double d when Math.Abs(d % 1) <= double.Epsilon * 100 => Convert.ToInt32(d), // Integer
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
            // Remove special characters, keep only letters, numbers, and underscores
            var sanitized = Regex.Replace(columnName, @"[^\w]", "_");

            // Ensure it starts with a letter or underscore
            if (!Regex.IsMatch(sanitized, @"^[a-zA-Z_]"))
                sanitized = "_" + sanitized;

            // Replace multiple underscores with single
            sanitized = Regex.Replace(sanitized, @"_+", "_");

            return sanitized;
        }

        #endregion

        #region Database Operations

        private async Task<int> ImportToDatabase(
            string database,
            string tableName,
            List<Dictionary<string, object>> data,
            bool truncateTable)
        {
            if (!data.Any())
            {
                throw new Exception("No data to import");
            }

            var connectionString = GetConnectionString(database);
            using var connection = new SqlConnection(connectionString);

            try
            {
                await connection.OpenAsync();
                using var transaction = (SqlTransaction)await connection.BeginTransactionAsync();

                try
                {
                    // Truncate table if requested
                    if (truncateTable)
                    {
                        await TruncateTable(connection, transaction, tableName);
                    }

                    // Get column names from first row
                    var columns = data[0].Keys.ToList();

                    // Insert data
                    var recordsImported = await BulkInsertData(
                        connection,
                        transaction,
                        tableName,
                        columns,
                        data);

                    await transaction.CommitAsync();

                    return recordsImported;
                }
                catch (Exception ex)
                {
                    await transaction.RollbackAsync();
                    throw new Exception($"Transaction failed: {ex.Message}", ex);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error importing to database {database}, table {tableName}");
                throw new Exception($"Database import failed: {ex.Message}", ex);
            }
        }

        private async Task TruncateTable(SqlConnection connection, SqlTransaction transaction, string tableName)
        {
            // Check if table exists
            var checkQuery = $"SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = @TableName";
            using var checkCmd = new SqlCommand(checkQuery, connection, transaction);
            checkCmd.Parameters.AddWithValue("@TableName", tableName);

            var tableExists = (int)await checkCmd.ExecuteScalarAsync() > 0;

            if (tableExists)
            {
                var truncateQuery = $"TRUNCATE TABLE [{tableName}]";
                using var truncateCmd = new SqlCommand(truncateQuery, connection, transaction);
                await truncateCmd.ExecuteNonQueryAsync();
                _logger.LogInformation($"Truncated table: {tableName}");
            }
            else
            {
                _logger.LogWarning($"Table {tableName} does not exist. Skipping truncate.");
            }
        }

        private async Task<int> BulkInsertData(
            SqlConnection connection,
            SqlTransaction transaction,
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
                        _logger.LogInformation($"Imported {i + 1} rows to {tableName}");
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning($"Error importing row {i + 1}: {ex.Message}");
                    // Continue with next row instead of throwing
                }
            }

            return recordsImported;
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
                throw new Exception($"Connection string not found for database: {database}");
            }

            return connectionString;
        }

        #endregion

        #region Other Service Methods

        public async Task<TableSchemaModel> GetTableSchema(string database, string schema, string tableName)
        {
            var tableSchema = new TableSchemaModel
            {
                Database = database,
                Schema = schema,
                TableName = tableName
            };

            try
            {
                var connectionString = GetConnectionString(database);
                using var connection = new SqlConnection(connectionString);
                await connection.OpenAsync();

                // Get columns
                var columnsQuery = @"
                    SELECT 
                        c.name as ColumnName,
                        t.name as DataType,
                        c.max_length as MaxLength,
                        c.precision as Precision,
                        c.scale as Scale,
                        c.is_nullable as IsNullable,
                        c.is_identity as IsIdentity,
                        c.is_computed as IsComputed,
                        OBJECT_DEFINITION(c.default_object_id) as DefaultValue
                    FROM sys.columns c
                    INNER JOIN sys.types t ON c.user_type_id = t.user_type_id
                    INNER JOIN sys.tables tab ON c.object_id = tab.object_id
                    INNER JOIN sys.schemas s ON tab.schema_id = s.schema_id
                    WHERE s.name = @Schema AND tab.name = @TableName
                    ORDER BY c.column_id";

                using var columnsCmd = new SqlCommand(columnsQuery, connection);
                columnsCmd.Parameters.AddWithValue("@Schema", schema);
                columnsCmd.Parameters.AddWithValue("@TableName", tableName);

                using var columnsReader = await columnsCmd.ExecuteReaderAsync();
                while (await columnsReader.ReadAsync())
                {
                    tableSchema.Columns.Add(new TableColumn
                    {
                        ColumnName = columnsReader["ColumnName"].ToString(),
                        DataType = columnsReader["DataType"].ToString(),
                        MaxLength = Convert.ToInt32(columnsReader["MaxLength"]),
                        Precision = columnsReader["Precision"] as int?,
                        Scale = columnsReader["Scale"] as int?,
                        IsNullable = Convert.ToBoolean(columnsReader["IsNullable"]),
                        IsIdentity = Convert.ToBoolean(columnsReader["IsIdentity"]),
                        IsComputed = Convert.ToBoolean(columnsReader["IsComputed"]),
                        DefaultValue = columnsReader["DefaultValue"]?.ToString()
                    });
                }

                await columnsReader.CloseAsync();

                // Get primary keys
                var pkQuery = @"
                    SELECT 
                        i.name as IndexName,
                        ic.is_included_column as IsIncluded,
                        c.name as ColumnName
                    FROM sys.indexes i
                    INNER JOIN sys.index_columns ic ON i.object_id = ic.object_id AND i.index_id = ic.index_id
                    INNER JOIN sys.columns c ON ic.object_id = c.object_id AND ic.column_id = c.column_id
                    INNER JOIN sys.tables t ON i.object_id = t.object_id
                    INNER JOIN sys.schemas s ON t.schema_id = s.schema_id
                    WHERE s.name = @Schema AND t.name = @TableName
                    AND i.is_primary_key = 1
                    ORDER BY ic.key_ordinal";

                using var pkCmd = new SqlCommand(pkQuery, connection);
                pkCmd.Parameters.AddWithValue("@Schema", schema);
                pkCmd.Parameters.AddWithValue("@TableName", tableName);

                using var pkReader = await pkCmd.ExecuteReaderAsync();
                while (await pkReader.ReadAsync())
                {
                    var columnName = pkReader["ColumnName"].ToString();
                    var column = tableSchema.Columns.FirstOrDefault(c => c.ColumnName == columnName);
                    if (column != null)
                    {
                        column.IsPrimaryKey = true;
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error getting schema for table {schema}.{tableName} in database {database}");
                throw;
            }

            return tableSchema;
        }

        #endregion
    }
}