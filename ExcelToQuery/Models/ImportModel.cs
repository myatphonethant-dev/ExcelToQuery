namespace ExcelToQuery.Models
{
    public class FileUploadRequest
    {
        public IFormFile File { get; set; }
        public string? TableName { get; set; } // Made nullable for auto-detection
        public string? TargetDatabase { get; set; } // Made nullable for auto-detection
        public bool HasHeaders { get; set; } = true;
        public string? SheetName { get; set; } = "Sheet1"; // Added sheet name
        public bool TruncateTable { get; set; } = false; // Added truncate option
        public List<ColumnMapping> ColumnMappings { get; set; } = new List<ColumnMapping>();
        public string? TransactionType { get; set; } // For custom logic
        public DateTime? ImportDate { get; set; } // Track import date
        public string? Notes { get; set; } // Additional notes
    }

    public class ColumnMapping
    {
        public string SourceColumn { get; set; } = string.Empty;
        public string TargetColumn { get; set; } = string.Empty;
        public string DataType { get; set; } = "string";
        public string? DefaultValue { get; set; } // Added default value
        public bool IsRequired { get; set; } = false; // Added required flag
        public string? Transformation { get; set; } // For custom transformations
        public int? MaxLength { get; set; } // For validation
        public bool IgnoreOnError { get; set; } = false; // Skip if conversion fails
    }

    // Response model
    public class FileUploadResponseModel
    {
        public bool Success { get; set; }
        public string Message { get; set; } = string.Empty;
        public int TotalRecords { get; set; }
        public int RecordsImported { get; set; }
        public int RecordsSkipped { get; set; }
        public string? TargetDatabase { get; set; }
        public string? TargetTable { get; set; }
        public string? ImportedFileName { get; set; }
        public DateTime ImportTimestamp { get; set; } = DateTime.UtcNow;
        public List<string> Errors { get; set; } = new List<string>();
        public List<string> Warnings { get; set; } = new List<string>();
        public TimeSpan ProcessingTime { get; set; }
        public Dictionary<string, object>? Metadata { get; set; } // Additional data
    }

    // For auto-detection
    public class FileAutoDetectModel
    {
        public string FileName { get; set; } = string.Empty;
        public string? SuggestedDatabase { get; set; }
        public string? SuggestedTableName { get; set; }
        public List<DetectedColumn> DetectedColumns { get; set; } = new List<DetectedColumn>();
        public int RowCount { get; set; }
        public int ColumnCount { get; set; }
        public string? SheetName { get; set; }
    }

    public class DetectedColumn
    {
        public string Name { get; set; } = string.Empty;
        public string DataType { get; set; } = "string";
        public bool HasNullValues { get; set; }
        public string? SampleValue { get; set; }
        public int MaxLength { get; set; }
    }

    // For bulk operations
    public class BulkUploadRequestModel
    {
        public List<IFormFile> Files { get; set; } = new List<IFormFile>();
        public bool ProcessInParallel { get; set; } = true;
        public string? DefaultDatabase { get; set; }
        public bool StopOnFirstError { get; set; } = false;
    }

    public class BulkUploadResponseModel
    {
        public bool OverallSuccess { get; set; }
        public int TotalFiles { get; set; }
        public int SuccessfulFiles { get; set; }
        public int FailedFiles { get; set; }
        public int TotalRecordsImported { get; set; }
        public List<FileResult> FileResults { get; set; } = new List<FileResult>();
        public DateTime ProcessedAt { get; set; } = DateTime.UtcNow;
    }

    public class FileResult
    {
        public string FileName { get; set; } = string.Empty;
        public bool Success { get; set; }
        public string? Database { get; set; }
        public string? TableName { get; set; }
        public int RecordsImported { get; set; }
        public string? ErrorMessage { get; set; }
        public TimeSpan ProcessingTime { get; set; }
    }

    // Database connection info
    public class DatabaseInfoModel
    {
        public string Name { get; set; } = string.Empty;
        public string DisplayName { get; set; } = string.Empty;
        public string ConnectionStringKey { get; set; } = string.Empty;
        public bool IsDefault { get; set; }
        public DateTime LastConnected { get; set; }
        public string? Description { get; set; }
    }

    // Table schema info
    public class TableSchemaModel
    {
        public string Database { get; set; } = string.Empty;
        public string Schema { get; set; } = "dbo";
        public string TableName { get; set; } = string.Empty;
        public string FullName => $"[{Database}].[{Schema}].[{TableName}]";
        public List<TableColumn> Columns { get; set; } = new List<TableColumn>();
        public List<TableIndex> Indexes { get; set; } = new List<TableIndex>();
        public long? RowCount { get; set; }
        public DateTime? CreatedDate { get; set; }
        public DateTime? ModifiedDate { get; set; }
    }

    public class TableColumn
    {
        public string ColumnName { get; set; } = string.Empty;
        public string DataType { get; set; } = string.Empty;
        public int MaxLength { get; set; }
        public int? Precision { get; set; }
        public int? Scale { get; set; }
        public bool IsNullable { get; set; }
        public bool IsPrimaryKey { get; set; }
        public bool IsIdentity { get; set; }
        public bool IsComputed { get; set; }
        public string? DefaultValue { get; set; }
        public string? Description { get; set; }
    }

    public class TableIndex
    {
        public string IndexName { get; set; } = string.Empty;
        public bool IsUnique { get; set; }
        public bool IsPrimaryKey { get; set; }
        public List<string> ColumnNames { get; set; } = new List<string>();
    }

    // Import configuration
    public class ImportConfigModel
    {
        public List<string> AllowedExtensions { get; set; } = new()
        {
            ".xlsx", ".xls", ".csv"
        };

        public long MaxFileSize { get; set; } = 50 * 1024 * 1024; // 50MB

        public Dictionary<string, string> DatabaseMappings { get; set; } = new()
        {
            { "Variya", "gtb_wallet" },
            { "WalletTranLog", "gtb_wallet" },
            { "ThirdPartyIntegration", "gtb_wallet" },
            { "BillPaymentLog", "gtb_wallet_log" }
        };

        public Dictionary<string, string> TableMappings { get; set; } = new()
        {
            { "Variya", "Tbl_VariyaLog" },
            { "WalletTranLog", "Tbl_WalletTranLog" },
            { "BillPaymentLog", "Tbl_BillPaymentLog" },
            { "ThirdPartyIntegration", "Tbl_ThirdPartyIntegration" }
        };

        public int BatchSize { get; set; } = 100;
        public int CommandTimeout { get; set; } = 300; // seconds
        public bool LogDetailedErrors { get; set; } = true;
    }

    // For advanced filtering
    public class ImportFilterModel
    {
        public string? WhereClause { get; set; }
        public List<string> SelectedColumns { get; set; } = new List<string>();
        public int? SkipRows { get; set; }
        public int? TakeRows { get; set; }
        public string? OrderBy { get; set; }
        public bool Distinct { get; set; }
    }

    // Data validation rules
    public class ValidationRule
    {
        public string ColumnName { get; set; } = string.Empty;
        public string RuleType { get; set; } = string.Empty; // "required", "regex", "range", "custom"
        public string? Pattern { get; set; }
        public object? MinValue { get; set; }
        public object? MaxValue { get; set; }
        public string? ErrorMessage { get; set; }
        public bool SkipInvalidRows { get; set; } = false;
    }
}

public class SimpleImportRequest
{
    public IFormFile File { get; set; }
    public string TableName { get; set; }
    public string TargetDatabase { get; set; }
}