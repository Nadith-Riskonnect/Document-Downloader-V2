namespace DocumentDownloader.Models
{
    /// <summary>
    /// Connection settings for the database
    /// </summary>
    public class ConnectionSettings
    {
        public string ServerName { get; set; } = string.Empty;
        public string DatabaseName { get; set; } = string.Empty;
        public string UserId { get; set; } = string.Empty;
        public string Password { get; set; } = string.Empty;
        public string OutputFolder { get; set; } = string.Empty;

        public string ConnectionString => 
            $"Server={ServerName};Database={DatabaseName};User Id={UserId};Password={Password};";

        public bool IsValid =>
            !string.IsNullOrWhiteSpace(ServerName) &&
            !string.IsNullOrWhiteSpace(DatabaseName) &&
            !string.IsNullOrWhiteSpace(UserId) &&
            !string.IsNullOrWhiteSpace(Password) &&
            !string.IsNullOrWhiteSpace(OutputFolder);
    }

    /// <summary>
    /// Represents a single document download progress item for display in the UI
    /// </summary>
    public class DownloadProgressItem : ViewModels.BaseViewModel
    {
        private string _category = string.Empty;
        private string _status = "Pending";
        private int _totalDocuments;
        private int _successCount;
        private int _failedCount;
        private int _skippedDuplicates;
        private string _currentFile = string.Empty;

        public string Category
        {
            get => _category;
            set { _category = value; OnPropertyChanged(); }
        }

        public string Status
        {
            get => _status;
            set { _status = value; OnPropertyChanged(); }
        }

        public int TotalDocuments
        {
            get => _totalDocuments;
            set { _totalDocuments = value; OnPropertyChanged(); }
        }

        public int SuccessCount
        {
            get => _successCount;
            set { _successCount = value; OnPropertyChanged(); }
        }

        public int FailedCount
        {
            get => _failedCount;
            set { _failedCount = value; OnPropertyChanged(); }
        }

        public int SkippedDuplicates
        {
            get => _skippedDuplicates;
            set { _skippedDuplicates = value; OnPropertyChanged(); }
        }

        public string CurrentFile
        {
            get => _currentFile;
            set { _currentFile = value; OnPropertyChanged(); }
        }
    }

    /// <summary>
    /// Represents a log entry for the download process
    /// </summary>
    public class DownloadLogEntry
    {
        public DateTime Timestamp { get; set; } = DateTime.Now;
        public string Category { get; set; } = string.Empty;
        public string Message { get; set; } = string.Empty;
        public LogLevel Level { get; set; } = LogLevel.Info;
        public string FileName { get; set; } = string.Empty;
        public string FileSize { get; set; } = string.Empty;
    }

    public enum LogLevel
    {
        Info,
        Success,
        Warning,
        Error
    }

    /// <summary>
    /// Document categories for processing
    /// </summary>
    public enum DocumentCategory
    {
        Risk,
        Incident,
        Control,
        Action,
        Compliance,
        AuditRecommendation,
        AuditDetails,
        AuditFinding,
        Policy
    }

    /// <summary>
    /// Tracks file information for duplicate detection
    /// </summary>
    public class FileTracker
    {
        public string FileName { get; set; } = string.Empty;
        public string FileHash { get; set; } = string.Empty;
        public long FileSize { get; set; }
        public string SourcePath { get; set; } = string.Empty;
    }
}
