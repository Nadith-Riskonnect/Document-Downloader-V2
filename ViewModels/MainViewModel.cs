using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using System.Windows.Input;
using DocumentDownloader.Models;
using DocumentDownloader.Services;
using DocumentDownloader.Shared.Helpers;

namespace DocumentDownloader.ViewModels
{
    public class MainViewModel : BaseViewModel
    {
        private readonly IDocumentDownloadService _downloadService;
        private CancellationTokenSource? _cancellationTokenSource;

        #region Properties

        private string _serverName = string.Empty;
        public string ServerName
        {
            get => _serverName;
            set { if (SetProperty(ref _serverName, value)) RaiseFormValidityChanged(); }
        }

        private string _databaseName = string.Empty;
        public string DatabaseName
        {
            get => _databaseName;
            set { if (SetProperty(ref _databaseName, value)) RaiseFormValidityChanged(); }
        }

        private string _userId = string.Empty;
        public string UserId
        {
            get => _userId;
            set { if (SetProperty(ref _userId, value)) RaiseFormValidityChanged(); }
        }

        private string _password = string.Empty;
        public string Password
        {
            get => _password;
            set { if (SetProperty(ref _password, value)) RaiseFormValidityChanged(); }
        }

        private string _outputFolder = string.Empty;
        public string OutputFolder
        {
            get => _outputFolder;
            set { if (SetProperty(ref _outputFolder, value)) RaiseFormValidityChanged(); }
        }

        private void RaiseFormValidityChanged()
        {
            OnPropertyChanged(nameof(IsFormValid));
            OnPropertyChanged(nameof(CanStartDownload));
            System.Windows.Input.CommandManager.InvalidateRequerySuggested();
        }

        private bool _isDownloading;
        public bool IsDownloading
        {
            get => _isDownloading;
            set
            {
                SetProperty(ref _isDownloading, value);
                OnPropertyChanged(nameof(CanStartDownload));
                OnPropertyChanged(nameof(CanCancel));
            }
        }

        private bool _isConnected;
        public bool IsConnected
        {
            get => _isConnected;
            set => SetProperty(ref _isConnected, value);
        }

        private string _statusMessage = "Ready";
        public string StatusMessage
        {
            get => _statusMessage;
            set => SetProperty(ref _statusMessage, value);
        }

        private string _connectionStatus = "Not Connected";
        public string ConnectionStatus
        {
            get => _connectionStatus;
            set => SetProperty(ref _connectionStatus, value);
        }

        public bool CanStartDownload => !IsDownloading && IsFormValid;
        public bool CanCancel => IsDownloading;

        public bool IsFormValid =>
            !string.IsNullOrWhiteSpace(ServerName) &&
            !string.IsNullOrWhiteSpace(DatabaseName) &&
            !string.IsNullOrWhiteSpace(UserId) &&
            !string.IsNullOrWhiteSpace(Password) &&
            !string.IsNullOrWhiteSpace(OutputFolder);

        // Summary properties
        private int _totalDocuments;
        public int TotalDocuments
        {
            get => _totalDocuments;
            set => SetProperty(ref _totalDocuments, value);
        }

        private int _successfulDownloads;
        public int SuccessfulDownloads
        {
            get => _successfulDownloads;
            set => SetProperty(ref _successfulDownloads, value);
        }

        private int _failedDownloads;
        public int FailedDownloads
        {
            get => _failedDownloads;
            set => SetProperty(ref _failedDownloads, value);
        }

        private int _duplicatesSkipped;
        public int DuplicatesSkipped
        {
            get => _duplicatesSkipped;
            set => SetProperty(ref _duplicatesSkipped, value);
        }

        #endregion

        #region Collections

        public ObservableCollection<DownloadProgressItem> ProgressItems { get; } = new();
        public ObservableCollection<DownloadLogEntry> LogEntries { get; } = new();

        #endregion

        #region Commands

        public ICommand TestConnectionCommand { get; }
        public ICommand StartDownloadCommand { get; }
        public ICommand CancelDownloadCommand { get; }
        public ICommand BrowseFolderCommand { get; }
        public ICommand ClearLogCommand { get; }

        #endregion

        public MainViewModel()
        {
            _downloadService = new DocumentDownloadService();

            // Initialize commands
            TestConnectionCommand = new RelayCommand(async _ => await TestConnectionAsync(), _ => IsFormValid && !IsDownloading);
            StartDownloadCommand = new RelayCommand(async _ => await StartDownloadAsync(), _ => CanStartDownload);
            CancelDownloadCommand = new RelayCommand(_ => CancelDownload(), _ => CanCancel);
            BrowseFolderCommand = new RelayCommand(_ => BrowseFolder());
            ClearLogCommand = new RelayCommand(_ => ClearLog());

            // Initialize progress items for each category
            InitializeProgressItems();
        }

        private void InitializeProgressItems()
        {
            ProgressItems.Clear();
            foreach (DocumentCategory category in Enum.GetValues<DocumentCategory>())
            {
                ProgressItems.Add(new DownloadProgressItem
                {
                    Category = category.ToString(),
                    Status = "Pending"
                });
            }
        }

        private ConnectionSettings GetConnectionSettings()
        {
            return new ConnectionSettings
            {
                ServerName = ServerName,
                DatabaseName = DatabaseName,
                UserId = UserId,
                Password = Password,
                OutputFolder = OutputFolder
            };
        }

        private async Task TestConnectionAsync()
        {
            var settings = GetConnectionSettings();
            StatusMessage = "Testing connection...";
            ConnectionStatus = "Connecting...";

            var (success, message) = await _downloadService.TestConnectionAsync(settings);

            if (success)
            {
                IsConnected = true;
                ConnectionStatus = "Connected";
                StatusMessage = message;
                AddLogEntry("System", message, LogLevel.Success);
            }
            else
            {
                IsConnected = false;
                ConnectionStatus = "Failed";
                StatusMessage = message;
                AddLogEntry("System", message, LogLevel.Error);
            }
        }

        private async Task StartDownloadAsync()
        {
            if (!IsFormValid)
            {
                StatusMessage = "Please fill in all required fields.";
                return;
            }

            // Validate output folder
            if (!IsValidPath(OutputFolder))
            {
                StatusMessage = "Invalid output folder path.";
                AddLogEntry("System", "Invalid output folder path specified.", LogLevel.Error);
                return;
            }

            IsDownloading = true;
            _cancellationTokenSource = new CancellationTokenSource();

            // Reset progress
            InitializeProgressItems();
            ResetSummary();
            StatusMessage = "Starting download...";

            try
            {
                var settings = GetConnectionSettings();

                await _downloadService.StartDownloadAsync(
                    settings,
                    ProgressItems,
                    entry => Application.Current.Dispatcher.Invoke(() => AddLogEntry(entry)),
                    _cancellationTokenSource.Token);

                // Calculate summary
                UpdateSummary();

                if (_cancellationTokenSource.Token.IsCancellationRequested)
                {
                    StatusMessage = "Download cancelled by user.";
                    AddLogEntry("System", "Download cancelled by user.", LogLevel.Warning);
                }
                else
                {
                    StatusMessage = $"Download complete! Total: {TotalDocuments}, Success: {SuccessfulDownloads}, Failed: {FailedDownloads}, Duplicates Skipped: {DuplicatesSkipped}";
                    AddLogEntry("System", StatusMessage, LogLevel.Success);
                }
            }
            catch (Exception ex)
            {
                StatusMessage = $"Error: {ex.Message}";
                AddLogEntry("System", $"Fatal error: {ex.Message}", LogLevel.Error);
            }
            finally
            {
                IsDownloading = false;
                _cancellationTokenSource?.Dispose();
                _cancellationTokenSource = null;
            }
        }

        private void CancelDownload()
        {
            _cancellationTokenSource?.Cancel();
            StatusMessage = "Cancelling...";
        }

        private void BrowseFolder()
        {
            var dialog = new Microsoft.Win32.OpenFolderDialog
            {
                Title = "Select Output Folder"
            };

            if (dialog.ShowDialog() == true)
            {
                OutputFolder = dialog.FolderName;
            }
        }

        private void ClearLog()
        {
            LogEntries.Clear();
        }

        private void AddLogEntry(string category, string message, LogLevel level)
        {
            LogEntries.Insert(0, new DownloadLogEntry
            {
                Timestamp = DateTime.Now,
                Category = category,
                Message = message,
                Level = level
            });

            // Limit log entries to prevent memory issues
            while (LogEntries.Count > 1000)
            {
                LogEntries.RemoveAt(LogEntries.Count - 1);
            }
        }

        private void AddLogEntry(DownloadLogEntry entry)
        {
            LogEntries.Insert(0, entry);

            // Limit log entries
            while (LogEntries.Count > 1000)
            {
                LogEntries.RemoveAt(LogEntries.Count - 1);
            }
        }

        private void ResetSummary()
        {
            TotalDocuments = 0;
            SuccessfulDownloads = 0;
            FailedDownloads = 0;
            DuplicatesSkipped = 0;
        }

        private void UpdateSummary()
        {
            TotalDocuments = ProgressItems.Sum(p => p.TotalDocuments);
            SuccessfulDownloads = ProgressItems.Sum(p => p.SuccessCount);
            FailedDownloads = ProgressItems.Sum(p => p.FailedCount);
            DuplicatesSkipped = ProgressItems.Sum(p => p.SkippedDuplicates);
        }

        private static bool IsValidPath(string path)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(path))
                    return false;

                // Check if path is rooted (absolute path)
                if (!Path.IsPathRooted(path))
                    return false;

                // Get full path to validate
                _ = Path.GetFullPath(path);
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
