using DocumentDownloader.Models;
using System.Collections.ObjectModel;

namespace DocumentDownloader.Services
{
    /// <summary>
    /// Interface for the document download service
    /// </summary>
    public interface IDocumentDownloadService
    {
        /// <summary>
        /// Starts the download process for all document categories
        /// </summary>
        Task StartDownloadAsync(
            ConnectionSettings settings,
            ObservableCollection<DownloadProgressItem> progressItems,
            Action<DownloadLogEntry> logCallback,
            CancellationToken cancellationToken);

        /// <summary>
        /// Tests the database connection
        /// </summary>
        Task<(bool Success, string Message)> TestConnectionAsync(ConnectionSettings settings);
    }
}
