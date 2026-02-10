using System.Collections.Concurrent;
using System.Collections.ObjectModel;
using System.Data.SqlClient;
using System.IO;
using System.Security.Cryptography;
using DocumentDownloader.Models;

namespace DocumentDownloader.Services
{
    /// <summary>
    /// Service for downloading documents from the database with duplicate detection and error handling
    /// </summary>
    public class DocumentDownloadService : IDocumentDownloadService
    {
        // Thread-safe dictionary to track downloaded files by hash for duplicate detection
        private readonly ConcurrentDictionary<string, FileTracker> _downloadedFiles = new();
        
        // Track files by normalized name within each category for duplicate folder detection
        private readonly ConcurrentDictionary<string, HashSet<string>> _categoryFolders = new();

        public async Task<(bool Success, string Message)> TestConnectionAsync(ConnectionSettings settings)
        {
            try
            {
                using var connection = new SqlConnection(settings.ConnectionString);
                await connection.OpenAsync();
                return (true, "Connection successful!");
            }
            catch (SqlException ex)
            {
                return (false, $"SQL Error: {ex.Message}");
            }
            catch (Exception ex)
            {
                return (false, $"Error: {ex.Message}");
            }
        }

        public async Task StartDownloadAsync(
            ConnectionSettings settings,
            ObservableCollection<DownloadProgressItem> progressItems,
            Action<DownloadLogEntry> logCallback,
            CancellationToken cancellationToken)
        {
            // Clear tracking dictionaries for fresh start
            _downloadedFiles.Clear();
            _categoryFolders.Clear();

            // Create output folder if it doesn't exist
            if (!Directory.Exists(settings.OutputFolder))
            {
                Directory.CreateDirectory(settings.OutputFolder);
                logCallback(new DownloadLogEntry
                {
                    Category = "System",
                    Message = $"Created output folder: {settings.OutputFolder}",
                    Level = LogLevel.Info
                });
            }

            using var connection = new SqlConnection(settings.ConnectionString);
            await connection.OpenAsync(cancellationToken);

            logCallback(new DownloadLogEntry
            {
                Category = "System",
                Message = "Connected to database successfully",
                Level = LogLevel.Success
            });

            // Process each document category
            var tasks = new (DocumentCategory Category, Func<SqlConnection, string, DownloadProgressItem, Action<DownloadLogEntry>, CancellationToken, Task> Processor)[]
            {
                (DocumentCategory.Risk, ProcessRiskDocumentsAsync),
                (DocumentCategory.Incident, ProcessIncidentDocumentsAsync),
                (DocumentCategory.Control, ProcessControlDocumentsAsync),
                (DocumentCategory.Action, ProcessActionDocumentsAsync),
                (DocumentCategory.Compliance, ProcessComplianceDocumentsAsync),
                (DocumentCategory.AuditRecommendation, ProcessAuditRecommendationDocumentsAsync),
                (DocumentCategory.AuditDetails, ProcessAuditDetailsDocumentsAsync),
                (DocumentCategory.AuditFinding, ProcessAuditFindingDocumentsAsync),
                (DocumentCategory.Policy, ProcessPolicyDocumentsAsync)
            };

            foreach (var (category, processor) in tasks)
            {
                if (cancellationToken.IsCancellationRequested)
                    break;

                var progressItem = progressItems.FirstOrDefault(p => p.Category == category.ToString());
                if (progressItem != null)
                {
                    progressItem.Status = "Processing...";
                    try
                    {
                        await processor(connection, settings.OutputFolder, progressItem, logCallback, cancellationToken);
                        progressItem.Status = "Completed";
                    }
                    catch (Exception ex)
                    {
                        progressItem.Status = "Error";
                        logCallback(new DownloadLogEntry
                        {
                            Category = category.ToString(),
                            Message = $"Fatal error: {ex.Message}",
                            Level = LogLevel.Error
                        });
                    }
                }
            }
        }

        #region Document Processors

        private async Task ProcessRiskDocumentsAsync(SqlConnection connection, string baseFolder,
            DownloadProgressItem progress, Action<DownloadLogEntry> logCallback, CancellationToken ct)
        {
            string riskFolder = Path.Combine(baseFolder, "Risk");
            EnsureDirectoryExists(riskFolder);

            string query = @"
                SELECT 
                    RADoc.AssessmentDocumentId, 
                    RADoc.AssessmentDetailId,
                    RADet.RiskCode,
                    RADet.RiskTypeId,
                    RRT.FieldName,
                    RADet.Title AS RiskAssessmentDetailTitle,
                    RADoc.Title, 
                    RADoc.FileName, 
                    RADoc.FileData  
                FROM 
                    RISK_AssessmentDocument AS RADoc 
                INNER JOIN 
                    RISK_AssessmentDetail AS RADet 
                ON
                    RADet.AssessmentDetailId = RADoc.AssessmentDetailId 
                INNER JOIN
                    RISK_RiskType AS RRT
                ON
                    RRT.RiskTypeId = RADet.RiskTypeId
                ORDER BY 
                    RADet.RiskTypeId";

            await ProcessDocumentsAsync(connection, query, progress, logCallback, ct, reader =>
            {
                string fieldName = reader["FieldName"]?.ToString() ?? "Unknown_FieldName";
                string riskCode = reader["RiskCode"]?.ToString() ?? "Unknown_RiskCode";
                string documentTitle = reader["Title"]?.ToString() ?? "Untitled";
                string fileName = reader["FileName"]?.ToString() ?? string.Empty;
                byte[]? fileData = reader["FileData"] as byte[];

                if (fileData == null || fileData.Length == 0)
                    return (null, "No file data", true);

                string fieldNameFolder = Path.Combine(riskFolder, SanitizeFolderName(fieldName));
                string subFolder = Path.Combine(fieldNameFolder, SanitizeFolderName(riskCode));
                EnsureDirectoryExists(subFolder);

                fileName = DetermineFileName(fileName, documentTitle, fileData);
                string filePath = Path.Combine(subFolder, fileName);

                return (new DocumentInfo { FileData = fileData, FilePath = filePath, FileName = fileName }, null, false);
            });
        }

        private async Task ProcessIncidentDocumentsAsync(SqlConnection connection, string baseFolder,
            DownloadProgressItem progress, Action<DownloadLogEntry> logCallback, CancellationToken ct)
        {
            string incidentFolder = Path.Combine(baseFolder, "Incident");
            EnsureDirectoryExists(incidentFolder);

            string query = @"
                SELECT
                    I.IncidentId,
                    I.IncidentTitle,
                    I.IncidentCode,
                    E.DocumentId,
                    E.[File],
                    E.Name,
                    E.FilePath
                FROM
                    Incident AS I
                INNER JOIN
                    EntityDocument AS E ON E.ObjectDataId = I.IncidentId
                WHERE
                    E.[File] IS NOT NULL AND E.IsDeleted = 0
                ORDER BY
                    I.IncidentCode";

            await ProcessDocumentsAsync(connection, query, progress, logCallback, ct, reader =>
            {
                string incidentCode = reader["IncidentCode"]?.ToString() ?? "Unknown_Code";
                string originalFileName = reader["Name"]?.ToString() ?? string.Empty;
                string filePath = reader["FilePath"]?.ToString() ?? string.Empty;
                byte[]? fileData = reader["File"] as byte[];

                if (fileData == null || fileData.Length == 0)
                    return (null, "No file data", true);

                string incidentSubFolder = Path.Combine(incidentFolder, SanitizeFolderName(incidentCode));
                EnsureDirectoryExists(incidentSubFolder);

                string namePart = !string.IsNullOrWhiteSpace(originalFileName)
                    ? Path.GetFileNameWithoutExtension(originalFileName)
                    : $"Document_{progress.TotalDocuments}";

                string extension = GetExtensionFromSource(originalFileName, filePath, fileData);
                string fileName = SanitizeFileName($"{incidentCode}_{namePart}{extension}");
                string fullPath = Path.Combine(incidentSubFolder, fileName);

                return (new DocumentInfo { FileData = fileData, FilePath = fullPath, FileName = fileName }, null, false);
            });
        }

        private async Task ProcessControlDocumentsAsync(SqlConnection connection, string baseFolder,
            DownloadProgressItem progress, Action<DownloadLogEntry> logCallback, CancellationToken ct)
        {
            string controlFolder = Path.Combine(baseFolder, "Control");
            EnsureDirectoryExists(controlFolder);

            string query = @"
                SELECT
                    ControlDetailId,
                    Title,
                    FileName,
                    FileData
                FROM
                    ControlDetails A INNER JOIN
                    ControlDocuments B ON A.id = B.ControlDetailId
                ORDER BY
                    ControlDetailId";

            await ProcessDocumentsAsync(connection, query, progress, logCallback, ct, reader =>
            {
                int controlDetailId = reader.GetInt32(reader.GetOrdinal("ControlDetailId"));
                string title = reader["Title"]?.ToString() ?? "Untitled";
                string fileName = reader["FileName"]?.ToString() ?? string.Empty;
                byte[]? fileData = reader["FileData"] as byte[];

                if (fileData == null || fileData.Length == 0)
                    return (null, $"Control ID {controlDetailId} - No file data", true);

                string controlSubFolder = Path.Combine(controlFolder, SanitizeFolderName($"Control_{controlDetailId}_{title}"));
                EnsureDirectoryExists(controlSubFolder);

                fileName = DetermineFileName(fileName, title, fileData);
                string filePath = Path.Combine(controlSubFolder, fileName);

                return (new DocumentInfo { FileData = fileData, FilePath = filePath, FileName = fileName }, null, false);
            });
        }

        private async Task ProcessActionDocumentsAsync(SqlConnection connection, string baseFolder,
            DownloadProgressItem progress, Action<DownloadLogEntry> logCallback, CancellationToken ct)
        {
            string actionFolder = Path.Combine(baseFolder, "Action");
            EnsureDirectoryExists(actionFolder);

            string query = @"
                SELECT
                    ActionDetailId,
                    Title,
                    FileName,
                    FileData
                FROM
                    Action_Document";

            await ProcessDocumentsAsync(connection, query, progress, logCallback, ct, reader =>
            {
                int actionDetailId = reader.GetInt32(reader.GetOrdinal("ActionDetailId"));
                string title = reader["Title"]?.ToString() ?? "Untitled";
                string fileName = reader["FileName"]?.ToString() ?? string.Empty;
                byte[]? fileData = reader["FileData"] as byte[];

                if (fileData == null || fileData.Length == 0)
                    return (null, $"Action ID {actionDetailId} - No file data", true);

                string actionSubFolder = Path.Combine(actionFolder, SanitizeFolderName($"Action_{actionDetailId}_{title}"));
                EnsureDirectoryExists(actionSubFolder);

                fileName = DetermineFileName(fileName, title, fileData);
                string filePath = Path.Combine(actionSubFolder, fileName);

                return (new DocumentInfo { FileData = fileData, FilePath = filePath, FileName = fileName }, null, false);
            });
        }

        private async Task ProcessComplianceDocumentsAsync(SqlConnection connection, string baseFolder,
            DownloadProgressItem progress, Action<DownloadLogEntry> logCallback, CancellationToken ct)
        {
            string query = @"
                SELECT
                    CASE 
                        WHEN I.IncidentID IS NOT NULL THEN 'Incident_Linked_Compliance_Documents' 
                        WHEN C.ComplianceId IS NOT NULL THEN 'Compliance' 
                        WHEN AD.AuthorityDocumentId IS NOT NULL THEN 'AuthorityDocument' 
                        WHEN P.PolicyId IS NOT NULL THEN 'Policy' 
                    END AS 'ApplicationFolderName',
                    CASE 
                        WHEN I.IncidentID IS NOT NULL THEN I.IncidentCode + ' - ' + I.IncidentTitle 
                        WHEN C.ComplianceId IS NOT NULL THEN C.Code + ' - ' + C.Title 
                        WHEN AD.AuthorityDocumentId IS NOT NULL THEN AD.Code + ' - ' + AD.Title 
                        WHEN P.PolicyId IS NOT NULL THEN P.Code + ' - ' + P.Title 
                    END AS 'EntityFolderName',
                    [FILE],
                    FilePath
                FROM EntityDocument ED
                LEFT OUTER JOIN Incident I ON ED.ObjectDataId = I.IncidentID AND ED.IMSApplicationID = 1
                LEFT OUTER JOIN Compliance C ON ED.ObjectDataId = C.ComplianceID AND ED.IMSApplicationID = 2 AND IMSSubApplicationID = 3
                LEFT OUTER JOIN AuthorityDocument AD ON ED.ObjectDataId = AD.AuthorityDocumentID AND ED.IMSApplicationID = 2 AND IMSSubApplicationID = 4
                LEFT OUTER JOIN Policy P ON ED.ObjectDataId = P.PolicyID AND ED.IMSApplicationID = 2 AND IMSSubApplicationID = 5
                WHERE ED.[File] IS NOT NULL AND ED.IsDeleted = 0
                ORDER BY ApplicationFolderName, EntityFolderName";

            await ProcessDocumentsAsync(connection, query, progress, logCallback, ct, reader =>
            {
                string? applicationFolder = reader["ApplicationFolderName"]?.ToString();
                string? entityFolder = reader["EntityFolderName"]?.ToString();
                string filePath = reader["FilePath"]?.ToString() ?? string.Empty;
                byte[]? fileData = reader["FILE"] as byte[];

                if (fileData == null || fileData.Length == 0)
                    return (null, "No file data", true);

                if (string.IsNullOrWhiteSpace(applicationFolder) || string.IsNullOrWhiteSpace(entityFolder))
                    return (null, "Missing folder information", true);

                string appFolder = Path.Combine(baseFolder, SanitizeFolderName(applicationFolder));
                string entityPath = Path.Combine(appFolder, SanitizeFolderName(entityFolder));
                EnsureDirectoryExists(entityPath);

                string fileName = Path.GetFileName(filePath);
                if (string.IsNullOrWhiteSpace(fileName))
                    fileName = $"document_{progress.TotalDocuments}";

                if (!Path.HasExtension(fileName))
                    fileName += GetExtensionFromContentType(null, fileData);

                fileName = SanitizeFileName(fileName);
                string fullPath = Path.Combine(entityPath, fileName);

                return (new DocumentInfo { FileData = fileData, FilePath = fullPath, FileName = fileName }, null, false);
            });
        }

        private async Task ProcessAuditRecommendationDocumentsAsync(SqlConnection connection, string baseFolder,
            DownloadProgressItem progress, Action<DownloadLogEntry> logCallback, CancellationToken ct)
        {
            string auditFolder = Path.Combine(baseFolder, "Audit_Recommendations");
            EnsureDirectoryExists(auditFolder);

            string query = @"
                SELECT
                    AR.RECOMMENDATIONID,
                    AR.RECOMMENDATIONNO,
                    AR.RECOMMENDATIONTITLE,
                    A.AttachmentID,
                    A.Title,
                    A.DocumentURL,
                    A.FileData,
                    A.ContentType
                FROM AUDITRECOMMENDATION AR
                INNER JOIN Attachment A ON AR.RECOMMENDATIONID = A.ObjectID
                WHERE A.FileData IS NOT NULL 
                    AND (A.Deleted IS NULL OR A.Deleted = 0)
                ORDER BY AR.RECOMMENDATIONNO";

            await ProcessDocumentsAsync(connection, query, progress, logCallback, ct, reader =>
            {
                string recommendationNo = reader["RECOMMENDATIONNO"]?.ToString() ?? "Unknown";
                string recommendationTitle = reader["RECOMMENDATIONTITLE"]?.ToString() ?? "Untitled";
                string title = reader["Title"]?.ToString() ?? "Untitled";
                string documentUrl = reader["DocumentURL"]?.ToString() ?? string.Empty;
                byte[]? fileData = reader["FileData"] as byte[];
                string contentType = reader["ContentType"]?.ToString() ?? string.Empty;

                if (fileData == null || fileData.Length == 0)
                    return (null, "No file data", true);

                string recommendationFolder = Path.Combine(auditFolder, 
                    SanitizeFolderName($"Recommendation_{recommendationNo}_{recommendationTitle}"));
                EnsureDirectoryExists(recommendationFolder);

                string fileName = DetermineFileNameFromUrl(documentUrl, title, contentType, fileData);
                string filePath = Path.Combine(recommendationFolder, fileName);

                return (new DocumentInfo { FileData = fileData, FilePath = filePath, FileName = fileName }, null, false);
            });
        }

        private async Task ProcessAuditDetailsDocumentsAsync(SqlConnection connection, string baseFolder,
            DownloadProgressItem progress, Action<DownloadLogEntry> logCallback, CancellationToken ct)
        {
            string auditFolder = Path.Combine(baseFolder, "Audit_Details_Attachments");
            EnsureDirectoryExists(auditFolder);

            string query = @"
                SELECT
                    AD.AuditDetailID,
                    AD.AuditNo,
                    AD.AuditTitle,
                    A.AttachmentID,
                    A.Title,
                    A.DocumentURL,
                    A.FileData,
                    A.ContentType
                FROM AUDITDETAIL AD
                INNER JOIN Attachment A ON AD.AuditDetailID = A.ObjectID
                WHERE A.FileData IS NOT NULL 
                    AND (A.Deleted IS NULL OR A.Deleted = 0)
                ORDER BY AD.AuditNo";

            await ProcessDocumentsAsync(connection, query, progress, logCallback, ct, reader =>
            {
                string auditNo = reader["AuditNo"]?.ToString() ?? "Unknown";
                string auditTitle = reader["AuditTitle"]?.ToString() ?? "Untitled";
                string title = reader["Title"]?.ToString() ?? "Untitled";
                string documentUrl = reader["DocumentURL"]?.ToString() ?? string.Empty;
                byte[]? fileData = reader["FileData"] as byte[];
                string contentType = reader["ContentType"]?.ToString() ?? string.Empty;

                if (fileData == null || fileData.Length == 0)
                    return (null, "No file data", true);

                string auditSubFolder = Path.Combine(auditFolder, SanitizeFolderName($"{auditNo}_{auditTitle}"));
                EnsureDirectoryExists(auditSubFolder);

                string fileName = DetermineFileNameFromUrl(documentUrl, title, contentType, fileData);
                string filePath = Path.Combine(auditSubFolder, fileName);

                return (new DocumentInfo { FileData = fileData, FilePath = filePath, FileName = fileName }, null, false);
            });
        }

        private async Task ProcessAuditFindingDocumentsAsync(SqlConnection connection, string baseFolder,
            DownloadProgressItem progress, Action<DownloadLogEntry> logCallback, CancellationToken ct)
        {
            string auditFolder = Path.Combine(baseFolder, "Audit_Finding_Attachments");
            EnsureDirectoryExists(auditFolder);

            string query = @"
                SELECT
                    AF.AuditFindingID,
                    AF.AuditFindingNo,
                    AD.AuditNo,
                    AD.AuditTitle,
                    A.AttachmentID,
                    A.Title,
                    A.DocumentURL,
                    A.FileData,
                    A.ContentType
                FROM AUDITFINDING AF
                INNER JOIN AUDITDETAIL AD ON AF.AuditDetailID = AD.AuditDetailID
                INNER JOIN Attachment A ON AF.AuditFindingID = A.ObjectID
                WHERE A.FileData IS NOT NULL 
                    AND (A.Deleted IS NULL OR A.Deleted = 0)
                ORDER BY AD.AuditNo, AF.AuditFindingNo";

            await ProcessDocumentsAsync(connection, query, progress, logCallback, ct, reader =>
            {
                string auditNo = reader["AuditNo"]?.ToString() ?? "Unknown";
                string auditTitle = reader["AuditTitle"]?.ToString() ?? "Untitled";
                string findingNo = reader["AuditFindingNo"]?.ToString() ?? "N/A";
                string title = reader["Title"]?.ToString() ?? "Untitled";
                string documentUrl = reader["DocumentURL"]?.ToString() ?? string.Empty;
                byte[]? fileData = reader["FileData"] as byte[];
                string contentType = reader["ContentType"]?.ToString() ?? string.Empty;

                if (fileData == null || fileData.Length == 0)
                    return (null, "No file data", true);

                string auditSubFolder = Path.Combine(auditFolder, SanitizeFolderName($"{auditNo}_{auditTitle}"));
                string findingFolder = Path.Combine(auditSubFolder, SanitizeFolderName($"Finding_{findingNo}"));
                EnsureDirectoryExists(findingFolder);

                string fileName = DetermineFileNameFromUrl(documentUrl, title, contentType, fileData);
                string filePath = Path.Combine(findingFolder, fileName);

                return (new DocumentInfo { FileData = fileData, FilePath = filePath, FileName = fileName }, null, false);
            });
        }

        private async Task ProcessPolicyDocumentsAsync(SqlConnection connection, string baseFolder,
            DownloadProgressItem progress, Action<DownloadLogEntry> logCallback, CancellationToken ct)
        {
            string policyFolder = Path.Combine(baseFolder, "Policy");
            EnsureDirectoryExists(policyFolder);

            string query = @"
                SELECT
                    P.PolicyId,
                    P.Code,
                    P.Title,
                    ED.DocumentId,
                    ED.Name AS DocumentName,
                    ED.FilePath,
                    ED.[File] AS FileData
                FROM Policy P
                INNER JOIN EntityDocument ED ON P.ObjectId = ED.ObjectId
                WHERE ED.[File] IS NOT NULL 
                    AND (ED.IsDeleted IS NULL OR ED.IsDeleted = 0)
                    AND (P.IsDeleted IS NULL OR P.IsDeleted = 0)
                ORDER BY P.Code, P.Title";

            await ProcessDocumentsAsync(connection, query, progress, logCallback, ct, reader =>
            {
                string code = reader["Code"]?.ToString() ?? "Unknown";
                string title = reader["Title"]?.ToString() ?? "Untitled";
                string documentName = reader["DocumentName"]?.ToString() ?? string.Empty;
                string filePath = reader["FilePath"]?.ToString() ?? string.Empty;
                byte[]? fileData = reader["FileData"] as byte[];

                if (fileData == null || fileData.Length == 0)
                    return (null, "No file data", true);

                string policySubFolder = Path.Combine(policyFolder, SanitizeFolderName($"{code}_{title}"));
                EnsureDirectoryExists(policySubFolder);

                string fileName;
                if (!string.IsNullOrWhiteSpace(filePath))
                    fileName = Path.GetFileName(filePath);
                else if (!string.IsNullOrWhiteSpace(documentName))
                    fileName = documentName;
                else
                    fileName = $"policy_document_{progress.TotalDocuments}";

                string extension = Path.GetExtension(fileName);
                if (string.IsNullOrWhiteSpace(extension) || extension == ".bin" || extension == ".dat")
                    fileName = Path.GetFileNameWithoutExtension(fileName) + GetExtensionFromContentType(null, fileData);

                fileName = SanitizeFileName(fileName);
                string fullPath = Path.Combine(policySubFolder, fileName);

                return (new DocumentInfo { FileData = fileData, FilePath = fullPath, FileName = fileName }, null, false);
            });
        }

        #endregion

        #region Common Processing Logic

        private class DocumentInfo
        {
            public byte[] FileData { get; set; } = Array.Empty<byte>();
            public string FilePath { get; set; } = string.Empty;
            public string FileName { get; set; } = string.Empty;
        }

        private async Task ProcessDocumentsAsync(
            SqlConnection connection,
            string query,
            DownloadProgressItem progress,
            Action<DownloadLogEntry> logCallback,
            CancellationToken ct,
            Func<SqlDataReader, (DocumentInfo? Doc, string? SkipReason, bool IsSkip)> documentExtractor)
        {
            using var command = new SqlCommand(query, connection);
            command.CommandTimeout = 300;

            try
            {
                using var reader = await command.ExecuteReaderAsync(ct);

                while (await reader.ReadAsync(ct))
                {
                    if (ct.IsCancellationRequested)
                        break;

                    progress.TotalDocuments++;

                    try
                    {
                        var (doc, skipReason, isSkip) = documentExtractor(reader);

                        if (isSkip || doc == null)
                        {
                            progress.FailedCount++;
                            logCallback(new DownloadLogEntry
                            {
                                Category = progress.Category,
                                Message = $"Skipped: {skipReason}",
                                Level = LogLevel.Warning
                            });
                            continue;
                        }

                        // Check for duplicates using file hash
                        string fileHash = ComputeFileHash(doc.FileData);
                        
                        if (_downloadedFiles.TryGetValue(fileHash, out var existingFile))
                        {
                            progress.SkippedDuplicates++;
                            logCallback(new DownloadLogEntry
                            {
                                Category = progress.Category,
                                Message = $"Duplicate detected: {doc.FileName} (same as {existingFile.FileName})",
                                Level = LogLevel.Warning,
                                FileName = doc.FileName
                            });
                            continue;
                        }

                        // Get unique file path to avoid overwriting
                        string finalPath = GetUniqueFilePath(doc.FilePath);
                        
                        // Write file
                        await File.WriteAllBytesAsync(finalPath, doc.FileData, ct);

                        // Track the file
                        _downloadedFiles.TryAdd(fileHash, new FileTracker
                        {
                            FileName = doc.FileName,
                            FileHash = fileHash,
                            FileSize = doc.FileData.Length,
                            SourcePath = finalPath
                        });

                        progress.SuccessCount++;
                        progress.CurrentFile = doc.FileName;

                        logCallback(new DownloadLogEntry
                        {
                            Category = progress.Category,
                            Message = $"Downloaded: {doc.FileName}",
                            Level = LogLevel.Success,
                            FileName = doc.FileName,
                            FileSize = FormatFileSize(doc.FileData.Length)
                        });
                    }
                    catch (Exception ex)
                    {
                        progress.FailedCount++;
                        logCallback(new DownloadLogEntry
                        {
                            Category = progress.Category,
                            Message = $"Error: {ex.Message}",
                            Level = LogLevel.Error
                        });
                    }
                }
            }
            catch (SqlException ex)
            {
                logCallback(new DownloadLogEntry
                {
                    Category = progress.Category,
                    Message = $"SQL Error: {ex.Message}. Table may not exist in this database.",
                    Level = LogLevel.Warning
                });
            }
        }

        #endregion

        #region Helper Methods

        private static string ComputeFileHash(byte[] fileData)
        {
            using var sha256 = SHA256.Create();
            byte[] hashBytes = sha256.ComputeHash(fileData);
            return Convert.ToBase64String(hashBytes);
        }

        private static void EnsureDirectoryExists(string path)
        {
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
        }

        private static string DetermineFileName(string fileName, string fallbackTitle, byte[] fileData)
        {
            if (string.IsNullOrWhiteSpace(fileName))
                fileName = fallbackTitle;

            string extension = Path.GetExtension(fileName);
            if (string.IsNullOrWhiteSpace(extension) || extension == ".bin" || extension == ".dat")
            {
                fileName = Path.GetFileNameWithoutExtension(fileName) + GetExtensionFromContentType(null, fileData);
            }

            return SanitizeFileName(fileName);
        }

        private static string DetermineFileNameFromUrl(string documentUrl, string title, string contentType, byte[] fileData)
        {
            string fileName;
            if (!string.IsNullOrWhiteSpace(documentUrl))
                fileName = Path.GetFileName(documentUrl);
            else if (!string.IsNullOrWhiteSpace(title))
                fileName = title;
            else
                fileName = $"attachment_{Guid.NewGuid():N}";

            if (!Path.HasExtension(fileName))
                fileName += GetExtensionFromContentType(contentType, fileData);

            return SanitizeFileName(fileName);
        }

        private static string GetExtensionFromSource(string originalFileName, string filePath, byte[] fileData)
        {
            string extension = string.Empty;

            if (!string.IsNullOrWhiteSpace(originalFileName))
                extension = Path.GetExtension(originalFileName);
            else if (!string.IsNullOrWhiteSpace(filePath))
                extension = Path.GetExtension(filePath);

            if (string.IsNullOrWhiteSpace(extension))
                extension = GetExtensionFromContentType(null, fileData);

            return extension;
        }

        private static string GetExtensionFromContentType(string? contentType, byte[] fileData)
        {
            // Check magic bytes first
            if (fileData != null && fileData.Length >= 4)
            {
                if (fileData[0] == 0x25 && fileData[1] == 0x50 && fileData[2] == 0x44 && fileData[3] == 0x46)
                    return ".pdf";

                if (fileData[0] == 0xD0 && fileData[1] == 0xCF && fileData[2] == 0x11 && fileData[3] == 0xE0)
                {
                    string mime = contentType?.ToLower() ?? "";
                    if (mime.Contains("msword") || mime.Contains("wordprocessingml")) return ".doc";
                    if (mime.Contains("ms-excel") || mime.Contains("spreadsheetml")) return ".xls";
                    if (mime.Contains("ms-powerpoint")) return ".ppt";
                    return ".msg";
                }

                if (fileData[0] == 0x50 && fileData[1] == 0x4B)
                {
                    string mime = contentType?.ToLower() ?? "";
                    if (mime.Contains("wordprocessingml")) return ".docx";
                    if (mime.Contains("spreadsheetml")) return ".xlsx";
                    if (mime.Contains("presentationml")) return ".pptx";
                    return ".docx";
                }

                if (fileData[0] == 0xFF && fileData[1] == 0xD8)
                    return ".jpg";

                if (fileData[0] == 0x89 && fileData[1] == 0x50 && fileData[2] == 0x4E && fileData[3] == 0x47)
                    return ".png";

                if (fileData[0] == 0x47 && fileData[1] == 0x49 && fileData[2] == 0x46)
                    return ".gif";
            }

            // Fall back to content type
            if (!string.IsNullOrWhiteSpace(contentType))
            {
                return contentType.ToLower().Trim() switch
                {
                    "application/pdf" => ".pdf",
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document" => ".docx",
                    "application/msword" => ".doc",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" => ".xlsx",
                    "application/vnd.ms-excel" => ".xls",
                    "application/vnd.openxmlformats-officedocument.presentationml.presentation" => ".pptx",
                    "application/vnd.ms-powerpoint" => ".ppt",
                    "image/jpeg" => ".jpg",
                    "image/png" => ".png",
                    "image/gif" => ".gif",
                    "text/plain" => ".txt",
                    "text/csv" => ".csv",
                    "application/vnd.ms-outlook" => ".msg",
                    "application/x-msg" => ".msg",
                    "message/rfc822" => ".msg",
                    _ => ".bin"
                };
            }

            return ".bin";
        }

        private static string SanitizeFolderName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return "Unknown";

            name = name.Trim();

            foreach (char c in Path.GetInvalidPathChars())
                name = name.Replace(c, '_');

            name = name.Replace("/", "_")
                       .Replace("\\", "_")
                       .Replace(":", "_")
                       .Replace("*", "_")
                       .Replace("?", "_")
                       .Replace("\"", "_")
                       .Replace("<", "_")
                       .Replace(">", "_")
                       .Replace("|", "_");

            if (name.Length > 100)
                name = name.Substring(0, 100);

            return name.Trim();
        }

        private static string SanitizeFileName(string fileName)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
                fileName = fileName.Replace(c, '_');
            return fileName;
        }

        private static string GetUniqueFilePath(string filePath)
        {
            if (!File.Exists(filePath))
                return filePath;

            string directory = Path.GetDirectoryName(filePath) ?? string.Empty;
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(filePath);
            string extension = Path.GetExtension(filePath);

            int counter = 1;
            string newFilePath;
            do
            {
                newFilePath = Path.Combine(directory, $"{fileNameWithoutExt}_{counter}{extension}");
                counter++;
            } while (File.Exists(newFilePath));

            return newFilePath;
        }

        private static string FormatFileSize(long bytes)
        {
            string[] sizes = { "B", "KB", "MB", "GB" };
            double len = bytes;
            int order = 0;
            while (len >= 1024 && order < sizes.Length - 1)
            {
                order++;
                len /= 1024;
            }
            return $"{len:0.##} {sizes[order]}";
        }

        #endregion
    }
}
