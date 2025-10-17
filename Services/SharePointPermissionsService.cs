using Microsoft.Graph;
using Microsoft.Graph.Models;
using SharePointPermissionsExporter.Models;
using System.Net;

namespace SharePointPermissionsExporter.Services;

/// <summary>
/// Service for retrieving SharePoint document library permissions using Microsoft Graph API
/// </summary>
public class SharePointPermissionsService
{
    private readonly GraphServiceClient _graphClient;
    private readonly int _delayBetweenRequestsMs;
    private readonly int _batchSize;
    private readonly int _maxRetryAttempts;

    /// <summary>
    /// Initializes a new instance of the SharePointPermissionsService
    /// </summary>
    /// <param name="graphClient">Authenticated GraphServiceClient</param>
    /// <param name="delayBetweenRequestsMs">Delay between requests in milliseconds (default: 100)</param>
    /// <param name="batchSize">Number of requests per batch (default: 20)</param>
    /// <param name="maxRetryAttempts">Maximum retry attempts for failed requests (default: 3)</param>
    public SharePointPermissionsService(
        GraphServiceClient graphClient,
        int delayBetweenRequestsMs = 100,
        int batchSize = 20,
        int maxRetryAttempts = 3)
    {
        _graphClient = graphClient ?? throw new ArgumentNullException(nameof(graphClient));
        _delayBetweenRequestsMs = delayBetweenRequestsMs;
        _batchSize = batchSize;
        _maxRetryAttempts = maxRetryAttempts;
    }

    /// <summary>
    /// Retrieves Site ID and Drive ID from SharePoint URL
    /// </summary>
    /// <param name="siteUrl">Full SharePoint site URL</param>
    /// <param name="documentLibraryName">Optional document library name</param>
    /// <returns>Tuple containing SiteId and DriveId</returns>
    public async Task<(string SiteId, string DriveId)> GetSiteAndDriveIdsAsync(
        string siteUrl,
        string? documentLibraryName = null)
    {
        try
        {
            Console.WriteLine($"Retrieving site information from: {siteUrl}");

            // Parse the site URL to extract host and site path
            var uri = new Uri(siteUrl);
            var hostName = uri.Host;
            var sitePath = uri.AbsolutePath;

            // Get the site using the hostname and site path
            var site = await _graphClient.Sites[$"{hostName}:{sitePath}"]
                .GetAsync();

            if (site == null || string.IsNullOrEmpty(site.Id))
            {
                throw new Exception("Failed to retrieve site information");
            }

            Console.WriteLine($"Site ID: {site.Id}");

            // Get the drive (document library)
            string driveId;

            if (!string.IsNullOrWhiteSpace(documentLibraryName))
            {
                // Get specific document library by name
                var drives = await _graphClient.Sites[site.Id].Drives
                    .GetAsync();

                var targetDrive = drives?.Value?.FirstOrDefault(d =>
                    d.Name?.Equals(documentLibraryName, StringComparison.OrdinalIgnoreCase) == true);

                if (targetDrive == null)
                {
                    throw new Exception($"Document library '{documentLibraryName}' not found");
                }

                driveId = targetDrive.Id!;
                Console.WriteLine($"Found document library '{documentLibraryName}' with Drive ID: {driveId}");
            }
            else
            {
                // Get default document library
                var drive = await _graphClient.Sites[site.Id].Drive
                    .GetAsync();

                if (drive == null || string.IsNullOrEmpty(drive.Id))
                {
                    throw new Exception("Failed to retrieve default document library");
                }

                driveId = drive.Id;
                Console.WriteLine($"Using default document library with Drive ID: {driveId}");
            }

            return (site.Id, driveId);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error retrieving site and drive IDs: {ex.Message}");
            throw;
        }
    }

    /// <summary>
    /// Recursively retrieves all files from the document library
    /// </summary>
    /// <param name="siteId">SharePoint Site ID</param>
    /// <param name="driveId">Document Library Drive ID</param>
    /// <returns>List of all DriveItem files</returns>
    public async Task<List<DriveItem>> GetAllFilesRecursivelyAsync(string siteId, string driveId)
    {
        var allFiles = new List<DriveItem>();

        try
        {
            Console.WriteLine("Starting recursive file retrieval...");

            // Start from root
            await TraverseFolderAsync(siteId, driveId, "root", allFiles);

            Console.WriteLine($"Total files found: {allFiles.Count}");
            return allFiles;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error retrieving files: {ex.Message}");
            throw;
        }
    }

    /// <summary>
    /// Recursively traverses folders to collect all files
    /// </summary>
    private async Task TraverseFolderAsync(
        string siteId,
        string driveId,
        string itemId,
        List<DriveItem> allFiles)
    {
        try
        {
            var children = await _graphClient.Drives[driveId]
                .Items[itemId]
                .Children
                .GetAsync(config =>
                {
                    config.QueryParameters.Top = 999;
                });

            if (children?.Value == null) return;

            // Use PageIterator to handle pagination
            var pageIterator = PageIterator<DriveItem, DriveItemCollectionResponse>
                .CreatePageIterator(
                    _graphClient,
                    children,
                    (item) =>
                    {
                        if (item.File != null)
                        {
                            // It's a file
                            allFiles.Add(item);
                        }
                        return true;
                    });

            await pageIterator.IterateAsync();

            // Now traverse subfolders
            foreach (var item in children.Value)
            {
                if (item.Folder != null && !string.IsNullOrEmpty(item.Id))
                {
                    // It's a folder, recurse into it
                    await TraverseFolderAsync(siteId, driveId, item.Id, allFiles);
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error traversing folder {itemId}: {ex.Message}");
            // Continue with other folders
        }
    }

    /// <summary>
    /// Retrieves permissions for a specific file
    /// </summary>
    /// <param name="siteId">SharePoint Site ID</param>
    /// <param name="driveId">Document Library Drive ID</param>
    /// <param name="fileId">File ID</param>
    /// <returns>List of permissions for the file</returns>
    public async Task<List<Permission>?> GetFilePermissionsAsync(
        string siteId,
        string driveId,
        string fileId)
    {
        try
        {
            var permissions = await _graphClient.Drives[driveId]
                .Items[fileId]
                .Permissions
                .GetAsync();

            return permissions?.Value?.ToList();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error retrieving permissions for file {fileId}: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Main orchestration method to get all file permissions with progress tracking
    /// </summary>
    /// <param name="siteId">SharePoint Site ID</param>
    /// <param name="driveId">Document Library Drive ID</param>
    /// <returns>List of FilePermissionInfo objects</returns>
    public async Task<List<FilePermissionInfo>> GetAllFilePermissionsAsync(
        string siteId,
        string driveId)
    {
        var allPermissions = new List<FilePermissionInfo>();
        var errorCount = 0;

        try
        {
            // Get all files
            var files = await GetAllFilesRecursivelyAsync(siteId, driveId);
            var totalFiles = files.Count;

            Console.WriteLine($"\nProcessing permissions for {totalFiles} files...");

            for (int i = 0; i < files.Count; i++)
            {
                var file = files[i];

                try
                {
                    // Progress indicator
                    if ((i + 1) % 10 == 0 || i == 0)
                    {
                        Console.WriteLine($"Progress: {i + 1}/{totalFiles} files processed");
                    }

                    // Get permissions for this file
                    var permissions = await GetFilePermissionsAsync(siteId, driveId, file.Id!);

                    if (permissions != null && permissions.Any())
                    {
                        foreach (var permission in permissions)
                        {
                            var permInfo = new FilePermissionInfo
                            {
                                FileName = file.Name ?? "Unknown",
                                WebUrl = file.WebUrl ?? string.Empty,
                                FileId = file.Id ?? string.Empty,
                                PermissionId = permission.Id ?? string.Empty,
                                Roles = permission.Roles?.ToList() ?? new List<string>(),
                                IsInherited = permission.InheritedFrom != null,
                                InheritedFrom = permission.InheritedFrom?.Path ?? string.Empty
                            };

                            // Extract grantee information
                            if (permission.GrantedToV2?.User != null)
                            {
                                permInfo.GrantedToDisplayName = permission.GrantedToV2.User.DisplayName ?? string.Empty;
                                permInfo.GrantedToEmail = permission.GrantedToV2.User.Id ?? string.Empty;
                            }
                            else if (permission.GrantedToV2?.Group != null)
                            {
                                permInfo.GrantedToDisplayName = permission.GrantedToV2.Group.DisplayName ?? string.Empty;
                                permInfo.GrantedToEmail = permission.GrantedToV2.Group.Id ?? string.Empty;
                            }
                            else if (permission.GrantedToIdentitiesV2 != null)
                            {
                                var firstIdentity = permission.GrantedToIdentitiesV2.FirstOrDefault();
                                if (firstIdentity?.User != null)
                                {
                                    permInfo.GrantedToDisplayName = firstIdentity.User.DisplayName ?? string.Empty;
                                    permInfo.GrantedToEmail = firstIdentity.User.Id ?? string.Empty;
                                }
                            }

                            allPermissions.Add(permInfo);
                        }
                    }

                    // Delay between requests to avoid throttling
                    await Task.Delay(_delayBetweenRequestsMs);
                }
                catch (Exception ex)
                {
                    errorCount++;
                    Console.WriteLine($"Error processing file '{file.Name}': {ex.Message}");
                    // Continue with next file
                }
            }

            Console.WriteLine($"\nCompleted! Total permissions retrieved: {allPermissions.Count}");
            Console.WriteLine($"Errors encountered: {errorCount}");

            return allPermissions;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error in GetAllFilePermissionsAsync: {ex.Message}");
            throw;
        }
    }

    /// <summary>
    /// Optimized version using Graph API batching with retry logic and throttling handling
    /// </summary>
    /// <param name="siteId">SharePoint Site ID</param>
    /// <param name="driveId">Document Library Drive ID</param>
    /// <returns>List of FilePermissionInfo objects</returns>
    public async Task<List<FilePermissionInfo>> GetAllFilePermissionsWithBatchingAsync(
        string siteId,
        string driveId)
    {
        var allPermissions = new List<FilePermissionInfo>();
        var errorCount = 0;
        var throttleCount = 0;

        try
        {
            // Get all files
            var files = await GetAllFilesRecursivelyAsync(siteId, driveId);
            var totalFiles = files.Count;

            Console.WriteLine($"\nProcessing permissions for {totalFiles} files using batching...");

            // Process files in batches
            for (int i = 0; i < files.Count; i += _batchSize)
            {
                var batch = files.Skip(i).Take(_batchSize).ToList();
                var batchNumber = (i / _batchSize) + 1;
                var totalBatches = (int)Math.Ceiling((double)totalFiles / _batchSize);

                Console.WriteLine($"Processing batch {batchNumber}/{totalBatches} ({batch.Count} files)");

                // Create batch request
                var batchRequestContent = new BatchRequestContentCollection(_graphClient);
                var requestIds = new Dictionary<string, DriveItem>();

                foreach (var file in batch)
                {
                    var requestId = Guid.NewGuid().ToString();
                    var request = _graphClient.Drives[driveId]
                        .Items[file.Id!]
                        .Permissions
                        .ToGetRequestInformation();

                    await batchRequestContent.AddBatchRequestStepAsync(request, requestId);

                    requestIds[requestId] = file;
                }

                // Execute batch with retry logic
                var batchResponse = await ExecuteBatchWithRetryAsync(batchRequestContent);

                // Process batch responses
                foreach (var kvp in requestIds)
                {
                    try
                    {
                        var file = kvp.Value;
                        var requestId = kvp.Key;

                        var response = await batchResponse.GetResponseByIdAsync<PermissionCollectionResponse>(requestId);

                        if (response?.Value != null && response.Value.Any())
                        {
                            foreach (var permission in response.Value)
                            {
                                var permInfo = new FilePermissionInfo
                                {
                                    FileName = file.Name ?? "Unknown",
                                    WebUrl = file.WebUrl ?? string.Empty,
                                    FileId = file.Id ?? string.Empty,
                                    PermissionId = permission.Id ?? string.Empty,
                                    Roles = permission.Roles?.ToList() ?? new List<string>(),
                                    IsInherited = permission.InheritedFrom != null,
                                    InheritedFrom = permission.InheritedFrom?.Path ?? string.Empty
                                };

                                // Extract grantee information
                                if (permission.GrantedToV2?.User != null)
                                {
                                    permInfo.GrantedToDisplayName = permission.GrantedToV2.User.DisplayName ?? string.Empty;
                                    permInfo.GrantedToEmail = permission.GrantedToV2.User.Id ?? string.Empty;
                                }
                                else if (permission.GrantedToV2?.Group != null)
                                {
                                    permInfo.GrantedToDisplayName = permission.GrantedToV2.Group.DisplayName ?? string.Empty;
                                    permInfo.GrantedToEmail = permission.GrantedToV2.Group.Id ?? string.Empty;
                                }
                                else if (permission.GrantedToIdentitiesV2 != null)
                                {
                                    var firstIdentity = permission.GrantedToIdentitiesV2.FirstOrDefault();
                                    if (firstIdentity?.User != null)
                                    {
                                        permInfo.GrantedToDisplayName = firstIdentity.User.DisplayName ?? string.Empty;
                                        permInfo.GrantedToEmail = firstIdentity.User.Id ?? string.Empty;
                                    }
                                }

                                allPermissions.Add(permInfo);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        errorCount++;
                        Console.WriteLine($"Error processing file in batch: {ex.Message}");
                    }
                }

                // Delay between batches
                await Task.Delay(_delayBetweenRequestsMs);
            }

            Console.WriteLine($"\nCompleted! Total permissions retrieved: {allPermissions.Count}");
            Console.WriteLine($"Errors encountered: {errorCount}");
            Console.WriteLine($"Throttle events: {throttleCount}");

            return allPermissions;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error in GetAllFilePermissionsWithBatchingAsync: {ex.Message}");
            throw;
        }
    }

    /// <summary>
    /// Executes a batch request with retry logic and exponential backoff
    /// </summary>
    private async Task<BatchResponseContentCollection> ExecuteBatchWithRetryAsync(
        BatchRequestContentCollection batchRequestContent)
    {
        int retryCount = 0;
        int delayMs = 1000;

        while (retryCount < _maxRetryAttempts)
        {
            try
            {
                var batchResponse = await _graphClient.Batch.PostAsync(batchRequestContent);

                if (batchResponse == null)
                {
                    throw new Exception("Batch response is null");
                }

                return batchResponse;
            }
            catch (ServiceException ex) when (ex.ResponseStatusCode == (int)HttpStatusCode.TooManyRequests)
            {
                retryCount++;
                Console.WriteLine($"⚠️ Throttled (429). Retry attempt {retryCount}/{_maxRetryAttempts}");

                // Check for Retry-After header
                if (ex.ResponseHeaders?.TryGetValues("Retry-After", out var retryAfterValues) == true)
                {
                    var retryAfter = retryAfterValues.FirstOrDefault();
                    if (int.TryParse(retryAfter, out int retryAfterSeconds))
                    {
                        delayMs = retryAfterSeconds * 1000;
                        Console.WriteLine($"Waiting {retryAfterSeconds} seconds as indicated by Retry-After header");
                    }
                }

                if (retryCount >= _maxRetryAttempts)
                {
                    Console.WriteLine("Max retry attempts reached. Throwing exception.");
                    throw;
                }

                await Task.Delay(delayMs);
                delayMs *= 2; // Exponential backoff
            }
            catch (Exception ex)
            {
                retryCount++;
                Console.WriteLine($"Error executing batch (attempt {retryCount}/{_maxRetryAttempts}): {ex.Message}");

                if (retryCount >= _maxRetryAttempts)
                {
                    throw;
                }

                await Task.Delay(delayMs);
                delayMs *= 2; // Exponential backoff
            }
        }

        throw new Exception("Failed to execute batch after maximum retry attempts");
    }
}