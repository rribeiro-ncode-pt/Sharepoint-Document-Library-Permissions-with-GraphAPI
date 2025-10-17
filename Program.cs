using Microsoft.Extensions.Configuration;
using SharePointPermissionsExporter.Services;
using System.Diagnostics;

namespace SharePointPermissionsExporter;

/// <summary>
/// Main program entry point for SharePoint Document Library Permissions Exporter
/// </summary>
class Program
{
    static async Task<int> Main(string[] args)
    {
        var stopwatch = Stopwatch.StartNew();

        try
        {
            Console.WriteLine("=================================================");
            Console.WriteLine("SharePoint Document Library Permissions Exporter");
            Console.WriteLine("=================================================\n");

            // Load configuration
            var configuration = LoadConfiguration();

            // Validate configuration
            if (!ValidateConfiguration(configuration))
            {
                Console.WriteLine("\n❌ Configuration validation failed. Please check appsettings.json");
                return 1;
            }

            // Extract configuration values
            var tenantId = configuration["AzureAd:TenantId"]!;
            var clientId = configuration["AzureAd:ClientId"]!;
            var clientSecret = configuration["AzureAd:ClientSecret"]!;
            var siteUrl = configuration["SharePoint:SiteUrl"]!;
            var documentLibraryName = configuration["SharePoint:DocumentLibraryName"];
            var baseFileName = configuration["Export:OutputFileName"] ?? "SharePointPermissions";
            var exportFormat = configuration["Export:ExportFormat"]?.ToLowerInvariant() ?? "csv";
            var outputFileName = $"{baseFileName}.{exportFormat}";
            var delayMs = int.Parse(configuration["Export:DelayBetweenRequestsMs"] ?? "100");
            var batchSize = int.Parse(configuration["Export:BatchSize"] ?? "20");
            var maxRetryAttempts = int.Parse(configuration["Export:MaxRetryAttempts"] ?? "3");

            Console.WriteLine("Configuration loaded successfully:");
            Console.WriteLine($"  - Site URL: {siteUrl}");
            Console.WriteLine($"  - Document Library: {(string.IsNullOrWhiteSpace(documentLibraryName) ? "(default)" : documentLibraryName)}");
            Console.WriteLine($"  - Output File: {outputFileName}");
            Console.WriteLine($"  - Export Format: {exportFormat.ToUpperInvariant()}");
            Console.WriteLine($"  - Batch Size: {batchSize}");
            Console.WriteLine($"  - Delay Between Requests: {delayMs}ms");
            Console.WriteLine($"  - Max Retry Attempts: {maxRetryAttempts}\n");

            // Step 1: Authenticate
            Console.WriteLine("Step 1: Authenticating with Microsoft Graph API...");
            var authService = new AuthenticationService(tenantId, clientId, clientSecret);

            if (!authService.ValidateCredentials())
            {
                Console.WriteLine("\n❌ Invalid credentials. Please check your Azure AD configuration.");
                return 1;
            }

            var graphClient = authService.GetAuthenticatedGraphClient();
            Console.WriteLine("✓ Authentication successful\n");

            // Step 2: Get Site and Drive IDs
            Console.WriteLine("Step 2: Retrieving SharePoint site and drive information...");
            var permissionsService = new SharePointPermissionsService(
                graphClient,
                delayMs,
                batchSize,
                maxRetryAttempts);

            var (siteId, driveId) = await permissionsService.GetSiteAndDriveIdsAsync(
                siteUrl,
                documentLibraryName);

            Console.WriteLine("✓ Site and drive information retrieved\n");

            // Step 3: Fetch all file permissions using batched method
            Console.WriteLine("Step 3: Fetching file permissions (this may take a while)...");
            var permissions = await permissionsService.GetAllFilePermissionsWithBatchingAsync(
                siteId,
                driveId);

            if (permissions.Count == 0)
            {
                Console.WriteLine("\n⚠️ No permissions found. The document library may be empty or inaccessible.");
                return 0;
            }

            Console.WriteLine($"✓ Retrieved {permissions.Count} permission records\n");

            // Step 4: Export to file
            Console.WriteLine($"Step 4: Exporting permissions to {exportFormat.ToUpperInvariant()}...");
            
            if (exportFormat == "json")
            {
                var jsonExportService = new JsonExportService();
                await jsonExportService.ExportToJsonAsync(permissions, outputFileName);
            }
            else
            {
                var csvExportService = new CsvExportService();
                await csvExportService.ExportToCsvAsync(permissions, outputFileName);
            }

            // Display summary statistics
            stopwatch.Stop();
            DisplaySummary(permissions, stopwatch.Elapsed);

            Console.WriteLine("\n✓ Process completed successfully!");
            return 0;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"\n❌ Fatal error: {ex.Message}");
            Console.WriteLine($"Stack trace: {ex.StackTrace}");
            return 1;
        }
    }

    /// <summary>
    /// Loads configuration from appsettings.json
    /// </summary>
    private static IConfiguration LoadConfiguration()
    {
        try
        {
            var builder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);

            return builder.Build();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading configuration: {ex.Message}");
            throw;
        }
    }

    /// <summary>
    /// Validates that all required configuration values are present
    /// </summary>
    private static bool ValidateConfiguration(IConfiguration configuration)
    {
        var requiredSettings = new[]
        {
            "AzureAd:TenantId",
            "AzureAd:ClientId",
            "AzureAd:ClientSecret",
            "SharePoint:SiteUrl"
        };

        var isValid = true;

        foreach (var setting in requiredSettings)
        {
            var value = configuration[setting];
            if (string.IsNullOrWhiteSpace(value) || value.Contains("your-") || value.Contains("-here"))
            {
                Console.WriteLine($"❌ Missing or invalid configuration: {setting}");
                isValid = false;
            }
        }

        return isValid;
    }

    /// <summary>
    /// Displays summary statistics about the export
    /// </summary>
    private static void DisplaySummary(List<Models.FilePermissionInfo> permissions, TimeSpan elapsed)
    {
        Console.WriteLine("\n=================================================");
        Console.WriteLine("Summary Statistics");
        Console.WriteLine("=================================================");
        Console.WriteLine($"Total Permissions Exported: {permissions.Count}");
        Console.WriteLine($"Unique Files: {permissions.Select(p => p.FileId).Distinct().Count()}");
        Console.WriteLine($"Inherited Permissions: {permissions.Count(p => p.IsInherited)}");
        Console.WriteLine($"Direct Permissions: {permissions.Count(p => !p.IsInherited)}");
        Console.WriteLine($"Unique Users/Groups: {permissions.Select(p => p.GrantedToEmail).Distinct().Count()}");
        Console.WriteLine($"Total Execution Time: {elapsed.TotalSeconds:F2} seconds");
        Console.WriteLine("=================================================");

        // Display top 5 most common roles
        var topRoles = permissions
            .SelectMany(p => p.Roles)
            .GroupBy(r => r)
            .OrderByDescending(g => g.Count())
            .Take(5)
            .ToList();

        if (topRoles.Any())
        {
            Console.WriteLine("\nTop 5 Most Common Roles:");
            foreach (var role in topRoles)
            {
                Console.WriteLine($"  - {role.Key}: {role.Count()} occurrences");
            }
        }
    }
}