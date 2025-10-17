using CsvHelper;
using CsvHelper.Configuration;
using SharePointPermissionsExporter.Models;
using System.Globalization;

namespace SharePointPermissionsExporter.Services;

/// <summary>
/// Service for exporting file permissions to CSV format
/// </summary>
public class CsvExportService
{
    /// <summary>
    /// Exports a list of FilePermissionInfo objects to a CSV file
    /// </summary>
    /// <param name="permissions">List of file permissions to export</param>
    /// <param name="outputFilePath">Path to the output CSV file</param>
    /// <returns>Task representing the async operation</returns>
    public async Task ExportToCsvAsync(List<FilePermissionInfo> permissions, string outputFilePath)
    {
        try
        {
            Console.WriteLine($"\nExporting {permissions.Count} permission records to CSV...");

            // Ensure the directory exists
            var directory = Path.GetDirectoryName(outputFilePath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Configure CSV writer
            var config = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                HasHeaderRecord = true,
                TrimOptions = TrimOptions.Trim,
                ShouldQuote = args => true // Quote all fields for safety
            };

            // Write to CSV
            await using var writer = new StreamWriter(outputFilePath);
            await using var csv = new CsvWriter(writer, config);

            // Register custom class map
            csv.Context.RegisterClassMap<FilePermissionInfoMap>();

            // Write records
            await csv.WriteRecordsAsync(permissions);

            Console.WriteLine($"✓ Successfully exported to: {Path.GetFullPath(outputFilePath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error exporting to CSV: {ex.Message}");
            throw;
        }
    }

    /// <summary>
    /// Exports permissions to CSV with custom formatting
    /// </summary>
    /// <param name="permissions">List of file permissions to export</param>
    /// <param name="outputFilePath">Path to the output CSV file</param>
    /// <param name="includeInheritedPermissions">Whether to include inherited permissions</param>
    /// <returns>Task representing the async operation</returns>
    public async Task ExportToCsvAsync(
        List<FilePermissionInfo> permissions,
        string outputFilePath,
        bool includeInheritedPermissions)
    {
        var filteredPermissions = includeInheritedPermissions
            ? permissions
            : permissions.Where(p => !p.IsInherited).ToList();

        Console.WriteLine($"Filtering: {filteredPermissions.Count} of {permissions.Count} permissions " +
                         $"(inherited permissions {(includeInheritedPermissions ? "included" : "excluded")})");

        await ExportToCsvAsync(filteredPermissions, outputFilePath);
    }
}

/// <summary>
/// Custom class map for FilePermissionInfo CSV export
/// </summary>
public sealed class FilePermissionInfoMap : ClassMap<FilePermissionInfo>
{
    public FilePermissionInfoMap()
    {
        Map(m => m.FileName).Name("File Name").Index(0);
        Map(m => m.WebUrl).Name("Web URL").Index(1);
        Map(m => m.FileId).Name("File ID").Index(2);
        Map(m => m.PermissionId).Name("Permission ID").Index(3);
        Map(m => m.RolesString).Name("Roles").Index(4);
        Map(m => m.GrantedToDisplayName).Name("Granted To (Name)").Index(5);
        Map(m => m.GrantedToEmail).Name("Granted To (Email)").Index(6);
        Map(m => m.IsInherited).Name("Is Inherited").Index(7);
        Map(m => m.InheritedFrom).Name("Inherited From").Index(8);
    }
}