using SharePointPermissionsExporter.Models;
using System.Text;
using System.Text.Json;

namespace SharePointPermissionsExporter.Services;

/// <summary>
/// Service for exporting file permissions to JSON format
/// </summary>
public class JsonExportService
{
    private readonly JsonSerializerOptions _jsonOptions;

    /// <summary>
    /// Initializes a new instance of the JsonExportService with configured serialization options
    /// </summary>
    public JsonExportService()
    {
        _jsonOptions = new JsonSerializerOptions
        {
            WriteIndented = true,
            PropertyNameCaseInsensitive = true,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase
        };
    }

    /// <summary>
    /// Exports a list of FilePermissionInfo objects to a JSON file
    /// </summary>
    /// <param name="permissions">List of file permissions to export</param>
    /// <param name="outputFilePath">Path to the output JSON file</param>
    /// <returns>Task representing the async operation</returns>
    public async Task ExportToJsonAsync(List<FilePermissionInfo> permissions, string outputFilePath)
    {
        try
        {
            Console.WriteLine($"\nExporting {permissions.Count} permission records to JSON...");

            // Ensure the directory exists
            var directory = Path.GetDirectoryName(outputFilePath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Group permissions by file and create hierarchical structure
            var groupedByFile = permissions
                .GroupBy(p => p.FileId)
                .Select(g => new
                {
                    fileName = DecodeBase64IfEncoded(g.First().FileName),
                    webUrl = DecodeBase64IfEncoded(g.First().WebUrl),
                    fileId = DecodeBase64IfEncoded(g.Key),
                    permissions = g.Select(p => new
                    {
                        permissionId = DecodeBase64IfEncoded(p.PermissionId),
                        roles = p.Roles.Select(DecodeBase64IfEncoded).ToList(),
                        grantedToDisplayName = DecodeBase64IfEncoded(p.GrantedToDisplayName),
                        grantedToEmail = DecodeBase64IfEncoded(p.GrantedToEmail),
                        isInherited = p.IsInherited,
                        inheritedFrom = DecodeBase64IfEncoded(p.InheritedFrom)
                    }).ToList()
                })
                .ToList();

            // Serialize to JSON
            var jsonString = JsonSerializer.Serialize(groupedByFile, _jsonOptions);

            // Write to file asynchronously
            await File.WriteAllTextAsync(outputFilePath, jsonString, Encoding.UTF8);

            Console.WriteLine($"✓ Successfully exported to: {Path.GetFullPath(outputFilePath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error exporting to JSON: {ex.Message}");
            throw;
        }
    }

    /// <summary>
    /// Exports permissions to JSON with optional filtering
    /// </summary>
    /// <param name="permissions">List of file permissions to export</param>
    /// <param name="outputFilePath">Path to the output JSON file</param>
    /// <param name="includeInheritedPermissions">Whether to include inherited permissions</param>
    /// <returns>Task representing the async operation</returns>
    public async Task ExportToJsonAsync(
        List<FilePermissionInfo> permissions,
        string outputFilePath,
        bool includeInheritedPermissions)
    {
        var filteredPermissions = includeInheritedPermissions
            ? permissions
            : permissions.Where(p => !p.IsInherited).ToList();

        Console.WriteLine($"Filtering: {filteredPermissions.Count} of {permissions.Count} permissions " +
                         $"(inherited permissions {(includeInheritedPermissions ? "included" : "excluded")})");

        await ExportToJsonAsync(filteredPermissions, outputFilePath);
    }

    /// <summary>
    /// Attempts to decode a base64 encoded string. If decoding fails or the string is not base64 encoded,
    /// returns the original value. This provides future-proofing for potential base64 encoded data from Graph API.
    /// </summary>
    /// <param name="value">The string value to potentially decode</param>
    /// <returns>The decoded string if successful, otherwise the original value</returns>
    private string DecodeBase64IfEncoded(string value)
    {
        if (string.IsNullOrEmpty(value))
        {
            return value;
        }

        try
        {
            // Check if the string could be base64 (basic validation)
            // Base64 strings have length that's a multiple of 4 and only contain valid base64 characters
            if (value.Length % 4 == 0 && System.Text.RegularExpressions.Regex.IsMatch(value, @"^[a-zA-Z0-9+/]*={0,2}$"))
            {
                var decodedBytes = Convert.FromBase64String(value);
                var decodedString = Encoding.UTF8.GetString(decodedBytes);
                
                // Only return decoded value if it contains printable characters
                // This helps avoid false positives where normal strings might look like base64
                if (!string.IsNullOrWhiteSpace(decodedString) && decodedString.All(c => !char.IsControl(c) || char.IsWhiteSpace(c)))
                {
                    return decodedString;
                }
            }
        }
        catch
        {
            // If decoding fails, return original value
        }

        return value;
    }
}