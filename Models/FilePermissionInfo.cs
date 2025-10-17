namespace SharePointPermissionsExporter.Models;

/// <summary>
/// Represents permission information for a file in SharePoint Document Library
/// </summary>
public class FilePermissionInfo
{
    /// <summary>
    /// Name of the file
    /// </summary>
    public string FileName { get; set; } = string.Empty;

    /// <summary>
    /// Full web URL to the file
    /// </summary>
    public string WebUrl { get; set; } = string.Empty;

    /// <summary>
    /// Unique identifier for the file
    /// </summary>
    public string FileId { get; set; } = string.Empty;

    /// <summary>
    /// Unique identifier for the permission
    /// </summary>
    public string PermissionId { get; set; } = string.Empty;

    /// <summary>
    /// List of roles assigned (e.g., read, write, owner)
    /// </summary>
    public List<string> Roles { get; set; } = new();

    /// <summary>
    /// Display name of the user or group granted permission
    /// </summary>
    public string GrantedToDisplayName { get; set; } = string.Empty;

    /// <summary>
    /// Email address of the user or group granted permission
    /// </summary>
    public string GrantedToEmail { get; set; } = string.Empty;

    /// <summary>
    /// Indicates whether the permission is inherited from parent
    /// </summary>
    public bool IsInherited { get; set; }

    /// <summary>
    /// Path or URL from which the permission is inherited
    /// </summary>
    public string InheritedFrom { get; set; } = string.Empty;

    /// <summary>
    /// Comma-separated string of roles for CSV export
    /// </summary>
    public string RolesString => string.Join(", ", Roles);
}